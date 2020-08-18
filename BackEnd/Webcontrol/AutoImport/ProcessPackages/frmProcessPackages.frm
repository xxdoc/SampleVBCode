VERSION 5.00
Object = "{B71A484A-57D1-11D2-821F-000086075197}#1.0#0"; "FTPX.OCX"
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MsgHoo32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8BAF5903-01D9-11D0-9E0A-444553540000}#5.1#0"; "mmail32.ocx"
Begin VB.Form frmProcessPackages 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Package (User name)"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "frmProcessPackages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin MailLib.mMail PP_MAIL 
      Left            =   3360
      Top             =   120
      _Version        =   327681
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Blocking        =   0   'False
      Debug           =   0
      Host            =   ""
      Timeout         =   0
      ConnectType     =   2
      PopPort         =   110
      SmtpPort        =   25
      AuthenticationType=   0
   End
   Begin FtpXCtl.FtpXCtl PP_FTP 
      Left            =   2640
      Top             =   120
      Blocking        =   -1  'True
      DebugMode       =   1
      Directory       =   ""
      DstFilename     =   ""
      Host            =   ""
      LogonPassword   =   ""
      Pattern         =   ""
      SrcFilename     =   ""
      Type            =   1
      LogonName       =   ""
      Account         =   ""
      Timeout         =   0
      Port            =   21
      DisablePasv     =   0   'False
      DirItemPattern  =   ""
      LibraryName     =   "WSOCK32.DLL"
      BlockingMode    =   0
      FirewallType    =   0
      FirewallHost    =   ""
      FirewallPort    =   0
      FirewallLogonName=   ""
      FirewallPassword=   ""
   End
   Begin VB.Timer Timer_Start 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer Timer_SpinMe 
      Enabled         =   0   'False
      Interval        =   200
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessPackages.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessPackages.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessPackages.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessPackages.frx":0C58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgBarLoss 
      Height          =   375
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Right-Click for Options"
      Top             =   8040
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtMess 
      Appearance      =   0  'Flat
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
      Height          =   8415
      Left            =   40
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   11295
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
Attribute VB_Name = "frmProcessPackages"
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
Private madoRSAssignments As ADODB.Recordset
Private mArv As V2ARViewer.clsARViewer
Private moForm As Form
Private mbHourly As Boolean
Private mbCheckingActiveFiles As Boolean
Private mbImporting As Boolean
Private mbExporting As Boolean
'Help process Packages
'Package Process params
Private msUserName As String
Private msPassWord As String
Private msAdjUserName As String
Private msAssignmentsID As String
Private msPackageID As String
Private msExportDocPath As String
Private msSendToFTP As String
Private msSendToFTPUserName As String
Private msSendToFTPPassWord As String
Private msSendToHTTP As String
Private msSendToHTTPUserName As String
Private msSendToHTTPPassWord As String
Private msEmailEntireClaim As String
Private msEmailEntireClaimCC As String
Private msEmailEntireClaimBCC As String
Private msEmailDocsOnly As String
Private msEmailDocsOnlyCC As String
Private msEmailDocsOnlyBCC As String
Private msEmailPhotosOnly As String
Private msEmailPhotosOnlyCC As String
Private msEmailPhotosOnlyBCC As String
Private msCarListClassName As String
'This is used when sending a Single PDF only not a delivery to client
Private mbCreateSinglePdfOnly As Boolean
Private msEmailSinglePdfOnlySubject As String
Private msEmailSinglePdfOnlyBody As String
Private msSendSinglePDFFileOnly As String
Private msZipNameSinglePDFFileOnly As String
Private msSendSinglePdfOnlyQueue As String
Private msPackageEmailQueueID As String
Private msSendSinglePdfOnlyQueueMess As String
'This is used when sending a Single PDF only not a delivery to client
Private msEmailSinglePdfOnlyTo As String
Private msEmailSinglePdfOnlyCC As String
Private msEmailSinglePdfOnlyBCC As String
Private mbDoEmailSinglePdfOnly As Boolean

'End Package Process params
Private msUsersID As String
Private msBuildAssgnPackPath As String
Private mbSHUTDOWN As Boolean
'Token info to be passed to Export Download (ExportDL)
Private msMyTokoutPath As String
Private msMyTokenData As String

'The Members will be set When the packager creates the Zip Files
Private msZipNameAll As String
Private msZipNameDocsOnly As String
Private msZipNamePhotosOnly As String
Private mbDoB2B As Boolean
Private msSentB2B As String
Private msBackupB2BMess As String
Private msSendFTP As String
Private msSentFTP As String
Private msBackupFTPMess As String
Private mbDoFTP As Boolean
Private msSendEmailEntireClaim As String
Private msSentEmailEntireClaim As String
Private msBackupEmailEntireClaimMess As String
Private mbDoEmailEntireClaim As Boolean
Private msSendEmailDocsOnly As String
Private msSentEmailDocsOnly As String
Private msBackupEmailDocsOnlyMess As String
Private mbDoEmailDocsOnly As Boolean
Private msSendEmailPhotosOnly As String
Private msSentEmailPhotosOnly As String
Private msBackupEmailPhotosOnlyMess As String
Private mbDoEmailPhotosOnly As Boolean

'Package Errors ...
Private msPackageErrors As String
'Mail State
Private mlMailState As Long
Private Const MAIL_STATE_SENDING = 1
Private Const MAIL_STATE_CONNECTING = 2
Private Const MAIL_STATE_DISCONNECTING = 3



Public Property Let adoRSAssignments(padoRS As ADODB.Recordset)
    Set madoRSAssignments = padoRS
End Property
Public Property Set adoRSAssignments(padoRS As ADODB.Recordset)
    Set madoRSAssignments = padoRS
End Property
Public Property Get adoRSAssignments() As ADODB.Recordset
    Set adoRSAssignments = madoRSAssignments
End Property

Public Property Let BuildAssgnPackPath(psBuildAssgnPackPath As String)
    msBuildAssgnPackPath = psBuildAssgnPackPath
End Property

Public Property Get BuildAssgnPackPath() As String
    BuildAssgnPackPath = msBuildAssgnPackPath
End Property

Public Property Let UserName(psUserName As String)
    msUserName = psUserName
End Property

Public Property Get UserName() As String
    UserName = msUserName
End Property

Public Property Let PassWord(psPassWord As String)
    msPassWord = psPassWord
End Property

Public Property Get PassWord() As String
     PassWord = msPassWord
End Property

Public Property Let AdjUserName(psAdjUserName As String)
    msAdjUserName = psAdjUserName
End Property

Public Property Get AdjUserName() As String
     AdjUserName = msAdjUserName
End Property

Public Property Let AssignmentsID(psAssignmentsID As String)
    msAssignmentsID = psAssignmentsID
End Property

Public Property Get AssignmentsID() As String
     AssignmentsID = msAssignmentsID
End Property

Public Property Let PackageID(psPackageID As String)
    msPackageID = psPackageID
End Property

Public Property Get PackageID() As String
     PackageID = msPackageID
End Property

Public Property Let ExportDocPath(psExportDocPath As String)
    msExportDocPath = psExportDocPath
End Property

Public Property Get ExportDocPath() As String
     ExportDocPath = msExportDocPath
End Property

Public Property Let SendToFTP(psSendToFTP As String)
    msSendToFTP = psSendToFTP
End Property

Public Property Get SendToFTP() As String
     SendToFTP = msSendToFTP
End Property

Public Property Let EmailEntireClaim(psEmailEntireClaim As String)
    msEmailEntireClaim = psEmailEntireClaim
End Property

Public Property Get EmailEntireClaim() As String
     EmailEntireClaim = msEmailEntireClaim
End Property


Public Property Let EmailDocsOnly(psEmailDocsOnly As String)
    msEmailDocsOnly = psEmailDocsOnly
End Property

Public Property Get EmailDocsOnly() As String
     EmailDocsOnly = msEmailDocsOnly
End Property

Public Property Let EmailPhotosOnly(psEmailPhotosOnly As String)
    msEmailPhotosOnly = psEmailPhotosOnly
End Property

Public Property Get EmailPhotosOnly() As String
     EmailPhotosOnly = msEmailPhotosOnly
End Property

Public Property Let CarListClassName(psCarListClassName As String)
    msCarListClassName = psCarListClassName
End Property

Public Property Get CarListClassName() As String
     CarListClassName = msCarListClassName
End Property

Public Property Let CreateSinglePdfOnly(pbCreateSinglePdfOnly As Boolean)
    mbCreateSinglePdfOnly = pbCreateSinglePdfOnly
End Property

Public Property Get CreateSinglePdfOnly() As Boolean
     CreateSinglePdfOnly = mbCreateSinglePdfOnly
End Property

Public Property Let PackageEmailQueueID(psPackageEmailQueueID As String)
    msPackageEmailQueueID = psPackageEmailQueueID
End Property

Public Property Get PackageEmailQueueID() As String
     PackageEmailQueueID = msPackageEmailQueueID
End Property


Private Sub Form_Load()
    On Error GoTo EH
    Dim sMess As String
    
    ' Don't want to be visible initially!
    Me.Visible = False
    Me.Caption = "Process Package (" & msAdjUserName & " - " & msAssignmentsID & ")"
    App.Title = Me.Caption
'    FormWinRegPos Me
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
    
    'Clear any previous files
    goUtil.utDeleteFile msBuildAssgnPackPath & "*.pdf"
    goUtil.utDeleteFile msBuildAssgnPackPath & "*.xml"
    
    Timer_SpinMe.Enabled = True
    Timer_Start.Enabled = True
    
   Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Form_Load" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Function REJECT00() As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sSQL As String
    Dim lRecordsAffected As Long
    Dim sProdDsn As String
    Dim sAdminComments As String
    
    sAdminComments = Left(msPackageErrors, 1000)
    
    
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    
    sSQL = "UPDATE Assignments SET "
    sSQL = sSQL & "[StatusID] = (SELECT [StatusID] FROM Status WHERE [Status] = 'REJECT00'), "
    sSQL = sSQL & "[DownLoadMe] = 1, "
    sSQL = sSQL & "[UpdateByUserID] = (SELECT [UsersID] FROM Users WHERE [UserName] = 'CFUSER'), "
    sSQL = sSQL & "[DateLastUpdated] = GetDate() "
    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
    
    gConn.Execute sSQL, lRecordsAffected
    
    sSQL = "UPDATE Package SET "
    sSQL = sSQL & "[PackageStatus] = (SELECT [Description] FROM Status WHERE [Status] = 'REJECT00'), "
    sSQL = sSQL & "[SendMe] = 0, "
    sSQL = sSQL & "[DownLoadMe] = 1, "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(sAdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = GetDate(), "
    sSQL = sSQL & "[UpdateByUserID] = (SELECT [UsersID] FROM Users WHERE [UserName] = 'CFUSER') "
    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
    sSQL = sSQL & "AND   [PackageID] = " & msPackageID & " "
    
    gConn.Execute sSQL, lRecordsAffected
    
    REJECT00 = True
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function REJECT00" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Function

Private Function DELIVERPackage() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sSortOrder As String
    Dim bDebugMode As Boolean
    Dim sLossFormat As String
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim RSEmail As ADODB.Recordset
    Dim sMess As String
    Dim sCurDocName As String
    Dim sCurDocDesc As String
    Dim sReportFormat As String
    Dim bPrintActiveReport As Boolean
    Dim sTickCount As String
    Dim sLossType As String
    Dim oWddxDeser As WDDXDeserializer
    Dim oWddxSer As WDDXSerializer
    Dim oWddxStruct As WDDXStruct
    Dim oWddxLossRS As WDDXRecordset
    Dim oWddxAssignmentDetailRS As WDDXRecordset
    Dim oWddxLossDetailRS As WDDXRecordset
    Dim oWddxVehicleDetailRS As WDDXRecordset
    Dim oWddxDocumentPropertiesRS As WDDXRecordset
    Dim oWddxIndemnityPaymentRS As WDDXRecordset
    Dim oWddxFeeBillPaymentRS As WDDXRecordset
    Dim sTemp As String
    Dim lDocRsPos As Long
    Dim sProdDsn As String
    'Sequence
    Dim lSequenceNumber As Long
    Dim lTotalDocs As Long
    Dim sGUID As String
    Dim bFail As Boolean
    Dim lRecordsAffected As Long
    'Parse Single PDF File Sent By User
    Dim sPackageAdminComments As String
    Dim sPEQPackageItemIDList As String
    Dim sCurrentPackageItemID As String
    Dim lCountLoop As Long
    Dim sFarmersXML01SubTypeList As String
    
    'Set the TickCount
    
    sTickCount = goUtil.utGetTickCount
    
    'Need to get the Package itmes for this Assingnment
    'Need to get the Package itmes for this Assingnment
    sSQL = "z_spsGetDeliverPackageItems "
    sSQL = sSQL & msAssignmentsID & ", "    '@AssignmentsID      int,
    sSQL = sSQL & "0, "                     '@bVerifyIntegrity   bit=0,
    '@OrderByPackageItemID   bit=0 --V2ECcarFarmers.clsLossXML01 Needs to Sort by PackageItemID
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 And Not mbCreateSinglePdfOnly Then
        sSQL = sSQL & "1 "                   '@OrderByPackageItemID   bit=0
    Else
        sSQL = sSQL & "0 "                   '@OrderByPackageItemID   bit=0
    End If
    
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    
    Set RS = New ADODB.Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 0 Then
        sMess = "No Records Found!"
        DELIVERPackage = False
        GoTo CLEAN_UP
    End If
    
    
    RS.MoveFirst
    lTotalDocs = RS.RecordCount
    
    ProgBarLoss.Max = lTotalDocs
    
    sLossFormat = goUtil.IsNullIsVbNullString(RS.Fields("LRFormat"))
    'Set some member variables
    msExportDocPath = goUtil.IsNullIsVbNullString(RS.Fields("B2BDir"))
    msSendToFTP = goUtil.IsNullIsVbNullString(RS.Fields("FTPSingleFileUrl"))
    msSendToFTPUserName = goUtil.IsNullIsVbNullString(RS.Fields("FTPSingleFileUserName"))
    msSendToFTPPassWord = goUtil.IsNullIsVbNullString(RS.Fields("FTPSingleFilePassword"))
    msSendToHTTP = goUtil.IsNullIsVbNullString(RS.Fields("HttpPostSingleFileUrl"))
    msSendToHTTPUserName = goUtil.IsNullIsVbNullString(RS.Fields("HttpPostSingleFileUserName"))
    msSendToHTTPPassWord = goUtil.IsNullIsVbNullString(RS.Fields("HttpPostSingleFilePassword"))
    msEmailEntireClaim = goUtil.IsNullIsVbNullString(RS.Fields("SingleFileEmail"))
    If StrComp(msEmailEntireClaim, "[EMPTY]", vbTextCompare) = 0 Then
        msEmailEntireClaim = vbNullString
    End If
    msEmailEntireClaimCC = goUtil.IsNullIsVbNullString(RS.Fields("SingleFileEmailCC"))
    msEmailEntireClaimBCC = goUtil.IsNullIsVbNullString(RS.Fields("SingleFileEmailBCC"))
    msEmailDocsOnly = goUtil.IsNullIsVbNullString(RS.Fields("EmailDocsOnly"))
    msEmailDocsOnlyCC = goUtil.IsNullIsVbNullString(RS.Fields("EmailDocsOnlyCC"))
    msEmailDocsOnlyBCC = goUtil.IsNullIsVbNullString(RS.Fields("EmailDocsOnlyBCC"))
    msEmailPhotosOnly = goUtil.IsNullIsVbNullString(RS.Fields("EmailPhotosOnly"))
    msEmailPhotosOnlyCC = goUtil.IsNullIsVbNullString(RS.Fields("EmailPhotosOnlyCC"))
    msEmailPhotosOnlyBCC = goUtil.IsNullIsVbNullString(RS.Fields("EmailPhotosOnlyBCC"))
    ' This is set in the Command line msCarListClassName
    'This is used when sending a Single PDF only not a delivery to client
    'mbCreateSinglePdfOnly As Boolean
    'This is used when sending a Single PDF only not a delivery to client
    'Need to Parse out the User entered Email To ... Cc and BCC
    If mbCreateSinglePdfOnly Then
        sSQL = "spsGetPackageEmailQueue "
        sSQL = sSQL & msPackageEmailQueueID & ", " '@PackageEmailQueueID        int=null,
        sSQL = sSQL & "0 "                         '@bDelPackageEmailQueueID    bit=0
        Set RSEmail = New ADODB.Recordset
        RSEmail.CursorLocation = adUseClient
        RSEmail.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
        Set RSEmail.ActiveConnection = Nothing
        
        If RSEmail.RecordCount = 0 Then
            sMess = "Email Items have been removed from queue! Email not sent!"
            moUL_ErrorMess sMess
            DELIVERPackage = False
            GoTo CLEAN_UP
        End If
        'TO:
        msEmailSinglePdfOnlyTo = goUtil.IsNullIsVbNullString(RSEmail.Fields("PEQEmailTo"))
        'CC:
        msEmailSinglePdfOnlyCC = goUtil.IsNullIsVbNullString(RSEmail.Fields("PEQEmailCC"))
        'BCC:
        msEmailSinglePdfOnlyBCC = goUtil.IsNullIsVbNullString(RSEmail.Fields("PEQEmailBCC"))
        'Subject:
        msEmailSinglePdfOnlySubject = goUtil.IsNullIsVbNullString(RSEmail.Fields("PEQEmailSubject"))
        'Body:
        msEmailSinglePdfOnlyBody = goUtil.IsNullIsVbNullString(RSEmail.Fields("PEQEmailMess"))
        'PEQ PackageItemID List
        sPEQPackageItemIDList = goUtil.IsNullIsVbNullString(RSEmail.Fields("PEQPackageItemIDList"))
    Else
        msEmailSinglePdfOnlyTo = goUtil.IsNullIsVbNullString(RS.Fields("SinglePDFEmail"))
        If StrComp(msEmailSinglePdfOnlyTo, "[EMPTY]", vbTextCompare) = 0 Then
            msEmailSinglePdfOnlyTo = vbNullString
        End If
        msEmailSinglePdfOnlyCC = goUtil.IsNullIsVbNullString(RS.Fields("SinglePDFEmailCC"))
        msEmailSinglePdfOnlyBCC = goUtil.IsNullIsVbNullString(RS.Fields("SinglePDFEmailBCC"))
        msEmailSinglePdfOnlySubject = vbNullString
        msEmailSinglePdfOnlyBody = vbNullString
    End If
   
    
    'Check for Farmers XML01 only stuff
    If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 And Not mbCreateSinglePdfOnly Then
        Set oWddxDeser = New WDDXDeserializer
        sTemp = RS.Fields("LossReport").Value
        Set oWddxStruct = oWddxDeser.deserialize(sTemp)
        Set oWddxAssignmentDetailRS = oWddxStruct.getProp("AssignmentDetailRS")
        sLossType = oWddxStruct.getProp("LossType")
        If InStr(1, sLossType, "Home", vbTextCompare) > 0 Then
            Set oWddxLossDetailRS = oWddxStruct.getProp("LossDetailRS")
        Else
            Set oWddxVehicleDetailRS = oWddxStruct.getProp("VehicleDetailRS")
        End If
    End If
    
    sMess = "Package Process Report " & Now() & vbCrLf
    sMess = sMess & String(100, "-") & vbCrLf & vbCrLf

    sMess = sMess & "Adjuster: " & msAdjUserName & vbCrLf
'    sMess = sMess & "Delivery: " & Trim(msExportDocPath & " " & msSendToFTP & " " & msEmailEntireClaim & " " & msEmailDocsOnly & " ") & vbCrLf
    sMess = sMess & "Total Documents: " & lTotalDocs & vbCrLf
    sCurDocName = "Last Document Processed: " & vbCrLf
    sCurDocDesc = "Process Activity: " & vbCrLf
    txtMess.Text = sMess
    DoEvents
    Sleep 100
    
    Do Until RS.EOF
        lDocRsPos = lDocRsPos + 1
        lSequenceNumber = lDocRsPos
        'Make the GUID  = {Sort_GUID_Name_Desc}.pdf
        If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 And Not mbCreateSinglePdfOnly Then
            sGUID = RS.Fields("PackageItemGUID").Value
            sGUID = "{" & sGUID & "}" & ".pdf"
        Else
            sGUID = RS.Fields("PackageItemGUID").Value
            sTemp = RS.Fields("SortOrder").Value
            sTemp = Format(sTemp, "000")
            sSortOrder = sTemp
            sTemp = Left(RS.Fields("Name").Value, 20)
            sTemp = sTemp & "_"
            sTemp = sTemp & Left(RS.Fields("Description").Value, 20)
            sGUID = sSortOrder & "_" & sTemp & "_" & "{" & sGUID & "}" & ".pdf"
        End If
        goUtil.utCleanFileFolderName sGUID, False
        sReportFormat = goUtil.IsNullIsVbNullString(RS.Fields("ReportFormat"))
        sCurDocName = "Last Document Processed: " & RS.Fields("Name") & vbCrLf
        sCurDocDesc = "Process Activity: " & vbCrLf & "Processed " & lDocRsPos & " Of " & lTotalDocs & vbCrLf
        sCurDocDesc = sCurDocDesc & "Document GUID: " & RS.Fields("PackageItemGUID").Value & vbCrLf
        sCurDocDesc = sCurDocDesc & "Document Marked to Send: " & CStr(CBool(RS.Fields("SendMe").Value)) & vbCrLf
        sCurDocDesc = sCurDocDesc & "Document Marked as Deleted: " & CStr(CBool(RS.Fields("IsDeleted").Value)) & vbCrLf
        sCurDocDesc = sCurDocDesc & "Document Marked as CO Approved: " & CStr(CBool(RS.Fields("IsCoApprove").Value))
        txtMess.Text = sMess & sCurDocName & sCurDocDesc
        DoEvents
        Sleep 100
'DEBUG
bDebugMode = False
If bDebugMode Or mbCreateSinglePdfOnly Then
    GoTo DEBUG_ME
End If
'DEBUG
        If CBool(RS.Fields("SendMe").Value) _
            And Not CBool(RS.Fields("IsDeleted").Value) _
            And CBool(RS.Fields("IsCoApprove").Value) Then
DEBUG_ME:
            'Check for Farmers XML01 only stuff
            If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 And Not mbCreateSinglePdfOnly Then
                'check for Indemnity Payment (check request) and Fee Bills
                Set oWddxStruct = New WDDXStruct
                Set oWddxDocumentPropertiesRS = New WDDXRecordset
                oWddxDocumentPropertiesRS.addColumn "UniqueID"
                oWddxDocumentPropertiesRS.addColumn "UnitNumber"
                oWddxDocumentPropertiesRS.addColumn "SequenceNumber"
                oWddxDocumentPropertiesRS.addColumn "TotalDocs"
                oWddxDocumentPropertiesRS.addColumn "GUID"
                oWddxDocumentPropertiesRS.addColumn "SubType"
                oWddxDocumentPropertiesRS.addColumn "Description"
                oWddxDocumentPropertiesRS.addColumn "State"
                oWddxDocumentPropertiesRS.addRows 1
                With oWddxDocumentPropertiesRS
                    If InStr(1, sLossType, "Home", vbTextCompare) > 0 Then
                        .setField 1, "UniqueID", oWddxAssignmentDetailRS.getField(1, "UniqueID")
                        .setField 1, "UnitNumber", oWddxAssignmentDetailRS.getField(1, "UnitNumber")
                        .setField 1, "SequenceNumber", CStr(lSequenceNumber)
                        .setField 1, "TotalDocs", CStr(lTotalDocs)
                        .setField 1, "GUID", RS.Fields("PackageItemGUID").Value
                        '<xs:enumeration value="Fee Bill" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Fee Bill|"
                        '<xs:enumeration value="IA Investigation Log" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "IA Investigation Log|"
                        '<xs:enumeration value="RCV Agreement" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "RCV Agreement|"
                        '<xs:enumeration value="ALE Worksheet" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "ALE Worksheet|"
                        '<xs:enumeration value="Contents worksheet" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Contents worksheet|"
                        '<xs:enumeration value="Diagram" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Diagram|"
                        '<xs:enumeration value="Scope Sheet" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Scope Sheet|"
                        '<xs:enumeration value="Roof calculation sheet" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Roof calculation sheet|"
                        '<xs:enumeration value="Wind/hail RDF worksheet" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Wind/hail RDF worksheet|"
                        '<xs:enumeration value="Estimate" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Estimate|"
                        '<xs:enumeration value="Cash-in-lieu form" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Cash-in-lieu form|"
                        '<xs:enumeration value="CCC total loss report" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "CCC total loss report|"
                        '<xs:enumeration value="Power of Attorney, POA" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Power of Attorney, POA|"
                        '<xs:enumeration value="637 Total loss evaluation form" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "637 Total loss evaluation form|"
                        '<xs:enumeration value="Image of title" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Image of title|"
                        '<xs:enumeration value="Photos" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Photos|"
                        '<xs:enumeration value="Invoices" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Invoices|"
                        '<xs:enumeration value="Misc. documents" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Misc. documents|"
                        '<xs:enumeration value="Indemnity Payment" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "Indemnity Payment|"
                        '<xs:enumeration value="LossReport" />
                        sFarmersXML01SubTypeList = sFarmersXML01SubTypeList & "LossReport|"
                        
                        If InStr(1, sReportFormat, "_arRptPhotos", vbTextCompare) > 0 Then
                            .setField 1, "SubType", "Photos"
                        ElseIf InStr(1, sReportFormat, "LRFormat", vbTextCompare) > 0 Then
                            .setField 1, "SubType", "LossReport"
                        ElseIf InStr(1, sReportFormat, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                            .setField 1, "SubType", "Diagram"
                        Else
                            sTemp = RS.Fields("Description").Value
                            If Trim(sTemp) = vbNullString Then
                                sTemp = "Misc. documents"
                            ElseIf InStr(1, sFarmersXML01SubTypeList, sTemp, vbTextCompare) = 0 Then
                                sTemp = "Misc. documents"
                            End If
                            .setField 1, "SubType", sTemp
                        End If
                        
                        .setField 1, "Description", RS.Fields("Name").Value
                        .setField 1, "State", RS.Fields("PAState").Value
                    Else
                        .setField 1, "UniqueID", oWddxAssignmentDetailRS.getField(1, "UniqueID")
                        .setField 1, "UnitNumber", oWddxAssignmentDetailRS.getField(1, "UnitNumber")
                        .setField 1, "SequenceNumber", CStr(lSequenceNumber)
                        .setField 1, "TotalDocs", CStr(lTotalDocs)
                        .setField 1, "GUID", RS.Fields("PackageItemGUID").Value
                        If InStr(1, sReportFormat, "_arRptPhotos", vbTextCompare) > 0 Then
                            .setField 1, "SubType", "Photos"
                        ElseIf InStr(1, sReportFormat, "LRFormat", vbTextCompare) > 0 Then
                            .setField 1, "SubType", "LossReport"
                        ElseIf InStr(1, sReportFormat, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                            .setField 1, "SubType", "Diagram"
                        Else
                            sTemp = RS.Fields("Description").Value
                            If Trim(sTemp) = vbNullString Then
                                sTemp = "Misc. documents"
                            End If
                            .setField 1, "SubType", sTemp
                        End If
                        .setField 1, "Description", RS.Fields("Name").Value
                        .setField 1, "State", RS.Fields("PAState").Value
                    End If
                End With
            End If
            'End of Farmers Only XML01
            
            bPrintActiveReport = False
            'Check for Farmers XML01 only stuff
            If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 And Not mbCreateSinglePdfOnly Then
                GoTo START_FARMERS_XML01
            Else
                GoTo START_ALL
            End If
START_FARMERS_XML01:
            If InStr(1, sReportFormat, "_arRptIB", vbTextCompare) > 0 Then
                bPrintActiveReport = PrintActiveReport(Nothing, , vbNullString, False, msBuildAssgnPackPath, sGUID, True, False, sReportFormat, msCarListClassName)
                If bPrintActiveReport Then
                    'Need to Get the xml output for the IB and Harvest the FeeBillPaymentRS
                    Set oWddxDeser = New WDDXDeserializer
                    sGUID = Replace(sGUID, ".pdf", ".xml", , , vbTextCompare)
                    sTemp = goUtil.utGetFileData(msBuildAssgnPackPath & sGUID)
                    Set oWddxStruct = oWddxDeser.deserialize(sTemp)
                    Set oWddxFeeBillPaymentRS = oWddxStruct.getProp("FeeBillPaymentRS")
                    'Need to finish populating some fields...
                    'BillID use the Package GUID !
                    '******************This Field Has been changed 9.19.2005
                    'No longer use the Document GUID for the BILLID
                    'BILLID = IB Number  This is filled out in the active report object
                    'ECRptFarmers_arRptIBFarmers02
'                    oWddxFeeBillPaymentRS.setField 1, "BillID", RS.Fields("PackageItemGUID").Value
                    '******************This Field Has been changed 9.19.2005
                    'Create the XML TransForm to be passed to Biz Talk
                    Set oWddxStruct = New WDDXStruct
                    oWddxStruct.setProp "DocumentPropertiesRS", oWddxDocumentPropertiesRS
                    oWddxStruct.setProp "FeeBillPaymentRS", oWddxFeeBillPaymentRS
                    Set oWddxSer = New WDDXSerializer
                    sTemp = oWddxSer.serialize(oWddxStruct)
                    'Now Remove the pre xml file
                    goUtil.utDeleteFile (msBuildAssgnPackPath & sGUID)
                    'Now save this new one in its place
                    goUtil.utSaveFileData msBuildAssgnPackPath & sGUID, sTemp
                    'At the End Of All Processing The Pdf move first to Bix Talk Dir
                    'Then the xml
                End If
            ElseIf InStr(1, sReportFormat, "_arRptAddlChk", vbTextCompare) > 0 Then
                bPrintActiveReport = PrintActiveReport(Nothing, , vbNullString, False, msBuildAssgnPackPath, sGUID, True, False, sReportFormat, msCarListClassName)
                If bPrintActiveReport Then
                    'Need to Get the xml output for the Indemnity Payment and Harvest the IndemnityPaymentRS
                    Set oWddxDeser = New WDDXDeserializer
                    sGUID = Replace(sGUID, ".pdf", ".xml", , , vbTextCompare)
                    sTemp = goUtil.utGetFileData(msBuildAssgnPackPath & sGUID)
                    Set oWddxStruct = oWddxDeser.deserialize(sTemp)
                    Set oWddxIndemnityPaymentRS = oWddxStruct.getProp("IndemnityPaymentRS")
                    'Need to finish populating some fields...
                    'BillID use the Package GUID !
                    If StrComp(oWddxIndemnityPaymentRS.getField(1, "PaymentGUID"), "[ENTER_PaymentGUID_PackageItemGUID]", vbTextCompare) = 0 Then
                        oWddxIndemnityPaymentRS.setField 1, "PaymentGUID", RS.Fields("PackageItemGUID").Value
                    End If
                    'Check the len of both payee line 1 and payee line 2
                    'They can not be longer than 37 chars.  Farmers really sucks ass.
                    'This is a total BULL SHIT Length for Company names, tax ids are all suppose to fit
                    'inside 37 characters!  What a CROCK ! not to mention that escape chars for & = &amp;amp;
                    'so Bob & Suzy smith = Bob &amp;amp; Suzy smith 24 chars the & = 9 chars out of the 37
                    sTemp = oWddxIndemnityPaymentRS.getField(1, "PayeeLineOne")
                    If Len(sTemp) > 37 Then
                        lCountLoop = 0
                        Do
                            sTemp = Left(sTemp, 37 - lCountLoop)
                            sTemp = CleanXML(sTemp)
                            lCountLoop = lCountLoop + 1
                        Loop Until Len(sTemp) <= 37
                        oWddxIndemnityPaymentRS.setField 1, "PayeeLineOne", sTemp
                    End If
                    sTemp = oWddxIndemnityPaymentRS.getField(1, "PayeeLineTwo")
                    If Len(sTemp) > 37 Then
                        lCountLoop = 0
                        Do
                            sTemp = Left(sTemp, 37 - lCountLoop)
                            sTemp = CleanXML(sTemp)
                            lCountLoop = lCountLoop + 1
                        Loop Until Len(sTemp) <= 37
                        oWddxIndemnityPaymentRS.setField 1, "PayeeLineTwo", sTemp
                    End If
                    If StrComp(oWddxIndemnityPaymentRS.getField(1, "PayeeLineThree"), "[ENTER_PayeeLineThree_MAStreet]", vbTextCompare) = 0 Then
                        sTemp = RS.Fields("MAStreet").Value
                        If Trim(sTemp) = vbNullString Then
                            sTemp = "MAILING STREET UNAVAILABLE"
                        End If
                        oWddxIndemnityPaymentRS.setField 1, "PayeeLineThree", CleanXML(sTemp)
                    ElseIf StrComp(oWddxIndemnityPaymentRS.getField(1, "PayeeLineFour"), "[ENTER_PayeeLineFour_MACity,MAState,MAZip]", vbTextCompare) = 0 Then
                        sTemp = RS.Fields("MAStreet").Value
                        If Trim(sTemp) = vbNullString Then
                            sTemp = "MAILING STREET UNAVAILABLE"
                        End If
                        oWddxIndemnityPaymentRS.setField 1, "PayeeLineThree", CleanXML(sTemp)
                    End If
                    sTemp = oWddxIndemnityPaymentRS.getField(1, "PayeeLineThree")
                    If Len(sTemp) > 37 Then
                        lCountLoop = 0
                        Do
                            sTemp = Left(sTemp, 37 - lCountLoop)
                            sTemp = CleanXML(sTemp)
                            lCountLoop = lCountLoop + 1
                        Loop Until Len(sTemp) <= 37
                        oWddxIndemnityPaymentRS.setField 1, "PayeeLineThree", sTemp
                    End If
                    If StrComp(oWddxIndemnityPaymentRS.getField(1, "PayeeLineFour"), "[ENTER_PayeeLineFour_MACity,MAState,MAZip]", vbTextCompare) = 0 Then
                        sTemp = RS.Fields("MACity").Value & ", " & RS.Fields("MAState").Value & ", " & RS.Fields("MAZIP").Value
                        oWddxIndemnityPaymentRS.setField 1, "PayeeLineFour", CleanXML(sTemp)
                    End If
                    'Create the XML TransForm to be passed to Biz Talk
                    Set oWddxStruct = New WDDXStruct
                    oWddxStruct.setProp "DocumentPropertiesRS", oWddxDocumentPropertiesRS
                    oWddxStruct.setProp "IndemnityPaymentRS", oWddxIndemnityPaymentRS
                    Set oWddxSer = New WDDXSerializer
                    sTemp = oWddxSer.serialize(oWddxStruct)
                    'Now Remove the pre xml file
                    goUtil.utDeleteFile (msBuildAssgnPackPath & sGUID)
                    'Now save this new one in its place
                    goUtil.utSaveFileData msBuildAssgnPackPath & sGUID, sTemp
                    'At the End Of All Processing The Pdf move first to Bix Talk Dir
                    'Then the xml
                End If
            Else
START_ALL:
                'When Creating a single pdf file only to be emailed
                'Need to be sure it is actually in the list of items to be
                'included in the email
                If mbCreateSinglePdfOnly Then
                    sCurrentPackageItemID = goUtil.IsNullIsVbNullString(RS.Fields("PackageItemID"))
                    'if the current packageitemid is found i the list then allow t to be processed
                    'for this email.  if not then skip to the next item in the package
                    If InStr(1, sPEQPackageItemIDList, sCurrentPackageItemID, vbBinaryCompare) = 0 Then
                        'If this item is not in the list then skip it
                        GoTo NEXT_RS:
                    End If
                End If
                bPrintActiveReport = PrintActiveReport(Nothing, , vbNullString, False, msBuildAssgnPackPath, sGUID, False, False, sReportFormat, msCarListClassName)
                'Check for Farmers XML01 only stuff
                If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 And Not mbCreateSinglePdfOnly Then
                    If bPrintActiveReport Then
                        sGUID = Replace(sGUID, ".pdf", ".xml", , , vbTextCompare)
                        'Create the XML TransForm to be passed to Biz Talk
                        Set oWddxStruct = New WDDXStruct
                        oWddxStruct.setProp "DocumentPropertiesRS", oWddxDocumentPropertiesRS
                        Set oWddxSer = New WDDXSerializer
                        sTemp = oWddxSer.serialize(oWddxStruct)
                        goUtil.utSaveFileData msBuildAssgnPackPath & sGUID, sTemp
                        'At the End Of All Processing The Pdf move first to Bix Talk Dir
                        'Then the xml
                    End If
                End If
            End If
            If bPrintActiveReport Then
                txtMess.Text = txtMess.Text & vbCrLf & "Delivering Document..."
                DoEvents
                Sleep 100
                'Move each doument one at a time to Biz talk Or Other Directory to be sent to client...
                'Check for Farmers XML01 only stuff
                If StrComp(sLossFormat, "V2ECcarFarmers.clsLossXML01", vbTextCompare) = 0 And Not mbCreateSinglePdfOnly Then
                    'First the PDF
                    'Copy Over the PDF First !
                    sGUID = Replace(sGUID, ".xml", ".pdf", , , vbTextCompare)
'DEBUG FARM
                    goUtil.utCopyFile msBuildAssgnPackPath & sGUID, msExportDocPath & "\" & sGUID
                    Sleep 1000 'wait or BizTalk pukes ' Can't go too fast for them pussies
                    'Copy over the XML Next
                    sGUID = Replace(sGUID, ".pdf", ".xml", , , vbTextCompare)
'DEBUG FARM
                    goUtil.utCopyFile msBuildAssgnPackPath & sGUID, msExportDocPath & "\" & sGUID
                    Sleep 1000
                    'also copy the xml file into the backup folder ....
                    'if it exisits
                    If goUtil.utFileExists(msExportDocPath & "\Backup", True) Then
                        goUtil.utCopyFile msBuildAssgnPackPath & sGUID, msExportDocPath & "\Backup\" & sGUID
                    End If
'Not Deleting Farmers Documents at this time after Delivery via B2B
'This will Allow for Backup of each package that was sent, or at least should
'Have Been sent via the B2B.  Items that were sent will be stored under SentB2B
'                    The Delete Both of the Originals
'                    sGUID = Replace(sGUID, ".xml", ".*", , , vbTextCompare)
'DEBUG FARM
'                    goUtil.utDeleteFile (msBuildAssgnPackPath & sGUID)
                End If
                
                If Not mbCreateSinglePdfOnly Then
                    'Now celebrate !! and Update!
                    sGUID = RS.Fields("PackageItemGUID")
                    sSQL = "UPDATE PackageItem SET "
                    sSQL = sSQL & "[SendMe] = 0, "
                    sSQL = sSQL & "[SentDate] = GetDate(), "
                    sSQL = sSQL & "[DownLoadMe] = 1, "
                    sSQL = sSQL & "[DateLastUpdated] = GetDate(), "
                    sSQL = sSQL & "[UpdateByUserID] = (SELECT [UsersID] FROM Users WHERE [UserName] = 'CFUSER') "
                    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
                    sSQL = sSQL & "AND   [PackageID] = " & msPackageID & " "
                    sSQL = sSQL & "AND  [SendMe] = 1 "
                    sSQL = sSQL & "AND [PackageItemGUID] = '" & sGUID & "' "
                    gConn.Execute sSQL, lRecordsAffected
                End If
            Else
                'Save Error Message
                If msExportDocPath <> vbNullString Then
                    goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & RS.Fields("Name").Value & "_" & sTickCount & "_Error.txt", sMess
                End If
                
            End If
        End If
NEXT_RS:
        RS.MoveNext
        ProgBarLoss.Value = lSequenceNumber
    Loop
            
    'After Createing All the Pdf Files Need To Zip them Up Delete the unziped
    If Not ZipUpFiles(RS, sTickCount) Then
        If mbCreateSinglePdfOnly Then
            If InStr(1, msPackageErrors, "Merge Documents Failed!", vbTextCompare) > 0 Then
                GoTo CLEANUP_QUEUE
            End If
        End If
        GoTo CLEAN_UP
    End If
    
    If mbCreateSinglePdfOnly Then
        bFail = False
        If Not EmailSinglePdfOnly(RS, sTickCount, sMess, msZipNameSinglePDFFileOnly, msSendSinglePdfOnlyQueue, sErrDesc) Then
            'Check for Certain Errors....
             'If its a timeout error then let this item remain in the queue
            'And it will be processed the next time around
            If InStr(1, msPackageErrors, "EmailSinglePdfOnly ERROR # 20008", vbTextCompare) = 0 Then
                'Let the Package Email Que be cleared of the Courrupt PDF messages that failed
                'To merge. Otherwise this will involve an infinite loop generating infinte error messages.
                If InStr(1, msPackageErrors, "Merge Documents Failed!", vbTextCompare) = 0 Then
                    bFail = True
                Else
                    bFail = False
                End If
            Else
                bFail = False
            End If
        End If
        
        If Not bFail Then
CLEANUP_QUEUE:
            sSQL = "spsGetPackageEmailQueue "
            sSQL = sSQL & msPackageEmailQueueID & ", " '@PackageEmailQueueID        int=null,
            sSQL = sSQL & "1 "                         '@bDelPackageEmailQueueID    bit=0
            
            gConn.Execute sSQL, lRecordsAffected
        End If
        
        DELIVERPackage = True
        GoTo CLEAN_UP
    End If
    
    
    'Sent to B2b
    If mbDoB2B And Not mbCreateSinglePdfOnly Then
        'Insert an Activity log Entry for this Action
        sSQL = " z_spuInsertActivityLogItem "
        'Use Null to for UpdateByUserID and CFUSER will be used by default
        sSQL = sSQL & "Null, "                                                  '@insUpdateByUserID  int=null,
        sSQL = sSQL & msAssignmentsID & ", "                                    '@insAssignmentsID   int,
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(msBackupB2BMess) & "', "    '@insActText     varchar(8000),
        sSQL = sSQL & "0, "                                                     '@insIsMgrEntry      bit=0,
        sSQL = sSQL & "'PACKAGE_UPDATE_B2B' "                                   '@insAdminComments   varchar(1000)=null,
        gConn.Execute sSQL, lRecordsAffected
    End If
    
    'Email Entire Claim
    If mbDoEmailEntireClaim And Not mbCreateSinglePdfOnly Then
        If Not MailPackageToClient(RS, sTickCount, msBackupEmailEntireClaimMess, msZipNameAll, msEmailEntireClaim, msSendEmailEntireClaim, msSentEmailEntireClaim, sErrDesc) Then
            'Insert an Activity log Entry for this Action
            msBackupEmailEntireClaimMess = msBackupEmailEntireClaimMess & vbCrLf & vbCrLf
            msBackupEmailEntireClaimMess = msBackupEmailEntireClaimMess & sErrDesc
            bFail = True
        End If
        sSQL = " z_spuInsertActivityLogItem "
        'Use Null to for UpdateByUserID and CFUSER will be used by default
        sSQL = sSQL & "Null, "                                                                  '@insUpdateByUserID  int=null,
        sSQL = sSQL & msAssignmentsID & ", "                                                    '@insAssignmentsID   int,
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(msBackupEmailEntireClaimMess) & "', "       '@insActText     varchar(8000),
        sSQL = sSQL & "0, "                                                                     '@insIsMgrEntry      bit=0,
        sSQL = sSQL & "'PACKAGE_UPDATE_EmailEntireClaim' "                                      '@insAdminComments   varchar(1000)=null,
        gConn.Execute sSQL, lRecordsAffected
        If bFail Then
            GoTo CLEAN_UP
        End If
    End If
    
    'Email Docs Only
    If mbDoEmailDocsOnly And Not mbCreateSinglePdfOnly Then
        If Not MailPackageToClient(RS, sTickCount, msBackupEmailDocsOnlyMess, msZipNameDocsOnly, msEmailDocsOnly, msSendEmailDocsOnly, msSentEmailDocsOnly, sErrDesc) Then
            'Insert an Activity log Entry for this Action
            msBackupEmailDocsOnlyMess = msBackupEmailDocsOnlyMess & vbCrLf & vbCrLf
            msBackupEmailDocsOnlyMess = msBackupEmailDocsOnlyMess & sErrDesc
            bFail = True
        End If
        sSQL = " z_spuInsertActivityLogItem "
        'Use Null to for UpdateByUserID and CFUSER will be used by default
        sSQL = sSQL & "Null, "                                                          '@insUpdateByUserID  int=null,
        sSQL = sSQL & msAssignmentsID & ", "                                            '@insAssignmentsID   int,
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(msBackupEmailDocsOnlyMess) & "', "  '@insActText     varchar(8000),
        sSQL = sSQL & "0, "                                                             '@insIsMgrEntry      bit=0,
        sSQL = sSQL & "'PACKAGE_UPDATE_EmailDocsOnly' "                                 '@insAdminComments   varchar(1000)=null,
        gConn.Execute sSQL, lRecordsAffected
        If bFail Then
            GoTo CLEAN_UP
        End If
    End If
    
    'Email Photos Only
    If mbDoEmailPhotosOnly And Not mbCreateSinglePdfOnly Then
        If Not MailPackageToClient(RS, sTickCount, msBackupEmailPhotosOnlyMess, msZipNamePhotosOnly, msEmailPhotosOnly, msSendEmailPhotosOnly, msSentEmailPhotosOnly, sErrDesc) Then
            'Insert an Activity log Entry for this Action
            msBackupEmailPhotosOnlyMess = msBackupEmailPhotosOnlyMess & vbCrLf & vbCrLf
            msBackupEmailPhotosOnlyMess = msBackupEmailPhotosOnlyMess & sErrDesc
            bFail = True
        End If
         sSQL = " z_spuInsertActivityLogItem "
        'Use Null to for UpdateByUserID and CFUSER will be used by default
        sSQL = sSQL & "Null, "                                                              '@insUpdateByUserID  int=null,
        sSQL = sSQL & msAssignmentsID & ", "                                                '@insAssignmentsID   int,
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(msBackupEmailPhotosOnlyMess) & "', "    '@insActText     varchar(8000),
        sSQL = sSQL & "0, "                                                                 '@insIsMgrEntry      bit=0,
        sSQL = sSQL & "'PACKAGE_UPDATE_EmailPhotosOnly' "                                   '@insAdminComments   varchar(1000)=null,
        gConn.Execute sSQL, lRecordsAffected
        If bFail Then
            GoTo CLEAN_UP
        End If
    End If
    
    'Send To FTP
    If mbDoFTP And Not mbCreateSinglePdfOnly Then
        If Not FTPPackageToClient(RS, sTickCount, msZipNameAll, msSendToFTP, msSendFTP, msSentFTP, sErrDesc) Then
            'Insert an Activity log Entry for this Action
            msBackupFTPMess = msBackupFTPMess & vbCrLf & vbCrLf
            msBackupFTPMess = msBackupFTPMess & sErrDesc
            bFail = True
        End If
        sSQL = " z_spuInsertActivityLogItem "
        'Use Null to for UpdateByUserID and CFUSER will be used by default
        sSQL = sSQL & "Null, "                                                  '@insUpdateByUserID  int=null,
        sSQL = sSQL & msAssignmentsID & ", "                                    '@insAssignmentsID   int,
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(msBackupFTPMess) & "', "    '@insActText     varchar(8000),
        sSQL = sSQL & "0, "                                                     '@insIsMgrEntry      bit=0,
        sSQL = sSQL & "'PACKAGE_UPDATE_FTP' "                                   '@insAdminComments   varchar(1000)=null,
        gConn.Execute sSQL, lRecordsAffected
        If bFail Then
            GoTo CLEAN_UP
        End If
    End If
        
    DELIVERPackage = True
    
CLEAN_UP:
    Set RS = Nothing
    Set RSEmail = Nothing
    Set oWddxDeser = Nothing
    Set oWddxSer = Nothing
    Set oWddxStruct = Nothing
    Set oWddxLossRS = Nothing
    Set oWddxAssignmentDetailRS = Nothing
    Set oWddxLossDetailRS = Nothing
    Set oWddxVehicleDetailRS = Nothing
    Set oWddxDocumentPropertiesRS = Nothing
    Set oWddxIndemnityPaymentRS = Nothing
    Set oWddxFeeBillPaymentRS = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    DELIVERPackage = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function DELIVERPackage" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    If msExportDocPath <> vbNullString Then
        If Not RS Is Nothing Then
            If RS.State = ADODB.adStateClosed Then
                goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & "RSCLosed_" & sTickCount & "_Error.txt", sMess
            Else
                If Not RS.EOF Then
                    goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & RS.Fields("Name").Value & "_" & sTickCount & "_Error.txt", sMess
                Else
                    goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & "RSEOF_" & sTickCount & "_Error.txt", sMess
                End If
            End If
        Else
            goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & "_" & sTickCount & "_Error.txt", sMess
        End If
    End If
    moUL_ErrorMess sMess
End Function

Private Function ZipUpFiles(pPackageRS As ADODB.Recordset, psTickcount As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sCOCatPassWord As String
    Dim sEncryptPassWord As String
    Dim sZipname As String
    Dim sZipNameAll As String
    Dim sZipNamePhotosOnly As String
    Dim sZipNameSinglePdfOnly As String
    Dim sNameSinglePdfOnly As String
    Dim sZipNameDocsOnly As String
    Dim sCOCode As String
    Dim sAdjUserName As String
    Dim sACIDDisplay As String
    Dim sClaimNo As String
    Dim sIBNo As String
    Dim sCatName As String
    Dim sYYYYMMDDHHMMSS As String
    'Send Queues / Sent Directories
    'B2B will only require a Sent Directory
    'Since the files being zipped were already Copied to
    'a B2B processing directory and sent
    Dim sSentB2B As String
    Dim bDoB2B As Boolean
    Dim sSendFTP As String
    Dim sSentFTP As String
    Dim bDoFTP As Boolean
    Dim sSendEmailEntireClaim As String
    Dim sSentEmailEntireClaim As String
    Dim bDoEmailEntireClaim As Boolean
    Dim sSendEmailDocsOnly As String
    Dim sSentEmailDocsOnly As String
    Dim bDoEmailDocsOnly As Boolean
    Dim sSendEmailPhotosOnly As String
    Dim sSentEmailPhotosOnly As String
    Dim bDoEmailPhotosOnly As Boolean
    'Set this to true if already Saved the file to one of the above directories
    Dim sTempStore As String
    Dim bSavedToSend As Boolean
    Dim sTemp As String
    'Single PDF
    Dim sSendSinglePdfOnly As String
    Dim sbDisableZipFilePassword As String
    Dim sOverRideZipFilePassword As String
    Dim sInsured As String
    Dim sInsTemp As String
    
    'First determine what method of sending and create the
    'appropriate directories
    
    If mbCreateSinglePdfOnly Then
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & "Creating Single PDF File. " & Now() & vbCrLf
        sMess = sMess & "To: " & msEmailSinglePdfOnlyTo & vbCrLf
        sMess = sMess & "CC: " & msEmailSinglePdfOnlyCC & vbCrLf
        sMess = sMess & "BCC: " & msEmailSinglePdfOnlyBCC & vbCrLf
        sMess = sMess & "Subject: " & msEmailSinglePdfOnlySubject & vbCrLf
        sMess = sMess & "Body: " & msEmailSinglePdfOnlyBody & vbCrLf
        txtMess.Text = sMess
        DoEvents
        Sleep 100
        'Build both Sent and Send directories
        sSendSinglePdfOnly = msBuildAssgnPackPath & "SendSinglePdfOnly"
        If Not goUtil.utFileExists(sSendSinglePdfOnly, True) Then
            goUtil.utMakeDir sSendSinglePdfOnly
        End If
        'there is no backup for single pdf files sent manually by a user
        GoTo BUILD_ME
    End If
    
    'B2B
    If goUtil.utFileExists(msExportDocPath, True) Then
        bDoB2B = True
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & "Creating Backup Package of files processed via B2B transaction. " & Now() & vbCrLf
        sMess = sMess & "B2B Process Directory: " & msExportDocPath
        txtMess.Text = sMess
        DoEvents
        Sleep 100
        'Build the Sent directory only
        sSentB2B = msBuildAssgnPackPath & "SentB2B"
        If Not goUtil.utFileExists(sSentB2B, True) Then
            goUtil.utMakeDir sSentB2B
        End If
    End If
    
    'FTP
    If Trim(msSendToFTP) <> vbNullString Then
        bDoFTP = True
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & "Creating Package to be sent via ftp. " & Now() & vbCrLf
        sMess = sMess & "FTP Delivery URL: " & msSendToFTP
        txtMess.Text = sMess
        DoEvents
        Sleep 100
        'Build both Sent and Send directories
        sSendFTP = msBuildAssgnPackPath & "SendFTP"
        If Not goUtil.utFileExists(sSendFTP, True) Then
            goUtil.utMakeDir sSendFTP
        End If
        sSentFTP = msBuildAssgnPackPath & "SentFTP"
        If Not goUtil.utFileExists(sSentFTP, True) Then
            goUtil.utMakeDir sSentFTP
        End If
    End If

    'Email of Entire Claim to one Email Address
    If Trim(msEmailEntireClaim) <> vbNullString Then
        bDoEmailEntireClaim = True
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & "Creating Package of Entire Claim to be sent via Email. " & Now() & vbCrLf
        sMess = sMess & "Email Delivery: " & msEmailEntireClaim
        txtMess.Text = sMess
        DoEvents
        Sleep 100
        'Build both Sent and Send directories
        sSendEmailEntireClaim = msBuildAssgnPackPath & "SendEmailEntireClaim"
        If Not goUtil.utFileExists(sSendEmailEntireClaim, True) Then
            goUtil.utMakeDir sSendEmailEntireClaim
        End If
        sSentEmailEntireClaim = msBuildAssgnPackPath & "SentEmailEntireClaim"
        If Not goUtil.utFileExists(sSentEmailEntireClaim, True) Then
            goUtil.utMakeDir sSentEmailEntireClaim
        End If
    End If

    'Email of Documents Only (No Photos)
    If Trim(msEmailDocsOnly) <> vbNullString Then
        bDoEmailDocsOnly = True
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & "Creating Package of Documents ONLY to be sent via Email. " & Now() & vbCrLf
        sMess = sMess & "Email Delivery: " & msEmailDocsOnly
        txtMess.Text = sMess
        DoEvents
        Sleep 100
        'Build both Sent and Send directories
        sSendEmailDocsOnly = msBuildAssgnPackPath & "SendEmailDocsOnly"
        If Not goUtil.utFileExists(sSendEmailDocsOnly, True) Then
            goUtil.utMakeDir sSendEmailDocsOnly
        End If
        sSentEmailDocsOnly = msBuildAssgnPackPath & "SentEmailDocsOnly"
        If Not goUtil.utFileExists(sSentEmailDocsOnly, True) Then
            goUtil.utMakeDir sSentEmailDocsOnly
        End If
    End If
    
    'Email of Photos Only (No Documents)
    If Trim(msEmailPhotosOnly) <> vbNullString Then
        bDoEmailPhotosOnly = True
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & "Creating Package of Photos ONLY to be sent via Email. " & Now() & vbCrLf
        sMess = sMess & "Email Delivery: " & msEmailPhotosOnly
        txtMess.Text = sMess
        DoEvents
        Sleep 100
        'Build both Sent and Send directories
        sSendEmailPhotosOnly = msBuildAssgnPackPath & "SendEmailPhotosOnly"
        If Not goUtil.utFileExists(sSendEmailPhotosOnly, True) Then
            goUtil.utMakeDir sSendEmailPhotosOnly
        End If
        sSentEmailPhotosOnly = msBuildAssgnPackPath & "SentEmailPhotosOnly"
        If Not goUtil.utFileExists(sSentEmailPhotosOnly, True) Then
            goUtil.utMakeDir sSentEmailPhotosOnly
        End If
    End If
    
BUILD_ME:
    'Need to build a Temp Directory that will hold all the files as they
    'now appear.  This will allow the ability to create the package in different
    'Ways depending on the Method(s) that are flagged to be sent to the Client.
    If Not mbCreateSinglePdfOnly Then
        sTempStore = msBuildAssgnPackPath & psTickcount & "_TempStore"
        If Not goUtil.utFileExists(sTempStore, True) Then
            goUtil.utMakeDir sTempStore
        End If
        
        'Now Copy over all the files in this directory to the Temp Store
        goUtil.utCopyFile msBuildAssgnPackPath & "*.*", sTempStore
    End If
    
    'I. Build the Zip File name
    '1. Get the Parts of the Zip Name
    pPackageRS.MoveFirst
    sCOCode = goUtil.IsNullIsVbNullString(pPackageRS.Fields("CoCode"))
    sCatName = goUtil.IsNullIsVbNullString(pPackageRS.Fields("CatName"))
    sAdjUserName = goUtil.IsNullIsVbNullString(pPackageRS.Fields("AdjUserName"))
    sACIDDisplay = goUtil.IsNullIsVbNullString(pPackageRS.Fields("ACIDDisplay"))
    sClaimNo = goUtil.IsNullIsVbNullString(pPackageRS.Fields("CLIENTNUM"))
    sIBNo = goUtil.IsNullIsVbNullString(pPackageRS.Fields("IBNUM"))
    sInsured = goUtil.IsNullIsVbNullString(pPackageRS.Fields("Insured"))
    sInsured = Trim(sInsured)
    
    'Need to get the Last Name of the Insured
    If InStr(1, sInsured, Chr(32), vbBinaryCompare) > 0 Then
        sInsTemp = sInsured
        sInsured = StrReverse(sInsured)
        sInsured = Left(sInsured, InStr(1, sInsured, Chr(32), vbBinaryCompare))
        sInsured = Trim(StrReverse(sInsured))
        'If the last name is really some sort of a thing like
        'Jr, Sr or some other dorky thing, need to go back one more
        'Word in the Insured name
        If Len(sInsured) <= 3 Then
            sInsTemp = Replace(sInsTemp, sInsured, vbNullString, , 1, vbBinaryCompare)
            sInsTemp = Trim(sInsTemp)
            If InStr(1, sInsTemp, Chr(32), vbBinaryCompare) > 0 Then
                sInsTemp = StrReverse(sInsTemp)
                sInsTemp = Left(sInsTemp, InStr(1, sInsTemp, Chr(32), vbBinaryCompare))
                sInsTemp = Trim(StrReverse(sInsTemp))
                sInsured = sInsTemp & Chr(32) & sInsured
            End If
        End If
    End If
    
    'Can't be greater than
    sInsured = Left(sInsured, 20)
    
    '2. Set the Year month Day Hour Min Sec String
    sYYYYMMDDHHMMSS = Format(Now(), "YYYY_MMDD_HHmmss")
    
    '3. Put them togetha
    If Not mbCreateSinglePdfOnly Then
        sZipname = sClaimNo & "_" & sInsured & "_"
        sZipname = sZipname & sCOCode & "_" & sCatName & "_{" & sYYYYMMDDHHMMSS & "}_"
        sZipname = sZipname & sACIDDisplay & "_"
        sZipname = sZipname & sAdjUserName & "_"
        sZipname = sZipname & sIBNo & "_"
    Else
        sZipname = sClaimNo & "_" & sInsured
    End If
    sZipNameAll = sZipname & "ALL.zip"
    sZipNameDocsOnly = sZipname & "Docs ONLY.zip"
    sZipNamePhotosOnly = sZipname & "Photos ONLY.zip"
    sZipNameSinglePdfOnly = sZipname & ".zip"
    sNameSinglePdfOnly = sZipname & ".pdf"
    
    goUtil.utCleanFileFolderName sZipNameAll
    goUtil.utCleanFileFolderName sZipNameDocsOnly
    goUtil.utCleanFileFolderName sZipNamePhotosOnly
    goUtil.utCleanFileFolderName sZipNameSinglePdfOnly
    
    'II. Create the Company Cat Password
    sCOCatPassWord = goUtil.IsNullIsVbNullString(pPackageRS.Fields("CoName"))
    sCOCatPassWord = sCOCatPassWord & "_" & goUtil.IsNullIsVbNullString(pPackageRS.Fields("CatName"))
    'make sure it is Lower Case
    sCOCatPassWord = LCase(sCOCatPassWord)
    
    
    'III. Save to Zip File
    sEncryptPassWord = goUtil.Encode(sCOCatPassWord)
    
    sbDisableZipFilePassword = GetSetting("V2WebControl", "SMTP", "DISABLE_ZIP_FILE_PASSWORD", vbNullString)
    sOverRideZipFilePassword = GetSetting("V2WebControl", "SMTP", "OVERRIDE_ZIP_FILE_PASSWORD", vbNullString)
    
    If StrComp(sbDisableZipFilePassword, "True", vbTextCompare) = 0 Then
        sEncryptPassWord = vbNullString
    ElseIf Trim(sOverRideZipFilePassword) <> vbNullString Then
        sEncryptPassWord = goUtil.Encode(sOverRideZipFilePassword)
    End If
    
    
    If mbCreateSinglePdfOnly Then
        'If Copy to self = file not found then no files were generated Skip creating a zip file.
        sTemp = goUtil.utCopyFile(msBuildAssgnPackPath & "*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            'First need to Merge multiple pdf files into a single Pdf file.
            'Create the SinglePdfOutPut folder exists
            msSendSinglePDFFileOnly = msBuildAssgnPackPath & "SinglePdfOutPut"
            If Not goUtil.utFileExists(msSendSinglePDFFileOnly, True) Then
                goUtil.utMakeDir msSendSinglePDFFileOnly
            End If
            'Now merge them
            If Not MergePDFFiles(msBuildAssgnPackPath, msSendSinglePDFFileOnly, sNameSinglePdfOnly) Then
                bSavedToSend = False
                sMess = txtMess.Text & vbCrLf & vbCrLf
                sMess = sMess & "Merge Documents Failed! " & Now() & vbCrLf
                sMess = sMess & "Email not sent"
                moUL_ErrorMess sMess
                txtMess.Text = sMess
            Else
                bSavedToSend = SaveToZipFile(sSendSinglePdfOnly, sZipNameSinglePdfOnly, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
                sMess = txtMess.Text & vbCrLf & vbCrLf
                sMess = sMess & sZipNameSinglePdfOnly & vbCrLf & "Single PDF File Created! " & Now() & vbCrLf
                txtMess.Text = sMess
            End If
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "Email not sent"
            moUL_ErrorMess sMess
            txtMess.Text = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
        GoTo BAIL_HERE
    End If
    
    '"*.*"
    'First Do the B2B
    If bDoB2B Then
        'If Copy to self = file not found then no files were generated Skip creating a zip file.
        sTemp = goUtil.utCopyFile(msBuildAssgnPackPath & "*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            bSavedToSend = SaveToZipFile(sSentB2B, sZipNameAll, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & sZipNameAll & vbCrLf & " B2B Backup Saved! " & Now() & vbCrLf
            txtMess.Text = sMess
            msBackupB2BMess = sMess
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "B2B Backup not Saved!"
            txtMess.Text = sMess
            msBackupB2BMess = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
    End If
    
    'FTP
    If bDoFTP And Not bSavedToSend Then
        sTemp = goUtil.utCopyFile(msBuildAssgnPackPath & "*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            bSavedToSend = SaveToZipFile(sSendFTP, sZipNameAll, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & sZipNameAll & vbCrLf & " FTP Backup Saved! " & Now() & vbCrLf
            txtMess.Text = sMess
            msBackupFTPMess = sMess
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "FTP Backup not Saved!"
            txtMess.Text = sMess
            msBackupFTPMess = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
    ElseIf bDoFTP And bSavedToSend And bDoB2B Then
        'if the file was already Saved to the B2B Directory...
        'Just copy it to the FTP one as well.  This is lots faster
        sTemp = goUtil.utCopyFile(sSentB2B & "\" & sZipNameAll, sSendFTP & "\" & sZipNameAll)
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & sZipNameAll & vbCrLf & " FTP Backup Saved! " & Now() & vbCrLf
        txtMess.Text = sMess
        msBackupFTPMess = sMess
    ElseIf bDoFTP Then
        sTemp = sSentB2B
        'make sure the current Build directory is reset
        goUtil.utDeleteFile msBuildAssgnPackPath & "*.*"
        'Then Reset the Build Files
        goUtil.utCopyFile sTempStore & "\*.*", msBuildAssgnPackPath
         'If Copy to self = file not found then no files were generated Skip creating a zip file.
        sTemp = goUtil.utCopyFile(msBuildAssgnPackPath & "*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            bSavedToSend = SaveToZipFile(sTemp, sZipNameAll, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
            sTemp = goUtil.utCopyFile(sSentB2B & "\" & sZipNameAll, sSendFTP & "\" & sZipNameAll)
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & sZipNameAll & vbCrLf & " FTP Backup Saved! " & Now() & vbCrLf
            txtMess.Text = sMess
            msBackupFTPMess = sMess
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "FTP Backup not Saved!"
            txtMess.Text = sMess
            msBackupFTPMess = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
    End If
    
    'Email Entire Claim
    If bDoEmailEntireClaim And Not bSavedToSend Then
        sTemp = goUtil.utCopyFile(msBuildAssgnPackPath & "*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            bSavedToSend = SaveToZipFile(sSendEmailEntireClaim, sZipNameAll, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & sZipNameAll & vbCrLf & " Email Entire Claim Backup Saved! " & Now() & vbCrLf
            txtMess.Text = sMess
            msBackupEmailEntireClaimMess = sMess
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "Email Entire Claim Backup not Saved!"
            txtMess.Text = sMess
            msBackupEmailEntireClaimMess = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
    ElseIf bDoEmailEntireClaim And bSavedToSend And (bDoB2B Or bDoFTP) Then
        If bDoB2B Then
            sTemp = sSentB2B
        ElseIf bDoFTP Then
            sTemp = sSendFTP
        End If
        'if the file was already Saved to the B2B Directory...
        'Just copy it to the FTP one as well.  This is lots faster
        sTemp = goUtil.utCopyFile(sTemp & "\" & sZipNameAll, sSendEmailEntireClaim & "\" & sZipNameAll)
        sMess = txtMess.Text & vbCrLf & vbCrLf
        sMess = sMess & sZipNameAll & vbCrLf & " Email Entire Claim Backup Saved! " & Now() & vbCrLf
        txtMess.Text = sMess
        msBackupEmailEntireClaimMess = sMess
    ElseIf bDoEmailEntireClaim Then
        'make sure the current Build directory is reset
        goUtil.utDeleteFile msBuildAssgnPackPath & "*.*"
        'Then Reset the Build Files
        sTemp = goUtil.utCopyFile(sTempStore & "\*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            bSavedToSend = SaveToZipFile(sSendEmailEntireClaim, sZipNameAll, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & sZipNameAll & vbCrLf & " Email Entire Claim Backup Saved! " & Now() & vbCrLf
            txtMess.Text = sMess
            msBackupEmailEntireClaimMess = sMess
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "Email Entire Claim Backup not Saved!"
            txtMess.Text = sMess
            msBackupEmailEntireClaimMess = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
    End If
    
    'The Photos and Docs Only need to always reset the Build Dir and Only
    'package Up either Docs or Photos
    If bDoEmailDocsOnly Then
        'make sure the current Build directory is reset
        goUtil.utDeleteFile msBuildAssgnPackPath & "*.*"
        'Then Reset the Build Files
        goUtil.utCopyFile sTempStore & "\*.*", msBuildAssgnPackPath
        'Then Remove ALL the Photos
        goUtil.utDeleteFile msBuildAssgnPackPath & "*Photo*.*"
        'If Copy to self = file not found then no files were generated Skip creating a zip file.
        sTemp = goUtil.utCopyFile(msBuildAssgnPackPath & "*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            bSavedToSend = SaveToZipFile(sSendEmailDocsOnly, sZipNameDocsOnly, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & sZipNameDocsOnly & vbCrLf & " Email Docs Only Backup Saved! " & Now() & vbCrLf
            txtMess.Text = sMess
            msBackupEmailDocsOnlyMess = sMess
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "Email Docs Only Backup not Saved!"
            txtMess.Text = sMess
            msBackupEmailDocsOnlyMess = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
    End If
    If bDoEmailPhotosOnly Then
        'make sure the current Build directory is reset
        goUtil.utDeleteFile msBuildAssgnPackPath & "*.*"
        'Then Reset the Build Files
        goUtil.utCopyFile sTempStore & "\*Photo*.*", msBuildAssgnPackPath
        'If Copy to self = file not found then no files were generated Skip creating a zip file.
        sTemp = goUtil.utCopyFile(msBuildAssgnPackPath & "*.*", msBuildAssgnPackPath)
        If sTemp = vbNullString Then
            bSavedToSend = SaveToZipFile(sSendEmailPhotosOnly, sZipNamePhotosOnly, "*.*", sEncryptPassWord, pPackageRS, psTickcount)
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & sZipNamePhotosOnly & vbCrLf & " Email Photos Only Backup Saved! " & Now() & vbCrLf
            txtMess.Text = sMess
            msBackupEmailPhotosOnlyMess = sMess
        Else
            sMess = txtMess.Text & vbCrLf & vbCrLf
            sMess = sMess & "No Documents to deliver! " & Now() & vbCrLf
            sMess = sMess & "Email Photos Only Backup not Saved!"
            txtMess.Text = sMess
            msBackupEmailPhotosOnlyMess = sMess
            DoEvents
            Sleep 100
            bSavedToSend = True
        End If
    End If
    
    'once everything is saved get rid of the TempStore
    goUtil.utDeleteDir sTempStore
    
    'Set the member vars so that the Email And Ftp processes
    'Will know Who What Where Why And how
    msZipNameAll = sZipNameAll
    msZipNameDocsOnly = sZipNameDocsOnly
    msZipNamePhotosOnly = sZipNamePhotosOnly
    mbDoB2B = bDoB2B
    msSentB2B = sSentB2B
    msSendFTP = sSendFTP
    msSentFTP = sSentFTP
    mbDoFTP = bDoFTP
    msSendEmailEntireClaim = sSendEmailEntireClaim
    msSentEmailEntireClaim = sSentEmailEntireClaim
    mbDoEmailEntireClaim = bDoEmailEntireClaim
    msSendEmailDocsOnly = sSendEmailDocsOnly
    msSentEmailDocsOnly = sSentEmailDocsOnly
    mbDoEmailDocsOnly = bDoEmailDocsOnly
    msSendEmailPhotosOnly = sSendEmailPhotosOnly
    msSentEmailPhotosOnly = sSentEmailPhotosOnly
    mbDoEmailPhotosOnly = bDoEmailPhotosOnly
    
BAIL_HERE:
    If mbCreateSinglePdfOnly Then
        msZipNameSinglePDFFileOnly = sZipNameSinglePdfOnly
        msSendSinglePdfOnlyQueue = sSendSinglePdfOnly
        goUtil.utDeleteFile msBuildAssgnPackPath & "*.*"
        goUtil.utDeleteDir msSendSinglePDFFileOnly
    End If
    ZipUpFiles = bSavedToSend
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function ZipUpFiles" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "Problems Creating Zip File! " & sZipname & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    'Save Error Message
    If msExportDocPath <> vbNullString Then
        goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & pPackageRS.Fields("Name").Value & "_" & psTickcount & "_Error.txt", sMess
    End If
    
    moUL_ErrorMess sMess
End Function

Private Function SaveToZipFile(psSendToDir As String, _
                                psZipName As String, _
                                psFilter As String, _
                                psEncryptPassWord As String, _
                                pPackageRS As ADODB.Recordset, _
                                psTickcount As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    'Zip Object
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim lCheckCount As Long

    Set oXZip = New V2ECKeyBoard.clsXZip
    oXZip.SetUtilObject goUtil
    
    'Because Zip Utility will create a temporary file under the Temp directory
    'when it is generating a zip file under the specified directory...
    'It is possible that Other Applications such as multiple instances of Process User Tokin will try to
    'Create a temp file with the same file name.  When this happens this
    'function will return an error , 504 cannot create temporary file.
    'This error can be used to indicate to the process receiving the error that it needs
    'to sleep, wait for the other process to finish up.  This will also help to
    'govern the process load on the server.
    lCheckCount = 0
CREATE_ZIP:
    On Error Resume Next
    'Be sure to Free up windows Tasks first !
    DoEvents
    Sleep 100
    If mbCreateSinglePdfOnly Then
        oXZip.SaveZIPFiles msSendSinglePDFFileOnly, psZipName, psFilter, psEncryptPassWord, psSendToDir
    Else
        oXZip.SaveZIPFiles msBuildAssgnPackPath, psZipName, psFilter, psEncryptPassWord, psSendToDir
    End If
    
    lErrNum = Err.Number
    If lErrNum <> 0 Then
        'if the error is a skipping event that means the zip file being created
        'is being created in the same directory as were all the files to be zipped
        'this error is ok, and can be ignored
        If lErrNum = 527 Then
            GoTo SUCCESS_ZIP
        End If
        lCheckCount = lCheckCount + 1
        'wait for 5 seconds total per file
        If lCheckCount <= 10 Then
            DoEvents
            Sleep 500
            GoTo CREATE_ZIP
        Else
            On Error GoTo EH
            Err.Raise Err.Number, , Err.Description
        End If
    End If
    
SUCCESS_ZIP:
    SaveToZipFile = True
    On Error GoTo EH
    SetAttr psSendToDir & "\" & psZipName, vbNormal
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function SaveToZipFile" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "Problems Creating Zip File! " & psZipName & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    'Save Error Message
    If msExportDocPath <> vbNullString Then
        goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & pPackageRS.Fields("Name").Value & "_" & psTickcount & "_Error.txt", sMess
    End If
    
    moUL_ErrorMess sMess
    Set oXZip = Nothing
End Function

Private Function MailPackageToClient(pPackageRS As ADODB.Recordset, _
                                        psTickcount As String, _
                                        psMessText As String, _
                                        psAttachName As String, _
                                        psEmailTo As String, _
                                        psSendEmailQueue As String, _
                                        psSentEmailBackup As String, _
                                        psErrorMess As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sTemp As String
    Dim sMessText As String
    Dim sAttachName As String
    Dim sEmailTo As String
    Dim sSendEmailQueue As String
    Dim sSentEmailBackup As String
    'SMTP Stuff
    Dim sSMTPHost As String
    Dim sSMTPUserName As String
    Dim sSMTPPassword As String
    'Attachments
    Dim boundary As Double
    
    'get SMTP Values
    sSMTPHost = GetECSCryptSetting("V2WebControl", "SMTP", "Host", vbNullString)
    sSMTPUserName = GetECSCryptSetting("V2WebControl", "SMTP", "UserName", vbNullString)
    sSMTPPassword = GetECSCryptSetting("V2WebControl", "SMTP", "Password", vbNullString)
    
    'Set Email variables
    sMessText = psMessText
    sAttachName = psAttachName
    sEmailTo = psEmailTo
    sSendEmailQueue = psSendEmailQueue
    sSentEmailBackup = psSentEmailBackup
    
SEND_EMAIL:
    'If the Attachment does not Exisit... that means there were no document
    'generated to Send Exit this process but also be sure to True this Function
    If Not goUtil.utFileExists(sSendEmailQueue & "\" & sAttachName) Then
        GoTo BAIL_HERE
    End If

    'Attach the File
    With PP_MAIL
        .Blocking = True
        .Debug = 1
        boundary = Fix(Rnd * 100000000000#)
        .ContentType = "multipart"
        .ContentSubtype = "mixed"
        .ContentSubtypeParameters = "boundary=" & CStr(boundary) & "_boundary"
        .MultipartBoundary = CStr(boundary) & "_boundary"
        .Action = MailActionCreatePart
        .Action = MailActionDescend
        .ContentTransferEncoding = "base64"
        .ContentType = "application"
        .ContentSubtype = "x-zip-compressed"
        .ContentSubtypeParameters = "name=" & Chr(34) & sSendEmailQueue & "\" & sAttachName & Chr(34)
        .ContentDisposition = "attachment; filename=" & Chr(34) & sAttachName & Chr(34)
        .Flags = MailSrcIsFile Or MailDstIsBody
        .SrcFilename = sSendEmailQueue & "\" & sAttachName
        .Action = MailActionEncode
        .Action = MailActionAscend
        .To = sEmailTo
        .Subject = "Claim Package " & Now() & " " & sAttachName
        .From = sSMTPUserName & "@eberls.com"
        .CC = vbNullString
        .BCC = vbNullString
        .Headers(PP_MAIL.HeadersCount) = "X-Mailer: Mabry"
        .Host = sSMTPHost
        .EmailAddress = sSMTPUserName & "@eberls.com"
        .MessageID = Year(Now) & Month(Now) & Day(Now) & Fix(Timer) & "_MabryMail"
        .Body(0) = psMessText
        .LogonName = sSMTPUserName
        .LogonPassword = sSMTPPassword
        .ConnectType = MailConnectTypeESMTP
        .Timeout = 60
        mlMailState = MAIL_STATE_CONNECTING
        .Action = MailActionConnect
        PP_MAIL_Done
    End With
    
    'Once the message has been sent... Need to Move the Attachment over to the Sent dir
    sTemp = goUtil.utCopyFile(sSendEmailQueue & "\" & sAttachName, sSentEmailBackup & "\" & sAttachName)
    'then Delete the File from the Send Queue
    sTemp = goUtil.utDeleteFile(sSendEmailQueue & "\" & sAttachName)
    
BAIL_HERE:
    
    'Check for Other attchments that were suppose to be sent but got stuck in queue
    sTemp = Dir(sSendEmailQueue & "\*.*", vbHidden)
    
    If sTemp <> vbNullString Then
        sAttachName = sTemp
        sTemp = sMessText & vbCrLf & vbCrLf
        sTemp = sTemp & String(100, "-") & vbCrLf
        sTemp = sTemp & "Sent Attachment Item: " & sAttachName & vbCrLf
        sMessText = sTemp
        GoTo SEND_EMAIL
    End If
    
    MailPackageToClient = True
    psMessText = sMessText
    
    Exit Function
EH:
    
ERROR_MESS:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function MailPackageToClient" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "Could Not Send Message" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    'Save Error Message
    If msBuildAssgnPackPath <> vbNullString Then
        goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\MailPackageToClient_" & psTickcount & "_Error.txt", sMess
    End If
    
    moUL_ErrorMess sMess
    psErrorMess = sMess
    
End Function

Private Function EmailSinglePdfOnly(pPackageRS As ADODB.Recordset, _
                                        psTickcount As String, _
                                        psMessText As String, _
                                        psAttachName As String, _
                                        psSendEmailQueue As String, _
                                        psErrorMess As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sTemp As String
    Dim sAttachName As String
    Dim sSendEmailQueue As String
    'SMTP Stuff
    Dim sSMTPHost As String
    Dim sSMTPUserName As String
    Dim sSMTPPassword As String
    'Attachments
    Dim boundary As Double
    
    'get SMTP Values
    sSMTPHost = GetECSCryptSetting("V2WebControl", "SMTP", "Host", vbNullString)
    sSMTPUserName = GetECSCryptSetting("V2WebControl", "SMTP", "UserName", vbNullString)
    sSMTPPassword = GetECSCryptSetting("V2WebControl", "SMTP", "Password", vbNullString)
    
    'Set Email variables
    sAttachName = psAttachName
    sSendEmailQueue = psSendEmailQueue
    
SEND_EMAIL:
    'If the Attachment does not Exisit... that means there were no document
    'generated to Send Exit this process but also be sure to True this Function
    If Not goUtil.utFileExists(sSendEmailQueue & "\" & sAttachName) Then
        GoTo BAIL_HERE
    End If

    'Attach the File
    With PP_MAIL
        .Blocking = True
        .Debug = 1
        boundary = Fix(Rnd * 100000000000#)
        .ContentType = "multipart"
        .ContentSubtype = "mixed"
        .ContentSubtypeParameters = "boundary=" & CStr(boundary) & "_boundary"
        .MultipartBoundary = CStr(boundary) & "_boundary"
        .Action = MailActionCreatePart
        .Action = MailActionDescend
        .ContentTransferEncoding = "base64"
        .ContentType = "application"
        .ContentSubtype = "x-zip-compressed"
        .ContentSubtypeParameters = "name=" & Chr(34) & sSendEmailQueue & "\" & sAttachName & Chr(34)
        .ContentDisposition = "attachment; filename=" & Chr(34) & sAttachName & Chr(34)
        .Flags = MailSrcIsFile Or MailDstIsBody
        .SrcFilename = sSendEmailQueue & "\" & sAttachName
        .Action = MailActionEncode
        .Action = MailActionAscend
        .To = msEmailSinglePdfOnlyTo
        .CC = msEmailSinglePdfOnlyCC
        .BCC = msEmailSinglePdfOnlyBCC
        .Subject = msEmailSinglePdfOnlySubject
        .Body(0) = msEmailSinglePdfOnlyBody
        .From = sSMTPUserName & "@eberls.com"
        .Headers(PP_MAIL.HeadersCount) = "X-Mailer: Mabry"
        .Host = sSMTPHost
        .EmailAddress = sSMTPUserName & "@eberls.com"
        .MessageID = Year(Now) & Month(Now) & Day(Now) & Fix(Timer) & "_MabryMail"
        .LogonName = sSMTPUserName
        .LogonPassword = sSMTPPassword
        .ConnectType = MailConnectTypeESMTP
        .Timeout = 60
        mlMailState = MAIL_STATE_CONNECTING
        .Action = MailActionConnect
        PP_MAIL_Done
    End With
    
    sTemp = goUtil.utDeleteFile(sSendEmailQueue & "\" & sAttachName)
    
BAIL_HERE:
    
    'Check for Other attchments that were suppose to be sent but got stuck in queue
    sTemp = Dir(sSendEmailQueue & "\*.*", vbHidden)
    
    If sTemp <> vbNullString Then
        sTemp = goUtil.utDeleteFile(sSendEmailQueue & "\" & sTemp)
    End If
    
    EmailSinglePdfOnly = True
    
    If goUtil.utFileExists(sSendEmailQueue & "\" & sAttachName) Then
        goUtil.utDeleteFile sSendEmailQueue & "\" & sAttachName
    End If
    
    Exit Function
EH:
    
ERROR_MESS:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function EmailSinglePdfOnly" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "Could Not Send Message" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    
    moUL_ErrorMess sMess
    psErrorMess = sMess
    
    If goUtil.utFileExists(sSendEmailQueue & "\" & sAttachName) Then
        goUtil.utDeleteFile sSendEmailQueue & "\" & sAttachName
    End If
    
End Function

Private Function FTPPackageToClient(pPackageRS As ADODB.Recordset, _
                                    psTickcount As String, _
                                    psAttachName As String, _
                                    psFTPTo As String, _
                                    psSendFTPQueue As String, _
                                    psSentFTPBackup As String, _
                                    psErrorMess As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    FTPPackageToClient = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function FTPPackageToClient" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    'Save Error Message
    If msExportDocPath <> vbNullString Then
        goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & pPackageRS.Fields("Name").Value & "_" & psTickcount & "_Error.txt", sMess
    End If
    
    moUL_ErrorMess sMess
    psErrorMess = sMess
End Function

Private Function CleanXML(psXML As String) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sXML As String
    Dim oSer As WDDXSerializer
    Dim sMess As String
    
    Set oSer = New WDDXSerializer
    sXML = psXML
    'First get rid of Yucky chars
    sXML = Replace(sXML, Chr(160), Chr(32), , , vbBinaryCompare)
    sXML = Replace(sXML, vbCrLf, Chr(32), , , vbBinaryCompare)
    sXML = oSer.serialize(sXML)
   
    sXML = Mid(sXML, InStr(1, sXML, "<string>", vbBinaryCompare) + 8)
    sXML = Left(sXML, InStr(1, sXML, "</string>", vbBinaryCompare) - 1)
    CleanXML = sXML
    
    Set oSer = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function CleanXML" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrDesc & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Private Function UpdatePackageAsDelivered() As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sSQL As String
    Dim lRecordsAffected As Long
    Dim sProdDsn As String
    Dim sTemp As String
    Dim sAdminComments As String
    
    sAdminComments = Left(msPackageErrors, 1000)
    
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    
    If Not mbCreateSinglePdfOnly Then
        sTemp = msExportDocPath & " " & msSendToFTP & " " & msEmailEntireClaim & " " & msEmailDocsOnly & " " & msEmailPhotosOnly
        sTemp = Trim(sTemp)
    Else
        sTemp = "[Manual Email]"
    End If
    
    If Not mbCreateSinglePdfOnly Then
        sSQL = "UPDATE Assignments SET "
        sSQL = sSQL & "[StatusID] = (SELECT [StatusID] FROM Status WHERE [Status] = 'DELIVERED'), "
        sSQL = sSQL & "[DownLoadMe] = 1, "
        sSQL = sSQL & "[UpdateByUserID] = (SELECT [UsersID] FROM Users WHERE [UserName] = 'CFUSER'), "
        sSQL = sSQL & "[DateLastUpdated] = GetDate() "
        sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
        
        gConn.Execute sSQL, lRecordsAffected
    End If
    
    If Not mbCreateSinglePdfOnly Then
        sSQL = "UPDATE Package SET "
        sSQL = sSQL & "[PackageStatus] = (SELECT [Description] FROM Status WHERE [Status] = 'DELIVERED'), "
        sSQL = sSQL & "[SentToEmail] = '" & Left(goUtil.utCleanSQLString(sTemp), 50) & "', "
        sSQL = sSQL & "[SendMe] = 0, "
        sSQL = sSQL & "[SentDate] = GetDate(), "
        sSQL = sSQL & "[DownLoadMe] = 1, "
        sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(sAdminComments) & "', "
        sSQL = sSQL & "[DateLastUpdated] = GetDate(), "
        sSQL = sSQL & "[UpdateByUserID] = (SELECT [UsersID] FROM Users WHERE [UserName] = 'CFUSER') "
        sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
        sSQL = sSQL & "AND   [PackageID] = " & msPackageID & " "
    Else
        sSQL = "UPDATE Package SET "
        sSQL = sSQL & "[SentToEmail] = '" & Left(goUtil.utCleanSQLString(sTemp), 50) & "', "
        sSQL = sSQL & "[SentDate] = GetDate(), "
        sSQL = sSQL & "[DownLoadMe] = 1, "
        sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(sAdminComments) & "', "
        sSQL = sSQL & "[DateLastUpdated] = GetDate(), "
        sSQL = sSQL & "[UpdateByUserID] = (SELECT [UsersID] FROM Users WHERE [UserName] = 'CFUSER') "
        sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
        sSQL = sSQL & "AND   [PackageID] = " & msPackageID & " "
    End If
    
    gConn.Execute sSQL, lRecordsAffected
    
    UpdatePackageAsDelivered = True
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function UpdatePackageAsDelivered" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    If UnloadMode = vbFormControlMenu Then
        ' Just hide form if user presses Closes by X
        Me.Visible = False
        Cancel = True
    ElseIf UnloadMode = vbFormCode Then
'        FormWinRegPos Me, True
        Call ShellNotifyIcon(NIM_DELETE, m_NID)
        Set madoRSAssignments = Nothing
        Set mArv = Nothing
        Set moForm = Nothing
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
    
'    FormWinRegPos Me, True
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


Private Sub PP_MAIL_AsyncError(ByVal ErrorCode As Integer, ByVal ErrorMessage As String)
    msPackageErrors = msPackageErrors & " AsyncError: " & CStr(ErrorCode) & vbCrLf & ErrorMessage & vbCrLf
End Sub

Private Sub PP_MAIL_Debug(ByVal Message As String)
    Debug.Print Message
    msPackageErrors = msPackageErrors & Message & vbCrLf
End Sub


Private Sub PP_MAIL_Done()
    On Error GoTo EH
    
     Select Case mlMailState
        Case MAIL_STATE_CONNECTING
            mlMailState = MAIL_STATE_SENDING
            PP_MAIL.Flags = MailDstIsHost
            PP_MAIL.Action = MailActionWriteMessage
            If (PP_MAIL.Blocking = True) Then
                PP_MAIL_Done
            End If
        Case MAIL_STATE_SENDING
            mlMailState = MAIL_STATE_DISCONNECTING
            PP_MAIL.Action = MailActionDisconnect
            If (PP_MAIL.Blocking = True) Then
                PP_MAIL_Done
            End If
        Case MAIL_STATE_DISCONNECTING
            PP_MAIL.NewMessage
    End Select
    
    Exit Sub
EH:
    
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



Private Sub moUL_ErrorMess(ByVal Mess As String)
    On Error GoTo EH
    
    
    ErrorLog msUserName & vbCrLf & Mess
    
    'Also build the Member variable for Package Errors
    msPackageErrors = msPackageErrors & Mess & vbCrLf & vbCrLf
        
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

Private Sub Timer_SpinMe_Timer()
    On Error GoTo EH
    Static lPic As Long
    lPic = lPic + 1
    'Change the Tray Icon
    m_NID.hIcon = imgList.ListImages(lPic).Picture
    If lPic = 4 Then
        lPic = 0
    End If
    Call ShellNotifyIcon(NIM_MODIFY, m_NID)
    If mbSHUTDOWN Then
        Unload Me
    End If
    
    Exit Sub
EH:
    Err.Clear
    Timer_SpinMe.Enabled = False
End Sub


Private Sub ShowBusyIcon()
    On Error GoTo EH
    Dim sMess As String
    
    Timer_SpinMe.Enabled = True
    
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
    
    Timer_SpinMe.Enabled = False
    
    ProgBarLoss.Value = 0
    
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

Public Function VerifyIntegrity(psAssignmentsID As String, psMess As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim sMess As String
    Dim sReportFormat As String
    Dim sJPEGFileName As String
    Dim sPDFFileName As String
    Dim sPhotoReposPath As String
    Dim sAttachReposPath As String
    Dim FI As V2ECKeyBoard.FILE_INFORMATION
    Dim FA As V2ECKeyBoard.FILE_ATTRIBUTES
    Dim bPrintActiveReport As Boolean
    Dim bVerifyPhoto As Boolean
    Dim sTickCount As String
    Dim sTempFileDir As String
    Dim WddxFilePath As String
    Dim WddxFileName As String
    Dim sWddxData As String
    Dim oWddxSer As WDDXDeserializer
    Dim oWddxStruct As WDDXStruct
    Dim oWddxPhotoRS As WDDXRecordset
    Dim oWddxDataRS As WDDXRecordset
    Dim sPropNames As String
    Dim sTemp As String
    'Package Doc Sequence
    Dim lDocBase As Long
    Dim lThisDocNo As Long
    Dim lShouldBeDocNo As Long
    Dim sDocNo As String
    Dim sDocFail As String
    'Photo Sequence
    Dim lPhotoBase As Long
    Dim lThisPhotoNo As Long
    Dim lShouldBePhotoNo As Long
    Dim sPhotoNo As String
    Dim sPhotoFail As String
    Dim lRSPos As Long
    Dim lDocRsPos As Long
    Dim sProdDsn As String
    Dim sFTPSitePath As String
    
    
    'Be sure Temp File Dir Exists
    If Not goUtil.CreateTempDir(sTempFileDir) Then
        sTempFileDir = App.Path
    Else
        sTempFileDir = "C:\Temp\"
    End If
    
    sTickCount = goUtil.utGetTickCount
    WddxFilePath = sTempFileDir
    
    VerifyIntegrity = True
    
    'Set the Attach and Photo Repository dirs
    sFTPSitePath = GetSetting("V2WebControl", "Dir", "FTPSitePath", vbNullString)
    sPhotoReposPath = sFTPSitePath & "PhotoRepos\"
    sAttachReposPath = sFTPSitePath & "AttachRepos\"
    
    goUtil.AttachReposPath = sAttachReposPath
    goUtil.PhotoReposPath = sPhotoReposPath
    
    'Need to get the Package itmes for this Assingnment
    sSQL = "z_spsGetDeliverPackageItems "
    sSQL = sSQL & msAssignmentsID & ", "    '@AssignmentsID      int,
    sSQL = sSQL & "1 "                      '@bVerifyIntegrity   bit=0
    
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    
    
    Set RS = New ADODB.Recordset
    
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 0 Then
        sMess = "No Records Found!"
        VerifyIntegrity = False
        GoTo CLEAN_UP
    Else
        sMess = String(36, "-") & "File Integrity Failure Report " & String(36, "-") & vbCrLf
        sMess = sMess & Now() & vbCrLf
        RS.MoveFirst
        Do Until RS.EOF
            lDocRsPos = lDocRsPos + 1
            'First Check the Sequence for this doc see if it is not saved with correct sort number
            sDocNo = goUtil.IsNullIsVbNullString(RS.Fields("SortOrder"))
            'Make sure the photo Number matches the Sequence.
            'if not then user needs to save the current sort!
            lDocBase = 1
            lThisDocNo = CLng(sDocNo)
            lShouldBeDocNo = lDocBase + (lDocRsPos - 1)
            If lThisDocNo <> lShouldBeDocNo Then
                VerifyIntegrity = False
                sDocFail = vbCrLf & "Invalid Document Sequence / Sort order!" & vbCrLf
                sDocFail = sDocFail & "Please make sure Documents are in order, then save sort order."
                sMess = sMess & BuildFailureReason(RS, sDocFail)
                GoTo CLEAN_UP
            End If
            '.1. If the Item has an External file Associated (PhotoReport, DiagramReport, or is an actual PDF Attachment)
            '   Does the file exist where it is suppose to.  IE did the adjuster manipulate the file
            '   Outside of Easy Claim.  Or was there some problem on the adjusters box that molested the
            'file ?  Is the File zero Length.
            sReportFormat = goUtil.IsNullIsVbNullString(RS.Fields("ReportFormat"))
            If InStr(1, sReportFormat, "_arRptPhotos", vbTextCompare) > 0 Then
                'Need to create Wddx Packet that conatins the list of photos
                'Build the WddxFileName
                WddxFileName = "Temp_" & sTickCount & "_Photos.xml"
                bPrintActiveReport = PrintActiveReport(Nothing, , vbNullString, False, WddxFilePath, WddxFileName, True, True, sReportFormat, msCarListClassName)
                'associated with this report and verify they exist and file len > 0
                'As well need to verify that the sort order is correct.
                If Not bPrintActiveReport Then
                    VerifyIntegrity = False
                    sMess = sMess & BuildFailureReason(RS, "System Error while Verifying Integrity!")
                    GoTo NEXT_RS
                Else
                    bVerifyPhoto = True
                    sPhotoFail = vbNullString
                    sWddxData = goUtil.utGetFileData(WddxFilePath & WddxFileName)
                    Set oWddxSer = New WDDXDeserializer
                    Set oWddxStruct = oWddxSer.deserialize(sWddxData)
                    sPropNames = Join(oWddxStruct.getPropNames, "|")
                    If InStr(1, sPropNames, "PhotosRS", vbTextCompare) = 0 Then
                        VerifyIntegrity = False
                        sPhotoFail = "Empty Photo Report! Please remove from package."
                        sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                        GoTo NEXT_RS
                    ElseIf InStr(1, sPropNames, "DataRS", vbTextCompare) = 0 Then
                        VerifyIntegrity = False
                        sPhotoFail = "Error Reading Data Recordset!"
                        sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                        GoTo NEXT_RS
                    End If
                    Set oWddxPhotoRS = oWddxStruct.getProp("PhotosRS")
                    Set oWddxDataRS = oWddxStruct.getProp("DataRS")
                    'Loop through each Photo and determine if every thing is Kosher
                    For lRSPos = 1 To oWddxPhotoRS.getRowCount
                        sJPEGFileName = oWddxPhotoRS.getField(lRSPos, "imgPhotoPath")
                        sPhotoNo = oWddxPhotoRS.getField(lRSPos, "fPhotoNo")
                        'Make sure the photo Number matches the Sequence.
                        'if not then user needs to save the current sort!
                        sTemp = oWddxDataRS.getField(1, "f_Description")
                        sTemp = Left(sTemp, 3)
                        lPhotoBase = CLng(sTemp)
                        lThisPhotoNo = CLng(sPhotoNo)
                        lShouldBePhotoNo = lPhotoBase + (lRSPos - 1)
                        If lThisPhotoNo <> lShouldBePhotoNo Then
                            VerifyIntegrity = False
                            sPhotoFail = vbCrLf & "Invalid Photo Sequence / Sort order!" & vbCrLf
                            sPhotoFail = sPhotoFail & "Please make sure photos are in order, then save sort order."
                            sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                            GoTo NEXT_RS
                        End If
                        
                        If Not goUtil.utFileExists(sJPEGFileName) Then
                            '10.3.2005 Also Set the Downloadme, uploadphoto flags
                            'so that the next time the adjusters connects they will
                            'upload the missing photo
                            SetJPEGRestoreFlags sJPEGFileName
                            bVerifyPhoto = False
                            sJPEGFileName = vbCrLf & "Photo No: [" & sPhotoNo & "] File Not Found!: " & sJPEGFileName
                            sPhotoFail = sPhotoFail & sJPEGFileName
                            GoTo NEXT_PHOTO
                        Else
                            'Need to ensure that the pdf file exists and the file len > 0
                            GetFileSettings sJPEGFileName, FI, FA
                            If FI.nFileSize = 0 Then
                                '10.3.2005 Also Set the Downloadme, uploadphoto flags
                                'so that the next time the adjusters connects they will
                                'upload the missing photo
                                SetJPEGRestoreFlags sJPEGFileName
                                bVerifyPhoto = False
                                sJPEGFileName = vbCrLf & "Photo No: [" & sPhotoNo & "] Invalid File Size!: " & sJPEGFileName
                                sPhotoFail = sPhotoFail & sJPEGFileName
                                GoTo NEXT_PHOTO
                            End If
                        End If
NEXT_PHOTO:
                    Next
                    If Not bVerifyPhoto Then
                        VerifyIntegrity = False
                        sMess = sMess & BuildFailureReason(RS, sPhotoFail)
                        GoTo NEXT_RS
                    End If
                End If
                
            ElseIf InStr(1, sReportFormat, ".pdf", vbTextCompare) > 0 Then
                'PDF ATTACHMENT
                'FRE27745_050608152131_1.pdf|
                'Need to ensure that the pdf file exists and the file len > 0
                sPDFFileName = Trim(Right(sReportFormat, 200))
                sPDFFileName = Replace(sPDFFileName, "|", vbNullString, , , vbBinaryCompare)
                sPDFFileName = sAttachReposPath & "\" & GetYYMMDDFolders(sPDFFileName) & sPDFFileName
                If Not goUtil.utFileExists(sPDFFileName) Then
                    SetPDFRestoreFlags sPDFFileName
                    VerifyIntegrity = False
                    sPDFFileName = vbCrLf & "File Not Found!: " & sPDFFileName
                    sMess = sMess & BuildFailureReason(RS, sPDFFileName)
                    GoTo NEXT_RS
                Else
                    'Need to ensure that the pdf file exists and the file len > 0
                    GetFileSettings sPDFFileName, FI, FA
                    If FI.nFileSize = 0 Then
                        SetPDFRestoreFlags sPDFFileName
                        VerifyIntegrity = False
                        sPDFFileName = vbCrLf & "Invalid File Size!: " & sPDFFileName
                        sMess = sMess & BuildFailureReason(RS, sPDFFileName)
                        GoTo NEXT_RS
                    End If
                End If
            ElseIf InStr(1, sReportFormat, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                'DIAGRAM WORKSHEET
                'ECrptFarmers_arWorkSheetDiag|clsLists|1|1|FRE27745_050608155648.jpg
                'need to verify the jpg photo exists and file len > 0
                sJPEGFileName = Trim(Right(sReportFormat, 200))
                sJPEGFileName = Mid(sJPEGFileName, InStrRev(sJPEGFileName, "|", , vbBinaryCompare) + 1)
                sJPEGFileName = sPhotoReposPath & "\" & GetYYMMDDFolders(sJPEGFileName) & sJPEGFileName
                If Not goUtil.utFileExists(sJPEGFileName) Then
                    SetWSRestoreFlags sJPEGFileName
                    VerifyIntegrity = False
                    sJPEGFileName = vbCrLf & "File Not Found!: " & sJPEGFileName
                    sMess = sMess & BuildFailureReason(RS, sJPEGFileName)
                    GoTo NEXT_RS
                Else
                    'Need to ensure that the pdf file exists and the file len > 0
                    GetFileSettings sJPEGFileName, FI, FA
                    If FI.nFileSize = 0 Then
                        SetWSRestoreFlags sJPEGFileName
                        VerifyIntegrity = False
                        sJPEGFileName = vbCrLf & "Invalid File Size!: " & sJPEGFileName
                        sMess = sMess & BuildFailureReason(RS, sJPEGFileName)
                        GoTo NEXT_RS
                    End If
                End If
            End If
            'Check each item for the following...
            '2. Is there an Item still flagged for Upload!
            '   If there is, then this will be cause to fail integrity becuase it
            '   Does not yet Exist on the Web Server
            If CBool(RS.Fields("UpLoadMe")) Then
                VerifyIntegrity = False
                sMess = sMess & BuildFailureReason(RS, "RECORD NEEDS TO BE UPLOADED!")
                GoTo NEXT_RS
            End If
NEXT_RS:
            RS.MoveNext
        Loop
    End If
    
CLEAN_UP:
    psMess = sMess
    Set RS = Nothing
    Set oWddxSer = Nothing
    Set oWddxStruct = Nothing
    Set oWddxPhotoRS = Nothing
    Set oWddxDataRS = Nothing
    Exit Function
EH:
    VerifyIntegrity = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function VerifyIntegrity" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    psMess = "System Error while Verifying Integrity!"
End Function

Private Function SetJPEGRestoreFlags(psJPGPath As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sSQL As String
    Dim sProdDsn As String
    Dim lRecordsAffected As Long
    Dim sPhotoName As String
    Dim sAdminComments As String
    
    'Get the Photoname from the photopath
    sPhotoName = psJPGPath
    sPhotoName = Mid(sPhotoName, InStrRev(sPhotoName, "\", , vbBinaryCompare) + 1, Len(sPhotoName))
           
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    
    'This update statement will force the adjuster to download the
    'upload photo flags.  Consequently the next time connecting
    'the photo upload will occur again.
    sAdminComments = "SYSTEM PHOTO RESTORE " & Now()
    sSQL = "UPDATE RTPhotoLog SET "
    sSQL = sSQL & "[DownLoadMe] = 1, "
    sSQL = sSQL & "[UpLoadMe] = 1, "
    sSQL = sSQL & "[UpLoadPhoto] = 1, "
    sSQL = sSQL & "[UpLoadPhotoThumb] = 1, "
    sSQL = sSQL & "[UpLoadPhotoHighRes] = 1, "
    sSQL = sSQL & "[AdminComments] = '" & Replace(sAdminComments, "'", "''", , , vbBinaryCompare) & "' "
    sSQL = sSQL & "WHERE [PhotoName] = '" & Replace(sPhotoName, "'", "''", , , vbBinaryCompare) & "' "
    
    gConn.Execute sSQL, lRecordsAffected
    
    SetJPEGRestoreFlags = CBool(lRecordsAffected)
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function SetJPEGRestoreFlags & vbCrLf"
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Private Function SetPDFRestoreFlags(psPDFPath As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sSQL As String
    Dim sProdDsn As String
    Dim lRecordsAffected As Long
    Dim sAttachment As String
    Dim sAdminComments As String
    
    'Get the Photoname from the photopath
    sAttachment = psPDFPath
    sAttachment = Mid(sAttachment, InStrRev(sAttachment, "\", , vbBinaryCompare) + 1, Len(sAttachment))
           
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    
    'This update statement will force the adjuster to download the
    'upload photo flags.  Consequently the next time connecting
    'the photo upload will occur again.
    sAdminComments = "SYSTEM PDF ATTACHMENT RESTORE " & Now()
    sSQL = "UPDATE RTAttachments SET "
    sSQL = sSQL & "[DownLoadMe] = 1, "
    sSQL = sSQL & "[UpLoadMe] = 1, "
    sSQL = sSQL & "[UpLoadAttachment] = 1, "
    sSQL = sSQL & "[AdminComments] = '" & Replace(sAdminComments, "'", "''", , , vbBinaryCompare) & "' "
    sSQL = sSQL & "WHERE [Attachment] = '" & Replace(sAttachment, "'", "''", , , vbBinaryCompare) & "' "
    
    gConn.Execute sSQL, lRecordsAffected
    
    SetPDFRestoreFlags = CBool(lRecordsAffected)
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function SetPDFRestoreFlags & vbCrLf"
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Private Function SetWSRestoreFlags(psJPGPath As String) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim sSQL As String
    Dim sProdDsn As String
    Dim lRecordsAffected As Long
    Dim sDiagramPhotoName As String
    Dim sAdminComments As String
    
    'Get the Diagram Photoname from the photopath
    sDiagramPhotoName = psJPGPath
    sDiagramPhotoName = Mid(sDiagramPhotoName, InStrRev(sDiagramPhotoName, "\", , vbBinaryCompare) + 1, Len(sDiagramPhotoName))
           
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    
    'This update statement will force the adjuster to download the
    'upload Diagram photo flags.  Consequently the next time connecting
    'the Diagram photo upload will occur again.
    sAdminComments = "SYSTEM DIAGRAM PHOTO RESTORE " & Now()
    sSQL = "UPDATE RTWSDiagram SET "
    sSQL = sSQL & "[DownLoadMe] = 1, "
    sSQL = sSQL & "[UpLoadMe] = 1, "
    sSQL = sSQL & "[UploadDiagramPhoto] = 1, "
    sSQL = sSQL & "[AdminComments] = '" & Replace(sAdminComments, "'", "''", , , vbBinaryCompare) & "' "
    sSQL = sSQL & "WHERE [DiagramPhotoName] = '" & Replace(sDiagramPhotoName, "'", "''", , , vbBinaryCompare) & "' "
    
    gConn.Execute sSQL, lRecordsAffected
    
    SetWSRestoreFlags = CBool(lRecordsAffected)
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function SetWSRestoreFlags & vbCrLf"
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Private Function GetYYMMDDFolders(psFileName As String) As String
    '10.3.2002 this fucntion will determine the month and year from the filename
    'and depending on the ADJ folder
    On Error GoTo EH
    Dim sMess As String
    Dim sTemp As String
    Dim sYY As String
    Dim sMM As String
    Dim sDD As String
    
    'Build the Month Year Folder Name based off of the File name
    'FRE26582_050414110631_1.pdf

    sTemp = Mid(psFileName, InStr(1, psFileName, "_", vbBinaryCompare) + 1)
    sYY = Left(sTemp, 2)
    sMM = Mid(sTemp, 3, 2)
    sDD = Mid(sTemp, 5, 2)
    
    
    
    GetYYMMDDFolders = sYY & sMM & "\" & sDD & "\"
    
    Exit Function
EH:
    
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function GetYYMMDDFolders" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Private Function BuildFailureReason(pRS As ADODB.Recordset, sFailureReason As String) As String
    On Error GoTo EH
    Dim sMess As String
    
    sMess = String(102, "-") & vbCrLf
    sMess = sMess & "[Sort]" & String(4, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("SortOrder")) & vbCrLf
    sMess = sMess & "[Name]" & String(4, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("Name")) & vbCrLf
    sMess = sMess & "[Description]" & String(3, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("Description")) & vbCrLf
    sMess = sMess & "[AttachmentName]" & String(2, vbTab) & goUtil.IsNullIsVbNullString(pRS.Fields("AttachmentName")) & vbCrLf
    sMess = sMess & "[Failure Reason]" & String(2, vbTab) & sFailureReason & vbCrLf
    sMess = sMess & String(102, "-") & vbCrLf & vbCrLf
    
    BuildFailureReason = sMess
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function BuildFailureReason" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Public Function PrintActiveReport(poReportItem As Object, _
                                Optional piMode As VBRUN.FormShowConstants = vbModeless, _
                                Optional psCopyName As String = vbNullString, _
                                Optional pbPrintPreview As Boolean = True, _
                                Optional psSaveToFilePath As String, _
                                Optional psSaveToFileName As String, _
                                Optional pbExportXML As Boolean, _
                                Optional pbExportXMLOnly As Boolean, _
                                Optional psReportFormat As String, _
                                Optional psCarListClassName As String) As Boolean
    On Error GoTo EH
    Dim sParams As String
    Dim sReportName As String
    Dim sReportTitle As String
    Dim srptProjectName As String
    Dim srptClassName As String
    Dim lrptVersion As Long
    Dim sData As String
    Dim saryData() As String
    Dim ocboReport As ComboBox
    Dim itmXReport As ListItem
    Dim MyActReport As ActiveReport
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'If using Adobe PDF Viewer
    Dim sPDFFilePath As String
    Dim sPrintPreview As String
    Dim bUseAdobeReader As Boolean
    'Some Reports need extra Params passed to them
    'Payments
    Dim sRTChecksID As String
    Dim sCheckNum As String
    'Internal Billing
    Dim sIBID As String
    Dim sSupplement As String
    'Photo Reports (Multi Report)
    Dim sPhotoReportNumber As String
    'Worksheet Diagram (Multi Report)
    Dim sDiagramNumber As String
    Dim sNumber As String
    'Loss Report
    Dim oLR As V2ECKeyBoard.clsLossReports
    Dim MyAssignmentsRS As ADODB.Recordset
    Dim sLRFormat As String
    Dim sLossReport As String
    Dim sLRData As String
    'Export to XML FileName
    Dim sXMLFilePath As String
    Dim sXMLFileName As String
    Dim sMess As String
    Dim sFTPSitePath As String
    Dim sPhotoReposPath As String
    Dim sAttachReposPath As String
    
     'Set the Attach and Photo Repository dirs
    sFTPSitePath = GetSetting("V2WebControl", "Dir", "FTPSitePath", vbNullString)
    sPhotoReposPath = sFTPSitePath & "PhotoRepos\"
    sAttachReposPath = sFTPSitePath & "AttachRepos\"
    
    sPrintPreview = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE")
    
    Select Case UCase(sPrintPreview)
        Case "USE_ADOBE"
            bUseAdobeReader = True
    End Select
    
    'if saving to file path always use adobe reader
    If psSaveToFilePath <> vbNullString Then
        bUseAdobeReader = True
    End If
    
    If psReportFormat <> vbNullString Then
        sData = psReportFormat
    Else
        Exit Function
    End If
    
    If sData <> vbNullString Then
        sReportTitle = Trim(Left(sData, 200))
        goUtil.utCleanFileFolderName sReportTitle, False
        sData = Mid(sData, InStr(1, sData, String(100, " "), vbBinaryCompare))
        sData = Trim(sData)
        saryData() = Split(sData, "|", , vbBinaryCompare)
        If UBound(saryData, 1) <= 1 Then
            'Check for Loss Report
            sLRFormat = saryData(0)
            If StrComp(sLRFormat, "LRFormat", vbTextCompare) = 0 Then
                Me.SetadoRSAssignments msAssignmentsID
                Set MyAssignmentsRS = Me.adoRSAssignments
                sLRFormat = goUtil.IsNullIsVbNullString(MyAssignmentsRS.Fields("LRFormat"))
                sLossReport = goUtil.IsNullIsVbNullString(MyAssignmentsRS.Fields("LossReport"))
                
                If InStr(1, sLRFormat, "OLEType_pdf", vbTextCompare) > 0 Then
                    sPDFFilePath = goUtil.AttachReposPath & sLossReport
                    If psSaveToFilePath <> vbNullString Then
                        'Do not do Loss Report if Export to xml only is true
                        If pbExportXML And pbExportXMLOnly Then
                            PrintActiveReport = False
                            Set ocboReport = Nothing
                            Set itmXReport = Nothing
                            Set oLR = Nothing
                            Set MyAssignmentsRS = Nothing
                            MsgBox "Loss Reports can not be part of and XML ONLY Export!", vbExclamation
                            Exit Function
                        End If
                        goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                    Else
                        If pbPrintPreview Then
                            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        Else
                            goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        End If
                        DoEvents
                        Sleep 100
                    End If
                Else
                    sPDFFilePath = goUtil.gsInstallDir & "\TempLossReport" & goUtil.utGetTickCount & ".pdf"
                    Set oLR = New V2ECKeyBoard.clsLossReports
                    If StrComp(sLRFormat, "TEXT", vbTextCompare) <> 0 Then
                        sLRData = sLRFormat & vbCrLf & sLossReport
                    Else
                        sLRData = sLossReport
                    End If
                    oLR.CreateExport sLRData, sPDFFilePath, ARPdf
                    If psSaveToFilePath <> vbNullString Then
                        'Do not do Loss Report if Export to xml only is true
                        If pbExportXML And pbExportXMLOnly Then
                            PrintActiveReport = False
                            Set ocboReport = Nothing
                            Set itmXReport = Nothing
                            Set oLR = Nothing
                            Set MyAssignmentsRS = Nothing
                            MsgBox "Loss Reports can not be part of and XML ONLY Export!", vbExclamation
                            Exit Function
                        End If
                        goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                    Else
                        If pbPrintPreview Then
                            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        Else
                            goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                        End If
                        DoEvents
                        Sleep 100
                    End If
                    goUtil.utDeleteFile sPDFFilePath
                End If
                PrintActiveReport = True
                GoTo CLEAN_UP
            End If
            
            sPDFFilePath = saryData(0)
            If InStr(1, sPDFFilePath, ".pdf", vbTextCompare) > 0 Then
                sPDFFilePath = goUtil.AttachReposPath & GetYYMMDDFolders(sPDFFilePath) & sPDFFilePath
                'Check for Pdf Attachment file
                If psSaveToFilePath <> vbNullString Then
                    'Do not do Attachments if Export to xml only is true
                    If pbExportXML And pbExportXMLOnly Then
                        PrintActiveReport = False
                        Set ocboReport = Nothing
                        Set itmXReport = Nothing
                        Set oLR = Nothing
                        Set MyAssignmentsRS = Nothing
                        MsgBox "Attachments can not be part of and XML ONLY Export!", vbExclamation
                        Exit Function
                    End If
                    goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                Else
                    If pbPrintPreview Then
                        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    Else
                        goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    End If
                    DoEvents
                    Sleep 100
                End If
            End If
            PrintActiveReport = True
            GoTo CLEAN_UP
        End If
        srptProjectName = saryData(0)
        srptClassName = saryData(1)
        lrptVersion = saryData(2)
        'Check For Multi Reports Here
        If psReportFormat <> vbNullString Then
            If UBound(saryData, 1) >= 3 Then
                sNumber = saryData(3)
            End If
            
            'If this is coming from the Package Screen need to populate the Number for certain reports
            If InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
                sPhotoReportNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                sDiagramNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
                sCheckNum = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
                sSupplement = sNumber
            End If
        ElseIf TypeOf poReportItem Is ListItem Then
            If UBound(saryData, 1) >= 3 Then
                sNumber = saryData(3)
            End If
            
            'If this is coming from the Package Screen need to populate the Number for certain reports
            If InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
                sPhotoReportNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
                sDiagramNumber = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
                sCheckNum = sNumber
            ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
                sSupplement = sNumber
            End If
        ElseIf InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
            'Photo Reports (Multi Report)
            sPhotoReportNumber = saryData(3)
        ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
            'Worksheet Diagram (Multi Report)
            sDiagramNumber = saryData(3)
        End If
    Else
        Exit Function
    End If
    
    'Build Params List to be passed in to Create Report Object
    'This Object will have list of Report Parameters it requires
    
    sParams = vbNullString
    sParams = sParams & "psAssignmentsID=" & msAssignmentsID & "|"
    'If using Adobe PDF Viewer
    If bUseAdobeReader Then
        sPDFFilePath = goUtil.gsInstallDir & "\TempActiveReport" & goUtil.utGetTickCount & ".pdf"
        sParams = sParams & "psXportPath=" & sPDFFilePath & "|"
        sParams = sParams & "pPDFJPEGQuality=" & "40" & "|"
        sParams = sParams & "pXportType=" & ExportType.ARPdf & "|"
    Else
        sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
    End If
    sParams = sParams & "pPhotoReposPath=" & sPhotoReposPath & "|"
    sParams = sParams & "pAttachReposPath=" & sAttachReposPath & "|"
    'Add the friggin Photo Repos Path
    
    
    'Certain Reports Need to have some more Params Passed in
    If InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
        'Need to Get the ChecksID and Check Number
        If Not ocboReport Is Nothing Then
            sRTChecksID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
            If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
                GoTo CLEAN_UP
            End If
        ElseIf Not itmXReport Is Nothing Then
            'the schecknum was already set above
            If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
                GoTo CLEAN_UP
            End If
        ElseIf psReportFormat <> vbNullString Then
            'the schecknum was already set above
            If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
                GoTo CLEAN_UP
            End If
        End If
        sParams = sParams & "pRTChecksID=" & sRTChecksID & "|"
        sParams = sParams & "psCheckNum=" & sCheckNum & "|"
    ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
        'If the IBID and Supplement Parameters already exist then use them
        'Otherwise have to do Data Call to get em.
        If InStr(1, sData, "pIBID=", vbTextCompare) > 0 And InStr(1, sData, "pSupplement=", vbTextCompare) > 0 Then
            sParams = sParams & saryData(3) & "|"
            sParams = sParams & saryData(4) & "|"
            'Check for Report Title As Well
            If InStr(1, sData, "psReportTitle=", vbTextCompare) > 0 Then
                sReportTitle = Mid(saryData(5), InStr(1, saryData(5), "=", vbTextCompare) + 1)
            End If
        Else
            If Not ocboReport Is Nothing Then
                sIBID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
                If Not GetIBParams(sIBID, sSupplement) Then
                    GoTo CLEAN_UP
                End If
            ElseIf Not itmXReport Is Nothing Then
                'The supplement was already set above
                If Not GetIBParams(sIBID, sSupplement) Then
                    GoTo CLEAN_UP
                End If
            ElseIf sSupplement <> vbNullString Then
                If Not GetIBParams(sIBID, sSupplement) Then
                    GoTo CLEAN_UP
                End If
            End If
            sParams = sParams & "pIBID=" & sIBID & "|"
            sParams = sParams & "pSupplement=" & sSupplement & "|"
        End If
        sParams = sParams & "pCopyName=" & psCopyName & "|"
    ElseIf InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
        'Photo Reports (Multi Report)
        sParams = sParams & "pNumber=" & sPhotoReportNumber & "|"
    ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
        'Worksheet Diagram (Multi Report)
        sParams = sParams & "pNumber=" & sDiagramNumber & "|"
    End If
    
    sReportName = srptProjectName & "." & srptClassName

    If StrComp(psCopyName, "(-ALL COPIES-)", vbTextCompare) = 0 Then
        'Do a recursive call until All Copies are printed
        'Client company Copy
        If Not PrintActiveReport(poReportItem, , goUtil.gsCurCarDBName & " Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Company Copy
        If Not PrintActiveReport(poReportItem, , GetSetting(goUtil.gsAppEXEName, "GENERAL", "CURRENT_COMPANY_NAME", "Company") & " Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Remit Copy
        If Not PrintActiveReport(poReportItem, , "Remit Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
        'Adjuster Copy
        If Not PrintActiveReport(poReportItem, , "Adjuster Copy", pbPrintPreview, psSaveToFilePath, psSaveToFileName) Then
            GoTo CLEAN_UP
        End If
    Else
        Set oCarList = CreateObject(psCarListClassName)
        If bUseAdobeReader Then
            'Add Export XML Parameters here
            If pbExportXML Then
                sParams = sParams & "pbExportXML=True|"
                If pbExportXMLOnly Then
                    sParams = sParams & "pbExportXMLOnly=True|"
                End If
            End If
            Set MyActReport = oCarList.GetARReport(sReportName, lrptVersion, sParams)
            If goUtil.utFileExists(sPDFFilePath) Or (pbExportXML And pbExportXMLOnly) Then
                If psSaveToFilePath <> vbNullString Then
                    If pbExportXML Then
                        If Not pbExportXMLOnly Then
                            If psCopyName <> vbNullString Then
                                goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & Replace(psSaveToFileName, ".pdf", "_" & psCopyName & ".pdf", , 1, vbTextCompare)
                            Else
                                goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                            End If
                        End If
                    Else
                        If psCopyName <> vbNullString Then
                            goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & Replace(psSaveToFileName, ".pdf", "_" & psCopyName & ".pdf", , 1, vbTextCompare)
                        Else
                            goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                        End If
                    End If
                    
                    If pbExportXML Then
                        'change the pdffile path the XML
                        sXMLFilePath = sPDFFilePath
                        sXMLFilePath = Left(sXMLFilePath, InStrRev(sXMLFilePath, ".", , vbBinaryCompare))
                        sXMLFilePath = sXMLFilePath & "xml"
                        'Change the pdf to XML file path
                        sXMLFileName = psSaveToFileName
                        sXMLFileName = Left(sXMLFileName, InStrRev(sXMLFileName, ".", , vbBinaryCompare))
                        sXMLFileName = sXMLFileName & "xml"
                       If psCopyName <> vbNullString Then
                            goUtil.utCopyFile sXMLFilePath, psSaveToFilePath & Replace(sXMLFileName, ".xml", "_" & psCopyName & ".xml", , 1, vbTextCompare)
                        Else
                            goUtil.utCopyFile sXMLFilePath, psSaveToFilePath & sXMLFileName
                        End If
                    End If
                Else
                    If pbPrintPreview Then
                        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    Else
                        goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, sReportTitle
                    End If
                    
                End If
                ' DoEvents
                Sleep 100
                goUtil.utDeleteFile sPDFFilePath
                goUtil.utDeleteFile sXMLFilePath
                If Not MyActReport Is Nothing Then
                    Unload MyActReport
                    Set MyActReport = Nothing
                End If
'                oCarList.CLEANUP
                Set oCarList = Nothing
            End If
        Else
            'Using Active Report Viewer
            Set MyActReport = oCarList.GetARReport(sReportName, lrptVersion, sParams)
        
            If pbPrintPreview Then
                If mArv Is Nothing Then
                    Set mArv = New V2ARViewer.clsARViewer
                    mArv.SetUtilObject goUtil
                End If
'                If Not moForm Is Nothing Then
'                    If StrComp(psCopyName, "(-ALL COPIES-)", vbTextCompare) <> 0 Then
'                        Unload moForm
'                        Set moForm = Nothing
'                    End If
'                End If
                With mArv
                    'Pass in true to have Active reports process on separate thread.
                    'This will allow the viewer to load while the report is processing
                    'false will force the report to run on single thread
                    MyActReport.Run False 'True
                    .objARvReport = MyActReport
                    .sRptTitle = sReportTitle
                    .HidePrintButton = False
                    .ShowReportOnForm moForm, piMode
        
                    Unload .objARvReport
                    Set .objARvReport = Nothing
                End With
            Else
                MyActReport.PrintReport False
            End If
            Unload MyActReport
            Set MyActReport = Nothing
'            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    End If
    PrintActiveReport = True
CLEAN_UP:
    'Cleanup
    Set ocboReport = Nothing
    Set itmXReport = Nothing
    Set oLR = Nothing
    Set MyAssignmentsRS = Nothing
    
    'Clear the Local ref to this report object only
    'The actual cleanup of this active report object will occur within gARV
    PrintActiveReport = True
    Exit Function
EH:
    PrintActiveReport = False
    Screen.MousePointer = vbDefault
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function PrintActiveReport" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Public Sub GetFileSettings(sFilePath, pFI As V2ECKeyBoard.FILE_INFORMATION, pFA As V2ECKeyBoard.FILE_ATTRIBUTES)
    On Error GoTo EH
    Dim oFI As V2ECKeyBoard.clsFileVersion
    Dim sMess As String
    
    Set oFI = New V2ECKeyBoard.clsFileVersion
    pFI = oFI.GetFileInformation(sFilePath)
    pFA = pFI.faFileAttributes
    Set oFI = Nothing
    
    'CleanUp
    Set oFI = Nothing
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Sub GetFileSettings" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Sub

Public Function SetadoRSAssignments(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim sProdDsn As String
    Dim sSQL As String
    Dim sMess As String
    
    'rese the typeof loss rs
    If Not madoRSAssignments Is Nothing Then
        Set madoRSAssignments = Nothing
    End If
    
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    Set madoRSAssignments = New ADODB.Recordset
       
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM Assignments "
    sSQL = sSQL & "WHERE ID = " & psIDAssignments & " "
    
    madoRSAssignments.CursorLocation = adUseClient
    madoRSAssignments.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSAssignments.ActiveConnection = Nothing
    
    SetadoRSAssignments = True
    
    Exit Function
EH:
    SetadoRSAssignments = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function SetadoRSAssignments" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Public Function GetPaymentsParams(psRTChecksID As String, psCheckNum As String) As Boolean
    On Error GoTo EH
    Dim sProdDsn As String
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sRTChecksID As String
    Dim sCheckNum As String
    Dim sMess As String
    
    sProdDsn = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDsn
    Set RS = New ADODB.Recordset
    
    sSQL = "SELECT RTC.[RTChecksID], "
    sSQL = sSQL & "RTC.[CheckNum] "
    sSQL = sSQL & "FROM RTChecks RTC "
    sSQL = sSQL & "WHERE RTC.[AssignmentsID] = " & msAssignmentsID & " "
    If psRTChecksID = vbNullString Then
        sSQL = sSQL & "AND RTC.[CheckNum] = " & psCheckNum & " "
    Else
        sSQL = sSQL & "AND RTC.[RTChecksID] = " & psRTChecksID & " "
    End If
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        sRTChecksID = RS.Fields("RTChecksID").Value
        sCheckNum = RS.Fields("CheckNum").Value
    End If
    
    
    psRTChecksID = sRTChecksID
    psCheckNum = sCheckNum
    GetPaymentsParams = True
    
    'Cleanup
    Set RS = Nothing
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function GetPaymentsParams" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function

Public Function GetIBParams(psIBID As String, psSupplement As String) As Boolean
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sIBID As String
    Dim sSupplement As String
    Dim sMess As String
    
    
    Set RS = New ADODB.Recordset
    
    sSQL = "SELECT IB.[IBID], "
    sSQL = sSQL & "BC.[Supplement] "
    sSQL = sSQL & "FROM IB "
    sSQL = sSQL & "INNER JOIN BillingCount BC ON IB.BillingCountID = BC.BillingCountID "
    sSQL = sSQL & "WHERE IB.[AssignmentsID] = " & msAssignmentsID & " "
    If psIBID = vbNullString Then
        sSQL = sSQL & "AND IB.[IB14a_sSupplement] = " & psSupplement & " "
    Else
        sSQL = sSQL & "AND IB.[IBID] = " & psIBID & " "
    End If
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        sIBID = RS.Fields("IBID").Value
        sSupplement = RS.Fields("Supplement").Value
    End If
    
    
    psIBID = sIBID
    psSupplement = sSupplement
    GetIBParams = True
    
    'Cleanup
    Set RS = Nothing
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function GetIBParams" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
End Function


Private Sub DoPack()
    On Error GoTo EH
    Dim sMess As String
    Dim sSQL As String
    Dim sTimeStamp As String
    Dim lRecordsAffected As Long
    
    
    
    sTimeStamp = Format(Now(), "MMDDYYYYHHmmss")
    
    If Not VerifyIntegrity(msAssignmentsID, sMess) Then
        sMess = sMess & vbCrLf & msPackageErrors
        goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & "FileIntegrity_" & sTimeStamp & ".txt", sMess
        If Not mbCreateSinglePdfOnly Then
            REJECT00
        Else
            sSQL = "UPDATE Package SET "
            sSQL = sSQL & "[DownLoadMe] = 1, "
            sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(Left(sMess, 1000)) & "', "
            sSQL = sSQL & "[DateLastUpdated] = GetDate(), "
            sSQL = sSQL & "[UpdateByUserID] = (SELECT [UsersID] FROM Users WHERE [UserName] = 'CFUSER') "
            sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
            sSQL = sSQL & "AND   [PackageID] = " & msPackageID & " "
            
            gConn.Execute sSQL, lRecordsAffected
            GoTo DELIVER
        End If
    Else
DELIVER:
        If DELIVERPackage() Then
            UpdatePackageAsDelivered
        Else
            sMess = "Error delivering package!" & vbCrLf & sMess & vbCrLf & msPackageErrors
            goUtil.utSaveFileData msBuildAssgnPackPath & "ERRORS\" & "FileIntegrity_" & sTimeStamp & ".txt", sMess
            If Not mbCreateSinglePdfOnly Then
                REJECT00
            End If
        End If
    End If
    
    mbSHUTDOWN = True
    
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub DoPack" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    mbSHUTDOWN = True
End Sub

Private Sub Timer_Start_Timer()
    Timer_Start.Enabled = False
    DoPack
End Sub


Private Function MergePDFFiles(psRawPDFFilesDir As String, _
                                psSinglePDFOutputDir As String, _
                                psSinglePDFOutputName As String) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMess As String
    
    Dim bFirstDoc As Boolean
    Dim sRawPDFFilesDir As String
    Dim sSinglePDFOutputDir As String
    Dim sSinglePDFOutputName As String
    Dim oMainDoc As Acrobat.CAcroPDDoc
    Dim oTempDoc As Acrobat.CAcroPDDoc
    'Need to use Adobe internal Java Object
    'in order to Add Book marks
    Dim oJSO As Object 'JavaScript Object
    Dim oBookMarkRoot As Object
    Dim oFolder As Scripting.Folder
    Dim saryFileSort() As String
    Dim oFile As Scripting.File
    Dim oFSO As Scripting.FileSystemObject
    Dim sBMName As String
    Dim lPos As Long
    Dim lFile As Long
    Dim lBMPageNo As Long
    'Use these for Insterts
    Dim lInsertPageAfter As Long
    Dim lNumPages As Long
    Dim lRet As Long
    Dim sTemp As String
    
    sRawPDFFilesDir = psRawPDFFilesDir
    sSinglePDFOutputDir = psSinglePDFOutputDir
    sSinglePDFOutputName = psSinglePDFOutputName
    
    Set oFSO = New Scripting.FileSystemObject
    
    Set oFolder = oFSO.GetFolder(sRawPDFFilesDir)
    
    bFirstDoc = True

    If oFolder.Files.Count = 0 Then
        Exit Function
    End If
    
    'Because the FSO folder files collection does not allow for
    'Native sorting, need to plug all the files into an array and sort that motha
    ReDim saryFileSort(1 To oFolder.Files.Count)
    lFile = 0
    For Each oFile In oFolder.Files
        lFile = lFile + 1
        saryFileSort(lFile) = oFile.Name
    Next
    
    'Once they is all in der sor the array
    goUtil.utBubbleSort saryFileSort
    
    For lFile = 1 To UBound(saryFileSort, 1)
        If LCase(Right(saryFileSort(lFile), 4)) = ".pdf" Then
            If bFirstDoc Then
                bFirstDoc = False
                Set oMainDoc = CreateObject("AcroExch.PDDoc")
                lRet = oMainDoc.Open(sRawPDFFilesDir & saryFileSort(lFile))
                Set oJSO = oMainDoc.GetJSObject
                Set oBookMarkRoot = oJSO.BookMarkRoot
                sBMName = saryFileSort(lFile)
                lPos = InStr(1, sBMName, "_{", vbBinaryCompare)
                If lPos > 0 Then
                    sBMName = Left(sBMName, lPos - 1) & ".pdf"
                End If
                lRet = oBookMarkRoot.CreateChild(sBMName, "this.pageNum =0", lFile - 1)
            Else
                Set oTempDoc = CreateObject("AcroExch.PDDoc")
                lRet = oTempDoc.Open(sRawPDFFilesDir & saryFileSort(lFile))
                'get the Book mark page number before the actual instert of new pages
                lBMPageNo = oMainDoc.GetNumPages
                lInsertPageAfter = lBMPageNo - 1
                lNumPages = oTempDoc.GetNumPages
                lRet = oMainDoc.InsertPages(lInsertPageAfter, oTempDoc, 0, lNumPages, 0)
                oTempDoc.Close
                If lRet = 0 Then
                    sBMName = saryFileSort(lFile)
                    lPos = InStr(1, sBMName, "_{", vbBinaryCompare)
                    If lPos > 0 Then
                        sBMName = Left(sBMName, lPos - 1) & ".pdf"
                    End If
                    'Need to copy the errored document over to be included in the enitre document
                    goUtil.utCopyFile sRawPDFFilesDir & saryFileSort(lFile), sSinglePDFOutputDir & "\" & sBMName
                    sBMName = "PDF Insert Page Error_" & sBMName
                Else
                    sBMName = saryFileSort(lFile)
                    lPos = InStr(1, sBMName, "_{", vbBinaryCompare)
                    If lPos > 0 Then
                        sBMName = Left(sBMName, lPos - 1) & ".pdf"
                    End If
                End If
                lRet = oBookMarkRoot.CreateChild(sBMName, "this.pageNum =" & lBMPageNo, lFile - 1)
            End If
        End If
    Next
    
    lRet = oMainDoc.Save(1, sSinglePDFOutputDir & "\" & sSinglePDFOutputName)
    oMainDoc.Close
    
    MergePDFFiles = True
    
CLEAN_UP:
    Set oFolder = Nothing
    Set oFile = Nothing
    Set oFSO = Nothing
    Set oBookMarkRoot = Nothing
    Set oJSO = Nothing
    Set oMainDoc = Nothing
    Set oTempDoc = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    MergePDFFiles = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Function MergePDFFiles" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    moUL_ErrorMess sMess
    Set oFolder = Nothing
    Set oFile = Nothing
    Set oFSO = Nothing
    Set oBookMarkRoot = Nothing
    Set oJSO = Nothing
    Set oMainDoc = Nothing
    Set oTempDoc = Nothing
End Function


