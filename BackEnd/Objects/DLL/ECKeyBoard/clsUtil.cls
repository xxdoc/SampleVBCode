VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'EVENTS

Public Event XactimateProgress(ByVal sMessage As String)

'Easy Claim Printer Type
Public Type ECPRINTER
    PRINTER_NAME As String
    PRINTER_PORT As String
    PRINTER_DRIVER As String
End Type

'
' Constants used by FindWindowPartial
'
Public Enum FindWindowPartialTypes
   FwpStartsWith = 0
   FwpContains = 1
   FwpMatches = 2
End Enum

'Used in ECKeyBoard Loss Reports List View
Public Enum LossReports
    DateAsgn = 1
    DateAsgnSort
    AssignmentType
    Status
    CatName
    CatCode
    Adjuster
    ACID
    CLIENTNUM
    IBNUM
    InsuredName
    HPhone
    WPhone
    RFormat
    RSort ' Hidden
    RKey 'Hidden
End Enum

'Used in ECKeyBoard Loss Reports Print Options
Public Enum PrintFormat
    RawText = 0
    Translated
End Enum

'Used in ECKeyBoard LossReports Image list
Public Enum LRPic
    lrError = 1
    lrUnknown
    Found
    SearchColumn
    lrHTMLHelp
    lrFORMAT01
    lrFORMAT02
    lrFORMAT03
    lrFORMAT04
    lrFORMAT05
    lrFORMAT06
    lrFORMAT07
    lrFORMAT08
    lrFORMAT09
    lrFORMAT10
    lrFORMAT11
    lrFORMAT12
    lrFORMAT13
    lrFORMAT14
    lrFORMAT15
    lrFORMAT16
    lrFORMAT17
    lrFORMAT18
    lrFORMAT19
    lrFORMAT20
    lrFORMAT21
    lrFORMAT22
    lrFORMAT23
    lrFORMAT24
    lrFORMAT25
    lrFORMAT26
    lrFORMAT27
    lrFORMAT28
    lrFORMAT29
    lrFORMAT30
    lrFORMATpdf
End Enum

'Used in ECKeyBoard Adjuster List View
Public Enum ADJ
    CRID = 1
    FACT
    Email
    ADJUpdateEmail
    ADJFName 'FirstName
    ADJLName 'LastName
    ADJSSN
    ADJPassword
    ADJContactPhone
    ADJTeamLeader
    ADJDateLastUpdated
    ADJLicDaysLeft
    ADJAPPVSInfo 'Application VS info
    Dirty 'Hidden
    RSort ' Hidden
    RKey 'Hidden
End Enum

'Used in ECKeyBoard Adjuster Image list
Public Enum ADJPic
    Pointer = 1
    Found
    SearchColumn
    DeleteMe
    UnDelete
End Enum

'If Data in a collection changed
Public Enum IsDirty
    NoChange = 0
    AddMe
    DeleteMe
End Enum

'Need UDT for Adjuster Table
Public Type udtAdjuster
    CRID As String
    FACT As String
    EmailAddress As String
    ADJUpdateEmail As String
    ADJFName As String
    ADJLName As String
    ADJSSN As String
    ADJPassword As String
    ADJContactPhone As String
    ADJTeamLeader As String
    ADJDateLastUpdated As Date
    ADJLicDaysLeft As String
    ADJAPPVSInfo  As String 'Application VS info
    IsDirty As IsDirty
    Key As String
End Type

'Needed for Print Report Event
Public Type udtAppDoc
    DocName As String
    Selected As Boolean
End Type

'Used to define which Report this is if it even is being chained
Public Enum ChainType
    NotChain = 0
    FirstInChain
    NextLink
End Enum

'Used for Export PDF
Public Enum ExportType
    ARExcel = 0
    ARPdf
    ARRtf
    ARHtml
    ARText
End Enum

'Used to Pass various misc Paramers
Public Type udtParameter
    ParamName As String
    ParamValue As Variant
End Type

'Used in frmECTray Tree Image List
Public Enum PicTree
    A01_ListView = 1
    A02_Add
    A03_Accounting
    A04_FeeSched
    A05_CatPref
    A06_Communications
    A07_Printer
    A08_MakeCurrent
    A09_ClosedFolder
    A10_OpenFolder
    A11_Support
    A12_GlobalPref
    A13_HandShake
    A14_ClosedHurc
    A15_OpenHurc
    A16_SetUpNewCat
    A17_RecycleBinFull
    A18_RecycleBinEmpty
    A19_DeleteAllFromRecycleBin
    A20_DeleteCat
    A21_RestoreCat
    A22_SendCatToXactimate
    A23_UpdateCurrentCat
    A24_SendCatToDisk
    A25_CreateCatBackup
    A26_ExportCat
    A27_GetCatFromDisk
    A28_GetCatBackup
    A29_ImportCat
    A30_Item0
    A31_Item1
    A32_Item2
    A33_Item3
    A34_Item4
    A35_Item5
    A36_Item6
    A37_Item7
    A38_Item8
    A39_Item9
    A40_Item10
    A41_ViewLossReports
    A42_closedHurc_Inactive
    A43_OpenHurc_Inactive
End Enum

'Used for Moving Items in ListView Up or Down
Public Enum MoveListItem
    MoveUp = -1
    MoveDown = 1
End Enum

'Assignment Types
Public Enum AssgnType
    iAssignmentType_Property = 1
    iAssignmentType_Auto = 2
End Enum

'Assignment Status
Public Enum AssgnStatus
    iAssignmentsStatus_CLOSED = 1
    iAssignmentsStatus_DELETED = 2
    iAssignmentsStatus_NEW = 3
    iAssignmentsStatus_REASSIGNED = 4
    iAssignmentsStatus_REOPEN = 5
    iAssignmentsStatus_INTERIM = 6
    iAssignmentsStatus_PENDING = 7
End Enum

'Member Vars used for Passing Global Vars
Private mbValidLic As Boolean                   'True if License is Valid
Private mbInitLic As Boolean                    'True if Initializing Lic
Private msAPPEXEName As String                  'Application Name IE EasyClaim
Private msMainAppEXEName As String              'Main Application name
Private msCommandString As String               'Command String passed to Application
Private msCompanyCarrierPrefix As String        'This is Used To Help Late Bind the Carrier Object
Private msInstallDir As String                  'Installed directory
Private msCurCatDir As String                   'The Current Cat Directory
Private msCurCompany As String                  'The Current Company ID
Private msCurCarDBName As String                'The Current Carrier DBName , 'IE Farmers, Safeco,
Private msCurCar As String                      'The Current Carrier ID
Private msCurCat As String                      'The Current CAT ID
Private msCurUsersID As String                  'The Current USersID
Private mbUCText As Boolean
Private mdicDB_PASSWORD As scripting.Dictionary 'Dictionary Object containing Encrypted Passwords used for DB and Files

'Member Vars used for Passing Global Objects
Private mclsKeyBoard As clsKeyBoard
Private moCurCarList As V2ECKeyBoard.clsCarLists          'The Current Carrier List Object (List Object is the main interface for The Carrier)
Private moSP As V2ECKeyBoard.clsSP                        'Spelling Object
Private WithEvents moXact As V2ECKeyBoard.clsXact                    'Send To Xactimate Object
Attribute moXact.VB_VarHelpID = -1
Private mWS As dao.Workspace                            'Access Work SPAce
Private mCurDB As dao.Database                          'Current Carrier Data Base
Private mMainDB As dao.Database
Private moECKeyBoardList As V2ECKeyBoard.clsLists         'Misc Stuff EasyCLaim (Cold Fusion Web Site uses the same class directly)
Private mARV As Object                                     'Active Reports Viewer Object
Private moProgForm As V2ECKeyBoard.clsProgForm            'Class the Wraps the Progress form
Private moTimeZone As V2ECKeyBoard.clsTimeZone          'Class for figuring out timzone stuff like GMT offest, Leap year, etc

'Main Easy Claim Objects
Private mfrmECTray As Form            'Main Easy Claim Form
Private mfrmCommStatus As Form        'Communications Status Form
Private mfrmWallPaper As Form         'Optional Background Work Area
'Paths
Private msFTPDLPath As String 'Download Path
Private msFTPULPath As String 'Upload Path
Private msAttachReposPath As String
Private msPhotoReposPath As String
Private msSPPath As String 'Software Package path

Public Property Get tzTimeZone() As clsTimeZone
    Set tzTimeZone = moTimeZone
End Property

Public Property Let SPPath(psPath As String)
    msSPPath = psPath
End Property
Public Property Get SPPath() As String
    SPPath = msSPPath
End Property

Public Property Let PhotoReposPath(psPath As String)
    msPhotoReposPath = psPath
End Property
Public Property Get PhotoReposPath() As String
    PhotoReposPath = msPhotoReposPath
End Property

Public Property Let AttachReposPath(psPath As String)
    msAttachReposPath = psPath
End Property
Public Property Get AttachReposPath() As String
    AttachReposPath = msAttachReposPath
End Property

Public Property Let FTPULPath(psPath As String)
    msFTPULPath = psPath
End Property
Public Property Get FTPULPath() As String
    FTPULPath = msFTPULPath
End Property

Public Property Let FTPDLPath(psPath As String)
    msFTPDLPath = psPath
End Property
Public Property Get FTPDLPath() As String
    FTPDLPath = msFTPDLPath
End Property

Public Property Let gbValidLic(pbFlag As Boolean)
    mbValidLic = pbFlag
End Property
Public Property Get gbValidLic() As Boolean
    gbValidLic = mbValidLic
End Property
    
Public Property Let gbInitLic(pbFlag As Boolean)
    mbInitLic = pbFlag
End Property
Public Property Get gbInitLic() As Boolean
    gbInitLic = mbInitLic
End Property

Public Property Let gsAppEXEName(psName As String)
    msAPPEXEName = psName
End Property
Public Property Get gsAppEXEName() As String
    gsAppEXEName = msAPPEXEName
End Property

Public Property Let gsMainAppEXEName(psName As String)
    msMainAppEXEName = psName
End Property
Public Property Get gsMainAppEXEName() As String
    gsMainAppEXEName = msMainAppEXEName
End Property

Public Property Let gsCommandString(psName As String)
    msCommandString = psName
End Property
Public Property Get gsCommandString() As String
    gsCommandString = msCommandString
End Property

Public Property Let gsCarPrefix(psName As String)
    msCompanyCarrierPrefix = psName
End Property
Public Property Get gsCarPrefix() As String
    gsCarPrefix = msCompanyCarrierPrefix
End Property

Public Property Let gsInstallDir(psDir As String)
    msInstallDir = psDir
End Property
Public Property Get gsInstallDir() As String
    gsInstallDir = msInstallDir
End Property

Public Property Let gsCurCatDir(psDir As String)
    msCurCatDir = psDir
End Property
Public Property Get gsCurCatDir() As String
    gsCurCatDir = msCurCatDir
End Property

Public Property Let gsCurCarDBName(psName As String)
    msCurCarDBName = psName
End Property
Public Property Get gsCurCarDBName() As String
    gsCurCarDBName = msCurCarDBName
End Property

Public Property Let gsCurCompany(psID As String)
    msCurCompany = psID
End Property
Public Property Get gsCurCompany() As String
    gsCurCompany = msCurCompany
End Property

Public Property Let gsCurCar(psDir As String)
    msCurCar = psDir
End Property
Public Property Get gsCurCar() As String
    gsCurCar = msCurCar
End Property

Public Property Let gsCurCat(psDir As String)
    msCurCat = psDir
End Property
Public Property Get gsCurCat() As String
    gsCurCat = msCurCat
End Property

Public Property Let gsCurUsersID(psDir As String)
    msCurUsersID = psDir
End Property
Public Property Get gsCurUsersID() As String
    gsCurUsersID = msCurUsersID
End Property


'Public Property Get gsCurCat() As String
'    Dim vAry As Variant
'    If msCurCatDir = vbNullString Then
'        gsCurCat = vbNullString
'    Else
'        vAry = Split(msCurCatDir, "\")
'        gsCurCat = vAry(UBound(vAry))
'    End If
'End Property
'
'Public Property Get gsCurCar() As String
'    Dim vAry As Variant
'    If msCurCatDir = vbNullString Then
'        gsCurCar = vbNullString
'    Else
'        vAry = Split(msCurCatDir, "\")
'        gsCurCar = vAry(UBound(vAry) - 1)
'    End If
'
'End Property

Public Property Let INIT_DB_PASSWORD(psInitKey As String)
    On Error Resume Next
    Dim sInitKey As String
    
    '12.11.2002 Data Base Password
    'Key "1" is the Latest Password.  If you need to change the Latest password
    'Be sure to move the Current Key "1" Password into Key "2" position and so on.
    'THIS IS THE ONLY PLACE IN CODE THAT NEEDS TO BE UPDATED FOR A PASSWORD CHANGE!
    '***Note***Communications Logon and Passwords are totally different and are Accomplished
    'in frmCommStatus (FTP_LOGON_NAME FTP_LOGON_PASS)
    '***End Note***
    sInitKey = psInitKey
    If Decode(sInitKey) = Decode(" c91223-n3j q02223-i2j b81223-t4k s12223-i1k") Then
        Set mdicDB_PASSWORD = New scripting.Dictionary
        mdicDB_PASSWORD.Add "1", " s70223-h51y r21223-o01f q12223-a1z h41223-g8i m02223-i2s r51223-t7e e90223-t31o i31223-a9e s80223-e41y u71223-n5j z81223-u4q j60223-i61n l91223-t3m m01223-v21l r61223-t6k v11223-c11k"
    End If
    
    Exit Property
EH:
    Err.Clear
End Property

Public Property Let gbUCText(pbFlag As Boolean)
    mbUCText = pbFlag
End Property
Public Property Get gbUCText() As Boolean
    gbUCText = mbUCText
End Property

Public Property Let goCurCarList(poCarList As V2ECKeyBoard.clsCarLists)
    Set moCurCarList = poCarList
End Property
Public Property Set goCurCarList(poCarList As V2ECKeyBoard.clsCarLists)
    Set moCurCarList = poCarList
End Property
Public Property Get goCurCarList() As V2ECKeyBoard.clsCarLists
    Set goCurCarList = moCurCarList
End Property

Public Property Let goSP(poSP As V2ECKeyBoard.clsSP)
    Set moSP = poSP
End Property
Public Property Set goSP(poSP As V2ECKeyBoard.clsSP)
    Set moSP = poSP
End Property
Public Property Get goSP() As V2ECKeyBoard.clsSP
    Set goSP = moSP
End Property

Public Property Let goXact(poXACT As V2ECKeyBoard.clsXact)
    Set moXact = poXACT
End Property
Public Property Set goXact(poXACT As V2ECKeyBoard.clsXact)
    Set moXact = poXACT
End Property
Public Property Get goXact() As V2ECKeyBoard.clsXact
    Set goXact = moXact
End Property

Public Property Let gWS(pWS As dao.Workspace)
    Set mWS = pWS
End Property
Public Property Set gWS(pWS As dao.Workspace)
    Set mWS = pWS
End Property
Public Property Get gWS() As dao.Workspace
    Set gWS = mWS
End Property

Public Property Let gCurDB(pCurDB As dao.Database)
    Set mCurDB = pCurDB
End Property
Public Property Set gCurDB(pCurDB As dao.Database)
    Set mCurDB = pCurDB
End Property
Public Property Get gCurDB() As dao.Database
    Set gCurDB = mCurDB
End Property

Public Property Let gMainDB(pMainDB As dao.Database)
    Set mMainDB = pMainDB
End Property
Public Property Set gMainDB(pMainDB As dao.Database)
    Set mMainDB = pMainDB
End Property
Public Property Get gMainDB() As dao.Database
    Set gMainDB = mMainDB
End Property

Public Property Get ECMainDBVersion() As String
    'Set the ECMain.mdb Version Number
    'This number should be the same as the Template Version Released
    'with this application version
    'Version 1 '2/10/2005 15:45
    'Version 2 '3/3/2005 19:30 ' Added StatusAlias to Status table
    ECMainDBVersion = 2
End Property

Public Property Let goECKeyBoardList(poECKeyBoardList As V2ECKeyBoard.clsLists)
    Set moECKeyBoardList = poECKeyBoardList
End Property
Public Property Set goECKeyBoardList(poECKeyBoardList As V2ECKeyBoard.clsLists)
    Set moECKeyBoardList = poECKeyBoardList
End Property
Public Property Get goECKeyBoardList() As V2ECKeyBoard.clsLists
    Set goECKeyBoardList = moECKeyBoardList
End Property

Public Property Let gfrmECTray(poForm As Object)
    Set mfrmECTray = poForm
End Property
Public Property Set gfrmECTray(poForm As Object)
    Set mfrmECTray = poForm
End Property
Public Property Get gfrmECTray() As Object
    Set gfrmECTray = mfrmECTray
End Property

Public Property Let gfrmCommStatus(poForm As Object)
    Set mfrmCommStatus = poForm
End Property
Public Property Set gfrmCommStatus(poForm As Object)
    Set mfrmCommStatus = poForm
End Property
Public Property Get gfrmCommStatus() As Object
    Set gfrmCommStatus = mfrmCommStatus
End Property

Public Property Let gfrmWallPaper(poForm As Object)
    Set mfrmWallPaper = poForm
End Property
Public Property Set gfrmWallPaper(poForm As Object)
    Set mfrmWallPaper = poForm
End Property
Public Property Get gfrmWallPaper() As Object
    Set gfrmWallPaper = mfrmWallPaper
End Property

Public Property Let goProgForm(poProgForm As V2ECKeyBoard.clsProgForm)
    Set moProgForm = poProgForm
End Property
Public Property Set goProgForm(poProgForm As V2ECKeyBoard.clsProgForm)
    Set moProgForm = poProgForm
End Property
Public Property Get goProgForm() As V2ECKeyBoard.clsProgForm
    Set goProgForm = moProgForm
End Property

Public Property Let gARV(poARV As Object)
    Set mARV = poARV
End Property
Public Property Set gARV(poARV As Object)
    Set mARV = poARV
End Property
Public Property Get gARV() As Object
    Set gARV = mARV
End Property

Public Property Get goForms() As Object
    Set goForms = Forms
End Property

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsUtil"
End Property

Public Sub utCreateWorkSpace()
    Set mWS = dao.CreateWorkspace("", "admin", "", dbUseJet)
End Sub

Public Sub BuildPathsEasyClaim()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Build Directory Paths
    msFTPDLPath = goUtil.gsInstallDir & "\FTP\DownLoad\"
    msFTPULPath = goUtil.gsInstallDir & "\FTP\UpLoad\"
    msSPPath = goUtil.gsInstallDir & "\SP\"
    msAttachReposPath = goUtil.gsInstallDir & "\AttachRepos\"
    msPhotoReposPath = goUtil.gsInstallDir & "\PhotoRepos\"
    
    'Be sure the entire structure exists
    If Not goUtil.utFileExists(msFTPDLPath, True) Then
        goUtil.utMakeDir goUtil.gsInstallDir & "\FTP\"
        goUtil.utMakeDir msFTPDLPath
    End If
    If Not goUtil.utFileExists(msFTPULPath, True) Then
        goUtil.utMakeDir goUtil.gsInstallDir & "\FTP\"
        goUtil.utMakeDir msFTPULPath
    End If
    'SP for Application, DataBase, Document and RegSetting
    If Not goUtil.utFileExists(msSPPath & "Application\SP\", True) Then
        goUtil.utMakeDir msSPPath
        goUtil.utMakeDir msSPPath & "Application\"
        goUtil.utMakeDir msSPPath & "Application\SP\"
    End If
    If Not goUtil.utFileExists(msSPPath & "DataBase\SP\", True) Then
        goUtil.utMakeDir msSPPath
        goUtil.utMakeDir msSPPath & "DataBase\"
        goUtil.utMakeDir msSPPath & "DataBase\SP\"
    End If
    If Not goUtil.utFileExists(msSPPath & "Document\SP\", True) Then
        goUtil.utMakeDir msSPPath
        goUtil.utMakeDir msSPPath & "Document\"
        goUtil.utMakeDir msSPPath & "Document\SP\"
    End If
    If Not goUtil.utFileExists(msSPPath & "RegSetting\SP\", True) Then
        goUtil.utMakeDir msSPPath
        goUtil.utMakeDir msSPPath & "RegSetting\"
        goUtil.utMakeDir msSPPath & "RegSetting\SP\"
    End If
    If Not goUtil.utFileExists(msAttachReposPath, True) Then
        goUtil.utMakeDir msAttachReposPath
    End If
    If Not goUtil.utFileExists(msPhotoReposPath, True) Then
        goUtil.utMakeDir msPhotoReposPath
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub BuildPathsEasyClaim" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub BuildPathsWebControl()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim saryPath() As String
    Dim lCount As Long
    Dim sFolder As String
    Dim sPath As String
    
    'Build Directory Paths
    
    'FTP Upload
    msFTPDLPath = vbNullString
    msFTPULPath = GetSetting("V2WebControl", "Dir", "FTPSitePath", vbNullString)
    
    If msFTPULPath = vbNullString Then
        'Cant build the paths until the setting for FTPSitePath is there
        Exit Sub
    End If
    'Take Off Trailing "\"
    If Right(msFTPULPath, 1) = "\" Then
        msFTPULPath = left(msFTPULPath, Len(msFTPULPath) - 1)
    End If
    
    'Photo and Attahc Repos under FTP Upload
    msAttachReposPath = msFTPULPath & "\AttachRepos"
    msPhotoReposPath = msFTPULPath & "\PhotoRepos"
    
    'Software Package
    msSPPath = left(msFTPULPath, InStr(1, msFTPULPath, "\Assignments", vbTextCompare) - 1)
    msSPPath = msSPPath & "\SP"
    
    
    'Be sure the entire structure exists
    
    'Build the Ftp Upload Path if it does not already exist
    saryPath() = Split(msFTPULPath, "\")
    For lCount = LBound(saryPath, 1) To UBound(saryPath, 1)
        sPath = sPath & saryPath(lCount) & "\"
        'Only start checking the path and build if the Length is > 3
        'This will account for both Mapped and UNC ie...
        'Mapped C:\  = Len 3
        If Len(sPath) > 3 Then
            If Not goUtil.utFileExists(sPath, True) Then
                goUtil.utMakeDir sPath
            End If
        End If
    Next
    'Build the Attach and Photo repos under FTP Upload
    If Not goUtil.utFileExists(msAttachReposPath, True) Then
        goUtil.utMakeDir msAttachReposPath
    End If
    If Not goUtil.utFileExists(msPhotoReposPath, True) Then
        goUtil.utMakeDir msPhotoReposPath
    End If
    
    'Build the Software Package Path if it does not already exist
    sPath = vbNullString
    saryPath() = Split(msSPPath, "\")
    For lCount = LBound(saryPath, 1) To UBound(saryPath, 1)
        sPath = sPath & saryPath(lCount) & "\"
        'Only start checking the path and build if the Length is > 3
        'This will account for both Mapped and UNC ie...
        'Mapped C:\  = Len 3
        If Len(sPath) > 3 Then
            If Not goUtil.utFileExists(sPath, True) Then
                goUtil.utMakeDir sPath
            End If
        End If
    Next
    
     'SP for Application, DataBase, Document and RegSetting
    If Not goUtil.utFileExists(msSPPath & "\Application\SP\", True) Then
        goUtil.utMakeDir msSPPath & "\Application\"
        goUtil.utMakeDir msSPPath & "\Application\SP\"
    End If
    If Not goUtil.utFileExists(msSPPath & "\DataBase\SP\", True) Then
        goUtil.utMakeDir msSPPath & "\DataBase\"
        goUtil.utMakeDir msSPPath & "\DataBase\SP\"
    End If
    If Not goUtil.utFileExists(msSPPath & "\Document\SP\", True) Then
        goUtil.utMakeDir msSPPath & "\Document\"
        goUtil.utMakeDir msSPPath & "\Document\SP\"
    End If
    If Not goUtil.utFileExists(msSPPath & "\RegSetting\SP\", True) Then
        goUtil.utMakeDir msSPPath & "\RegSetting\"
        goUtil.utMakeDir msSPPath & "\RegSetting\SP\"
    End If
    
    'Add "\" to end
    msFTPDLPath = vbNullString
    
    msFTPULPath = msFTPULPath & "\"
    
    'Photo and Attahc Repos under FTP Upload
    msAttachReposPath = msAttachReposPath & "\"
    msPhotoReposPath = msPhotoReposPath & "\"
    
    'Software Package
    msSPPath = msSPPath & "\"
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub BuildPathsEasyClaim" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub utOpenDatabase(pDB As dao.Database, psDBPath As String, pvOptions As Variant, pbReadOnly As Boolean)
    On Error Resume Next
    Dim lKey As Long
    Dim sKey As String
    Dim sPassword As String
    Dim lErrorNum As Long
    Dim sErrorDescription As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
Set_Key:
    On Error Resume Next
    lKey = lKey + 1
    sKey = CStr(lKey)
    sPassword = DB_PASSWORD(sKey)
    sPassword = Decode(sPassword)
    If Err.Number <> 0 Then
        lErrorNum = Err.Number
        sErrorDescription = Err.Description
        Err.Clear
        On Error GoTo EH
        Err.Raise lErrorNum, , sErrorDescription & vbCrLf & "Invalid Password!" & vbCrLf & msClassName & vbCrLf & "Public Sub utOpenDatabase"
    End If
    Set pDB = mWS.OpenDatabase(psDBPath, pvOptions, pbReadOnly, ";PWD =" & sPassword)
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
            Err.Raise lErrorNum, , sErrorDescription & vbCrLf & msClassName & vbCrLf & "Public Sub utOpenDatabase"
        End If
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, sErrDesc
End Sub

Public Sub utOpenDatabaseADOConn(poADOConn As ADODB.Connection, psDBPath As String)
    On Error Resume Next
    Dim lKey As Long
    Dim sKey As String
    Dim sPassword As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
Set_Key:
    On Error Resume Next
    lKey = lKey + 1
    sKey = CStr(lKey)
    sPassword = DB_PASSWORD(sKey)
    sPassword = Decode(sPassword)
    If Err.Number <> 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        Err.Clear
        On Error GoTo EH
        Err.Raise lErrNum, , sErrDesc & vbCrLf & "Invalid Password!" & vbCrLf & msClassName & vbCrLf & "Public Sub utOpenDatabaseADOConn"
    End If
    poADOConn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
    poADOConn.ConnectionString = "Data Source=" & psDBPath
    poADOConn.Properties("Jet OLEDB:Database Password") = sPassword
    poADOConn.Open
    If Err.Number <> 0 Then
        'Check the Error Message, if it is an invalid Password...
        'Keep Checking the Dictionary until one works.
        If Err.Number = -2147217843 Then
            Err.Clear
            GoTo Set_Key
        Else
            lErrNum = Err.Number
            sErrDesc = Err.Description
            Err.Clear
            On Error GoTo EH
            Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub utOpenDatabaseADOConn"
        End If
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, sErrDesc
End Sub

Public Sub utAlwaysOnTop(pForm As Object, pbSetOnTop As Boolean, Optional plHwnd As Long = 0)
    AlwaysOnTop pForm, pbSetOnTop, plHwnd
End Sub

Public Sub utShowError(psAppEXEName As String, pErr As ErrObject, psProc As String, _
                        Optional pFrmOwner As Object = Nothing, _
                        Optional psMod As String, _
                        Optional plRetErrNum As Long, _
                        Optional psRetErrDescription As String)
    Dim sErrMess As String

    
    plRetErrNum = pErr.Number
    psRetErrDescription = pErr.Description
    
    sErrMess = Now() & vbCrLf
    sErrMess = sErrMess & "AppName: " & psAppEXEName & vbCrLf
    sErrMess = sErrMess & "ClassName: " & psMod & vbCrLf
    sErrMess = sErrMess & "ProcName: " & psProc & vbCrLf
    sErrMess = sErrMess & "ERROR # " & plRetErrNum & vbCrLf
    sErrMess = sErrMess & psRetErrDescription & vbCrLf & vbCrLf
    If Not goUtil.goProgForm Is Nothing Then
        goUtil.goProgForm.ShowForm False
    End If
    MsgBox sErrMess, vbCritical + vbOKOnly, "ERROR in " & psProc
End Sub

Public Sub utSelText(pTextBox As Object)
    SelText pTextBox
End Sub

Public Sub utUCText(pControl As Object)
    UCText pControl
End Sub

Public Function utFormWinRegPos(psAppEXEName As String, poMyForm As Object, Optional pbSave As Boolean, _
                              Optional pfrmOffset As Object, Optional pctrlOffset As Object, _
                              Optional pbUseFullCaption As Boolean = True, Optional pbUseFrmName As Boolean) As Boolean
    utFormWinRegPos = FormWinRegPos(psAppEXEName, poMyForm, pbSave, pfrmOffset, pctrlOffset, pbUseFullCaption, pbUseFrmName)
                              
End Function

Public Function utFileExists(strFile As String, Optional pbDirOnly As Boolean) As Boolean
    utFileExists = FileExists(strFile, pbDirOnly)
End Function

Public Function utGetFileData(psFilePath As String, Optional pbLock As Boolean = False) As String
    utGetFileData = GetFileData(psFilePath, pbLock)
End Function

Public Sub utSaveFileData(psFilePath As String, psFileData As String, Optional psDelimeter As String, Optional pbLock As Boolean = False)
    SaveFileData psFilePath, psFileData, psDelimeter, pbLock
End Sub

Public Function utFindSetForm(goForms As Object, psFormName As String, poForm As Object) As Boolean
    utFindSetForm = FindSetForm(goForms, psFormName, poForm)
End Function

Public Function utFormExists(goForms As Object, psFormName As String) As Boolean
    utFormExists = FormExists(goForms, psFormName)
End Function

Public Function utCloseApplication(ByVal psAppCaption As String, plHwnd As Long) As Boolean
    utCloseApplication = CloseApplication(psAppCaption, plHwnd)
End Function

Public Function utGetCaption(ByVal lHwnd As Long) As String
    utGetCaption = GetCaption(lHwnd)
End Function

Public Function utFindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    utFindWindow = FindWindow(lpClassName, lpWindowName)
End Function

Public Function utFindNextWindow(psAppCaption As String, plhWndStart As Long, plhWndFound As Long) As Boolean
    utFindNextWindow = FindNextWindow(psAppCaption, plhWndStart, plhWndFound)
End Function

Public Function utLookForWindow(psWindowCaption As String, Optional plTimeOutSeconds As Long = 60) As Boolean
    On Error GoTo EH
    Dim bFound As Boolean
    Dim lSleep As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    bFound = True
    Do Until WindowFound(psWindowCaption)
        lSleep = lSleep + 1
        If lSleep > plTimeOutSeconds Then
            bFound = False
            Exit Do
        End If
    Loop
    
    utLookForWindow = bFound
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function utLookForWindow" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function utFindWindowPartial(AppTitle As String, _
   Optional Method As FindWindowPartialTypes = FwpStartsWith, _
   Optional CaseSensitive As Boolean = False, _
   Optional MustBeVisible As Boolean = False) As Long
   utFindWindowPartial = FindWindowPartial(AppTitle, Method, CaseSensitive, MustBeVisible)
End Function

Public Function utGetECSCryptSetting(psAPP As String, psSection As String, psKey As String, _
                                   Optional pvDefault As Variant = vbNullString) As Variant
    utGetECSCryptSetting = GetECSCryptSetting(psAPP, psSection, psKey, pvDefault)
End Function

Public Sub utSaveECSCryptSetting(psAPP As String, psSection As String, psKey As String, psSetting As String)
    SaveECSCryptSetting psAPP, psSection, psKey, psSetting
End Sub

Public Sub utEnterUserPass(psSection As String, psUserName As String, pvPass As Variant, _
                         Optional poForm As Object)
    EnterUserPass psSection, psUserName, pvPass, poForm
End Sub

Public Function utGetAppVSInfo(psAppEXEName As String, psAppPath As String) As String
    utGetAppVSInfo = GetAppVSInfo(psAppEXEName, psAppPath)
End Function

Public Function utGetECSWinSysObjectsPaths() As Variant
    utGetECSWinSysObjectsPaths = GetECSWinSysObjectsPaths
End Function

Public Function utGetWinOSVersion() As String
    utGetWinOSVersion = GetWinOSVersion
End Function

Public Function utGetMyComputerName() As String
   utGetMyComputerName = GetMyComputerName
End Function

Public Function utGetMyUserName() As String
   utGetMyUserName = GetMyUserName
End Function

Public Sub utErrorLog(poErr As ErrObject, psAppName As String, psClassName As String, psProcName As String, Optional pbShowError As Boolean = True)
    Dim sErrMess As String
    Dim sAdjName As String
    Dim sFileName As String
    Dim sErrorNum As String
    Dim sErrorDescp As String
    Dim sErrorLogDir As String
    Dim sSilentError As String
    
    
    'Set these before the Error handler for this Proc
    sErrorNum = poErr.Number
    sErrorDescp = poErr.Description
    'Show the error
    sSilentError = GetSetting("ECS", "Dir", "SILENT_ERROR", "True")
    
    If sSilentError = "False" Then
        If pbShowError Then
            ShowError psAppName, poErr, psProcName, , psClassName
        End If
    End If
    
    'Need to set this here so we won't clear the first Error
    On Error GoTo EH
    sErrMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sErrMess = sErrMess & Now() & vbCrLf
    sErrMess = sErrMess & "AppName: " & psAppName & vbCrLf
    sErrMess = sErrMess & "ClassName: " & psClassName & vbCrLf
    sErrMess = sErrMess & "ProcName: " & psProcName & vbCrLf
    sErrMess = sErrMess & "ERROR # " & sErrorNum & " " & Now & vbCrLf
    sErrMess = sErrMess & sErrorDescp & vbCrLf
    sErrMess = sErrMess & "<<<<<<<<<<  END ERROR MESSAGE  >>>>>>>>>>" & vbCrLf & vbCrLf
    sErrMess = sErrMess & "<<<<<<<<<< BEGIN WIN OS INFO >>>>>>>>>>" & vbCrLf
    sErrMess = sErrMess & GetWinOSVersion & vbCrLf
    sErrMess = sErrMess & "<<<<<<<<<<  END WIN OS INFO  >>>>>>>>>>" & vbCrLf
    
    'Get the Error Log Install Dir
    sErrorLogDir = GetSetting("ECS", "Dir", "ERRORLOG_DIR", App.Path)
    
    'Check to be sure the Error Log Folder is there
    If Not FileExists(sErrorLogDir & "\ErrorLog", True) Then
        MakeDir sErrorLogDir & "\ErrorLog"
    End If
    'Get Adjuster Name
    sAdjName = GetSetting(gsAppEXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", psAppName)
    sAdjName = sAdjName & "_"
    sAdjName = sAdjName & GetSetting(gsAppEXEName, "GENERAL", "ADJUSTOR_LAST_NAME", vbNullString)
    
    sFileName = sErrorLogDir & "\ErrorLog\" & sAdjName & "_" & Format(Now(), "YYMMDDHHMMSS") & "_Error.Log"
    SaveFileData sFileName, sErrMess
    Exit Sub
EH:
    MsgBox App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub ErrorLog" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Public Function utCleanSQLString(pvText As Variant) As Variant
    utCleanSQLString = CleanSQLString(pvText)
End Function

Public Function utCleanValString(psValText As String) As String
    utCleanValString = CleanValString(psValText)
End Function

Public Sub utCleanValTextBox(pvText As Variant)
    CleanValTextBox pvText
End Sub


Public Function utSaveDefaultPrinterSettings(psAppEXEName As String, Optional psDefaultPrinterName As String) As Boolean
    utSaveDefaultPrinterSettings = SaveDefaultPrinterSettings(psAppEXEName, psDefaultPrinterName)
End Function

Public Function utGetDefaultPrinterSettings(psAppEXEName As String) As V2ECKeyBoard.ECPRINTER
    utGetDefaultPrinterSettings = GetDefaultPrinterSettings(psAppEXEName)
End Function

Public Sub utLoadSP()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    If Not moSP Is Nothing Then
        moSP.CLEANUP
        Set moSP = Nothing
    End If
    LoadSP msInstallDir, moSP
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub utLoadSP"
End Sub

Public Sub utLoadCarriers(pcboList As Object)
    LoadCarriers pcboList
End Sub

Public Sub utLoadTOL(pcboList As Object)
    LoadTOL pcboList
End Sub

Public Sub utLoadStates(pcboList As Object)
    LoadStates pcboList
End Sub

Public Function utGetPath(psAppEXEName As String, psName As String, psMess As String, psFileMess As String, psDeFaultPath As String, plHwnd As Long, _
                         Optional psFilter As String = vbNullString, _
                         Optional psSelFile As String, _
                         Optional plFlags As Long, _
                         Optional pbCenterForm As Boolean = True, _
                         Optional pbShowOpen As Boolean = True) As String
                         
    utGetPath = GetPath(psAppEXEName, psName, psMess, psFileMess, psDeFaultPath, plHwnd, psFilter, psSelFile, plFlags, pbCenterForm, pbShowOpen)
                         
End Function

Public Function utGetSavePath(psAppEXEName As String, psName As String, psMess As String, psFileMess As String, psDeFaultPath As String, plHwnd As Long, _
                         Optional psFilter As String = vbNullString, _
                         Optional psSelFile As String, _
                         Optional plFlags As Long, _
                         Optional pbCenterForm As Boolean = True, _
                         Optional pbShowOpen As Boolean = True) As String
    utGetSavePath = GetSavePath(psAppEXEName, psName, psMess, psFileMess, psDeFaultPath, plHwnd, psFilter, psSelFile, plFlags, pbCenterForm, pbShowOpen)

End Function

Public Function utGetTickCount() As String
                         
    utGetTickCount = CStr(timeGetTime)
                         
End Function

Public Function utFetchFLSFile(psRawPathFileName As String, psCopyToPath As String, _
                             psLifeSpanServerPath As String, Optional plLifeSpanDays = 1, _
                             Optional pbDelRawFile As Boolean) As String
                             
          utFetchFLSFile = FetchFLSFile(psRawPathFileName, psCopyToPath, psLifeSpanServerPath, plLifeSpanDays, pbDelRawFile)
End Function

Public Sub utDelFLSFiles(psLifeSpanServerPath As String)
    DelFLSFiles psLifeSpanServerPath
End Sub

Public Function utDeleteAllFLSFiles(psFLSDir As String, plFLSDays As Long) As Boolean
    utDeleteAllFLSFiles = DeleteAllFLSFiles(psFLSDir, plFLSDays)
End Function

Public Sub utSuffixLabels(plblArray As Object, Optional plLen As Long = 25)
    SuffixLabels plblArray, plLen
End Sub

Public Function utPrintListView(psAppEXEName As String, polvw As Object, psTitle As String, _
                              Optional pPrtOrientation As DDActiveReports.PrtOrientation = ddOPortrait, _
                              Optional pShowMode As VBRUN.FormShowConstants = vbModeless, _
                              Optional plPageBreakAfterXItems As Long = 0, _
                              Optional pbSelectedOnly As Boolean = False) As Boolean
    Dim sExt As String
    sExt = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "SHOWLIST_FORMAT", ".xls")
    ShowItemList polvw, , psTitle, , True, sExt
                              
End Function

Private Sub ShowItemList(poLstView As Object, _
                         Optional plMaxColLen As Long = 100, _
                         Optional psOutPutName As String = vbNullString, _
                         Optional psOutPutPath As String = vbNullString, _
                         Optional pbUseTempPrefix As Boolean = False, _
                         Optional psExt As String)
    On Error GoTo EH
    'Error
    Dim lRet As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    'File names
    Dim sTickCount As String
    Dim sFileName As String
    Dim sFullPathName As String
    Dim sTempDir As String
    Dim sExt As String
    Dim bValidExt As Boolean
    Dim bDelAppApthFile As Boolean
    'Objects
    Dim itmX As listItem
    Dim oLstView As ListView
    'Build Print Data
    Dim lColPos As Long
    Dim lFillLen As Long
    Dim aryColMaxLen() As Long
    Dim sHeader As String
    Dim sData As String
    Dim sTemp As String
     
    sTickCount = goUtil.utGetTickCount
    
    'Set the File Name. use .txt or .xls
    'Only Support .txt and .xls
    If psExt = vbNullString Then
        psExt = ".txt"
    Else
        sExt = psExt
    End If
    
    'Check for Valid Ext
    If StrComp(sExt, ".txt", vbTextCompare) = 0 Then
        bValidExt = True
    End If
    
    If StrComp(sExt, ".xls", vbTextCompare) = 0 Then
        bValidExt = True
    End If
    
    If Not bValidExt Then
        'Bail
        Exit Sub
    End If
    
    'Set the List View Object
    Set oLstView = poLstView
    
    If psOutPutName = vbNullString Then
        sFileName = "Temp_" & sTickCount & "_ItemList" & sExt
    Else
        If pbUseTempPrefix Then
            sFileName = "Temp_" & sTickCount & "_" & psOutPutName & sExt
        Else
            sFileName = psOutPutName & sExt
        End If
    End If
    
    'Set Output path
    If psOutPutPath = vbNullString Then
        'Try to use temp directory
        sTempDir = "C:\Temp\"
    Else
        sTempDir = psOutPutPath
    End If
    
    sFullPathName = sTempDir & sFileName
    
    If Not goUtil.utFileExists(sTempDir, True) Then
        bDelAppApthFile = True
        sTempDir = App.Path & "\"
    End If
    
    'Build the Data
    Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
    
    '1. Build the Header
    ReDim aryColMaxLen(1 To oLstView.ColumnHeaders.Count)
    For lColPos = 1 To oLstView.ColumnHeaders.Count
        'Only show Cols visibile on the List View
        'This will allow the user to change
        'what cols show on the out put !
        If oLstView.ColumnHeaders(lColPos).Width > 0 Then
            'Only need to get Max Len for txt format
            If StrComp(sExt, ".txt", vbTextCompare) = 0 Then
                aryColMaxLen(lColPos) = GetMaxLenthForCol(oLstView, lColPos)
            End If
            sTemp = oLstView.ColumnHeaders(lColPos).Text
            sTemp = "[" & sTemp & "]" 'wrap the col name
            If StrComp(sExt, ".txt", vbTextCompare) = 0 Then
                'If the data max len is less then Col len
                'then use the Col length
                If aryColMaxLen(lColPos) < Len(sTemp) Then
                    aryColMaxLen(lColPos) = Len(sTemp)
                End If
                'figure out the filler space needed
                'to keep the report format spaced correctly
                lFillLen = aryColMaxLen(lColPos)
                lFillLen = (lFillLen - Len(sTemp))
                If lFillLen > 0 Then
                    sTemp = sTemp & String(lFillLen, Chr(32))
                End If
            End If
            'Add to the header
            sHeader = sHeader & sTemp & vbTab
        End If
    Next
    If sHeader <> vbNullString Then
        'Add CRLF to separate the header
        sHeader = sHeader & vbCrLf
    End If
    
    'Set the Header to Data
    sData = sHeader
    
    '2. Now build the body
    For Each itmX In oLstView.ListItems
        For lColPos = 1 To oLstView.ColumnHeaders.Count
            If oLstView.ColumnHeaders(lColPos).Width > 0 Then
                'First Col item will be in Text
                If lColPos = 1 Then
                    sTemp = itmX.Text
                Else
                    'Subsequent Cols will be in ListSubItems
                    sTemp = itmX.ListSubItems(lColPos - 1).Text
                End If
                '***Note need to clear out certain chars
                'to maintian format integrity
                'Need to clear out any CRLF present in the data
                'Replace with 2 spaces
                sTemp = Replace(sTemp, vbCrLf, String(2, Chr(32)))
                'Can't have any extra tabs,
                'since this is a Tab delimited report!
                sTemp = Replace(sTemp, vbTab, " ")
                'Check for Flags
                If goUtil.GetFlagFromText(sTemp) Then
                    'Convert to "X"
                    sTemp = "X"
                End If
                
                '+ 3 accoutns for "..."
                If Len(sTemp) > (plMaxColLen + 3) Then
                    sTemp = left(sTemp, plMaxColLen) & "..."
                End If
                'Only need to get Max Len for txt format
                If StrComp(sExt, ".txt", vbTextCompare) = 0 Then
                    lFillLen = aryColMaxLen(lColPos)
                    lFillLen = lFillLen - Len(sTemp)
                    If lFillLen > 0 Then
                        sTemp = sTemp & String(lFillLen, Chr(32))
                    End If
                End If
                sData = sData & sTemp & vbTab
            End If
        Next
        sData = sData & vbCrLf
    Next
    
    
    'Save it to temp directory
    goUtil.utSaveFileData sFullPathName, sData
    
    If goUtil.utFileExists(sFullPathName) Then
        lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sFullPathName, vbNullString, App.Path, vbNormalFocus, False, False, True)
    End If
    
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    
    Set oLstView = Nothing
    Set itmX = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Private Sub ShowItemList"
End Sub

Private Function GetMaxLenthForCol(poLstView As Object, _
                                    lColPos As Long, _
                                    Optional plMaxColLen As Long = 100) As Long
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim itmX As listItem
    Dim oLstView As ListView
    Dim sTemp As String
    Dim lThisLen As Long
    Dim lLen As Long

    Set oLstView = poLstView
    
    For Each itmX In oLstView.ListItems
        If lColPos = 1 Then
            sTemp = itmX.Text
        Else
            sTemp = itmX.ListSubItems(lColPos - 1).Text
        End If
        lThisLen = Len(sTemp)
        If lThisLen > lLen Then
            lLen = lThisLen
        End If
    Next
    
    If lLen > plMaxColLen Then
        ' Add 3 to the max length to account for "..."
        lLen = plMaxColLen + 3
    End If
    
    GetMaxLenthForCol = lLen
    
    Set itmX = Nothing
    Set oLstView = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Private Function GetMaxLenthForCol"
End Function


Public Sub utHideAllForms(goForms As Object, Optional psSkipForms As String)
    On Error GoTo EH
    Dim MyForm As Form
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    For Each MyForm In goForms
        If InStr(1, psSkipForms, MyForm.Name, vbTextCompare) = 0 Then
            MyForm.Visible = False
        End If
    Next
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub utHideAllForms"
End Sub

Public Sub utShowAllForms(goForms As Object, oECTray As Object, psNavPos As String, Optional psSkipForms As String)
    On Error GoTo EH
    Dim MyForm As Form
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bCheckModalFlag As Boolean
    
    'First check for Modal forms.  If we find one need to Set focus
    'to the most current Modal form
    
    For Each MyForm In goForms
        On Error Resume Next
        'Add any more forms that need this modal flag  here
        If StrComp(MyForm.Name, "frmFixXactProjects", vbTextCompare) = 0 Then
            bCheckModalFlag = True
        End If
        If bCheckModalFlag Then
            If MyForm.ModalFlag Then
                If Err.Number = 0 Then
                    MyForm.SetFocus
                    If Err.Number = 0 Then
                        Exit Sub
                    End If
                Else
                    Err.Clear
                End If
            End If
        End If
    Next
    
    On Error GoTo EH
    
    For Each MyForm In goForms
        If InStr(1, psSkipForms, MyForm.Name, vbTextCompare) = 0 Then
            If InStr(1, MyForm.Name, "frmWallPaper", vbTextCompare) = 0 And InStr(1, MyForm.Name, "frmProgress", vbTextCompare) = 0 Then
                'If this is the frmClaimsList then check to see if there is
                'an open claim. if so do not show the claims List
                If StrComp(MyForm.Name, "frmClaimsList", vbTextCompare) = 0 Then
                    If Not MyForm.MyfrmClaim Is Nothing Then
                        GoTo NEXT_FORM
                    End If
                End If
                
                If MyForm.Visible = False Then
                    MyForm.Visible = True
                End If
                If MyForm.WindowState <> vbMaximized Then
                    MyForm.WindowState = vbNormal
                    'Check the Pos of the open forms to see if they
                    'are hidden behind the Nav Screen. If they are, Move the
                    'Left to make it fully visible
                    'Do not move the Claim Form
                    If StrComp(MyForm.Name, "frmClaim", vbTextCompare) = 0 Then
                        GoTo NEXT_FORM
                    ElseIf MyForm.Tag <> vbNullString Then
                        GoTo NEXT_FORM
                    End If
                    If psNavPos = "LEFT" Then
                        If MyForm.left < oECTray.Width Then
                            MyForm.left = oECTray.left + oECTray.Width
                        End If
                    Else
                        If MyForm.left + MyForm.Width > oECTray.left Then
                            MyForm.left = oECTray.left - MyForm.Width
                        End If
                    End If
                End If
                On Error Resume Next
                MyForm.Show
                If Err.Number > 0 Then
                    Err.Clear
                End If
                On Error GoTo EH
            'Check for Wall paper
            Else
                If InStr(1, MyForm.Name, "frmWallPaper", vbTextCompare) = 0 Then
                    oECTray.ShowWallPaper
                End If
            End If
        End If
NEXT_FORM:
    Next
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub utShowAllForms"
End Sub

Public Function utShowFormIgnoreModality(poForm As Object, Optional pbShow As Boolean = True) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not poForm Is Nothing Then
        If pbShow Then
            If poForm.left < 0 Then
                poForm.left = poForm.left + 20000
                poForm.Refresh
            End If
        Else
            If poForm.left > 0 Then
                poForm.left = poForm.left - 20000
                poForm.Refresh
            End If
        End If
    End If
    
    utShowFormIgnoreModality = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function utShowFormIgnoreModality"
End Function

Public Sub utShowFTP(psShellCommand As String)
    On Error GoTo EH
    Dim lHwnd As Long
    Dim dblRet As Double
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    lHwnd = utFindWindowPartial("Communications Status", FwpStartsWith, True, False)
    If lHwnd = 0 Then
        dblRet = Shell(psShellCommand, vbNormalFocus)
    End If
    SaveSetting "ECFTP", "MSG", "COMMAND", "SHOW_FTP"
    utLookForWindow "Communications Status", 2
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub utShowFTP"
End Sub

Public Sub utShutDownFTP()
    On Error GoTo EH
    Dim lHwnd As Long
    Dim dblRet As Double
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    lHwnd = utFindWindowPartial("Communications Status", FwpStartsWith, True, False)
    If lHwnd <> 0 Then
        SaveSetting "ECFTP", "MSG", "COMMAND", "SHUT_DOWN_FTP"
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "utShutDownFTP"
End Sub

Public Function utFTPConnected() As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim dblRet As Double
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bFTPConnected As Boolean
    
    lHwnd = utFindWindowPartial("Communications Status", FwpStartsWith, True, False)
    If lHwnd <> 0 Then
        bFTPConnected = CBool(GetSetting("ECFTP", "MSG", "CONNECTED", False))
    End If
    
    utFTPConnected = bFTPConnected
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "utFTPConnected"
End Function

Public Function utFTPExists() As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim dblRet As Double
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim bFTPExists As Boolean
    
    lHwnd = utFindWindowPartial("Communications Status", FwpStartsWith, True, False)
    If lHwnd <> 0 Then
        bFTPExists = True
    End If
    
    utFTPExists = bFTPExists
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "utFTPExists"
End Function

Public Function utValidate(Optional poForm As Object, Optional poControl As Object) As Boolean
    utValidate = Validate(poForm, poControl)
End Function

Public Sub utCleanFileFolderName(pvText As Variant, Optional pbFilePath As Boolean)
     CleanFileFolderName pvText, pbFilePath
End Sub

Public Function utCreateCat(psAppEXEName As String, psInstallDir As String, _
                            pcolCatPrefs As Collection, _
                            Optional pbHideError As Boolean) As Boolean
                                
    utCreateCat = CreateCat(psAppEXEName, psInstallDir, pcolCatPrefs, pbHideError)
End Function

Public Function utCreateCatPref(pcolCatPrefs As Collection) As String
    utCreateCatPref = CreateCatPref(pcolCatPrefs)
End Function

Public Sub utCompRepair(psAppEXEName As String, psSourcePath As String, Optional pbSkipBackup As Boolean)
    CompRepair psAppEXEName, psSourcePath, pbSkipBackup
End Sub

Public Sub utFillAddressFields(psAddress As String, _
                             psZip As String, _
                             psState As String, _
                             psCity As String, _
                             psStreet As String)
                             
    FillAddressFields psAddress, psZip, psState, psCity, psStreet
    
End Sub

Public Sub utUpdateAddress(psAddress As String, _
                         psZip As String, _
                         psState As String, _
                         psCity As String, _
                         psStreet As String)
    UpdateAddress psAddress, psZip, psState, psCity, psStreet
End Sub

Public Sub utBubbleSort(pvArray As Variant, Optional psEndString As String, Optional pbRebillPresent As Boolean)
    BubbleSort pvArray, psEndString, pbRebillPresent
End Sub

Public Function utValidDate(psDate As String) As String
    utValidDate = ValidDate(psDate)
End Function

Public Function utValidSSN(psSSN As String) As String
    utValidSSN = ValidSSN(psSSN)
End Function

Public Function utGetTaskbarHeight() As Long
    utGetTaskbarHeight = GetTaskbarHeight
End Function

Public Sub utSaveLic(plDays As Long)
    SaveLic plDays
End Sub

Public Function utWindowFound(psName As String, Optional plSleepSeconds As Long = 1) As Boolean
    utWindowFound = WindowFound(psName, plSleepSeconds)
End Function

Public Function utDeleteFile(psFilePath As String) As String
    utDeleteFile = DeleteFile(psFilePath)
End Function

Public Function utCopyFile(psSourceFilePath As String, psDestFilePath As String) As String
    utCopyFile = CopyFile(psSourceFilePath, psDestFilePath)
End Function

Public Function utDeleteDir(psDirPath As String) As String
    utDeleteDir = DeleteDir(psDirPath)
End Function

Public Function utMakeDir(psDirPath As String) As String
    utMakeDir = MakeDir(psDirPath)
End Function

Public Function utCopyDir(psSourceDirPath As String, psDestDirPath As String) As String
    utCopyDir = CopyDir(psSourceDirPath, psDestDirPath)
End Function

Public Function utGetSystemDir() As String
    utGetSystemDir = GetSystemDir()
End Function

Public Function utJoinToAddress(psStreet As String, psCity As String, psState As String, psZip As String) As String
    utJoinToAddress = JoinToAddress(psStreet, psCity, psState, psZip)
End Function

Public Function utGetProgForm(Optional pShowMode As FormShowConstants = vbModeless) As Object
    On Error GoTo EH
    Dim MyProgForm As frmProgress
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Set MyProgForm = New frmProgress
    Load MyProgForm
    MyProgForm.Show pShowMode
    
    Set utGetProgForm = MyProgForm
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function utGetProgForm"
End Function

Public Function utBadAutoNum(Optional pofID As Object, Optional psTable As String = "Assignments", _
                           Optional plMaxID As Long, Optional psFieldName = "ID") As Boolean
    'BGS 2.28.2002 151  "Duplicate IB"
    'EasyClaim cannot save the entry because it is a duplicate IB
    'Need to check if the autonumber was bad.
    'You will also get same error when the autonum creates an already existing ID.
    'If BadAutoNum returns true, that means the Autonum was Bad , and BadAutoNum
    'function has set the autonum to the next correct ID. Which means we can resume
    'adding the record.  This is a known bug with JET4.0 Engine, once the autonum field
    'gets corrupted theres no fixing it permanently. It may reset itself again.  So if
    'it does we will catch it here and fix it.
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;Q257408

    On Error GoTo EH
    Dim sSQL As String
    Dim lMaxID As Long
    Dim RS As dao.Recordset
    Dim oField As dao.Field
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Set oField = pofID
    
    'First need to get the Max ID
    sSQL = "SELECT MAX(A." & psFieldName & ") As MaxID "
    sSQL = sSQL & "FROM " & psTable & " A "
    
    Set RS = gCurDB.OpenRecordset(sSQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        lMaxID = IIf(IsNull(RS!MaxID), 0, RS!MaxID)
    End If
    Set RS = Nothing
    
    'If we passed in a Field Object update here
    If Not oField Is Nothing Then
        If oField.Value > lMaxID Then
            'If the autonum ID is greater than the MAX ID (The last ID )
            'Then The AutoNum is not Bad so set this function to false
            utBadAutoNum = False
        Else
            'if the autonum is < or = to the MAX ID then we have a bad Autonum.
            'So Change the ID to the next ID number
            oField.Value = lMaxID + 1
            utBadAutoNum = True
        End If
    Else
        plMaxID = lMaxID
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function utBadAutoNum"
End Function

Public Sub utFindListItem(poForm As Object, poListBox As Object, psItem As String, _
                         Optional pLastIndex As Long = 0, Optional pCompare As VbCompareMethod = vbTextCompare)
    On Error GoTo EH
    
    Dim lCount As Long
    Dim lCount2 As Long
    Dim bFound As Boolean
    Dim itmX As listItem
    Dim sFind As String
    Dim sTemp As String
    Dim oListBox As ListView
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If TypeOf poListBox Is ListView Then
        Set oListBox = poListBox
    End If
                
    'BGS 3.4.2001 This will Find and Highlight Desired Item
    'In the ListView Box
    'Go through entire list and search each Item and select if the string exists in
    'any of the columns
    If pLastIndex = 0 Then
        'if the index is o then unselect any already selected items
        If oListBox.ListItems.Count > 0 Then
            For lCount = 1 To oListBox.ListItems.Count
                Set itmX = oListBox.ListItems(lCount)
                itmX.Selected = False
            Next
        End If
        sFind = InputBox("Enter search text", "Find", psItem, poForm.left, poForm.top)
        psItem = sFind 'Set the return parameter here
    Else
        sFind = psItem
    End If
    
    With oListBox
        For lCount = IIf(pLastIndex = 0, 1, pLastIndex) To .ListItems.Count
            Set itmX = .ListItems(lCount)
            With itmX
                If Trim(sFind) <> vbNullString Then
                    sTemp = itmX.Text
                    If InStr(1, sTemp, sFind, pCompare) > 0 Then
                        bFound = True
                    End If
                    'If did not find it in the main text then
                    'look in the subitems...
                    If Not bFound Then
                        For lCount2 = 1 To itmX.ListSubItems.Count
                            sTemp = itmX.ListSubItems(lCount2).Text
                            If InStr(1, sTemp, sFind, pCompare) > 0 Then
                                bFound = True
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                If bFound Then
                    itmX.Selected = True
                    itmX.EnsureVisible
                    pLastIndex = lCount + 1
                    Exit For
                End If
            End With
        Next
    End With
    
    If Not bFound And sFind <> vbNullString Then
        MsgBox """" & sFind & """ Not found.", vbInformation + vbOKOnly, "Find"
    Else
        If oListBox.Visible Then
            If oListBox.Enabled Then
                oListBox.SetFocus
            End If
        End If
    End If
    
    
    Set itmX = Nothing
    Set oListBox = Nothing
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub utFindListItem"
End Sub

Public Function utGetIBPREFIX() As String
    On Error GoTo EH
    Dim sIBPrefix As String
    Dim sFName As String
    Dim sLName As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sIBPrefix = Trim(GetSetting("ECS", "WEB_SECURITY", "IB_PREFIX", vbNullString))
    
    If sIBPrefix = vbNullString Then
        sLName = UCase(Trim(GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_LAST_NAME", "X")))
        sFName = UCase(Trim(GetSetting(App.EXEName, "GENERAL", "ADJUSTOR_FIRST_NAME", "X")))
        sIBPrefix = left(sFName, 1) & left(sLName, 1)
    End If
    
    utGetIBPREFIX = sIBPrefix
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function utGetIBPREFIX"
End Function

Public Function utNumInTextSortFormat(psText As String) As String
    'This function will search Text string and format any numerics
    'with 15 digit 0 place holder so numbers can be sorted within a string.
    On Error GoTo EH
    Dim sText As String
    Dim sTemp As String
    Dim sTemp2 As String
    Dim lCount As Long
    Dim lCount2 As Long
    Dim varyText As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sText = psText
    
    varyText = Split(sText, Chr(32))
    
    If IsArray(varyText) Then
        For lCount = LBound(varyText, 1) To UBound(varyText, 1)
            sText = varyText(lCount)
            If IsNumeric(sText) Then
                sText = Format(sText, "000000000000000.0000")
                varyText(lCount) = sText
            Else
                lCount2 = 1
                sTemp = vbNullString
                sTemp2 = vbNullString
                'If there is a Numeric + Alpha I.E. 1ST street or 2ND etc
                'Need to find the pos where the number stops and format it as well!
                'Need to format a number in the middle of text.
                Do Until lCount2 > Len(sText)
                    If Not IsNumeric(Mid(sText, lCount2, 1)) Then
                        sTemp = sTemp & Mid(sText, lCount2, 1)
                        lCount2 = lCount2 + 1
                    Else
                        sTemp2 = vbNullString
                        Do Until Not IsNumeric(Mid(sText, lCount2, 1))
                            sTemp2 = sTemp2 & Mid(sText, lCount2, 1)
                            lCount2 = lCount2 + 1
                        Loop
                        sTemp = sTemp & Format(sTemp2, "000000000000000.0000")
                    End If
                Loop
                varyText(lCount) = sTemp
            End If
        Next
        utNumInTextSortFormat = Join(varyText, Chr(32))
    Else
        utNumInTextSortFormat = vbNullString
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
   Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function utNumInTextSortFormat"
End Function

Public Sub utMoveListItem(poListView As Object, pMoveItem As MoveListItem)
    'USe this Function to Move a list view's Item data from one
    'list item to the next Up or down in the list.  Includes Icons and report Icons
    Dim strSwap As String
    Dim lIconSwap As Long
    Dim lSmallIconSwap As Long
    Dim sKeySwap As String
    Dim sKeyItem As String
    Dim oListView As MSComctlLib.ListView
    Dim listItem As MSComctlLib.listItem
    Dim swapItem As MSComctlLib.listItem
    Dim lngCounter As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set oListView = poListView

    If oListView.SelectedItem Is Nothing Then
        Exit Sub
    End If
    If Not oListView.SelectedItem.Selected Then
        Exit Sub
    End If
    If pMoveItem = MoveUp Then
        If oListView.SelectedItem.Index = 1 Then
            Exit Sub
        End If
    End If
    If pMoveItem = MoveDown Then
        If oListView.SelectedItem.Index = oListView.ListItems.Count Then
            Exit Sub
        End If
    End If

    oListView.Visible = False
    
    Set listItem = oListView.SelectedItem
    Set swapItem = oListView.ListItems(oListView.SelectedItem.Index + pMoveItem)
    
    strSwap = listItem.Text
    listItem.Text = swapItem.Text
    swapItem.Text = strSwap
    sKeySwap = swapItem.Key
    sKeyItem = listItem.Key
    listItem.Key = vbNullString
    swapItem.Key = vbNullString
    listItem.Key = sKeySwap
    swapItem.Key = sKeyItem

    lIconSwap = listItem.Icon
    lSmallIconSwap = listItem.SmallIcon
    listItem.Icon = swapItem.Icon
    listItem.SmallIcon = swapItem.SmallIcon
    swapItem.Icon = lIconSwap
    swapItem.SmallIcon = lSmallIconSwap


    For lngCounter = 1 To oListView.ColumnHeaders.Count - 1
        strSwap = listItem.SubItems(lngCounter)
        listItem.SubItems(lngCounter) = swapItem.SubItems(lngCounter)
        swapItem.SubItems(lngCounter) = strSwap
        lIconSwap = listItem.ListSubItems(lngCounter).ReportIcon
        listItem.ListSubItems(lngCounter).ReportIcon = swapItem.ListSubItems(lngCounter).ReportIcon
        swapItem.ListSubItems(lngCounter).ReportIcon = lIconSwap
    Next

    listItem.Selected = False
    swapItem.Selected = True
    swapItem.EnsureVisible
    
    oListView.Visible = True
    
    'CleanUp
    Set oListView = Nothing
    Set listItem = Nothing
    Set swapItem = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub utMoveListItem"
End Sub

Public Function utScrubAlphaNumeric(psScrubMe As String, Optional pbAllowSpaces As Boolean = False) As String
    On Error GoTo EH
    Dim sSCrub As String
    Dim sClean As String
    Dim lPos As Long
    Dim sLetter As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sSCrub = psScrubMe
    sClean = psScrubMe
    
    For lPos = 1 To Len(sSCrub)
        sLetter = Mid(sSCrub, lPos, 1)
        If Not utIsAlphaNumeric(sLetter) Then
            sClean = Replace(sClean, sLetter, vbNullString, , , vbBinaryCompare)
        End If
    Next
    
    utScrubAlphaNumeric = sClean
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Private Function utScrubAlphaNumeric"
End Function

Public Function utIsAlphaNumeric(psLetter As String, Optional pbAllowSpaces As Boolean = False) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    If pbAllowSpaces Then
        utIsAlphaNumeric = psLetter Like "[A-Za-z0-9_- ]"
    Else
        utIsAlphaNumeric = psLetter Like "[A-Za-z0-9]"
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Private Function utIsAlphaNumeric"
End Function

Public Function utActiveFile(psActiveFilePath As String, plCurrentTry As Long) As Boolean
    '3.3.2004 Active file will return true IF
    'The file is is actively being written to
    On Error GoTo EH
    Dim iFFile As Integer
    Dim sActiveFile As String
    
    sActiveFile = psActiveFilePath
    plCurrentTry = plCurrentTry + 1
    iFFile = FreeFile
    Open sActiveFile For Binary Access Read Lock Read As #iFFile
    Close #iFFile
    utActiveFile = False
    Exit Function
EH:
    Close #iFFile
    utActiveFile = True
End Function

Public Sub DecodeStr(psEncryptedText As String, psPlainText As String)
    Dim oKeyBoard As V2ECKeyBoard.clsKeyBoard
    Set oKeyBoard = New V2ECKeyBoard.clsKeyBoard
    psPlainText = oKeyBoard.Decode(psEncryptedText)
    Set oKeyBoard = Nothing
End Sub
Public Sub EncodeStr(psPlainText As String, psEncryptText As String)
    Dim oKeyBoard As V2ECKeyBoard.clsKeyBoard
    Set oKeyBoard = New V2ECKeyBoard.clsKeyBoard
    psEncryptText = oKeyBoard.Encode(psPlainText)
    Set oKeyBoard = Nothing
End Sub

Public Function Decode(psEncryptedText As String, Optional pbSeedCheck As Boolean = False) As String
    Decode = mclsKeyBoard.Decode(psEncryptedText, pbSeedCheck)
End Function

Public Function Encode(psPlainText As String) As String
    Encode = mclsKeyBoard.Encode(psPlainText)
End Function

Public Function SendGetCATFromDisk(psAppEXEName As String, pbSave As Boolean, Optional pbBackup As Boolean, _
                                   Optional psCar As String, _
                                   Optional psCAT As String) As Boolean
    On Error GoTo EH
    Dim sSSN As String
    Dim sToDiskMess As String
    Dim sFromDiskMess As String
    Dim sMyFilter As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If pbBackup Then
        'BGS 12.28.2001 check to see if SSN has been filled out.
        'if not can not continue
        sSSN = GetECSCryptSetting("ECS", "WEB_SECURITY", "SSN")
        If sSSN = vbNullString Then
            MsgBox "All fields in File|Preferences must be filled out!", vbInformation + vbOKOnly, "CAT BACKUP"
            GoTo CLEAN_ME_UP
        End If
        sToDiskMess = "Backup CAT To Disk"
        sFromDiskMess = "Get Backup CAT From Disk"
        sMyFilter = "CAT Backup File" & " (*." & "bkp" & ")" & SD & "*." & "bkp" & SD
    Else
        sToDiskMess = "Export CAT To Disk"
        sFromDiskMess = "Import CAT From Disk"
        sMyFilter = "CAT Export File" & " (*." & "exp" & ")" & SD & "*." & "exp" & SD
    End If

    'BGS if we not saving then getting
    If Not pbSave Then
        SendGetCATFromDisk = GetFromDisk(psAppEXEName, sSSN, pbBackup, sFromDiskMess, sMyFilter)
    Else ' we are Saving :)
        SendGetCATFromDisk = SendToDisk(psAppEXEName, psCar, psCAT, sSSN, pbBackup, sToDiskMess, sMyFilter)
    End If
    
CLEAN_ME_UP:
    Screen.MousePointer = vbDefault
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function SendGetCATFromDisk"
End Function

Private Function GetFromDisk(psAppEXEName As String, psSSN As String, pbBackup As Boolean, _
                             psFromDiskMess As String, psMyFilter As String) As Boolean
    On Error GoTo EH
    Dim sCarrierName As String
    Dim sCatName As String
    Dim sNewCatName As String
    Dim sExtractName As String
    Dim sCurDirectory As String
    Dim sTemp As String
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sBackupSSN As String
    Dim sBackUpTime As String
    Dim varyBackUp As Variant
    Dim varyName As Variant
    Dim sMess As String
    Dim sPath As String
    Dim sSelFile As String
    Dim sExt As String
    Dim bPreservePaths As Boolean
    Dim sPreserveDir As String
    Dim sMainPresDirFolder As String
    Dim sExported As String
    Dim sFilter As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    bPreservePaths = True
    If pbBackup Then
        sExt = ".zkp"
    Else
        sExt = ".zxp"
    End If
    sFilter = psMyFilter
    If Not pbBackup Then
        'Allow for Previous Version Exports... Previous Verison Backups Not allowed!
        sFilter = sFilter & "VS 1 Export File" & " (*." & "dat" & ")" & SD & "*." & "dat" & SD
    End If
    
    sPath = utGetPath(psAppEXEName, psFromDiskMess, psFromDiskMess, vbNullString, sPath, gfrmECTray.hWnd, sFilter, sSelFile, , False)
    If FileExists(sSelFile) Then
        sExtractName = GetFileData(sSelFile)
        'BGS 12.28.2001 IF we are just importing Check to see if this is really a BackUp floppy
        If Not pbBackup Then
            If InStr(1, sExtractName, vbCrLf, vbBinaryCompare) > 0 Then
                MsgBox "Import not allowed from CAT Backup", vbInformation + vbOKOnly, "Can't Import CAT"
                GoTo CLEAN_ME_UP
            End If
            varyName = Split(sExtractName, "|")
            If UBound(varyName, 1) > 0 Then
                sExtractName = varyName(0)
                sPreserveDir = Decode(CStr(varyName(1)))
            Else
                'Dealing with previous version Export File
                sExt = ".zip"
                bPreservePaths = False
            End If
        Else
            If InStr(1, sExtractName, vbCrLf, vbBinaryCompare) = 0 Then
                MsgBox "BackUp not allowed from CAT Export", vbInformation + vbOKOnly, "Can't Get CAT BackUp"
                GoTo CLEAN_ME_UP
            End If
            varyBackUp = Split(sExtractName, vbCrLf)
            sExtractName = varyBackUp(0)
            varyName = Split(sExtractName, "|")
            If UBound(varyName, 1) > 0 Then
                sExtractName = varyName(0)
                sPreserveDir = Decode(CStr(varyName(1)))
            Else
                'Dealing with previous version Export File
                MsgBox "Previous version Backups not allowed!", vbExclamation + vbOKOnly, "Can't Get CAT BackUP"
                GoTo CLEAN_ME_UP
            End If
            sBackUpTime = varyBackUp(1)
            sBackupSSN = Decode(CStr(varyBackUp(2)))

            'BGS We need to check to see if the psSSN matches
            'Can not let anyone other than the Adjuster who made the backup
            'use this backup.
            If psSSN <> sBackupSSN Then
                MsgBox "Fields in File|Preferences do not match this CAT Backup", vbInformation + vbOKOnly, "Can't Get CAT BackUP"
                GoTo CLEAN_ME_UP
            End If
            'Ask if they want to proceed with the backup
            If MsgBox("CAT BackUp " & sExtractName & " (" & sBackUpTime & ")" & vbCrLf & _
                      "This BackUp will over-write the previous CAT if it still exists." & vbCrLf & vbCrLf & _
                      "Do you wish to proceed ? ", vbYesNo, "CAT BACKUP") = vbNo Then
                GoTo CLEAN_ME_UP
            End If
        End If

        'Issue 236 9.5.2002 Export to Floppy or Backup to Floppy
        'If the name of your cat has the name of the carrier in it,
        'IE you are working on Farmers Cat and the name of your cat is
        'FarmerTestCat or you are working on a SafeCo cat and the name of
        'your cat is SafecoTest, then you will run into problems when
        'importing the cat back into EasyClaim. The import process will
        'take the "Safeco" text out of the Cat name or "Farmer" text out
        'of the FarmersCat. This will cause the Unzip process to fail when
        'it can't find the correct cat name.
        sCarrierName = left(sExtractName, InStr(1, sExtractName, "\") - 1)
        'Issue 236 Need to replace only once and include the "\"
        sExtractName = Replace(sExtractName, sCarrierName & "\", vbNullString, , 1)
        sCatName = sExtractName
        
        
        'If geting a Backup ,First need to check if this Cat was previously Exported
        If pbBackup Then
            sExported = GetSetting(psAppEXEName, "CAT_STRUCTURE\" & sCarrierName & "\" & sCatName, sCatName, vbNullString)
            If sExported <> vbNullString Then
                sExported = Decode(sExported)
                If StrComp(sExported, "EXPORTED", vbTextCompare) = 0 Then
                    MsgBox "You may not Restore a CAT that was previously Exported!", vbExclamation + vbOKOnly, "Can't Get CAT BackUP"
                    GoTo CLEAN_ME_UP
                End If
            End If
        End If
        
        
        'BGS 10.23.2001 if the cat does not exists then make the cat dir
        sCurDirectory = gsInstallDir & "\Cats\" & sCarrierName & "\" & sCatName
        
        'Need to be sure Nobody Deleted the Cats Directory
        If Not FileExists(gsInstallDir & "\Cats", True) Then
            MakeDir gsInstallDir & "\Cats"
        End If
        If Not FileExists(gsInstallDir & "\Cats\" & sCarrierName, True) Then
            MakeDir gsInstallDir & "\Cats\" & sCarrierName
        End If
        If Not FileExists(gsInstallDir & "\Cats\01_Photos", True) Then
            MakeDir gsInstallDir & "\Cats\01_Photos"
        End If
        If Not FileExists(gsInstallDir & "\Cats\01_Photos\" & sCarrierName, True) Then
            MakeDir gsInstallDir & "\Cats\01_Photos\" & sCarrierName
        End If
        If Not FileExists(gsInstallDir & "\Cats\02_Attachments", True) Then
            MakeDir gsInstallDir & "\Cats\02_Attachments"
        End If
        If Not FileExists(gsInstallDir & "\Cats\02_Attachments\" & sCarrierName, True) Then
            MakeDir gsInstallDir & "\Cats\02_Attachments\" & sCarrierName
        End If
        
        If Not FileExists(sCurDirectory, True) Then
            sMess = MakeDir(sCurDirectory)
            If sMess <> vbNullString Then
                Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
            End If
        Else
            If pbBackup Then
                GoTo DEL_CAT
            End If
             sTemp = sCatName
CAT_EXISTS:     'If the Cat does exists ask if they want to write over it or
            'rename it to something else

            If MsgBox(sTemp & " already exists ! " & vbCrLf & vbCrLf & _
            "Press ""Yes"" to write over existing claim info." & vbCrLf & vbCrLf & _
            "Press ""No"" to Change the name of " & sTemp & " to something else.", vbInformation + vbYesNo, psFromDiskMess) = vbNo Then
                sNewCatName = InputBox("Enter a new name for this Cat", "CAT NAME", "New" & sTemp)
                'If they rename it to something that already exists then
                'ask them again.
                If sNewCatName = vbNullString Then
                    'Bail becuase they hit the cancel button
                    GetFromDisk = False
                    GoTo CLEAN_ME_UP
                End If
                
                sCurDirectory = gsInstallDir & "\Cats\" & sCarrierName & "\" & sNewCatName
                
                If FileExists(sCurDirectory, True) Then
                    sTemp = sNewCatName
                    GoTo CAT_EXISTS
                Else
                    sMess = MakeDir(sCurDirectory)
                    If sMess <> vbNullString Then
                        MsgBox "Problems Replacing " & sNewCatName & vbCrLf & vbCrLf & "Error: " & sMess, vbCritical + vbOKOnly
                        GoTo CLEAN_ME_UP
                    End If
                End If
            Else
                'Need to delete the current cat first before attempting to
                'unzip files into it
DEL_CAT:
                sMess = DeleteDir(gsInstallDir & "\Cats\" & sCarrierName & "\" & sCatName)
                If sMess <> vbNullString Then
                    MsgBox "Problems Replacing " & sCatName & vbCrLf & vbCrLf & "Error: " & sMess, vbCritical + vbOKOnly
                    GoTo CLEAN_ME_UP
                End If
            End If
        End If
        If sNewCatName = vbNullString Then
            sNewCatName = sCatName
        End If

        Set oXZip = New clsXZip
        
        sCurDirectory = gsInstallDir & "\Cats\" & sCarrierName & "\" & sNewCatName
        
        If oXZip.UNZipFiles(sCurDirectory, sPath & "\" & sCatName & sExt, bPreservePaths) Then
            
            If sPreserveDir <> vbNullString Then
                'Now Copy over the Files, Photos and Attachments if they exist
                'Cat Files
                If FileExists(sCurDirectory & "\" & sPreserveDir & "\Files", True) Then
                    sMess = CopyDir(sCurDirectory & "\" & sPreserveDir & "\Files", gsInstallDir & "\Cats\" & sCarrierName & "\" & sNewCatName)
                End If
                'Cat Photos
                If FileExists(sCurDirectory & "\" & sPreserveDir & "\Photos", True) Then
                    sMess = sMess & CopyDir(sCurDirectory & "\" & sPreserveDir & "\Photos", gsInstallDir & "\Cats\01_Photos\" & sCarrierName & "\" & sNewCatName)
                End If
                'Cat Attachments
                If FileExists(sCurDirectory & "\" & sPreserveDir & "\Attachments", True) Then
                    sMess = sMess & CopyDir(sCurDirectory & "\" & sPreserveDir & "\Attachments", gsInstallDir & "\Cats\02_Attachments\" & sCarrierName & "\" & sNewCatName)
                End If
                If sMess <> vbNullString Then
                    Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
                End If
                'Now Delete the Main Preserve Dir Folder
                sMainPresDirFolder = left(sPreserveDir, InStr(1, sPreserveDir, "\", vbBinaryCompare) - 1)
                sMainPresDirFolder = sCurDirectory & "\" & sMainPresDirFolder
                DeleteDir sMainPresDirFolder
                
            End If
            'Need to Add this Cat to Registry for CAT_STRUCTURE
            SaveSetting psAppEXEName, "CAT_STRUCTURE\" & sCarrierName, sCarrierName, Encode(sCarrierName)
            SaveSetting psAppEXEName, "CAT_STRUCTURE\" & sCarrierName & "\" & sNewCatName, sNewCatName, Encode(sNewCatName)
            
            
            'BGS 11.20.2001 need to remove the Cat from the Floppy since it
            'should only exist on one easy claim at any given time.
            'BGS 12.31.2001 If it is a Backup we can leave it on the floppy
            If Not pbBackup Then
                If FileExists(sPath & "\" & sCatName & sExt) Then
                    sMess = DeleteFile(sPath & "\" & sCatName & sExt)
                End If
                If FileExists(sSelFile) Then
                    sMess = sMess & DeleteFile(sSelFile)
                End If
                If FileExists(sPath & "\" & sSelFile) Then
                    sMess = sMess & DeleteFile(sPath & "\" & sSelFile)
                End If
                If sMess <> vbNullString Then
                    Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
                End If
            End If
            
            GetFromDisk = True
        End If

        Set oXZip = Nothing
    End If
CLEAN_ME_UP:
    Set oXZip = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oXZip = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Private Function GetFromDisk"
End Function

Private Function SendToDisk(psAppEXEName As String, psCar As String, psCAT As String, psSSN As String, pbBackup As Boolean, _
                             psToDiskMess As String, psMyFilter As String) As Boolean
    On Error GoTo EH
    Dim sCarrierName As String
    Dim sCatName As String
    Dim sZIPName As String
    Dim sCurDirectory As String
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sMess As String
    Dim sPath As String
    Dim sSelFile As String
    Dim sPhotoDir As String
    Dim sAttachDir As String
    Dim sZipDir As String
    Dim sTemp As String
    Dim sPreserveDir As String
    Dim sZip As String
    Dim varyName As Variant
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sCurDirectory = Mid(gCurDB.Name, 1, InStrRev(gCurDB.Name, "\") - 1)
    sPhotoDir = gsInstallDir & "\Cats\01_Photos\" & psCar & "\" & psCAT
    sAttachDir = gsInstallDir & "\Cats\02_Attachments\" & psCar & "\" & psCAT
    sZip = gsInstallDir & "\Cats\03_Zip\"
    sPreserveDir = sZip & psCar & "\" & psCAT
    
    If pbBackup Then
        sSelFile = psCAT & ".bkp"
    Else
        sSelFile = psCAT & ".exp"
    End If
    
    'Only need the very first directory from the install dir.
    'So take out the C:\ or D:\, whatever the install Root Drive is out!
    sPreserveDir = Mid(sPreserveDir, InStr(1, sPreserveDir, "\", vbBinaryCompare) + 1)
    If sPreserveDir <> vbNullString Then
        Do Until left(sPreserveDir, 1) <> "\"
            sPreserveDir = Mid(sPreserveDir, 2)
        Loop
    End If
    
    sCarrierName = psCar

    sPath = utGetPath(psAppEXEName, psToDiskMess, psToDiskMess, sSelFile, sPath, gfrmECTray.hWnd, psMyFilter, , , False, False)
    If Not FileExists(sPath, True) Or sPath = vbNullString Then
        GoTo CLEAN_ME_UP
    End If
    Screen.MousePointer = vbHourglass

    'BGS Get the Carrier and Cat name into the CatName
    sCatName = sCarrierName & "\" & psCAT

    'Need to Close the current DB
    If goCurCarList Is Nothing Then
        goCurCarList.CLEANUP
        Set goCurCarList = Nothing
    End If
    'Close Currnet DB
    CloseCurDB
    gsCurCatDir = vbNullString
    gfrmECTray.cmdConnect.Enabled = False
    gfrmECTray.cmdPrint.Enabled = False
    
    'BGS check to see if we have any existing and give chance to not
    'overwrite them
    If pbBackup Then
        If FileExists(sPath & "\" & psCAT & ".bkp") Then
            sTemp = GetFileData(sPath & "\" & psCAT & ".bkp")
            If sTemp = vbNullString Then
                DeleteFile sPath & "\" & psCAT & ".bkp"
                DeleteFile sPath & "\" & psCAT & ".zkp"
            Else
                sTemp = left(sTemp, InStrRev(sTemp, vbCrLf) - 1)
                varyName = Split(sTemp, "|")
                If UBound(varyName, 1) > 0 Then
                    sTemp = varyName(1)
                    sTemp = Mid(sTemp, InStr(1, sTemp, vbCrLf))
                    sTemp = varyName(0) & sTemp
                End If
                If MsgBox(sTemp & vbCrLf & " already exists." & vbCrLf & vbCrLf & "Do you want to over-write the CAT?", vbQuestion + vbYesNo, psToDiskMess) = vbNo Then
                    GoTo CLEAN_ME_UP
                End If
            End If
        End If
    Else
        If FileExists(sPath & "\" & psCAT & ".exp") Then
            sTemp = GetFileData(sPath & "\" & psCAT & ".exp")
            If sTemp = vbNullString Then
                DeleteFile sPath & "\" & psCAT & ".exp"
                DeleteFile sPath & "\" & psCAT & ".zxp"
            Else
                varyName = Split(sTemp, "|")
                If UBound(varyName, 1) > 0 Then
                    sTemp = varyName(0)
                End If
                If MsgBox(sTemp & vbCrLf & " already exists." & vbCrLf & vbCrLf & "Do you want to over-write the CAT?", vbQuestion + vbYesNo, psToDiskMess) = vbNo Then
                    GoTo CLEAN_ME_UP
                End If
            End If
        End If
    End If
    

    'bgs 10.23.2001 be sure the read only is off
    
    'Copy over The Cat Dir and The Cat Photos dir and Cat Attachments Dir
    'Cat Dir
    sZipDir = sZip
    If Not FileExists(sZipDir, True) Then
        MakeDir sZipDir
    End If
    
    sZipDir = sZipDir & psCar & "\"
    If Not FileExists(sZipDir, True) Then
        MakeDir sZipDir
    End If
    
    sZipDir = sZipDir & psCAT
    If Not FileExists(sZipDir, True) Then
        MakeDir sZipDir
    End If
    
    'Cat Files Dir
    sMess = CopyDir(sCurDirectory, sZipDir & "\Files")
    If sMess <> vbNullString Then
        Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
    End If
    'Cat Photos Dir
    If FileExists(sPhotoDir, True) Then
        sMess = CopyDir(sPhotoDir, sZipDir & "\Photos")
        If sMess <> vbNullString Then
            Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
        End If
    End If
    'Cat Attachments Dir
    If FileExists(sAttachDir, True) Then
        sMess = CopyDir(sAttachDir, sZipDir & "\Attachments")
        If sMess <> vbNullString Then
            Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
        End If
    End If
    
    'BGS Get the zip name
    If pbBackup Then
        sZIPName = psCAT & ".zkp"
    Else
        sZIPName = psCAT & ".zxp"
    End If

    'BGS 10.23.2001 be sure that we get rid of any existing files
    If FileExists(sPath & sZIPName) Then
        sMess = DeleteFile(sPath & sZIPName)
        If sMess <> vbNullString Then
            Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
        End If
    End If
    If pbBackup Then
        If FileExists(sPath & psCAT & ".bkp") Then
            DeleteFile sPath & psCAT & ".bkp"
        End If
    Else
        If FileExists(sPath & psCAT & ".exp") Then
            DeleteFile sPath & psCAT & ".exp"
        End If
    End If


    'BGS the SaveZipFiles will zip up all the files we copied over into the
    'zip diretory in the the Catname.zip.  And it will also delete the files
    'we just zipped up leaving just the Catname.zip

    Set oXZip = New clsXZip

    If oXZip.SaveZIPFiles(sZipDir, sZIPName, "*\*", DB_PASSWORD("1"), sPath) Then
'        sMess = CopyFile(sZipDir & "\" & sZIPName, sPath & "\" & sZIPName)
'        'Check for Disk Errors...
'        If sMess <> vbNullString Then
'            If CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)) = 70 Then
'                MsgBox "Can't write to disk!  Permission denied.", vbExclamation, "ERROR"
'                DeleteDir (sZip & "*")
'            ElseIf CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)) = -2147024784 Then
'                MsgBox "Can't write to disk!  Not enough space on disk.", vbExclamation, "ERROR"
'                DeleteDir (sZip & "*")
'            Else
'                Err.Raise CLng(left(sMess, InStr(1, sMess, vbCrLf) - 1)), , sMess
'            End If
'            GoTo CLEAN_ME_UP
'        End If
'        'BGS now that we have copied over the Ziped cat to the Floppy we can delete
'        'the zipped cat from the Catname\Zip\ directory.
'        sMess = DeleteDir(sZip & "*")
        Set oXZip = Nothing
    Else
        SendToDisk = False
        GoTo CLEAN_ME_UP
    End If

    'BGS now create the floppy catname file with the name of the cat
    If pbBackup Then
        SaveFileData sPath & "\" & psCAT & ".bkp", sCatName & "|" & Encode(sPreserveDir) & vbCrLf & Now & vbCrLf & Encode(psSSN)
    Else
        SaveFileData sPath & "\" & psCAT & ".exp", sCatName & "|" & Encode(sPreserveDir)
        'BGS 11.20.2001 need to remove the Cat from this easyclaim since it was exported
        sMess = DeleteDir(gsInstallDir & "\Cats\" & psCar & "\" & psCAT)
        If sMess = vbNullString Then
            'This is very important! Need to update the registry that this cat
            'Was Exported.  This will minimize the possibility of the adjuster
            'making a Backup then exporting then Getting the Backup.
            SaveSetting psAppEXEName, "CAT_STRUCTURE\" & psCar & "\" & psCAT, psCAT, Encode("EXPORTED")
            DeleteDir sPhotoDir
            DeleteDir sAttachDir
        Else
            Exit Function
        End If
    End If
    
    SendToDisk = True
CLEAN_ME_UP:
    Set oXZip = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set oXZip = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Private Function SendToDisk"
End Function

Public Function CloseCurDB() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    CloseCurDB = True
    If Not gCurDB Is Nothing Then
        On Error Resume Next
        gCurDB.Close
        On Error GoTo EH
        Set gCurDB = Nothing
    End If
    If Not gWS Is Nothing Then
        On Error Resume Next
        gWS.Close
        On Error GoTo EH
        Set gWS = Nothing
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CloseCurDB = False
    Set gCurDB = Nothing
    Set gWS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function CloseCurDB"
End Function

Public Function CloseMainDB() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    CloseMainDB = True
    If Not gMainDB Is Nothing Then
        Set gMainDB = Nothing
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CloseMainDB = False
    Set gMainDB = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function CloseMainDB"
End Function

Public Function SetCurDB(psAppEXEName As String, psCompanyName As String, psCarName As String, psDBPath As String, _
                         Optional psDir As String = "App.Path", _
                         Optional pbCheckVersion As Boolean = True, _
                         Optional pbCompRepair As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    SetCurDB = True
    'BGS 1.10.2001 Need to be sure the default directory is
    'Ap.path or will get strange errors when creating WorkSpace
    'especially when the cur dir is the floppy a drive.
    If psDir = "App.Path" Then
        psDir = gsInstallDir
    End If
    ChDir psDir
    If Not gCurDB Is Nothing Then
        'Check to see if we already have theSame DB open
        If Not pbCompRepair Then
            If StrComp(gCurDB.Name, psDBPath, vbTextCompare) = 0 Then
                Exit Function
            End If
        End If
        Set gCurDB = Nothing
    End If
    
    If gWS Is Nothing Then
        utCreateWorkSpace
    End If
    
    utOpenDatabase mCurDB, psDBPath, False, False
    
    If pbCheckVersion Then
        CheckVSCurDB psAppEXEName, psCompanyName, psCarName, psDBPath, , , pbCompRepair
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    SetCurDB = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function SetCurDB"
End Function

Public Function SetMainDB(psAppEXEName As String, psDBPath As String, _
                         Optional psDir As String = "App.Path", _
                         Optional pbCheckVersion As Boolean = True, _
                         Optional pbCompRepair As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    SetMainDB = True
    'BGS 1.10.2001 Need to be sure the default directory is
    'Ap.path or will get strange errors when creating WorkSpace
    'especially when the cur dir is the floppy a drive.
    If psDir = "App.Path" Then
        psDir = gsInstallDir
    End If
    ChDir psDir
    If Not gMainDB Is Nothing Then
        'Check to see if we already have theSame DB open
        If Not pbCompRepair Then
            If StrComp(gMainDB.Name, psDBPath, vbTextCompare) = 0 Then
                Exit Function
            End If
        End If
        Set gMainDB = Nothing
    End If
    
    If gWS Is Nothing Then
        utCreateWorkSpace
    End If
    
    utOpenDatabase mMainDB, psDBPath, False, False
    
    If pbCheckVersion Then
        CheckVSMainDB psAppEXEName, psDBPath, , pbCompRepair
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    SetMainDB = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function SetMainDB"
End Function

Public Sub CheckVSCurDB(psAppEXEName As String, psCompanyName, psCarName As String, psCurDBPath As String, _
                        Optional pbDoPicDB As Boolean, _
                        Optional pbForceUpdate As Boolean, _
                        Optional pbCompRepair As Boolean)
    On Error GoTo EH
    Dim sCompanyName As String
    Dim sCarrierName As String
    Dim sCarrierVSNum As Integer
    Dim sTemplateDBPath As String
    Dim sTemplateReportsDir As String
    Dim sCatReportsDir As String
    Dim TemplateDB As Database
    Dim RSCur As Recordset
    Dim RSTemplate As Recordset
    Dim iPos As Integer
    Dim sSQL As String
    Dim bCheckVersionTable As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sCompanyName = psCompanyName
    sCarrierName = psCarName

    sTemplateReportsDir = gsInstallDir & "\Templates\" & sCompanyName & "\Carriers\" & sCarrierName & "\"
    sTemplateDBPath = sTemplateReportsDir & sCarrierName & ".mdb"
    sCarrierVSNum = goCurCarList.ECDBVersion

    'BGS check the version of the Cur Db to that of the Template DB
    
    If sTemplateDBPath <> vbNullString Then
    
        bCheckVersionTable = True
        sSQL = "SELECT A.CarrierName, A.vsNumber FROM DB_VERSION A "
        utOpenDatabase TemplateDB, sTemplateDBPath, False, True
        
        Set RSCur = mCurDB.OpenRecordset(sSQL)
        Set RSTemplate = TemplateDB.OpenRecordset(sSQL)
        bCheckVersionTable = False
        
        If Not RSCur.EOF Then
            RSCur.MoveFirst
        End If
        If Not RSTemplate.EOF Then
            RSTemplate.MoveFirst
        End If
        
        'BGS we need to check the be sure that the CAT actually has the correct
        'Carrier DB .  If it doesn't then this is very bad !  give them a message
        'telling them the wrong DB is there
        If UCase(RSCur!CarrierName) <> UCase(sCarrierName) Then
            MsgBox "Your " & sCarrierName & " CAT DB has the WRONG Carrier!!  """ & RSCur!CarrierName & """", vbCritical + vbOKOnly
            GoTo CLEANUP
        End If
        
        'Also check the template DB if it has been screwed with
        If UCase(RSTemplate!CarrierName) <> UCase(sCarrierName) Then
            MsgBox "Your " & sCarrierName & " TEMPLATE DB is the WRONG Carrier!!  """ & RSTemplate!CarrierName & """", vbCritical + vbOKOnly
            GoTo CLEANUP
        'BGS check the template Version number to be sure that the correct
        'version has been packaged with the Eberls.exe
        ElseIf RSTemplate!vsNumber <> sCarrierVSNum Then
            MsgBox "The current " & sCarrierName & " TEMPLATE DB VS """ & RSTemplate!vsNumber & """ is the wrong DB Version for this Easy Claim release." & _
                   "You need the " & sCarrierName & " TEMPLATE DB VS """ & sCarrierVSNum & """.  Please contact support.", vbCritical + vbOKOnly
            GoTo CLEANUP
        End If
        
        'BGS If we are good to go up to here then...
        'check the Cat DB vs against the Template DB vs and prompt
        'that the DB will be updated if applicable
        If RSCur!vsNumber <> RSTemplate!vsNumber Then
            If Not pbDoPicDB Then
                MsgBox "Easy Claim needs to update the " & sCarrierName & " CAT DB VS """ & RSCur!vsNumber & """ to VS """ & RSTemplate!vsNumber & """.", vbInformation + vbOKOnly
            Else
                MsgBox "Easy Claim needs to update the " & sCarrierName & " PIC DB VS """ & RSCur!vsNumber & """ to VS """ & RSTemplate!vsNumber & """.", vbInformation + vbOKOnly
            End If
            RSCur.Close
            RSTemplate.Close
            TemplateDB.Close
            Set TemplateDB = Nothing
            GoTo UPDATE_TO_NEW_VERSION
        ElseIf pbForceUpdate Then
            GoTo UPDATE_TO_NEW_VERSION
        End If
        
        'Bgs if we get to here then the version is upto date so cleanup
        GoTo CLEANUP
        
UPDATE_TO_NEW_VERSION:
        'BGS 10.4.2001 Need to copy over the TemplateDB into the same path as the
        'CurDb PAth, But change  the Name to "TemplateDB.db"
        iPos = InStrRev(psCurDBPath, "\")
        'bgs 10.23.2001 be sure the read only is off
        
        If utFileExists(left(psCurDBPath, iPos) & "TemplateDB.db") Then
            SetAttr left(psCurDBPath, iPos) & "TemplateDB.db", vbNormal
        End If
        goUtil.utCopyFile sTemplateDBPath, left(psCurDBPath, iPos) & "TemplateDB.db"
        'BGS now change the template path string
        sTemplateDBPath = left(psCurDBPath, iPos) & "TemplateDB.db"
        
        'BGS 10.5.2001 set up the Cat Reports Dir
        sCatReportsDir = left(psCurDBPath, iPos)
        
        'now set the Template db variable to the template db we just copied over
        'into the Cat directory as the current DB path.
        utOpenDatabase TemplateDB, sTemplateDBPath, False, False
        
        If UpDateOldDB(psAppEXEName, mCurDB, psCurDBPath, TemplateDB, sTemplateDBPath, sTemplateReportsDir, sCatReportsDir) Then
            
            MsgBox "DB Update Successful!", vbInformation + vbOKOnly
            'Bgs now reset the gCurDB
            If Not pbDoPicDB Then
                Set gCurDB = Nothing
                SetCurDB psAppEXEName, sCompanyName, sCarrierName, psCurDBPath, gsCurCatDir, False
            End If
        Else
            MsgBox "DB" & psCurDBPath & " Not Updated!", vbCritical + vbOKOnly
        End If
    End If
    
CLEANUP:

    'BGS 1.10.2002 Only comp repair when selecting Make current from Navigator
    'BGS 1.14.2002 And When Updating Current Cat
    If pbCompRepair Then
        mCurDB.Close
        utCompRepair psAppEXEName, psCurDBPath
        'BGS open it again
        utOpenDatabase mCurDB, psCurDBPath, False, False
    End If
    
    Set TemplateDB = Nothing
    Set RSCur = Nothing
    Set RSTemplate = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    If Err.Number = NO_DB_VERSION_TABLE_ERROR And bCheckVersionTable Then
        Err.Clear
        MsgBox "Easy Claim needs to update the " & sCarrierName & " CAT DB VS ""(Version unavailable)"" to VS """ & sCarrierVSNum & """.", vbInformation + vbOKOnly
        GoTo UPDATE_TO_NEW_VERSION
    End If
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub CheckVSCurDB"
End Sub

Public Sub CheckVSMainDB(psAppEXEName As String, psMainDBPath As String, _
                        Optional pbForceUpdate As Boolean, _
                        Optional pbCompRepair As Boolean)
    On Error GoTo EH
    Dim iVSNum As Integer
    Dim sTemplateDBPath As String
    Dim TemplateDB As Database
    Dim RSMain As Recordset
    Dim RSTemplate As Recordset
    Dim iPos As Integer
    Dim sSQL As String
    Dim bCheckVersionTable As Boolean
    Dim sECMainDBVSNUM As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    

    sTemplateDBPath = gsInstallDir & "\Templates\"
    sTemplateDBPath = sTemplateDBPath & "ECMain.mdb"
    sECMainDBVSNUM = ECMainDBVersion

    'BGS check the version of the Main Db to that of the Template DB
    
    If sTemplateDBPath <> vbNullString Then
    
        bCheckVersionTable = True
        sSQL = "SELECT A.Version FROM DB_VERSION A "
        utOpenDatabase TemplateDB, sTemplateDBPath, False, True
        
        Set RSMain = mMainDB.OpenRecordset(sSQL)
        Set RSTemplate = TemplateDB.OpenRecordset(sSQL)
        bCheckVersionTable = False
        
        If Not RSMain.EOF Then
            RSMain.MoveFirst
        End If
        If Not RSTemplate.EOF Then
            RSTemplate.MoveFirst
        End If
        
        
        'Also check the template DB if it has been screwed with
        If RSTemplate!Version <> sECMainDBVSNUM Then
            MsgBox "The current ECmain.mdb TEMPLATE DB VS """ & RSTemplate!Version & """ is the wrong DB Version for this " & goUtil.gsMainAppEXEName & " release " & _
                   "You need the ECMain.mdb TEMPLATE DB VS """ & sECMainDBVSNUM & """.  Please contact support.", vbCritical + vbOKOnly
            GoTo CLEANUP
        End If
        
        'BGS If we are good to go up to here then...
        'check the current ECMain DB vs against the Template DB vs and prompt
        'that the DB will be updated if applicable
        If RSMain!Version <> RSTemplate!Version Then
            MsgBox goUtil.gsMainAppEXEName & " needs to update the ECmain.mdb DB VS """ & RSMain!Version & """ to VS """ & RSTemplate!Version & """.", vbInformation + vbOKOnly
            RSMain.Close
            RSTemplate.Close
            TemplateDB.Close
            Set TemplateDB = Nothing
            GoTo UPDATE_TO_NEW_VERSION
        ElseIf pbForceUpdate Then
            GoTo UPDATE_TO_NEW_VERSION
        End If
        
        'Bgs if we get to here then the version is upto date so cleanup
        GoTo CLEANUP
        
UPDATE_TO_NEW_VERSION:
        'BGS 10.4.2001 Need to copy over the TemplateDB into the same path as the
        'MainDB PAth, But change  the Name to "TemplateDB.db"
        iPos = InStrRev(psMainDBPath, "\")
        'bgs 10.23.2001 be sure the read only is off
        
        If utFileExists(left(psMainDBPath, iPos) & "TemplateDB.db") Then
            SetAttr left(psMainDBPath, iPos) & "TemplateDB.db", vbNormal
        End If
        FileCopy sTemplateDBPath, left(psMainDBPath, iPos) & "TemplateDB.mdb"
        'BGS now change the template path string
        sTemplateDBPath = left(psMainDBPath, iPos) & "TemplateDB.mdb"
        
        'now set the Template db variable to the template db we just copied over
        'into the Cat directory as the current DB path.
        utOpenDatabase TemplateDB, sTemplateDBPath, False, False
        
        If UpDateOldDB(psAppEXEName, mMainDB, psMainDBPath, TemplateDB, sTemplateDBPath, vbNullString, vbNullString, True) Then
            
            MsgBox "DB Update Successful!", vbInformation + vbOKOnly
            'Bgs now reset the gMainDB
            SetMainDB psAppEXEName, psMainDBPath, , False
        Else
            MsgBox "DB" & psMainDBPath & " Not Updated!", vbCritical + vbOKOnly
        End If
    End If
    
CLEANUP:
    If pbCompRepair Then
        mMainDB.Close
        utCompRepair psAppEXEName, psMainDBPath
        'BGS open it again
        utOpenDatabase mMainDB, psMainDBPath, False, False
    End If
    
    Set TemplateDB = Nothing
    Set RSMain = Nothing
    Set RSTemplate = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    If lErrNum = NO_DB_VERSION_TABLE_ERROR And bCheckVersionTable Then
        Err.Clear
        MsgBox goUtil.gsMainAppEXEName & " needs to update the ECMain.mdb DB VS ""(Version unavailable)"" to VS """ & sECMainDBVSNUM & """.", vbInformation + vbOKOnly
        GoTo UPDATE_TO_NEW_VERSION
    End If
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Sub CheckVSMainDB"
End Sub


Public Function UpDateOldDB(psAppEXEName As String, pOldDB As Database, psOldDbPath, pNewDB As Database, _
                             psNewDBPath As String, psReportSourceDir As String, _
                             psReportDestDir As String, Optional pbSkipReports As Boolean = False, _
                             Optional pbMakeCopy As Boolean = False, Optional pbMaxSpace As Boolean = False) As Boolean
    On Error GoTo EH
    
    Dim oldTablDef As dao.TableDef
    Dim oldRS As dao.Recordset
    Dim oldField As dao.Field
    Dim newTablDef As dao.TableDef
    Dim newRS As dao.Recordset
    Dim newField As dao.Field
    
    Dim frmProg As New frmProgress
    Dim iTableCount As Integer
    Dim iRecordCount As Long
    Dim sReport As String
    Dim iReportCount As Integer
    Dim saryReports() As String
    Dim iCount As Integer
    
    'BGS 10.26.2001 Use These variables for MakeCopy stuff
    Dim sOldDBName As String
    Dim sTemplateDBDir As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sOldDBName = Mid(psOldDbPath, InStrRev(psOldDbPath, "\") + 1)
    
    'BGS 10.4.2001 get the Progress bars ready
    Load frmProg
    
    'BGS see how many tables we are dealing with
    For Each oldTablDef In pOldDB.TableDefs
       If InStr(1, oldTablDef.Name, "msys", vbTextCompare) = 0 And InStr(1, oldTablDef.Name, "DB_VERSION", vbTextCompare) = 0 Then
          iTableCount = iTableCount + 1
       End If
    Next
    
    'BGS 10.5.2001 see how many reports we are dealing with
    
    If Not pbSkipReports Then
        'BGS 10.24.2001 If making a copy need to copy Everything From
        'Current Cat including Text Files and even the .db files.
        'The Eberls .db May or may not be compacted by the Zip function
        'when storing files to floppy.
        'BGS 2.8.2002 to maximize space when exporting or backng up to
        'floppy we are going to exclude files that can be replaced from
        'the Templates folder.  Files that will be excluded are:
        '*_Backup.db (This is made every time the Db is Made Current or Compacted)
        '*.RPT (These reports can be retrieved fom Templates folder)
        '*.ldb (This temp lock file totlally useless)
        '"Pictures.db" (This is a relic piece of poop this process will nuke it yippy skippy")
        If pbMakeCopy Then
            sReport = Dir(psReportSourceDir & "*.*")
        Else
            sReport = Dir(psReportSourceDir & "*.*")
        End If
        Do Until sReport = vbNullString
            If pbMakeCopy And pbMaxSpace Then
                If InStr(1, sReport, "_BackUp.db", vbTextCompare) > 0 Then
                    GoTo NEXT_FILE
                ElseIf InStr(1, sReport, ".rpt", vbTextCompare) > 0 Then
                    GoTo NEXT_FILE
                ElseIf InStr(1, sReport, ".ldb", vbTextCompare) > 0 Then
                    GoTo NEXT_FILE
                ElseIf InStr(1, sReport, "Pictures.db", vbTextCompare) > 0 Then
                    GoTo NEXT_FILE
                End If
            End If
            iReportCount = iReportCount + 1
            ReDim Preserve saryReports(1 To iReportCount)
            saryReports(iReportCount) = sReport
NEXT_FILE:
            sReport = Dir
        Loop
    End If
    
    
    'BGS Pbar Stuff
    With frmProg
        .PBarTable.Max = iTableCount
        .PBarTable.Value = 0
        If Not pbSkipReports Then
            .PBarFile.Max = iReportCount
            .PBarFile.Value = 0
        End If
        .Show vbModeless
        .Refresh
    End With
    UpDateOldDB = True
    

    'BGS Pbar Stuff
    iTableCount = 0
    
        With frmProg
        For Each oldTablDef In pOldDB.TableDefs
            If InStr(1, oldTablDef.Name, "msys", vbTextCompare) = 0 And InStr(1, oldTablDef.Name, "DB_VERSION", vbTextCompare) = 0 Then
                'BGS Pbar Stuff
                iTableCount = iTableCount + 1
                .PBarTable.Value = iTableCount
                .lblTable.Caption = iTableCount & " Of " & .PBarTable.Max & " (" & oldTablDef.Name & ")"
                .RefreshMe
                
                For Each newTablDef In pNewDB.TableDefs
                    If oldTablDef.Name = newTablDef.Name Then
                        Set oldRS = oldTablDef.OpenRecordset
                        Set newRS = newTablDef.OpenRecordset
                        
                        If Not oldRS.EOF Then oldRS.MoveFirst
                        'BGS Pbar Stuff
                        iRecordCount = oldRS.RecordCount
                        If iRecordCount > 0 Then
                            .PBarRecord.Max = iRecordCount
                            iRecordCount = 0
                        End If
                        
                        Do While Not oldRS.EOF
                            'BGS Pbar Stuff
                            iRecordCount = iRecordCount + 1
                            .PBarRecord.Value = iRecordCount
                            newRS.AddNew
                            For Each oldField In oldRS.Fields
                                newRS.Fields(oldField.Name).Value = oldRS.Fields(oldField.Name).Value
                                'BGS Pbar Stuff
                                .lblField.Caption = iRecordCount & " Of " & .PBarRecord.Max & " (" & oldField.Name & ")"
                                .lblField.Refresh
                            Next
                            newRS.Update
                            oldRS.MoveNext
                        Loop
                        newRS.Close
                        oldRS.Close
                        Exit For
                    End If
                Next
            End If
        Next
        
        If Not pbSkipReports Then
            iReportCount = 0
            For iCount = 1 To UBound(saryReports, 1)
                iReportCount = iReportCount + 1
                'BGS 10.23.2001 if it is readonly need to make it normal
                If utFileExists(psReportDestDir & saryReports(iCount)) Then
                    SetAttr psReportDestDir & saryReports(iCount), vbNormal
                End If
                If pbMakeCopy And saryReports(iCount) <> sOldDBName And saryReports(iCount) <> Replace(sOldDBName, ".db", ".ldb") Then
                    goUtil.utCopyFile psReportSourceDir & saryReports(iCount), psReportDestDir & saryReports(iCount)
                ElseIf Not pbMakeCopy Then
                    '8.27.2002 Only copy over Rates.ini and LossCauses.ini if they do not already exist
                    If StrComp(saryReports(iCount), "Rates.ini", vbTextCompare) = 0 Or StrComp(saryReports(iCount), "LossCauses.ini", vbTextCompare) = 0 Then
                        If Not utFileExists(psReportDestDir & saryReports(iCount)) Then
                            goUtil.utCopyFile psReportSourceDir & saryReports(iCount), psReportDestDir & saryReports(iCount)
                        End If
                    ElseIf InStr(1, saryReports(iCount), ".db", vbTextCompare) = 0 And InStr(1, saryReports(iCount), ".ldb", vbTextCompare) = 0 And InStr(1, saryReports(iCount), ".mdb", vbTextCompare) = 0 And InStr(1, saryReports(iCount), ".dat", vbTextCompare) = 0 Then
                        goUtil.utCopyFile psReportSourceDir & saryReports(iCount), psReportDestDir & saryReports(iCount)
                    End If
                End If
                
                .lblFile.Caption = iReportCount & " Of " & .PBarFile.Max & " (" & saryReports(iCount) & ")"
                .PBarFile.Value = iReportCount
                .lblFile.Refresh
            Next
        End If
    End With
    
    
    'Clean up
    Set oldRS = Nothing
    Set newRS = Nothing
    Set oldField = Nothing
    Set newField = Nothing
    Set oldTablDef = Nothing
    Set newTablDef = Nothing

    'BGS now close both dbs and copy over the New to the old
    'BGS 10.23.2001 ONLY if we are making a copy, leave the oldDB open :)
    If Not pbMakeCopy Then
        pOldDB.Close
        Set pOldDB = Nothing
    End If
    pNewDB.Close
    Set pNewDB = Nothing
    
    If Not pbMakeCopy Then
        FileCopy psNewDBPath, psOldDbPath
        'Kill the temp one
        Kill psNewDBPath
    Else 'If we making a copy need to change the Template Name
        sTemplateDBDir = left(psNewDBPath, InStrRev(psNewDBPath, "\"))
        'BGS 2.8.2002 need to compact the db to maximize space
        'Skip backup oreles we will have double the space DB to floppy
        If pbMaxSpace Then
            utCompRepair psAppEXEName, psNewDBPath, True
        End If
        FileCopy psNewDBPath, sTemplateDBDir & sOldDBName
        Kill psNewDBPath
    End If
    
    'BGS Pbar Stuff
    Unload frmProg
    Set frmProg = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    If lErrNum = DB_VERSION_DUPLICATE_VALUE_ERROR Then
        Err.Clear
        Resume Next
    ElseIf lErrNum = DB_VERSION_ITEM_NOT_FOUND_ERROR Then
        Err.Clear
        Resume Next
    End If
    
    Unload frmProg
    Set frmProg = Nothing
    UpDateOldDB = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function UpDateOldDB"
End Function

Public Function DB_PASSWORD(psKey As String) As String
    On Error GoTo EH
    If Not mdicDB_PASSWORD Is Nothing Then
        DB_PASSWORD = mdicDB_PASSWORD(psKey)
    End If
    If Err.Number > 0 Then
        Err.Clear
    End If
    
    Exit Function
EH:
    Err.Clear
    DB_PASSWORD = "Invalid Key!"
End Function

Public Sub ShowAllForms(oECTray As Object, psNavPos As String, Optional psSkipForms As String)
    utShowAllForms Forms, oECTray, psNavPos, psSkipForms & "frmSP"
End Sub

Public Sub HideAllForms(Optional psSkipForms As String)
    Dim sSkipForms As String
    sSkipForms = psSkipForms & "frmProgress"
    utHideAllForms Forms, sSkipForms
End Sub

Public Function AddItemToDictionary(pDictionary As scripting.Dictionary, psItem As String, psKey As String) As Boolean
    On Error GoTo EH
    
    If pDictionary Is Nothing Then
        Set pDictionary = New scripting.Dictionary
    End If
    
    pDictionary.Add psKey, psItem
    
    AddItemToDictionary = True
    
    Exit Function
EH:
    Err.Clear
    AddItemToDictionary = False
End Function

Public Function RemoveItemFromDictionary(pDictionary As scripting.Dictionary, psKey As String) As Boolean
    On Error GoTo EH
    
    If Not pDictionary Is Nothing Then
        If pDictionary.Exists(psKey) Then
            pDictionary.Remove psKey
            RemoveItemFromDictionary = True
        Else
            RemoveItemFromDictionary = False
        End If
    Else
        RemoveItemFromDictionary = False
    End If
    
    Exit Function
EH:
    Err.Clear
    RemoveItemFromDictionary = False
End Function

Public Function GetItemFromDictionary(pDictionary As scripting.Dictionary, psKey As String) As String
    On Error GoTo EH
    
    If Not pDictionary Is Nothing Then
        If pDictionary.Exists(psKey) Then
            GetItemFromDictionary = pDictionary(psKey)
        Else
            GetItemFromDictionary = vbNullString
        End If
    Else
        GetItemFromDictionary = vbNullString
    End If
    
    Exit Function
EH:
    Err.Clear
    GetItemFromDictionary = vbNullString
End Function

Public Function AddItemToCollection(pCol As Collection, pvItem As Variant, psKey As String) As Boolean
    On Error GoTo EH
    
    If pCol Is Nothing Then
        Set pCol = New Collection
    End If
    
    pCol.Add psKey, pvItem
    
    AddItemToCollection = True
    
    Exit Function
EH:
    Err.Clear
    AddItemToCollection = False
End Function

Public Function RemoveItemFromCollection(pCol As Collection, psKey As String) As Boolean
    On Error GoTo EH
    
    If Not pCol Is Nothing Then
        pCol.Remove psKey
    Else
        RemoveItemFromCollection = False
    End If
    
    RemoveItemFromCollection = True
    
    Exit Function
EH:
    Err.Clear
    RemoveItemFromCollection = False
End Function

Public Function GetItemFromCollection(pCol As Collection, psKey As String) As Variant
    On Error GoTo EH
    
    If Not pCol Is Nothing Then
        GetItemFromCollection = pCol(psKey)
    Else
        GetItemFromCollection = vbNullString
    End If
    
    Exit Function
EH:
    Err.Clear
    GetItemFromCollection = vbNullString
End Function


Public Function UpdateCatDB(psCompany As String, sCar As String, sCat As String, sDir As String) As Boolean
    On Error GoTo EH
    UpdateCatDB = True
    'Update Cat with FTP Download from Server
    
    Exit Function
EH:
    UpdateCatDB = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UpdateCatDB"
End Function

Public Function ValidCatStructure(psSection As String, psKey As String, _
                                    Optional poNodX As Object, _
                                    Optional poTree As Object, _
                                    Optional psCompany As String, _
                                    Optional psCompanyName As String, _
                                    Optional psCar As String, _
                                    Optional psCarName As String, _
                                    Optional psCAT As String, _
                                    Optional psCatName As String, _
                                    Optional pbAddedDeleteAllFromRecycleBin As Boolean, _
                                    Optional psAddedCompanyRecycleBin As String, _
                                    Optional psAddedCarRecycleBin As String, _
                                    Optional pbCatActive As Boolean) As Boolean
    On Error GoTo EH
    Dim sTemp As String
    Dim Nodx As Node
    Dim oECtree As TreeView
    Dim sRelative As String
    
    If Not poNodX Is Nothing And Not poTree Is Nothing Then
        Set Nodx = poNodX
        Set oECtree = poTree
    End If

    'Check against Registry structure first, it may be a bogus Carrier folder
    'The Name of the Carrier should match the Encrpted Key value for that Carrier
    
    sTemp = GetSetting(goUtil.gsMainAppEXEName, "CAT_STRUCTURE\" & psSection, psKey, vbNullString)
    On Error Resume Next
    sTemp = goUtil.Decode(sTemp)
    'If the decode errors then either the encrypted registry setting was tampered with
    'or the folder was added manually
    If Err.Number > 0 Then
        Err.Clear
        Set poNodX = Nothing
        ValidCatStructure = False
    'If the string compare does not match then the Encrypted setting was manipulated
    ElseIf StrComp(sTemp, psKey, vbTextCompare) <> 0 Then
        On Error GoTo EH
        If Not Nodx Is Nothing Then
            'See if this has been marked for Recyclebin.
            If StrComp(sTemp, "RECYCLEBIN", vbTextCompare) = 0 Then
                Nodx.Image = PicTree.A17_RecycleBinFull
                'Add the Empty Recycle Bin Icon if it hasn't already
                If Not pbAddedDeleteAllFromRecycleBin Then
                    Set Nodx = oECtree.Nodes.Add(goUtil.gsMainAppEXEName & "|RecycleBin", tvwChild, goUtil.gsMainAppEXEName & "|RecycleBin|Empty", "Empty Recycle Bin", PicTree.A19_DeleteAllFromRecycleBin)
                    pbAddedDeleteAllFromRecycleBin = True
                End If
                
                'Add this Company Folder to the Recycle Bin if it hasn't already
                sTemp = goUtil.gsMainAppEXEName & "|RecycleBin|" & psCompany
                If InStr(1, psAddedCompanyRecycleBin, psCompany, vbTextCompare) = 0 Then
                    Set Nodx = oECtree.Nodes.Add(goUtil.gsMainAppEXEName & "|RecycleBin", tvwChild, sTemp, psCompanyName, PicTree.A09_ClosedFolder)
                    psAddedCompanyRecycleBin = psAddedCompanyRecycleBin & psCompany & "|"
                End If
                sRelative = sTemp
                
                
                'Add this Car Folder to the Recycle Bin if it hasn't already
                sTemp = sTemp & "|" & psCar
                If InStr(1, psAddedCarRecycleBin, psCar, vbTextCompare) = 0 Then
                    Set Nodx = oECtree.Nodes.Add(sRelative, tvwChild, sTemp, psCarName, PicTree.A09_ClosedFolder)
                    psAddedCarRecycleBin = psAddedCarRecycleBin & psCar & "|"
                End If
                sRelative = sTemp
                'Add this Cat to this Car Folder
                Set Nodx = oECtree.Nodes.Add(sRelative, tvwChild, sTemp & "|" & psCAT, psCatName, IIf(pbCatActive, PicTree.A14_ClosedHurc, PicTree.A42_closedHurc_Inactive))
                'Add Restor and Delete icons to this cat
                Set Nodx = oECtree.Nodes.Add(sTemp & "|" & psCAT, tvwChild, sTemp & "|" & psCAT & "|RestoreCat", "Restore Cat", PicTree.A21_RestoreCat)
                Set Nodx = oECtree.Nodes.Add(sTemp & "|" & psCAT, tvwChild, sTemp & "|" & psCAT & "|DeleteCat", "Delete Cat", PicTree.A20_DeleteCat)
            Else
                Set poNodX = Nothing
            End If
        End If
        ValidCatStructure = False
    'If we get here that means the Carrier folder matched the Registry, good to go
    Else
        ValidCatStructure = True
    End If
    
    'Clean up
    Set oECtree = Nothing
    Exit Function
EH:
    ValidCatStructure = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ValidCatStructure"
End Function

Public Function utShellExecute(Optional plHwnd As Long = -1, _
                                Optional pslpOperation As String = "OPEN", _
                                Optional pslpFile As String, _
                                Optional pslpParameters As String = vbNullString, _
                                Optional pslpDirectory As String = "App.Path", _
                                Optional plnShowCmd As VBA.VbAppWinStyle = vbNormalFocus, _
                                Optional pbCreateTempFile As Boolean = False, _
                                Optional pbUseTimeStampFileName As Boolean = False, _
                                Optional pbShowMessage As Boolean = False, _
                                Optional psTempFileCaption As String) As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim slpOperation As String
    Dim slpFile As String
    Dim slpParameters As String
    Dim slpDirectory As String
    Dim lnShowCmd As VBA.VbAppWinStyle
    Dim sErrorMess As String
    Dim sTmpExt As String
    Dim sTmpFile As String
    Dim lRet As Long
    Dim sDir As String
    
    utShellExecute = False
    
    'Get info from Params
    If plHwnd = -1 Then
        lHwnd = GetDesktopWindow
    End If
    slpOperation = pslpOperation
    If pslpFile = vbNullString Then
        Exit Function
    Else
        slpFile = pslpFile
    End If
     
    slpParameters = pslpParameters
    If pslpDirectory = "App.Path" Then
        If goUtil.gsInstallDir <> vbNullString Then
            slpDirectory = goUtil.gsInstallDir
        Else
            slpDirectory = App.Path
        End If
    Else
        slpDirectory = pslpDirectory
    End If
    
    lnShowCmd = plnShowCmd
    
    'See if the file exists if not Then Exit
    If goUtil.utFileExists(slpFile) Or InStr(1, slpFile, "MAPIMAIL", vbTextCompare) > 0 Then
    
        If pbCreateTempFile Then
            'Be sure the Temp dir exixts
            If Not CreateTempDir(sErrorMess) Then
                GoTo SHOW_ERROR
            End If
            'See if need to create a Temp file name
            If pbUseTimeStampFileName Then
                sTmpExt = Mid(slpFile, InStrRev(slpFile, ".", , vbBinaryCompare))
                sTmpFile = "C:\Temp\Temp_" & psTempFileCaption & "_" & CStr(timeGetTime()) & sTmpExt
            Else
                sTmpFile = "C:\Temp\Temp_" & psTempFileCaption & "_" & Mid(slpFile, InStrRev(slpFile, "\", , vbBinaryCompare) + 1)
            End If
            sErrorMess = goUtil.utCopyFile(slpFile, sTmpFile)
        Else
            sTmpFile = slpFile
        End If
        
        lRet = ShellExecute(lHwnd, slpOperation, sTmpFile, slpParameters, slpDirectory, lnShowCmd)
            'Check to see if the Associated Application opened the file
            'If not that means there is no application associated with the file
            'in that case open Explorer to give the user a chance to open the file
        If lRet = SE_ERR_NOASSOC Then
            sDir = Space(260)
            lRet = GetSystemDirectory(sDir, Len(sDir))
            sDir = left(sDir, lRet)
            lRet = ShellExecute(lHwnd, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & sTmpFile, sDir, lnShowCmd)
        End If
    Else
SHOW_ERROR:
        If pbShowMessage Then
            If sErrorMess = vbNullString Then
                sErrorMess = "File Not Found!" & vbCrLf & psTempFileCaption & vbCrLf & slpFile
            End If
            MsgBox sErrorMess, vbExclamation + vbOKOnly, "File Error"
        End If
    End If
        utShellExecute = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function utShellExecute"
End Function


Public Function SaveCatStructure(psSection As String, psKey As String) As Boolean
    On Error GoTo EH
    SaveSetting goUtil.gsMainAppEXEName, "CAT_STRUCTURE\" & psSection, psKey, goUtil.Encode(psKey)
    SaveCatStructure = True
    Exit Function
EH:
    SaveCatStructure = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveCatStructure"
End Function


Private Sub Class_Initialize()
    Set mclsKeyBoard = New clsKeyBoard
    Set moTimeZone = New clsTimeZone
End Sub

Public Function SetGlobalObjects(pcolObjects As Collection) As Boolean
    On Error Resume Next
    Set goUtil = pcolObjects("goUtil")
    If Err.Number = 0 Then
        SetGlobalObjects = True
    Else
        SetGlobalObjects = False
    End If
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        SetUtilObject = True
    Else
        SetUtilObject = False
    End If
End Function

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mfrmECTray = Nothing
    Set mfrmWallPaper = Nothing
    
    If Not moCurCarList Is Nothing Then
        moCurCarList.CLEANUP
        Set moCurCarList = Nothing
    End If
    
    If Not moProgForm Is Nothing Then
        moProgForm.CLEANUP
        Set moProgForm = Nothing
    End If
    If Not mclsKeyBoard Is Nothing Then
        mclsKeyBoard.CLEANUP
        Set mclsKeyBoard = Nothing
    End If
    
    If Not moSP Is Nothing Then
        moSP.CLEANUP
        Set moSP = Nothing
    End If
    If Not moXact Is Nothing Then
        moXact.CLEANUP
        Set moXact = Nothing
    End If
    If Not moECKeyBoardList Is Nothing Then
        moECKeyBoardList.CLEANUP
        Set moECKeyBoardList = Nothing
    End If
    If Not mARV Is Nothing Then
        mARV.CLEANUP
        Set mARV = Nothing
    End If
    Set mdicDB_PASSWORD = Nothing
   
    CLEANUPModUtil
    
    CloseMainDB
    CloseCurDB
    
    'Delete any files that were created under Temp Directory
    DelTempDirFiles
    
    Set goUtil = Nothing
    
    CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CLEANUP = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function CLEANUP"
End Function


Public Function CreateTempDir(Optional psErrorMess As String) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sError As String
    
    CreateTempDir = False
    
    If Not goUtil.utFileExists("C:\Temp", True) Then
        'Create C:\Temp dir
        sError = goUtil.utMakeDir("C:\Temp")
        If sError <> vbNullString Then
            psErrorMess = sError
            Exit Function
        End If
    End If
    
    CreateTempDir = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CreateTempDir = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function CreateTempDir"
End Function

Public Function DelTempDirFiles(Optional psTempDir As String = "C:\Temp", _
                                Optional psTempFile As String = "Temp_*") As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sTempDir As String
    Dim sTempFile As String
    Dim sErrDesc As String
    Dim sError As String
    
    DelTempDirFiles = False
    
    sTempDir = psTempDir
    sTempFile = psTempFile
    
    If FileExists(sTempDir, True) Then
        'Clear any temp files created with "Temp_*"
        If Dir(sTempDir & "\" & sTempFile, vbNormal) <> vbNullString Then
            sError = DeleteFile(sTempDir & "\" & sTempFile)
        End If
    End If
    
    DelTempDirFiles = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description & vbCrLf & sError
    DelTempDirFiles = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function DelTempDirFiles"
End Function

Public Function ConvertTwipsToPixels(plTwips As Long) As Long
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    ConvertTwipsToPixels = plTwips / 15
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function ConvertTwipsToPixels"
End Function

Public Function ConvertPixelsToTwips(plPixels As Long) As Long
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    ConvertPixelsToTwips = plPixels * 15
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function ConvertPixelsToTwips"
End Function

Public Function IsNullIsVbNullString(poField As ADODB.Field) As String
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If IsNull(poField.Value) Then
        Select Case poField.Type
            Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp
                IsNullIsVbNullString = vbNullString
            Case ADODB.adBigInt, ADODB.adBoolean, ADODB.adCurrency, ADODB.adDecimal, ADODB.adDouble, ADODB.adFldLong, ADODB.adInteger, ADODB.adNumeric, ADODB.adSingle, ADODB.adSmallInt, ADODB.adTinyInt
                IsNullIsVbNullString = "0"
            Case Else
                IsNullIsVbNullString = vbNullString
        End Select
    Else
        Select Case poField.Type
            Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp
                If poField.Value = NULL_DATE Then
                    IsNullIsVbNullString = vbNullString
                Else
                    IsNullIsVbNullString = poField.Value
                End If
            Case Else
                IsNullIsVbNullString = poField.Value
        End Select
        
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function IsNullIsVbNullString"
End Function

Public Function GetFlagText(pbFlag As Boolean) As String
    On Error GoTo EH
    Dim sFlagText As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If pbFlag Then
        sFlagText = Chr(160)
    Else
        sFlagText = Chr(32)
    End If
    
    GetFlagText = sFlagText
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function GetFlagText"
End Function

Public Function GetFlagFromText(psFlagText As String) As Boolean
    On Error GoTo EH
    Dim bFlag As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If psFlagText = Chr(160) Then
        bFlag = True
    Else
        bFlag = False
    End If
    
    GetFlagFromText = bFlag
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function GetFlagFromText"
End Function

Private Sub Class_Terminate()
    Set moTimeZone = Nothing
End Sub

Private Sub moXact_ProgressMessage(ByVal sMessage As String)
     RaiseEvent XactimateProgress(sMessage)
End Sub

Public Function GetAccessDBUID(psIDName As String, psTableName As String) As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String
    Dim sID As String
    Dim sIDName As String
    Dim sTableName As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sIDName = psIDName
    sTableName = psTableName
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    utOpenDatabaseADOConn oConn, gMainDB.Name
    
    'Get the next ID
    sSQL = "SELECT MIN(" & sIDName & ") -1 As ID  "
    sSQL = sSQL & "FROM " & sTableName & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    If Not adoRS.EOF Then
        adoRS.MoveFirst
        sID = IsNullIsVbNullString(adoRS.Fields("ID"))
        'If the Minimum is not a Negative Number then start the
        'Negative number at -1
        If IsNumeric(sID) Then
            If CDbl(sID) > -1 Then
                sID = "-1"
            End If
        Else
            sID = "-1"
        End If
    Else
        sID = "-1"
    End If
    
    GetAccessDBUID = sID
    
    'cleanup
    Set adoRS = Nothing
    Set oConn = Nothing
    
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function GetAccessDBUID"
End Function

Public Function GetvbVarTypeFromText(psvbVarTypeText As String) As VBA.VbVarType
    On Error GoTo EH
    Dim sDataType As String
    Dim iDataType As VBA.VbVarType
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sDataType = psvbVarTypeText
    
    If StrComp(sDataType, "vbArray", vbTextCompare) = 0 Then
        iDataType = vbArray
    ElseIf StrComp(sDataType, "vbBoolean", vbTextCompare) = 0 Then
        iDataType = vbBoolean
    ElseIf StrComp(sDataType, "vbByte", vbTextCompare) = 0 Then
        iDataType = vbByte
    ElseIf StrComp(sDataType, "vbCurrency", vbTextCompare) = 0 Then
        iDataType = vbCurrency
    ElseIf StrComp(sDataType, "vbDataObject", vbTextCompare) = 0 Then
        iDataType = vbDataObject
    ElseIf StrComp(sDataType, "vbDate", vbTextCompare) = 0 Then
       iDataType = vbDate
    ElseIf StrComp(sDataType, "vbDecimal", vbTextCompare) = 0 Then
       iDataType = vbDecimal
    ElseIf StrComp(sDataType, "vbDouble", vbTextCompare) = 0 Then
       iDataType = vbDouble
    ElseIf StrComp(sDataType, "vbEmpty", vbTextCompare) = 0 Then
       iDataType = vbEmpty
    ElseIf StrComp(sDataType, "vbError", vbTextCompare) = 0 Then
       iDataType = vbError
    ElseIf StrComp(sDataType, "vbInteger", vbTextCompare) = 0 Then
       iDataType = vbInteger
    ElseIf StrComp(sDataType, "vbLong", vbTextCompare) = 0 Then
       iDataType = vbLong
    ElseIf StrComp(sDataType, "vbNull", vbTextCompare) = 0 Then
       iDataType = vbNull
    ElseIf StrComp(sDataType, "vbObject", vbTextCompare) = 0 Then
       iDataType = vbObject
    ElseIf StrComp(sDataType, "vbSingle", vbTextCompare) = 0 Then
       iDataType = vbSingle
    ElseIf StrComp(sDataType, "vbString", vbTextCompare) = 0 Then
        iDataType = vbString
    ElseIf StrComp(sDataType, "vbUserDefinedType", vbTextCompare) = 0 Then
        iDataType = vbUserDefinedType
    ElseIf StrComp(sDataType, "vbVariant", vbTextCompare) = 0 Then
        iDataType = vbVariant
    Else
        iDataType = vbVariant
    End If
    
    GetvbVarTypeFromText = iDataType
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function GetvbVarTypeFromText"
End Function


Public Function GetApplicationVersionNumber(psMainSPVersion As Long, psProjectName As String, poConn As ADODB.Connection) As Long
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim lVersion As Long
    
    sSQL = "SELECT  Version "
    sSQL = sSQL & "FROM     Application "
    sSQL = sSQL & "WHERE    '" & utCleanSQLString(psProjectName) & "' = ProjectName "
    sSQL = sSQL & "AND      " & psMainSPVersion & " >= SPVersionBase "
    sSQL = sSQL & "AND      " & psMainSPVersion & " <= SPVersion "
    sSQL = sSQL & "AND      IsDeleted = 0 "

    Set RS = New ADODB.Recordset
    'Use passed in Connection since this can be used
    'on ACCESS or SQL Server Connection
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, poConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        RS.MoveFirst
        lVersion = RS.Fields("Version").Value
    Else
        Set RS = New ADODB.Recordset
        'If the version number is not found in the main table then
        'It must be in the History table
        sSQL = "SELECT  Version "
        sSQL = sSQL & "FROM     ApplicationHistory "
        sSQL = sSQL & "WHERE    '" & utCleanSQLString(psProjectName) & "' = ProjectName "
        sSQL = sSQL & "AND      " & psMainSPVersion & " >= SPVersionBase "
        sSQL = sSQL & "AND      " & psMainSPVersion & " <= SPVersion "
        sSQL = sSQL & "AND      IsDeleted = 0 "
        
        RS.CursorLocation = adUseClient
        RS.Open sSQL, poConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
        
        If Not RS.EOF Then
            RS.MoveFirst
            lVersion = RS.Fields("Version").Value
        Else
            lVersion = 0
        End If
    End If
    
    GetApplicationVersionNumber = lVersion
    
    'Cleanup
    Set RS = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function GetApplicationVersionNumber"
End Function

