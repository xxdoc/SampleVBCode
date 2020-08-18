VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mXZip As XceedZip
Attribute mXZip.VB_VarHelpID = -1
Private mobjZip As clsXZip
Private mlKey As Long
Private msKey As String

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsZip"
End Property

Public Property Get XZip() As XceedZip
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mXZip = Nothing
    Set mXZip = New XceedZip
    Set XZip = mXZip
    Exit Property
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Property Get XZip"
End Property

Public Function SaveZIPFiles(psSourceDir As String, psZipName As String, _
                             Optional psFilter As String = "*", _
                             Optional psEncryptedPassWord As String, _
                             Optional psDestDir As String) As Boolean

    On Error GoTo EH
    Dim ResultCode As xcdError
    Dim lPermissionErrorCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mobjZip = New clsXZip
    Set mXZip = mobjZip.XZip
    
    ' All properties keep their default values except the two below
    mXZip.FilesToProcess = psSourceDir & "\" & psFilter
    If FileExists(psDestDir, True) Then
        mXZip.ZipFilename = psDestDir & "\" & psZipName
    Else
        mXZip.ZipFilename = psSourceDir & "\" & psZipName
    End If
    If psEncryptedPassWord <> vbNullString Then
        mXZip.EncryptionPassword = goUtil.Decode(psEncryptedPassWord)
    End If
    mXZip.DeleteZippedFiles = True
    mXZip.CompressionLevel = xclHigh
    mXZip.ExtraHeaders = xehSecurityDescriptor
    mXZip.PreservePaths = False
    ' Start zipping
    Screen.MousePointer = vbHourglass
    lPermissionErrorCount = 0
ZIPME:
    ResultCode = mXZip.Zip
   ' Check the return value.
    If ResultCode <> xerSuccess Then
        'If we get a persmision error means we tried to retrieve
        'the file while it was still being written to disk.
        'Give it a chance to release permissions.
        If ResultCode = xerOpenZipFile And ResultCode <> xerFilesSkipped Then
            lPermissionErrorCount = lPermissionErrorCount + 1
            Sleep 500
            If lPermissionErrorCount <= 20 Then
                GoTo ZIPME
            End If
        End If
        If ResultCode <> xerWarnings Then
            Err.Raise ResultCode, , mXZip.GetErrorDescription(xvtError, ResultCode) & vbCrLf & mXZip.ZipFilename
        End If
    Else
        SaveZIPFiles = True
    End If
    'Clean up
    Set mXZip = Nothing
    Set mobjZip = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    SaveZIPFiles = False
    Screen.MousePointer = vbDefault
    'Clean up
    Set mXZip = Nothing
    Set mobjZip = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function SaveZIPFiles"
End Function

Public Function UNZipFiles(psUnZipToDir As String, psNameZipPath As String, _
                           Optional pbPreservePaths As Boolean) As Boolean
    On Error GoTo EH
    Dim ResultCode As xcdError
    Dim lPermissionErrorCount As Long
    Dim sPassword As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    mlKey = 1
    msKey = CStr(mlKey)
    sPassword = goUtil.DB_PASSWORD(msKey)
    If sPassword = INVALID_DB_PASSWORD_KEY Then
        Err.Raise -999, , "Invalid Password!"
        Exit Function
    End If
    
    Set mobjZip = New clsXZip
    Set mXZip = mobjZip.XZip
    
    
    ' All properties keep their default values except the two below
    mXZip.UnzipToFolder = psUnZipToDir
    mXZip.ZipFilename = psNameZipPath
    mXZip.EncryptionPassword = goUtil.Decode(sPassword)
    
    mXZip.PreservePaths = pbPreservePaths
    mXZip.FilesToProcess = vbNullString

    ' Start unzipping
    Screen.MousePointer = vbHourglass
    lPermissionErrorCount = 0
UNZIPME:
    ResultCode = mXZip.Unzip
   ' Check the return value.
    If ResultCode <> xerSuccess Then
        'If we get a persmision error means we tried to retrieve
        'the file while it was still being written to disk.
        'Give it a chance to release permissions.
        If ResultCode = xerOpenZipFile Then
            lPermissionErrorCount = lPermissionErrorCount + 1
            Sleep 500
            If lPermissionErrorCount <= 20 Then
                GoTo UNZIPME
            End If
        End If
        If ResultCode <> xerWarnings Then
            Err.Raise ResultCode, , mXZip.GetErrorDescription(xvtError, ResultCode) & vbCrLf & mXZip.ZipFilename
        End If
    Else
        Screen.MousePointer = vbDefault
        UNZipFiles = True
    End If
    
    'Clean up
    Set mXZip = Nothing
    Set mobjZip = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    UNZipFiles = False
    Screen.MousePointer = vbDefault
    'Clean up
    Set mXZip = Nothing
    Set mobjZip = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function UNZipFiles"
End Function




Private Sub mXZip_InsertDisk(ByVal lDiskNumber As Long, bDiskInserted As Boolean)
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    sMess = "---------------------------------------------------------------------" & vbCrLf
    sMess = sMess & "When saving/retrieving to/from multiple Disks, " & vbCrLf
    sMess = sMess & "it is VERY IMPORTANT to label and insert the " & vbCrLf
    sMess = sMess & "appropriate Disk number when asked.  If you insert " & vbCrLf
    sMess = sMess & "the wrong disk number YOU WILL LOSE YOUR DATA!" & vbCrLf
    sMess = sMess & "***ALWAYS USE AN EMPTY DISK IF SAVING TO FLOPPY***" & vbCrLf
    sMess = sMess & "When retrieving from multiple disks, you must insert " & vbCrLf
    sMess = sMess & "the last numbered disk first, then follow the prompts.  The last " & vbCrLf
    sMess = sMess & "disk is where the Backup(*.bkp) or Export(*.exp) file will be saved." & vbCrLf
    sMess = sMess & "---------------------------------------------------------------------" & vbCrLf
    
'    sMess = "Please Insert Disk # " & lDiskNumber & " and press OK when ready." & vbCrLf & vbCrLf & sMess
'    If MsgBox(sMess, vbOKCancel, "Insert Another Disk") = vbOK Then
'        bDiskInserted = True
'    End If

    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub mXZip_InsertDisk"
End Sub

Private Sub mXZip_InvalidPassword(ByVal sFileName As String, sNewPassword As String, bRetry As Boolean)
    On Error GoTo EH
    Dim sPassword As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    mlKey = mlKey + 1
    msKey = CStr(mlKey)
    
    sPassword = goUtil.DB_PASSWORD(msKey)
    If sPassword = INVALID_DB_PASSWORD_KEY Then
        bRetry = False
    Else
        If sPassword = vbNullString Then
            bRetry = False
        Else
            sNewPassword = goUtil.Decode(sPassword)
            bRetry = True
        End If
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub mXZip_InvalidPassword"
End Sub

Public Function SetUtilObject(oUtilObject As Object) As Boolean
    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        SetUtilObject = True
    Else
        SetUtilObject = False
    End If
End Function

