Attribute VB_Name = "Globals"
Option Explicit

'WC Service Enum
Public Enum Netcmd
    NetStart = 0
    NetStop
    NetPause
    NetContinue
    NetRestart
End Enum


'Application Constants
Public Const LIST_ADJFTP_DELIM As String = "|"
Public Const INVALID_NULL As Long = 94
Public gWS As DAO.Workspace
Public gDB As DAO.Database
Public gConn As ADODB.Connection
Public goUtil As V2ECKeyBoard.clsUtil

Public Function SetDB(psDBPath As String, Optional psDir As String = "App.Path") As Boolean
    On Error GoTo EH
    
     SetDB = True
    If Not gDB Is Nothing Then
        'if its already set to this DB then we can exit
        If gDB.Name = Replace(psDBPath, "\\", "\") Then
            Exit Function 'BAIL
        Else
            Set gDB = Nothing
        End If
    End If
    
    'BGS 1.10.2001 Need to be sure the default directory is
    'Ap.path or will get strange errors when creating WorkSpace
    'especially when the cur dir is the floppy a drive.
    If psDir = "App.Path" Then
        psDir = App.Path
    End If
    ChDir psDir
    
    If gWS Is Nothing Then
        Set gWS = CreateWorkspace("", "admin", "", dbUseJet)
    End If

    Set gDB = gWS.OpenDatabase(psDBPath, False, False)
    
    Exit Function
EH:
    ShowError Err, "Public Function SetDB", , "Globals.bas"
End Function

Public Sub ErrorLog(psErrMess)
    Dim sMess As String
    Dim sErrMess As String
    Dim FileName As String
    Dim oForm As Form
    
    sErrMess = "<------------" & App.EXEName & " " & Now() & "------------>" & vbCrLf
    sErrMess = sErrMess & psErrMess
    
    If Not FindSetForm("frmProcessData", oForm) Then
        SaveSetting "V2WebControl", "Msg", "ErrorMess", sErrMess
    Else
        FileName = App.Path & "\" & Format(Now(), "YYMMDD") & "_Error.Log"
        sMess = GetFileData(FileName) & sErrMess
        SaveFileData FileName, sMess
        'svcMessageError = 109
        'svcEventError = 1
        oForm.NTSvcWebControl.LogEvent 109, 1, sErrMess
    End If
End Sub

Public Function OpenConnection(psConnString As String, Optional plOptions As Long = -1, Optional psErrorMess As String) As Boolean
    On Error GoTo EH
    Dim sUserID As String
    Dim sPassword As String
    Dim sMess As String
    
    OpenConnection = True
    
    If Not gConn Is Nothing Then
        If InStr(1, gConn.ConnectionString, psConnString, vbTextCompare) > 0 Then
            If gConn.State = ADODB.adStateOpen Then
                Exit Function
            End If
        End If
        Set gConn = Nothing
    End If
    
    sUserID = GetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    sPassword = GetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
    
    Set gConn = New ADODB.Connection
    gConn.ConnectionTimeout = 0 ' This will make the connection wait indefinately which is desirable for this application
    gConn.Open psConnString, sUserID, sPassword, plOptions
           
    Exit Function
EH:
    OpenConnection = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function OpenConnection" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
    psErrorMess = sMess
End Function

Public Function CloseConnection() As Boolean
    On Error GoTo EH
    Dim sMess As String
    
    CloseConnection = True
    
    If Not gConn Is Nothing Then
        gConn.Close
        Set gConn = Nothing
    End If
       
    Exit Function
EH:
    CloseConnection = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Function CloseConnection" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Function

Public Sub WCService(cmdType As Netcmd)
    On Error GoTo EH
    Dim sMess As String
    Dim bWC As Boolean
    Dim vcmdType As Variant
    Dim vServiceName As Variant
    
    bWC = GetSetting("V2WebControl", "WCService", "cmdActive", False)
    If bWC Then
        Exit Sub
    End If
    Select Case cmdType
        Case Netcmd.NetContinue
            vcmdType = "Continue"
        Case Netcmd.NetPause
            vcmdType = "Pause"
        Case Netcmd.NetStart
            vcmdType = "Start"
        Case Netcmd.NetStop
            vcmdType = "Stop"
        Case Netcmd.NetRestart
            vcmdType = "Restart"
        Case Else
            vcmdType = "Unknown (" & cmdType & ")"
            Err.Raise -999, , "Unknown Command"
            
    End Select
    
    'Flag Active command
    SaveSetting "V2WebControl", "WCService", "cmdActive", True
    If cmdType = NetRestart Then
        ReDim vcmdType(1 To 2)
        ReDim vServiceName(1 To 2)
        vcmdType(1) = "Stop"
        vServiceName(1) = "Webcontrol Service Vs 2.0"
        vcmdType(2) = "Start"
        vServiceName(2) = "Webcontrol Service Vs 2.0"
        ExecuteNetCommand vcmdType, vServiceName, False
    Else
        vServiceName = "Webcontrol Service Vs 2.0"
        ExecuteNetCommand vcmdType, vServiceName
    End If
    'UNFlag Active command
    SaveSetting "V2WebControl", "WCService", "cmdActive", False
    Exit Sub
EH:
    'UNFlag Active command
    SaveSetting "V2WebControl", "WCService", "cmdActive", False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Sub WCService" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    If IsArray(vcmdType) Then
        sMess = sMess & "Command Type = " & Join(vcmdType, ",") & vbCrLf
    Else
        sMess = sMess & "Command Type = " & vcmdType & vbCrLf
    End If
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Public Sub ExecuteNetCommand(pvcmdType As Variant, pvServiceName As Variant, Optional pbKill As Boolean = True)
    On Error GoTo EH
    Dim scmdType As String
    Dim sServiceName As String
    Dim sBatName As String
    Dim sPath As String
    Dim sCommand As String
    Dim sEnd As String
    Dim lCount As Long
    Dim lLBound As Long
    Dim lUBound As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'Check to see if there are multiple commands
    If IsArray(pvcmdType) Then
        lLBound = LBound(pvcmdType)
        lUBound = UBound(pvcmdType)
        
        For lCount = lLBound To lUBound
            'Build the names
            scmdType = pvcmdType(lCount)
            sServiceName = pvServiceName(lCount)
            sBatName = sServiceName & ".bat"
            If lCount = 1 Then
                sPath = App.Path & "\" & scmdType & sBatName
            End If
            If lCount = lUBound Then
                sEnd = vbNullString
            Else
                sEnd = vbCrLf
            End If
            sCommand = sCommand & "Net " & scmdType & " """ & sServiceName & """" & sEnd
        Next
    Else
        'Build the names
        scmdType = pvcmdType
        sServiceName = pvServiceName
        sBatName = sServiceName & ".bat"
        sPath = App.Path & "\" & scmdType & sBatName
        sCommand = "Net " & scmdType & " """ & sServiceName & """"
    End If
    
    'Save and Execute the Net Batch file
    SaveFileData sPath, sCommand
    Shell sPath, vbHide
    If pbKill Then
        'done with it can get rid of it
        DoEvents
        Sleep 500
        Kill sPath
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & "Public Sub ExecuteNetCommand"
End Sub
Public Sub SetUtilObject()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'Set Public Objects Here
    Set goUtil = New V2ECKeyBoard.clsUtil
    Set goUtil.goECKeyBoardList = New V2ECKeyBoard.clsLists
    goUtil.goECKeyBoardList.SetUtilObject goUtil
    Set goUtil.gARV = New V2ARViewer.clsARViewer
    goUtil.gARV.SetUtilObject goUtil
    goUtil.gsAppEXEName = App.EXEName 'Application Name
    goUtil.gsMainAppEXEName = App.EXEName
    goUtil.gsInstallDir = App.Path
    goUtil.gsCarPrefix = GetSetting(goUtil.gsAppEXEName, "COMPANY", "CAR_PREFIX", "V2ECcar")
    'Intialize the DB_PASSWORD Property
    goUtil.INIT_DB_PASSWORD = " r91223-i3q j12223-t1e z02223-n2d q81223-i4x"
    goUtil.SetUtilObject goUtil
    
    SaveSetting goUtil.gsAppEXEName, "COMPANY", "CAR_PREFIX", goUtil.gsCarPrefix
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & "Public Sub SetUtilObject"
End Sub

Public Function GetUID(Optional psUserName As String, Optional pbShowError As Boolean = False) As Long
    Dim sUserName As String
    Dim sProdDSN As String
    Dim sErrorMess As String
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sUserName = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    psUserName = sUserName
    If sUserName = vbNullString Then
        If pbShowError Then
            MsgBox "Must Set up Database Login under Settings!", vbExclamation + vbOKOnly, "User Name Not Found"
        End If
        Exit Function
    End If
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    CloseConnection
    If Not OpenConnection(sProdDSN, , sErrorMess) Then
        If pbShowError Then
            MsgBox sErrorMess, vbCritical + vbOKOnly, "Error"
        End If
        Exit Function
    End If
    Set RS = New ADODB.Recordset

    sSQL = "z_spsGetCompanyUsersInfo 1,0,null,null,null,null,null,null,null,null,null,'and username=''" & sUserName & "'' ' "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, gConn, adOpenStatic, adLockReadOnly

    If Not RS.EOF Then
        RS.MoveFirst
        GetUID = RS!USERSID
    Else
        GoTo CLEAN_UP
    End If
CLEAN_UP:
    RS.Close
    Set RS = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & "Public Function GetUID"
End Function

