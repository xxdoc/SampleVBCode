VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'BGS 1.28.2002 Need to see if Xactimate is even installed.
'***********LoadXact**************
'then check to see if it is already running, if it is then
'we need to set the main project window to focus, if it is not running then we need to launch
'it and check for Pasword window then get to the main project window.  At this point we can
' update Xactimate with the current project.  We are going to allow a single
'project or multiple projects.  Single project would be sent from ClaimMiant Screen.
'Multiple projects would be sent from List View Screen. (Note this class should Validate all
'fields before attempting to insert into Xactimate. If any field Fails the validation then we
'will not insert any project(s) until all projects and their fields are filled out correctly.
'We will prompt the User with a list of projects that are not filled out correctly.

Option Explicit

'EVENTS

'Update WebControl or Easy Claim Database
Public Event ProgressMessage(ByVal sMessage As String)

'These data types are what we need to populate into
'Xactimate Base project info
Public Type udtXactMain
    Name As String
    Street As String
    City As String
    State As String
    Zip As String
    Zip4 As String
    DateOfLoss As String
    DateReceived As String
    DateInspected As String
    DateEntered As String '(Should be todays Date )
End Type

Public Type udtXactLoss
    ClaimNumber As String
    PolicyNumber As String
    TypeOfLoss As String
    Deductible As String
    CatCode As String
End Type

Public Enum FixXactProj
    Status = 1
    Data = 2
    Name = 3
    Street
    StreetSort
    City
    State
    Zip
    Zip4
    DateOfLoss
    DateOfLossSort
    DateReceived
    DateReceivedSort
    DateInspected
    DateInspectedSort
    DateEntered
    DateEnteredSort
    ClaimNumber
    ClaimNumberSort
    PolicyNumber
    PolicyNumberSort
    TypeOfLoss
    Deductible
    DeductibleSort
    CatCode
    CatCodeSort
End Enum
Public Enum FixXactPic
    NeedsSentToXact = 1
    AlreadySentToXact
    BadData
    GoodData
    Skip
End Enum

Public Type udtXactProject
    Main As udtXactMain
    Loss As udtXactLoss
    IsDirty As Boolean 'True if validation failed and adjuster fixed data and it then passed validation
    ValidData As Boolean
    SentToXact As Boolean 'True if this item was sent to xactimate
    SkipThisProject As Boolean ' True if not want to send this one to Xactimate
End Type

Public Type udtLogOn
    UserID As String
    Password As String
End Type

Private moKB As clsKeyBoard
Private moReg As clsRegSetting
Private mcolXactProjects As Collection 'Holds multiple Projects to be added into Xactimate project list screen
Private mcolValidXProjects As Collection  ' Holds Projects that have been validated and or sent to Xactimate
Private moXProcess As clsProcess
Private msErrorInfo As String
Private mfrmTimer As frmTimer
Private mbCreatedOnce As Boolean 'Set this true the first time created
Private mbLookUpLoaded As Boolean 'True if we successfully loaded lookup info
Private mbXProjectsValidated As Boolean
Private mvaryStates As Variant ' lookup array of States derived from Xact Tables
Private mvaryTOL As Variant    ' lookup array of Types of loss derived from Xact Tables
Private mbSkipAll As Boolean 'True if exiting without sending to Xactimate
Private mlStartHwnd As Long
Private mbCancel As Boolean 'True if we ran clean
Private mbSendToExport As Boolean 'True if Sending to Export File ONLY!
Private mbGetFromExport As Boolean 'True If Sending to Xactimate from an Export File!
Private msExportFilePath As String ' The current Location of the Export File to save or Get
Private msExportData As String ' Build the Export String Data


Public Property Let ExportData(psData As String)
    msExportData = psData
End Property
Public Property Get ExportData() As String
    ExportData = msExportData
End Property

Public Property Let ExportFilePath(psPath As String)
    msExportFilePath = psPath
End Property
Public Property Get ExportFilePath() As String
    ExportFilePath = msExportFilePath
End Property

Public Property Let GetFromExport(pbFlag As Boolean)
    mbGetFromExport = pbFlag
End Property
Public Property Get GetFromExport() As Boolean
    GetFromExport = mbGetFromExport
End Property

Public Property Let SendToExport(pbFlag As Boolean)
    mbSendToExport = pbFlag
End Property
Public Property Get SendToExport() As Boolean
    SendToExport = mbSendToExport
End Property

Public Property Let CancelSendToXactimate(pbFlag As Boolean)
    mbCancel = pbFlag
End Property
Public Property Get CancelSendToXactimate() As Boolean
    CancelSendToXactimate = mbCancel
End Property

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsXact"
End Property

Public Property Let StartHwnd(plHwnd As Long)
    mlStartHwnd = plHwnd
End Property
Public Property Get SkipAll() As Boolean
    SkipAll = mbSkipAll
End Property
Public Property Get LookUpLoaded() As Boolean
    LookUpLoaded = mbLookUpLoaded
End Property
Public Property Get TOL() As Variant
    TOL = mvaryTOL
End Property
Public Property Get States() As Variant
    States = mvaryStates
End Property

Public Property Let XactProjects(pcolXactProjects As Collection)
    Set mcolXactProjects = pcolXactProjects
End Property
Public Property Set XactProjects(pcolXactProjects As Collection)
    Set mcolXactProjects = pcolXactProjects
End Property
Public Property Get XactProjects() As Collection
    Set XactProjects = mcolValidXProjects
End Property

Public Property Get ErrorInfo() As String
    ErrorInfo = msErrorInfo
End Property

Private Sub Class_Initialize()
    Set moKB = New clsKeyBoard
    Set moReg = New clsRegSetting
    Set mcolXactProjects = New Collection
    Set mcolValidXProjects = New Collection
    Set moXProcess = New clsProcess
    Set mfrmTimer = New frmTimer
    Load mfrmTimer
    PopulateLookUp
End Sub

Private Function LoadXact(Optional psLastWindow As String) As Boolean
    On Error GoTo EH
    Dim sXAppPath As String
    Dim sXCorePath As String 'Current directory must be here before we can create process for Xactimate
    Dim lSleep As Long
    Dim RetVal As Long
    Dim bFound As Boolean 'True if found login on create process
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMessage As String
    
    If mbCancel Then
        Exit Function
    End If
    
    sMessage = "Loading Xactimate... "
    RaiseEvent ProgressMessage(sMessage)
    
    'Check to see if the Login screen is already there
    Select Case psLastWindow
        Case vbNullString, "Xactimate - Login"
            If WindowFound("Xactimate - Login") Then
                moXProcess.Active = True
                'Wait for 5 seconds for user to logon
                'and try again
                Sleepy 5
                'This will keep flashing the Xactimate Login Screen
                'Until the user puts in their Login Info.  Also it will
                'build up the call stack every 5 seconds until Xactimate
                'logon times out.  This works nicely... A soon as Xactimate closes
                'or they log in, Every call on the stack will pop off.
                LoadXact = LoadXact("Xactimate - Login")  'Recursive Call
                Exit Function
            ElseIf WindowFound("Xactimate - Estimate") Then
                moXProcess.Active = True
                'If we find ourselves on the Estimate screen then
                'we need to save and close it out
                moKB.PressKeyVK keyAlt, True 'Hold alt key
                moKB.PressKey "t" 'Estimate
                moKB.PressKey "s" 'Save
                Sleepy 2
                moKB.PressKey "t" 'Estimate
                moKB.PressKey "x" 'Exit
                moKB.PressKeyVK keyAlt, , True 'Release the alt key
                LoadXact = LoadXact("Xactimate - Estimate")  'Recursive Call
                Exit Function
            ElseIf WindowFound("New Project") Then
                moXProcess.Active = True
                'If we are in the middle of adding a new project then
                'Cancel that
                moKB.PressKeyVK keyEscape
                LoadXact = LoadXact("New Project")  'Recursive Call
                Exit Function
            ElseIf WindowFound("Project") Then
                moXProcess.Active = True
                'If we are in the middle of opening a project then
                'Cancel that
                moKB.PressKeyVK keyEscape
                LoadXact = LoadXact("Project")  'Recursive Call
                Exit Function
            ElseIf WindowFound("Xactimate - Projects SmartList") Then
                moXProcess.Active = True
                LoadXact = True
                Exit Function
            ElseIf WindowFound("Xactimate - Control Center") Then
                moXProcess.Active = True
XACT_CONTROL_CENTER:
                Sleepy 5
                'If we are at the control centre then we need to open
                'the Projects smart list
                moKB.PressKeyVK keyAlt, True 'Hold Alt key
                moKB.PressKey "f" 'File
                moKB.PressKey "p" 'Project SmartList
                moKB.PressKeyVK keyAlt, , True 'Release the alt key
                LoadXact = LoadXact("Xactimate - Control Center")
                Exit Function
            End If
        Case "Xactimate - Estimate", "New Project", "Project"
            If WindowFound("Xactimate - Control Center") Then
                moKB.PressKeyVK keyAlt, True 'Hold alt key
                moKB.PressKey "f" 'File
                moKB.PressKey "p" 'Project SmartList
                moKB.PressKeyVK keyAlt, , True 'Release the Alt Key
                If WindowFound("Xactimate - Projects SmartList") Then
                    LoadXact = True
                    Exit Function
                End If
            End If
        Case "Xactimate - Control Center"
            If moXProcess.Active Then
                bFound = LookForWindow("Xactimate - Projects SmartList")
                If bFound Then
                    LoadXact = True
                    Exit Function
                End If
            Else
                If WindowFound("Xactimate - Projects SmartList") Then
                    LoadXact = True
                    Exit Function
                End If
            End If
            
    End Select
    
    If mbCancel Then
        Exit Function
    End If
    'If Xactimate is not already running then we need to ceate a process
    'for it. Once the process is created we will not come into this block again
    ' the above code will keep recursing every 5 seconds until either the Xactimate Logon Screen times out
    'or the user types in their info and the control center screen comes up.
    '
    If Not LoadXact And Not moXProcess.Active And Not mbCreatedOnce Then
        'See if Xactimate is installed
        'HKEY_LOCAL_MACHINE\SOFTWARE\XACTWARE\97
        'Find wherever the user installed their Xactimate Application
        'It should be here ... C:\Program Files\Xactware\
        'But it is possible the user decided not to use the default directory
        'so we must look it up everytime in their registry.
        sXAppPath = moReg.Query_Value(HKEY_LOCAL_MACHINE, "Software\Xactware\97", "Location")
        sXAppPath = sXAppPath & "\XM8\x.exe"
        sXCorePath = moReg.Query_Value(HKEY_LOCAL_MACHINE, "Software\Xactware\97", "Location")
        sXCorePath = sXCorePath & "\CORE"
        
        If Not FileExists(sXAppPath) Then
            msErrorInfo = msErrorInfo & "Xactimate Application path invalid (" & sXAppPath & ")" & vbCrLf
            Exit Function
        End If
        
        If Not FileExists(sXCorePath, True) Then
            msErrorInfo = msErrorInfo & "Xactimate Core path invalid (" & sXCorePath & ")" & vbCrLf
            Exit Function
        Else
            'Change current directory to this path or
            'create process for Xactimate x.exe will fail.
            If Not moXProcess.Active Then
            
                On Error Resume Next
                'MUST Change drive to same drive as Core path
                ChDrive left(sXCorePath, InStr(1, sXCorePath, ":"))
                'Must change the Current directory as well to core path
                ChDir sXCorePath
                 If Err.Number > 0 Then
                    Err.Clear
                End If
                On Error GoTo EH
                
                moXProcess.ProcCaption = "Xactimate - Control Center"
                RetVal = moXProcess.StartProgram(sXAppPath) 'Start Xactimate HERE !!
            Else
                RetVal = True
            End If
            
            If RetVal Then
'                Set mfrmTimer.Process = moXProcess
'                Load mfrmTimer
'                mfrmTimer.StartHwnd = mlStartHwnd
'                mfrmTimer.TimerProcess.Enabled = True
                mbCreatedOnce = True
                'Wait for the Logon screen to be found
                'give it at least 60 seconds.
                bFound = LookForWindow("Xactimate - Login")
                If bFound Then
                    'Wait as long as they keep the login screen up
                    Do Until Not WindowFound("Xactimate - Login")
                    Loop
                    bFound = LookForWindow("Xactimate - Control Center")
                    If bFound Then
                        GoTo XACT_CONTROL_CENTER:
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
        
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function LoadXact" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Sub Sleepy(plSeconds As Long)
    Dim lSleep As Long
    
    For lSleep = 1 To plSeconds * 10
       DoEvents
       Sleep 100
    Next
End Sub

Public Function ValidateXactProjects(Optional pbShowFixXactProjects As Boolean = True) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcolXactProjects Is Nothing Then
        Exit Function
    End If
    If Not ShowFixXact Then
        Exit Function
    End If
    mbXProjectsValidated = True
    ValidateXactProjects = mbXProjectsValidated
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    mbXProjectsValidated = False
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function ValidXProjects" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function ShowFixXact() As Boolean
    On Error GoTo EH
    Dim bProjectsValidated As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mbCancel Then
        Exit Function
    End If
    
    With frmFixXactProjects
        .XactProjects = mcolXactProjects
        .Xact = Me
        .ModalFlag = True
        Load frmFixXactProjects
        .Show vbModal 'MUST BE SHOWN MODAL !!!
        bProjectsValidated = .ProjectsValidated
        mbSkipAll = .SkipAll
        .CLEANUP
    End With
    Unload frmFixXactProjects
    Set frmFixXactProjects = Nothing
    
    ShowFixXact = bProjectsValidated
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function ShowFixXact" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not moKB Is Nothing Then
        moKB.CLEANUP
        Set moKB = Nothing
    End If
    If Not moReg Is Nothing Then
        Set moReg = Nothing
    End If
    If Not mcolXactProjects Is Nothing Then
        Set mcolXactProjects = Nothing
    End If
    If Not mcolValidXProjects Is Nothing Then
        Set mcolValidXProjects = Nothing
    End If
    If Not moXProcess Is Nothing Then
        moXProcess.CLEANUP
        Set moXProcess = Nothing
    End If
    
    If Not mfrmTimer Is Nothing Then
        mfrmTimer.CLEANUP
        Set mfrmTimer = Nothing
    End If
    mbSendToExport = False
    mbGetFromExport = False
    msExportFilePath = vbNullString
    msExportData = vbNullString

    
    CLEANUP = True
    mbCancel = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanUp" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function AddXProject(pvXProj As Variant, Optional pbValidated As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mbCancel Then
        Exit Function
    End If
    If Not pbValidated Then
        mcolXactProjects.Add pvXProj, pvXProj.Loss.ClaimNumber
    Else
        mcolValidXProjects.Add pvXProj, pvXProj.Loss.ClaimNumber
    End If
    AddXProject = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function AddXProject" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function SendToXact() As Boolean
    On Error GoTo EH
    Dim vXProject As Variant
    Dim XProject As udtXactProject
    Dim bFound As Boolean
    Dim sStartWindow As String
    Dim bCapsLock As Boolean
    Dim bNumLock As Boolean
    Dim bScrollLock As Boolean
    Dim sXactVS As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMessage As String
    Dim lCurProject As Long
    Dim lMaxProject As Long
    
    If mbCancel Then
        Exit Function
    End If
    
    If Not mbXProjectsValidated Then
        Exit Function
    End If
    'Don't load xactimate if only sending to Export File
    If Not mbSendToExport Then
        If Not LoadXact Then
            Exit Function
        End If
    End If
    
    Set mcolXactProjects = Nothing
    Set mcolXactProjects = mcolValidXProjects
    Set mcolValidXProjects = Nothing
    Set mcolValidXProjects = New Collection
    
    If Not mbSendToExport Then
        'Issue 183 9.8.2002 Send to xact puts info in all in Caps
        'or lowercase letters (opposite of what's in ec)
        moKB.GetLockStatus bCapsLock, bNumLock, bScrollLock
        
        If bCapsLock Then
            moKB.PressKeyVK keyCapsLock
        End If
        
        If Not bNumLock Then
            moKB.PressKeyVK keyNumLock
        End If
        
        If bScrollLock Then
            moKB.PressKeyVK keyScrollLock
        End If
    End If
    
    
    'Get list of projects to update message
    
    For Each vXProject In mcolXactProjects
        If mbCancel Then
            Exit Function
        End If
        XProject = vXProject
        If Not XProject.SentToXact Then
            If (XProject.ValidData Or mbSendToExport) And Not XProject.SkipThisProject Then
                lMaxProject = lMaxProject + 1
            End If
        End If
    Next
    
    If lMaxProject > 0 Then
        sMessage = "Sending " & lMaxProject & " Project(s)..."
        RaiseEvent ProgressMessage(sMessage)
    End If
    
    For Each vXProject In mcolXactProjects
        If mbCancel Then
            Exit Function
        End If
        XProject = vXProject
        If Not XProject.SentToXact Then
            If (XProject.ValidData Or mbSendToExport) And Not XProject.SkipThisProject Then
            
                lCurProject = lCurProject + 1
                sMessage = "Sending " & lCurProject & " of " & lMaxProject & " Project(s)..."
                RaiseEvent ProgressMessage(sMessage)
                
                If mbSendToExport Then
                    GoTo SEND_TO_EXPORT
                End If
                'Issue 224  9.23.2002 Send to Xactimate no work with Xactimate 2002
                'Need to code for Different versions here...
                sXactVS = GetSetting("ECS", "XACTIMATE", "Version", "2001")
                Select Case UCase(sXactVS)
                    Case "2001" '<------------------Xactimate 2001 ------------------>
                        '1
                        Sleepy 2
                        bFound = LookForWindow("Xactimate - Projects SmartList")
                        If Not bFound Then
                            GoTo SKIP_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                        moKB.PressKeyVK keyAlt, True                    'Hold Alt key
                        moKB.PressKey "a"                               'Add
                        moKB.PressKeyVK keyAlt, , True                  'Release Alt Key
                        
                        '2
                        Sleepy 1
                        bFound = LookForWindow("New Project")
                        If Not bFound Then
                            GoTo SKIP_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                        moKB.PressString XProject.Loss.ClaimNumber      'Project = Client Claim Number
                        PressTabKey 1                                   'Tab to Type Combo box
                        moKB.PressKey "e"                               'Type = Estimate
                        moKB.PressKeyVK keyReturn                       'Hit  Enter (Return Key)
                        
                        '3
                        Sleepy 1
                        bFound = LookForWindow("Xactimate - Estimate")
                        If Not bFound Then
                            GoTo NEXT_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                    
                        moKB.PressString XProject.Main.Name             'Name
                        PressTabKey 4                                   'Tab 4 spaces to Street
                        moKB.PressString XProject.Main.Street           'Street
                        PressTabKey 2                                   'Tab 2 spaces to City
                        moKB.PressString Replace(XProject.Main.City, Chr(160), " ") 'City
                        PressTabKey 1                                   'Tab 1 space to State
                        moKB.PressString XProject.Main.State             'State
                        PressTabKey 1                                   'Tab 1 space to ZIP
                        moKB.PressString XProject.Main.Zip & XProject.Main.Zip4
                        PressTabKey 9                                   'Tab 9 Spaces to Date of Loss
                        moKB.PressString XProject.Main.DateOfLoss       'Date Of Loss
                        PressTabKey 1                                   'Tab 1 space to Date Received
                        moKB.PressString XProject.Main.DateReceived     'Date Received
                        PressTabKey 1                                   'Tab 1 Space to Date Inspected
                        moKB.PressString XProject.Main.DateInspected    'Date Inspected
                        PressTabKey 1                                   'Tab 1 space to Date Entered
                        moKB.PressString XProject.Main.DateEntered      'Date Entered
                        PressTabKey 1                                   'Tab 1 space to Add button (Type of Loss)
                        moKB.PressKeyVK keyReturn                       'Hit Enter key while focus on Add Button (Type of Loss)
                        
                        '4
                        Sleepy 1
                        bFound = LookForWindow("Edit Loss Information")
                        If Not bFound Then
                            GoTo NEXT_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                        moKB.PressString XProject.Loss.ClaimNumber      'Claim Number (Client Claim Number)
                        PressTabKey 1                                   'Tab 1 space to Policy Number
                        moKB.PressString XProject.Loss.PolicyNumber     'Policy Number
                        PressTabKey 1                                   'Tab 1 space to Type of Loss
                        moKB.PressString XProject.Loss.TypeOfLoss       'Type of Loss
                        PressTabKey 1                                   'Tab 1 space to Deductible
                        moKB.PressString XProject.Loss.Deductible       'Deductible
                        PressTabKey 1                                   'Tab 1 space to Catastrophe Code
                        moKB.PressString XProject.Loss.CatCode          'Cat Code
                        PressTabKey 1                                   'Tab 1 space to OK button
                        moKB.PressKeyVK keyReturn                       'Hit Eneter key (Return Key)
                        
                        '5
                        Sleepy 1
                        bFound = LookForWindow("Xactimate - Estimate")
                        If Not bFound Then
                            GoTo NEXT_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                        moKB.PressKeyVK keyAlt, True                    'Hold Alt Key
                        moKB.PressKey "t"                               'Estimate Menu
                        moKB.PressKey "x"                               'Exit Menu
                        moKB.PressKeyVK keyAlt, , True                  'Release Alt Key
                        moKB.PressKeyVK keyReturn                       'Hit Enter (Yes button will have focus for Save Project prompt)
                    Case "2002", "2002(05-29-2003 272)", "2002(05-25-2004 276)" '<------------------Xactimate 2002 ------------------>
                        '1
                        Sleepy 2
                        bFound = LookForWindow("Xactimate - Projects SmartList")
                        If Not bFound Then
                            GoTo SKIP_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                        moKB.PressKeyVK keyAlt, True                    'Hold Alt key
                        moKB.PressKey "a"                               'Add
                        moKB.PressKeyVK keyAlt, , True                  'Release Alt Key
                        
                        '2
                        Sleepy 1
                        bFound = LookForWindow("New Project")
                        If Not bFound Then
                            GoTo SKIP_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                        moKB.PressString XProject.Loss.ClaimNumber      'Project = Client Claim Number
                        PressTabKey 1                                   'Tab to Type Combo box
                        moKB.PressKey "e"                               'Type = Estimate
                        moKB.PressKeyVK keyReturn                       'Hit  Enter (Return Key)
                        
                        '3
                        Sleepy 1
                        bFound = LookForWindow("Xactimate - Estimate")
                        If Not bFound Then
                            GoTo NEXT_PROJECT
                        End If
                        If mbCancel Then
                            Exit Function
                        End If
                        moKB.PressString XProject.Main.Name             'Name
                        PressTabKey 2                                   'Tab 2 spaces to Street (2001 Tab 4 spaces)
                        moKB.PressString XProject.Main.Street           'Street
                        PressTabKey 1                                   'Tab 1 spaces to City (2001 Tab 2 Spaces)
                        moKB.PressString Replace(XProject.Main.City, Chr(160), " ") 'City
                        PressTabKey 1                                   'Tab 1 space to State
                        moKB.PressString XProject.Main.State             'State
                        PressTabKey 1                                   'Tab 1 space to ZIP
                        moKB.PressString XProject.Main.Zip & XProject.Main.Zip4 'ZIP
                        '7.28.2003 Xactimate Updated Admin GUI
                        Select Case UCase(sXactVS)
                            Case "2002(05-29-2003 272)", "2002(05-25-2004 276)"
                                PressTabKey 12                          'Tab 12 Spaces to Date of Loss (2001 Tab 9 Spaces)
                            Case Else
                                PressTabKey 11                          'Tab 11 Spaces to Date of Loss (2001 Tab 9 Spaces)
                        End Select
                                                           
                        moKB.PressString XProject.Main.DateOfLoss       'Date Of Loss
                        PressTabKey 1                                   'Tab 1 space
                        '6.14.2004 Xactimate Updated Admin GUI
                        Select Case UCase(sXactVS)
                            Case "2002(05-25-2004 276)"
                                PressTabKey 1                           'Tab 1 space
                            Case Else
                        End Select
                        moKB.PressString XProject.Main.DateReceived     'Date Received
                        PressTabKey 1                                   'Tab 1 Space
                        '6.14.2004 Xactimate Updated Admin GUI
                        Select Case UCase(sXactVS)
                            Case "2002(05-25-2004 276)"
                                PressTabKey 1                           'Tab 1 space
                            Case Else
                        End Select
                        moKB.PressString XProject.Main.DateInspected    'Date Inspected
                        PressTabKey 1                                   'Tab 1 space
                        '6.14.2004 Xactimate Updated Admin GUI
                        Select Case UCase(sXactVS)
                            Case "2002(05-25-2004 276)"
                                PressTabKey 1                           'Tab 1 space
                            Case Else
                        End Select
                        moKB.PressString XProject.Main.DateEntered      'Date Entered
                        'There is no Add button for Type Of Loss in 2002.
                        'Just need to tab to the Type Of Loss
                        PressTabKey 1                                   'Tab 1 space
                        '6.14.2004 Xactimate Updated Admin GUI
                         Select Case UCase(sXactVS)
                            Case "2002(05-25-2004 276)"
                                PressTabKey 1                                   'Tab 1 space
                                moKB.PressKeyVK keyF5
                                bFound = LookForWindow("Type of Loss SmartList")
                                If Not bFound Then
                                    GoTo NEXT_PROJECT
                                End If
                                moKB.PressString XProject.Loss.TypeOfLoss       'Type of Loss
                                Sleep 500
                                moKB.PressKeyVK keyReturn
                                Sleep 500
                            Case Else
                                moKB.PressString XProject.Loss.TypeOfLoss       'Type of Loss
                                PressTabKey 1                                   'Tab 1 space to Claim Number
                        End Select    'Tab 1 space to Claim Number
                        moKB.PressString XProject.Loss.ClaimNumber      'Claim Number (Client Claim Number)
                        PressTabKey 1                                   'Tab 1 space to Policy Number
                        moKB.PressString XProject.Loss.PolicyNumber     'Policy Number
                        PressTabKey 1                                   'Tab 1 space to Deductible
                        moKB.PressString XProject.Loss.Deductible       'Deductible
                        PressTabKey 1                                   'Tab 1 space to Catastrophe Code
                        moKB.PressString XProject.Loss.CatCode          'Cat Code
                        '5
                        moKB.PressKeyVK keyAlt, True                    'Hold Alt Key
                        moKB.PressKey "t"                               'Estimate Menu
                        moKB.PressKey "x"                               'Exit Menu
                        moKB.PressKeyVK keyAlt, , True                  'Release Alt Key
                        moKB.PressKeyVK keyReturn                       'Hit Enter (Yes button will have focus for Save Project prompt)
                    Case Else
                        GoTo SKIP_PROJECT
                End Select
SEND_TO_EXPORT:
                'BGS Set flag that was sent to Xact
                XProject.SentToXact = True

                'Need to Update DB as we go along...
                If Not goUtil Is Nothing And Not mbGetFromExport Then
                    If Not goUtil.goCurCarList Is Nothing Then
                        If Not goUtil.goCurCarList.FlagSentToXactimate(XProject) Then
                            'set the Cancel flag if we fail to update the
                            mbCancel = True
                        Else
                            If mbSendToExport Then
                                'Reset this flag
                                XProject.SentToXact = False
                                If Not BuildXactExport(XProject) Then
                                    mbCancel = True
                                End If
                                XProject.SentToXact = True
                            End If
                        End If
                    End If
                Else
                    If mbSendToExport Then
                        XProject.SentToXact = False
                        If Not BuildXactExport(XProject) Then
                            mbCancel = True
                        End If
                        XProject.SentToXact = True
                    End If
                End If
            Else
                GoTo SKIP_PROJECT
            End If
        End If
        
NEXT_PROJECT:
        AddXProject XProject, True
SKIP_PROJECT:
    If (mbGetFromExport And Not mbSendToExport) Then
        If Not BuildXactExport(XProject) Then
            mbCancel = True
        End If
    End If
    Next
    
    If mbSendToExport Or mbGetFromExport Then
        If msExportData <> vbNullString Then
            If FileExists(msExportFilePath) Then
                DeleteFile msExportFilePath
            End If
            SaveFileData msExportFilePath, msExportData
            msExportData = vbNullString
        End If
    End If
    
    If lMaxProject > 0 Then
        sMessage = "Send Complete " & lCurProject & " of " & lMaxProject & " Project(s)..."
        RaiseEvent ProgressMessage(sMessage)
    End If
    
    SendToXact = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function SendToXact" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function BuildXactExport(pvXProj As Variant) As Boolean
    On Error GoTo EH
    Dim sExportData As String
    Dim XProject As udtXactProject
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    XProject = pvXProj
    
    With XProject
        sExportData = .IsDirty & F_DELIM
        With .Loss
            sExportData = sExportData & .CatCode & F_DELIM
            sExportData = sExportData & .ClaimNumber & F_DELIM
            sExportData = sExportData & .Deductible & F_DELIM
            sExportData = sExportData & .PolicyNumber & F_DELIM
            sExportData = sExportData & .TypeOfLoss & F_DELIM
        End With
        With .Main
            sExportData = sExportData & .City & F_DELIM
            sExportData = sExportData & .DateEntered & F_DELIM
            sExportData = sExportData & .DateInspected & F_DELIM
            sExportData = sExportData & .DateOfLoss & F_DELIM
            sExportData = sExportData & .DateReceived & F_DELIM
            sExportData = sExportData & .Name & F_DELIM
            sExportData = sExportData & .State & F_DELIM
            sExportData = sExportData & .Street & F_DELIM
            sExportData = sExportData & .Zip & F_DELIM
            sExportData = sExportData & .Zip4 & F_DELIM
        End With
        sExportData = sExportData & .SentToXact & F_DELIM
        sExportData = sExportData & .SkipThisProject & F_DELIM
        sExportData = sExportData & .ValidData & F_DELIM
    End With
    sExportData = sExportData & vbCrLf
    
    'Add on to the existing Export Data
    msExportData = msExportData & sExportData
    
    BuildXactExport = True
    
    'Build The Export File
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    BuildXactExport = False
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function BuildXactExport" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function PopulateFromXactExport() As Boolean
    On Error GoTo EH
    Dim saryExport() As String
    Dim saryXProj() As String
    Dim sExportData As String
    Dim XProject As udtXactProject
    Dim lExportCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Clear Any Exisiting Xact projects
    Set frmFixXactProjects.XactProjects = Nothing
    Set mcolXactProjects = Nothing
    Set mcolXactProjects = New Collection
    
        
    sExportData = GetFileData(msExportFilePath)
    saryExport = Split(sExportData, vbCrLf, , vbBinaryCompare)
    
    For lExportCount = LBound(saryExport) To UBound(saryExport)
        If saryExport(lExportCount) <> vbNullString Then
            saryXProj = Split(saryExport(lExportCount), F_DELIM, , vbBinaryCompare)
            With XProject
                .IsDirty = saryXProj(0)
                With .Loss
                    .CatCode = saryXProj(1)
                    .ClaimNumber = saryXProj(2)
                    .Deductible = saryXProj(3)
                    .PolicyNumber = saryXProj(4)
                    .TypeOfLoss = saryXProj(5)
                End With
                With .Main
                    .City = saryXProj(6)
                    .DateEntered = saryXProj(7)
                    .DateInspected = saryXProj(8)
                    .DateOfLoss = saryXProj(9)
                    .DateReceived = saryXProj(10)
                    .Name = saryXProj(11)
                    .State = saryXProj(12)
                    .Street = saryXProj(13)
                    .Zip = saryXProj(14)
                    .Zip4 = saryXProj(15)
                End With
                .SentToXact = saryXProj(15)
                .SkipThisProject = saryXProj(16)
                .ValidData = saryXProj(17)
            End With
            AddXProject XProject
        End If
    Next
    
    Set frmFixXactProjects.XactProjects = mcolXactProjects
    
    PopulateFromXactExport = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    PopulateFromXactExport = False
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function PopulateFromXactExport" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function LookForWindow(psWindowCaption As String, Optional plTimeOutSeconds As Long = 60) As Boolean
    On Error GoTo EH
    Dim bFound As Boolean
    Dim lSleep As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mbCancel Then
        Exit Function
    End If
    bFound = True
    Do Until WindowFound(psWindowCaption)
        lSleep = lSleep + 1
        If lSleep > plTimeOutSeconds Then
            bFound = False
            Exit Do
        End If
        If mbCancel Then
            Exit Function
        End If
        If Not moXProcess.Active Then
            bFound = False
            Exit Do
        End If
    Loop
    
    LookForWindow = bFound
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function LookForWindow" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Sub PressTabKey(plCount As Long)
    Dim lCount As Long
    If mbCancel Then
        Exit Sub
    End If
    For lCount = 1 To plCount
        moKB.PressKeyVK keyTab
    Next
End Sub

Friend Function ValidateXProject(pvXProj As Variant, Optional pitmX As listItem) As Boolean
    On Error GoTo EH
    Dim XProj As udtXactProject
    Dim itmX As listItem
    Dim bitmX As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mbCancel Then
        Exit Function
    End If
    
    If Not IsEmpty(pitmX) Then
        Set itmX = pitmX
        bitmX = True
    End If
    XProj = pvXProj
    XProj.ValidData = True
    'Validate the Main Data

    With XProj.Main
        If Not DateValid(.DateEntered) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.DateEntered - 1) = .DateEntered
                itmX.SubItems(FixXactProj.DateEnteredSort - 1) = Format(.DateEntered, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateEntered - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.DateEntered - 1) = .DateEntered
                itmX.SubItems(FixXactProj.DateEnteredSort - 1) = Format(.DateEntered, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateEntered - 1).ReportIcon = Empty
            End If
        End If
        
        If Not DateValid(.DateInspected) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.DateInspected - 1) = .DateInspected
                itmX.SubItems(FixXactProj.DateInspectedSort - 1) = Format(.DateInspected, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateInspected - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.DateInspected - 1) = .DateInspected
                itmX.SubItems(FixXactProj.DateInspectedSort - 1) = Format(.DateInspected, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateInspected - 1).ReportIcon = Empty
            End If
        End If
        
        If Not DateValid(.DateOfLoss) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.DateOfLoss - 1) = .DateOfLoss
                itmX.SubItems(FixXactProj.DateOfLossSort - 1) = Format(.DateOfLoss, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateOfLoss - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.DateOfLoss - 1) = .DateOfLoss
                itmX.SubItems(FixXactProj.DateOfLossSort - 1) = Format(.DateOfLoss, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateOfLoss - 1).ReportIcon = Empty
            End If
        End If
        
        If Not DateValid(.DateReceived) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.DateReceived - 1) = .DateReceived
                itmX.SubItems(FixXactProj.DateReceivedSort - 1) = Format(.DateReceived, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateReceived - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.DateReceived - 1) = .DateReceived
                itmX.SubItems(FixXactProj.DateReceivedSort - 1) = Format(.DateReceived, "YYYY/MM/DD")
                itmX.ListSubItems(FixXactProj.DateReceived - 1).ReportIcon = Empty
            End If
        End If
        
'        .Name      'check for null string
        If Trim(.Name) = vbNullString Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.Name - 1) = .Name
                itmX.ListSubItems(FixXactProj.Name - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.Name - 1) = .Name
                itmX.ListSubItems(FixXactProj.Name - 1).ReportIcon = Empty
            End If
        End If
'        .Street    'check for null string
        If Trim(.Street) = vbNullString Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.Street - 1) = .Street
                itmX.SubItems(FixXactProj.StreetSort - 1) = goUtil.utNumInTextSortFormat(.Street)
                itmX.ListSubItems(FixXactProj.Street - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.Street - 1) = .Street
                itmX.SubItems(FixXactProj.StreetSort - 1) = goUtil.utNumInTextSortFormat(.Street)
                itmX.ListSubItems(FixXactProj.Street - 1).ReportIcon = Empty
            End If
        End If
'        .City      'check for null string
        If Trim(.City) = vbNullString Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.City - 1) = .City
                itmX.ListSubItems(FixXactProj.City - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.City - 1) = .City
                itmX.ListSubItems(FixXactProj.City - 1).ReportIcon = Empty
            End If
        End If

        If Not StateValid(.State) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.State - 1) = .State
                itmX.ListSubItems(FixXactProj.State - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.State - 1) = .State
                itmX.ListSubItems(FixXactProj.State - 1).ReportIcon = Empty
            End If
        End If
        
        If Not Zip5Valid(.Zip) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.Zip - 1) = .Zip
                itmX.ListSubItems(FixXactProj.Zip - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.Zip - 1) = .Zip
                itmX.ListSubItems(FixXactProj.Zip - 1).ReportIcon = Empty
            End If
        End If
        
        If Not Zip4Valid(.Zip4) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.Zip4 - 1) = .Zip4
                itmX.ListSubItems(FixXactProj.Zip4 - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.Zip4 - 1) = .Zip4
                itmX.ListSubItems(FixXactProj.Zip4 - 1).ReportIcon = Empty
            End If
        End If
    End With
    
    'Validate Loss Data
    With XProj.Loss
        If Trim(.ClaimNumber) = vbNullString Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.ClaimNumber - 1) = .ClaimNumber
                itmX.SubItems(FixXactProj.ClaimNumberSort - 1) = goUtil.utNumInTextSortFormat(.ClaimNumber)
                itmX.ListSubItems(FixXactProj.ClaimNumber - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.ClaimNumber - 1) = .ClaimNumber
                itmX.SubItems(FixXactProj.ClaimNumberSort - 1) = goUtil.utNumInTextSortFormat(.ClaimNumber)
                itmX.ListSubItems(FixXactProj.ClaimNumber - 1).ReportIcon = Empty
            End If
        End If
        
        If Trim(.PolicyNumber) = vbNullString Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.PolicyNumber - 1) = .PolicyNumber
                itmX.SubItems(FixXactProj.PolicyNumberSort - 1) = goUtil.utNumInTextSortFormat(.PolicyNumber)
                itmX.ListSubItems(FixXactProj.PolicyNumber - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.PolicyNumber - 1) = .PolicyNumber
                itmX.SubItems(FixXactProj.PolicyNumberSort - 1) = goUtil.utNumInTextSortFormat(.PolicyNumber)
                itmX.ListSubItems(FixXactProj.PolicyNumber - 1).ReportIcon = Empty
            End If
        End If
        
        If Not TolValid(.TypeOfLoss) Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.TypeOfLoss - 1) = .TypeOfLoss
                itmX.ListSubItems(FixXactProj.TypeOfLoss - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.TypeOfLoss - 1) = .TypeOfLoss
                itmX.ListSubItems(FixXactProj.TypeOfLoss - 1).ReportIcon = Empty
            End If
        End If
        
        If Trim(.Deductible) = vbNullString Or IsNumeric(.Deductible) = False Then
            XProj.ValidData = False
            If bitmX Then
                itmX.SubItems(FixXactProj.Deductible - 1) = .Deductible
                itmX.SubItems(FixXactProj.DeductibleSort - 1) = goUtil.utNumInTextSortFormat(.Deductible)
                itmX.ListSubItems(FixXactProj.Deductible - 1).ReportIcon = FixXactPic.BadData
            End If
        Else
            If bitmX Then
                itmX.SubItems(FixXactProj.Deductible - 1) = .Deductible
                itmX.SubItems(FixXactProj.DeductibleSort - 1) = goUtil.utNumInTextSortFormat(.Deductible)
                itmX.ListSubItems(FixXactProj.Deductible - 1).ReportIcon = Empty
            End If
        End If
'        .CatCode       'No validation required
        If bitmX Then
            itmX.SubItems(FixXactProj.CatCode - 1) = .CatCode
            itmX.SubItems(FixXactProj.CatCodeSort - 1) = goUtil.utNumInTextSortFormat(.CatCode)
        End If
    End With
    
    If XProj.ValidData Then
        ValidateXProject = True
        If bitmX Then
            itmX.SubItems(FixXactProj.Data - 1) = vbNullString
            itmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.GoodData
        End If
    Else
        If bitmX Then
            itmX.SubItems(FixXactProj.Data - 1) = vbNullString
            itmX.ListSubItems(FixXactProj.Data - 1).ReportIcon = FixXactPic.BadData
        End If
    End If
    
    ValidateXProject = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Friend Function ValidateXProject" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function DateValid(psDate As String) As Boolean
    On Error GoTo EH
    Dim dDate As Date
    Dim dNow As Date
    
    If mbCancel Then
        Exit Function
    End If
    
    If IsDate(psDate) Then
        dDate = Format(psDate, "MM/DD/YYYY")
        dNow = Format(Now, "MM/DD/YYYY")
    Else
        If psDate = vbNullString Then
            DateValid = True
        ElseIf psDate = "//" Then
            psDate = vbNullString
            DateValid = True
        End If
        Exit Function
    End If
    
    If dDate <= dNow Then
        'Date must be greater that 1900
        If dDate > CDate("1/1/1900") Then
            DateValid = True
            psDate = Format(dDate, "MM/DD/YYYY")
        End If
    End If
    
    Exit Function
EH:
    DateValid = False
End Function
 Private Function StateValid(psState As String) As Boolean
    On Error GoTo EH
    Dim lCount As Long
    
    If mbCancel Then
        Exit Function
    End If
    
    If Trim(psState) = vbNullString Then
        StateValid = False
        Exit Function
    End If
    If IsArray(mvaryStates) Then
        For lCount = LBound(mvaryStates, 1) To UBound(mvaryStates, 1)
            If StrComp(CStr(mvaryStates(lCount)), psState, vbTextCompare) = 0 Then
                StateValid = True
                Exit For
            End If
        Next
    End If
    
    Exit Function
EH:
    StateValid = False
 End Function

Private Function TolValid(psTOL As String) As Boolean
    On Error GoTo EH
    Dim lCount As Long
    
    If mbCancel Then
        Exit Function
    End If
    
    If IsArray(mvaryTOL) Then
        For lCount = LBound(mvaryTOL, 1) To UBound(mvaryTOL, 1)
            If StrComp(CStr(mvaryTOL(lCount)), psTOL, vbTextCompare) = 0 Then
                TolValid = True
                Exit For
            End If
        Next
    End If
    
    Exit Function
EH:
    TolValid = False
End Function

Friend Function ZipValid(psZip As String, Optional psZip5, Optional psZip4) As Boolean
    On Error GoTo EH
    Dim sZip As String
    Dim sZip5 As String
    Dim sZip4 As String
    If mbCancel Then
        Exit Function
    End If
    Select Case Trim(psZip)
        Case vbNullString, "-"
            ZipValid = True
            psZip = "00000-0000"
            psZip5 = "00000"
            psZip4 = "0000"
        Case Is > vbNullString
            If InStr(1, psZip, "-") > 0 Then
                sZip = Trim(psZip)
                sZip = Replace(sZip, Chr(32), vbNullString)
                sZip = Replace(sZip, Chr(160), vbNullString)
                If Len(sZip) = 10 Then
                    If IsNumeric(left(sZip, 1)) Then
                        sZip5 = left(sZip, InStr(1, sZip, "-") - 1)
                        sZip5 = Format(sZip5, "00000")
                        If Not IsEmpty(psZip5) Then
                            psZip5 = sZip5
                        End If
                    End If
                    If IsNumeric(Right(sZip, 1)) Then
                        sZip4 = Mid(sZip, InStr(1, sZip, "-") + 1)
                        sZip4 = Format(sZip4, "0000")
                        If Not IsEmpty(psZip4) Then
                            psZip4 = sZip4
                        End If
                    End If
                    
                    If IsNumeric(sZip5) And IsNumeric(sZip4) Then
                        ZipValid = True
                        psZip = sZip5 & "-" & sZip4
                    End If
                End If
            Else
                sZip = Trim(psZip)
                sZip = Trim(psZip)
                sZip = Replace(sZip, Chr(32), vbNullString)
                sZip = Replace(sZip, Chr(160), vbNullString)
                If Len(sZip) = 5 Then
                    If IsNumeric(sZip) Then
                        ZipValid = True
                        psZip = sZip
                        If Not IsEmpty(psZip5) Then
                            psZip5 = psZip
                        End If
                        If Not IsEmpty(psZip4) Then
                            psZip4 = "0000"
                        End If
                    End If
                End If
            End If
    End Select
    
    Exit Function
EH:
    ZipValid = False
End Function

Friend Function Zip5Valid(psZip5 As String) As Boolean
    On Error GoTo EH
    Dim sZip5 As String
    If mbCancel Then
        Exit Function
    End If
    Select Case psZip5
        Case vbNullString
            Zip5Valid = True
            psZip5 = "00000"
        Case Is > vbNullString
            sZip5 = Trim(psZip5)
            If IsNumeric(sZip5) Then
                sZip5 = Format(sZip5, "00000")
                psZip5 = sZip5
                Zip5Valid = True
            End If
    End Select
    
    Exit Function
EH:
    Zip5Valid = False
End Function

Friend Function Zip4Valid(psZip4 As String) As Boolean
    On Error GoTo EH
    Dim sZip4 As String
    If mbCancel Then
        Exit Function
    End If
    Select Case psZip4
        Case vbNullString
            Zip4Valid = True
            psZip4 = "0000"
        Case Is > vbNullString
            sZip4 = Trim(psZip4)
            If IsNumeric(sZip4) Then
                sZip4 = Format(sZip4, "0000")
                psZip4 = sZip4
                Zip4Valid = True
            End If
    End Select
    
    Exit Function
EH:
    Zip4Valid = False
End Function

Private Sub PopulateLookUp(Optional pbHideError As Boolean)
    On Error GoTo EH
    Dim sXAppPath As String
    Dim sXCorePath As String
    Dim sStatesPath As String
    Dim sTOLPath As String
    Dim sTemp As String
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mbCancel Then
        Exit Sub
    End If
    mbLookUpLoaded = True
    
    sXAppPath = moReg.Query_Value(HKEY_LOCAL_MACHINE, "Software\Xactware\97", "Location")
    sXAppPath = sXAppPath & "\XM8"
    sXCorePath = moReg.Query_Value(HKEY_LOCAL_MACHINE, "Software\Xactware\97", "Location")
    sXCorePath = sXCorePath & "\CORE"
    
    sStatesPath = sXCorePath & "\data\STATES.DBX"
    sTOLPath = sXAppPath & "\Profiles\Carrier\SYSTEM\TOL.DBX"
    
    'Get the States from Xact
    If FileExists(sStatesPath) Then
        On Error Resume Next
        FileCopy sStatesPath, App.Path & "\STATES.DBX"
        sStatesPath = App.Path & "\STATES.DBX"
        SetAttr sStatesPath, vbNormal
        
        If Not FileExists(sStatesPath) Then
            mbLookUpLoaded = False
            If Not pbHideError Then
                MsgBox "Please close Xactimate and try again.", vbExclamation, "CLOSE XACTIMATE"
            End If
            Exit Sub
        End If
        On Error GoTo EH
        '196  Error #5 ECKEYBOARD  6.14.2002
        'Xactimate has States table messed up
        sTemp = GetFileData(sStatesPath)
        If sTemp = vbNullString Then
            mbLookUpLoaded = False
            sTemp = "Invalid Xactimate STATES.DBX file." & vbCrLf & vbCrLf
            sTemp = sTemp & "Please contact Xactimate support to" & vbCrLf
            sTemp = sTemp & "fix this problem, if you are unable to" & vbCrLf
            sTemp = sTemp & "reconfigure (US/Canada) State information" & vbCrLf
            sTemp = sTemp & "on your own."
            If Not pbHideError Then
                MsgBox sTemp, vbExclamation, "Xactimate Configuration Error"
            End If
            Exit Sub
        End If
        mvaryStates = Split(sTemp, Chr(32))
        If IsArray(mvaryStates) Then
            sTemp = vbNullString
            'Skip first element Header info
            For lCount = LBound(mvaryStates, 1) + 1 To UBound(mvaryStates, 1)
                If Len(mvaryStates(lCount)) > 3 Then
                    If Val(mvaryStates(lCount)) = 0 Then
                        'Add it to stemp
                        sTemp = sTemp & left(mvaryStates(lCount), 2) & vbCrLf
                        'skip past the next several elements
                        lCount = lCount + 5
                    End If
                End If
            Next
            'Get rid of tail vbcrlf
            sTemp = left(sTemp, InStrRev(sTemp, vbCrLf) - 1)
            'Finally split stemp into States array
            mvaryStates = Split(sTemp, vbCrLf)
            'sort the states array
            BubbleSort mvaryStates
        End If
    End If
    
    'Get the Type Of Loss from Xact
    If FileExists(sTOLPath) Then
        On Error Resume Next
        FileCopy sTOLPath, App.Path & "\TOL.DBX"
        sTOLPath = App.Path & "\TOL.DBX"
        SetAttr sTOLPath, vbNormal
        
        If Not FileExists(sTOLPath) Then
            mbLookUpLoaded = False
            If Not pbHideError Then
                MsgBox "Please close Xactimate and try again.", vbExclamation, "CLOSE XACTIMATE"
            End If
            Exit Sub
        End If
        On Error GoTo EH
        sTemp = GetFileData(sTOLPath)
        '196  Error #5 ECKEYBOARD  6.14.2002
        'Xactimate has TOL Type Of Loss table messed up
        If sTemp = vbNullString Then
            mbLookUpLoaded = False
            sTemp = "Invalid Xactimate TOL.DBX file." & vbCrLf & vbCrLf
            sTemp = sTemp & "Please contact Xactimate support to" & vbCrLf
            sTemp = sTemp & "fix this problem, if you are unable to" & vbCrLf
            sTemp = sTemp & "reconfigure (TOL) Type of Loss information" & vbCrLf
            sTemp = sTemp & "on your own."
            If Not pbHideError Then
                MsgBox sTemp, vbExclamation, "Xactimate Configuration Error"
            End If
            Exit Sub
        End If
        mvaryTOL = Split(sTemp, Chr(32))
        If IsArray(mvaryTOL) Then
            sTemp = vbNullString
            'Skip first element Header info
            For lCount = LBound(mvaryTOL, 1) + 1 To UBound(mvaryTOL, 1)
                If Len(mvaryTOL(lCount)) > 3 Then
                    'Only get the Ucase tol
                    mvaryTOL(lCount) = UCaseOnly(CStr(mvaryTOL(lCount)))
                    If Len(mvaryTOL(lCount)) > 0 Then
                        If InStr(1, mvaryTOL(lCount), "*") = 0 Then
                            'Add it to stemp
                            sTemp = sTemp & mvaryTOL(lCount) & vbCrLf
                        End If
                        'skip past the next several elements
                        lCount = lCount + 10
                    End If
                End If
            Next
            'Get rid of tail vbcrlf
            sTemp = left(sTemp, InStrRev(sTemp, vbCrLf) - 1)
            'Finally split stemp into TOL array
            mvaryTOL = Split(sTemp, vbCrLf)
            'sort the TOL array
            BubbleSort mvaryTOL
        End If
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    mbLookUpLoaded = False
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub PopulateLookUp" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Private Function UCaseOnly(psText As String) As String
    On Error GoTo EH
    Dim lCount As Long
    Dim lPos As Long
    Dim sTemp As String
    Dim sChar As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mbCancel Then
        Exit Function
    End If
    
    If Len(psText) > 1 Then
        For lCount = 1 To Len(psText)
            sChar = Mid(psText, lCount, 1)
            If sChar = UCase(sChar) Then
                sTemp = sTemp & sChar
            Else
                If Len(sTemp) > 1 Then
                    sTemp = left(sTemp, Len(sTemp) - 1)
                    Exit For
                End If
            End If
        Next
        UCaseOnly = sTemp
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function UcaseOnly" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function SameDate(psDate1 As String, psDate2 As String) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    SameDate = True
    
    If IsDate(psDate1) And IsDate(psDate2) Then
        If CDate(psDate1) <> CDate(psDate2) Then
            SameDate = False
        End If
    Else
        If psDate1 <> psDate2 Then
            SameDate = False
        End If
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function SameDate" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Friend Sub CheckIsDirtyAndAdd(pCheckXactProjects As Collection)
    On Error GoTo EH
    If mbCancel Then
        Exit Sub
    End If
    Dim CheckXProj As udtXactProject
    Dim XProj As udtXactProject
    Dim lCount As Long
    Dim bIsDirty As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not mcolXactProjects Is Nothing Then
        Set mcolValidXProjects = Nothing
        Set mcolValidXProjects = New Collection
    Else
        Exit Sub
    End If
    
    For lCount = 1 To mcolXactProjects.Count
        bIsDirty = False
        XProj = mcolXactProjects.Item(lCount)
        CheckXProj = pCheckXactProjects.Item(lCount)
        With CheckXProj
            With .Main
            
                If .City <> XProj.Main.City Then
                    bIsDirty = True
                End If
                
                If Not SameDate(.DateEntered, XProj.Main.DateEntered) Then
                    bIsDirty = True
                End If
                
                If Not SameDate(.DateInspected, XProj.Main.DateInspected) Then
                    bIsDirty = True
                End If
                
                If Not SameDate(.DateOfLoss, XProj.Main.DateOfLoss) Then
                    bIsDirty = True
                End If
                
                If Not SameDate(.DateReceived, XProj.Main.DateReceived) Then
                    bIsDirty = True
                End If
                
                If .Name <> XProj.Main.Name Then
                    bIsDirty = True
                End If
                
                If .State <> XProj.Main.State Then
                    bIsDirty = True
                End If
                
                If .Street <> XProj.Main.Street Then
                    bIsDirty = True
                End If
                
                If .Zip <> XProj.Main.Zip Then
                    bIsDirty = True
                End If
                
                If .Zip4 <> XProj.Main.Zip4 Then
                    bIsDirty = True
                End If
                
            End With
            With .Loss
                
                If .CatCode <> XProj.Loss.CatCode Then
                    bIsDirty = True
                End If
                
                If .ClaimNumber <> XProj.Loss.ClaimNumber Then
                    bIsDirty = True
                End If
                
                If .Deductible <> XProj.Loss.Deductible Then
                    bIsDirty = True
                End If
                
                If .PolicyNumber <> XProj.Loss.PolicyNumber Then
                    bIsDirty = True
                End If
                
                If .TypeOfLoss <> XProj.Loss.TypeOfLoss Then
                    bIsDirty = True
                End If
                
            End With
            
        End With
        
        CheckXProj.IsDirty = bIsDirty
        AddXProject CheckXProj, True
        
    Next
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private sub CheckIsDirtyAndAdd" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
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
