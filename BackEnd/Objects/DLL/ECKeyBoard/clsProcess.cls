VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type


Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private mProc As PROCESS_INFORMATION
Private mStart As STARTUPINFO
Private mbActive As Boolean
Private msProcCaption As String

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsProcess"
End Property

Public Property Let ProcCaption(psName As String)
    msProcCaption = psName
End Property
Public Property Get ProcCaption() As String
    ProcCaption = msProcCaption
End Property

Public Property Get Active() As Boolean
    Active = mbActive
End Property
Public Property Let Active(pbFlag As Boolean)
    mbActive = pbFlag
End Property
Public Property Get ProcessInfo() As PROCESS_INFORMATION
    ProcessInfo = mProc
End Property

Public Function StartProgram(ByVal ProgramName$, Optional pbUseSHell As Boolean = True) As Long
    On Error GoTo EH
    Dim RetVal As Long
    Dim dShellVal As Double
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not pbUseSHell Then
        
        RetVal = CreateProcessA(0&, ProgramName$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, mStart, mProc)
        StartProgram = RetVal
    Else
        dShellVal = Shell(ProgramName$, vbNormalFocus)
        StartProgram = CLng(dShellVal)
    End If
    
    If StartProgram Then
        mbActive = True
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function StartProgram" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

'Call this if you need to wait for the process this class created to
'end Once you call it it will loop until it sees that the process
'is over. then it will set the Active flag to False. while it is looping
'only Event code may execute.

Public Function WaitForProgramToEnd() As Boolean
    On Error GoTo EH
    Dim RetVal As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Only check for a finite amount of time (100 mili) since we
    'are dealing with single thread(In apartmentthread model) vb6
    'other wise only event driven code will excecute on the
    'doevents during this process
    If mbActive Then
        RetVal = 1
        RetVal = WaitForSingleObject(mProc.hProcess, 100)
        If RetVal = 0 Then
            mbActive = False
            WaitForProgramToEnd = True 'set to true if the program ended
        End If
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function WaitForProgramToEnd" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
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
    
    mbActive = False
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanUp" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

'Don't use TerminateProcess It should only be used as last resort
'Like ctrl alt delete. Let user terminate any processes that are
'created through this class by clicking on file exit or the X button.
'One could use ExitProcess but the code involved
'is not not worth what we get out of it and there are side effects if you do not
'compile this class in native code and try to use ExitProcess. GPF will untimately occur
'if using ExitProcess while debugging and Memory leaks occur if using TerminateProcess.
'We will Be better off using API made for whatever process you are creating to exit that application
'if you just gotta kill that process in code.
'Public Sub KillProgram()
'    Dim RetVal As Long
'    RetVal = TerminateProcess(mProc.hProcess, 1)
'End Sub
