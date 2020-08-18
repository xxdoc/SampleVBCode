VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Public Sub PressKey
'Public Sub PressString
'Public Sub PressKeyVK
'Public Sub GetLockStatus
'Source code http://www.freevbcode.com/ShowCode.Asp?ID=340
'Provided by :
'Name: scott Seligman
'E -mail: scott@ scottandmichelle.net
'URL: http://www.scottandmichelle.net/
Option Explicit

Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
'BGS used for Close window
Private Const WM_CLOSE = &H10


Public Enum EnumKeys
    KEYBACKSPACE = &H8
    keyTab = &H9
    keyReturn = &HD
    keyShift = &H10
    keyControl = &H11
    keyAlt = &H12
    keyPause = &H13
    keyEscape = &H1B
    keySpace = &H20
    keyEnd = &H23
    keyHome = &H24
    keyLeft = &H25
    KeyUp = &H26
    keyRight = &H27
    KeyDown = &H28
    keyInsert = &H2D
    keyDelete = &H2E
    keyF1 = &H70
    keyF2 = &H71
    keyF3 = &H72
    keyF4 = &H73
    keyF5 = &H74
    keyF6 = &H75
    keyF7 = &H76
    keyF8 = &H77
    keyF9 = &H78
    keyF10 = &H79
    keyF11 = &H7A
    keyF12 = &H7B
    keyNumLock = &H90
    keyScrollLock = &H91
    keyCapsLock = &H14
End Enum

Public Enum Variable
    Name = 0
    Value
End Enum

Private Const BAD_SEED As Long = -9999
Private Const SECURITY_SEED As Long = 91234

'BGS Diff Script Identifiers, this is what starts each line in the .scrpt file
Private Const ENCRYPT_TEXT As String = "~"
Private Const PLAIN_TEXT As String = "!"
Private Const VARIABLE_TEXT As String = "["
Private Const KEYSTROKE_TEXT As String = "{"
Private Const SLEEP_SECONDS As String = "*"
Private Const USER_INSTRUCTIONS As String = "@"

Private mbCancel As Boolean
Private mvVariables As Variant

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsKeyBoard"
End Property

Public Property Let Cancel(pbFlag As Boolean)
    mbCancel = pbFlag
End Property

Public Property Let Variables(pvVariables As Variant)
    mvVariables = pvVariables
End Property
Public Property Get Variables() As Variant
    Variables = mvVariables
End Property

Public Function AddVariable(psName As String, psValue As String) As Boolean
    On Error GoTo EH
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'BGS 1.23.2002 See if the Variable array has been set
    If Not IsArray(mvVariables) Then
        lCount = 0
        ReDim mvVariables(0 To 0, Variable.Name To Variable.Value)
    Else
        lCount = UBound(mvVariables, 1) + 1
        ReDim Preserve mvVariables(0 To lCount, Variable.Name To Variable.Value)
    End If
    
    mvVariables(lCount, Variable.Name) = psName
    mvVariables(lCount, Variable.Value) = psValue
    
    AddVariable = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function AddVariable" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function RemoveVariable(psName As String, Optional pbRemoveAll As Boolean = False) As Boolean
    On Error GoTo EH
    Dim lCount As Long
    Dim lNewCount As Long
    Dim bFound As Boolean
    Dim saryVariables() As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If IsArray(mvVariables) Then
        If pbRemoveAll Then
            Erase mvVariables
            RemoveVariable = True
            Exit Function
        End If
        
        For lCount = 0 To UBound(mvVariables, 1)
            If StrComp(mvVariables(lCount, Variable.Name), psName, vbTextCompare) = 0 Then
                bFound = True
                Exit For
            End If
        Next
    Else
        Exit Function
    End If
    
    If bFound Then
        mvVariables(lCount, Variable.Name) = vbNullString
    End If
    
    'This will clear out any Varibles that are null strings
    ReDim saryVariables(0 To 0, Variable.Name To Variable.Value)
    If IsArray(mvVariables) Then
        For lCount = 0 To UBound(mvVariables, 1)
            If mvVariables(lCount, Variable.Name) > vbNullString Then
                'Start lNewCount at 0
                ReDim Preserve saryVariables(0 To lNewCount, Variable.Name To Variable.Value)
                saryVariables(lNewCount, Variable.Name) = mvVariables(lCount, Variable.Name)
                saryVariables(lNewCount, Variable.Value) = mvVariables(lCount, Variable.Value)
                lNewCount = lNewCount + 1
            End If
        Next
    End If
    
    Erase mvVariables
    mvVariables = saryVariables
    Erase saryVariables
    RemoveVariable = bFound
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function AddVariable" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

'Presses the single key represented by sKey
Public Sub PressKey(sKey As String, Optional bHold As Boolean, Optional _
   bRelease As Boolean)

    Dim nVK As Long
    nVK = VkKeyScan(Asc(sKey))
    
    If nVK = 0 Then
        Exit Sub
    End If
    
    Dim nScan As Long
    Dim nExtended As Long
        
    nScan = MapVirtualKey(nVK, 2)
    nExtended = 0
    If nScan = 0 Then
        nExtended = KEYEVENTF_EXTENDEDKEY
    End If
    nScan = MapVirtualKey(nVK, 0)
    
    Dim bShift As Boolean
    Dim bCtrl As Boolean
    Dim bAlt As Boolean
    
    bShift = (nVK And &H100)
    bCtrl = (nVK And &H200)
    bAlt = (nVK And &H400)
    
    nVK = (nVK And &HFF)
    
    If Not bRelease Then
        If bShift Then
            keybd_event EnumKeys.keyShift, 0, 0, 0
        End If
        If bCtrl Then
            keybd_event EnumKeys.keyControl, 0, 0, 0
        End If
        If bAlt Then
            keybd_event EnumKeys.keyAlt, 0, 0, 0
        End If
    
        keybd_event nVK, nScan, nExtended, 0
    End If
    
    If Not bHold Then
        keybd_event nVK, nScan, KEYEVENTF_KEYUP Or nExtended, 0
    
        If bShift Then
            keybd_event EnumKeys.keyShift, 0, KEYEVENTF_KEYUP, 0
        End If
        If bCtrl Then
            keybd_event EnumKeys.keyControl, 0, KEYEVENTF_KEYUP, 0
        End If
        If bAlt Then
            keybd_event EnumKeys.keyAlt, 0, KEYEVENTF_KEYUP, 0
        End If
    End If

End Sub

'Loop through a string and calls PressKey for each character (Does not
' parse strings like SendKeys)
Public Sub PressString(ByVal sString As String, Optional bDoEvents As Boolean = True)
    Dim lSleep As Long
    Dim sSPeed As String
    
    '"Entry Speed - FAST"
    '"Entry Speed - MEDIUM"
    '"Entry Speed - SLOW"
    
    sSPeed = GetSetting("ECS", "KEYBOARD", "SPEED", "Entry Speed - FAST")
    
    If StrComp(sSPeed, "Entry Speed - FAST", vbTextCompare) = 0 Then
        lSleep = 0
    ElseIf StrComp(sSPeed, "Entry Speed - MEDIUM", vbTextCompare) = 0 Then
        lSleep = 40
    ElseIf StrComp(sSPeed, "Entry Speed - SLOW", vbTextCompare) = 0 Then
        lSleep = 90
    End If
    
    Do While sString <> vbNullString
        PressKey Mid(sString, 1, 1)
        If lSleep > 0 Then
            Sleep lSleep
        End If
        If bDoEvents Then
            DoEvents
        End If
        
        sString = Mid(sString, 2)
    Loop

End Sub

'Presses a specific key (this is used for keys that don't have a
' ascii equilivant)
Public Sub PressKeyVK(keyPress As EnumKeys, Optional bHold As Boolean, _
   Optional bRelease As Boolean, Optional bCompatible As Boolean)
    
    Dim nScan As Long
    Dim nExtended As Long
        
    nScan = MapVirtualKey(keyPress, 2)
    nExtended = 0
    If nScan = 0 Then
        nExtended = KEYEVENTF_EXTENDEDKEY
    End If
    nScan = MapVirtualKey(keyPress, 0)
    
    If bCompatible Then
        nExtended = 0
    End If
    
    If Not bRelease Then
        keybd_event keyPress, nScan, nExtended, 0
    End If
    
    If Not bHold Then
        keybd_event keyPress, nScan, KEYEVENTF_KEYUP Or nExtended, 0
    End If

End Sub

'Returns (in the boolean variables) the status of the various Lock keys
Public Sub GetLockStatus(bCapsLock As Boolean, bNumLock As Boolean, _
   bScrollLock As Boolean)

    bCapsLock = GetKeyState(EnumKeys.keyCapsLock)
    bNumLock = GetKeyState(EnumKeys.keyNumLock)
    bScrollLock = GetKeyState(EnumKeys.keyScrollLock)

End Sub

Public Function ProcessScript(psScriptName As String, psPath As String) As Boolean
    On Error GoTo EH
    Dim saryScript() As String
    Dim sScript As String
    Dim sScriptPath As String
    Dim lCount As Long
    Dim lSleep As Long
    Dim bNumLock As Boolean
    Dim bCapsLock As Boolean
    Dim bScrollLock As Boolean
    Dim lKey As EnumKeys
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'BGS 12.5.2001 Get Script data
    sScriptPath = psPath & "\" & psScriptName
    
    If FileExists(sScriptPath) Then
        sScript = GetFileData(sScriptPath)
        If sScript = vbNullString Then
            MsgBox "No script data found in " & sScriptPath & " !", vbExclamation, "NO SCRIPT DATA"
            Exit Function
        Else
            saryScript = Split(sScript, vbCrLf)
        End If
    Else
        MsgBox sScriptPath & " not found !", vbExclamation, "NO SCRIPT DATA"
        Exit Function
    End If
    
    'BGS 12.5.2001 check for num lock scroll lock caps lock
    'the only one we want on is Num lock
    GetLockStatus bCapsLock, bNumLock, bScrollLock
    
    If bCapsLock Then
        PressKeyVK keyCapsLock
    End If
    
    If Not bNumLock Then
        PressKeyVK keyNumLock
    End If
    
    If bScrollLock Then
        PressKeyVK keyScrollLock
    End If
    
    'BGS 12.5.2001 loop through the script commands and process them
    'according to what kind of command it is
    
    For lCount = 0 To UBound(saryScript, 1)
        If saryScript(lCount) <> vbNullString Then
            sScript = Mid(saryScript(lCount), 2)
        End If
        Select Case left(saryScript(lCount), 1)
            Case ENCRYPT_TEXT
                sScript = Decode(sScript)
                PressString sScript, True
                
            Case PLAIN_TEXT
                PressString sScript, True
                
            Case VARIABLE_TEXT
                sScript = GetVariable(sScript)
                If mbCancel Then
                    GoTo CLEAN_UP
                End If
                PressString sScript, True
                
            Case KEYSTROKE_TEXT
                lKey = GetKeyStroke(sScript)
                If mbCancel Then
                    GoTo CLEAN_UP
                End If
                
                PressKeyVK lKey
                Sleep 200
                DoEvents
                
            Case SLEEP_SECONDS
                'BGS 12.6.2001 Need to loop through the sleep and doevents
                'this will allow the single thread app (which VB6 is) to
                'finish doing stuff before we continue on with the script
                
                For lSleep = 1 To CDbl(sScript) * 10
                    Sleep 100
                    DoEvents
                Next
            Case USER_INSTRUCTIONS
                MsgBox sScript, vbInformation, "MANUAL INSTRUCTIONS"
WAIT_FOR_USER:
                'wait for 5 seconds
                For lSleep = 1 To 5 * 10
                    Sleep 100
                    DoEvents
                Next
                
                If MsgBox("Did you accomplish the user instructions ?" & vbCrLf & vbCrLf & sScript, vbYesNo) = vbNo Then
                    GoTo WAIT_FOR_USER
                End If
                
            Case Else
                If MsgBox("There was an error detected in this script !" & vbCrLf & _
                          sScriptPath & vbCrLf & _
                          "Do you want to abort this Script?" _
                          , vbYesNo, "ERROR IN " & psScriptName) = vbYes Then
                    mbCancel = True
                End If
                
        End Select
        'BGS 12.5.2001 check to see if the Cancel flag was tripped
        If mbCancel Then
            GoTo CLEAN_UP
        End If
    Next
    
    ProcessScript = True
CLEAN_UP:
    Erase saryScript
    Exit Function

EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function ProcessScript" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function GetVariable(psName As String) As String
    On Error GoTo EH
    Dim lCount As Long
    Dim bFound As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If IsArray(mvVariables) Then
        For lCount = 0 To UBound(mvVariables, 1)
            If StrComp(mvVariables(lCount, Variable.Name), psName, vbTextCompare) = 0 Then
                GetVariable = mvVariables(lCount, Variable.Value)
                bFound = True
                Exit For
            End If
        Next
    End If
    
    If Not bFound Then
        If MsgBox(psName & " Variable not found !" & vbCrLf & _
                  "Do you want to abort this Script?" _
                  , vbYesNo, "ERROR IN " & psName) = vbYes Then
            mbCancel = True
        End If
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function GetVariable" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function GetKeyStroke(psName As String) As EnumKeys
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Select Case UCase(psName)
        Case "KEYBACKSPACE"
            GetKeyStroke = KEYBACKSPACE
        Case "KEYTAB"
            GetKeyStroke = keyTab
        Case "KEYRETURN"
            GetKeyStroke = keyReturn
        Case "KEYSHIFT"
            GetKeyStroke = keyShift
        Case "KEYCONTROL"
            GetKeyStroke = keyControl
        Case "KEYALT"
            GetKeyStroke = keyAlt
        Case "KEYPAUSE"
            GetKeyStroke = keyPause
        Case "KEYESCAPE"
            GetKeyStroke = keyEscape
        Case "KEYSPACE"
            GetKeyStroke = keySpace
        Case "KEYEND"
            GetKeyStroke = keyEnd
        Case "KEYHOME"
            GetKeyStroke = keyHome
        Case "KEYLEFT"
            GetKeyStroke = keyLeft
        Case "KEYUP"
            GetKeyStroke = KeyUp
        Case "KEYRIGHT"
            GetKeyStroke = keyRight
        Case "KEYDOWN"
            GetKeyStroke = KeyDown
        Case "KEYINSERT"
            GetKeyStroke = keyInsert
        Case "KEYDELETE"
            GetKeyStroke = keyDelete
        Case "KEYF1"
            GetKeyStroke = keyF1
        Case "KEYF2"
            GetKeyStroke = keyF2
        Case "KEYF3"
            GetKeyStroke = keyF3
        Case "KEYF4"
            GetKeyStroke = keyF4
        Case "KEYF5"
            GetKeyStroke = keyF5
        Case "KEYF6"
            GetKeyStroke = keyF6
        Case "KEYF7"
            GetKeyStroke = keyF7
        Case "KEYF8"
            GetKeyStroke = keyF8
        Case "KEYF9"
            GetKeyStroke = keyF9
        Case "KEYF10"
            GetKeyStroke = keyF10
        Case "KEYF11"
            GetKeyStroke = keyF11
        Case "KEYF12"
            GetKeyStroke = keyF12
        Case "KEYNumLock"
            GetKeyStroke = keyNumLock
        Case "KEYScrollLock"
            GetKeyStroke = keyScrollLock
        Case "KEYCapsLock"
            GetKeyStroke = keyCapsLock
        Case Else
            If MsgBox(psName & " KeyStroke not found !" & vbCrLf & _
                      "Do you want to abort this Script?" _
                      , vbYesNo, "ERROR IN " & psName) = vbYes Then
                mbCancel = True
            End If
            
    End Select
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function GetKeyStroke" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Friend Function Encode(psPlainText As String) As String
'Purpose: Encodes String IE Password

'Parameters : psPassword As String

'Returns: Encrypted String

'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    Dim iChar As Integer    'ascii = random letter
    Dim iChar2 As Integer   'ascii = random letter
    Dim lMyKey() As Long    'array of (Long) elements indicating positions where characters of sMyCode belong
    Dim lx As Long, ly As Long 'lx for looping from begin to end of len(sMyString), ly end to Begin
    Dim sMyCode() As String 'array of (String) elements containing encoded characters
    Dim sMyCodeOutPut As String 'Concatenation of sMycode() string elements
    Dim sMyString As String 'String to be encoded
    Dim sMyCharPos As New Collection
    Dim sPlainText As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sPlainText = psPlainText

    Call Randomize
    sMyString = QtrStrHash(sPlainText, True) 'Sends input text to QtrStrHash function for initial encoding
    ReDim lMyKey(1 To Len(sMyString))   'Dimensions array to have same number of elements as there are characters in sMyString
    ReDim sMyCode(1 To Len(sMyString))  ' ""

    For lx = 1 To Len(sMyString)
        sMyCharPos.Add lx
    Next
    For lx = 1 To Len(sMyString) 'each element of lMyKey() represents a random position generated for each character of sMyString
        ly = 1 + Int(Rnd * sMyCharPos.Count)
        lMyKey(lx) = sMyCharPos.Item(ly)
        sMyCharPos.Remove (ly)
    Next
    Set sMyCharPos = Nothing
    lx = 0
    ly = Len(sMyString) + 1
    
    Do
        iChar = 97 + Int(Rnd * 26)
        iChar2 = 97 + Int(Rnd * 26)
        lx = lx + 1
        ly = ly - 1                 '''Reverses key order Respectively to sMystring'''       '123456 is used to offset glSecuritySeed with lMyKey using Additon and Subtraction
        sMyCode(lx) = Chr(iChar) & lMyKey(ly) & Mid(sMyString, lMyKey(lx), 1) & (lMyKey(ly) - 123456) + SECURITY_SEED & Chr(iChar2) & Chr(160)
        sMyCodeOutPut = sMyCodeOutPut & sMyCode(lx)                                                                                                         'ascii = (space)
    Loop Until lx = Len(sMyString)
    
    Encode = StrReverse(sMyCodeOutPut)

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Friend Function Encode" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Friend Function Decode(psEncryptedText As String, Optional bSeedCheck As Boolean) As String
'Purpose: Encodes String IE Password

'Parameters : psEncryptedText As String Optional bSeedCheck As Boolean
'                                       bSeedCheck used to make sure that
'                                       the Security File Has not been corrupted
'                                       or Someone Tried to Use a Different Security
'                                       File to Access the Current Machine.

'Returns: DeCrypted String

'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description

    On Error GoTo EH
    Dim lFindKeyStart() As Long
    Dim iKeyLen() As Integer
    Dim lx As Long, ly As Long
    Dim sMyKey As String
    Dim sMyCode As String
    Dim sMyDeCode() As Long
    Dim sMyDeCodeOutPut As String
    Dim sMyCodeTemp As String
    Dim sMyCodeTemp2 As String
    Dim lTestKey1 As Long
    Dim lTestKey2 As Long
    Dim sEncryptedText As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sEncryptedText = psEncryptedText
    
    sMyCode = StrReverse(sEncryptedText)
    lTestKey1 = Mid(sMyCode, InStr(1, sMyCode, Chr(160)) - 7, 6)
    lTestKey2 = Mid(sMyCode, 2, (InStr(1, sMyCode, Chr(160)) - 8) - 2)
    If Abs((lTestKey2 - 123456) - lTestKey1) <> SECURITY_SEED Then GoTo EH
    
    Do
        sMyCodeTemp = Mid(sMyCode, InStr(1, sMyCode, Chr(160)) + 1, Len(sMyCode) - InStr(1, sMyCode, Chr(160)))
        sMyKey = sMyKey & "," & Mid(sMyCode, 2, (InStr(1, sMyCode, Chr(160)) - 8) - 2)
        sMyCodeTemp2 = sMyCodeTemp2 & Mid(sMyCode, InStr(1, sMyCode, Chr(160)) - 8, 1)
        sMyCode = sMyCodeTemp
    Loop Until Len(sMyCode) < 8
    
    

    sMyCode = StrReverse(sMyCodeTemp2)
    sMyKey = sMyKey & ","
    ReDim lFindKeyStart(1 To Len(sMyCode))
    ReDim iKeyLen(1 To Len(sMyCode))
    ReDim sMyDeCode(1 To Len(sMyCode))

    For lx = 1 To Len(sMyCode)
        lFindKeyStart(lx) = InStr(1, sMyKey, ",") + 1
        iKeyLen(lx) = InStr(Val(lFindKeyStart(lx)), sMyKey, ",") - Val(lFindKeyStart(lx))
        sMyKey = Replace(sMyKey, ",", "!", 1, 1)
        sMyDeCode(lx) = Mid(sMyKey, lFindKeyStart(lx), iKeyLen(lx))
    Next lx

    For lx = 1 To Len(sMyCode)
        For ly = 1 To Len(sMyCode)
            If sMyDeCode(ly) = lx Then
                sMyDeCodeOutPut = sMyDeCodeOutPut & Mid(sMyCode, ly, 1)
                Exit For
            End If
        Next ly
    Next lx
    
    Decode = Trim(QtrStrHash(sMyDeCodeOutPut, False))

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    If bSeedCheck Then
        Decode = BAD_SEED
    Else
        Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Friend Function Decode" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
    End If
End Function

Private Function QtrStrHash(sMyText As String, bHash As Boolean)
'Purpose:  Used to Further Hash An encrypted String or De Hash

'Parameters : sMyText As String, bHash As Boolean

'Returns: N/A

'Author : BGS 3/10/2000

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    Dim s1Qtr As String
    Dim s2Qtr As String
    Dim s3Qtr As String
    Dim s4Qtr As String
    Dim s1Half As String
    Dim s2Half As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If bHash Then
        Do Until Len(sMyText) Mod 4 = False
            sMyText = sMyText & " "
        Loop
    End If
    s1Half = Mid(sMyText, 1, Int(Len(sMyText) / 2))
    s2Half = Mid(sMyText, Int(Len(sMyText) / 2) + 1, Len(sMyText) - (Int(Len(sMyText) / 2)))
    s1Qtr = Mid(s1Half, 1, Int(Len(s1Half) / 2))
    s2Qtr = Mid(s1Half, Int(Len(s1Half) / 2) + 1, Len(s1Half) - (Int(Len(s1Half) / 2)))
    s3Qtr = Mid(s2Half, 1, Int(Len(s2Half) / 2))
    s4Qtr = Mid(s2Half, Int(Len(s2Half) / 2) + 1, Len(s2Half) - (Int(Len(s2Half) / 2)))
    If bHash Then
            sMyText = StrReverse(s2Qtr) & s4Qtr & s3Qtr & StrReverse(s1Qtr)
    Else
            sMyText = StrReverse(s4Qtr) & StrReverse(s1Qtr) & s3Qtr & s2Qtr
    End If
    QtrStrHash = sMyText

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function QtrStrHash" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If IsObject(mvVariables) Then
        If Not mvVariables Is Nothing Then
            Set mvVariables = Nothing
        End If
    End If
    
    CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanUp" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function


