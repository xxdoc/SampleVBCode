VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSend2Billings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'These data types are what we need to populate into
'Send2Billings Base project info
Public Type udtClaims
    b01sCATNO As String
    b02sCLAIMNO As String
    b03sIB As String
    b04cGROSSLOSS As Currency
    b05cSERVICEFEE As Currency
    b06cEXPENSEREIM As Currency
    b07cADMINFEE As Currency
    b08sFULLNAME As String  'Insured Name
    b09sADJNAME As String
    b10sCLAIMCITY As String 'Enter City Before State
    b11sCLAIMSTATE As String
    b12dtFILESRECD As Date  'This is Date Of Loss
    b13dtDATECLOSED As Date
    b14sPROPADDR As String  'Loss Location
    b15dtDateUploaded As Date
    b16dtDateEntered As Date
    b17sComments As String
End Type

Public Enum FixClaimItem
    Status = 1
    Data = 2
    BillingDup
    b01sCATNO
    b01sCATNOSort
    b03sIB
    b03sIBSort
    b02sCLAIMNO
    b02sCLAIMNOSort
    b15dtDateUploaded
    b15dtDateUploadedSort
    b16dtDateEntered
    b16dtDateEnteredSort
    b13dtDATECLOSED
    b13dtDATECLOSEDSort
    b12dtFILESRECD 'Is Date Of Loss
    b12dtFILESRECDSort
    b09sADJNAME
    b08sFULLNAME
    b10sCLAIMCITY
    b11sCLAIMSTATE
    b04cGROSSLOSS
    b04cGROSSLOSSSort
    b05cSERVICEFEE
    b05cSERVICEFEESort
    b06cEXPENSEREIM
    b06cEXPENSEREIMSort
    b07cADMINFEE
    b07cADMINFEESort
    b14sPROPADDR
    b14sPROPADDRSort
    b17sComments
End Enum

Public Enum FixClaimPic
    NeedsSentToBilling = 1
    AlreadySentToBilling
    BadData
    GoodData
    Skip
    PointAt
    BillingDup
End Enum

Public Type udtClaimItem
    BillingDup As Boolean
    Main As udtClaims
    IsDirty As Boolean 'True if validation failed and adjuster fixed data and it then passed validation
    ValidData As Boolean
    SentToBillings As Boolean 'True if this item was sent to Send2Billings
    SkipThisItem As Boolean ' True if not want to send this one to Send2Billings
End Type

Private moKB As clsKeyBoard
Private moReg As clsRegSetting
Private mcolClaimItems As Collection 'Holds multiple Items to be added into Fix Claims List
Private mcolValidClaimItems As Collection  ' Holds Items that have been validated and or sent to Billings
Private msErrorInfo As String
Private mbCreatedOnce As Boolean 'Set this true the first time created
Private mbLookUpLoaded As Boolean 'True if we successfully loaded lookup info
Private mbItemsValidated As Boolean
Private mbSkipAll As Boolean 'True if exiting without sending to Billings
Private mlStartHwnd As Long
Private mbCancel As Boolean 'True if we ran clean
Private mConn As ADODB.Connection
Private mRS As ADODB.Recordset
Private mfrmFixClaims As frmFixClaims
Private moPB As ProgressBar
Private mdtStart As Date
Private mdtEnd As Date
Private msPrevClientClaimNo As String
Private msPrevIBNUM As String
Private mlItemsSkipped As Long
Private mlItemsSent As Long
Private mbSQLSERVER As Boolean

Public Property Get SQLServer() As Boolean
    SQLServer = mbSQLSERVER
End Property

Public Property Get StartDate() As Date
    StartDate = mdtStart
End Property

Public Property Get EndDate() As Date
    EndDate = mdtEnd
End Property

Public Property Let ProgBar(poProgBar As Object)
    If TypeOf poProgBar Is ProgressBar Then
        Set moPB = poProgBar
    End If
End Property
Public Property Set ProgBar(poProgBar As Object)
    If TypeOf poProgBar Is ProgressBar Then
        Set moPB = poProgBar
    End If
End Property
Public Property Get ProgBar() As Object
    Set ProgBar = moPB
End Property

Public Property Get Conn() As ADODB.Connection
    Set Conn = mConn
End Property

Public Property Get FailedValidFlag() As String
    FailedValidFlag = "?Unknown?"
End Property

Public Property Let CancelSendToSend2Billings(pbFlag As Boolean)
    mbCancel = pbFlag
End Property
Public Property Get CancelSendToSend2Billings() As Boolean
    CancelSendToSend2Billings = mbCancel
End Property

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsSend2Billings"
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

Public Property Let ClaimItems(pcolClaimItems As Collection)
    Set mcolClaimItems = pcolClaimItems
End Property
Public Property Set ClaimItems(pcolClaimItems As Collection)
    Set mcolClaimItems = pcolClaimItems
End Property
Public Property Get ClaimItems() As Collection
    Set ClaimItems = mcolValidClaimItems
End Property

Public Property Get ErrorInfo() As String
    ErrorInfo = msErrorInfo
End Property

Private Sub Sleepy(plSeconds As Long)
    Dim lSleep As Long
    
    For lSleep = 1 To plSeconds * 10
       DoEvents
       Sleep 100
    Next
End Sub

Public Function ValidateClaimItems(Optional pbShowFixClaimItemects As Boolean = True) As Boolean
    On Error GoTo EH
    
    If mcolClaimItems Is Nothing Then
        Exit Function
    End If
    If Not ShowFixSend2Billings Then
        Exit Function
    End If
    mbItemsValidated = True
    ValidateClaimItems = mbItemsValidated
    Exit Function
EH:
    mbItemsValidated = False
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function ValidClaimItems" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Public Function ShowFixSend2Billings() As Boolean
    On Error GoTo EH
    Dim bProjectsValidated As Boolean
    
    If mbCancel Then
        Exit Function
    End If
   
    With mfrmFixClaims
        .ClaimItems = mcolClaimItems
        .Send2Billings = Me
        .ModalFlag = True
        Load mfrmFixClaims
        .Show vbModeless
        Do
            DoEvents
            Sleep 10
            If gofrmMain.UnloadFlag Then
                Exit Function
            End If
        Loop Until mfrmFixClaims.Visible = False
        
        bProjectsValidated = .ItemsValidated
        mbSkipAll = .SkipAll
        .CLEANUP
    End With
    'Do Not set to nothing because need to maintain the Lookup Collection values
    'This form will be cleaned up when the main app exits.
    Unload mfrmFixClaims
    
    
    ShowFixSend2Billings = True
    
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function ShowFixSend2Billings" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    CloseConnection
    
    If Not moKB Is Nothing Then
        moKB.CLEANUP
        Set moKB = Nothing
    End If
    If Not moReg Is Nothing Then
        Set moReg = Nothing
    End If
    If Not mcolClaimItems Is Nothing Then
        Set mcolClaimItems = Nothing
    End If
    If Not mcolValidClaimItems Is Nothing Then
        Set mcolValidClaimItems = Nothing
    End If
    
    If Not gofrmMain Is Nothing Then
        If gofrmMain.UnloadFlag Then
            Unload mfrmFixClaims
            mfrmFixClaims.CLEANUP
            Set mfrmFixClaims = Nothing
            Set gofrmMain = Nothing
            Set goUtil = Nothing
            Set moPB = Nothing
        End If
    End If
    
    CLEANUP = True
    mbCancel = True
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanUp" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Public Function AddClaimItem(pvClaimItem As Variant, Optional pbValidated As Boolean) As Boolean
    On Error GoTo EH
    Dim ClaimItem As V2EberlsBillings.udtClaimItem
    If mbCancel Then
        Exit Function
    End If
    
    ClaimItem = pvClaimItem
    
    If Not pbValidated Then
        mcolClaimItems.Add ClaimItem, ClaimItem.Main.b03sIB
    Else
        mcolValidClaimItems.Add ClaimItem, ClaimItem.Main.b03sIB
    End If
    AddClaimItem = True
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function AddClaimItem" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Public Function SendKeysToBillings() As Boolean
    On Error GoTo EH
    Dim vClaimItem As Variant
    Dim ClaimItem As udtClaimItem
    Dim bFound As Boolean
    Dim sStartWindow As String
    Dim bCapsLock As Boolean
    Dim bNumLock As Boolean
    Dim bScrollLock As Boolean
    Dim sSendToBillingsVS As String
    Dim lCount As Long
    Dim MyStatus As frmStatus
    Dim sMess As String
    Dim bOKOn As Boolean
    Dim sClaimNoScrub As String '4.20.2003 Need to take out any " ", "-" from the number
    
    If mbCancel Then
        GoTo ABORT
    End If
    
    If Not mbItemsValidated Then
        Exit Function
    End If
    
    Set mcolClaimItems = Nothing
    Set mcolClaimItems = mcolValidClaimItems
    Set mcolValidClaimItems = Nothing
    Set mcolValidClaimItems = New Collection
    
    'Issue 183 9.8.2002 Send to Send2Billings puts info in all in Caps
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
    
    'See if there are any Items to be sent
    If mcolClaimItems.Count > 0 Then
        bFound = LookForWindow("Approach - [BILLINGS.APR:EC Send To Billings]")
        If Not bFound Then
            GoTo ABORT
        End If
        If mbCancel Then
            Exit Function
        End If
        moKB.PressKeyVK keyControl, True                'Hold Alt key
        moKB.PressKey "n"                               'Add
        moKB.PressKeyVK keyControl, , True              'Release Alt Key
        
        'Wait for 2 seconds
        Sleepy 2
        Set MyStatus = New frmStatus
        Load MyStatus
        MyStatus.Show vbModeless
        MyStatus.Visible = False
        bOKOn = CBool(GetSetting(App.EXEName, "GENERAL", "OK_ON", True))
        If Not gofrmMain Is Nothing Then
            gofrmMain.WindowState = vbMinimized
        End If
    End If
    
    For lCount = 1 To mcolClaimItems.Count
        If mbCancel Then
            GoTo ABORT
        End If
        vClaimItem = mcolClaimItems.Item(lCount)
        ClaimItem = vClaimItem
        If Not ClaimItem.SentToBillings Then
            If ClaimItem.ValidData And Not ClaimItem.SkipThisItem Then
                'Need to code for Different versions here...
                sSendToBillingsVS = GetSetting(goUtil.gsAppEXEName, "Version", "LatestVS", "1")
                Select Case UCase(sSendToBillingsVS)
                    Case "1" '<------------------Send2Billings vs 1 ------------------>

                        'Enter Information
                        If mbCancel Then
                            GoTo ABORT
                        End If
                        moKB.PressKeyVK keySpace                            'Press Space Bar
                        moKB.PressKeyVK keyShift, True                      'Hold Shift Key
                        moKB.PressKeyVK keyEnd                              'Press End Key
                        moKB.PressKeyVK keyShift, , True                    'Release Shift Key
                        
                        moKB.PressString ClaimItem.Main.b01sCATNO           'CATNO
                        PressTabKey 1
                                                                            'Tab 1 space to CLAIMNO
                        '4.20.2003 Need to make sure that the number is alphanumeric
                        sClaimNoScrub = ClaimItem.Main.b02sCLAIMNO
                        sClaimNoScrub = ScrubClientNumber(sClaimNoScrub)
                        moKB.PressString sClaimNoScrub                      'Scrubbed Client Claim Number
                        
                        PressTabKey 1                                       'Tab 1 space to IB
                        moKB.PressString ClaimItem.Main.b03sIB              'IB NUMBER
                        PressTabKey 4                                       'Tab 4 spaces to GROSSLOSS
                         'If there was an error sending...Abort the send process
                        If LookForWindow("Lotus Approach", 1) Then
                            GoTo ABORT
                        End If
                        moKB.PressString ClaimItem.Main.b04cGROSSLOSS       'GROSS LOSS
                        PressTabKey 1                                       'Tab 1 space to SERVICE FEE
                        moKB.PressString ClaimItem.Main.b05cSERVICEFEE      'SERVICE FEE
                        PressTabKey 1                                       'Tab 1 Space Expense Reimbersments (MISC EXPENSE)
                        moKB.PressString ClaimItem.Main.b06cEXPENSEREIM     'EXPENSE REIMBERSMENT
                        PressTabKey 1                                       'Tab 1 space to Admin Fee
                        moKB.PressString ClaimItem.Main.b07cADMINFEE        'ADMIN FEE
                        PressTabKey 1                                       'Tab 1 Space to FULL NAME
                        moKB.PressString ClaimItem.Main.b08sFULLNAME        'FULL NAME
                        PressTabKey 2                                       'Tab 2 spaces to ADJ NAME
                        moKB.PressString ClaimItem.Main.b09sADJNAME         'ADJUSTER NAME
                        PressTabKey 4                                       'Tab 4 spaces to CLAIM CITY
                        moKB.PressString ClaimItem.Main.b10sCLAIMCITY       'CLAIM CITY
                        'need to Back tab once
                        moKB.PressKeyVK keyShift, True                      'Hold Shift Key
                        moKB.PressKeyVK keyTab                              'Press End Key
                        moKB.PressKeyVK keyShift, , True                    'Release Shift Key
                        moKB.PressString ClaimItem.Main.b11sCLAIMSTATE      'CLAIM STATE
                        PressTabKey 5                                       'Tab 5 To FILESRECVD (DATE OF LOSS)
                        moKB.PressString ClaimItem.Main.b12dtFILESRECD      'FILESRECVD (DATE OF LOSS)
                        PressTabKey 1                                       'Tab 1 Space to DATE CLOSED
                        moKB.PressString ClaimItem.Main.b13dtDATECLOSED     'DATE CLOSED
                        PressTabKey 10                                      'Tab 10 Spaces to PROP_ADDR
                        moKB.PressString ClaimItem.Main.b14sPROPADDR        'PROPERTY ADDRESS (LOSS LOCATION)
                
                        moKB.PressKeyVK keyReturn                           'Hit Enter key
                        
                        MyStatus.ResetME
                        sMess = "Item " & lCount & " Of " & mcolClaimItems.Count & vbCrLf & vbCrLf
                        sMess = sMess & "Continue?"
                        MyStatus.lblMess.Caption = sMess
                        MyStatus.Visible = True
                        bFound = LookForWindow("Send Progress", 5)
                        If Not bFound Then
                            GoTo ABORT
                        Else
                            If bOKOn Then
                                Do
                                    DoEvents
                                    Sleep 10
                                Loop Until MyStatus.OKMe Or MyStatus.CancelMe
                                If Not MyStatus.OKMe Then
                                    GoTo ABORT
                                End If
                            Else
                                moKB.PressKeyVK keyReturn                           'Hit Enter key
                            End If
                        End If
                        
                        'If there are more Claim Items, be sure the current Window was not changed
                        If lCount < mcolClaimItems.Count Then
                            bFound = LookForWindow("Approach - [BILLINGS.APR:EC Send To Billings]")
                            If Not bFound Then
                                GoTo ABORT
                            End If
                            If mbCancel Then
                                GoTo ABORT
                            End If
                            moKB.PressKeyVK keyControl, True                'Hold Alt key
                            moKB.PressKey "n"                               'Add
                            moKB.PressKeyVK keyControl, , True              'Release Alt Key
                        End If
                    Case Else
                        mlItemsSkipped = mlItemsSkipped + 1
                        GoTo SKIP_PROJECT
                End Select
                
                'BGS Set flag that was sent to Send2Billings
                ClaimItem.SentToBillings = True
                'Need to Update DB as we go along...
                If Not FlagSend2Billings(ClaimItem) Then
                    'set the Cancel flag if we fail to update the
                    mbCancel = True
                Else
                    mlItemsSent = mlItemsSent + 1
                End If
            Else
                mlItemsSkipped = mlItemsSkipped + 1
                GoTo SKIP_PROJECT
            End If
        End If
        
NEXT_PROJECT:
        AddClaimItem ClaimItem, True
SKIP_PROJECT:
    Next
    SendKeysToBillings = True
ABORT:
    If Not MyStatus Is Nothing Then
        Unload MyStatus
        Set MyStatus = Nothing
    End If
    
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function SendKeysToBillings" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Private Function ScrubClientNumber(psScrubMe As String, Optional pbAllowSpaces As Boolean = False) As String
    On Error GoTo EH
    Dim sSCrub As String
    Dim sClean As String
    Dim lPos As Long
    Dim sLetter As String
    
    sSCrub = psScrubMe
    sClean = psScrubMe
    
    For lPos = 1 To Len(sSCrub)
        sLetter = Mid(sSCrub, lPos, 1)
        If Not IsAlphaNumeric(sLetter) Then
            sClean = Replace(sClean, sLetter, vbNullString, , , vbBinaryCompare)
        End If
    Next
    
    ScrubClientNumber = sClean
    
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function ScrubClientNumber" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Private Function IsAlphaNumeric(psLetter As String, Optional pbAllowSpaces As Boolean = False) As Boolean
    On Error GoTo EH
    If pbAllowSpaces Then
        IsAlphaNumeric = psLetter Like "[A-Za-z0-9_- ]"
    Else
        IsAlphaNumeric = psLetter Like "[A-Za-z0-9]"
    End If
    
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function IsAlphaNumeric" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Public Function LookForWindow(psWindowCaption As String, Optional plTimeOutSeconds As Long = 60) As Boolean
    On Error GoTo EH
    Dim bFound As Boolean
    Dim lSleep As Long
    If mbCancel Then
        Exit Function
    End If
    bFound = True
    Do Until goUtil.utWindowFound(psWindowCaption)
        lSleep = lSleep + 1
        If lSleep > plTimeOutSeconds Then
            bFound = False
            Exit Do
        End If
        If mbCancel Then
            Exit Function
        End If
    Loop
    
    LookForWindow = bFound
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function LookForWindow" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
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

Friend Function ValidateClaimItem(pvClaimItem As Variant, Optional pitmX As ListItem, Optional pbIgnorePrevClientClaimNo As Boolean) As Boolean
    On Error GoTo EH
    Dim ClaimItem As udtClaimItem
    Dim itmX As ListItem
    Dim bitmX As Boolean
    Dim iPic As FixClaimPic
    Dim bRebillsOn As Boolean
    
    bRebillsOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "REBILLS_ON", False))
    
    If mbCancel Then
        Exit Function
    End If
    
    If Not IsEmpty(pitmX) Then
        Set itmX = pitmX
        bitmX = True
    End If
    ClaimItem = pvClaimItem
    ClaimItem.ValidData = True
    'Validate the Main Data

    With ClaimItem.Main
        If ClaimItem.BillingDup Then
            EditListItem True, itmX, FixClaimItem.BillingDup, "X", FixClaimPic.BillingDup
        Else
            EditListItem True, itmX, FixClaimItem.BillingDup, vbNullString, Empty
        End If
        If Not FindItemInCol(mfrmFixClaims.colCATNO, .b01sCATNO) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b01sCATNO, .b01sCATNO, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b01sCATNOSort, goUtil.utNumInTextSortFormat(.b01sCATNO), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b01sCATNO, .b01sCATNO, Empty
            EditListItem bitmX, itmX, FixClaimItem.b01sCATNOSort, goUtil.utNumInTextSortFormat(.b01sCATNO), Empty
        End If
        
        If Not ValidStringData(.b02sCLAIMNO) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b02sCLAIMNO, .b02sCLAIMNO, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b02sCLAIMNOSort, goUtil.utNumInTextSortFormat(.b02sCLAIMNO), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b02sCLAIMNO, .b02sCLAIMNO, Empty
            EditListItem bitmX, itmX, FixClaimItem.b02sCLAIMNOSort, goUtil.utNumInTextSortFormat(.b02sCLAIMNO), Empty
        End If
        
        'When validating the IB number, check the Client Claim Number against the previous one.
        'If the current Client Claim Number is the same as the previous one then put a Pointer
        'next to the IB number because these items need to researched.
        If Not ValidStringData(.b03sIB) Then
            ClaimItem.ValidData = False
            If Not pbIgnorePrevClientClaimNo Then
                If StrComp(msPrevClientClaimNo, .b02sCLAIMNO, vbTextCompare) = 0 And bRebillsOn Then
                    iPic = PointAt
                Else
                    iPic = BadData
                End If
            Else
                iPic = BadData
            End If
            EditListItem bitmX, itmX, FixClaimItem.b03sIB, .b03sIB, iPic
            EditListItem bitmX, itmX, FixClaimItem.b03sIBSort, goUtil.utNumInTextSortFormat(.b03sIB), iPic
        Else
            If Not pbIgnorePrevClientClaimNo Then
                If StrComp(msPrevClientClaimNo, .b02sCLAIMNO, vbTextCompare) = 0 _
                    And bRebillsOn _
                    And itmX.Index <> 1 _
                    And IsTheSameSupplement(msPrevIBNUM, .b03sIB) Then
                    iPic = PointAt
                    EditListItem bitmX, itmX, FixClaimItem.b03sIB, .b03sIB, iPic
                    EditListItem bitmX, itmX, FixClaimItem.b03sIBSort, goUtil.utNumInTextSortFormat(.b03sIB), iPic
                Else
                    EditListItem bitmX, itmX, FixClaimItem.b03sIB, .b03sIB, Empty
                    EditListItem bitmX, itmX, FixClaimItem.b03sIBSort, goUtil.utNumInTextSortFormat(.b03sIB), Empty
                End If
            Else
                If itmX.ListSubItems(FixClaimItem.b03sIB - 1).ReportIcon <> Empty Then
                    iPic = itmX.ListSubItems(FixClaimItem.b03sIB - 1).ReportIcon
                    EditListItem bitmX, itmX, FixClaimItem.b03sIB, .b03sIB, iPic
                    EditListItem bitmX, itmX, FixClaimItem.b03sIBSort, goUtil.utNumInTextSortFormat(.b03sIB), iPic
                Else
                    EditListItem bitmX, itmX, FixClaimItem.b03sIB, .b03sIB, Empty
                    EditListItem bitmX, itmX, FixClaimItem.b03sIBSort, goUtil.utNumInTextSortFormat(.b03sIB), Empty
                End If
            End If
            
        End If
        
        'Set the prev client claim number for the next validation
        msPrevClientClaimNo = .b02sCLAIMNO
        msPrevIBNUM = .b03sIB
        
        If Not ValidCurrencyData(.b04cGROSSLOSS) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b04cGROSSLOSS, .b04cGROSSLOSS, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b04cGROSSLOSSSort, goUtil.utNumInTextSortFormat(CStr(.b04cGROSSLOSS)), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b04cGROSSLOSS, .b04cGROSSLOSS, Empty
            EditListItem bitmX, itmX, FixClaimItem.b04cGROSSLOSSSort, goUtil.utNumInTextSortFormat(CStr(.b04cGROSSLOSS)), Empty
        End If
        
        If Not ValidCurrencyData(.b05cSERVICEFEE) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b05cSERVICEFEE, .b05cSERVICEFEE, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b05cSERVICEFEESort, goUtil.utNumInTextSortFormat(CStr(.b05cSERVICEFEE)), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b05cSERVICEFEE, .b05cSERVICEFEE, Empty
            EditListItem bitmX, itmX, FixClaimItem.b05cSERVICEFEESort, goUtil.utNumInTextSortFormat(CStr(.b05cSERVICEFEE)), Empty
        End If
        
        If Not ValidCurrencyData(.b06cEXPENSEREIM) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b06cEXPENSEREIM, .b06cEXPENSEREIM, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b06cEXPENSEREIMSort, goUtil.utNumInTextSortFormat(CStr(.b06cEXPENSEREIM)), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b06cEXPENSEREIM, .b06cEXPENSEREIM, Empty
            EditListItem bitmX, itmX, FixClaimItem.b06cEXPENSEREIMSort, goUtil.utNumInTextSortFormat(CStr(.b06cEXPENSEREIM)), Empty
        End If
        
        If Not ValidCurrencyData(.b07cADMINFEE) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b07cADMINFEE, .b07cADMINFEE, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b07cADMINFEESort, goUtil.utNumInTextSortFormat(CStr(.b07cADMINFEE)), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b07cADMINFEE, .b07cADMINFEE, Empty
            EditListItem bitmX, itmX, FixClaimItem.b07cADMINFEESort, goUtil.utNumInTextSortFormat(CStr(.b07cADMINFEE)), Empty
        End If
        
        If Not ValidStringData(.b08sFULLNAME) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b08sFULLNAME, .b08sFULLNAME, FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b08sFULLNAME, .b08sFULLNAME, Empty
        End If
        
        If Not FindItemInCol(mfrmFixClaims.colADJNAME, .b09sADJNAME) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b09sADJNAME, .b09sADJNAME, FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b09sADJNAME, .b09sADJNAME, Empty
        End If
        
        'When Validating the CLAIM CITY SEE if the STATE is in the Taxable
        'State Collection.  If it is then Check to see if the Claim City is in
        'Collection of Taxable cities.
        If FindItemInCol(mfrmFixClaims.colTAXSTATES, .b11sCLAIMSTATE) Then
            If Not FindItemInCol(mfrmFixClaims.colCLAIMCITY, .b10sCLAIMCITY) Then
                ClaimItem.ValidData = False
                EditListItem bitmX, itmX, FixClaimItem.b10sCLAIMCITY, .b10sCLAIMCITY, FixClaimPic.BadData
            Else
                EditListItem bitmX, itmX, FixClaimItem.b10sCLAIMCITY, .b10sCLAIMCITY, Empty
            End If
        Else
            If Not ValidStringData(.b10sCLAIMCITY) Then
                ClaimItem.ValidData = False
                EditListItem bitmX, itmX, FixClaimItem.b10sCLAIMCITY, .b10sCLAIMCITY, FixClaimPic.BadData
            Else
                EditListItem bitmX, itmX, FixClaimItem.b10sCLAIMCITY, .b10sCLAIMCITY, Empty
            End If
        End If
        
        If Not ValidStringData(.b11sCLAIMSTATE) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b11sCLAIMSTATE, .b11sCLAIMSTATE, FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b11sCLAIMSTATE, .b11sCLAIMSTATE, Empty
        End If
        
        If Not DateValid(CStr(.b12dtFILESRECD)) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b12dtFILESRECD, .b12dtFILESRECD, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b12dtFILESRECDSort, Format(CStr(.b12dtFILESRECD), "YYYY/MM/DD"), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b12dtFILESRECD, .b12dtFILESRECD, Empty
            EditListItem bitmX, itmX, FixClaimItem.b12dtFILESRECDSort, Format(CStr(.b12dtFILESRECD), "YYYY/MM/DD"), Empty
        End If
        
        If Not DateValid(CStr(.b13dtDATECLOSED)) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b13dtDATECLOSED, .b13dtDATECLOSED, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b13dtDATECLOSEDSort, Format(CStr(.b13dtDATECLOSED), "YYYY/MM/DD"), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b13dtDATECLOSED, .b13dtDATECLOSED, Empty
            EditListItem bitmX, itmX, FixClaimItem.b13dtDATECLOSEDSort, Format(CStr(.b13dtDATECLOSED), "YYYY/MM/DD"), Empty
        End If
        
        'These items are determined programatically (not by the user)
        'do not need validation.
        EditListItem bitmX, itmX, FixClaimItem.b15dtDateUploaded, .b15dtDateUploaded, Empty
        EditListItem bitmX, itmX, FixClaimItem.b15dtDateUploadedSort, Format(CStr(.b15dtDateUploaded), "YYYY/MM/DD"), Empty
        EditListItem bitmX, itmX, FixClaimItem.b16dtDateEntered, .b16dtDateEntered, Empty
        EditListItem bitmX, itmX, FixClaimItem.b16dtDateEnteredSort, Format(CStr(.b16dtDateEntered), "YYYY/MM/DD"), Empty
        
        If Not ValidStringData(.b14sPROPADDR) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b14sPROPADDR, .b14sPROPADDR, FixClaimPic.BadData
            EditListItem bitmX, itmX, FixClaimItem.b14sPROPADDRSort, goUtil.utNumInTextSortFormat(.b14sPROPADDR), FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b14sPROPADDR, .b14sPROPADDR, Empty
            EditListItem bitmX, itmX, FixClaimItem.b14sPROPADDRSort, goUtil.utNumInTextSortFormat(.b14sPROPADDR), Empty
        End If
        
        If Not ValidStringData(.b17sComments) Then
            ClaimItem.ValidData = False
            EditListItem bitmX, itmX, FixClaimItem.b17sComments, .b17sComments, FixClaimPic.BadData
        Else
            EditListItem bitmX, itmX, FixClaimItem.b17sComments, .b17sComments, Empty
        End If
        
        
    End With
    
    If ClaimItem.ValidData Then
        ValidateClaimItem = True
        EditListItem bitmX, itmX, FixClaimItem.Data, vbNullString, FixClaimPic.GoodData
    Else
        EditListItem bitmX, itmX, FixClaimItem.Data, "X", FixClaimPic.BadData
    End If
    
    ValidateClaimItem = True
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Friend Function ValidateClaimItem" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Private Function IsTheSameSupplement(psPrevIBNUM As String, psThisIBNUM As String) As Boolean
    On Error GoTo EH
    Dim sTemp As String
    'Prev
    Dim sPrevSup As String
    Dim lPrevSupS As Long
    Dim lPrevSupValue As Long
    Dim sPrevSupS As String
    Dim lPrevSupR As Long
    Dim sPrevSupR As String
    'This
    Dim sThisSup As String
    Dim lThisSupS As Long
    Dim lThisSupValue As Long
    Dim sThisSupS As String
    Dim lThisSupR As Long
    Dim sThisSupR As String
    
    'Need to compare the previous IBNUM to see if it is part of the same supplement
    
    sPrevSup = Mid(psPrevIBNUM, 6, Len(psPrevIBNUM))
    lPrevSupS = InStr(1, sPrevSup, "S", vbTextCompare)
    lPrevSupR = InStr(1, sPrevSup, "R", vbTextCompare)
    If lPrevSupR > 0 Then
        If lPrevSupS > 0 Then
            sTemp = Mid(sPrevSup, lPrevSupS + 1)
            sTemp = Left(sTemp, InStr(1, sTemp, "R", vbTextCompare) - 1)
            lPrevSupValue = CLng(sTemp)
        Else
            lPrevSupValue = 0
        End If
    ElseIf lPrevSupS > 0 Then
        sTemp = Mid(sPrevSup, lPrevSupS + 1)
        lPrevSupValue = CLng(sTemp)
    Else
        lPrevSupValue = 0
    End If
    
    sThisSup = Mid(psThisIBNUM, 6, Len(psThisIBNUM))
    lThisSupS = InStr(1, sThisSup, "S", vbTextCompare)
    lThisSupR = InStr(1, sThisSup, "R", vbTextCompare)
    If lThisSupR > 0 Then
        If lThisSupS > 0 Then
            sTemp = Mid(sThisSup, lThisSupS + 1)
            sTemp = Left(sTemp, InStr(1, sTemp, "R", vbTextCompare) - 1)
            lThisSupValue = CLng(sTemp)
        Else
            lThisSupValue = 0
        End If
    ElseIf lThisSupS > 0 Then
        sTemp = Mid(sThisSup, lThisSupS + 1)
        lThisSupValue = CLng(sTemp)
    Else
        lThisSupValue = 0
    End If
    
    'need to compare this sup to previous supp to see if it is part of the same Supplemnt
    If lPrevSupValue = lThisSupValue Then
        IsTheSameSupplement = True
    Else
        IsTheSameSupplement = False
    End If
    
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function IsTheSameSupplement" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Private Sub EditListItem(pbItemX As Boolean, pitmX As ListItem, piSubItemIndex As Long, pvValue As Variant, pvReportIcon As Variant)
    On Error GoTo EH
    Dim vValue As Variant
    
    vValue = pvValue
    
    If VarType(vValue) = vbCurrency Then
        vValue = Format(vValue, "###,##0.00")
    ElseIf VarType(vValue) = vbDate Then
        If vValue = NULL_DATE Then
            vValue = vbNullString
        Else
            vValue = Format(vValue, "mm/dd/yyyy")
        End If
    End If
    
    If pbItemX Then
        If pitmX.SubItems(piSubItemIndex - 1) <> vValue Or vValue = vbNullString Then
            pitmX.SubItems(piSubItemIndex - 1) = vValue
        End If
        If pitmX.ListSubItems(piSubItemIndex - 1).ReportIcon <> pvReportIcon Then
            pitmX.ListSubItems(piSubItemIndex - 1).ReportIcon = pvReportIcon
        End If
    End If
    
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub EditListItem" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub


Private Function ValidCurrencyData(pcCur As Currency) As Boolean
    On Error GoTo EH
    If mbCancel Then
        Exit Function
    End If
    
    If pcCur >= 0 Then
        If pcCur <= 99000000 Then
            ValidCurrencyData = True
        End If
    End If
    
    Exit Function
EH:
    ValidCurrencyData = False
End Function

Private Function ValidStringData(psText As String) As Boolean
    On Error GoTo EH
    If mbCancel Then
        Exit Function
    End If
    
    If InStr(1, psText, FailedValidFlag, vbTextCompare) = 0 Then
        ValidStringData = True
    End If
    
    Exit Function
EH:
    ValidStringData = False
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

Public Sub PopulateLookUp(Optional pbHideError As Boolean)
    
    '1.27.2003 Populate the From and End Dates, the dates available to import
    'from the Batches table.  As Well, Populate the Lookup items for FIx Claims Screen.
    'These Items will be used by admin folks to correct validation failed records.
    On Error GoTo EH
    Dim sClaimsDSN As String
    Dim bPB As Boolean
    
    If Not ProgBar Is Nothing Then
        ProgBar.Max = 100
        ProgBar.Value = 0
        bPB = True
    End If
    
    If Not mfrmFixClaims Is Nothing Then
        Unload mfrmFixClaims
        mfrmFixClaims.CLEANUP
        Set mfrmFixClaims = Nothing
    End If
    
    Set mfrmFixClaims = New frmFixClaims

    'If the Claims DSN is balnk need to prompt to set it up
    sClaimsDSN = GetSetting(goUtil.gsAppEXEName, "DBConn", "Claims", vbNullString)
    
    If sClaimsDSN = vbNullString Then
        MsgBox "Please set up Claims Billing DSN!", vbExclamation + vbOKOnly, "Claims Billing DSN"
        Exit Sub
    End If
        
    If Not OpenConnection Then
        GoTo CLEAN_UP
    End If
    
    'Use adAsyncExecute so the user can do other stuff while the Lookup boxes are loading.
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    
    
    'CATNO
    If gofrmMain.UnloadFlag Then
        GoTo CLEAN_UP
    End If
    'SQL SERVER is using Linked Servers to FOXPRO Tabels SO must use OpenQuery command.
    'ClaimProdV1 = name of the linked server = Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=\\eberls01\h$\Claims Prgm Data;SourceType=DBF;Exclusive=No
    If mbSQLSERVER Then
        PopulateCATNO_SQLSVR bPB
    Else
        PopulateCATNO bPB
    End If
    
    'ADJNAME
    If gofrmMain.UnloadFlag Then
        GoTo CLEAN_UP
    End If
    
    'SQL SERVER is using Linked Servers to FOXPRO Tabels SO must use OpenQuery command.
    'ClaimProdV1 = name of the linked server = Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=\\eberls01\h$\Claims Prgm Data;SourceType=DBF;Exclusive=No
    If mbSQLSERVER Then
        PopulateADJNAME_SQLSVR bPB
    Else
        PopulateADJNAME bPB
    End If
    
    'TAXABLE CLAIM CITY
    If gofrmMain.UnloadFlag Then
        GoTo CLEAN_UP
    End If
    
    'SQL SERVER is using Linked Servers to FOXPRO Tabels SO must use OpenQuery command.
    'ClaimProdV1 = name of the linked server = Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=\\eberls01\h$\Claims Prgm Data;SourceType=DBF;Exclusive=No
    If mbSQLSERVER Then
        PopulateCLAIMCITY_SQLSVR bPB
    Else
        PopulateCLAIMCITY bPB
    End If
    
    
    'TAXABLE CLAIM STATES
    If gofrmMain.UnloadFlag Then
        GoTo CLEAN_UP
    End If
    'SQL SERVER is using Linked Servers to FOXPRO Tabels SO must use OpenQuery command.
    'ClaimProdV1 = name of the linked server = Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=\\eberls01\h$\Claims Prgm Data;SourceType=DBF;Exclusive=No
    If mbSQLSERVER Then
        PopulateTAXSTATES_SQLSVR bPB
    Else
        PopulateTAXSTATES bPB
    End If
    
    
    'Need to Populate Start and End Dates
    If gofrmMain.UnloadFlag Then
        GoTo CLEAN_UP
    End If
    'SQL SERVER is using Linked Servers to FOXPRO Tabels SO must use OpenQuery command.
    'ClaimProdV1 = name of the linked server = Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=\\eberls01\h$\Claims Prgm Data;SourceType=DBF;Exclusive=No
    If mbSQLSERVER Then
        PopulateStartEnd_SQLSVR bPB
    Else
        PopulateStartEnd bPB
    End If
    
    'Populate By Cat Site Drop Down
    If mbSQLSERVER Then
        PopulateByCatSite_SQLSVR bPB
    Else
        'your hosed'
    End If
    
    'Populate By adjuster Drop Down
    If mbSQLSERVER Then
        PopulateByAdjuster_SQLSVR bPB
    Else
        'your still hosed'
    End If
    
    mbLookUpLoaded = True
CLEAN_UP:
    Set mRS = Nothing
    Exit Sub
EH:
    mbLookUpLoaded = False
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub PopulateLookUp" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateCATNO(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    'CATNO
    sSQL = "SELECT "
    sSQL = sSQL & "SITE.SITE_NO "
    sSQL = sSQL & "FROM SITE "
    sSQL = sSQL & "GROUP BY SITE.SITE_NO "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!SITE_NO), vbNullString, RTrim(mRS!SITE_NO))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colCATNO = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub PopulateCATNO" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateCATNO_SQLSVR(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    'CATNO
    sSQL = "SELECT "
    sSQL = sSQL & "SITE.SITE_NO "
    sSQL = sSQL & "FROM SITE "
    sSQL = sSQL & "GROUP BY SITE.SITE_NO "
    
    sSQL = "Select * FROM OpenQuery(ClaimProdV1,'" & sSQL & "')"
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!SITE_NO), vbNullString, RTrim(mRS!SITE_NO))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colCATNO = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub PopulateCATNO_SQLSVR" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateADJNAME(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    'ADJNAME
    sSQL = "SELECT "
    sSQL = sSQL & "ECSADJ.LAST_NAME, "
    sSQL = sSQL & "ECSADJ.FIRST_NAME "
    sSQL = sSQL & "FROM ECSADJ "
    sSQL = sSQL & "ORDER BY ECSADJ.LAST_NAME, ECSADJ.FIRST_NAME "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
           
            If Not (IsNull(mRS!LAST_NAME)) And Not (IsNull(mRS!FIRST_NAME)) Then
                sValue = RTrim(mRS!LAST_NAME) & " " & RTrim(mRS!FIRST_NAME)
                ADDItemToCol colObject, sValue, sValue
            End If
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colADJNAME = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub PopulateADJNAME" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Public Sub PopulateADJNAME_SQLSVR(pbPB As Boolean, Optional psCatSite As String = vbNullString)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    Dim sCatSite As String
    
    sCatSite = Trim(psCatSite)
    bPB = pbPB
    
    'ADJNAME
    '4.19.2003 Use the Special Table 4.19.2003
    sSQL = "SELECT "
    sSQL = sSQL & "ADJUSTER_N "
    sSQL = sSQL & "FROM SPECIAL "
    If sCatSite <> vbNullString Then
        sSQL = sSQL & "WHERE EFFECTIVE_ = ''" & sCatSite & "'' "
    End If
    sSQL = sSQL & "ORDER BY ADJUSTER_N "
    sSQL = "Select * FROM OpenQuery(ClaimProdV1,'" & sSQL & "') "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
           
            If Not IsNull(mRS!adjuster_n) Then
                sValue = RTrim(mRS!adjuster_n)
                ADDItemToCol colObject, sValue, sValue
            End If
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colADJNAME = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub PopulateADJNAME_SQLSVR" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub


Private Sub PopulateCLAIMCITY(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    'CITY
    sSQL = "SELECT "
    sSQL = sSQL & "TXTAX.CITY_NAME "
    sSQL = sSQL & "FROM TXTAX "
    sSQL = sSQL & "WHERE LEN(RTRIM(TXTAX.ALIAS)) > 0 "
    sSQL = sSQL & "ORDER BY CITY_NAME "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!CITY_NAME), vbNullString, UCase(RTrim(mRS!CITY_NAME)))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colCLAIMCITY = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub PopulateCLAIMCITY" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateCLAIMCITY_SQLSVR(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    'CITY
    sSQL = "SELECT "
    sSQL = sSQL & "TXTAX.CITY_NAME "
    sSQL = sSQL & "FROM TXTAX "
    sSQL = sSQL & "WHERE LEN(RTRIM(TXTAX.ALIAS)) > 0 "
    sSQL = sSQL & "ORDER BY CITY_NAME "
    
    sSQL = "Select * FROM OpenQuery(ClaimProdV1,'" & sSQL & "') "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!CITY_NAME), vbNullString, UCase(RTrim(mRS!CITY_NAME)))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colCLAIMCITY = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub PopulateCLAIMCITY_SQLSVR" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateTAXSTATES(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    'TAXABLE STATE
    sSQL = "SELECT "
    sSQL = sSQL & "TXTAX.STATE "
    sSQL = sSQL & "FROM TXTAX "
    sSQL = sSQL & "WHERE LEN(RTRIM(TXTAX.ALIAS)) > 0 "
    sSQL = sSQL & "GROUP BY STATE "
    sSQL = sSQL & "ORDER BY STATE "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!State), vbNullString, UCase(RTrim(mRS!State)))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colTAXSTATES = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub PopulateTAXSTATES" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateTAXSTATES_SQLSVR(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    'TAXABLE STATE
    sSQL = "SELECT "
    sSQL = sSQL & "TXTAX.STATE "
    sSQL = sSQL & "FROM TXTAX "
    sSQL = sSQL & "WHERE LEN(RTRIM(TXTAX.ALIAS)) > 0 "
    sSQL = sSQL & "GROUP BY STATE "
    sSQL = sSQL & "ORDER BY STATE "
    
    sSQL = "Select * FROM OpenQuery(ClaimProdV1,'" & sSQL & "') "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!State), vbNullString, UCase(RTrim(mRS!State)))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set mfrmFixClaims.colTAXSTATES = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub PopulateTAXSTATES_SQLSVR" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateStartEnd(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    sSQL = "SELECT "
    sSQL = sSQL & "CTOD(BATCHES.DATE) AS GETDATE "
    sSQL = sSQL & "FROM BATCHES "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "BATCHES.ECUPDATED = 1 "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "BATCHES.DUPLICATE = 0 "
    sSQL = sSQL & "GROUP BY GETDATE "
    sSQL = sSQL & "ORDER BY GETDATE DESC "
    
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!GETDATE), vbNullString, RTrim(mRS!GETDATE))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set gofrmMain.colDates = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub PopulateStartEnd" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateStartEnd_SQLSVR(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    sSQL = "SELECT "
    sSQL = sSQL & "CONVERT(DATETIME,(BATCHES.DATE)) AS GETDATE "
    sSQL = sSQL & "FROM BATCHES "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "BATCHES.ECUPDATED = 1 "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "BATCHES.DUPLICATE = 0 "
    sSQL = sSQL & "GROUP BY CONVERT(DATETIME,(BATCHES.DATE)) "
    sSQL = sSQL & "ORDER BY CONVERT(DATETIME,(BATCHES.DATE)) DESC "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!GETDATE), vbNullString, RTrim(mRS!GETDATE))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set gofrmMain.colDates = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub PopulateStartEnd_SQLSVR" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateByCatSite_SQLSVR(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    sSQL = "SELECT "
    sSQL = sSQL & "CATSITE "
    sSQL = sSQL & "FROM BATCHES "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "ECUPDATED = 1 "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "DUPLICATE = 0 "
    sSQL = sSQL & "GROUP BY CATSITE "
    sSQL = sSQL & "ORDER BY CATSITE "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!catsite), vbNullString, RTrim(mRS!catsite))
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set gofrmMain.colCatSites = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub PopulateByCatSite_SQLSVR" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub PopulateByAdjuster_SQLSVR(pbPB As Boolean)
    On Error GoTo EH
    Dim sSQL As String
    Dim colObject As Collection
    Dim bPB As Boolean
    Dim sValue As String
    
    bPB = pbPB
    
    sSQL = "SELECT "
    sSQL = sSQL & "ADJUSTER_N, SSN "
    sSQL = sSQL & "FROM BATCHES "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "ECUPDATED = 1 "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "DUPLICATE = 0 "
    sSQL = sSQL & "GROUP BY ADJUSTER_N, SSN "
    sSQL = sSQL & "ORDER BY ADJUSTER_N "
    
    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo CLEAN_UP
        End If
    Loop
    
    If Not mRS.EOF Then
        mRS.MoveFirst
        Set colObject = New Collection
        Do Until mRS.EOF
            sValue = IIf(IsNull(mRS!adjuster_n), vbNullString, RTrim(mRS!adjuster_n))
            sValue = sValue & " (" & IIf(IsNull(mRS!SSN), 0, mRS!SSN) & ")"
            ADDItemToCol colObject, sValue, sValue
            mRS.MoveNext
        Loop
    End If
    
    If Not colObject Is Nothing Then
        Set gofrmMain.colAdjusters = colObject
        Set colObject = Nothing
    End If
    If bPB Then
        ProgBar.Value = 0
    End If
    mRS.Close
    
CLEAN_UP:
    Set colObject = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub PopulateByAdjuster_SQLSVR" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Function UCaseOnly(psText As String) As String
    On Error GoTo EH
    Dim lCount As Long
    Dim lPos As Long
    Dim sTemp As String
    Dim sChar As String
    
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
                    sTemp = Left(sTemp, Len(sTemp) - 1)
                    Exit For
                End If
            End If
        Next
        UCaseOnly = sTemp
    End If
    
    Exit Function
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function UcaseOnly" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Function

Friend Sub CheckIsDirtyAndAdd(pCheckClaimItems As Collection)
    On Error GoTo EH
    If mbCancel Then
        Exit Sub
    End If
    Dim CheckClaimItem As V2EberlsBillings.udtClaimItem
    Dim vClaimItem As Variant
    Dim ClaimItem As V2EberlsBillings.udtClaimItem
    Dim lCount As Long
    Dim bIsDirty As Boolean
    
    Set mcolValidClaimItems = Nothing
    Set mcolValidClaimItems = New Collection
    
    For lCount = 1 To mcolClaimItems.Count
        bIsDirty = False
        vClaimItem = mcolClaimItems.Item(lCount)
        ClaimItem = vClaimItem
        FindItemInCol pCheckClaimItems, ClaimItem.Main.b03sIB, CheckClaimItem
        
        With CheckClaimItem
            If .SkipThisItem <> ClaimItem.SkipThisItem Then
                bIsDirty = True
            End If
            With .Main
            
                If .b01sCATNO <> ClaimItem.Main.b01sCATNO Then
                    bIsDirty = True
                End If
                
                If .b02sCLAIMNO <> ClaimItem.Main.b02sCLAIMNO Then
                    bIsDirty = True
                End If
                
                If .b03sIB <> ClaimItem.Main.b03sIB Then
                    bIsDirty = True
                End If
                
                If .b04cGROSSLOSS <> ClaimItem.Main.b04cGROSSLOSS Then
                    bIsDirty = True
                End If
                
                If .b05cSERVICEFEE <> ClaimItem.Main.b05cSERVICEFEE Then
                    bIsDirty = True
                End If
                
                If .b06cEXPENSEREIM <> ClaimItem.Main.b06cEXPENSEREIM Then
                    bIsDirty = True
                End If
                
                If .b07cADMINFEE <> ClaimItem.Main.b07cADMINFEE Then
                    bIsDirty = True
                End If
                
                If .b08sFULLNAME <> ClaimItem.Main.b08sFULLNAME Then
                    bIsDirty = True
                End If
                
                If .b09sADJNAME <> ClaimItem.Main.b09sADJNAME Then
                    bIsDirty = True
                End If
                
                If .b10sCLAIMCITY <> ClaimItem.Main.b10sCLAIMCITY Then
                    bIsDirty = True
                End If
                
                If .b11sCLAIMSTATE <> ClaimItem.Main.b11sCLAIMSTATE Then
                    bIsDirty = True
                End If
                
                If .b12dtFILESRECD <> ClaimItem.Main.b12dtFILESRECD Then
                    bIsDirty = True
                End If
                
                If .b13dtDATECLOSED <> ClaimItem.Main.b13dtDATECLOSED Then
                    bIsDirty = True
                End If
                
                If .b14sPROPADDR <> ClaimItem.Main.b14sPROPADDR Then
                    bIsDirty = True
                End If
                
                If .b17sComments <> ClaimItem.Main.b17sComments Then
                    bIsDirty = True
                End If
                
            End With
            
        End With
        
        CheckClaimItem.IsDirty = bIsDirty
        AddClaimItem CheckClaimItem, True
        
    Next
    
    Exit Sub
EH:
    Err.Raise Err.Number, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private sub CheckIsDirtyAndAdd" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Public Function SetGlobalObjects(pcolObjects As Collection) As Boolean
    On Error Resume Next
    Set goUtil = pcolObjects("goUtil")
    Set gofrmMain = pcolObjects("frmMain")
End Function

Public Function OpenConnection() As Boolean
    On Error GoTo EH
    Dim sUserID As String
    Dim sPassword As String
    Dim lUseSQLSERVER
    
    OpenConnection = True
    
    'Get user ID And Pass Word
    sUserID = goUtil.utGetECSCryptSetting(goUtil.gsAppEXEName, "DBConn", "USERID")
    sPassword = goUtil.utGetECSCryptSetting(goUtil.gsAppEXEName, "DBConn", "PASSWORD")
    lUseSQLSERVER = GetSetting(goUtil.gsAppEXEName, "DBConn", "USE_SQLSERVER", vbUnchecked)
    'Check for Use of SQL Server
    If lUseSQLSERVER = vbChecked Then
        mbSQLSERVER = True
        'Check for USER PASS
        If Trim(sUserID) = vbNullString Or Trim(sPassword) = vbNullString Then
INVALID_USER_OR_PASSWORD:
            MsgBox "Invalid User ID or Password!", vbCritical + vbOKOnly, "SQL SERVER"
            OpenConnection = False
            Exit Function
        End If
    Else
        mbSQLSERVER = False
    End If
    
    'ADO Object
    If Not mConn Is Nothing Then
        Set mConn = Nothing
    End If
    
    Set mConn = New ADODB.Connection
    mConn.Open GetSetting(goUtil.gsAppEXEName, "DBConn", "Claims", vbNullString), sUserID, sPassword
           
    Exit Function
EH:
    OpenConnection = False
    If Err.Number = -2147217843 Then
        GoTo INVALID_USER_OR_PASSWORD
        Exit Function
    End If
    Err.Raise Err.Number, , Err.Description & vbCrLf & msClassName & vbCrLf & "Public Function OpenConnection"
End Function

Public Function CloseConnection() As Boolean
    On Error GoTo EH
    CloseConnection = True
    'ADO Object
    If Not mConn Is Nothing Then
        If mConn.State = adStateOpen Then
            mConn.Close
        End If
        Set mConn = Nothing
    End If
    Exit Function
EH:
    CloseConnection = False
    Err.Raise Err.Number, , Err.Description & vbCrLf & msClassName & vbCrLf & "Public Function CloseConnection"
End Function

Private Function FlagSend2Billings(pvClaimItem As Variant) As Boolean

    On Error GoTo EH
    Dim ClaimItem As V2EberlsBillings.udtClaimItem
    Dim sSQL As String
    
    ClaimItem = pvClaimItem
    
    With ClaimItem
        sSQL = "UPDATE BATCHES SET "
        sSQL = sSQL & "BATCHES.COPIED = " & IIf(.SentToBillings, FixClaimPic.AlreadySentToBilling, 0) & ",  "
        sSQL = sSQL & "BATCHES.ENTEREDDATE = CONVERT(DATETIME,('" & Now() & "')) "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "BATCHES.IBNUMBER = '" & goUtil.utCleanSQLString(.Main.b03sIB) & "' "
        sSQL = sSQL & "AND "
        sSQL = sSQL & "BATCHES.CLAIMNUMBER = '" & goUtil.utCleanSQLString(.Main.b02sCLAIMNO) & "' "
        If Not mConn Is Nothing Then
            mConn.Execute sSQL
        Else
            Exit Function
        End If
    End With
    
    FlagSend2Billings = True
    
    Exit Function
EH:
    FlagSend2Billings = False
    Err.Raise Err.Number, , Err.Description & vbCrLf & msClassName & vbCrLf & "Private Function FlagSend2Billings"
End Function

Public Function SendToBillings(psStartDate As String, _
                                psEndDate As String, _
                                psCatSite As String, _
                                psAdjuster As String) As Boolean
    On Error GoTo EH
    Dim ClaimItem As V2EberlsBillings.udtClaimItem
    Dim vClaimItem As Variant
    Dim sSQL As String
    Dim colClaimItems As Collection
    Dim bPB As Boolean
    Dim dtDateClosed As Date
    Dim lCount As Long
    Dim lCopied As Long
    Dim bBillingDup As Boolean
    Dim sMess As String
    
    If mbSQLSERVER Then
        SendToBillings = SendToBillings_SQLSVR(psStartDate, psEndDate, psCatSite, psAdjuster)
        Exit Function
    End If
    
    mbCancel = False
    Set moKB = New clsKeyBoard
    Set moReg = New clsRegSetting
    Set mcolClaimItems = New Collection
    Set mcolValidClaimItems = New Collection
    mdtStart = CDate(psStartDate)
    mdtEnd = CDate(psEndDate)
    
    OpenConnection
    
    If Not ProgBar Is Nothing Then
        ProgBar.Max = 100
        ProgBar.Value = 0
        bPB = True
    End If

    sSQL = "SELECT "
    sSQL = sSQL & "CTOD(BATCHES.DATE) AS ULDATE, "
    sSQL = sSQL & "BATCHES.CATSITE AS CATNO, "
    sSQL = sSQL & "BATCHES.CLAIMNUMBE AS CLAIMNO, "
    sSQL = sSQL & "BATCHES.IBNUMBER AS IB, "
    sSQL = sSQL & "BATCHES.GROSSLOSS, "
    sSQL = sSQL & "BATCHES.TOTALSERVI AS SERVICEFEE, "
    sSQL = sSQL & "BATCHES.MISCCHARGE AS EXPENSEREIM, "
    sSQL = sSQL & "BATCHES.ADMINISTRA AS ADMINFEE, "
    sSQL = sSQL & "BATCHES.INSUREDNAM AS FULLNAME, "
    sSQL = sSQL & "BATCHES.ADJUSTER_N AS ADJNAME, "
    sSQL = sSQL & "BATCHES.LOSSCITY AS CLAIMCITY, "
    sSQL = sSQL & "BATCHES.LOSSSTATE AS CLAIMSTATE, "
    sSQL = sSQL & "BATCHES.DATEOFLOSS AS FILESRECD, "
    sSQL = sSQL & "BATCHES.DATECLOSED, "
    sSQL = sSQL & "BATCHES.LOSS_LOC AS PROPADDR, "
    sSQL = sSQL & "BATCHES.BILLINGDUP, "
    sSQL = sSQL & "BATCHES.COPIED "
    sSQL = sSQL & "FROM BATCHES "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "CTOD(BATCHES.DATE) >= CTOD('" & psStartDate & "') "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "CTOD(BATCHES.DATE) <= CTOD('" & psEndDate & "') "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "BATCHES.ECUPDATED = 1 "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "BATCHES.DUPLICATE = 0 "
    sSQL = sSQL & "ORDER BY ULDATE, IB "

    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient

    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo SKIPBILLINGS
        End If
    Loop

    If Not mRS.EOF Then
        mRS.MoveFirst
        If bPB Then
            ProgBar.Max = mRS.RecordCount
        End If
    Else
        SendToBillings = True 'No errors
        GoTo SKIPBILLINGS
    End If

    Me.StartHwnd = gofrmMain.hWnd

    If Not Me.LookUpLoaded Then
        GoTo SKIPBILLINGS
    End If

    'First load ClaimItem user defined type
    'to be keyed into billings

    Do Until mRS.EOF
        With ClaimItem
            With .Main
                .b01sCATNO = IIf(IsNull(mRS!CATNO), vbNullString, UCase(RTrim(mRS!CATNO)))
                .b02sCLAIMNO = IIf(IsNull(mRS!ClaimNo), vbNullString, UCase(RTrim(mRS!ClaimNo)))
                .b03sIB = IIf(IsNull(mRS!IB), vbNullString, UCase(RTrim(mRS!IB)))
                .b04cGROSSLOSS = IIf(IsNull(mRS!grossloss), 0, CCur(mRS!grossloss))
                .b05cSERVICEFEE = IIf(IsNull(mRS!SERVICEFEE), 0, CCur(mRS!SERVICEFEE))
                .b06cEXPENSEREIM = IIf(IsNull(mRS!EXPENSEREIM), 0, CCur(mRS!EXPENSEREIM))
                .b07cADMINFEE = IIf(IsNull(mRS!ADMINFEE), 0, CCur(mRS!ADMINFEE))
                .b08sFULLNAME = IIf(IsNull(mRS!FULLNAME), vbNullString, UCase(RTrim(mRS!FULLNAME)))
                .b09sADJNAME = IIf(IsNull(mRS!ADJNAME), vbNullString, UCase(RTrim(mRS!ADJNAME)))
                .b10sCLAIMCITY = IIf(IsNull(mRS!CLAIMCITY), vbNullString, UCase(Replace(RTrim(mRS!CLAIMCITY), FailedValidFlag, vbNullString, , , vbTextCompare)))
                .b11sCLAIMSTATE = IIf(IsNull(mRS!CLAIMSTATE), vbNullString, UCase(RTrim(mRS!CLAIMSTATE)))
                .b12dtFILESRECD = ValidDateField(mRS!FILESRECD, mRS!uldate)
                .b13dtDATECLOSED = ValidDateField(mRS!dateclosed, mRS!uldate)
                .b14sPROPADDR = IIf(IsNull(mRS!PROPADDR), vbNullString, RTrim(mRS!PROPADDR))
                .b14sPROPADDR = UCase(Replace(.b14sPROPADDR, S_z, " ", , , vbBinaryCompare))
            End With
            
            If IsNull(mRS!copied) Then
                lCopied = 0
            Else
                lCopied = mRS!copied
            End If
            
            Select Case lCopied
                Case 0, FixClaimPic.NeedsSentToBilling
                    .SentToBillings = False
                    .SkipThisItem = False
                Case FixClaimPic.AlreadySentToBilling
                    .SentToBillings = True
                    .SkipThisItem = False
                Case FixClaimPic.BadData

                Case FixClaimPic.GoodData
                
                Case FixClaimPic.Skip
                    .SkipThisItem = True
                    .SentToBillings = False
                Case FixClaimPic.PointAt
            End Select
            
            'Check for Billing Duplicate.  This will be true if
            'the bill that was uploaded was already sent via Snail Mail
            'and manually entered. If this happens need to put a Stop flag for that
            'Item
            If IsNull(mRS!BillingDup) Then
                bBillingDup = False
            Else
                bBillingDup = CBool(mRS!BillingDup)
            End If
            
            If bBillingDup Then
                .SkipThisItem = True
                .BillingDup = True
            Else
                .BillingDup = False
            End If
            
            If InStr(1, .Main.b02sCLAIMNO, "TEST", vbTextCompare) > 0 Then
                .SkipThisItem = True
            End If
           
        End With
        
        If Not gofrmMain.UnloadFlag Then
            Me.AddClaimItem ClaimItem
        Else
            GoTo SKIPBILLINGS
        End If
        If bPB Then
            ProgBar.Value = mRS.AbsolutePosition
        End If
        mRS.MoveNext
    Loop

    Set mRS = Nothing
    
    If bPB Then
        ProgBar.Value = 0
    End If
    
    If gofrmMain.UnloadFlag Then
        GoTo SKIPBILLINGS
    End If

    'after ValidateClaimItems Passes then update Batches Table with
    'any changes made during validation.

    If Me.ValidateClaimItems() Then
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo SKIPBILLINGS
        End If
        
        gofrmMain.lblWarning.Caption = "Updating Changes.  Please Wait..."
        DoEvents
        Sleep 10
        'BGS 2/1/2002 need to check for any dirty Items
        'means they were changed due to the validation process
        'and update Batches table with new data.
        Set colClaimItems = Me.ClaimItems
        If bPB Then
            ProgBar.Max = colClaimItems.Count
        End If
        For Each vClaimItem In colClaimItems
            ClaimItem = vClaimItem
            If ClaimItem.IsDirty Then
                With ClaimItem
                    sSQL = "UPDATE BATCHES "
                    sSQL = sSQL & "SET "
                    'Check the Icon
                    If .SkipThisItem Then
                        sSQL = sSQL & "BATCHES.COPIED = " & FixClaimPic.Skip & ", "
                    ElseIf .SentToBillings Then
                        sSQL = sSQL & "BATCHES.COPIED = " & FixClaimPic.AlreadySentToBilling & ", "
                    Else
                        sSQL = sSQL & "BATCHES.COPIED = " & FixClaimPic.NeedsSentToBilling & ", "
                    End If
                    sSQL = sSQL & "BATCHES.CATSITE = '" & goUtil.utCleanSQLString(.Main.b01sCATNO) & "', "
                    sSQL = sSQL & "BATCHES.CLAIMNUMBE = '" & goUtil.utCleanSQLString(.Main.b02sCLAIMNO) & "', "
                    sSQL = sSQL & "BATCHES.GROSSLOSS = " & .Main.b04cGROSSLOSS & ", "
                    sSQL = sSQL & "BATCHES.TOTALSERVI = " & .Main.b05cSERVICEFEE & ", "
                    sSQL = sSQL & "BATCHES.MISCCHARGE = " & .Main.b06cEXPENSEREIM & ", "
                    sSQL = sSQL & "BATCHES.ADMINISTRA = " & .Main.b07cADMINFEE & ", "
                    sSQL = sSQL & "BATCHES.INSUREDNAM = '" & goUtil.utCleanSQLString(.Main.b08sFULLNAME) & "', "
                    sSQL = sSQL & "BATCHES.ADJUSTER_N = '" & goUtil.utCleanSQLString(.Main.b09sADJNAME) & "', "
                    sSQL = sSQL & "BATCHES.LOSSCITY = '" & goUtil.utCleanSQLString(.Main.b10sCLAIMCITY) & "', "
                    sSQL = sSQL & "BATCHES.LOSSSTATE = '" & goUtil.utCleanSQLString(.Main.b11sCLAIMSTATE) & "', "
                    sSQL = sSQL & "BATCHES.DATEOFLOSS = CTOD('" & .Main.b12dtFILESRECD & "'), "
                    sSQL = sSQL & "BATCHES.DATECLOSED = CTOD('" & .Main.b13dtDATECLOSED & "'), "
                    sSQL = sSQL & "BATCHES.LOSS_LOC = '" & goUtil.utCleanSQLString(.Main.b14sPROPADDR) & "' "
                    sSQL = sSQL & "WHERE "
                    sSQL = sSQL & "BATCHES.IBNUMBER = '" & goUtil.utCleanSQLString(.Main.b03sIB) & "' "
                    mConn.Execute sSQL
                End With
            End If
            If bPB Then
                lCount = lCount + 1
                ProgBar.Value = lCount
            End If
            DoEvents
            Sleep 10
            If gofrmMain.UnloadFlag Then
                GoTo SKIPBILLINGS
            End If
        Next
        
        If bPB Then
            ProgBar.Value = 0
        End If
        
        If gofrmMain.UnloadFlag Then
            GoTo SKIPBILLINGS
        End If
        
        gofrmMain.lblWarning.Caption = vbNullString
        DoEvents
        Sleep 10

        'If we are not skiping all then we need to SendToXactimate
        'SendToXact will launch Xactimate if it isn't already loaded and set focus to it
        If Not Me.SkipAll Then
            gofrmMain.lblWarning.Caption = "Sending Keys to Billings.  Please Stand By..."
            Screen.MousePointer = vbHourglass
            DoEvents
            Sleep 10
            
            'Reset The Sent and Skipped variables
            mlItemsSent = 0
            mlItemsSkipped = 0
            
            If Me.SendKeysToBillings() Then
                'Check here incase user tries to unload in middle of sending
                If gofrmMain.UnloadFlag Then
                    GoTo SKIPBILLINGS
                End If

                Me.LookForWindow Left(gofrmMain.Caption, 10)
                sMess = "Item(s) Sent to Billings!" & vbCrLf & vbCrLf
                sMess = sMess & "ITEM(S) SENT = " & mlItemsSent & vbCrLf
                sMess = sMess & "ITEM(S) SKIPPED = " & mlItemsSkipped
                MsgBox sMess, vbInformation + vbOKOnly, "Send To Billings"
            Else
                'OK we need a check here incase user tries to unload in middle of sending
                If gofrmMain.UnloadFlag Then
                    GoTo SKIPBILLINGS
                End If

                Me.LookForWindow Left(gofrmMain.Caption, 10)
                sMess = "Item(s) NOT Sent to Billings!" & vbCrLf & vbCrLf
                sMess = sMess & "ITEM(S) SENT = " & mlItemsSent & vbCrLf
                sMess = sMess & "ITEM(S) SKIPPED = " & mlItemsSkipped
                MsgBox sMess, vbInformation + vbOKOnly, "Send To Billings"
            End If
        End If
    End If
    SendToBillings = True
    Me.CLEANUP
SKIPBILLINGS:
    Set colClaimItems = Nothing
    Set mRS = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
EH:
    SendToBillings = False
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, , Err.Description & vbCrLf & msClassName & vbCrLf & "Public Function SendToBillings"
End Function

Public Function SendToBillings_SQLSVR(psStartDate As String, _
                                        psEndDate As String, _
                                        psCatSite As String, _
                                        psAdjuster As String) As Boolean
    On Error GoTo EH
    Dim ClaimItem As V2EberlsBillings.udtClaimItem
    Dim vClaimItem As Variant
    Dim sSQL As String
    Dim colClaimItems As Collection
    Dim bPB As Boolean
    Dim dtDateClosed As Date
    Dim lCount As Long
    Dim lCopied As Long
    Dim bBillingDup As Boolean
    Dim sMess As String
    '4.19.2003 allow to weed out what we see
    Dim sSSN As String
    Dim lADJSSN As Long
    Dim sCatSite As String
    Dim bByAllCatSites As Boolean
    Dim bByAllAdjusters As Boolean
    Dim bHideSkippedOn As Boolean
    Dim bHideSnailMailOn As Boolean
    Dim bHideSendOn As Boolean
    Dim bHideSentOn As Boolean
    
    bHideSkippedOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESKIPPED_ON", False))
    bHideSnailMailOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESNAILMAIL_ON", False))
    bHideSendOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESEND_ON", False))
    bHideSentOn = CBool(GetSetting(goUtil.gsAppEXEName, "GENERAL", "HIDESENT_ON", False))
    
    ' 4.19.2003 Parse the ssn from the Adjuster
    sSSN = Trim(psAdjuster)
    If StrComp(sSSN, "(--By All Adjusters--)", vbTextCompare) = 0 Then
        bByAllAdjusters = True
    Else
        sSSN = Mid(sSSN, InStr(1, sSSN, "(", vbBinaryCompare) + 1)
        sSSN = Replace(sSSN, ")", vbNullString, , , vbBinaryCompare)
        lADJSSN = sSSN
    End If
    
    'set the cat site
    sCatSite = Trim(psCatSite)
    If StrComp(sCatSite, "(--By All Cat Sites--)", vbTextCompare) = 0 Then
        bByAllCatSites = True
    Else
        'Use adAsyncExecute so the user can do other stuff while the Lookup boxes are loading.
        Set mRS = New ADODB.Recordset
        OpenConnection
        mRS.CursorLocation = adUseClient
        PopulateADJNAME_SQLSVR True, sCatSite
        Set mRS = Nothing
    End If
    
    
    mbCancel = False
    Set moKB = New clsKeyBoard
    Set moReg = New clsRegSetting
    Set mcolClaimItems = New Collection
    Set mcolValidClaimItems = New Collection
    mdtStart = CDate(psStartDate)
    mdtEnd = CDate(psEndDate)
    
    OpenConnection
    
    If Not ProgBar Is Nothing Then
        ProgBar.Max = 100
        ProgBar.Value = 0
        bPB = True
    End If

    sSQL = "SELECT "
    sSQL = sSQL & "CONVERT(DATETIME,(BATCHES.[date])) AS ULDATE, "
    sSQL = sSQL & "BATCHES.[EnteredDate], "
    sSQL = sSQL & "BATCHES.[catsite] AS CATNO, "
    sSQL = sSQL & "BATCHES.[claimnumber] AS CLAIMNO, "
    sSQL = sSQL & "BATCHES.[ibnumber] AS IB, "
    sSQL = sSQL & "BATCHES.[grossloss], "
    sSQL = sSQL & "BATCHES.[totalservice] AS SERVICEFEE, "
    sSQL = sSQL & "BATCHES.[misccharge] AS EXPENSEREIM, "
    sSQL = sSQL & "BATCHES.[administrative] AS ADMINFEE, "
    sSQL = sSQL & "BATCHES.[insuredname] AS FULLNAME, "
    sSQL = sSQL & "BATCHES.[adjuster_n] AS ADJNAME, "
    sSQL = sSQL & "BATCHES.[losscity] AS CLAIMCITY, "
    sSQL = sSQL & "BATCHES.[lossstate] AS CLAIMSTATE, "
    sSQL = sSQL & "BATCHES.[dateofloss] AS FILESRECD, "
    sSQL = sSQL & "BATCHES.[dateclosed], "
    sSQL = sSQL & "BATCHES.[loss_loc] AS PROPADDR, "
    sSQL = sSQL & "BATCHES.[billingdup], "
    sSQL = sSQL & "BATCHES.[copied], "
    sSQL = sSQL & "BATCHES.[Comments], "
    sSQL = sSQL & "BATCHES.[reassigned] "
    sSQL = sSQL & "FROM BATCHES "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & "CONVERT(DATETIME,(BATCHES.[date])) >= CONVERT(DATETIME,('" & psStartDate & "')) "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "CONVERT(DATETIME,(BATCHES.[date])) <= CONVERT(DATETIME,('" & psEndDate & "')) "
    If Not bByAllCatSites Then
        sSQL = sSQL & "AND "
        sSQL = sSQL & "BATCHES.[catsite] ='" & sCatSite & "' "
    End If
    If Not bByAllAdjusters Then
        sSQL = sSQL & "AND "
        sSQL = sSQL & "BATCHES.[ssn] =" & lADJSSN & " "
    End If
    If bHideSkippedOn Then
        sSQL = sSQL & "AND "
        sSQL = sSQL & "BATCHES.[copied] <> " & FixClaimPic.Skip & " "
    End If
    If bHideSnailMailOn Then
        sSQL = sSQL & "AND "
        sSQL = sSQL & "BATCHES.[billingdup] <> 1 "
    End If
    If bHideSendOn Then
        sSQL = sSQL & "AND "
        sSQL = sSQL & "BATCHES.[EnteredDate] Is Not Null "
    End If
    If bHideSentOn Then
        sSQL = sSQL & "AND "
        sSQL = sSQL & "BATCHES.[EnteredDate] Is Null "
    End If
    sSQL = sSQL & "AND "
    sSQL = sSQL & "BATCHES.[ecupdated] = 1 "
    sSQL = sSQL & "AND "
    sSQL = sSQL & "BATCHES.[duplicate] = 0 "
    sSQL = sSQL & "ORDER BY IB DESC, ULDATE "

    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient

    mRS.Open sSQL, mConn, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until mRS.State = adStateOpen
        If bPB Then
            If ProgBar.Value + 1 < ProgBar.Max Then
                ProgBar.Value = ProgBar.Value + 1
            End If
        End If
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo SKIPBILLINGS
        End If
    Loop

    If Not mRS.EOF Then
        mRS.MoveFirst
        If bPB Then
            ProgBar.Max = mRS.RecordCount
        End If
    Else
        SendToBillings_SQLSVR = True 'No errors
        MsgBox "No Records Found.", vbInformation + vbOKOnly, "Send To Billings"
        gofrmMain.lblWarning.Caption = "No Records Found."
        ProgBar.Value = 0
        GoTo SKIPBILLINGS
    End If

    Me.StartHwnd = gofrmMain.hWnd

    If Not Me.LookUpLoaded Then
        GoTo SKIPBILLINGS
    End If

    'First load ClaimItem user defined type
    'to be keyed into billings

    Do Until mRS.EOF
        With ClaimItem
            With .Main
                .b01sCATNO = IIf(IsNull(mRS!CATNO), vbNullString, UCase(RTrim(mRS!CATNO)))
                .b02sCLAIMNO = IIf(IsNull(mRS!ClaimNo), vbNullString, UCase(RTrim(mRS!ClaimNo)))
                .b03sIB = IIf(IsNull(mRS!IB), vbNullString, UCase(RTrim(mRS!IB)))
                .b04cGROSSLOSS = IIf(IsNull(mRS!grossloss), 0, CCur(mRS!grossloss))
                .b05cSERVICEFEE = IIf(IsNull(mRS!SERVICEFEE), 0, CCur(mRS!SERVICEFEE))
                .b06cEXPENSEREIM = IIf(IsNull(mRS!EXPENSEREIM), 0, CCur(mRS!EXPENSEREIM))
                .b07cADMINFEE = IIf(IsNull(mRS!ADMINFEE), 0, CCur(mRS!ADMINFEE))
                .b08sFULLNAME = IIf(IsNull(mRS!FULLNAME), vbNullString, UCase(RTrim(mRS!FULLNAME)))
                .b09sADJNAME = IIf(IsNull(mRS!ADJNAME), vbNullString, UCase(RTrim(mRS!ADJNAME)))
                .b10sCLAIMCITY = IIf(IsNull(mRS!CLAIMCITY), vbNullString, UCase(Replace(RTrim(mRS!CLAIMCITY), FailedValidFlag, vbNullString, , , vbTextCompare)))
                .b11sCLAIMSTATE = IIf(IsNull(mRS!CLAIMSTATE), vbNullString, UCase(RTrim(mRS!CLAIMSTATE)))
                .b12dtFILESRECD = ValidDateField(mRS!FILESRECD, mRS!uldate)
                .b13dtDATECLOSED = ValidDateField(mRS!dateclosed, mRS!uldate)
                .b14sPROPADDR = IIf(IsNull(mRS!PROPADDR), vbNullString, RTrim(mRS!PROPADDR))
                .b14sPROPADDR = UCase(Replace(.b14sPROPADDR, S_z, " ", , , vbBinaryCompare))
                .b15dtDateUploaded = IIf(IsNull(mRS!uldate), Format(NULL_DATE, "mm/dd/yyyy"), Format(mRS!uldate, "mm/dd/yyyy"))
                .b16dtDateEntered = IIf(IsNull(mRS!EnteredDate), Format(NULL_DATE, "mm/dd/yyyy"), Format(mRS!EnteredDate, "mm/dd/yyyy"))
                .b17sComments = IIf(IsNull(mRS!Comments), vbNullString, mRS!Comments)
            End With
            
            If IsNull(mRS!copied) Then
                lCopied = 0
            Else
                lCopied = mRS!copied
            End If
            
            Select Case lCopied
                Case 0, FixClaimPic.NeedsSentToBilling
                    'Also check the Entered date
                    If Not IsNull(mRS!EnteredDate) Then
                        .SentToBillings = True
                    Else
                        .SentToBillings = False
                    End If
                    .SkipThisItem = False
                Case FixClaimPic.AlreadySentToBilling
                    .SentToBillings = True
                    .SkipThisItem = False
                Case FixClaimPic.BadData

                Case FixClaimPic.GoodData
                
                Case FixClaimPic.Skip
                    .SkipThisItem = True
                    .SentToBillings = False
                Case FixClaimPic.PointAt
            End Select
            
            'Check for Billing Duplicate.  This will be true if
            'the bill that was uploaded was already sent via Snail Mail
            'and manually entered. If this happens need to put a Stop flag for that
            'Item
            If IsNull(mRS!BillingDup) Then
                bBillingDup = False
            Else
                bBillingDup = CBool(mRS!BillingDup)
            End If
            
            If bBillingDup Then
                .SkipThisItem = True
                .BillingDup = True
            Else
                .BillingDup = False
            End If
            
             '6.4.2003 Check for Reassigned flag
            'But allow admin to mark this item as skipped.
            If mRS!Reassigned > 0 Then
                .BillingDup = True
            End If
            
            If InStr(1, .Main.b02sCLAIMNO, "TEST", vbTextCompare) > 0 Then
                .SkipThisItem = True
            End If
           
        End With
        
        If Not gofrmMain.UnloadFlag Then
            Me.AddClaimItem ClaimItem
        Else
            GoTo SKIPBILLINGS
        End If
        If bPB Then
            ProgBar.Value = mRS.AbsolutePosition
        End If
        mRS.MoveNext
    Loop

    Set mRS = Nothing
    
    If bPB Then
        ProgBar.Value = 0
    End If
    
    If gofrmMain.UnloadFlag Then
        GoTo SKIPBILLINGS
    End If

    'after ValidateClaimItems Passes then update Batches Table with
    'any changes made during validation.

    If Me.ValidateClaimItems() Then
        DoEvents
        Sleep 10
        If gofrmMain.UnloadFlag Then
            GoTo SKIPBILLINGS
        End If
        
        gofrmMain.lblWarning.Caption = "Updating Changes.  Please Wait..."
        DoEvents
        Sleep 10
        'BGS 2/1/2002 need to check for any dirty Items
        'means they were changed due to the validation process
        'and update Batches table with new data.
        Set colClaimItems = Me.ClaimItems
        If bPB Then
            ProgBar.Max = colClaimItems.Count
        End If
        For Each vClaimItem In colClaimItems
            ClaimItem = vClaimItem
            If ClaimItem.IsDirty Then
                With ClaimItem
                    sSQL = "UPDATE BATCHES "
                    sSQL = sSQL & "SET "
                    'Check the Icon
                    If .SkipThisItem Then
                        sSQL = sSQL & "BATCHES.[copied] = " & FixClaimPic.Skip & ", "
                    ElseIf .SentToBillings Then
                        sSQL = sSQL & "BATCHES.[copied] = " & FixClaimPic.AlreadySentToBilling & ", "
                    Else
                        sSQL = sSQL & "BATCHES.[copied] = " & FixClaimPic.NeedsSentToBilling & ", "
                    End If
                    sSQL = sSQL & "BATCHES.[catsite] = '" & goUtil.utCleanSQLString(.Main.b01sCATNO) & "', "
                    sSQL = sSQL & "BATCHES.[claimnumber] = '" & goUtil.utCleanSQLString(.Main.b02sCLAIMNO) & "', "
                    sSQL = sSQL & "BATCHES.[grossloss] = " & .Main.b04cGROSSLOSS & ", "
                    sSQL = sSQL & "BATCHES.[totalservice] = " & .Main.b05cSERVICEFEE & ", "
                    sSQL = sSQL & "BATCHES.[misccharge] = " & .Main.b06cEXPENSEREIM & ", "
                    sSQL = sSQL & "BATCHES.[administrative] = " & .Main.b07cADMINFEE & ", "
                    sSQL = sSQL & "BATCHES.[insuredname] = '" & goUtil.utCleanSQLString(.Main.b08sFULLNAME) & "', "
                    sSQL = sSQL & "BATCHES.[adjuster_n] = '" & goUtil.utCleanSQLString(.Main.b09sADJNAME) & "', "
                    sSQL = sSQL & "BATCHES.[losscity] = '" & goUtil.utCleanSQLString(.Main.b10sCLAIMCITY) & "', "
                    sSQL = sSQL & "BATCHES.[lossstate] = '" & goUtil.utCleanSQLString(.Main.b11sCLAIMSTATE) & "', "
                    sSQL = sSQL & "BATCHES.[dateofloss] = CONVERT(DATETIME,('" & .Main.b12dtFILESRECD & "')), "
                    sSQL = sSQL & "BATCHES.[dateclosed] = CONVERT(DATETIME,('" & .Main.b13dtDATECLOSED & "')), "
                    sSQL = sSQL & "BATCHES.[loss_loc] = '" & goUtil.utCleanSQLString(.Main.b14sPROPADDR) & "', "
                    sSQL = sSQL & "BATCHES.[Comments] = '" & goUtil.utCleanSQLString(.Main.b17sComments) & "' "
                    sSQL = sSQL & "WHERE "
                    sSQL = sSQL & "BATCHES.[ibnumber] = '" & goUtil.utCleanSQLString(.Main.b03sIB) & "' "
                    mConn.Execute sSQL
                End With
            End If
            If bPB Then
                lCount = lCount + 1
                ProgBar.Value = lCount
            End If
            DoEvents
            Sleep 10
            If gofrmMain.UnloadFlag Then
                GoTo SKIPBILLINGS
            End If
        Next
        
        If bPB Then
            ProgBar.Value = 0
        End If
        
        If gofrmMain.UnloadFlag Then
            GoTo SKIPBILLINGS
        End If
        
        gofrmMain.lblWarning.Caption = vbNullString
        DoEvents
        Sleep 10

        'If we are not skiping all then we need to SendToXactimate
        'SendToXact will launch Xactimate if it isn't already loaded and set focus to it
        If Not Me.SkipAll Then
            gofrmMain.lblWarning.Caption = "Sending Keys to Billings.  Please Stand By..."
            Screen.MousePointer = vbHourglass
            DoEvents
            Sleep 10
            
            'Reset The Sent and Skipped variables
            mlItemsSent = 0
            mlItemsSkipped = 0
            
            If Me.SendKeysToBillings() Then
                'Check here incase user tries to unload in middle of sending
                If gofrmMain.UnloadFlag Then
                    GoTo SKIPBILLINGS
                End If

                Me.LookForWindow Left(gofrmMain.Caption, 10)
                sMess = "Item(s) Sent to Billings!" & vbCrLf & vbCrLf
                sMess = sMess & "ITEM(S) SENT = " & mlItemsSent & vbCrLf
                sMess = sMess & "ITEM(S) SKIPPED = " & mlItemsSkipped
                MsgBox sMess, vbInformation + vbOKOnly, "Send To Billings"
            Else
                'OK we need a check here incase user tries to unload in middle of sending
                If gofrmMain.UnloadFlag Then
                    GoTo SKIPBILLINGS
                End If

                Me.LookForWindow Left(gofrmMain.Caption, 10)
                sMess = "Item(s) NOT Sent to Billings!" & vbCrLf & vbCrLf
                sMess = sMess & "ITEM(S) SENT = " & mlItemsSent & vbCrLf
                sMess = sMess & "ITEM(S) SKIPPED = " & mlItemsSkipped
                MsgBox sMess, vbInformation + vbOKOnly, "Send To Billings"
            End If
        End If
    End If
    SendToBillings_SQLSVR = True
    Me.CLEANUP
SKIPBILLINGS:
    Set colClaimItems = Nothing
    Set mRS = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
EH:
    SendToBillings_SQLSVR = False
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, , Err.Description & vbCrLf & msClassName & vbCrLf & "Public Function SendToBillings_SQLSVR"
End Function


Public Function ADDItemToCol(pCol As Collection, pvItem As Variant, psKey As String) As Boolean
    On Error GoTo EH
    
    pCol.Add pvItem, RTrim(psKey)
    
    ADDItemToCol = True
    Exit Function
EH:
    Err.Clear
End Function

Public Function FindItemInCol(pCol As Collection, psKey As String, Optional pvItem As Variant) As Boolean
    On Error GoTo EH
    
    pvItem = pCol(psKey)
    
    FindItemInCol = True
    Exit Function
EH:
    Err.Clear
End Function

Public Function ValidDateField(pdtField As ADODB.Field, Optional pdtDefault As ADODB.Field) As Date
    On Error GoTo EH
    
    
    If IsNull(pdtField.Value) Then
        If Not pdtDefault Is Nothing Then
            If IsNull(pdtDefault.Value) Then
                ValidDateField = Format(Now(), "MM/DD/YYYY")
            Else
                If IsDate(pdtDefault.Value) Then
                    If CDate(pdtDefault.Value) = NULL_DATE Then
                        ValidDateField = Format(Now(), "MM/DD/YYYY")
                    Else
                        ValidDateField = Format(pdtDefault.Value, "mm/dd/yyyy")
                    End If
                Else
                    ValidDateField = Format(Now(), "MM/DD/YYYY")
                End If
            End If
        Else
            ValidDateField = Format(Now(), "MM/DD/YYYY")
        End If
    Else
        If IsDate(pdtField.Value) Then
            If CDate(pdtField.Value) = NULL_DATE Then
                ValidDateField = NULL_DATE
            Else
                ValidDateField = Format(pdtField.Value, "mm/dd/yyyy")
            End If
        Else
            ValidDateField = Format(Now(), "MM/DD/YYYY")
        End If
    End If
    
    Exit Function
EH:
    Err.Raise Err.Number, , Err.Description & vbCrLf & msClassName & vbCrLf & "Public Function ValidDateField"
End Function
