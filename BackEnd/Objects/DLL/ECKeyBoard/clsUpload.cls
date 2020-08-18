VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Error Event
Public Event ErrorMess(ByVal Mess As String)

'Update WebControl or Easy Claim Database
Public Event UpdateDB(ByVal vBatches As Variant)

'Update WebControl Database Real Time Tables
Public Event UpdateDBRT(ByVal vBatches As Variant, poBatch As V2ECKeyBoard.clsBatches, poULRT As V2ECKeyBoard.clsCarUL, poUL As V2ECKeyBoard.clsUpload)

Private moULCar As V2ECKeyBoard.clsCarUL
Private mcolFields As Collection
Private moParam As V2ECKeyBoard.clsLossReports
'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
Private msAPPEXEName As String 'THIS IDENTIFIES CARRRIER OBJECT !! IE FARMERS, AMFAM, ETC. ETC.
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Property Let Param(poParam As V2ECKeyBoard.clsLossReports)
    Set moParam = poParam
End Property
Public Property Set Param(poParam As V2ECKeyBoard.clsLossReports)
    Set moParam = poParam
End Property
Public Property Get Param() As V2ECKeyBoard.clsLossReports
    Set Param = moParam
End Property

Public Property Let APPEXEName(psAppEXEName As String)
    msAPPEXEName = psAppEXEName
End Property
Public Property Get APPEXEName() As String
    APPEXEName = msAPPEXEName
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsUpLoad"
End Property

Public Property Let Carrier(poCarrier As V2ECKeyBoard.clsCarUL)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set moULCar = poCarrier

    Set mcolFields = New Collection
    Exit Property
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Property Let Carrier"
End Property

Public Property Set Carrier(poCarrier As V2ECKeyBoard.clsCarUL)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set moULCar = poCarrier
    
    Set mcolFields = New Collection
    
    Exit Property
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Property Set Carrier"
End Property

Public Property Get Fields() As Collection
    Set Fields = mcolFields
End Property

Public Sub AddField(pvValue As Variant, psFieldName As String, _
                    Optional piLen As Integer = -1, _
                    Optional piForceDataType As VbVarType = Empty)
    On Error GoTo EH
    Dim oField As clsField
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set oField = New clsField
    With oField
        'Check for
        If piForceDataType = 0 Then
            .DataType = VarType(pvValue)
        Else
            .DataType = piForceDataType
        End If
        .FieldName = psFieldName
        .Length = piLen
        .Value = pvValue
    End With
    
    mcolFields.Add oField, oField.FieldName
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub AddField"
End Sub

Private Function CreateUpLoad(psUploadDir As String, psClaimNo As String, psCheckNum As String, pCurDB As Object, _
                              Optional psRT As String = vbNullString, _
                              Optional psSort As String = vbNullString) As Boolean
    On Error GoTo EH
    Dim saryUL() As String
    Dim oField As clsField
    Dim iCount As Integer
    Dim sFilename As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    
    CreateUpLoad = True
    
    'BGS 10.11.2001 check to see if the Class has been created first
    'If so then we can populate fields into mcolFields with whichever
    'clsCarUL carrier we are dealing with.
    If Not moULCar Is Nothing Then
        moULCar.PopulateFields psClaimNo, psCheckNum, pCurDB, Me
    Else
        CreateUpLoad = False
        Exit Function
    End If
    
    'BGS 10.11.2001 Dimension appropriate fields
    ReDim saryUL(0 To mcolFields.Count)
    saryUL(0) = moULCar.ClassName

    
    'BGS 10.11.2001 Create the File name
    If moULCar.ULType = ClosedIB Then
        sFilename = psSort & moULCar.IBNumber & "_" & Format(Now(), "mmddyyhhmmss") & ".OUT"
    ElseIf moULCar.ULType = RealTime Then
        sFilename = psSort & moULCar.IBNumber & "_" & psRT & Format(Now(), "mmddyyhhmmss") & ".OUT"
    End If
    
    For Each oField In mcolFields
        iCount = iCount + 1
        With oField
            saryUL(iCount) = .DataType & F_DELIM & .FieldName & F_DELIM & Replace(.Value, vbCrLf, F_VBCRLF) & F_DELIM
        End With
    Next
    'BGS Now that we have populated the saryUL with all the field Data
    'need to join this into the out put file
    
    goUtil.utSaveFileData psUploadDir & "\" & sFilename, Join(saryUL, vbCrLf)
    
    'BGS 10.17.2001 Clean UP
    Set mcolFields = Nothing
    Set mcolFields = New Collection
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CreateUpLoad = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function CreateUpLoad"
End Function

Public Function CreateUploadZOT(psUploadDir As String, pCurDB As Object, psAppEXEName As String, _
                                poPBar As Object, _
                                psPLabelText As String, _
                                poPList As Object) As String
    On Error GoTo EH
    Dim sSQL As String
    Dim RS As dao.Recordset
    Dim iCount As Integer
    Dim sZOTName As String
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'BGS need to get all the billing info that has not already been uploaded
    sSQL = "SELECT A.ClaimNo, A.CheckNumIB, A.UpLoadMe, A.Status "
    sSQL = sSQL & "FROM Billing A "
    sSQL = sSQL & "WHERE A.UpLoadMe = True "
    sSQL = sSQL & "ORDER BY A.ID "
    
    Set RS = pCurDB.OpenRecordset(sSQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        poPBar.Max = RS.RecordCount
        poPBar.Value = 0
        Do Until RS.EOF
            psPLabelText = "Creating File for (" & RS!ClaimNo & " - " & RS!CheckNumIB & ")"
            poPList.AddItem psPLabelText
            If Not CreateUpLoad(psUploadDir, RS!ClaimNo, RS!CheckNumIB, pCurDB) Then
                CreateUploadZOT = vbNullString
                Exit Function
            Else
                'BGS 11.15.2001 once we have created the upload file
                'we can Reset the Upload flag
                RS.Edit
                RS!UploadMe = False
                RS!Status = "SENT"
                RS.Update
                
                'BGS 11.15.2001 Once we uploaded the file need to change the
                'Status on the assignments table
                sSQL = "UPDATE Assignments SET "
                sSQL = sSQL & "Assignments.Status = 'SENT' "
                sSQL = sSQL & "WHERE Assignments.ClaimNo = '" & goUtil.utCleanSQLString(RS!ClaimNo) & "' "
                
                pCurDB.Execute sSQL
            End If
            poPBar.Value = poPBar.Value + 1
            
            RS.MoveNext
        Loop
    Else
        CreateUploadZOT = vbNullString
        Exit Function
    End If
    
    Set RS = Nothing
    
    'BGS 11.15.2001 Now need to create the ZOT file
    'This will put all the .OUT filesin the upload dir into 1 ZOT file (ZIPED with Xceed ZIP)
    'And remove all the .OUT files from the upload dir.
    
    sZOTName = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME", "UNKNOWN_ID")
    sZOTName = sZOTName & "_" & Format(Now, "MMDDYYHHMMSS") & ".ZOT"
    Set oXZip = New V2ECKeyBoard.clsXZip
    
    If oXZip.SaveZIPFiles(psUploadDir, sZOTName, "*.OUT", goUtil.DB_PASSWORD("1")) Then
        CreateUploadZOT = sZOTName
    End If
    
    Set oXZip = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CreateUploadZOT"
End Function

Public Function CreateUploadZOTRT(psUploadDir As String, pCurDB As Object, psAppEXEName As String, _
                                poPBar As Object, _
                                poPLabel As Object, _
                                poPList As Object) As String
    On Error GoTo EH
    Dim sSQL As String
    Dim RS As Recordset
    Dim iCount As Integer
    Dim sZOTName As String
    Dim sSort As String
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'BGS need to get all the RT Assignment infor that has not yet been uploaded
    'We will upload even if the claim is closed.  If an adjuster enters the
    'claim and changes photos or act log on closed claim then we need to upload it.
    
    sSQL = "SELECT A.ClaimNo, B.CheckNum, A.RTUpLoadMe, A.Status "
    sSQL = sSQL & "FROM Assignments As A "
    sSQL = sSQL & "LEFT JOIN Checks As B "
    sSQL = sSQL & "ON A.ClaimNo = B.ClaimNo "
    sSQL = sSQL & "WHERE A.RTUpLoadMe = True "
    sSQL = sSQL & "AND B.CheckNum Is Not Null "
    sSQL = sSQL & "ORDER BY B.ID "
    
    Set RS = pCurDB.OpenRecordset(sSQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        poPBar.Max = RS.RecordCount
        poPBar.Value = 0
        Do Until RS.EOF
            poPLabel.Caption = "Creating RT File for (" & RS!ClaimNo & " - " & RS!CheckNum & ")"
            poPList.AddItem poPLabel.Caption
            sSort = Format(iCount, "000")
            iCount = iCount + 1
            If Not CreateUpLoad(psUploadDir, RS!ClaimNo, RS!CheckNum, pCurDB, "@" & RS!CheckNum, "ZZ" & sSort & "_") Then
                CreateUploadZOTRT = vbNullString
                Exit Function
            Else
                'BGS 11.15.2001 once we have created the upload file
                'we can reset the Upload flag.
                RS.Edit
                RS!RTUploadMe = False
                RS.Update
            End If
            poPBar.Value = poPBar.Value + 1
            
            RS.MoveNext
        Loop
    Else
        CreateUploadZOTRT = vbNullString
        Exit Function
    End If
    
    Set RS = Nothing
    
    'BGS 11.15.2001 Now need to create the ZOT file
    'This will put all the .OUT filesin the upload dir into 1 ZOT file (ZIPED with Xceed ZIP)
    'And remove all the .OUT files from the upload dir.
    
    sZOTName = goUtil.utGetECSCryptSetting("ECS", "WEB_SECURITY", "USER_NAME", "UNKNOWN_ID")
    sZOTName = sZOTName & "_" & Format(Now, "MMDDYYHHMMSS") & ".ZOT"
    Set oXZip = New V2ECKeyBoard.clsXZip
    
    If oXZip.SaveZIPFiles(psUploadDir, sZOTName, "*.OUT", goUtil.DB_PASSWORD("1")) Then
        CreateUploadZOTRT = sZOTName
    End If
    
    Set oXZip = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CreateUploadZOTRT"
End Function

Public Function UpdateBatches(psOutFileDir As String, pPBar As Object, ptxtMess As Object) As Boolean
    On Error GoTo EH
    'BGS File Variables
    Dim iFileCount As Integer
    Dim iCount As Integer
    Dim saryUL() As String
    Dim sOutFile() As String
    Dim sZOTFile() As String  'BGS 10.30.2001 contains list of ZIPed out files
    Dim sClassName As String
    'BGS Field Object variables
    Dim vValue As Variant
    Dim sFieldName As String
    Dim iVarType As VariantTypeConstants
    Dim sField As String
    Dim iPos As Integer
    'BGS Objects
    Dim oField As V2ECKeyBoard.clsField
    Dim oBatches As V2ECKeyBoard.clsBatches
    Dim oCarrier As V2ECKeyBoard.clsCarUL
    Dim oXZip As V2ECKeyBoard.clsXZip
    
    'BGS Flags
    Static bFirstBatch As Boolean
    Dim bRebillPresent As Boolean
    'Error mess
    Dim sMess As String
    
    UpdateBatches = True
    
    Screen.MousePointer = vbHourglass
    
    'BGS 10.30.2001 need to loop through all the ZIPed OUT files "ZOT" files
    'and unzip them before we can process them.
    ReDim sZOTFile(1 To 1) As String
    sZOTFile(1) = Dir(psOutFileDir & "\*.ZOT", vbNormal)
    iFileCount = 1
    Do Until sZOTFile(iFileCount) = vbNullString
        iFileCount = iFileCount + 1
        ReDim Preserve sZOTFile(1 To iFileCount)
        sZOTFile(iFileCount) = Dir
    Loop
    
    Set oXZip = New V2ECKeyBoard.clsXZip
    iFileCount = 1
    Do Until sZOTFile(iFileCount) = vbNullString
        oXZip.UNZipFiles psOutFileDir, psOutFileDir & "\" & sZOTFile(iFileCount)
        Kill psOutFileDir & "\" & sZOTFile(iFileCount)
        iFileCount = iFileCount + 1
    Loop
    
    'BGS 10.15.2001 see if there are any "*.OUT" files to process.
    'Get a count of them and update the progress bar max
    ReDim sOutFile(1 To 1) As String
    sOutFile(1) = Dir(psOutFileDir & "\*.OUT", vbNormal)
    
    iFileCount = 1
    Do Until sOutFile(iFileCount) = vbNullString
        'BGS 10.17.2001 The Last element in the sOutFile array  should be
        'vbnullstring so the do loop below will know when to end.
        iFileCount = iFileCount + 1
        ReDim Preserve sOutFile(1 To iFileCount)
        sOutFile(iFileCount) = Dir
    Loop
    
    'BGS 11.18.2001 SOrt the Files ascending in the Array
    'BGS 11.26.2001 pass in Rebill present flag
    'Bubble sort will set it if applicable
    goUtil.utBubbleSort sOutFile, "_", bRebillPresent
    
    If iFileCount = 1 Then
        UpdateBatches = False
        'BGS Bail here if there are no files to update
        Screen.MousePointer = vbDefault
        If TypeOf ptxtMess Is TextBox Then
            ptxtMess.Text = "Nothing to import. " & Now()
        ElseIf TypeOf ptxtMess Is Label Then
            ptxtMess.Caption = "Nothing to import. " & Now()
        End If
        Exit Function
    End If
    
    pPBar.Max = iFileCount
    iFileCount = 1
    
    Do Until sOutFile(iFileCount) = vbNullString
        
        If TypeOf ptxtMess Is TextBox Then
            ptxtMess.Text = "Processing " & sOutFile(iFileCount)
        ElseIf TypeOf ptxtMess Is Label Then
            ptxtMess.Caption = "Processing " & sOutFile(iFileCount)
        End If
        ptxtMess.Refresh
        pPBar.Value = iFileCount
        saryUL = Split(goUtil.utGetFileData(psOutFileDir & "\" & sOutFile(iFileCount)), vbCrLf)
        sClassName = saryUL(0)
        
        'BGS 4.11.2002
        'Check to see if there is App.exename in there ('.')
        'If not then it was made with an older version clsCarULFarm
        'when it was still in ECKeyBoard So we need to append ECcarFarmers app name.
        If InStr(1, sClassName, ".", vbBinaryCompare) = 0 Then
            sClassName = "ECcarFarmers." & sClassName
        ElseIf InStr(1, sClassName, "ECFarmers.", vbTextCompare) > 0 Then
            sClassName = Replace(sClassName, "ECFarmers.", "ECcarFarmers.", , 1)
        End If
        
        'BGS check to see if we are dealing with the same Carrier Object
        'if we are then do not set it to nothing 'use the same object'
        '6.14.2002 use str comp not instr
        If Not oCarrier Is Nothing Then
            If StrComp(oCarrier.ClassName, sClassName, vbTextCompare) <> 0 Then
                Set oCarrier = Nothing
            End If
        End If
        
        'BGS 10.15.2001 Need to populate the fields collection with
        'Fields in the OUT text file
        
        Set mcolFields = New Collection
        For iCount = 1 To UBound(saryUL)
            sField = saryUL(iCount)
            
            'VARTYPE
            iPos = InStr(1, sField, F_DELIM, vbBinaryCompare) - 1
            If iPos > 0 Then
                iVarType = CInt(left(sField, iPos))
            Else
                iVarType = vbEmpty
            End If
            'BGS Clear out the VarType info
            sField = Replace(sField, iVarType & F_DELIM, vbNullString, 1, 1)
            
            'FIELDNAME
            iPos = InStr(1, sField, F_DELIM, vbBinaryCompare) - 1
            If iPos > 0 Then
                sFieldName = left(sField, iPos)
            Else
                sFieldName = vbNullString
            End If
            'BGS Clear out the FieldName info
            sField = Replace(sField, sFieldName & F_DELIM, vbNullString, 1, 1)
            
            'VALUE
            iPos = InStr(1, sField, F_DELIM, vbBinaryCompare) - 1
            If iPos > 0 Then
                'Check for encrypted string
                If left(sField, 1) = Chr(160) Then
                    If goUtil Is Nothing Then
                        Set goUtil = New V2ECKeyBoard.clsUtil
                    End If
                    'Decrypt it
                    vValue = goUtil.Decode(left(sField, iPos))
                Else
                    vValue = left(sField, iPos)
                End If
                
            Else
                vValue = vbNullString
            End If
            'BGS 10.15.2001 Need to convert to the correct DataType
            vValue = ConvertVarType(iVarType, vValue)
            
            'Add the field to the fields collection
            Set oField = New clsField
            With oField
                .DataType = iVarType
                .FieldName = sFieldName
                .Length = -1 'Not Udating this at this time
                .Value = Replace(vValue, F_VBCRLF, vbCrLf)
            End With
            
            mcolFields.Add oField, oField.FieldName
        Next
        
        'BGS 10.15.2001 Now update the batches table
        If oBatches Is Nothing Then
            Set oBatches = New clsBatches
            If Not oBatches.OpenConnection(WebControlDB) Then
                UpdateBatches = False
                GoTo CLEANUP
            End If
            If Not oBatches.OpenConnection(ClientBatchesDB) Then
                UpdateBatches = False
                GoTo CLEANUP
            End If
        End If
        
        If oCarrier Is Nothing Then
            Set oCarrier = CreateObject(sClassName)
        End If
        
        'BGS 11.26.2001 Need to set the Rebillpresent Flag so we know
        'if we need to load the RSbdt (RS that holds the BillData table info)
        'This is very important it takes around 5 minutes to init this RS
        'but after that don't need to open up the RS again for others
        'It is only necessary if we have a rebill file in the lot of them.
        oBatches.RebillPresent = bRebillPresent
        
        If Not bFirstBatch Then
            bFirstBatch = True
            If TypeOf ptxtMess Is TextBox Then
                ptxtMess.Text = ptxtMess.Text & " (First Batch Please Wait!)"
            ElseIf TypeOf ptxtMess Is Label Then
                ptxtMess.Caption = ptxtMess.Text & " (First Batch Please Wait!)"
            End If
            ptxtMess.Refresh
            
            If bRebillPresent Then
                If TypeOf ptxtMess Is TextBox Then
                    ptxtMess.Text = ptxtMess.Text & vbCrLf & Now & vbCrLf
                    ptxtMess.Text = ptxtMess.Text & "Rebill Present.  Please wait!  This could take several minutes."
                ElseIf TypeOf ptxtMess Is Label Then
                    ptxtMess.Caption = ptxtMess.Text & " " & Now & " "
                    ptxtMess.Caption = ptxtMess.Caption & "Rebill Present.  Please wait!  This could take several minutes."
                End If
                
                ptxtMess.Refresh
            End If
        End If
        
        If Not oCarrier.UpdateBatchesUDT(oBatches, Me) Then
            sMess = "Could not import: " & sOutFile(iFileCount) & vbCrLf & vbCrLf
            FireErrorMess sMess
            oBatches.CloseConnection WebControlDB
            oBatches.CloseConnection ClientBatchesDB
            Set oBatches = Nothing
        Else
            'If we updated Batches then we can update Webcontrol DB
            'or whatever other DB that uses this in the future
            oCarrier.UpdateDB oBatches, Me
        End If
        
        'BGS Need to reset the Fields object for the next UL File
        Set mcolFields = Nothing
        
        'BGS Now Make a backup of the UL File and remove it from the
        'UL Directory
        If Not goUtil.utFileExists(psOutFileDir & "\BackUp", True) Then
            MkDir psOutFileDir & "\BackUp"
        End If
        FileCopy psOutFileDir & "\" & sOutFile(iFileCount), psOutFileDir & "\BackUp\" & sOutFile(iFileCount)
        Kill psOutFileDir & "\" & sOutFile(iFileCount)
        
        'BGS 10.17.2001 INcrement the file count here
        iFileCount = iFileCount + 1
    Loop
    
    
    'BGS 10.30.2001 now zip up the Backed up out files into one big old
    'file for the day
    If goUtil.utFileExists(psOutFileDir & "\Backup", True) Then
        If goUtil.utFileExists(psOutFileDir & "\Backup\*.OUT", False) Then
            oXZip.SaveZIPFiles psOutFileDir & "\Backup", "ReceivedClientBatchesDB_" & Format(Now, "MMDDYYHHMMSS") & ".zip", "*.OUT", goUtil.DB_PASSWORD("1")
        End If
    End If
    
    
CLEANUP:
    If Not oBatches Is Nothing Then
        oBatches.CloseConnection WebControlDB
        oBatches.CloseConnection ClientBatchesDB
        Set oBatches = Nothing
    End If
    Set oXZip = Nothing
    Set mcolFields = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    UpdateBatches = False
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "Public Function UpdateBatches" & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    FireErrorMess sMess
End Function

Private Function ConvertVarType(piVarType As VbVarType, pvValue As Variant) As Variant
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Select Case piVarType
        Case VbVarType.vbArray
            ConvertVarType = CVar(pvValue)
        Case VbVarType.vbBoolean
            ConvertVarType = CBool(pvValue)
        Case VbVarType.vbByte
            ConvertVarType = CByte(pvValue)
        Case VbVarType.vbCurrency
            ConvertVarType = CCur(pvValue)
        Case VbVarType.vbDataObject
            ConvertVarType = CVar(pvValue)
        Case VbVarType.vbDate
            ConvertVarType = CDate(pvValue)
        Case VbVarType.vbDecimal
            ConvertVarType = CDec(pvValue)
        Case VbVarType.vbDouble
            ConvertVarType = CDbl(pvValue)
        Case VbVarType.vbEmpty
            ConvertVarType = CVar(pvValue)
        Case VbVarType.vbError
            ConvertVarType = CVar(pvValue)
        Case VbVarType.vbInteger
            ConvertVarType = CInt(pvValue)
        Case VbVarType.vbLong
            ConvertVarType = CLng(pvValue)
        Case VbVarType.vbNull
            ConvertVarType = CVar(pvValue)
        Case VbVarType.vbObject
            ConvertVarType = CVar(pvValue)
        Case VbVarType.vbSingle
            ConvertVarType = CSng(pvValue)
        Case VbVarType.vbString
            ConvertVarType = CStr(pvValue)
        Case VbVarType.vbUserDefinedType
            ConvertVarType = CVar(pvValue)
        Case VbVarType.vbVariant
        
    End Select
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
     Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function ConvertVarType"
End Function

Private Sub Class_Terminate()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mcolFields = Nothing
    Set moULCar = Nothing
    Set moParam = Nothing
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub Class_Terminate"
End Sub

Public Sub FireErrorMess(psMess As String)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    RaiseEvent ErrorMess(ByVal psMess)
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireErrorMess"
End Sub

Public Sub FireUpdateDB(pvBatches As Variant)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    RaiseEvent UpdateDB(ByVal pvBatches)
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireUpdateDB"
End Sub

Public Sub FireUpdateDBRT(pvBatchesRT As Variant, poBatch As V2ECKeyBoard.clsBatches, poRTUL As V2ECKeyBoard.clsCarUL)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    RaiseEvent UpdateDBRT(ByVal pvBatchesRT, poBatch, poRTUL, Me)
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub FireUpdateDBRT"
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
