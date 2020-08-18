Attribute VB_Name = "modBatches"
Option Explicit

Private mConnClaims As ADODB.Connection
Private mrsBat As ADODB.Recordset
Private mrsBill As ADODB.Recordset
Private mrsECS As ADODB.Recordset
Private mrsTAX As ADODB.Recordset
Private mrsGPTax As ADODB.Recordset
Private mfrmECUpdate As frmECUpdate
Public gbCancel As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const S_z As String = "¶Ññ"
Public gbSQLSVR As Boolean

Public Function OpenConnection() As Boolean
    On Error GoTo EH
    Dim sUserID As String
    Dim sPassword As String
    Dim sProdDSN As String
    
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    sUserID = GetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    sPassword = GetECSCryptSetting("V2WebControl", "DBConn", "PASSWORD")
    gbSQLSVR = True
        
    Set mConnClaims = New ADODB.Connection
    mConnClaims.ConnectionTimeout = 0 ' This will make the connection wait indefinately which is desirable for this application
    mConnClaims.Open sProdDSN, sUserID, sPassword
    OpenConnection = True
    Exit Function
EH:
    OpenConnection = False
    Err.Raise Err.Number, , Err.Description & vbCrLf & "Public Function OpenConnection"
End Function

Public Function CloseConnection() As Boolean
    On Error GoTo EH
    CloseConnection = True
    If Not mConnClaims Is Nothing Then
        mConnClaims.Close
        Set mConnClaims = Nothing
    End If
   
    Exit Function
EH:
    CloseConnection = False
    Err.Raise Err.Number, , Err.Description & vbCrLf & "Public Function CloseConnection"
End Function

Public Sub Main()
    On Error GoTo EH
    Dim sMess As String
    'If we already have this running then Bail
    If App.PrevInstance Then
        SaveSetting "ECUpdateBatches", "Msg", "PrevInstance", True
        End
        Exit Sub
    Else
        If Command$ = "RunAsDepOfV2AutoImport" Then
            GoTo RUN_UPDATE
        Else
            MsgBox "ECUpdateBatches may only be run as a dependant of Auto Import.", vbExclamation
            End
            Exit Sub
        End If
    End If
RUN_UPDATE:

    If OpenConnection Then
        If gbSQLSVR Then
            UpdateBatches_SQLSVR
        End If
        
        CloseConnection
    End If
    End
    Exit Sub
EH:
    sMess = "<------------" & App.EXEName & " " & Now() & "------------>" & vbCrLf
    sMess = sMess & "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Public Sub Main" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    SaveSetting "ECUpdateBatches", "Msg", "ErrorMess", sMess
    
   
    Set mrsBat = Nothing
    Set mrsECS = Nothing
'    Set mrsBill = Nothing
    Set mrsTAX = Nothing
    Set mrsGPTax = Nothing
    
    Unload mfrmECUpdate
    Set mfrmECUpdate = Nothing
    
End Sub


Public Function UpdateBatches_SQLSVR() As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim lCount As Long
    Dim RS As ADODB.Recordset
    Dim sSeek As String
    Dim lSSNSeek As Long
    Dim sCityName As String
    Dim sLossCity As String
    Dim sTemp1 As String
    Dim sTemp2 As String
    Dim supdSQL As String
    Dim supdLossCitySQL As String
    
    UpdateBatches_SQLSVR = True
    'Load the Progress Form
    Set mfrmECUpdate = New frmECUpdate
    Load mfrmECUpdate
    mfrmECUpdate.Show vbModeless
    mfrmECUpdate.lblField.Caption = "Loading, Please wait..."
    mfrmECUpdate.Refresh

    'Get Batches info
    sSQL = "SELECT "
    sSQL = sSQL & "BATCHES.[BatchesID], "
    sSQL = sSQL & "BATCHES.[billingdup], "
    sSQL = sSQL & "BATCHES.[date], "
    sSQL = sSQL & "BATCHES.[catsite], "
    sSQL = sSQL & "BATCHES.[adj_name], "
    sSQL = sSQL & "BATCHES.[adjuster_n], "
    sSQL = sSQL & "BATCHES.[ecupdated], "
    sSQL = sSQL & "BATCHES.[ibnumber], "
    sSQL = sSQL & "BATCHES.[lossstate], "
    sSQL = sSQL & "BATCHES.[loss_loc], "
    sSQL = sSQL & "BATCHES.[losscity], "
    sSQL = sSQL & "BATCHES.[ssn], "
    sSQL = sSQL & "BATCHES.[copied], " 'Updated 10.14.2002
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   [Catcode] "
    sSQL = sSQL & "FROM     ClientCompanyCatSpec "
    sSQL = sSQL & "WHERE    [ClientCompanyCatSpecID] = BATCHES.[ClientCompanyCatSpecID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "As [CATCODE] "
    sSQL = sSQL & "FROM BATCHES "
    sSQL = sSQL & "WHERE BATCHES.[ssn] > 0 "
    sSQL = sSQL & "AND BATCHES.[copied] Is Null "
    sSQL = sSQL & "AND BATCHES.[ecupdated] Is Null Or BATCHES.[ECUPDATED] = 0 "
    
    Set mrsBat = New ADODB.Recordset
    mrsBat.CursorType = adOpenKeyset
    mrsBat.LockType = adLockOptimistic
    mrsBat.Open sSQL, mConnClaims, , , adCmdText
    
    'If there is nothing to process then bail
    If mrsBat.RecordCount = 0 Then
        GoTo CLEANUP
    End If
    
    lCount = 0
    mfrmECUpdate.PBarRecord.Value = 0
    
    'Get Scrub RS from Stored Proc.
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "Exec spsClaimScrub", mConnClaims, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    Do Until RS.State = adStateOpen
        DoEvents
        Sleep 100
        If gbCancel Then
            GoTo CLEANUP
        End If
    Loop
    
    Set mrsECS = RS
    Set mrsBill = RS.NextRecordset
    Set mrsTAX = RS.NextRecordset
    Set mrsGPTax = RS.NextRecordset
    
    'Set up the progress bar
    mfrmECUpdate.PBarRecord.Max = mrsBat.RecordCount
    mrsBat.MoveFirst
    Do Until mrsBat.EOF
        supdSQL = "UPDATE BATCHES SET "
        'I. Check for cancel button
        If gbCancel Then
            Exit Do
        End If
        'IV Need to Copy over original ADJ Name
        supdSQL = supdSQL & "ADJ_NAME = '" & Replace(Trim(mrsBat!adjuster_n), "'", "''") & "', "
        
        '2. Look for Loss state and SSN match
        If Not mrsECS.EOF Then
            mrsECS.MoveFirst
            lSSNSeek = IIf(IsNull(mrsBat!SSN), -999, mrsBat!SSN)
            If lSSNSeek <> -999 Then
                Do Until mrsECS.EOF
                    If Not IsNull(mrsECS!SS_NUM) Then
                        If mrsECS!SS_NUM = lSSNSeek Then
                            Exit Do
                        End If
                    End If
                    mrsECS.MoveNext
                Loop
            End If
            If Not mrsECS.EOF And lSSNSeek <> -999 Then
                supdSQL = supdSQL & "ADJUSTER_N = '" & Replace(RTrim(mrsECS!LAST_NAME), "'", "''") & " " & Replace(RTrim(mrsECS!First_Name), "'", "''") & "', "
            Else
               supdSQL = supdSQL & "ADJUSTER_N = '" & "?Unknown?" & lSSNSeek & "', "
            End If
           
            mrsECS.MoveFirst
        End If
        
        '3 Need to See if the City Name needs to be checked against the
        'Tax Table SO we get the Exact City Name Populated
        If Not mrsGPTax.EOF Then
            mrsGPTax.MoveFirst
            sSeek = IIf(IsNull(mrsBat!lossstate), vbNullString, Replace(RTrim(mrsBat!lossstate), "'", "''"))
            If sSeek <> vbNullString Then
                Do Until mrsGPTax.EOF
                    If Not IsNull(mrsGPTax!State) Then
                        If StrComp(RTrim(mrsGPTax!State), sSeek, vbTextCompare) = 0 Then
                            mrsGPTax.MoveFirst
                            'Get the Loss City from Batches
                            sLossCity = IIf(IsNull(mrsBat!losscity), vbNullString, Replace(RTrim(mrsBat!losscity), "'", "''"))
                            GoTo CHANGE_CITY_NAME
                        End If
                    End If
                    mrsGPTax.MoveNext
                Loop
            End If
            mrsGPTax.MoveFirst
        End If
        supdLossCitySQL = vbNullString
        GoTo SKIP_CITY
CHANGE_CITY_NAME:
        If Not mrsTAX.EOF Then
            mrsTAX.MoveFirst
            If sSeek <> vbNullString And sLossCity <> vbNullString Then
                Do Until mrsTAX.EOF
                    If Not IsNull(mrsTAX!State) Then
                        If StrComp(RTrim(mrsTAX!State), sSeek, vbTextCompare) = 0 Then
                            'Get the City name from the Tax Table
                            sCityName = IIf(IsNull(mrsTAX!CITY_NAME), vbNullString, RTrim(mrsTAX!CITY_NAME))
                            'Check for "(" this needs to be weeded out from the Tx table item
                            If InStr(1, sCityName, "(", vbBinaryCompare) > 0 Then
                                sCityName = RTrim(Left(sCityName, InStr(1, sCityName, "(", vbBinaryCompare) - 1))
                            End If

                            'Now do the Check to see if this city name is a match against the Uploaded LossCity
                            If InStr(1, sLossCity, sCityName, vbTextCompare) > 0 Then
                                'If we find the City name in the loss city also need to check the
                                'first couple of chars
                                sTemp1 = Left(sLossCity, 3)
                                sTemp2 = Left(sCityName, 3)
                                If StrComp(sTemp1, sTemp2, vbTextCompare) = 0 Then
                                    supdLossCitySQL = "LOSSCITY = '" & Replace(mrsTAX!CITY_NAME, "'", "''") & "', "
                                    mrsTAX.MoveFirst
                                    GoTo SKIP_CITY
                                End If
                            End If
                        End If
                    End If
                    mrsTAX.MoveNext
                Loop
            End If
            'If we get here then didn't find the City for a state that needs it
            supdLossCitySQL = "LOSSCITY = '" & Left("?Unknown?" & Replace(RTrim(mrsBat!losscity), "'", "''"), 50) & "', "
            mrsTAX.MoveFirst
        End If
        '4. Mark Record as Updated
SKIP_CITY:
        supdSQL = supdSQL & supdLossCitySQL
        supdSQL = supdSQL & "ECUPDATED = 1, "
        supdSQL = supdSQL & "COPIED = 0 "  'Updated 10.14.2002
        supdSQL = supdSQL & "WHERE BatchesID = " & mrsBat!BatchesID & " "
        
       mConnClaims.Execute supdSQL
        
        lCount = lCount + 1
        If lCount <= mfrmECUpdate.PBarRecord.Max Then
            mfrmECUpdate.PBarRecord.Value = lCount
        End If
        mfrmECUpdate.lblField.Caption = lCount & " Of " & mfrmECUpdate.PBarRecord.Max
        mrsBat.MoveNext
        DoEvents
        'Check for cancel button
        If gbCancel Then
            Exit Do
        End If
    Loop
    
CLEANUP:
    
    Set mrsBat = Nothing
    Set mrsECS = Nothing
    Set mrsBill = Nothing
    Set mrsTAX = Nothing
    Set mrsGPTax = Nothing
    Set RS = Nothing
    
    Unload mfrmECUpdate
    Set mfrmECUpdate = Nothing
    Exit Function
EH:
    UpdateBatches_SQLSVR = False
    Err.Raise Err.Number, , Err.Description & vbCrLf & "Public Function UpdateBatches_SQLSVR"
End Function

Private Function GetCity(psAddress As String) As String
    On Error GoTo EH
    Dim sAddress As String
    Dim sZip As String
    Dim sState As String
    Dim sCity As String
    Dim sStreet As String
    
    sAddress = psAddress
    
    FillAddressFields sAddress, sZip, sState, sCity, sStreet
    
    If sCity <> vbNullString Then
        GetCity = sCity
    Else
        GetCity = psAddress
    End If

    Exit Function
EH:
    GetCity = psAddress
End Function

Private Sub FillAddressFields(psAddress As String, _
                             psZip As String, _
                             psState As String, _
                             psCity As String, _
                             psStreet As String)
    On Error GoTo EH
    Dim sTemp As String
    Dim sAddress As String
    Dim sValTemp As String
    Dim lPos As Long
    
    sAddress = Trim(Replace(psAddress, vbCrLf, vbNullString))
    
    'Zip code
    If InStr(1, sAddress, " ", vbBinaryCompare) > 0 Then
        sTemp = Trim(Mid(sAddress, InStrRev(sAddress, " ", , vbBinaryCompare)))
        'Val function Bug in VB6
        'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
        'Need to parse out both % and !  because these trailing equate to Double and Single
        'and Val bugs because it can't convert Double or single into integer
        sValTemp = Replace(sTemp, "-", vbNullString)
        If Val(CleanValString(sValTemp)) > 0 Then
            If Len(sTemp) >= 5 Then
                'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
                'Need to use string reverse to get proper Left length
                'Using Replace can not work here, must use right to left logic.
                lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
                If lPos >= 0 Then
                    sAddress = Trim(Left(sAddress, lPos))
                End If
                sTemp = Replace(sTemp, ",", vbNullString)
                psZip = sTemp
            Else
                psZip = vbNullString
                psState = vbNullString
                psCity = vbNullString
                GoTo ADDRESS
            End If
        Else
            psZip = vbNullString
            psState = vbNullString
            psCity = vbNullString
            GoTo ADDRESS
        End If
    Else
        psZip = vbNullString
        psState = vbNullString
        psCity = vbNullString
        GoTo ADDRESS
    End If
    
    'State
    If Len(sAddress) > 2 Then
        sTemp = Right(sAddress, 2)
        If Val(CleanValString(sTemp)) = 0 Then
            'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
            lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
            If lPos >= 0 Then
                sAddress = Trim(Left(sAddress, lPos))
            End If
            sTemp = Replace(sTemp, ",", vbNullString)
            psState = sTemp
        Else
            psState = vbNullString
            psCity = vbNullString
            GoTo ADDRESS
        End If
    Else
        psState = vbNullString
        psCity = vbNullString
        GoTo ADDRESS
    End If
    
    'City
    If InStr(1, sAddress, S_z, vbBinaryCompare) > 0 Then
        sTemp = Trim(Mid(sAddress, InStrRev(sAddress, S_z, , vbBinaryCompare)))
        If Val(CleanValString(sTemp)) = 0 Then
            'Issue 243 9.10.2002 Copy Button for Address Chops of Letters
            lPos = InStrRev(sAddress, sTemp, , vbBinaryCompare) - 1
            If lPos >= 0 Then
                sAddress = Trim(Left(sAddress, lPos))
            End If
            sTemp = Replace(sTemp, ",", vbNullString)
            sTemp = Replace(sTemp, S_z, vbNullString)
            sTemp = Replace(sTemp, Chr(32), Chr(160))
            psCity = sTemp
        Else
            psCity = vbNullString
        End If
    Else
        psCity = vbNullString
    End If
ADDRESS:
    'Address
    sAddress = Replace(sAddress, ",", vbNullString)
    sAddress = Replace(sAddress, S_z, vbNullString)
    psStreet = sAddress
    
    Exit Sub
EH:
   
End Sub

Private Function CleanValString(psValText As String) As String
    'Val function Bug in VB6
    'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
    'Need to parse out both % and !  because these trailing equate to Double and Single
    'and Val() bugs because it can't convert Double or single into integer
    On Error GoTo EH
    
    psValText = Replace(psValText, "%", vbNullString)
    psValText = Replace(psValText, "!", vbNullString)
    CleanValString = psValText
    Exit Function
EH:
    CleanValString = vbNullString
End Function

Public Function GetECSCryptSetting(psAPP As String, psSECTION As String, psKEY As String, _
                                   Optional pvDefault As Variant = vbNullString) As Variant
    On Error GoTo EH
    Dim sCryptSetting As String
    Dim oUtil As V2ECKeyBoard.clsUtil
    
    Set oUtil = New V2ECKeyBoard.clsUtil
    
    sCryptSetting = GetSetting(psAPP, psSECTION, psKEY, vbNullString)
    
    If sCryptSetting <> vbNullString Then
        GetECSCryptSetting = CStr(oUtil.Decode(sCryptSetting))
    Else
        GetECSCryptSetting = pvDefault
    End If
    
    Set oUtil = Nothing
    Exit Function
EH:
    Err.Raise Err.Number, , Err.Description & vbCrLf & "Public Function GetECSCryptSetting"
End Function

