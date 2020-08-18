VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'BGS DBASE ENUM
Public Enum DBConn
    WebControlDB = 1        'WebControlDB
    ClientBatchesDB         'ClientBatchesDB
End Enum

Public Type udtBatches
    dtContacted As Date             'Not in Batches Table used to update WebControl DB
    lSSN As String                    '11.21.2001 SSN used to match Cat_no/Site
    sBatchID As String              'char    17  0
    sCarrierCode As String          'nvarchar    2   1
    sIBNumber As String             'nvarchar    10  1
    cPayment As Currency            'money   8   1
    sPaymentID As String            'nvarchar    20  1
    dtDate As Date                  'datetime    8   1
    sAdjuster_N As String           'nvarchar    50  1
    sAdjuster_I As String           'nvarchar    50  1
    sClaimNumber As String          'nvarchar    250 1
    sInsuredName As String          'nvarchar    50  1
    sLossLocationCity As String     'nvarchar    30  1
    sLossLocationState As String    'nvarchar    5   1
    dtDateOfLoss As Date            'datetime    8   1
    dtDateClosed  As Date           'datetime    8   1
    cGrossLoss As Currency          'nvarchar    20  1
    sCatCode As String              'nvarchar    30  1
    cTotalServiceFees As Currency   'nvarchar    20  1
    cAdministrativeFee As Currency  'nvarchar    20  1
    cMiscCharges As Currency
    cTaxesTotal  As Currency        'nvarcharas    20  1
    cTotalFee As Currency           'nvarchar    20  1
End Type

Public Type RTDeleteAction
    ID As Long
    sTableName As String
    sIBNumber As String
    sClientClaimNumber As String
End Type


Public Type udtBatchesRT
    sStatus As String
    sSiteAddress As String
    sSiteCity As String
    sSiteState As String
    sSiteZip As String
    dtInspectedDate As Date
    dtReceivedDate As Date
    dtContactedDate As Date
    dtAssignedDate As Date
    sTypeOfLoss As String
    sPolicyNumber As String
    sMailAddress As String
    sLossLocation As String
    sMortgageeName As String
    cBuildingLimits As Currency
    cContentsLimits As Currency
    cAPSLimits As Currency
    cBuildingRCLimits As Currency
    cContentsRCLimits As Currency
    cAPSRCLimits As Currency
    sDeductibles As String
    lSSN As String                  '11.21.2001 SSN used to match Cat_no/Site
    sRTID As String                 'char    17  0
    sCarrierCode As String          'nvarchar    2   1
    sIBNumber As String             'nvarchar    10  1
    cPayment As Currency            'money   8   1
    sPaymentID As String            'nvarchar    20  1
    dtDate As Date                  'datetime    8   1
    sAdjuster_N As String           'nvarchar    50  1
    sAdjuster_I As String           'nvarchar    50  1
    sClaimNumber As String          'nvarchar    250 1
    sInsuredName As String          'nvarchar    50  1
    sLossLocationCity As String     'nvarchar    30  1
    sLossLocationState As String    'nvarchar    5   1
    dtDateOfLoss As Date            'datetime    8   1
    dtDateClosed  As Date           'datetime    8   1
    cGrossLoss As Currency          'nvarchar    20  1
    sCatCode As String              'nvarchar    30  1
    cTotalServiceFees As Currency   'nvarchar    20  1
    cAdministrativeFee As Currency  'nvarchar    20  1
    cMiscCharges As Currency
    cTaxesTotal  As Currency        'nvarcharas    20  1
    cTotalFee As Currency           'nvarchar    20  1
End Type

Private mConnWebControlDB As ADODB.Connection
Private mConnClientBatchesDB As ADODB.Connection
Private mudtBatches As udtBatches
Private mudtBatchesRT As udtBatchesRT
Private msIBTable As String
Private msBatchID As String
Private msRTID As String
Private mbDuplicateIB As Boolean
Private mRS As ADODB.Recordset
Private mRSBdt As ADODB.Recordset 'This will stay open with one time pass to get data
Private mRSAdj As ADODB.Recordset 'This will stay open with one time pass to get data
Private mbRebillPresent As Boolean 'This is true if there is a Rebill Present.

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsBatches"
End Property

Public Property Get Batches() As udtBatches
    Batches = mudtBatches
End Property
Public Property Let Batches(pudtBatches As udtBatches)
    mudtBatches = pudtBatches
    msBatchID = mudtBatches.sBatchID
End Property

Public Property Get BatchesRT() As udtBatchesRT
    BatchesRT = mudtBatchesRT
End Property
Public Property Let BatchesRT(pudtBatchesRT As udtBatchesRT)
    mudtBatchesRT = pudtBatchesRT
    msRTID = mudtBatchesRT.sRTID
End Property

Public Property Get RebillPresent() As Boolean
    RebillPresent = mbRebillPresent
End Property
Public Property Let RebillPresent(pbFlag As Boolean)
    mbRebillPresent = pbFlag
End Property

Public Property Get BatchID() As String
    BatchID = msBatchID
End Property
Public Property Let BatchID(psID As String)
    msBatchID = psID
End Property

Public Property Get RTID() As String
    RTID = msRTID
End Property
Public Property Let RTID(psID As String)
    msRTID = psID
End Property

Public Function OpenConnection(piDBConn As DBConn) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    OpenConnection = True
    'ADO Object
    Select Case piDBConn
        Case DBConn.WebControlDB
            If mConnWebControlDB Is Nothing Then
                Set mConnWebControlDB = New ADODB.Connection
                mConnWebControlDB.Open GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
            End If
        Case DBConn.ClientBatchesDB
'            'If we are not updating ClientBatchesDB (BillDataPV table) then Don't open connection.
'            If Not CBool(GetSetting(App.EXEName, "DBConn", "DisableClientBatchesDBUpdate", True)) Then
'                If mConnClientBatchesDB Is Nothing Then
'                    Set mConnClientBatchesDB = New ADODB.Connection
'                    mConnClientBatchesDB.Open GetSetting(App.EXEName, "DBConn", "ClientBatchesDB", vbNullString)
'                End If
'            End If
            
    End Select
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    OpenConnection = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function OpenConnection"
End Function

Public Function CloseConnection(piDBConn As DBConn) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    CloseConnection = True
    'ADO Object For Detail Records
    Select Case piDBConn
        Case DBConn.WebControlDB
            If Not mConnWebControlDB Is Nothing Then
                mConnWebControlDB.Close
                Set mConnWebControlDB = Nothing
            End If
        Case DBConn.ClientBatchesDB
'            If Not mConnClientBatchesDB Is Nothing Then
'                mConnClientBatchesDB.Close
'                Set mConnClientBatchesDB = Nothing
'            End If
    End Select
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CloseConnection = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CloseConnection"
End Function

Public Function UpdateBatchesTable() As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim bDuplicate As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Static bCollateError As Boolean
    
    UpdateBatchesTable = True
    
    'BGS 10.31.2001 first need to check for dups
    'Use Disconnected RS
    sSQL = "SELECT BATCHES.IBNUMBER FROM BATCHES "
    sSQL = sSQL & "WHERE BATCHES.IBNUMBER = '" & goUtil.utCleanSQLString(mudtBatches.sIBNumber) & "'"
    
    'If we are not updating ClientBatchesDB (Batches table then bail here
    '5.3.2002
    If CBool(GetSetting(App.EXEName, "DBConn", "DisableClientBatchesDBUpdate", True)) Then
        Exit Function
    End If
    
    Set mRS = New ADODB.Recordset
    mRS.CursorLocation = adUseClient
    mRS.Open sSQL, mConnClientBatchesDB, adOpenStatic, adLockBatchOptimistic
    Set mRS.ActiveConnection = Nothing

    If Not mRS.EOF Then
        bDuplicate = True
    Else
        bDuplicate = False
    End If
    
    mbDuplicateIB = bDuplicate
    'clean up
    Set mRS = Nothing

    
    With mudtBatches
        sSQL = "INSERT INTO BATCHES (BatchID, "
        sSQL = sSQL & "ssn, "
        sSQL = sSQL & "carriercod, "
        sSQL = sSQL & "IBNumber, "
        sSQL = sSQL & "Payment, "
        sSQL = sSQL & "PaymentID, "
        'BGS 11.28.2001 Using [Microsoft][ODBC Visual FoxPro Driver] Date Format is funcky = "{mm-dd-yyyy}"
        sSQL = sSQL & "[Date], "
        sSQL = sSQL & "Adjuster_N, "
        sSQL = sSQL & "Adjuster_I, "
        sSQL = sSQL & "ClaimNumbe, "
        sSQL = sSQL & "InsuredNam, "
        sSQL = sSQL & "LossCity, "
        sSQL = sSQL & "LossState, "
        'BGS 11.28.2001 Using [Microsoft][ODBC Visual FoxPro Driver] Date Format is funcky = "{mm-dd-yyyy}"
        sSQL = sSQL & "DateOfLoss, "
        sSQL = sSQL & "DateClosed, "
        sSQL = sSQL & "GrossLoss, "
        sSQL = sSQL & "CatCode, "
        sSQL = sSQL & "TotalServi, "
        sSQL = sSQL & "Administra, "
        sSQL = sSQL & "MiscCharge, "
        sSQL = sSQL & "TaxesTotal, "
        sSQL = sSQL & "TotalFee, "
        sSQL = sSQL & "Duplicate ) "

        sSQL = sSQL & "VALUES ('" & goUtil.utCleanSQLString(.sBatchID) & "', "
        sSQL = sSQL & IIf(.lSSN = vbNullString, "0", goUtil.utCleanSQLString(.lSSN)) & ", "
        sSQL = sSQL & IIf(.sCarrierCode = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sCarrierCode) & "'") & ", "
        sSQL = sSQL & IIf(.sIBNumber = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sIBNumber) & "'") & ", "
        sSQL = sSQL & .cPayment & ", "
        sSQL = sSQL & IIf(.sPaymentID = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sPaymentID) & "'") & ", "
        'BGS 11.28.2001 Using [Microsoft][ODBC Visual FoxPro Driver] Date Format is funcky = "{mm-dd-yyyy}"
        sSQL = sSQL & IIf(.dtDate = NULL_DATE, Format(NULL_DATE, "{mm-dd-yyyy}"), Format(.dtDate, "{mm-dd-yyyy}")) & ", "
        sSQL = sSQL & IIf(.sAdjuster_N = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sAdjuster_N) & "'") & ", "
        sSQL = sSQL & IIf(.sAdjuster_I = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sAdjuster_I) & "'") & ", "
        sSQL = sSQL & IIf(.sClaimNumber = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sClaimNumber) & "'") & ", "
        sSQL = sSQL & IIf(.sInsuredName = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sInsuredName) & "'") & ", "
        sSQL = sSQL & IIf(.sLossLocationCity = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sLossLocationCity) & "'") & ", "
        sSQL = sSQL & IIf(.sLossLocationState = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sLossLocationState) & "'") & ", "
        'BGS 11.28.2001 Using [Microsoft][ODBC Visual FoxPro Driver] Date Format is funcky = "{mm-dd-yyyy}"
        sSQL = sSQL & IIf(.dtDateOfLoss = NULL_DATE, Format(NULL_DATE, "{mm-dd-yyyy}"), Format(.dtDateOfLoss, "{mm-dd-yyyy}")) & ", "
        sSQL = sSQL & IIf(.dtDateClosed = NULL_DATE, Format(NULL_DATE, "{mm-dd-yyyy}"), Format(.dtDateClosed, "{mm-dd-yyyy}")) & ", "
        sSQL = sSQL & .cGrossLoss & ", "
        sSQL = sSQL & IIf(.sCatCode = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sCatCode) & "'") & ", "
        sSQL = sSQL & .cTotalServiceFees & ", "
        sSQL = sSQL & .cAdministrativeFee & ", "
        sSQL = sSQL & .cMiscCharges & ", "
        sSQL = sSQL & .cTaxesTotal & ", "
        sSQL = sSQL & .cTotalFee & ", "
        sSQL = sSQL & CInt(bDuplicate) * -1 & " ) "
    End With
    mConnClientBatchesDB.Execute sSQL
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    'Collating error is occurring on the initial opening of recordset. After the first error
    'do not receive this error on subsequent opening of recordsets. So if we get this error
    'then we need to Close the connection and repoen it to get around this error
    'This error needs to be investigated I looked it up on Microsoft support...
    'http://support.microsoft.com/support/kb/articles/Q246/1/67.ASP  This may only be
    'occuring when in the IDE ??  here is the work around anyhow  Only disconnect and reconnect once
    'use the bcollateerror flag so that we don't possibly get in a never ending loop.
    
    If lErrNum = -2147467259 And Not bCollateError Then
        
        CloseConnection WebControlDB
        OpenConnection WebControlDB
        If CBool(GetSetting(App.EXEName, "DEBUG", "COLLATE_ERROR", False)) Then
            MsgBox "Collate error !", vbExclamation, "Debug Error Message"
        End If
        bCollateError = True
        Resume
    End If
    
    UpdateBatchesTable = False
    Set mRS = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function UpdateBatchesTable"
End Function

Public Function UpdateIBTable(poCarrier As V2ECKeyBoard.clsCarUL, poUL As V2ECKeyBoard.clsUpload) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oField As clsField
    Dim colFields As Collection
    Dim sQuote As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    UpdateIBTable = True
    
    sSQL = "INSERT INTO " & poCarrier.TableName & " ( BATCHID, "
    Set colFields = poUL.Fields
    For Each oField In colFields
        sSQL = sSQL & oField.FieldName & ", "
    Next
    
    'BGS trim off the last ,
    sSQL = left(sSQL, InStrRev(sSQL, ",") - 1) & ") "
    
    sSQL = sSQL & "VALUES ( '" & goUtil.utCleanSQLString(msBatchID) & "', "
    
    For Each oField In colFields
        sQuote = GetQuote(oField)
        sSQL = sSQL & sQuote & goUtil.utCleanSQLString(oField.Value) & sQuote & ", "
    Next
    
    'BGS trim off the last ,
    sSQL = left(sSQL, InStrRev(sSQL, ",") - 1) & ") "
    sSQL = Replace(sSQL, vbCrLf, " ")
    mConnWebControlDB.Execute sSQL
    
    'cleanup
    Set colFields = Nothing
    Set oField = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    UpdateIBTable = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function UpdateIBTable"
End Function

Public Function UpdateAllRTTables(poCarrier As V2ECKeyBoard.clsCarUL, poUL As V2ECKeyBoard.clsUpload) As Boolean
    On Error GoTo EH
    Dim oField As clsField
    Dim colFields As Collection
    Dim colTableFields As Collection
    Dim bDeleteFromRTTable As Boolean
    Dim bUpdateRtTable As Boolean
    Dim sTableName As String
    Dim sIBFieldName As String
    Dim sIB As String
    Dim sClientClaimFieldName As String
    Dim sClientClaim As String
    Dim sCheckFieldName As String
    Dim sCheck As String
    Dim sFieldName As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    UpdateAllRTTables = True
    Set colFields = poUL.Fields
    'Need to loop through all the fields looking for Tables
    'with records to be (Added, updated) or deleted from the Server.
    For Each oField In colFields
        'Look for special defined fields that are really labels for Tables
        If oField.DataType = vbUserDefinedType Then
            If InStr(1, oField.FieldName, "UPDATE_RT_TABLE", vbBinaryCompare) > 0 _
            Or InStr(1, oField.FieldName, "DELETE_FROM_RT_TABLE", vbBinaryCompare) > 0 Then
                If colTableFields Is Nothing Then
                    Set colTableFields = New Collection
                    If InStr(1, oField.FieldName, "UPDATE_RT_TABLE", vbBinaryCompare) > 0 Then
                        bUpdateRtTable = True
                        bDeleteFromRTTable = False
                        sTableName = oField.Value
                    ElseIf InStr(1, oField.FieldName, "DELETE_FROM_RT_TABLE", vbBinaryCompare) > 0 Then
                        bDeleteFromRTTable = True
                        bUpdateRtTable = False
                    End If
                Else
                    If sTableName <> vbNullString Then
                        GoTo RUN_TABLE
                    End If
                End If
            ElseIf InStr(1, oField.FieldName, "END_OF_RT_TABLE", vbBinaryCompare) > 0 Then
RUN_TABLE:
                If bUpdateRtTable Then
                    'Run the table update
                    UpdateRTTable sTableName, colTableFields, sIBFieldName, sIB, sClientClaimFieldName, sClientClaim, sCheckFieldName, sCheck
                    bUpdateRtTable = False
                ElseIf bDeleteFromRTTable Then
                    'Run the Delete from table
                    DeleteFromRTTable sTableName, colTableFields, sIBFieldName, sIB, sClientClaimFieldName, sClientClaim
                    bDeleteFromRTTable = False
                End If
                'Reset the Table variables for the next table
                Set colTableFields = Nothing
                sTableName = vbNullString
                sIBFieldName = vbNullString
                sIB = vbNullString
                sClientClaimFieldName = vbNullString
                sClientClaim = vbNullString
                sCheckFieldName = vbNullString
                sCheck = vbNullString
            End If
        Else
            If colTableFields Is Nothing Then
                Set colTableFields = New Collection
                bUpdateRtTable = True
            End If
            'Parse out the field name
            sFieldName = Mid(oField.FieldName, InStr(1, oField.FieldName, "=", vbBinaryCompare) + 1)
            If InStr(1, sFieldName, ":", vbBinaryCompare) > 0 Then
                sFieldName = left(sFieldName, InStr(1, sFieldName, ":", vbBinaryCompare) - 1)
            End If
            If InStr(1, oField.FieldName, "IB_NAME=", vbBinaryCompare) > 0 Then
                sIBFieldName = sFieldName
                sIB = oField.Value
            ElseIf InStr(1, oField.FieldName, "CLIENT_CLAIM_NAME=", vbBinaryCompare) > 0 Then
                sClientClaimFieldName = sFieldName
                sClientClaim = oField.Value
            ElseIf InStr(1, oField.FieldName, "CHECK_NAME=", vbBinaryCompare) > 0 Then
                sCheckFieldName = sFieldName
                sCheck = oField.Value
            ElseIf InStr(1, oField.FieldName, "TABLE_NAME=", vbBinaryCompare) > 0 Then
                sTableName = sFieldName
            End If
            'Set the field name to the parsed field name
            oField.FieldName = sFieldName
            'Add the fields to the TableFields collection
            colTableFields.Add oField, oField.FieldName
        End If
    Next
    
    'clean up
    Set oField = Nothing
    Set colFields = Nothing
    Set colTableFields = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    UpdateAllRTTables = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function UpdateAllRTTables"
End Function

Public Function UpdateRTTable(psTableName As String, pColFields As Collection, _
                              psIBFieldName As String, _
                              psIB As String, _
                              psClientClaimFieldName As String, _
                              psClientClaim As String, _
                              psCheckFieldName As String, _
                              psCheck As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim oField As clsField
    Dim colFields As Collection
    Dim sQuote As String
    Dim lRecAffected As Long
    Dim bResetRTTable As Boolean
    Dim sRTTableName As String
    Dim lID As Long 'Used for RT Tables unique ID for Multi records
    'Need to update Memo fields lil different
    Dim colMemoFields As Collection
    Dim rstMemo As ADODB.Recordset
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    UpdateRTTable = True
    Set colFields = pColFields
    
    sSQL = "UPDATE " & psTableName & " SET "
    For Each oField In colFields
        sQuote = GetQuote(oField)
        'Check for memo field
        If Len(Trim(oField.Value)) > 255 Then
            sSQL = sSQL & psTableName & "." & oField.FieldName & " = '', "
            If colMemoFields Is Nothing Then
                Set colMemoFields = New Collection
            End If
            colMemoFields.Add oField
        Else
            '267  FoxPro Activity Log Memo field Bug
            '1.28.2003 Need to Check the Length of the trimmed Field
            'If it is 0 then Need to put a printable char in there so
            'it can't inherit another Memo field value.
            If Len(Trim(oField.Value)) = 0 Then
                oField.Value = Chr(160)
            End If
            sSQL = sSQL & psTableName & "." & oField.FieldName & " = " & sQuote & goUtil.utCleanSQLString(Trim(oField.Value)) & sQuote & ", "
        End If
        If StrComp(oField.FieldName, "lID", vbTextCompare) = 0 Then
            lID = oField.Value
        End If
    Next
    'BGS trim off the last ,
    sSQL = left(sSQL, InStrRev(sSQL, ",") - 1) & " "
    
    sSQL = sSQL & "WHERE " & psTableName & "." & psIBFieldName & " = '" & goUtil.utCleanSQLString(psIB) & "' "
    sSQL = sSQL & "AND " & psTableName & "." & psClientClaimFieldName & " = '" & goUtil.utCleanSQLString(psClientClaim) & "' "
    'Was check
    If lID > 0 Then
        sSQL = sSQL & "AND " & psTableName & ".LID = " & lID & " "
    End If
    
    sSQL = Replace(sSQL, vbCrLf, F_VBCRLF)
    
    mConnWebControlDB.Execute sSQL, lRecAffected
    If lRecAffected = 0 Then
        sSQL = "INSERT INTO " & psTableName & " ( RTID, "
        For Each oField In colFields
            sSQL = sSQL & oField.FieldName & ", "
        Next
        'BGS trim off the last ,
        sSQL = left(sSQL, InStrRev(sSQL, ",") - 1) & ") "
        
        sSQL = sSQL & "VALUES ( '" & goUtil.utCleanSQLString(msRTID) & "', "
        For Each oField In colFields
            sQuote = GetQuote(oField)
            If Len(oField.Value) > 255 Then
                sSQL = sSQL & "'', "
            Else
                sSQL = sSQL & sQuote & goUtil.utCleanSQLString(Trim(oField.Value)) & sQuote & ", "
            End If
        Next
        'BGS trim off the last ,
        sSQL = left(sSQL, InStrRev(sSQL, ",") - 1) & ") "
            
        sSQL = Replace(sSQL, vbCrLf, " ")
        
        mConnWebControlDB.Execute sSQL, lRecAffected
        If lRecAffected = 0 Then
            Err.Raise -999, , "No Records Affected."
        Else
            GoTo UPDATE_MEMO
        End If
    Else
UPDATE_MEMO:
        If Not colMemoFields Is Nothing Then
            For Each oField In colMemoFields
                sSQL = "SELECT " & UCase(psTableName) & "." & UCase(oField.FieldName) & " "
                sSQL = sSQL & "FROM " & UCase(psTableName) & " "
                sSQL = sSQL & "WHERE " & UCase(psTableName) & "." & UCase(psIBFieldName) & " = '" & goUtil.utCleanSQLString(psIB) & "' "
                sSQL = sSQL & "AND " & UCase(psTableName) & "." & UCase(psClientClaimFieldName) & " = '" & goUtil.utCleanSQLString(psClientClaim) & "' "
                If lID > 0 Then
                    sSQL = sSQL & "AND " & UCase(psTableName) & ".LID = " & lID & " "
                End If
                Set rstMemo = New ADODB.Recordset
                rstMemo.CursorType = adOpenKeyset
                rstMemo.LockType = adLockOptimistic
                rstMemo.Open sSQL, mConnWebControlDB, , , adCmdText
                rstMemo.Fields(0).Value = Trim(oField.Value)
                rstMemo.Update
                rstMemo.Close
                Set rstMemo = Nothing
            Next
        End If
    End If
    'cleanup
    Set colMemoFields = Nothing
    Set colFields = Nothing
    Set oField = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    UpdateRTTable = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function UpdateRTTable"
End Function

Public Function DeleteFromRTTable(psTableName As String, pColFields As Collection, _
                              psIBFieldName As String, _
                              psIB As String, _
                              psClientClaimFieldName As String, _
                              psClientClaim As String) As Boolean
    On Error GoTo EH
    Dim MyRTDel As RTDeleteAction
    Dim oField As clsField
    Dim sSQL As String
    Dim sTableName As String
    Dim saryTables() As String
    Dim lTableCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    With MyRTDel
        For Each oField In pColFields
            Select Case UCase(oField.FieldName)
                Case "ID"
                    .ID = oField.Value
                Case "STABLENAME"
                    .sTableName = oField.Value
            End Select
        Next
        .sClientClaimNumber = psClientClaim
        .sIBNumber = psIB
    End With
    
    'If the ID is set to -999 that means Del the entire Claim
    'if not then just delete the specific ID from the Specific Table
    'By IBNumber and ClientClaimNumber
    If MyRTDel.ID = -999 Then
        sTableName = MyRTDel.sTableName
        sTableName = Mid(sTableName, InStr(1, sTableName, ":") + 1)
        saryTables = Split(sTableName, ",")
        
        For lTableCount = LBound(saryTables, 1) To UBound(saryTables, 1)
            MyRTDel.sTableName = saryTables(lTableCount)
            With MyRTDel
                'Issue 240 9.6.2002 Supplement not being deleted from Web
                'Need to account for Rebill\supplement suffixes on the IBnumber
                'in all carrier RT tables (all other tables like indem photo etc etc
                'are ok since they have the ibnumber without suffix). Any supplements
                'or rebills that were being uploaded and then deleted would not get deleted
                'because the IB number we pass in does not contain the IB suffix,
                'just the IB. This only happens on delete actions.
                sSQL = "DELETE FROM " & UCase(.sTableName) & " "
                sSQL = sSQL & "WHERE Left(" & UCase(.sTableName) & ".SIBNUMBER, " & Len(.sIBNumber) & ") = '" & .sIBNumber & "' "
                sSQL = sSQL & "AND " & UCase(.sTableName) & ".SCLIENTCLA = '" & .sClientClaimNumber & "' "
            End With
            mConnWebControlDB.Execute sSQL
        Next
    Else
        With MyRTDel
            'Issue 240 9.6.2002 Supplement not being deleted from Web
            sSQL = "DELETE FROM " & UCase(.sTableName) & " "
            sSQL = sSQL & "WHERE Left(" & UCase(.sTableName) & ".SIBNUMBER, " & Len(.sIBNumber) & ") = '" & .sIBNumber & "' "
            sSQL = sSQL & "AND " & UCase(.sTableName) & ".SCLIENTCLA = '" & .sClientClaimNumber & "' "
            sSQL = sSQL & "AND " & UCase(.sTableName) & ".LID = " & .ID & " "
        End With
        mConnWebControlDB.Execute sSQL
    End If
    
    DeleteFromRTTable = True
    
    'cleanup
    Set oField = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    DeleteFromRTTable = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function DeleteFromRTTable"
End Function

Public Function UpdateBillDataTable() As Boolean
'**********************************************************************************************************************
'BGS 11.16.2001  first need to check for Existing IBs (If the current IB we have Has An 'R#=Rebilled' with the same
'IB number that has been Paid (This will be indicated by the Rec_Num being populated)
'(**Note** The rec_num field should be populated when the payment is received by ECS from the Carrier)
'Also need to take into account that the following may be true for preexisting IBs
'1. The prexisting IB may have already been paid or may not have been payed.
'  1a. If The IB has already been paid then we need to Figure out the Difference
    'from the already paid IB for Service Fee, Misc Fee, and Admin fee if applicable
    'Also need to take into account that there may be multiple rebills for this single
    'IB and figure the difference for the current IB appropriately.
   '1b. If the IB has not been paid then we can  insert the The Rebill as it is (Checking for previous rebills and figuring the differences appropriately if applicable)
    'The reason we can not just delete the unpaid IB in this instance is because there actually may be
    'a pending payment for this IB.  If there is then we can not delete it.  The decision to delete this
    '"Duplicate" IB has to be made by admin folks.  We can make their job a little easier by marking the
    'the IB field with {?}
'**********************************************************************************************************************


    On Error GoTo EH
    Dim sSQL As String
    Dim iConn As DBConn
    
    Dim cServiceFee As Currency
    Dim cMiscFee As Currency
    Dim cAdminFee As Currency
    
    Dim sIBMatch As String
    Dim iIBPos As Integer
    Dim sIBmess As String
    Dim vValues(1 To 4) As Variant 'Hold Values that need to be added to RSbdt
    
    Dim bDuplicate As Boolean
    Dim bResearchNeeded As Boolean
    Dim bUseDiff As Boolean
    Static bCollateErrorWebControlDB As Boolean
    Static bCollateErrorClientBatchesDB As Boolean
    
    'BGS 11.21.2001 "Cat_no" in ClientBatchesDB BillData.dbf
    'BGS 11.21.2001 "Site" in WebControlDB ECSADJ.dbf
    'Need to match up "Site_State" and "Ss_num" to get CatNo\Site
    Dim sCatNo_Site As String
    
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    UpdateBillDataTable = True
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '9.27.2002 All this is now being handled in a separate process
    'ECUpdateBatches.exe until further notice
    Exit Function
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    'If we are not updating ClientBatchesDB (BillDataPV table then bail here
    '5.3.2002
    If CBool(GetSetting(App.EXEName, "DBConn", "DisableClientBatchesDBUpdate", True)) Then
        Exit Function
    End If
    
    'We are still on an open connection to WebControlDB
    iConn = WebControlDB
    
    'BGS 11.21.2001 Need to get the Cat_no / Site from ECSADJ if possible
    'Noticed that this SELECT statement takes About 3 seconds to Execute.
    'Sheesh man. thats like forever.
    
    If mRSAdj Is Nothing Then
        sSQL = "SELECT SITE, SITE_STATE, SS_NUM "
        sSQL = sSQL & "FROM ECSADJ "
        
        Set mRSAdj = New ADODB.Recordset
        mRSAdj.CursorLocation = adUseClient
        mRSAdj.Open sSQL, mConnWebControlDB, adOpenStatic, adLockBatchOptimistic
        Set mRSAdj.ActiveConnection = Nothing
    End If

    If Not mRSAdj.EOF Then
        mRSAdj.MoveFirst
        'BGS loop through the RS instead of using SQL to parse for the State and SSN.
        'looping and using strComp actually is faster than doing it in SQL DBASE III
        Do Until mRSAdj.EOF
            If StrComp(mRSAdj!SS_NUM, mudtBatches.lSSN, vbBinaryCompare) = 0 Then
                If StrComp(Right(RTrim(mRSAdj!SITE_STATE), 2), mudtBatches.sLossLocationState, vbTextCompare) = 0 Then
                    sCatNo_Site = mRSAdj!SITE
                    Exit Do
                Else
                    mRSAdj.MoveNext
                End If
            Else
                mRSAdj.MoveNext
            End If
        Loop
        mRSAdj.MoveFirst
    End If
    
    If sCatNo_Site = vbNullString Then
        sCatNo_Site = "Error"
    End If
    
    'BGS 11.26.2001 Need to use ClientBatchesDB connection here
    iConn = ClientBatchesDB
    If mbRebillPresent Then
        If mRSBdt Is Nothing Then
            OpenRSBdt
        End If
    End If
    
    'BGS 11.26.2001 only make this check if we have a Rebill in the current IB
    'we want to insert into the BillData.dbf table
    'BGS 12.14.2001 start looking for the "R" after the first 4 chars otherwise
    'we will loop unecessarily for those IB with R in the beginning

    If InStr(5, mudtBatches.sIBNumber, "R", vbTextCompare) > 0 Then
        'BGS then we need to check for Duplicates and Rebills if it was paid or not
        'if it wasn't paid yet then we will put {Research IB} in the IB that we are about to insert
        If Not mRSBdt.EOF Then
            'BGS 11.26.2001 need to match up the Current IB we want to insert with
            'any previous IB's.  If we are dealing with a Supplemental, need to match
            'it up with any previous Supplementals that have not been paid.
            'the IB could look like this FLYYMMDDHHMMSS R1.S1.
            'most likely will look like this FLYYMMDDHHMMSS
            'F= first name L= Last Name YY=YEAR MM = MONTH DD = DAY HH = HOUR MM = MIN SS = SEC
            'R1. = Rebill 1. S1. = Supplemental 1.
            sIBMatch = mudtBatches.sIBNumber
            If Len(sIBMatch) > 14 Then
                sIBMatch = Trim(Mid(sIBMatch, 15))
                iIBPos = InStr(1, sIBMatch, "S", vbTextCompare)
                If iIBPos > 0 Then
                    sIBMatch = Mid(sIBMatch, iIBPos)
                Else
                    sIBMatch = "NOSUP"
                End If
            Else
                sIBMatch = "NOSUP"
            End If
            mRSBdt.MoveFirst
                Do Until mRSBdt.EOF
                    'BGS loop through the RS instead of using SQL to parse for the State and SSN.
                    'looping and using strComp actually is faster than doing it in SQL DBASE III
                    If StrComp(left(LTrim(mRSBdt!IB), 14), left(mudtBatches.sIBNumber, 14), vbTextCompare) = 0 Then
                        'BGS 11.26.2001 If we have previous IB (Rebilling) need to add up the
                        'previous Fees that have not yet been paid If we are not dealing with Supplements
                        'at this point we can just add up the previous unpaid IB
                        If sIBMatch = "NOSUP" Then
                            bUseDiff = True
                            cServiceFee = cServiceFee + IIf(IsNull(mRSBdt!SERVICE_FE), 0, mRSBdt!SERVICE_FE)
                            cAdminFee = cAdminFee + IIf(IsNull(mRSBdt!ADMIN_FEE), 0, mRSBdt!ADMIN_FEE)
                            cMiscFee = cMiscFee + IIf(IsNull(mRSBdt!IB_ADJ_EXP), 0, mRSBdt!IB_ADJ_EXP)
                        Else
                            'BGS 11.26.2001 only add up the matching supplementals
                            'Do not include any supplementals that do not match the current
                            'supplementals. Having to do this should be very rare but just incase
                            'we need to do this code for when it does happen once in a blue moon.
                            If InStr(1, mRSBdt!IB, sIBMatch, vbTextCompare) > 0 Then
                                bUseDiff = True
                                cServiceFee = cServiceFee + IIf(IsNull(mRSBdt!SERVICE_FE), 0, mRSBdt!SERVICE_FE)
                                cAdminFee = cAdminFee + IIf(IsNull(mRSBdt!ADMIN_FEE), 0, mRSBdt!ADMIN_FEE)
                                cMiscFee = cMiscFee + IIf(IsNull(mRSBdt!IB_ADJ_EXP), 0, mRSBdt!IB_ADJ_EXP)
                            End If
                        End If
                        'BGS if it is null then this previous IB has not yet been paid
                        'so we have to mark the current entry IB field with {?} meaning that
                        'IB and previous unpaid IB needs research because the new fees may not
                        'work out right if we have a previous billing that has not already been paid.
                        If Not IsNull(mRSBdt!RECORD_NUM) Then
                            If mRSBdt!RECORD_NUM = 0 Then
                                bResearchNeeded = True
                            End If
                        Else
                            bResearchNeeded = True
                        End If
                    End If
                    mRSBdt.MoveNext
                Loop
                mRSBdt.MoveFirst
        End If
        
    End If
    
    'BGS 11.26.2001 need to populate the sibmess with appropriate message
    'if we are dealing with rebilled item that needs research or if we have
    'a duplicate ib entry wich should never ever happen but just in case.
    'The IB field can only contain 25 Chars So the max chars the iBmess can be is 4 {?D}.
    '{?} means admin folks have to research this IB entry for previous rebills that
    '{D} means there is a duplicate IB entry (This should never happen but just in case}
    '{D?} means there is duplicate entry that also has rebilling issue to take care of
    
    If mbDuplicateIB Or bResearchNeeded Then
        If bResearchNeeded Then
            sIBmess = "{?}"
        End If
        If mbDuplicateIB Then
            If sIBmess = vbNullString Then
                sIBmess = "{D}"
            Else
                sIBmess = "{?D}"
            End If
        End If
    End If
    
    
    With mudtBatches
        msBatchID = .sBatchID
        'BGS 12.14.2001 Insert Into BILLDATAPV   PV is the preview table.
        'WebControlDB will be set up to have the Admin people preview these items and they
        'can Verify these entries before they import them into the BillData table.
        sSQL = "INSERT INTO BILLDATAPV (CAT_NO, "
        sSQL = sSQL & "CLAIM_NO, "
        sSQL = sSQL & "FULL_NAME, "
        sSQL = sSQL & "CLAIM_CITY, "
        sSQL = sSQL & "CLAIM_STAT, "
        sSQL = sSQL & "ADJUSTOR_N, "
        'BGS 11.28.2001 Using [Microsoft][ODBC Visual FoxPro Driver]Date Format is funcky= "{mm-dd-yyyy}"
        sSQL = sSQL & "ENTRY_DATE, "
        sSQL = sSQL & "FILES_RECD, "
        sSQL = sSQL & "DATE_CLOSE, "
        sSQL = sSQL & "IB, "
        sSQL = sSQL & "SERVICE_FE, "
        sSQL = sSQL & "GROSS_LOSS, "
        sSQL = sSQL & "IB_ADJ_EXP, "
        sSQL = sSQL & "ADMIN_FEE) "
        
        sSQL = sSQL & "VALUES ('" & goUtil.utCleanSQLString(sCatNo_Site) & "', "
        sSQL = sSQL & IIf(.sClaimNumber = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sClaimNumber) & "'") & ", "
        sSQL = sSQL & IIf(.sInsuredName = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sInsuredName) & "'") & ", "
        sSQL = sSQL & IIf(.sLossLocationCity = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sLossLocationCity) & "'") & ", "
        sSQL = sSQL & IIf(.sLossLocationState = vbNullString, "''", "'" & goUtil.utCleanSQLString(.sLossLocationState) & "'") & ", "
        sSQL = sSQL & IIf(.sAdjuster_N = vbNullString, vbNullString, "'" & goUtil.utCleanSQLString(.sAdjuster_N) & "'") & ", "
        'BGS 11.28.2001 Using [Microsoft][ODBC Visual FoxPro Driver] Date Format is funcky = "{mm-dd-yyyy}"
        sSQL = sSQL & Format(Now, "{mm-dd-yyyy}") & ", "
        sSQL = sSQL & IIf(.dtDateOfLoss = NULL_DATE, Format(NULL_DATE, "{mm-dd-yyyy}"), Format(.dtDateOfLoss, "{mm-dd-yyyy}")) & ", "
        sSQL = sSQL & IIf(.dtDateClosed = NULL_DATE, Format(NULL_DATE, "{mm-dd-yyyy}"), Format(.dtDateClosed, "{mm-dd-yyyy}")) & ", "
        vValues(1) = .sIBNumber & sIBmess
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(.sIBNumber & sIBmess) & "', "
        vValues(2) = IIf(bUseDiff, .cTotalServiceFees - cServiceFee, .cTotalServiceFees)
        sSQL = sSQL & IIf(bUseDiff, .cTotalServiceFees - cServiceFee, .cTotalServiceFees) & ", "
        sSQL = sSQL & .cGrossLoss & ", "
        vValues(3) = IIf(bUseDiff, .cMiscCharges - cMiscFee, .cMiscCharges)
        sSQL = sSQL & IIf(bUseDiff, .cMiscCharges - cMiscFee, .cMiscCharges) & ", "
        vValues(4) = IIf(bUseDiff, .cAdministrativeFee - cAdminFee, .cAdministrativeFee)
        sSQL = sSQL & IIf(bUseDiff, .cAdministrativeFee - cAdminFee, .cAdministrativeFee) & ") "
 
        'BGS 11.26.2001 ONLY if Rebill is present then we need to add each record we insert
        'also into the Billdata rs
        If mbRebillPresent Then
            mRSBdt.AddNew
            mRSBdt!IB = vValues(1)
            mRSBdt!SERVICE_FE = vValues(2)
            mRSBdt!IB_ADJ_EXP = vValues(3)
            mRSBdt!ADMIN_FEE = vValues(4)
            mRSBdt.Update
        End If
    End With
    'BGS 11.16.2001 Need to update the BillDataTable. It is loacted in
    'another directory other than WebControlDB directory so use the ClientBatchesDB Directory.
    iConn = ClientBatchesDB
    mConnClientBatchesDB.Execute sSQL

    Exit Function
EH:
    'Collating error is occurring on the initial opening of recordset. After the first error
    'do not receive this error on subsequent opening of recordsets. So if we get this error
    'then we need to Close the connection and repoen it to get around this error
    'This error needs to be investigated I looked it up on Microsoft support...
    'http://support.microsoft.com/support/kb/articles/Q246/1/67.ASP  This may only be
    'occuring when in the IDE ??  here is the work around anyhow  Only disconnect and reconnect once
    'use the bcollateerror flag so that we don't possibly get in a never ending loop.

    If Err.Number = -2147467259 Then
        Select Case iConn
            Case DBConn.WebControlDB
                If Not bCollateErrorWebControlDB Then
                    CloseConnection WebControlDB
                    OpenConnection WebControlDB
                    bCollateErrorWebControlDB = True
                Else
                    GoTo EH2
                End If
            Case DBConn.ClientBatchesDB
                If Not bCollateErrorClientBatchesDB Then
                    CloseConnection ClientBatchesDB
                    OpenConnection ClientBatchesDB
                    bCollateErrorClientBatchesDB = True
                Else
                    GoTo EH2
                End If
        End Select
        If CBool(GetSetting(App.EXEName, "DEBUG", "COLLATE_ERROR", False)) Then
            MsgBox "Collate error!", vbExclamation, "Debug Error Message"
        End If
        Resume
    End If
EH2:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    UpdateBillDataTable = False
    Set mRSBdt = Nothing
    Set mRSAdj = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function UpdateBillDataTable"
End Function

Private Function OpenRSBdt() As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    OpenRSBdt = True
    If mRSBdt Is Nothing Then
        sSQL = "SELECT IB, RECORD_NUM, SERVICE_FE, "
        sSQL = sSQL & "IB_ADJ_EXP, ADMIN_FEE "
        sSQL = sSQL & "FROM BILLDATA "
        
        Set mRSBdt = New ADODB.Recordset
        mRSBdt.CursorLocation = adUseClient
        'BGS 11.26.2001 the init time on this RS is around 5 min with DBF table size of 100 MB
        'This should only hit here once then the rest of the processing uses this same RS.
        'Archiving this table will make this go faster.  Again this code will happen whenever
        'we have a Rebill In the lot of IBs.
        mRSBdt.Open sSQL, mConnClientBatchesDB, adOpenStatic, adLockBatchOptimistic
        Set mRSBdt.ActiveConnection = Nothing
    End If
        
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    OpenRSBdt = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function OpenRSBdt"
End Function

Public Function OpenDCRS(psSQL As String, pConn As DBConn) As ADODB.Recordset

    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Set OpenDCRS = New ADODB.Recordset
    OpenDCRS.CursorLocation = adUseClient
    OpenDCRS.CursorType = adOpenStatic
    OpenDCRS.LockType = adLockBatchOptimistic
    Select Case pConn
        Case DBConn.WebControlDB
            OpenDCRS.Open psSQL, mConnWebControlDB, , , adCmdText
        Case DBConn.ClientBatchesDB
            OpenDCRS.Open psSQL, mConnClientBatchesDB, , , adCmdText
    End Select
    Set OpenDCRS.ActiveConnection = Nothing
        
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function OpenDCRS"
End Function

Private Function GetQuote(pField As clsField) As String
    On Error GoTo EH
    Static bLeft As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    Select Case pField.DataType
        Case VbVarType.vbString
            GetQuote = "'"
        Case VbVarType.vbDate
            GetQuote = vbNullString
            If pField.Value = NULL_DATE Then
                pField.Value = Format(NULL_DATE, "{mm-dd-yyyy}")
            Else
                pField.Value = Format(pField.Value, "{mm-dd-yyyy}")
            End If
        Case Else
            GetQuote = vbNullString
    End Select
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Function GetQuote"
End Function

Private Sub Class_Terminate()
    Set mRSAdj = Nothing
    Set mRSBdt = Nothing
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
