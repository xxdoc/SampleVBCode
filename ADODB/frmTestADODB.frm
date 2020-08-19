VERSION 5.00
Begin VB.Form frmTestADODB 
   Caption         =   "frmTestADODB"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmTestADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public Enum enCmd
    CommandTimeout = 500
    
End Enum
Public Enum enConn
    CommandTimeout = 500
    ConnectionTimeout = 30
End Enum

Private moConn As ADODB.Connection
Private moCmd As ADODB.Command
Private moRS As ADODB.Recordset
Private msConnStr As String


Private Property Get Conn() As ADODB.Connection
    Set Conn = moConn
End Property
Private Property Set Conn(ByRef poConn As ADODB.Connection)
    Set moConn = poConn
End Property

Private Property Get Cmd() As ADODB.Command
    Set Cmd = moCmd
End Property
Private Property Set Cmd(ByRef poCmd As ADODB.Command)
    Set moCmd = poCmd
End Property

Private Property Get RS() As ADODB.Recordset
    Set RS = moRS
End Property
Private Property Set RS(ByRef poRS As ADODB.Recordset)
    Set moRS = poRS
End Property

Private Property Get ConnStr() As String
    ConnStr = msConnStr
End Property
Private Property Let ConnStr(ByRef psConnStr As String)
    msConnStr = psConnStr
End Property


Private Sub Command1_Click()
    On Error GoTo EH
    
    Dim myDistEmpties As modTypes.typDistEmpties

    With myDistEmpties
        .strTAPAppVersion = "TAP2.5"
        .strRegion = "West"
        .strOrderBy = "Name"
        .strCustSel = "100001"
        .lngCompany = 10
        .intDebugOn = 1
    End With
    
    Set Cmd = New ADODB.Command
    Set Conn = New ADODB.Connection
    'Sanitized
    ConnStr = "Provider=SQLOLEDB;Password=**********;Persist Security Info=True;User ID=***********;Initial Catalog=***;Data Source=**-**.********.com"
    
    
    Cmd.CommandTimeout = enCmd.CommandTimeout
    Conn.CursorLocation = adUseClient
    Conn.CommandTimeout = enConn.CommandTimeout
    Conn.ConnectionTimeout = enConn.ConnectionTimeout
    
    Conn.Open (ConnStr)
    Set Cmd.ActiveConnection = Conn
    Cmd.CommandText = "tap_spDistEmpties"
    Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    
'Just accessing Paramters and hoping everything works out.
'    With myDistEmpties
'        '1. pTAPAppVersion --Versioning, just in case [different strokes for different folks]
'        '"declare @pTAPAppVersion varchar(50) "
'        Cmd.Parameters("@pTAPAppVersion") = IsNullOrEmptyDefault(.strTAPAppVersion, vbNullString)
'        '2. pRegion --dbo.[Users].Region
'        '"declare @pRegion nvarchar(50) "
'        Cmd.Parameters("@pRegion") = IsNullOrEmptyDefault(.strRegion, vbNullString)
'        '3. pOrderBy --Fields to order by e.g. [ShipState], [Name]
'        '"declare @pOrderBy nvarchar(100) "
'        Cmd.Parameters("@pOrderBy") = IsNullOrEmptyDefault(.strOrderBy, vbNullString)
'        '4. pCustSel --dbo.[INVENTORYREPORT].[UserID]
'        '"declare @pCustSel nvarchar(50) "
'        Cmd.Parameters("@pCustSel") = IsNullOrEmptyDefault(.strCustSel, vbNullString)
'        '5. pCompany --Set @pCompany = -1 for ALL Companies in results.  Otherwise, ALL other companies besides @pCompany value will be r
'        '"declare @pCompany bigint "
'        Cmd.Parameters("@pCompany") = IsNullOrEmptyDefault(.lngCompany, -1)
'        '6. pDebugOn --Debugging?  SET @pDebugOn = 1 IF NOT SET @pDebugOn = 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
'        '"declare @pDebugOn bit "
'        'Debug OFF
'        Cmd.Parameters("@pDebugOn") = IsNullOrEmptyDefault(.intDebugOn, 0)
'    End With


'The table below shows the ADO Data Type mapping between Access, SQL Server, and Oracle:
'
'DataType Enum      Value           Access                                                      SQLServer                               Oracle
'adBigInt           20                                                                          BigInt (SQL Server 2000 +)
'adBinary           128                                                                         Binary-TimeStamp                        Raw *
'adBoolean          11              YesNo                                                       Bit
'adChar             129                                                                         Char                                    Char
'adCurrency         6               Currency                                                    Money-SmallMoney
'adDate             7               Date                                                        DateTime
'adDBTimeStamp      135             DateTime (Access 97 (ODBC))                                 DateTime-SmallDateTime                  Date
'adDecimal          14                                                                                                                  Decimal *
'adDouble           5               Double                                                      Float                                   Float
'adGUID             72              ReplicationID (Access 97 (OLEDB)), (Access 2000 (OLEDB))    UniqueIdentifier (SQL Server 7.0 +)
'adIDispatch        9
'adInteger          3               AutoNumber                                                  Identity (SQL Server 6.5)               Int *
'                                   Integer Int
'                                   Long
'adLongVarBinary    205             OLEObject                                                   Image                                   Long Raw *
'                                                                                                                                       Blob (Oracle 8.1.x)
'adLongVarChar      201             Memo (Access 97)                                            Text                                    Long *
'                                   Hyperlink (Access 97)                                                                               Clob ( Oracle 8.1.x)
'
'adLongVarWChar     203             Memo (Access 2000 (OLEDB))                                  NText (SQL Server 7.0 +)                NClob (Oracle 8.1.x)
'                                   Hyperlink (Access 2000 (OLEDB))
'
'adNumeric          131             Decimal (Access 2000 (OLEDB))                               Decimal                                 Decimal
'                                                                                               Numeric                                 Integer
'                                                                                                                                       Number
'                                                                                                                                       SmallInt
'adSingle           4               Single                                                      Real
'adSmallInt         2               Integer                                                     SmallInt
'adUnsignedTinyInt  17              Byte                                                        TinyInt
'adVarBinary        204             ReplicationID (Access 97)                                   VarBinary
'adVarChar          200             Text (Access 97)                                            VarChar                                 VarChar
'adVariant          12                                                                          Sql_Variant (SQL Server 2000 +)         VarChar2
'adVarWChar         202             Text (Access 2000 (OLEDB))                                  NVarChar (SQL Server 7.0 +)             NVarChar2
'adWChar            130                                                                         NChar (SQL Server 7.0 +)

'Using Parameters.Append
    With myDistEmpties
        '1. pTAPAppVersion --Versioning, just in case [different strokes for different folks]
        '"declare @pTAPAppVersion varchar(50) "
        Cmd.Parameters.Append Cmd.CreateParameter("@pTAPAppVersion", adVarChar, adParamInput, 50, IsNullOrEmptyDefault(.strTAPAppVersion, vbNullString))
        '2. pRegion --dbo.[Users].Region
        '"declare @pRegion nvarchar(50) "
        Cmd.Parameters.Append Cmd.CreateParameter("@pRegion", adVarWChar, adParamInput, 50, IsNullOrEmptyDefault(.strRegion, vbNullString))
        '3. pOrderBy --Fields to order by e.g. [ShipState], [Name]
        '"declare @pOrderBy nvarchar(100) "
        Cmd.Parameters.Append Cmd.CreateParameter("@pOrderBy", adVarWChar, adParamInput, 100, IsNullOrEmptyDefault(.strOrderBy, vbNullString))
        '4. pCustSel --dbo.[INVENTORYREPORT].[UserID]
        '"declare @pCustSel nvarchar(50) "
        Cmd.Parameters.Append Cmd.CreateParameter("@pCustSel", adVarWChar, adParamInput, 50, IsNullOrEmptyDefault(.strCustSel, vbNullString))
        '5. pCompany --Set @pCompany = -1 for ALL Companies in results.  Otherwise, ALL other companies besides @pCompany value will be r
        '"declare @pCompany bigint "
        Cmd.Parameters.Append Cmd.CreateParameter("@pCompany", adBigInt, adParamInput, , IsNullOrEmptyDefault(.lngCompany, -1))
        '6. pDebugOn --Debugging?  SET @pDebugOn = 1 IF NOT SET @pDebugOn = 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
        '"declare @pDebugOn bit "
        'Debug OFF
        Cmd.Parameters.Append Cmd.CreateParameter("@pDebugOn", adBoolean, adParamInput, , IsNullOrEmptyDefault(.intDebugOn, 0))
    End With
    
    Set RS = Cmd.Execute
    'Disconnect the recordset
    Set RS.ActiveConnection = Nothing
    
    If Not RS.BOF And Not RS.EOF Then
        RS.MoveFirst
    End If
    
    RS.Close
    Conn.Close
    
    
    'clean up
GoTo CLEANUP
EH:
    MsgBox "Error: #" & Err.Number & " Description: " & Err.Description, vbCritical, "Error"
CLEANUP:
    Set RS = Nothing
    Set Cmd = Nothing
    Set Conn = Nothing
End Sub

Public Function IsNullOrEmptyDefault(ByRef pVar, ByRef pDefaultValue)
    Dim varRet

    If IsNull(pVar) Or IsEmpty(pVar) Then
        varRet = pDefaultValue
    ElseIf Not IsNumeric(pVar) Then
        If Trim(pVar) = vbNullString Then
            varRet = pDefaultValue
        Else
            varRet = pVar
        End If
    Else
        varRet = pVar
    End If

    IsNullOrEmptyDefault = varRet
End Function
