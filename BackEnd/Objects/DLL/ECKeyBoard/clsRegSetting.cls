VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Submitted on: 1/18/2001 4:41:47 AM
'By: Stuart Roberts
'Level: Intermediate
'User Rating:      By 3 Users
'Compatibility: VB 6#
'
'Users have accessed this code 1322 times.
'  (About the author)
'http://www.planetsourcecode.com/xq/ASP/txtCodeId.14499/lngWId.1/qx/vb/scripts/ShowCode.htm
'

Option Explicit

'# Error constants
Private Const cnErrBadRegKey = 20000
Private Const cdErrBadRegKey = "The registry key specified is invalid."

'# Events
Public Event ErrorOccurred(Number As Long, Source As String, Description As String)
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'# Public Enumerations for Registry Key handles
Public Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'# Public Enumerations for Registry Key values
Public Enum Reg_Value
    NULL_VALUE = 0
    STRING_VALUE = 1
    BINARY_VALUE = 3
    DWORD_VALUE = 4
End Enum

'# Public Enumerations for Registry query results
Public Enum Reg_Result
    STRING_RESULT = 0
    BOOLEAN_RESULT = 1
    INTEGER_RESULT = 2
End Enum

'
' Reg Key Security Options
'
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))



'# Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

'# Key creation/open disposition
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

'# Error codes
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const ERROR_SUCCESS = 0&

'# Error messages
Private Const MSG_BADDB = "Error with database."
Private Const MSG_BADKEY = "The key specified is invalid or cannot be found."
Private Const MSG_CANTOPEN = "Cannot open specified key."
Private Const MSG_CANTREAD = "Cannot read specified key."
Private Const MSG_CANTWRITE = "Cannot write to specified key."
Private Const MSG_OUTOFMEMORY = "Out of memory while performing registry task."
Private Const MSG_INVALID_PARAMETER = "The parameter specified is invalid."
Private Const MSG_ACCESS_DENIED = "Access was denied."
Private Const MSG_INVALID_PARAMETERS = "The parameters specified is invalid."
Private Const MSG_NO_MORE_ITEMS = "There are no more items to return."
Private Const MSG_UNKNOWN = "Error performing task on the Registry."

Private Const REG_SZ = 1                         ' Unicode nul terminated string

'# Private property varibles
Private m_ErrorSource As String
Private m_ErrorDescription As String
Private m_ErrorNumber As Long
Private m_ErrorExist As Boolean

Property Get ErrorSource() As String
On Error Resume Next
    ErrorSource = m_ErrorSource
End Property

Property Get ErrorDescription() As String
On Error Resume Next
    ErrorDescription = m_ErrorDescription
End Property

Property Get ErrorNumber() As Long
On Error Resume Next
    ErrorNumber = m_ErrorNumber
End Property

Property Get ErrorExist() As Boolean
On Error Resume Next
    ErrorExist = m_ErrorExist
    m_ErrorExist = False ' Reset error status
End Property

Public Function Delete_Key(lngKey As hKey, sKeyName As String) As Boolean
' Description:
'   This Function will Delete a key from the Registry
On Error GoTo HandleError
    Dim lRetVal As Long
    
    Delete_Key = False
    lRetVal = RegDeleteKey(lngKey, sKeyName)
    If lRetVal = ERROR_NONE Then
        Delete_Key = True
    Else
        ProcessError lRetVal, "Reg.Delete_Key" & vbCrLf & vbCrLf & Err.Source, Err.Description, True
    End If
    
    Exit Function
    
HandleError:
    ProcessError Err.Number, "Reg.Delete_Key (" & sKeyName & ")" & vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Public Function Delete_Value(lngKey As hKey, sKeyName As String, sValueName As String) As Boolean
' Description:
'   This Function will delete a value from the Registry
On Error GoTo HandleError
    Dim lRetVal As Long
    Dim hdKey As Long
    
    Delete_Value = False
    lRetVal = RegOpenKeyEx(lngKey, sKeyName, 0, KEY_ALL_ACCESS, hdKey)
    lRetVal = RegDeleteValue(hdKey, sValueName)
    RegCloseKey (hdKey)
    If lRetVal = ERROR_NONE Then
        Delete_Value = True
    Else
        ProcessError lRetVal, "Reg.Delete_Value" & vbCrLf & vbCrLf & Err.Source, Err.Description, True
    End If
    
    Exit Function
    
HandleError:
    ProcessError Err.Number, "Reg.Delete_Value (" & sKeyName & "," & sValueName & ")" & vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Public Function Set_Key(lngKey As hKey, sNewKeyName As String) As Boolean
' Description:
'   This Function will create a new key in the Registry
On Error GoTo HandleError
    Dim hNewKey As Long
    Dim lRetVal As Long
    
    Set_Key = False
    lRetVal = RegCreateKeyEx(lngKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
    If lRetVal = ERROR_NONE Then
        Set_Key = True
    Else
        ProcessError lRetVal, "Reg.Set_Key" & vbCrLf & vbCrLf & Err.Source, Err.Description, True
    End If
    
    Exit Function
    
HandleError:
    ProcessError Err.Number, "Reg.Set_Key (" & sNewKeyName & ")" & vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Public Function Set_Value(lngKey As hKey, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Reg_Value) As Boolean
' Description:
'   This Function will set the data field of a value in the Registry
On Error GoTo HandleError
    Dim lRetVal As Long
    Dim hdKey As Long
    Dim sValue As String
    Dim bOpenKey As Boolean
    
    Set_Value = False
    bOpenKey = False
    lRetVal = RegOpenKeyEx(lngKey, sKeyName, 0, KEY_ALL_ACCESS, hdKey)
    If lRetVal <> ERROR_NONE Then
        lRetVal = Set_Key(lngKey, sKeyName)
        If lRetVal = ERROR_NONE Then
            lRetVal = RegOpenKeyEx(lngKey, sKeyName, 0, KEY_ALL_ACCESS, hdKey)
            If lRetVal = ERROR_NONE Then
                bOpenKey = True
            End If
        End If
    Else
        bOpenKey = True
    End If
    
    If lRetVal = ERROR_NONE Then
        Select Case lValueType
            Case NULL_VALUE
                vValueSetting = vValueSetting & Chr(0)
                lRetVal = RegSetValueEx(hdKey, sValueName, 0&, lValueType, vValueSetting, CLng(Len(vValueSetting)))
            Case STRING_VALUE
                sValue = CStr(vValueSetting) & Chr(0)
                lRetVal = RegSetValueEx(hdKey, sValueName, 0&, lValueType, sValue, CLng(Len(sValue)))
            Case BINARY_VALUE
                vValueSetting = vValueSetting & Chr(0)
                lRetVal = RegSetValueEx(hdKey, sValueName, 0&, lValueType, vValueSetting, CLng(Len(vValueSetting)))
            Case DWORD_VALUE
                lRetVal = RegSetValueEx(hdKey, sValueName, 0&, lValueType, CLng(vValueSetting), 4)
        End Select
        If lRetVal = ERROR_NONE Then
            Set_Value = True
        Else ' Error setting the value
            ProcessError lRetVal, "Reg.Set_Value" & vbCrLf & vbCrLf & Err.Source, Err.Description, True
        End If
    Else ' Error from first block of code
        ProcessError lRetVal, "Reg.Set_Value" & vbCrLf & vbCrLf & Err.Source, Err.Description, True
    End If
    
    If bOpenKey Then
        lRetVal = RegCloseKey(hdKey)
        If lRetVal <> ERROR_NONE Then ' Error closing open key
            ProcessError lRetVal, "Reg.Set_Value" & vbCrLf & vbCrLf & Err.Source, Err.Description, True
        End If
    End If
    
    Exit Function
    
HandleError:
    ProcessError Err.Number, "Reg.Set_Value (" & sKeyName & "," & sValueName & "," & vValueSetting & ")" & _
                 vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Public Function Query_Value(lngKey As hKey, sKeyName As String, sValueName As String, Optional sDefault As String = "", Optional ResultType As Reg_Result = STRING_RESULT) As Variant
' Description:
'   This Function will return the data field of a value from the Registry
On Error GoTo HandleError
    Dim lRetVal As Long
    Dim hdKey As Long
    Dim sValue As String
    Dim lngBuffer As Long
    
    Query_Value = ""
    lRetVal = RegOpenKeyEx(lngKey, sKeyName, 0, KEY_QUERY_VALUE, hdKey)
    If lRetVal = ERROR_NONE Then
        sValue = Space(255)
        lngBuffer = Len(sValue)
        lRetVal = RegQueryValueEx(hdKey, sValueName, 0&, 0, sValue, lngBuffer)
        If lRetVal = ERROR_NONE Then
            If lngBuffer <= 0 Then
                Query_Value = ""
            Else
                Query_Value = Left(sValue, lngBuffer - 1)
            End If
        Else ' Error querying value
            ProcessError lRetVal, "Reg.Query_Value" & vbCrLf & vbCrLf & Err.Source, Err.Description, True
            Query_Value = sDefault
        End If
        RegCloseKey (hdKey)
    Else ' Error opening key
        Query_Value = sDefault
        ProcessError cnErrBadRegKey, "Reg.Query_Value", cdErrBadRegKey
    End If
    
'# Make sure variant returned to calling function is in correct format
    Select Case ResultType
        Case BOOLEAN_RESULT
            If UCase(Query_Value) <> "TRUE" Then
                Query_Value = "FALSE"
            End If
        Case INTEGER_RESULT
            If Not IsNumeric(Query_Value) Then
                Query_Value = 0
            End If
    End Select
    
    Exit Function
    
HandleError:
    ProcessError Err.Number, "Reg.Query_Value (" & sKeyName & "," & sValueName & ")" & vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Public Function Query_Key(lngKey As hKey, sKeyName As String) As Boolean
' Description:
'   This Function will return if a key exists in the Registry
On Error GoTo HandleError
    Dim lRetVal As Long
    Dim hdKey As Long
    
    Query_Key = False
    lRetVal = RegOpenKeyEx(lngKey, sKeyName, 0, KEY_QUERY_VALUE, hdKey)
    If lRetVal = ERROR_NONE Then
        Query_Key = True
    End If
    
    Exit Function
    
HandleError:
    ProcessError Err.Number, "Reg.Query_Key (" & sKeyName & ")" & vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Public Function EnumKeys(lngKey As hKey, sKeyName As String) As String()
'Description:
'   This function returns an array of strings containing all keys contained within the one specified
On Error GoTo HandleError
    Dim nPos As Integer
    Dim sEnumName As String
    Dim ReturnArray() As String
    Dim hdKey As Long, lRetVal As Long
    
    lRetVal = RegOpenKeyEx(lngKey, sKeyName, 0, KEY_ALL_ACCESS, hdKey)
    If lRetVal = ERROR_NONE Then
'# Get all the folder lists
        nPos = 0
        Do
            sEnumName = Space$(256)
            If RegEnumKey(hdKey, nPos, sEnumName, Len(sEnumName)) <> 0 Then
                Exit Do
            End If
            sEnumName = Left$(sEnumName, InStr(sEnumName, Chr$(0)) - 1)
            ReDim Preserve ReturnArray(nPos)
            ReturnArray(nPos) = sEnumName
            nPos = nPos + 1
        Loop
        EnumKeys = ReturnArray()
    Else ' Error opening key
        ReDim ReturnArray(0)
        EnumKeys = ReturnArray()
        ProcessError cnErrBadRegKey, "Reg.Query_Value", cdErrBadRegKey
    End If
    
    Exit Function
    
HandleError:
    ReDim ReturnArray(0)
    EnumKeys = ReturnArray()
    ProcessError Err.Number, "Reg.EnumKeys (" & sKeyName & ")" & vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Public Function EnumValues(lngKey As hKey, sKeyName As String) As Variant
   Dim nRet As Long
   Dim hKey As Long
   Dim nMaxValueNameLen As Long
   Dim nMaxValueLen As Long
   Dim nValueNameLen As Long
   Dim nValueLen As Long
   Dim nType As Long
   Dim nIndex As Long
   Dim nStrings As Long
   Dim ValueName As String
   Dim Value As String
   Dim Values() As String
   Dim Results() As String
   Dim i As Long

   ' Open key
   nRet = RegOpenKeyEx(lngKey, sKeyName, 0&, KEY_READ, hKey)
   If nRet = ERROR_SUCCESS Then
   
      ' Get a quick snapshot of what we're facing.
      nRet = RegQueryInfoKey(hKey, vbNullString, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, nMaxValueNameLen, nMaxValueLen, ByVal 0&, ByVal 0&)
      If nRet = ERROR_SUCCESS Then
      
         ' Allocate buffers.
         ValueName = Space(nMaxValueNameLen + 1)
         Value = Space(nMaxValueLen + 1)
         
         ' Get value names and associated values.
         Do
            nValueNameLen = Len(ValueName)
            nValueLen = Len(Value)
            nRet = RegEnumValue(hKey, nIndex, ValueName, nValueNameLen, ByVal 0&, nType, ByVal Value, nValueLen)
            If nRet = ERROR_SUCCESS Then
               ' Only return string values.
               If nType = REG_SZ Then
                  ReDim Preserve Values(0 To 1, 0 To nStrings) As String
                  Values(0, nStrings) = Left$(ValueName, nValueNameLen)
                  Values(1, nStrings) = Left$(Value, nValueLen - 1)
                  nStrings = nStrings + 1
               End If
               nIndex = nIndex + 1
            Else 'ERROR_NO_MORE_ITEMS
               Exit Do
            End If
         Loop
         
         ' Transpose array to match VB's output, and
         ' return Results if any were obtained.
         If nStrings >= 1 Then
            ReDim Results(0 To nStrings - 1, 0 To 1) As String
            For i = 0 To nStrings - 1
               Results(i, 0) = Values(0, i)
               Results(i, 1) = Values(1, i)
            Next i
            EnumValues = Results
         End If
      End If
      
      Call RegCloseKey(hKey)
   End If
    Exit Function
HandleError:
    ProcessError Err.Number, "Reg.EnumValues (" & sKeyName & ")" & vbCrLf & vbCrLf & Err.Source, Err.Description
End Function

Private Sub ProcessError(ErrNumber As Long, ErrSource As String, ErrDescription As String, _
                         Optional InternalRegError As Boolean = False)
    m_ErrorExist = True
    m_ErrorNumber = ErrNumber
    m_ErrorSource = ErrSource
    If InternalRegError Then
        Select Case ErrNumber
            Case ERROR_BADDB
                m_ErrorDescription = MSG_BADDB
            Case ERROR_BADKEY
                m_ErrorDescription = MSG_BADKEY
            Case ERROR_CANTOPEN
                m_ErrorDescription = MSG_CANTOPEN
            Case ERROR_CANTREAD
                m_ErrorDescription = MSG_CANTREAD
            Case ERROR_CANTWRITE
                m_ErrorDescription = MSG_CANTWRITE
            Case ERROR_OUTOFMEMORY
                m_ErrorDescription = MSG_OUTOFMEMORY
            Case ERROR_INVALID_PARAMETER
                m_ErrorDescription = MSG_INVALID_PARAMETER
            Case ERROR_ACCESS_DENIED
                m_ErrorDescription = MSG_ACCESS_DENIED
            Case ERROR_INVALID_PARAMETERS
                m_ErrorDescription = MSG_INVALID_PARAMETERS
            Case ERROR_NO_MORE_ITEMS
                m_ErrorDescription = MSG_NO_MORE_ITEMS
            Case Else
                m_ErrorDescription = MSG_UNKNOWN
        End Select
    Else
        m_ErrorDescription = ErrDescription
    End If
    RaiseEvent ErrorOccurred(m_ErrorNumber, m_ErrorSource, m_ErrorDescription)
End Sub

Public Function SaveSetting(ByVal plHKEY As Long, ByVal psSection As String, ByVal psKey As String, ByVal psSetting As String) As Boolean
    Dim lRet As Long
    Dim lhKey As Long
    Dim lResult As Long
   
    lRet = RegCreateKey(plHKEY, psSection, lhKey)
    If lRet = ERROR_SUCCESS Then
        psSetting = psSetting & vbNullChar
        lRet = RegSetValueEx(lhKey, psKey, 0&, STRING_VALUE, ByVal psSetting, Len(psSetting))
        lRet = RegCloseKey(lhKey)
    End If
    SaveSetting = (lRet = ERROR_SUCCESS)
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
