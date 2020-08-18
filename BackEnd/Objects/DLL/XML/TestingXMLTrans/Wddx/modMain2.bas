Attribute VB_Name = "modMain2"
Option Explicit

'Turn off and On Grid lines in ListView
Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1
Public Const CLASS_PREFIX As String = "cls"
Public Const CLASS_MAX_LEN As Long = 30
Public Const WEB_REFRESH_ERROR As Long = -2147467259

'11.27.2002 Page Break Flag
Public Const INSERT_PAGE_BREAK As String = "——∂Ò"

'BGS 10.11.2001 clsUpload Special chars
Public Const F_DELIM As String = "ﬁ"
Public Const F_VBCRLF As String = "∂"
'ACCESS
Public Const NULL_DATE As String = "12:00:00 AM"

'BGS 11.20.2001 NO_SSN
Public Const NO_SSN As Long = 999999999

Public Const INVALID_DB_PASSWORD_KEY As String = "Invalid Key!"

'BGS 4.7.2002 Used In Building SQL INSERT STATEMENT
Public Const S_z As String = "∂—Ò" '"""" ' Begin SQL String Field
Public Const z_S As String = "Ò—∂" '""", " ' End SQL String Field
Public Const S_z_SET As String = """"
Public Const z_S_SET As String = """, "

Public Const COLUMN_DELIM As String = "ﬁ"
Public Const RECORD_DELIM As String = "∂"
'Pass this one Global Object between Apps
Public goUtil As V2ECKeyBoard.clsUtil

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

