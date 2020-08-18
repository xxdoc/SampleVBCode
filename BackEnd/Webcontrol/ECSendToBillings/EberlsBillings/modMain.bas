Attribute VB_Name = "modMain"
Option Explicit

Public Const NULL_DATE As String = "12:00:00 AM"
Public Const S_z As String = "¶Ññ"

'Pass this one Global Object between Apps
Public goUtil As V2ECKeyBoard.clsUtil
Public gofrmMain As Object

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
