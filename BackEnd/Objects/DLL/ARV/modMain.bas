Attribute VB_Name = "modMain"

'Pass this one Global Object between Apps
Public goUtil As V2ECKeyBoard.clsUtil

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Property Get msClassName() As String
    msClassName = "modMain"
End Property



