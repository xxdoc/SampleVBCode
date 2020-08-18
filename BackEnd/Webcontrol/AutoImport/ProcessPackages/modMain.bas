Attribute VB_Name = "modMain"
Option Explicit
'BGS 10.11.2001 clsUpload Special chars
Public Const F_DELIM As String = "ﬁ"
Public Const F_VBCRLF As String = "∂"
'ACCESS
Public Const NULL_DATE As String = "12:00:00 AM"

'BGS 11.20.2001 NO_SSN
Public Const NO_SSN As Long = 999999999

'BGS 4.7.2002 Used In Building SQL INSERT STATEMENT
Public Const S_z As String = "∂—Ò" '"""" ' Begin SQL String Field
Public Const z_S As String = "Ò—∂" '""", " ' End SQL String Field
Public Const DT_z As String = "—Ò∂" 'Used to indicate SQL Server Date quote
Public Const S_z_SET As String = "'"
Public Const z_S_SET As String = "', "

Public Const COLUMN_DELIM As String = "ﬁ"
Public Const COLUMN_DELIM_REP As String = "§"
Public Const RECORD_DELIM As String = "∂"
Public Const RECORD_DELIM_REP As String = "•"
Private mclsARViewer As V2ARViewer.clsARViewer


Public Property Let LetARv(pARV As V2ARViewer.clsARViewer)
    Set mclsARViewer = pARV
End Property

Public Property Get GetARV() As V2ARViewer.clsARViewer
    Set GetARV = mclsARViewer
End Property

Public Function GetDateQuote() As String
    On Error GoTo EH
    Dim sDSN As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sDSN = GetSetting("V2WebControl", "DSN", "NAME", "ACCESS_2000")
    
    If sDSN = "ACCESS_2000" Then
        GetDateQuote = "#"
    Else
        GetDateQuote = DT_z
    End If

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & "Public Function GetDateQuote"
End Function
