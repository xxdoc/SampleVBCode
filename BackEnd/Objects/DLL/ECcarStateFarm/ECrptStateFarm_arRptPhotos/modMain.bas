Attribute VB_Name = "modMain"

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

Private Property Get msClassName() As String
    msClassName = "modMain"
End Property

Public Function GetPreviousParamsCol(psMiscDelimSettings As String, psClassName As String) As Collection
    On Error GoTo EH
    'Begin Items For Assignments MiscDelimSettings
    Dim sMiscDelimSettings As String
    Dim saryDelim() As String
    Dim lCountDelim As Long
    Dim sDelimItem As String
    Dim saryColumn() As String
    Dim lCountColumn As Long
    Dim sColumnName As String
    Dim sColumnValue As String
    Dim saryCVItems() As String
    Dim lCountCVItems As Long
    Dim sCVData As String
    Dim saryCVDataItems() As String
    Dim lPosCVDataItems As Long
    Dim sCVName As String
    Dim sCVValue As String
    Dim saryValue() As String
    Dim MyParams As rptMiscDelimParam
    'End Itmes For Assignments MiscDelimSettings
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sMiscDelimSettings = Trim(psMiscDelimSettings)
    
    If sMiscDelimSettings <> vbNullString Then
        saryDelim() = Split(sMiscDelimSettings, RECORD_DELIM)
        For lCountDelim = LBound(saryDelim, 1) To UBound(saryDelim, 1)
            sDelimItem = saryDelim(lCountDelim)
            If sDelimItem <> vbNullString Then
                saryColumn() = Split(sDelimItem, COLUMN_DELIM)
                For lCountColumn = LBound(saryColumn, 1) To UBound(saryColumn, 1)
                    sColumnName = saryColumn(lCountColumn)
                    If sColumnName <> vbNullString Then
                        If StrComp(sColumnName, psClassName & "_DELIMPARAMS", vbTextCompare) = 0 Then
                            sColumnValue = saryColumn(lCountColumn + 1)
                            saryCVItems() = Split(sColumnValue, "^")
                            For lCountCVItems = LBound(saryCVItems, 1) To UBound(saryCVItems, 1)
                                sCVData = saryCVItems(lCountCVItems)
                                If sCVData <> vbNullString Then
                                    saryCVDataItems() = Split(sCVData, "|")
                                    For lPosCVDataItems = LBound(saryCVDataItems, 1) To UBound(saryCVDataItems, 1)
                                        sCVName = saryCVDataItems(lPosCVDataItems)
                                        If sCVName <> vbNullString Then
                                            sCVValue = saryCVDataItems(lPosCVDataItems)
                                            saryValue() = Split(sCVValue, "=")
                                            With MyParams
                                                If StrComp(saryValue(0), "ClassName", vbTextCompare) = 0 Then
                                                    .ClassName = saryValue(1)
                                                ElseIf StrComp(saryValue(0), "ParamCaption", vbTextCompare) = 0 Then
                                                    .ParamCaption = saryValue(1)
                                                ElseIf StrComp(saryValue(0), "ParamDataType", vbTextCompare) = 0 Then
                                                    .ParamDataType = saryValue(1)
                                                ElseIf StrComp(saryValue(0), "ParamName", vbTextCompare) = 0 Then
                                                    .ParamName = saryValue(1)
                                                ElseIf StrComp(saryValue(0), "ParamValue", vbTextCompare) = 0 Then
                                                    .ParamValue = saryValue(1)
                                                ElseIf StrComp(saryValue(0), "ProjectName", vbTextCompare) = 0 Then
                                                    .ProjectName = saryValue(1)
                                                End If
                                            End With
                                        End If
                                    Next
                                    If GetPreviousParamsCol Is Nothing Then
                                        Set GetPreviousParamsCol = New Collection
                                    End If
                                    GetPreviousParamsCol.Add MyParams, MyParams.ParamName
                                End If
                            Next
                            GoTo BAIL_HERE
                        End If
                    End If
                Next
            End If
        Next
    End If
BAIL_HERE:

Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Set GetPreviousParamsCol = Nothing
    Err.Raise lErrNum, , sErrDesc & vbCrLf & msClassName & vbCrLf & "Public Function GetPreviousParamsCol"
End Function

Public Function RemoveParam(psName As String, pColParams As Collection) As Boolean
    On Error GoTo EH
    If Not pColParams Is Nothing Then
        pColParams.Remove psName
        RemoveParam = True
    End If
    Exit Function
EH:
    Err.Clear
End Function

Public Function GetParam(psName As String, pColParams As Collection) As Variant
    On Error GoTo EH
    
    If Not pColParams Is Nothing Then
        GetParam = pColParams(psName)
    Else
        GetParam = psName & ": Parameter collection not set!"
    End If
    
    Exit Function
EH:
    GetParam = vbNullString
    Err.Clear
End Function







