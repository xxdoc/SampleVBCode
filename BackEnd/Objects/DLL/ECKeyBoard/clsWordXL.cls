VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWordXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const DELIM As String = ","
Private Const VARS_DELIM As String = "©"
Private Const VAR_DELIM As String = "®"
Private Const COMMA_DELIM As String = "Þ"
Private Const VBCRLF_DELIM As String = "¶"
Public Enum AvailDocs
    Name = 1
    DateLastUpdated
    DateLastUpdatedSort
    DateCreated
    DateCreatedSort
End Enum

Public Enum SavedDocs
    Name = 1
    Date
    Variables
End Enum

Public Enum Pic
    Word = 1
    XL
    Hourglass
End Enum

Public Type udtSavedDocVar
    a01Name As String
    a02Date As String
    a03Variables As Variant
End Type

Public Type QVariable
    Name As String
    Value As String
End Type

Private moWord As Word.Application
Private mlhWndWord As Long 'Windows handle for Word
Private moDoc As Word.Document

Private moXL As Excel.Application
Private mlhWndXL As Long 'Windows handle for XL
Private moWB As Excel.Workbook

'BGS 1.2.2002 Claim Stuff
Private msDocVarID As String
Private msWordXLDocPath As String
Private msDVRPath As String
Private maryQVariables() As QVariable
Private maryDocVariables() As QVariable

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsWordXL"
End Property

Public Property Get XLIsValid() As Boolean
    XLIsValid = IsValidObject(moXL)
End Property

Public Property Let DocVarID(psID As String)
    msDocVarID = psID
End Property
Public Property Get DocVarID() As String
    DocVarID = msDocVarID
End Property

Public Property Let WordXLDocPath(psPath As String)
    msWordXLDocPath = psPath
End Property
Public Property Get WordXLDocPath() As String
    WordXLDocPath = msWordXLDocPath
End Property

Public Property Let DVRPath(psPath As String)
    msDVRPath = psPath
End Property
Public Property Get DVRPath() As String
    DVRPath = msDVRPath
End Property

Public Property Let aryQVariables(paryQVariables As Variant)
    If Not IsEmpty(paryQVariables) Then
        maryQVariables = paryQVariables
    End If
End Property
Public Property Get aryQVariables() As Variant
    aryQVariables = maryQVariables
End Property

Public Property Let aryDocVariables(paryDocVariables As Variant)
    maryDocVariables = paryDocVariables
End Property
Public Property Get aryDocVariables() As Variant
    aryDocVariables = maryDocVariables
End Property

Public Property Let hWndWord(plHwnd As Long)
    mlhWndWord = plHwnd
End Property
Public Property Get hWndWord() As Long
    hWndWord = mlhWndWord
End Property

Public Property Let hWndXL(plHwnd As Long)
    mlhWndXL = plHwnd
End Property
Public Property Get hWndXL() As Long
    hWndXL = mlhWndXL
End Property

Public Function LoadDocVar(Optional pShowMode As FormShowConstants = vbModeless) As Boolean
    On Error GoTo EH
    Dim MyDocVar As frmDocVar
    Dim lErrNum As Long
    Dim sErrDesc As String
        
    Set MyDocVar = New frmDocVar
    
    If Not FindSetForm(Forms, "frmDocVar", MyDocVar) Then
        Set MyDocVar.WordXL = Me
        Load MyDocVar
    Else
        Set MyDocVar.WordXL = Me
    End If
   
    LoadDocVar = True
    MyDocVar.Show pShowMode

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function LoadDocVar" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function LoadWordXLAPP(pofrmDocVar As Object) As Boolean
    On Error GoTo EH
    Dim MyDocVar As Form
    Set MyDocVar = pofrmDocVar
    Dim bWord As Boolean
    Dim bXL As Boolean
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'BGS 1.7.2002 Check to see if Word and Excel 2000 have been installed.
    'If not can not let them use this form.
    
    With MyDocVar
        If moWord Is Nothing Then
            .lblName.Caption = "Loading MS WORD, Please Wait..."
            .lblName.Refresh
            .imgSelected.Picture = .imgVarDoc.ListImages.Item(Pic.Word).Picture
            .imgSelected.Refresh
            bWord = WordIsInstalled
            If bWord Then
                mlhWndWord = FindWindow(vbNullString, moWord.Name)
            End If
        Else
            bWord = WordIsInstalled
        End If
        If moXL Is Nothing Then
            .lblName.Caption = "Loading MS EXCEL, Please Wait..."
            .lblName.Refresh
            .imgSelected.Picture = .imgVarDoc.ListImages.Item(Pic.XL).Picture
            .imgSelected.Refresh
            bXL = XLIsInstalled
            If bXL Then
                mlhWndXL = FindWindow(vbNullString, moXL.Name)
            End If
        Else
            bXL = XLIsInstalled
        End If
        .lblName.Caption = vbNullString
        .lblName.Refresh
        .imgSelected.Picture = Nothing
        .imgSelected.Refresh
    End With
    
    If bWord And bXL Then
        LoadWordXLAPP = True
    Else
        sMess = "MS Word Excel not found!"
        If bWord Then
            sMess = Replace(sMess, "Word", vbNullString)
        End If
        If bXL Then
            sMess = Replace(sMess, "Excel", vbNullString)
        End If
        MsgBox sMess, vbExclamation + vbOKOnly, "MS OFFICE NOT INSTALLED"
        Exit Function
    End If
    
    'cleannup
    Set MyDocVar = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    AlwaysOnTop MyDocVar, False
    Set MyDocVar = Nothing
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function LoadWordAPP" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not moDoc Is Nothing Then
        If IsValidObject(moDoc) Then
            moDoc.Close False
        End If
        Set moDoc = Nothing
    End If
    If Not moWord Is Nothing Then
        If IsValidObject(moWord) Then
            moWord.Quit
        End If
        Set moWord = Nothing
    End If
    If Not moWB Is Nothing Then
        If IsValidObject(moWB) Then
            moWB.Close False
        End If
        Set moWB = Nothing
    End If
    If Not moXL Is Nothing Then
        If IsValidObject(moXL) Then
            moXL.Quit
        End If
        Set moXL = Nothing
    End If
    Set frmDocVar = Nothing
    CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
   Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanUp" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function LoadaryDocVariables(pDocType As Pic, psDocName As String, _
                                     Optional pbSaved As Boolean, Optional psSaved As String) As Boolean
    'This will Load and Scan WOrd and XL documents for Variables
    On Error GoTo EH
    Dim lSCount As Long 'Section (For Headers and Footers only)
    Dim lFCount As Long 'Fields
    Dim lVarCount As Long 'Variables
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'XL stuff
    Dim oWS As Excel.Worksheet
    Dim oName As Excel.Name
    Dim bFoundSheet As Boolean
    
    Screen.MousePointer = vbHourglass
    
    Erase maryDocVariables
    
    Select Case pDocType
        Case Pic.Word
            If Not moDoc Is Nothing Then
                If IsValidObject(moDoc) Then
                    moDoc.Close False
                End If
                Set moDoc = Nothing
            End If
            Set moDoc = moWord.Documents.Open(goUtil.AttachReposPath & "\" & psDocName)
            
            If Not pbSaved Then
                'BGS 1.3.2002 Sections applies only to Headers and Footers If they happen to exist
                'in the particular document we are scanning for variables
                For lSCount = 1 To moDoc.Sections.Count
                    'Header Variables
                    For lFCount = 1 To moDoc.Sections(lSCount).Headers(wdHeaderFooterPrimary).Range.Fields.Count
                        If moDoc.Sections(lSCount).Headers(wdHeaderFooterPrimary).Range.Fields(lFCount).Type = wdFieldDocVariable Then
                            AddWordVariable moDoc.Sections(lSCount).Headers(wdHeaderFooterPrimary).Range.Fields(lFCount).code.Text, lVarCount
                        End If
                    Next
                    'Footer Variables
                    For lFCount = 1 To moDoc.Sections(lSCount).Footers(wdHeaderFooterPrimary).Range.Fields.Count
                        If moDoc.Sections(lSCount).Footers(wdHeaderFooterPrimary).Range.Fields(lFCount).Type = wdFieldDocVariable Then
                            AddWordVariable moDoc.Sections(lSCount).Footers(wdHeaderFooterPrimary).Range.Fields(lFCount).code.Text, lVarCount
                        End If
                    Next
                Next
                'Main Body Variables
                For lFCount = 1 To moDoc.Fields.Count
                    If moDoc.Fields(lFCount).Type = wdFieldDocVariable Then
                        AddWordVariable moDoc.Fields(lFCount).code.Text, lVarCount
                    End If
                Next
            Else
                AddSavedVariables psSaved, lVarCount
            End If
            
        Case Pic.XL
            If Not moWB Is Nothing Then
                If IsValidObject(moWB) Then
                    moWB.Close False
                End If
                Set moWB = Nothing
            End If
            Set moWB = moXL.Workbooks.Open(goUtil.AttachReposPath & "\" & psDocName)
            
            If Not pbSaved Then
                'Need to loop through SHeets to find the Variable tab
                For Each oWS In moWB.Sheets
                    If oWS.Name = "Variables" Then
                        bFoundSheet = True
                        Exit For
                    End If
                Next
                If bFoundSheet Then
                    For Each oName In moWB.Names
                        If InStr(1, oName.RefersTo, "Variables", vbTextCompare) > 0 Then
                            AddXLVariable oName.Name, lVarCount
                        End If
                    Next
                End If
                
            Else
                AddSavedVariables psSaved, lVarCount
            End If
            
        Case Else
            Exit Function
    End Select
    
    Screen.MousePointer = vbDefault
    If lVarCount > 0 Then
        LoadaryDocVariables = True
    End If
    
    'clean up
    Set oWS = Nothing
    Set oName = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function LoadaryDocVariables" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function FillaryDocVariables(pDocType As Pic, psDocName As String) As Boolean
    'This will Populate WOrd and XL documents with the appropriate Variable Data
    On Error GoTo EH
    Dim lSCount As Long
    Dim lFCount As Long
    Dim sVarName As String
    Dim sVarValue As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'XL stuff
    Dim oWS As Excel.Worksheet
    Dim oName As Excel.Name
    Dim bFoundSheet As Boolean
    
    Select Case pDocType
        Case Pic.Word
            For lSCount = 1 To moDoc.Sections.Count
                'Header Variables
                For lFCount = 1 To moDoc.Sections(lSCount).Headers(wdHeaderFooterPrimary).Range.Fields.Count
                    If moDoc.Sections(lSCount).Headers(wdHeaderFooterPrimary).Range.Fields(lFCount).Type = wdFieldDocVariable Then
                        sVarName = moDoc.Sections(lSCount).Headers(wdHeaderFooterPrimary).Range.Fields(lFCount).code.Text
                        sVarName = GetVariableName(sVarName)
                        sVarValue = QueryVariable(sVarName, True)
                        moDoc.Sections(lSCount).Headers(wdHeaderFooterPrimary).Range.Fields(lFCount).Result.Text = sVarValue
                        FillaryDocVariables = True
                    End If
                Next
                'Footer Variables
                For lFCount = 1 To moDoc.Sections(lSCount).Footers(wdHeaderFooterPrimary).Range.Fields.Count
                    If moDoc.Sections(lSCount).Footers(wdHeaderFooterPrimary).Range.Fields(lFCount).Type = wdFieldDocVariable Then
                        sVarName = moDoc.Sections(lSCount).Footers(wdHeaderFooterPrimary).Range.Fields(lFCount).code.Text
                        sVarName = GetVariableName(sVarName)
                        sVarValue = QueryVariable(sVarName, True)
                        moDoc.Sections(lSCount).Footers(wdHeaderFooterPrimary).Range.Fields(lFCount).Result.Text = sVarValue
                        FillaryDocVariables = True
                    End If
                Next
            Next
            'Main Body Variables
            For lFCount = 1 To moDoc.Fields.Count
                If moDoc.Fields(lFCount).Type = wdFieldDocVariable Then
                    sVarName = moDoc.Fields(lFCount).code.Text
                    sVarName = GetVariableName(sVarName)
                    sVarValue = QueryVariable(sVarName, True)
                    moDoc.Fields(lFCount).Result.Text = sVarValue
                    FillaryDocVariables = True
                End If
            Next
        Case Pic.XL
            'Need to loop through SHeets to find the Variable tab
            For Each oWS In moWB.Sheets
                If oWS.Name = "Variables" Then
                    bFoundSheet = True
                    Exit For
                End If
            Next
            If bFoundSheet Then
                For Each oName In moWB.Names
                    If InStr(1, oName.RefersTo, "Variables", vbTextCompare) > 0 Then
                        sVarName = oName.Name
                        sVarValue = QueryVariable(sVarName, True)
                        oName.RefersToRange = sVarValue
                        FillaryDocVariables = True
                    End If
                Next
                moXL.Calculate
            End If
    End Select
    
    'clean up
    Set oWS = Nothing
    Set oName = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function PopulatearyDocVariables" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Sub AddWordVariable(psCodeText As String, plVarCOunt As Long)
    'This will Add  Word Variable to the VAriable array
    On Error GoTo EH
    Dim MyVariable As QVariable
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    MyVariable.Name = GetVariableName(psCodeText)
    MyVariable.Value = QueryVariable(MyVariable.Name)
    plVarCOunt = plVarCOunt + 1
    ReDim Preserve maryDocVariables(1 To plVarCOunt)
    maryDocVariables(plVarCOunt) = MyVariable
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub AddVariable" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Private Sub AddXLVariable(psName As String, plVarCOunt As Long)
    'This will Add  XL Variable to the VAriable array
    On Error GoTo EH
    Dim MyVariable As QVariable
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    MyVariable.Name = psName
    MyVariable.Value = QueryVariable(MyVariable.Name)
    plVarCOunt = plVarCOunt + 1
    ReDim Preserve maryDocVariables(1 To plVarCOunt)
    maryDocVariables(plVarCOunt) = MyVariable
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub AddVariable" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Private Function GetVariableName(psCodeText As String) As String
    On Error GoTo EH
    Dim sTemp As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    sTemp = psCodeText
    sTemp = Mid(sTemp, InStr(1, sTemp, """", vbBinaryCompare) + 1)
    sTemp = left(sTemp, InStr(1, sTemp, """", vbBinaryCompare) - 1)
    
    GetVariableName = sTemp
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function GetVariableName" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function QueryVariable(psName As String, Optional pbUseDocVariables As Boolean) As String
    On Error GoTo EH
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not pbUseDocVariables Then
        For lCount = LBound(maryQVariables, 1) To UBound(maryQVariables, 1)
            If StrComp(maryQVariables(lCount).Name, psName, vbTextCompare) = 0 Then
                QueryVariable = maryQVariables(lCount).Value
                Exit Function
            End If
        Next
    Else
        If IsArray(maryDocVariables) Then
            For lCount = LBound(maryDocVariables, 1) To UBound(maryDocVariables, 1)
                If StrComp(maryDocVariables(lCount).Name, psName, vbTextCompare) = 0 Then
                    QueryVariable = maryDocVariables(lCount).Value
                    Exit Function
                End If
            Next
        End If
    End If
    If Not pbUseDocVariables Then
        QueryVariable = " "
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function QueryVariable" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function AddSavedVariables(psSaved As String, plVarCOunt As Long) As Boolean
    'This will Add saved variables both Word and Excel to the VAriable array
    On Error GoTo EH
    Dim sVars() As String
    Dim sVar() As String
    Dim MyVariable As QVariable
    Dim lCount As Integer
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sVars = Split(psSaved, VARS_DELIM, , vbBinaryCompare)
    
    For lCount = LBound(sVars, 1) To UBound(sVars, 1)
        sVar = Split(sVars(lCount), VAR_DELIM, , vbBinaryCompare)
        MyVariable.Name = sVar(0)
        MyVariable.Value = sVar(1)
        plVarCOunt = plVarCOunt + 1
        ReDim Preserve maryDocVariables(1 To plVarCOunt)
        maryDocVariables(plVarCOunt) = MyVariable
    Next
    AddSavedVariables = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function AddSavedVariables" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function GetAvail(pvaryAvail As Variant) As Boolean
    On Error GoTo EH
    Dim sReport As String
    Dim saryReports() As String
    Dim iReportCount As Integer
    Dim bFoundAvail As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As Long
    
    'BGS get the XL reports
    sReport = Dir(msWordXLDocPath & "\*.xls")
    Do Until sReport = vbNullString
        bFoundAvail = True
        iReportCount = iReportCount + 1
        ReDim Preserve saryReports(1 To iReportCount)
        saryReports(iReportCount) = sReport
        sReport = Dir
    Loop
    'BGS Get the DOC reports
    sReport = Dir(msWordXLDocPath & "\*.doc")

    Do Until sReport = vbNullString
        bFoundAvail = True
        iReportCount = iReportCount + 1
        ReDim Preserve saryReports(1 To iReportCount)
        saryReports(iReportCount) = sReport
        sReport = Dir
    Loop
   
    If bFoundAvail Then
        pvaryAvail = saryReports
        GetAvail = True
    Else
        GetAvail = False
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetRSAvail" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function GetSaved(pvDocVar As Variant) As Boolean
    On Error GoTo EH
    Dim arySavedDocVar() As udtSavedDocVar
    Dim sarySavedDocVar() As String
    Dim vReports As Variant
    Dim iReport As Integer
    Dim sVariables As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If FileExists(msDVRPath & "\" & msDocVarID & ".dvr") Then
        'BGS get the XL reports
        vReports = Split(GetFileData(msDVRPath & "\" & msDocVarID & ".dvr"), vbCrLf, , vbBinaryCompare)
        If IsArray(vReports) Then
            For iReport = LBound(vReports, 1) To UBound(vReports, 1)
                ReDim Preserve arySavedDocVar(0 To iReport)
                sarySavedDocVar = Split(vReports(iReport), DELIM, , vbBinaryCompare)
                arySavedDocVar(iReport).a01Name = sarySavedDocVar(0)
                arySavedDocVar(iReport).a02Date = sarySavedDocVar(1)
                sVariables = Replace(sarySavedDocVar(2), COMMA_DELIM, ",", , , vbBinaryCompare)
                sVariables = Replace(sVariables, VBCRLF_DELIM, vbCrLf, , , vbBinaryCompare)
                arySavedDocVar(iReport).a03Variables = sVariables
            Next
        End If
    Else
        Exit Function
    End If
    pvDocVar = arySavedDocVar
    GetSaved = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function GetRSSaved" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function StoreSaved(polvwSaved As Object) As Boolean
    On Error GoTo EH
    Dim lvwSaved As ListView
    Dim itmX As listItem
    Dim sarySave() As String
    Dim bSaveME As Boolean
    Dim lCount As Long
    Dim sVariable As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set lvwSaved = polvwSaved
    
    For Each itmX In lvwSaved.ListItems
        bSaveME = True
        lCount = lCount + 1
        ReDim Preserve sarySave(1 To lCount)
        sVariable = Replace(itmX.SubItems(SavedDocs.Variables - 1), vbCrLf, VBCRLF_DELIM)
        sVariable = Replace(sVariable, ",", COMMA_DELIM)
        sarySave(lCount) = itmX.Text & "," & itmX.SubItems(SavedDocs.Date - 1) & "," & sVariable
    Next
    
    'BGS 1.4.2002 Kill the existing .dvr file first
    'the .dvr file contains variables as they existed when document was printed
    If FileExists(msDVRPath & "\" & msDocVarID & ".dvr") Then
        SetAttr msDVRPath & "\" & msDocVarID & ".dvr", vbNormal
        Kill msDVRPath & "\" & msDocVarID & ".dvr"
    End If
    
    If bSaveME Then
        SaveFileData msDVRPath & "\" & msDocVarID & ".dvr", Join(sarySave, vbCrLf)
        StoreSaved = True
    End If
    
    'Clean up
    Set lvwSaved = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function StoreSaved" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function PrintIt(pofrmDocVar As Object, pDocType As Pic, psDocName As String, psDate As String, polblVarName As Object, potxtValue As Object, polvwSaved As Object) As Boolean
    On Error GoTo EH
    Dim lVarCount As Long
    Dim lblVarName As Label
    Dim txtValue As TextBox
    Dim lvwSaved As ListView
    Dim itmX As listItem
    Dim saryVars() As String
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    'If not using the preview for Varibles just use the Loaded Varibles
    If Not polblVarName Is Nothing Then
        'BGS 1.4.2002 need to add the current document which is in Header
        'and framDocVar into Variable array
        For lCount = 1 To potxtValue.UBound
            'BGS if it is locked then it was deleted do not add it
            If Not potxtValue(lCount).Locked Then
                lVarCount = lVarCount + 1
                ReDim Preserve saryVars(1 To lVarCount)
                saryVars(lVarCount) = polblVarName(lCount).Caption & VAR_DELIM & potxtValue(lCount).Text
            End If
        Next
    Else
        
    End If
    
    Select Case pDocType
        Case Pic.Word
            If Not moDoc Is Nothing Then
                If IsValidObject(moDoc) Then
                    If Not potxtValue Is Nothing Then
                        If potxtValue.UBound > 0 Then
                            lVarCount = 0
                            Erase maryDocVariables
                            If AddSavedVariables(Join(saryVars, VARS_DELIM), lVarCount) Then
                                If FillaryDocVariables(pDocType, psDocName) Then
                                    PrintIt = True
                                End If
                            End If
                        End If
                    Else
                        If FillaryDocVariables(pDocType, psDocName) Then
                            PrintIt = True
                        End If
                    End If
                    moWord.WindowState = wdWindowStateNormal
                    moWord.Visible = True
                    moWord.Activate
                End If
            End If
        Case Pic.XL
            If Not moXL Is Nothing Then
                If IsValidObject(moXL) Then
                    If Not potxtValue Is Nothing Then
                        If potxtValue.UBound > 0 Then
                            lVarCount = 0
                            Erase maryDocVariables
                            If AddSavedVariables(Join(saryVars, VARS_DELIM), lVarCount) Then
                                If FillaryDocVariables(pDocType, psDocName) Then
                                    PrintIt = True
                                End If
                            End If
                        End If
                    Else
                        If FillaryDocVariables(pDocType, psDocName) Then
                            PrintIt = True
                        End If
                    End If
                    moXL.WindowState = xlNormal
                    moXL.Visible = True
                    moXL.ActiveWorkbook.Activate
                End If
            End If
    End Select
    'If not using the preview for Varibles just use the Loaded Varibles
    If Not polblVarName Is Nothing Then
        If PrintIt Then
            'BGS 1.15.2002 Lets add it to the Printed even if it has been printed before
            Set itmX = polvwSaved.ListItems.Add(, , psDocName, , pDocType)
            itmX.SubItems(SavedDocs.Date - 1) = Now
            itmX.SubItems(SavedDocs.Variables - 1) = Join(saryVars, VARS_DELIM)
        End If
    End If
    
    'Cleanup
    Set lblVarName = Nothing
    Set txtValue = Nothing
    Set lvwSaved = Nothing
    Set itmX = Nothing
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function PrintIt" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function IsValidObject(poObject As Object) As Boolean
    On Error GoTo EH
    Dim sName As String
    
    If poObject Is Nothing Then
        IsValidObject = False
        Exit Function
    End If
    
    If poObject.Application.WindowState = wdWindowStateNormal Then
        IsValidObject = True
        Exit Function
    ElseIf poObject.Application.WindowState = wdWindowStateMaximize Then
        IsValidObject = True
        Exit Function
    ElseIf poObject.Application.WindowState = xlNormal Then
        IsValidObject = True
        Exit Function
    ElseIf poObject.Application.WindowState = xlMaximized Then
        IsValidObject = True
        Exit Function
    End If

    IsValidObject = True
    Exit Function
EH:
    Err.Clear
    IsValidObject = False
End Function

Private Function WordIsInstalled() As Boolean
    On Error GoTo EH
    
    If Not IsValidObject(moWord) Then
        Set moWord = Nothing
        Set moWord = New Word.Application
    End If
    WordIsInstalled = True
   
    Exit Function
EH:
    Err.Clear
End Function

Private Function XLIsInstalled() As Boolean
    On Error GoTo EH
    
    If Not IsValidObject(moXL) Then
        Set moXL = New Excel.Application
    End If
    XLIsInstalled = True

    Exit Function
EH:
    Err.Clear
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

