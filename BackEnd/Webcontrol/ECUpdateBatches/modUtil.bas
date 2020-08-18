Attribute VB_Name = "modUtil"
Option Explicit
'Used in Loss Reports List View
Public Enum LossReports
    DateAsgn = 1
    Adjuster
    ClaimNo
    InsuredName
    HPhone
    WPhone
    RFormat
    RSort ' Hidden
    RKey 'Hidden
End Enum

'Used in Loss Reports Print Options
Public Enum PrintFormat
    RawText = 0
    Translated
End Enum

'Used in LossReports Image list
'This is the only Thing that will have to be updated
'When Adding more Report formats. Add a Picture to the imglist
'The actual refrences to the pic will be made by the
Public Enum LRPic
    lrError = 1
    lrASN
    lrCCMS
    lrUnknown
    Found
    SearchColumn
    lrHTMLHelp
End Enum

'Used in Adjuster List View
Public Enum ADJ
    CRID = 1
    FACT
    EMAIL
    ADJUpdateEmail
    ADJFName 'FirstName
    ADJLName 'LastName
    ADJSSN
    ADJPassword
    ADJContactPhone
    ADJTeamLeader
    ADJDateLastUpdated
    ADJLicDaysLeft
    ADJAPPVSInfo 'Application VS info
    Dirty 'Hidden
    RSort ' Hidden
    RKey 'Hidden
End Enum

'Used in Adjuster Image list
Public Enum ADJPic
    Pointer = 1
    Found
    SearchColumn
    DeleteMe
    UnDelete
End Enum


Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1
Public Const CLASS_PREFIX As String = "cls"
Public Const CLASS_MAX_LEN As Long = 30
Public Const WEB_REFRESH_ERROR As Long = -2147467259

Public gbUCText As Boolean
Public gARV As ARViewer.clsARViewer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Property Get ModName() As String
    ModName = App.EXEName & ".modUtil.bas"
End Property


Public Function FileExists(strFile As String, Optional pbDirOnly As Boolean) As Boolean
    On Error GoTo EH
    If strFile <> vbNullString Then
        If Not pbDirOnly Then
            FileExists = Dir(strFile, vbHidden) <> vbNullString
        Else
            SaveFileData strFile & "\Dir.tmp", "Dir"
            FileExists = Dir(strFile, vbDirectory) <> vbNullString
            SetAttr strFile & "\Dir.tmp", vbNormal
            Kill strFile & "\Dir.tmp"
        End If
    End If
    
    Exit Function
EH:
    FileExists = False
End Function


Public Function GetFileData(psFilePath As String, Optional pbLock As Boolean = False, Optional piFFile As Integer, Optional pbSkipMess As Boolean = True) As String
'Purpose:

'Parameters :

'Returns:

'Author :

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    Dim lMyFileLen As Long
    Dim iFFile As Long
    
    iFFile = FreeFile
    piFFile = iFFile
    If pbLock Then
        Open psFilePath For Binary Access Read Lock Read As #iFFile
    Else
        Open psFilePath For Binary Access Read As #iFFile
    End If
    lMyFileLen = FileLen(psFilePath) + 2
    GetFileData = Input(lMyFileLen, #iFFile)
    If Not pbLock Then
        Close #iFFile
    End If
    
    Exit Function
EH:
    Close #iFFile
    If Not pbSkipMess Then
        If MsgBox("Could not read file... " & vbCrLf & psFilePath & vbCrLf & "(" & Err.Description & ")" & vbCrLf & vbCrLf & _
                  "The network or file is busy." & vbCrLf & "Press ""Yes"" to try again." & vbCrLf & "Press ""No"" to abort this process", vbYesNo, "File is Busy") = vbYes Then
            Resume
        End If
    End If
    
End Function

Public Sub SaveFileData(psFilePath As String, psFileData As String, Optional psDelimeter As String, Optional pbLock As Boolean = False, Optional piFFile As Integer)
'Purpose:

'Parameters :

'Returns:

'Author :

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    Dim lMyFileLen As Long
    Dim iFFile As Integer
    
    iFFile = FreeFile
    piFFile = iFFile
    Open psFilePath For Binary Access Write As #iFFile
    Put #iFFile, 1, psFileData & psDelimeter
    If Not pbLock Then
        Close #iFFile
    End If
    Exit Sub
EH:
    Close #iFFile
    Err.Raise Err.Number, , "Public Sub SaveFileData" & vbCrLf & vbCrLf & Err.Description & vbCrLf & psFilePath
End Sub


Public Function isControlArray(MyForm As Form, Mycontrol As Control) As Boolean
    
    'BGS 8/1/1999 Added this function to determin if a Control is part of
    'a control array or not. I had to do this because VB does not have a
    'function that figures this out. (IsArray does not work on Control Arrays)

    On Error GoTo EH
    Dim MyCount As Integer
    Dim CheckMyControl As Control
    
    For Each CheckMyControl In MyForm.Controls
        If CheckMyControl.Name = Mycontrol.Name Then
            MyCount = MyCount + 1
            If MyCount > 1 Then
                Exit For
            End If
        End If
    Next
    
    isControlArray = MyCount - 1
    
    Exit Function
EH:
    ShowError Err, "isControlArray", , ModName
    
End Function

Public Sub SelText(pTextBox As Control)
'Purpose: Highlights All Text

'Parameters : TextBox

'Returns: Just Highlights TExt in the pTextBox

'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    With pTextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    Exit Sub
EH:
    Err.Clear
    
End Sub

Public Function FormWinRegPos(pMyForm As Form, Optional pbSave As Boolean, _
                              Optional pfrmOffset As Form, Optional pctrlOffset As Control, _
                              Optional pbUseFullCaption As Boolean = True) As Boolean
'Purpose: This Procedure can be used by AnyForm to Get or Save the Form Position
'         from the Windows Registry using Save Setting and GetSetting :)

'Parameters : pMyForm As Form, Optional pbSave As Boolean

'Returns: FormWinRegPos Returns True Only if  Retrieving and Finds Stored  Values
'         FormWinRegPos Returns False if Retreieving and does not find Stored Values
'         FormWinRegPos Returns False When Saveing IE pbSave is Set to True.


'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description
'                     1     BGS         3/21/2000   Added Optional pfrOffset incase you want to Offset the posn
'                                                   in realation to another form.
'                     2     BGS         10/23/2001  Changed the SECTION to enter ALL forms under FORM_POSN
'                                                   Also check for Borderstyle do Change width or height on non sizable windows :)
                    
    'This Procedure will Either Retrieve or Save Form Posn values
    'Best used on Form Load and Unload or QueryUnLoad
    Dim sCap As String
    
    On Error GoTo EH
    With pMyForm
        If Not pbUseFullCaption Then
            sCap = .Name
        Else
            sCap = .Caption & " "
        End If
        If pbSave Then
            'If Saving then do this...
            'If Form was minimized or Maximized then Closed Need to Save Windowstate
            'THEN... set Back to Normal Or previous non Max or Min State then Save
            'Posn Parameters
            
            SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_WindowState", .WindowState
            
            If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
                .WindowState = vbNormal
            End If
            
            'Save AppName...FrmName...KeyName...Value
            If pfrmOffset Is Nothing Then
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Top", .top
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Left", .left
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Height", .Height
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Width", .Width
            Else
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Top", .top - pfrmOffset.top - pctrlOffset.top
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Left", .left - pfrmOffset.left
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Height", .Height
                SaveSetting App.EXEName, "FORM_POSN", .Name & sCap & "_Width", .Width
            End If
        Else
            'If Not Saveing Must Be Getting ..
            'Need to ref AppName...FrmName...KeyName
            '(If nothing Stored Use The Exisiting Form value)
            If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
                .WindowState = vbNormal
            End If
            If pfrmOffset Is Nothing Then
                .top = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Top", .top)
                .left = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Left", .left)
                If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
                    .Height = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Height", .Height)
                    .Width = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Width", .Width)
                End If
                'Be Sure WindowState is set last (Can't Change POSN if vbMinimized Or Maximized)
                .WindowState = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_WindowState", .WindowState)
            Else
                .top = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Top", .top) + pfrmOffset.top + pctrlOffset.top
                .left = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Left", .left) + pfrmOffset.left
                If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
                    .Height = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Height", .Height)
                    .Width = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_Width", .Width)
                End If
                'Be Sure WindowState is set last (Can't Change POSN if vbMinimized Or Maximized)
                .WindowState = GetSetting(App.EXEName, "FORM_POSN", .Name & sCap & "_WindowState", .WindowState)
            End If
        End If
    End With

    
    Exit Function
EH:
    ShowError Err, "FormWinRegPos", , ModName
    
End Function


Public Sub UCText(pControl As Control)
    On Error GoTo EH
    Dim iSelpos As Integer
    If Not gbUCText Then
        gbUCText = True
        iSelpos = pControl.SelStart
        With pControl
            .Text = UCase(.Text)
            .SelStart = iSelpos
            .SelLength = 0
        End With
        gbUCText = False
    End If
    
    Exit Sub
EH:
    ShowError Err, "UCText", , ModName
End Sub


Public Function DBTblExists(psDBPath As String, psDBTbl As String) As Boolean
    On Error GoTo EH
    Dim WS As Workspace
    ' SourceVariables
    Dim dbSource As Database
    Dim tblSource As TableDef
    Dim sTemp As String
    
    'BGS 1.10.2001 Need to be sure the default directory is
    'Ap.path or will get strange errors when creating WorkSpace
    ChDir App.Path
    Set WS = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbSource = WS.OpenDatabase(psDBPath, False, True)
    
    For Each tblSource In dbSource.TableDefs
        sTemp = tblSource.Name
        If InStr(1, psDBTbl, sTemp, vbTextCompare) > 0 Then
            DBTblExists = True
            GoTo CLEANUP
        End If
    Next
    DBTblExists = False
CLEANUP:
    Set tblSource = Nothing
    dbSource.Close: Set dbSource = Nothing
    WS.Close: Set WS = Nothing
    
    Exit Function
EH:
    Set tblSource = Nothing
    Set dbSource = Nothing
    Set WS = Nothing
    ShowError Err, "DBTblExists", , ModName
End Function


Public Function DynamicArraySet(pVarArray As Variant) As Boolean
    'Purpose: To see if a Dynamic array has been set
    'Parameters : pVarArray As Variant: Send in any Dynamic array data type
    'Returns: True if has been set, false if not
    'Author : BGS-3/24/2000
    'Revision History:  SMR     Initials    Date    Description
    
    On Error GoTo NOT_SET
    Dim iRet As Integer
    
    If IsArray(pVarArray) Then
        iRet = LBound(pVarArray, 1)
        'if the Lbound call to the first dimension of
        'pVarArray does not error then the dynamic array must
        'be set so...
        DynamicArraySet = True
        Exit Function
    End If
    
NOT_SET:
    DynamicArraySet = False
End Function

Public Sub CompRepair(psAppEXEName As String, psSourcePath As String, Optional psPass As String, Optional pbSkipBackup As Boolean)
    On Error GoTo EH
    Dim sTemp As String
    Dim sDBBackup
    Dim sDBName As String
    
    'BGS 12.15.2000 Make the Temp Same Path as the Original
    sTemp = left(psSourcePath, InStrRev(psSourcePath, "\") - 1) & "\DB.tmp"
    sDBBackup = left(psSourcePath, InStrRev(psSourcePath, ".") - 1)
    sDBBackup = sDBBackup & "_BackUp.mdb"
    sDBName = Mid(psSourcePath, InStrRev(psSourcePath, "\") + 1)
    'Kill Temp  if it there for some reason
    If FileExists(sTemp) Then
        If MsgBox("There were system problems the last time you launched the " & App.EXEName & "." & vbCrLf & _
               "The " & sDBName & " should be restored with " & vbCrLf & sDBBackup & "." & vbCrLf & _
               "Press OK to Restore or Cancel to Exit " & App.EXEName & ".", vbOKCancel) = vbOK Then
               Kill sTemp
               FileCopy sDBBackup, psSourcePath
        Else
'            End 'BAIL !!!
        End If
    End If
    'Copy the DB to the Temp PAth
    FileCopy psSourcePath, sTemp
    
    'BGS 12.13.2000 Make Backup Data base
    'In case there is some problems
    If Not pbSkipBackup Then
        FileCopy psSourcePath, sDBBackup
    End If
    
    
    'Kill the Source since it is copied to Temp
    Kill psSourcePath
    
    'Compact the Temp and send it back to where it was originally
    'copied from
    If Len(psPass) > 0 Then
        CompactDatabase sTemp, psSourcePath, dbLangGeneral & ";PWD=" & psPass, , ";PWD=" & psPass
    Else
        CompactDatabase sTemp, psSourcePath
    End If
    
    'Repair the just compacted database
'    RepairDatabase psSourcePath
    
    'Finally kill the Temp since we done with it
    Kill sTemp
    
    Exit Sub
EH:
    ShowError Err, "CompRepair", , "modUtil"
    'if for some reason we errored while compacting and repairing
    'See if we can recover the Source back to what it was before
    'If this fails the Data Base we were trying to compact was really
    'Hosed up.
    If Not FileExists(psSourcePath) Then
        If FileExists(sTemp) Then
            FileCopy sTemp, psSourcePath
            Kill sTemp
        End If
    End If
End Sub


Public Sub ShowError(pobjErr As ErrObject, psProc As String, Optional pFormOwner As Form, Optional psMod As String)
    Dim oError As ECKeyBoard.clsUtil
    
    Set oError = New ECKeyBoard.clsUtil
    oError.utShowError App.EXEName, pobjErr, psProc, pFormOwner, psMod
    Set oError = Nothing

End Sub


Public Sub BubbleSort(pvArray As Variant, Optional psEndString As String, Optional pbRebillPresent As Boolean)
    Dim lMainLoop As Long
    Dim lSubLoop As Long
    Dim sStringA As String
    Dim sStringB As String
    Dim sTempA As String
    Dim sTempB As String
    
    If DynamicArraySet(pvArray) Then
        For lMainLoop = UBound(pvArray) To LBound(pvArray) Step -1
            For lSubLoop = LBound(pvArray) + 1 To lMainLoop
                'Only sort non nullstrings
                If pvArray(lSubLoop) <> vbNullString Then
                    'BGS 11.26.2001 check for Rebill
                    If InStr(1, pvArray(lSubLoop), "R", vbTextCompare) > 0 Then
                        pbRebillPresent = True
                    End If
                    If psEndString > vbNullString Then
                        'BGS 7.12.2001 need to sort with S as "A" because the
                        'Supplements have precendence over rebilling in the sort
                        sStringA = left(pvArray(lSubLoop - 1), InStr(1, pvArray(lSubLoop - 1), psEndString) - 1)
                        sTempA = Right(sStringA, 6)
                        sStringA = Replace(sStringA, sTempA, vbNullString)
                        sTempA = Replace(sTempA, "S", "A", , , vbTextCompare)
                        sStringA = sStringA & sTempA
                        
                        sStringB = left(pvArray(lSubLoop), InStr(1, pvArray(lSubLoop), psEndString) - 1)
                        sTempB = Right(sStringB, 6)
                        sStringB = Replace(sStringB, sTempB, vbNullString)
                        sTempB = Replace(sTempB, "S", "A", , , vbTextCompare)
                        sStringB = sStringB & sTempB
                        
                        If sStringA > sStringB Then
                            Call SwitchPlace(pvArray(lSubLoop - 1), pvArray(lSubLoop))
                        End If
                    Else
                        If pvArray(lSubLoop - 1) > pvArray(lSubLoop) Then
                            Call SwitchPlace(pvArray(lSubLoop - 1), pvArray(lSubLoop))
                        End If
                    End If
                End If
            Next
        Next
    End If
    
End Sub

Public Sub SwitchPlace(a As Variant, b As Variant)
  Dim c As Variant
  c = a
  a = b
  b = c
End Sub

Public Function FormExists(psFormName As String) As Boolean
    On Error GoTo EH
    Dim iCount As Integer
    
    For iCount = 0 To Forms.Count - 1
        If UCase(Forms(iCount).Name) = UCase(psFormName) Then
            FormExists = True
            Exit For
        End If
    Next
    Exit Function
EH:
     ShowError Err, "Public Function FormExists", , ModName
End Function

Public Function FindSetForm(psFormName As String, pForm As Form) As Boolean
    On Error GoTo EH
    Dim iCount As Integer
    
    For iCount = 0 To Forms.Count - 1
        If UCase(Forms(iCount).Name) = UCase(psFormName) Then
            Set pForm = Forms(iCount)
            FindSetForm = True
            Exit For
        End If
    Next
    Exit Function
EH:
     ShowError Err, "Public Function FindSetForm", , ModName
End Function

Public Function CleanString(pvText As Variant) As Variant
    On Error GoTo EH
    'BGS 11.21.2001 add some more cleaning
    If VarType(pvText) = vbString Then
        CleanString = Replace(pvText, "'", "''")
'        CleanString = Replace(CleanString, vbCrLf, " ")
    Else
        CleanString = pvText
    End If
    Exit Function
EH:
    CleanString = pvText
    ShowError Err, "Public Function CleanString", , ModName
End Function

Public Function CleanValString(psValText As String) As String
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
  ShowError Err, "Public Function CleanValString", , ModName
End Function

Public Sub UpdateAddress(psAddress As String, _
                         psZip As String, _
                         psState As String, _
                         psCity As String, _
                         psStreet As String)
    On Error GoTo EH
    Dim sTemp As String
    Dim sAddress As String
    
    sTemp = Trim(Replace(psStreet, vbCrLf, vbNullString)) & String(2, " ") & vbCrLf
    sAddress = sTemp
    sTemp = Trim(Replace(psCity, vbCrLf, vbNullString))
    sTemp = Replace(sTemp, ",", vbNullString) & ", "
    sAddress = sAddress & sTemp
    sTemp = Trim(Replace(psState, vbCrLf, vbNullString)) & " "
    sAddress = sAddress & sTemp
    sTemp = Trim(Replace(psZip, vbCrLf, vbNullString))
    sAddress = sAddress & sTemp
    
     
    If Right(sAddress, 5) = vbCrLf & ", " & " " Then
        On Error Resume Next
        sAddress = RTrim(left(sAddress, InStrRev(sAddress, vbCrLf) - 1))
        Dim l As Long
'        For l = 1 To Len(sAddress)
'            Debug.Print Mid(sAddress, l, 1) & " ---->" & Asc(Mid(sAddress, l, 1))
'        Next
    End If
    
    psAddress = sAddress
    Exit Sub
EH:
    ShowError Err, "Public Sub UpdateAddress", , ModName
End Sub

Public Sub FillAddressFields(psAddress As String, _
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
                    sAddress = Trim(left(sAddress, lPos))
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
                sAddress = Trim(left(sAddress, lPos))
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
                sAddress = Trim(left(sAddress, lPos))
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
    ShowError Err, "Public Sub FillAddressFields", , ModName

End Sub

Public Function ValidSSN(psSSN As String) As String
    'Returns "Error" if invliad SSN
    On Error GoTo EH
    Dim sSSN As String
    
    sSSN = psSSN

    sSSN = Val(CleanValString(sSSN))
    If Len(sSSN) < 8 Or Len(sSSN) > 9 Then
        sSSN = "Error"
    End If
    
    ValidSSN = sSSN
    
    Exit Function
EH:
    ShowError Err, "Public Function ValidSSN", , ModName
End Function
        
Public Function ValidDate(psDate As String) As String
    On Error GoTo EH
    'will retunrn "12:00:00 AM" if invalid
    Dim sDate As String
    
    sDate = psDate
    
    If IsDate(sDate) Then
        'BGS 5.8.2002 Check to see if they entered like 12:00AM
        'it will default to year 1899 All dates enterd should be at least
        '1900 and above. Unless we have over 100 year old claims.. Hmm you think ?
        If Format(sDate, "YYYY") > 1900 Then
            sDate = Format(sDate, "MM/DD/YYYY")
        Else
            sDate = NULL_DATE
        End If
    Else
        sDate = NULL_DATE
    End If
    
    ValidDate = sDate
    
    Exit Function
EH:
    ShowError Err, "Public Function ValidDate", , ModName
End Function
    
Public Function GetAppVSInfo(psAppEXEName As String, psAppPath As String) As String
    On Error GoTo EH
    Dim vOPaths As Variant
    Dim lCount As Long
    Dim sFI As String
    Dim sText As String
    Dim sEXE As String
    Dim vEXE As Variant
    Dim colEXE As Collection
    Dim oFI As ECKeyBoard.clsFileVersion
    Dim FI As ECKeyBoard.FILE_INFORMATION
    
    sEXE = Dir(psAppPath & "\*.exe")
    Do Until sEXE = vbNullString
        If colEXE Is Nothing Then
            Set colEXE = New Collection
        End If
        colEXE.Add sEXE, sEXE
        sEXE = Dir
    Loop
    
    Set oFI = New ECKeyBoard.clsFileVersion
    
    If Not colEXE Is Nothing Then
        For Each vEXE In colEXE
            sEXE = vEXE
            FI = oFI.GetFileInformation(psAppPath & "\" & sEXE)
            sText = sText & "(" & FI.cFilename & ") "
            sText = sText & "VS " & FI.nVerMajor & "." & FI.nVerMinor & "." & FI.nVerRevision & " "
            sText = sText & FI.dtLastModifyTime & vbCrLf
        Next
    End If
    
    GetAppVSInfo = sText
    
    vOPaths = GetECSWinSysObjectsPaths
    If IsArray(vOPaths) Then
        For lCount = LBound(vOPaths, 1) To UBound(vOPaths, 1)
            FI = oFI.GetFileInformation(vOPaths(lCount))
            sFI = sFI & "(" & FI.cFilename & ") "
            sFI = sFI & "VS " & FI.nVerMajor & "." & FI.nVerMinor & "." & FI.nVerRevision & " "
            sFI = sFI & FI.dtLastModifyTime & vbCrLf
        Next
        GetAppVSInfo = GetAppVSInfo & sFI
    End If
    
    'Cleanup
    If Not oFI Is Nothing Then
        Set oFI = Nothing
    End If
    Set colEXE = Nothing
  Exit Function
EH:
    If Not oFI Is Nothing Then
        Set oFI = Nothing
    End If
    ShowError Err, "Public Function GetAppVSInfo", , ModName

End Function
    
Public Function GetECSWinSysObjectsPaths() As Variant
    On Error GoTo EH
    Dim sBaseDLL As String
    Dim sDLL As String
    Dim sBaseOCX As String
    Dim sOCX As String
    Dim saryObjectPaths()
    Dim lCount As Long
    Dim bFound As Boolean
    
    If FileExists("C:\WINDOWS\SYSTEM\ECS\DLL", True) Then
        sBaseDLL = "C:\WINDOWS\SYSTEM\ECS\DLL"
        sBaseOCX = "C:\WINDOWS\SYSTEM\ECS\OCX"
    ElseIf FileExists("C:\WINDOWS\SYSTEM32\ECS\DLL", True) Then
        sBaseDLL = "C:\WINDOWS\SYSTEM32\ECS\DLL"
        sBaseOCX = "C:\WINDOWS\SYSTEM32\ECS\OCX"
    ElseIf FileExists("C:\WINNT\SYSTEM32\ECS\DLL", True) Then
        sBaseDLL = "C:\WINNT\SYSTEM32\ECS\DLL"
        sBaseOCX = "C:\WINNT\SYSTEM32\ECS\OCX"
    End If

    If FileExists(sBaseDLL, True) Then
        sDLL = Dir(sBaseDLL & "\*.dll")
        If sDLL > vbNullString Then
            bFound = True
        End If
        Do Until sDLL = vbNullString
            lCount = lCount + 1
            ReDim Preserve saryObjectPaths(1 To lCount)
            saryObjectPaths(lCount) = sBaseDLL & "\" & sDLL
            sDLL = Dir
        Loop
    End If
    
    If FileExists(sBaseOCX, True) Then
        sOCX = Dir(sBaseOCX & "\*.ocx")
        If sOCX > vbNullString Then
            bFound = True
        End If
        Do Until sOCX = vbNullString
            lCount = lCount + 1
            ReDim Preserve saryObjectPaths(1 To lCount)
            saryObjectPaths(lCount) = sBaseOCX & "\" & sOCX
            sOCX = Dir
        Loop
    End If
            
    GetECSWinSysObjectsPaths = saryObjectPaths
    Exit Function
EH:
    ShowError Err, "Public Function GetECSWinSysObjectsPaths", , ModName

End Function

Public Function FindCBOItem(psSearchText As String, pCBO As ComboBox, piPos As Integer) As String
    On Error GoTo EH
    Dim iCount As Integer
    
    If pCBO.ListCount > 0 Then
        For iCount = 0 To pCBO.ListCount - 1
            If InStr(1, left(pCBO.List(iCount), piPos), psSearchText, vbTextCompare) > 0 Then
                FindCBOItem = pCBO.List(iCount)
                Exit Function
            End If
        Next
        'BGS if we did not find the search item then pass it back what was sent in
        FindCBOItem = psSearchText
    End If
    Exit Function
EH:
    ShowError Err, "Public Function FindCBOItem", , ModName
End Function

Public Function Validate(Optional poForm As Object, Optional poControl As Object) As Boolean
    On Error GoTo EH
    Dim Mycontrol As Control
    Dim sValidMess As String
    Dim lPos As Long
    
    Validate = True
    If poControl Is Nothing Then
        For Each Mycontrol In poForm.Controls
            ValidControl Mycontrol, sValidMess, Validate
        Next
    Else
        ValidControl poControl, sValidMess, Validate
    End If
    
    If Not Validate Then
        MsgBox "Please fix the following value(s)..." & vbCrLf & vbCrLf & sValidMess, vbExclamation + vbOKOnly, "Validation"
    End If
    
    Exit Function
EH:
    ShowError Err, "Public Function Validate", , ModName
End Function

Private Sub ValidControl(pMyControl As Object, psValidMess As String, pbValidate)
    On Error GoTo EH
    Dim lPos As Long
    
    'Numeric validation
        If InStr(1, pMyControl.Tag, "Numeric", vbTextCompare) > 0 Then
            If Not IsNumeric(pMyControl.Text) Then
                If pMyControl.Text = vbNullString Then
                    pMyControl.Text = 0
                Else
                    psValidMess = psValidMess & pMyControl.Text & " Is not a valid amount!" & vbCrLf
                    pbValidate = False
                End If
            Else
                If CDbl(pMyControl.Text) < 0 Then
                    psValidMess = psValidMess & pMyControl.Text & " Is not a valid amount!" & vbCrLf
                    pbValidate = False
                End If
            End If
        End If
        
        'Percent Validation
        If InStr(1, pMyControl.Tag, "Percent", vbTextCompare) > 0 Then
            If Not IsNumeric(pMyControl.Text) Then
                If pMyControl.Text = vbNullString Then
                    pMyControl.Text = "0.000"
                Else
                    psValidMess = psValidMess & pMyControl.Text & " Is not a valid percent!" & vbCrLf
                    pbValidate = False
                End If
            Else
                lPos = InStr(1, pMyControl.Text, ".", vbBinaryCompare)
                If CDbl(pMyControl.Text) > 100 Or CDbl(pMyControl.Text) < 0 Then
                    psValidMess = psValidMess & pMyControl.Text & " Is not a valid percent!" & vbCrLf
                    pbValidate = False
                ElseIf lPos > 0 And InStr(1, pMyControl.Tag, "TaxPercent", vbTextCompare) > 0 Then
                    'issue 178 Force Taxes for Texas, New Mexico, and West Virginia
                    'The percent will need to go 3 digits to right of decimal for all
                    'states not just Texas New Mexico West virginia.
                    If Len(Mid(pMyControl.Text, lPos + 1)) < 3 Then
TAX_PERCENT:
                        psValidMess = psValidMess & pMyControl.Text & " Tax Percent must have at least 3 decimal places!" & vbCrLf
                        pbValidate = False
                    End If
                ElseIf lPos = 0 And InStr(1, pMyControl.Tag, "TaxPercent", vbTextCompare) > 0 Then
                    If Val(pMyControl.Text) <> 0 Then
                        GoTo TAX_PERCENT
                    End If
                End If
            End If
        End If
        
    Exit Sub
EH:
    ShowError Err, "Private Sub ValidControl", , ModName
End Sub

Public Sub EnterUserPass(psSECTION As String, psUserName As String, pvPass As Variant, _
                         Optional poForm As Object)
    On Error GoTo EH
    Dim sCryptUserName As String
    Dim sCryptPass As String
    Dim sRet As String
    
    Dim oUtil As ECKeyBoard.clsUtil
    
    Set oUtil = New ECKeyBoard.clsUtil
    'If pass is control that has text then we are
    'prompting for confirmation of the old password first if it exists
    'and then re enter the new password
    If IsObject(pvPass) Then
        sCryptPass = GetECSCryptSetting("ECS", psSECTION, "PASSWORD")
        If sCryptPass <> vbNullString Then
            If Not poForm Is Nothing Then
                'PUt input box top left of form
                sRet = InputBox("Please enter old password.", "OLD PASSWORD", , poForm.left, poForm.top)
                If sRet = vbNullString Then
                    'they clicked on cancel
                    GoTo CLEANUP
                End If
            Else
                'Put input box default windows pos
                sRet = InputBox("Please enter old password.", "OLD PASSWORD")
                If sRet = vbNullString Then
                    'they clicked on cancel
                    GoTo CLEANUP
                End If
            End If
            'check sret against the old password
            If StrComp(sCryptPass, sRet, vbBinaryCompare) <> 0 Then
                MsgBox "The password you entered does not match.", vbOKOnly + vbExclamation, "INCORRECT PASSWORD"
                GoTo CLEANUP
            Else
                'Save the Old Pass word
                SaveECSCryptSetting "ECS", psSECTION, "OLD_PASSWORD", sCryptPass
                'Need to reset password on server when connecting to server
                SaveSetting "ECS", psSECTION, "RESET_PASSWORD", True
                GoTo ENTER_NEWPASS
            End If
            
        Else
            'if there is no password saved yet then just ask for
            'Password
ENTER_NEWPASS:
            sRet = InputBox("Please enter a new password.", "ENTER NEW PASSWORD", , poForm.left, poForm.top)
            If sRet = vbNullString Then
                'they clicked on cancel
                'Need to undo reset password on server when connecting to server
                SaveSetting "ECS", psSECTION, "RESET_PASSWORD", False
                GoTo CLEANUP
            Else
                sCryptPass = sRet
                'Ask them to double check the password they just entered
                sRet = InputBox("Please enter the same password again.", "ENTER PASSWORD AGAIN", , poForm.left, poForm.top)
                If sRet = vbNullString Then
                    'they clicked on cancel
                    'Need to undo reset password on server when connecting to server
                    SaveSetting "ECS", psSECTION, "RESET_PASSWORD", False
                     GoTo CLEANUP
                Else
                    If StrComp(sCryptPass, sRet, vbBinaryCompare) <> 0 Then
                        MsgBox "The password you entered does not match.", vbOKOnly + vbExclamation, "INCORRECT PASSWORD"
                        'Need to undo reset password on server when connecting to server
                        SaveSetting "ECS", psSECTION, "RESET_PASSWORD", False
                        GoTo CLEANUP
                    End If
                End If
            End If
            
            'Disaplay password
            pvPass.Text = sCryptPass
        End If
    Else
        sCryptPass = CStr(pvPass)
    End If
    'set the user Name
    sCryptUserName = psUserName
    
    If sCryptPass <> vbNullString Then
        SaveECSCryptSetting "ECS", psSECTION, "PASSWORD", sCryptPass
    End If
    If sCryptUserName <> vbNullString Then
        SaveECSCryptSetting "ECS", psSECTION, "USERNAME", sCryptUserName
    End If
    
CLEANUP:
    Set oUtil = Nothing
    Exit Sub
EH:
    ShowError Err, "Public Sub EnterPass", , ModName
End Sub

Public Function GetECSCryptSetting(psAPP As String, psSECTION As String, psKEY As String, _
                                   Optional pvDefault As Variant = vbNullString) As Variant
    On Error GoTo EH
    Dim sCryptSetting As String
    Dim oUtil As ECKeyBoard.clsUtil
    
    Set oUtil = New ECKeyBoard.clsUtil
    
    sCryptSetting = GetSetting(psAPP, psSECTION, psKEY, vbNullString)
    
    If sCryptSetting <> vbNullString Then
        GetECSCryptSetting = CStr(oUtil.Decode(sCryptSetting))
    Else
        GetECSCryptSetting = pvDefault
    End If
    
    Set oUtil = Nothing
    Exit Function
EH:
    ShowError Err, "GetECSCryptSetting", , ModName
End Function

Public Sub SaveECSCryptSetting(psAPP As String, psSECTION As String, psKEY As String, psSetting As String)
    On Error GoTo EH
    Dim sCryptSetting As String
    Dim oUtil As ECKeyBoard.clsUtil
    
    sCryptSetting = psSetting
    
    Set oUtil = New ECKeyBoard.clsUtil
    
    sCryptSetting = oUtil.Encode(sCryptSetting)
    
    SaveSetting psAPP, psSECTION, psKEY, sCryptSetting
    
    Set oUtil = Nothing
    Exit Sub
EH:
    ShowError Err, "Public Sub SaveECSCryptSetting", , ModName
End Sub

