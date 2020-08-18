VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Used in arDeResPhoto
Public Type udtDeResPhoto
    sBuildPath As String
    sPhotoPath As String
    colParameters As Collection
End Type

Public Enum PhotoSettings
    Optimal_H = 5250
    Optimal_W = 7050
End Enum

'--------Tokens Files FTP to Server request for info----'
'Always Add New Items to Bottom of the Enum... otherwise
'You will have to Break Compatibility and update All EXE and DLL that
'Reference this Object
Public Enum TokenType
    Security = 0
    GetLookUp
End Enum

'Used in Web security
'Always Add New Items to Bottom of the Enum... otherwise
'You will have to Break Compatibility and update All EXE and DLL that
'Reference this Object, have to Recompile this Object and Rereference this
'Object and recompile in those items as well.
Public Enum SecurityToken
    TokenType = 0       '0
    Carrier             '1
    UserName            '2
    SSN                 '3
    Pass                '4
    OldPass             '5
    ResetPass           '6
    LicDaysLeft         '7
    AppVSInfo           '8
    Email               '9
    TeamLeader          '10
    ContactPhone        '11
    FName               '12
    LName               '13
    IBPrefix            '14
    Company             '15
    sEmergencyPhone     '16
    sAddress            '17
    sCity               '18
    sState              '19
    iZip                '20
    iZip4               '21
    sOtherPostCode      '22
End Enum

'Changes made to this Type will Require Breaking Compatiblity
Public Type TokenInfo
    iTokenType As TokenType
    sCarrier As String
    sPath As String
    sToken As String
End Type

Private moExport As clsLossReports
Private moARReport As Object

Public Property Let ARReport(poAR As Object)
    Set moARReport = poAR
End Property
Public Property Set ARReport(poAR As Object)
    Set moARReport = poAR
End Property
Public Property Get ARReport() As Object
    Set ARReport = moARReport
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsLists"
End Property

Public Function GetARReport(psReportName As String) As Object
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Select Case UCase(psReportName)
        Case "ARGENTEXT"
            Set GetARReport = New V2ECKeyBoard.arGenText
    End Select
    
    'Set ref to this Lists Instance
    Set GetARReport.Lists = Me
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function GetARReport"
End Function

Public Function DeResPhoto(pudtDeResPhoto As V2ECKeyBoard.udtDeResPhoto) As Boolean

    On Error GoTo EH
    Dim sBuildPath As String
    Dim sCmdStr As String
    Dim sNewFileName As String
    Dim sIrfanViewPath As String
    Dim sIrfanViewIniPath As String
    Dim sIrfanViewIniData As String
    Dim aryIrfanViewIni() As String
    Dim lIniCount As Long
    Dim bEditJpeg As Boolean
    Dim bEditTiff As Boolean
    Dim bEditSave As Boolean
    Dim bFoundJpeg As Boolean
    Dim bFoundTiff As Boolean
    Dim bFoundSave As Boolean
    Dim lActiveFileTry As Long
    'Parameters in pudtDeResPhoto.colParameters
    Dim pbInitINI As Boolean
    Dim plReSizePercent As Long
    Dim plPhotoQuality As Long
    Dim plOptimalPixelWidth As Long
    Dim plOptimalPixelHeight As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Populate the Parameters
    pbInitINI = pudtDeResPhoto.colParameters("pbInitINI")
    plReSizePercent = pudtDeResPhoto.colParameters("plReSizePercent")
    plPhotoQuality = pudtDeResPhoto.colParameters("plPhotoQuality")
    plOptimalPixelWidth = pudtDeResPhoto.colParameters("plOptimalPixelWidth")
    plOptimalPixelHeight = pudtDeResPhoto.colParameters("plOptimalPixelHeight")
    
'****************USe IrfanView to handle photo resize and quality*****************
    sIrfanViewPath = App.Path & "\" & "i_view32.exe"
    sIrfanViewIniPath = App.Path & "\" & "i_view32.ini"
    
    sIrfanViewPath = LCase(sIrfanViewPath)
    sIrfanViewIniData = goUtil.utGetFileData(sIrfanViewIniPath)
    aryIrfanViewIni() = Split(sIrfanViewIniData, vbCrLf)
    
    If pbInitINI Then
        'Edit the IniFile
        For lIniCount = LBound(aryIrfanViewIni, 1) To UBound(aryIrfanViewIni, 1)
            sIrfanViewIniData = aryIrfanViewIni(lIniCount)
            If InStr(1, sIrfanViewIniData, "[JPEG]", vbTextCompare) > 0 Then
                bEditJpeg = True
                bFoundJpeg = True
                bEditTiff = False
                bEditSave = False
            ElseIf InStr(1, sIrfanViewIniData, "[TIFF]", vbTextCompare) > 0 Then
                bEditJpeg = False
                bEditTiff = True
                bFoundTiff = True
                bEditSave = False
            ElseIf InStr(1, sIrfanViewIniData, "[SAVE]", vbTextCompare) > 0 Then
                bEditJpeg = False
                bEditTiff = False
                bEditSave = True
                bFoundSave = True
            ElseIf InStr(1, sIrfanViewIniData, "[", vbTextCompare) > 0 Then
                bEditJpeg = False
                bEditTiff = False
                bEditSave = False
            End If
            If bEditJpeg Then
                If InStr(1, sIrfanViewIniData, "Save Quality=", vbTextCompare) > 0 Then
                    sIrfanViewIniData = "Save Quality=" & plPhotoQuality
                End If
            End If
            If bEditTiff Then
                If InStr(1, sIrfanViewIniData, "Save Compression=", vbTextCompare) > 0 Then
                    sIrfanViewIniData = "Save Compression=7"
                End If
            End If
            If bEditSave Then
                If InStr(1, sIrfanViewIniData, "SaveOldDate=", vbTextCompare) > 0 Then
                    sIrfanViewIniData = "SaveOldDate=1"
                End If
            End If
            aryIrfanViewIni(lIniCount) = sIrfanViewIniData
        Next
        sIrfanViewIniData = Join(aryIrfanViewIni, vbCrLf)
        'See if the ini entries were even found
        'if not add them
        If Not bFoundJpeg Then
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "[JPEG]"
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "Save Quality=" & plPhotoQuality
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "Save Progressive=0"
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "Save Grayscale=0"
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "KeepExif=1"
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "KeepCom=1"
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "KeepIptc=1"
        End If
        If Not bFoundTiff Then
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "[TIFF]"
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "Save Compression=7"
        End If
        If Not bFoundSave Then
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "[Save]"
            sIrfanViewIniData = sIrfanViewIniData & vbCrLf & "SaveOldDate=1"
        End If
        
        'Check to see if any other process is Locking the
        'INI file.  Only check for maximum of 5 seconds
        lActiveFileTry = 0
        Do Until Not goUtil.utActiveFile(sIrfanViewIniPath, lActiveFileTry)
            If lActiveFileTry > 5 Then
                Exit Do
            End If
            DoEvents
            Sleep 1000
        Loop
        goUtil.utDeleteFile sIrfanViewIniPath
        goUtil.utSaveFileData sIrfanViewIniPath, sIrfanViewIniData
    End If
    
    sBuildPath = pudtDeResPhoto.sBuildPath
    Screen.MousePointer = vbHourglass
    sNewFileName = "0.jpg"
    sCmdStr = """" & sIrfanViewPath & """ "
    sCmdStr = sCmdStr & LCase(pudtDeResPhoto.sPhotoPath) & " "
    'Add 20% to the Width and Height of the Optimal for Source Photo
    'Since it Will not be uploaded we can afford it
    sCmdStr = sCmdStr & "/resize=(" & plOptimalPixelWidth * 1.2 & "," & plOptimalPixelHeight * 1.2 & ") /aspectratio /convert="
    sCmdStr = sCmdStr & LCase(sBuildPath) & sNewFileName
'    ShellBat "Temp0", sCmdStr
    Shell sCmdStr, vbHide
    DoEvents
    Sleep 200
    
     'Allow the User to adjust the Resize Percentage
    sNewFileName = "1.jpg"
    sCmdStr = """" & sIrfanViewPath & """ "
    sCmdStr = sCmdStr & LCase(pudtDeResPhoto.sPhotoPath) & " "
    sCmdStr = sCmdStr & "/resize=(" & plReSizePercent & "p," & plReSizePercent & "p) /aspectratio /convert="
    sCmdStr = sCmdStr & LCase(sBuildPath) & sNewFileName
'    ShellBat "Temp1", sCmdStr
    Shell sCmdStr, vbHide
    DoEvents
    Sleep 200
    
     sNewFileName = "2.jpg"
    sCmdStr = """" & sIrfanViewPath & """ "
    sCmdStr = sCmdStr & LCase(pudtDeResPhoto.sPhotoPath) & " "
    'Make the Thumb 20% of the Optimal Pixels width
    sCmdStr = sCmdStr & "/resize=(" & plOptimalPixelWidth * 0.2 & "," & plOptimalPixelHeight * 0.2 & ") /aspectratio /convert="
    sCmdStr = sCmdStr & LCase(sBuildPath) & sNewFileName
'    ShellBat "Temp2", sCmdStr
    Shell sCmdStr, vbHide
    DoEvents
    Sleep 200
    
    Screen.MousePointer = vbDefault
    
    DeResPhoto = True

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function DeResPhoto"
End Function

Private Sub ShellBat(psBatName As String, psBatData As String)
    On Error GoTo EH
    Dim sBatFile As String
    Dim lActiveFileTry As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'See if Temp dir Exisits
    If Not goUtil.utFileExists("C:\Temp", True) Then
        goUtil.utMakeDir "C:\Temp"
    End If
    sBatFile = "C:\Temp\" & psBatName & ".bat"
    
    If goUtil.utFileExists(sBatFile) Then
        'Only check for maximum of 5 seconds
        lActiveFileTry = 0
        Do Until Not goUtil.utActiveFile(sBatFile, lActiveFileTry)
            If lActiveFileTry > 5 Then
                Exit Do
            End If
            Sleep 1000
        Loop
        goUtil.utDeleteFile sBatFile
    End If
    goUtil.utSaveFileData sBatFile, psBatData
    
    Shell sBatFile, vbHide
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ShellBat"
End Sub

Public Function CFExportARReport(psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If Not moARReport Is Nothing Then
        CFExportARReport = moARReport.ExportME(psXportPath, pXportType)
    End If
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
     Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CFExportARReport"
End Function

Public Function ExportARReport(poARReport As Object, psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim oLRs As V2ECKeyBoard.clsLossReports
    Dim oAR As Object
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not poARReport Is Nothing Then
        Set moARReport = poARReport
    End If
    If Not moARReport Is Nothing Then
        Set oLRs = New V2ECKeyBoard.clsLossReports
    Else
        Exit Function
    End If
    With moARReport
        .Printer.Orientation = ddOPortrait
    End With
    
    If oLRs.ExportFile(moARReport, psXportPath, pXportType) Then
        ExportARReport = True
    End If
    
    'Clean up
    Unload moARReport
    Set moARReport = Nothing
    oLRs.CLEANUP
    Set oLRs = Nothing
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ExportME"
End Function

Public Function PrintListView(psAppEXEName As String, polvw As Object, psTitle As String, _
                              Optional pPrtOrientation As DDActiveReports.PrtOrientation = ddOPortrait, _
                              Optional pShowMode As VBRUN.FormShowConstants = vbModeless, _
                              Optional plPageBreakAfterXItems As Long = 0, _
                              Optional pbSelectedOnly As Boolean = False) As Boolean
    On Error GoTo EH
    Dim LVW As ListView
    Dim ItemX As listItem
    Dim ColHead As ColumnHeader
    Dim lMaxHeadLen As Long
    Dim lMaxItemLen As Long
    Dim lMaxLen As Long
    Dim lTitleLen As Long
    Dim sTitleLineLeft As String
    Dim sTitleLineRight As String
    Dim sTitle As String
    Dim sBody As String
    Dim sTextData As String
    Dim lItemCount As Long
    Dim MyGenText As arGenText
    Dim lSelected As Long
    Dim bSelectedOnly
    Dim sPrintDevice As String
    Dim sPrintPort As String
    Dim lPageBreakAfterXItems As Long
    Dim lPageLineCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    lPageBreakAfterXItems = plPageBreakAfterXItems
    
    Set LVW = polvw
    
    Screen.MousePointer = vbHourglass
    'Get the Max Header Len
    For Each ColHead In LVW.ColumnHeaders
        If lMaxHeadLen < Len(ColHead.Text) Then
            lMaxHeadLen = Len(ColHead.Text)
        End If
    Next
    
    'Add "..." at the end of the Header
    lMaxHeadLen = lMaxHeadLen + Len("... ")
    
    'Get the Max Item Length
    For Each ItemX In LVW.ListItems
        'See If there are any Selected items
        If ItemX.Selected Then
            lSelected = lSelected + 1
        End If
        For Each ColHead In LVW.ColumnHeaders
            'Only look at Visible columns
            If ColHead.Width > 0 Then
                If ColHead.Index = 1 Then
                    If lMaxItemLen < Len(ItemX.Text) Then
                        lMaxItemLen = Len(ItemX.Text)
                    End If
                Else
                    If lMaxItemLen < Len(ItemX.ListSubItems(ColHead.Index - 1).Text) Then
                         lMaxItemLen = Len(ItemX.ListSubItems(ColHead.Index - 1).Text)
                    End If
                End If
            End If
        Next
    Next
    
    'Build the Title
    If pPrtOrientation = ddODefault Or pPrtOrientation = ddOPortrait Then
        lTitleLen = 88
    ElseIf pPrtOrientation = ddOLandscape Then
        lTitleLen = 118
    End If
    lMaxLen = lMaxHeadLen + lMaxItemLen
    sTitleLineLeft = String((lTitleLen - Len(psTitle)) / 2, "-")
    sTitleLineRight = String(((lTitleLen - Len(psTitle)) / 2) - Len(Now()) - 2, "-")
    
    sTitle = vbCrLf & sTitleLineLeft & psTitle & sTitleLineRight & Now()
    
    'Check the Selected Only flag, There must be at least one item selected
    'For the FLag to apply.
    If pbSelectedOnly Then
        If lSelected > 0 Then
            bSelectedOnly = True
        End If
    End If
    
     'If the Page Break is 0 then
    'Figure out the Page Break Automatically
    If lPageBreakAfterXItems = 0 Then
        'Build the Body
        For Each ItemX In LVW.ListItems
            lPageLineCount = lPageLineCount + 1             'Line for Number 1.
            For Each ColHead In LVW.ColumnHeaders
                'Only look at Visible columns
                If ColHead.Width > 0 Then
                    lPageLineCount = lPageLineCount + 1     'Line For Each Sub Item
                End If
            Next
            lPageLineCount = lPageLineCount + 1             'Line for Space before next Item
            Exit For
        Next
        If lPageLineCount > 0 Then
            lPageBreakAfterXItems = 71 / lPageLineCount
        End If
    End If
    
    lItemCount = 0
    'Build the Body
    For Each ItemX In LVW.ListItems
        'Check if the Selected Only flag is true
        If bSelectedOnly Then
            If Not ItemX.Selected Then
                GoTo NEXT_ITEM
            End If
        End If
        'Check for Page Breaks
        If lItemCount > 0 Then
            If lPageBreakAfterXItems > 0 Then
                If lItemCount Mod lPageBreakAfterXItems = 0 Then
                    sBody = sBody & INSERT_PAGE_BREAK & vbCrLf
                    sBody = sBody & sTitle & vbCrLf & vbCrLf
                End If
            End If
        End If
        lItemCount = lItemCount + 1
        sBody = sBody & lItemCount & "." & vbCrLf
        For Each ColHead In LVW.ColumnHeaders
            'Only look at Visible columns
            If ColHead.Width > 0 Then
                If ColHead.Index = 1 Then
                    sBody = sBody & ColHead.Text & String(lMaxHeadLen - Len(ColHead.Text) - 1, ".") & vbTab & ItemX.Text & vbCrLf
                Else
                    sBody = sBody & ColHead.Text & String(lMaxHeadLen - Len(ColHead.Text) - 1, ".") & vbTab & ItemX.ListSubItems(ColHead.Index - 1).Text & vbCrLf
                End If
            End If
        Next
        sBody = sBody & vbCrLf & vbCrLf
NEXT_ITEM:
    Next
    
    sTextData = sTitle & vbCrLf & vbCrLf
    sTextData = sTextData & sBody
    
    Set MyGenText = New arGenText
    With MyGenText
        On Error Resume Next
        sPrintDevice = GetSetting(psAppEXEName, "PRINTER", "PRINTER_NAME", vbNullString)
        If sPrintDevice <> vbNullString Then
            .Printer.DeviceName = sPrintDevice
            If Err.Number > 0 Then
                Err.Clear
            End If
        End If
        
        sPrintPort = GetSetting(psAppEXEName, "PRINTER", "PRINTER_PORT", vbNullString)
        If sPrintPort <> vbNullString Then
            .Printer.Port = sPrintPort
            If Err.Number > 0 Then
                Err.Clear
            End If
        End If
        
        .Printer.Orientation = pPrtOrientation
        On Error GoTo EH
        If pPrtOrientation = ddOLandscape Then
            .PrintWidth = 15540
            .fText.Width = 15500
        End If
        .Detail.Height = 75
        .fText.Height = 165
        .fText.Font.Name = "Courier New"
        .fText.Font.Size = 11
        .TextData = sTextData
    End With
    
    Screen.MousePointer = vbDefault
    
    DoEvents
    Sleep 500
    MyGenText.Run True  ' = Run Asynchronous
    
    Set goUtil.gARV.objARvReport = MyGenText
    goUtil.gARV.SetUtilObject goUtil
    goUtil.gARV.sRptTitle = psTitle
    goUtil.gARV.ShowReport pShowMode
    
    
    
    PrintListView = True
     
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    PrintListView = False
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function PrintListView"
End Function

Public Function ViewImage(psFilePath) As Boolean
    On Error GoTo EH
    Dim sCmdStr As String
    Dim sIrfanViewPath As String
    Dim lErrNum As Long
    Dim sErrDesc As String

    Screen.MousePointer = vbHourglass
    
    sIrfanViewPath = App.Path & "\" & "i_view32.exe"
    
    sCmdStr = """" & sIrfanViewPath & """ "
    sCmdStr = sCmdStr & LCase(psFilePath) & " "
    Shell sCmdStr, vbHide
    
    Screen.MousePointer = vbDefault
    ViewImage = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Screen.MousePointer = vbDefault
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ViewImage"
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

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
     Set moExport = Nothing
    If Not moARReport Is Nothing Then
        Unload moARReport
        Set moARReport = Nothing
    End If
    
    CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CLEANUP"
End Function

