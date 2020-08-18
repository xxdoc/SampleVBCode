VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6765
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin DDActiveReportsViewerCtl.ARViewer ARv 
      Height          =   5060
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8916
      SectionData     =   "frmPreview.frx":030A
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu miFOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu miFSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu miFsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Begin VB.Menu miFXLSExport 
            Caption         =   "E&XCEL Export"
         End
         Begin VB.Menu sep00 
            Caption         =   "-"
         End
         Begin VB.Menu miFHTMLExport 
            Caption         =   "&HTML Export"
            Begin VB.Menu miFHTMLExportNow 
               Caption         =   "&Export to HTML"
            End
            Begin VB.Menu miFHTMLExportJPG 
               Caption         =   "JPG Export &Quality%"
            End
         End
         Begin VB.Menu sep0 
            Caption         =   "-"
         End
         Begin VB.Menu miFPDFExport 
            Caption         =   "P&DF Export"
            Begin VB.Menu miFPDFExportNow 
               Caption         =   "&Export to PDF"
            End
            Begin VB.Menu miFPDFExportJPG 
               Caption         =   "JPG Export Quality%"
            End
         End
         Begin VB.Menu sep1 
            Caption         =   "-"
         End
         Begin VB.Menu miFRTFExport 
            Caption         =   "&RTF Export"
         End
         Begin VB.Menu Sep2 
            Caption         =   "-"
         End
         Begin VB.Menu miFTXTExport 
            Caption         =   "&Text Export"
         End
      End
      Begin VB.Menu miFsep2 
         Caption         =   "-"
      End
      Begin VB.Menu miFPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu miPrinterSetup 
         Caption         =   "Printe&r Setup"
      End
      Begin VB.Menu miFsep3 
         Caption         =   "-"
      End
      Begin VB.Menu miFExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mARViewer As V2ARViewer.clsARViewer
Private mPDF As ActiveReportsPDFExport.ARExportPDF
Private mRTF As ActiveReportsRTFExport.ARExportRTF
Private mTXT As ActiveReportsTextExport.ARExportText
Private mXLS As ActiveReportsExcelExport.ARExportExcel
Private mHTML As ActiveReportsHTMLExport.HTMLexport
Private mbLoadCompleted As Boolean

Public Property Let ARViewer(poARV As V2ARViewer.clsARViewer)
    Set mARViewer = poARV
End Property
Public Property Set ARViewer(poARV As V2ARViewer.clsARViewer)
    Set mARViewer = poARV
End Property
Public Property Get ARViewer() As V2ARViewer.clsARViewer
    Set ARViewer = mARViewer
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub ARv_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    If Not mARViewer.HidePrintButton Then
        Select Case KeyCode
            Case vbKeyF
                Me.PopupMenu mnuFile
            Case vbKeyO
                miFOpen_Click
            Case vbKeyS
                miFSave_Click
            Case vbKeyE
               Me.PopupMenu mnuExport
            Case vbKeyP
                If MsgBox("Print " & Me.Caption & "?", vbYesNo + vbQuestion, "Print Now ?") = vbYes Then
                    miFPrint_Click
                End If
            Case vbKeyR
                miPrinterSetup_Click
            Case vbKeyEscape
                Unload Me
        End Select
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ARv_KeyDown"
End Sub

Private Sub ARv_LoadCompleted()
    mbLoadCompleted = True
End Sub

Private Sub ARv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    If Not mARViewer.HidePrintButton Then
        If Button = vbRightButton Then
            Me.PopupMenu mnuFile
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub ARv_MouseUp"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    With mARViewer
        Me.Caption = .sRptTitle
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , False
    End With
    
    'Check for Hide Print button
    If mARViewer.HidePrintButton Then
        mnuFile.Enabled = False
        ARv.ToolBar.Tools.Item(2).Visible = False
    Else
        mnuFile.Enabled = True
        ARv.ToolBar.Tools.Item(2).Visible = True
    End If
    
    'Set the Export HTML and PDF JPQ quality defaults...
    miFHTMLExportJPG.Caption = "JPG Export Quality%(" & GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "HTML_JPG", 100) & ")"
    miFPDFExportJPG.Caption = "JPG Export Quality%(" & GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "PDF_JPG", 100) & ")"
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim lcount As Long

    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , False
    Me.Visible = False
    For lcount = 1 To 50
'        DoEvents
        Sleep 100
        If mbLoadCompleted Then
            Exit For
        End If
    Next
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    'BGS 12.12.2000 This is the best
    'Solution to Resizeing , using the On Error resume next
    On Error Resume Next
    With ARv
        .Width = Me.Width - 170
        .Height = Me.Height - 700
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    
    Set mARViewer = Nothing
    Set ARv.ReportSource = Nothing
 
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Public Sub RunReport(pobjRpt As Object)
    On Error GoTo EH
    'BGS 8.18.2000 IF the object is nothing
    'let the Viewer load without a report.
    'The user can then open a saved report :)
    If Not pobjRpt Is Nothing Then
        Set ARv.ReportSource = pobjRpt
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub RunReport"
End Sub

Public Function OpenFile(pError As tErr) As Boolean
    On Error GoTo EH
    
    If Not ARv.ReportSource Is Nothing Then
        Set ARv.ReportSource = Nothing
    End If
    
    'BGS 8.11.2000 Need to Open File and
    'set the Caption to the File name and
    'Posn form appropriately
    With mARViewer
        ARv.Pages.Load .sOpenFile
        Me.Caption = .sOpenFile
        goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , False
    End With
    OpenFile = True
    Exit Function
EH:
    OpenFile = False
    With pError
       .Description = Err.Description
       .HelpContext = Err.HelpContext
       .HelpFile = Err.HelpFile
       .Number = Err.Number
       .Source = Err.Source
    End With
    
End Function

Private Sub miFHTMLExportJPG_Click()
    On Error GoTo EH
    Dim sRet As String
    Dim lJPGqual As Long
    'Get new setting from user
    lJPGqual = GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "HTML_JPG", 100)
    sRet = InputBox("Please enter 1 - 100", "HTML JPG Export Quality Percentage", lJPGqual)
    'Check to be sure it is valid
    sRet = Replace(sRet, "%", vbNullString)
    If IsNumeric(sRet) Then
        lJPGqual = sRet
        If lJPGqual >= 1 And lJPGqual <= 100 Then
            SaveSetting App.EXEName, "EXPORT_PHOTO_QUALITY", "HTML_JPG", lJPGqual
        End If
    End If
    miFHTMLExportJPG.Caption = "JPG Export Quality%(" & GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "HTML_JPG", 100) & ")"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFHTMLExportJPG_Click"
End Sub

Private Sub miFHTMLExportNow_Click()
    Set mHTML = New ActiveReportsHTMLExport.HTMLexport
    mHTML.JPEGQuality = GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "HTML_JPG", 100)
    ExportFile mHTML, "HTML Document", "html"
End Sub

Private Sub miFOpen_Click()
    Dim sOpen As SelectedFile
    Dim sDir As String
    Dim sMyFilter As String
    Dim sInitDir As String
    On Error GoTo EH
    
    sMyFilter = sMyFilter & "Report Document File (*.rdf)" & SD & "*.rdf" & SD
    FileDialog.sFilter = sMyFilter
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES
    FileDialog.sDlgTitle = "Open Active Report File"
    sInitDir = GetSetting(App.EXEName, "Dir", "Init", "Error")
    If sInitDir <> "Error" Then
        FileDialog.sInitDir = sInitDir
    Else
        FileDialog.sInitDir = App.Path & "\"
    End If
    
    sOpen = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        sDir = sOpen.sLastDirectory
        SaveSetting App.EXEName, "Dir", "Init", sDir
        If Right(sDir, 1) <> "\" Then
            sDir = sDir & "\"
        End If
        
        If Not ARv.ReportSource Is Nothing Then
            Set ARv.ReportSource = Nothing
        End If
        ARv.Pages.Load sDir & sOpen.sFiles(1)
        'BGS 8.11.2000 Need to set the Caption to the
        'File name and posn form appropriately
        With mARViewer
            Me.Caption = sDir & sOpen.sFiles(1)
            goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, , , , False
        End With
     End If
   Exit Sub
EH:
    If Err.Number > 0 Then
        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFOpen_Click"
    End If
    Err.Clear
End Sub

Private Sub miFPDFExportJPG_Click()
    On Error GoTo EH
    Dim sRet As String
    Dim lJPGqual As Long
    'Get new setting from user
    lJPGqual = GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "PDF_JPG", 100)
    sRet = InputBox("Please enter 1 - 100", "PDF JPG Export Quality Percentage", lJPGqual)
    'Check to be sure it is valid
    sRet = Replace(sRet, "%", vbNullString)
    If IsNumeric(sRet) Then
        lJPGqual = sRet
        If lJPGqual >= 1 And lJPGqual <= 100 Then
            SaveSetting App.EXEName, "EXPORT_PHOTO_QUALITY", "PDF_JPG", lJPGqual
        End If
    End If
    miFPDFExportJPG.Caption = "JPG Export Quality%(" & GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "PDF_JPG", 100) & ")"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFPDFExportJPG_Click"
End Sub

Private Sub miFPDFExportNow_Click()
    'Users can download Reader for free here...
    'ftp://ftp.adobe.com/pub/adobe/acrobatreader/win/4.x/ar405eng.exe
    Set mPDF = New ActiveReportsPDFExport.ARExportPDF
    mPDF.JPGQuality = GetSetting(App.EXEName, "EXPORT_PHOTO_QUALITY", "PDF_JPG", 100)
    ExportFile mPDF, "Portable Document Format", "pdf"
End Sub

Private Sub miFSave_Click()
    Dim sSave As SelectedFile
    Dim sDir As String
    Dim sMyFilter As String
    Dim sSaveDir As String
    Dim sFilePath As String
    Dim bUseFilePath As Boolean
    
    On Error GoTo EH
    
    sMyFilter = sMyFilter & "Report Document File (*.rdf)" & SD & "*.rdf" & SD
    FileDialog.sFilter = sMyFilter
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT
    FileDialog.sDlgTitle = "Save Active Report File"
    sSaveDir = GetSetting(App.EXEName, "Dir", "Save", "Error")
    If sSaveDir <> "Error" Then
        FileDialog.sInitDir = sSaveDir
    Else
        FileDialog.sInitDir = "C:\"
    End If
    
    sSave = ShowSave(Me.hWnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
        sDir = sSave.sLastDirectory
        If goUtil.utFileExists(sDir, True) Then
            SaveSetting App.EXEName, "Dir", "Save", sDir
            If Right(sDir, 1) <> "\" Then
                sDir = sDir & "\"
            End If
        Else
            If InStr(1, sSave.sFiles(1), "\", vbTextCompare) > 0 Then
                bUseFilePath = True
                sFilePath = left(sSave.sFiles(1), InStrRev(sSave.sFiles(1), "\"))
                SaveSetting App.EXEName, "Dir", "Save", sFilePath
            End If
        End If
    
        If sSave.nFilesSelected > 0 Then
            'BGS make sure the extention is there
            If Not UCase(Right(sSave.sFiles(1), 4)) = ".RDF" Then
                sSave.sFiles(1) = sSave.sFiles(1) & ".rdf"
            End If
            If Not ARv.ReportSource Is Nothing Then
                If Not bUseFilePath Then
                    ARv.ReportSource.Pages.Save sDir & sSave.sFiles(1)
                Else
                    ARv.ReportSource.Pages.Save sSave.sFiles(1)
                End If
            Else
                If Not bUseFilePath Then
                    ARv.Pages.Save sDir & sSave.sFiles(1)
                Else
                    ARv.Pages.Save sSave.sFiles(1)
                End If
            End If
        End If
    End If
    
   Exit Sub
EH:
    If Err.Number > 0 Then
        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFSave_Click"
    End If
    Err.Clear
End Sub

Private Sub miFXLSExport_Click()
    On Error GoTo EH
    Set mXLS = New ActiveReportsExcelExport.ARExportExcel
    ExportFile mXLS, "Excel Spreadsheet", "xls"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFXLSExport_Click"
End Sub



Private Sub miFRTFExport_Click()
    On Error GoTo EH
    Set mRTF = New ActiveReportsRTFExport.ARExportRTF
    ExportFile mRTF, "Rich Text Format", "rtf"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFRTFExport_Click"
End Sub

Private Sub miFTXTExport_Click()
    On Error GoTo EH
    Set mTXT = New ActiveReportsTextExport.ARExportText
    ExportFile mTXT, "Text Format", "txt"
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFTXTExport_Click"
End Sub

Private Sub miFPrint_Click()
    On Error GoTo EH
    ARv.UseSourcePrinter = True
    ARv.PrintReport False
EH:
    If Err.Number > 0 Then
        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miFPrint_Click"
    End If
    Err.Clear
End Sub

Private Sub miPrinterSetup_Click()
    On Error GoTo EH
    ARv.UseSourcePrinter = True
    ARv.Printer.SetupDialog
EH:
    If Err.Number > 0 Then
        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub miPrinterSetup_Click"
    End If
    Err.Clear
End Sub

Private Sub miFExit_Click()
    Unload Me
End Sub

Private Function ExportFile(pobjExport As Object, psFileFormat As String, _
                            psFileExt As String) As Boolean

    Dim sExport As SelectedFile
    Dim sDir As String
    Dim sMyFilter As String
    Dim sExportDir As String
    Dim sFilePath As String
    Dim bUseFilePath As Boolean
    On Error GoTo EH
    
    sMyFilter = sMyFilter & psFileFormat & " (*." & psFileExt & ")" & SD & "*." & psFileExt & SD
    FileDialog.sFilter = sMyFilter
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT
    FileDialog.sDlgTitle = "Export " & psFileFormat
    sExportDir = GetSetting(App.EXEName, "Dir", "Export" & psFileExt, "Error")
    If sExportDir <> "Error" Then
        FileDialog.sInitDir = sExportDir
    Else
        FileDialog.sInitDir = "C:\"
    End If
    
    sExport = ShowSave(Me.hWnd)
    If Err.Number <> 32755 And sExport.bCanceled = False Then
        sDir = sExport.sLastDirectory
        If goUtil.utFileExists(sDir, True) Then
            SaveSetting App.EXEName, "Dir", "Export" & psFileExt, sDir
            If Right(sDir, 1) <> "\" Then
                sDir = sDir & "\"
            End If
        Else
            If InStr(1, sExport.sFiles(1), "\", vbTextCompare) > 0 Then
                bUseFilePath = True
                sFilePath = left(sExport.sFiles(1), InStrRev(sExport.sFiles(1), "\"))
                SaveSetting App.EXEName, "Dir", "Save", sFilePath
            End If
        End If
    
        If sExport.nFilesSelected > 0 Then
            'BGS make sure the extention is there
            If Not UCase(Right(sExport.sFiles(1), 4)) = "." & UCase(psFileExt) Then
                If Not UCase(Right(sExport.sFiles(1), 5)) = "." & UCase(psFileExt) Then
                    sExport.sFiles(1) = sExport.sFiles(1) & "." & psFileExt
                End If
            End If
            
            If Not bUseFilePath Then
                pobjExport.FileName = sDir & sExport.sFiles(1)
            Else
                pobjExport.FileName = sExport.sFiles(1)
            End If
            'BGS 9.4.2002 All we have to do is Check the ARV object page count.
            'Checking the report source page count may or may not work.
            'Report source may have been set to nothing else where. ARV maintains
            'a copy of the report in ARV object.
            If ARv.Pages.Count > 0 Then
                pobjExport.Export ARv.Pages
            End If
        End If
    End If
    
    'BGS Cleanup
    Set pobjExport = Nothing
   Exit Function
EH:
    If Err.Number > 0 Then
        goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function ExportFile"
    End If
    Err.Clear
End Function

