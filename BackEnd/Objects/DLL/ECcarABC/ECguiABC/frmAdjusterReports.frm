VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdjusterReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjuster Reports"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdjusterReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framAdjSavedPackages 
      Caption         =   "Saved Adjuster Reports (Available to Adjuster only)"
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   8880
      Begin VB.CommandButton cmdDelItem 
         Caption         =   "&Delete File"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewFile 
         Caption         =   "View &File"
         Height          =   375
         Left            =   7410
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdMailTo 
         Caption         =   "Mail To:"
         Height          =   375
         Left            =   3765
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwSavedPackages 
         Height          =   2010
         Left            =   120
         TabIndex        =   23
         Tag             =   "Enable"
         Top             =   675
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   3545
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgVarDoc"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   855
      Left            =   8025
      MaskColor       =   &H00000000&
      Picture         =   "frmAdjusterReports.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Exit"
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame framAdjReports 
      Caption         =   "Adjuster Reports"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CheckBox chkXMLOnly 
         Alignment       =   1  'Right Justify
         Caption         =   "Export XML Only"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         ToolTipText     =   "Loss Report and Attachments will not have XML export files!"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox chkXMLExport 
         Alignment       =   1  'Right Justify
         Caption         =   "Include XML Export"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         ToolTipText     =   "Loss Report and Attachments will not have XML export files!"
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkUsePass 
         Alignment       =   1  'Right Justify
         Caption         =   "Use password"
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdSavePackage 
         Caption         =   "P&rint to File"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         Picture         =   "frmAdjusterReports.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2685
         Width           =   1830
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   7920
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   2355
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   6960
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   2355
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkPrintPreview 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Preview"
         Height          =   240
         Left            =   6720
         TabIndex        =   9
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkHideDetails 
         Alignment       =   1  'Right Justify
         Caption         =   "Hide details"
         Height          =   240
         Left            =   6720
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintReport 
         Caption         =   "&Print"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   1605
         Width           =   2055
      End
      Begin VB.TextBox txtCommissionPct 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "Numeric ALPHANUM"
         Text            =   "65"
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtSunDaysDate 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "Date"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdSundaysDate 
         Height          =   375
         Left            =   1905
         Picture         =   "frmAdjusterReports.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Date"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboAdjusterReports 
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   8655
      End
      Begin VB.Label lblPass 
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7920
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblPass 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6960
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblpctCommission 
         Caption         =   "Commission Percentage (Enter % example 65% as 65)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label lblSunDayDate 
         Caption         =   "Enter Sunday's Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblAdjusterReports 
         Caption         =   "Adjuster Reports:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmAdjusterReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private mbFormLoading As Boolean
Private msReportName As String
Private msrptProjectName As String
Private msrptClassName As String
Private mlrptVersion As Long
Private mbUnloadMe As Boolean
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mArv As V2ARViewer.clsARViewer
Private moForm As Form
Private mbXMLExport As Boolean
Private mbXMLOnly As Boolean
Private msCatName As String

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Set MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Get MyGUI() As V2ECKeyBoard.clsCarGUI
    Set MyGUI = moGUI
End Property

Private Sub cboAdjusterReports_Click()
    On Error GoTo EH
    Dim sData As String
    Dim saryData() As String
    
    sData = cboAdjusterReports.Text
    
    If sData <> vbNullString Then
        msReportName = Trim(left(sData, 200))
        sData = Mid(sData, 200)
        sData = Trim(sData)
        saryData() = Split(sData, "|", , vbBinaryCompare)
        msrptProjectName = saryData(0)
        msrptClassName = saryData(1)
        mlrptVersion = saryData(2)
    Else
        Exit Sub
    End If
    
    If StrComp(msrptProjectName, "ECrpt_arCar", vbTextCompare) = 0 Then
        chkHideDetails.Visible = True
        lblpctCommission.Visible = False
        txtCommissionPct.Visible = False
        lblSunDayDate.Visible = False
        txtSunDaysDate.Visible = False
        cmdSundaysDate.Visible = False
    ElseIf StrComp(msrptProjectName, "ECrpt_arCommission", vbTextCompare) = 0 Then
        chkHideDetails.Visible = True
        lblpctCommission.Visible = True
        txtCommissionPct.Visible = True
        lblSunDayDate.Visible = False
        txtSunDaysDate.Visible = False
        cmdSundaysDate.Visible = False
    ElseIf StrComp(msrptProjectName, "ECrpt_arProduction", vbTextCompare) = 0 Then
        chkHideDetails.Visible = True
        lblpctCommission.Visible = False
        txtCommissionPct.Visible = False
        lblSunDayDate.Visible = False
        txtSunDaysDate.Visible = False
        cmdSundaysDate.Visible = False
    ElseIf StrComp(msrptProjectName, "ECrpt_arWCTL", vbTextCompare) = 0 Then
        chkHideDetails.Visible = True
        lblpctCommission.Visible = False
        txtCommissionPct.Visible = False
        lblSunDayDate.Visible = True
        txtSunDaysDate.Visible = True
        cmdSundaysDate.Visible = True
    End If
    
    cmdPrintReport.Enabled = True
    cmdSavePackage.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboAdjusterReports_Click"
End Sub

Private Sub chkUsePass_Click()
    On Error GoTo EH
    
    If chkUsePass.Value = vbChecked Then
        lblPass(0).Visible = True
        txtPassWord(0).Visible = True
        lblPass(1).Visible = True
        txtPassWord(1).Visible = True
    Else
        lblPass(0).Visible = False
        txtPassWord(0).Visible = False
        lblPass(1).Visible = False
        txtPassWord(1).Visible = False
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUsePass_Click"
End Sub

Private Sub chkXMLExport_Click()
    On Error GoTo EH
    If mbFormLoading Then
        Exit Sub
    End If
    
    If chkXMLExport.Value = vbChecked Then
        mbXMLExport = True
        If chkXMLExport.Enabled Then
            chkXMLOnly.Enabled = True
        End If
    Else
        mbXMLExport = False
        chkXMLOnly.Enabled = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "XML_EXPORT", mbXMLExport
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkXMLExport_Click"
End Sub

Private Sub chkXMLOnly_Click()
    On Error GoTo EH
    If mbFormLoading Then
        Exit Sub
    End If
    
    If chkXMLOnly.Value = vbChecked Then
        mbXMLOnly = True
    Else
        mbXMLOnly = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "XML_ONLY", mbXMLOnly
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkXMLOnly_Click"
End Sub

Private Sub cmdDelItem_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim sFilePath As String
    Dim sFileName As String
    
    If lvwSavedPackages.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    cmdDelItem.Enabled = False
    
    sFileName = lvwSavedPackages.SelectedItem.Text
    sFilePath = goUtil.AttachReposPath & sFileName
    
    sMess = "Are you sure you want to Delete the Selected File " & sFileName & " ? "
    
    If MsgBox(sMess, vbQuestion + vbYesNo, "Delete Selected File") = vbYes Then
        sMess = goUtil.utDeleteFile(sFilePath)
        If sMess <> vbNullString Then
            sMess = "Error " & sMess
            MsgBox sMess, vbCritical + vbOKOnly, "Error"
        End If
    End If
    
    PopulatelvwSavedPackages
    
    cmdDelItem.Enabled = True
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelItem_Click"
End Sub

Public Sub PopulatelvwSavedPackages()
    On Error GoTo EH
    'Source Variables
    Dim varySavedPackages As Variant
    Dim sSavedPackage As String
    Dim sSavedPackagePath As String
    Dim iCount As Integer
    Dim itmX As ListItem
    Dim sDateLastUpdated As String
    Dim sDateCreated As String
    Dim oFI As V2ECKeyBoard.clsFileVersion
    Dim myFI As V2ECKeyBoard.FILE_INFORMATION

    lvwSavedPackages.ListItems.Clear
    'BGS 1.2.2002 load the Avail reports
    If Not GetSavedPackages(varySavedPackages) Then
        Exit Sub
    Else
        If IsArray(varySavedPackages) Then
            'Set the File info Object
            Set oFI = New V2ECKeyBoard.clsFileVersion
            
            For iCount = LBound(varySavedPackages) To UBound(varySavedPackages)
                sSavedPackage = varySavedPackages(iCount)
                
                sSavedPackagePath = goUtil.AttachReposPath & sSavedPackage
                
                myFI = oFI.GetFileInformation(sSavedPackagePath)
                sDateLastUpdated = myFI.dtLastModifyTime
                sDateCreated = myFI.dtCreationDate
                
                'If the DateLastUpdated is Earlier than the Date Created...
                'That means the User has not Done anything with it
                If IsDate(sDateLastUpdated) And IsDate(sDateCreated) Then
                    If CDate(sDateLastUpdated) < CDate(sDateCreated) Then
                        'So Make it Blank
                        sDateLastUpdated = vbNullString
                    End If
                End If
                
                Set itmX = lvwSavedPackages.ListItems.Add(, , sSavedPackage)
                
                itmX.SubItems(SavedPackagesListView.DateLastUpdated - 1) = Format(sDateLastUpdated, "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(SavedPackagesListView.DateLastUpdatedSort - 1) = Format(sDateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                itmX.SubItems(SavedPackagesListView.DateCreated - 1) = Format(sDateCreated, "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(SavedPackagesListView.DateCreatedSort - 1) = Format(sDateCreated, "YYYY/MM/DD HH:MM:SS")
            Next
        End If
    End If

    'cleanup
    Set itmX = Nothing
    Set oFI = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwSavedPackages"
End Sub

Public Function GetSavedPackages(pvarySavedPackages As Variant) As Boolean
    On Error GoTo EH
    Dim sReport As String
    Dim saryReports() As String
    Dim iReportCount As Integer
    Dim bFound As Boolean
    
    'BGS get the .zip reports
    sReport = Dir(goUtil.AttachReposPath & "\" & msCatName & "*.zip")
    Do Until sReport = vbNullString
        bFound = True
        iReportCount = iReportCount + 1
        ReDim Preserve saryReports(1 To iReportCount)
        saryReports(iReportCount) = sReport
        sReport = Dir
    Loop
    If bFound Then
        pvarySavedPackages = saryReports
        GetSavedPackages = True
    Else
        GetSavedPackages = False
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetSavedPackages"
End Function

Private Sub cmdExit_Click()
    On Error GoTo EH
    Dim sMess As String
    mbUnloadMe = True
    cmdExit.Enabled = False
    CLEANUP
    Unload Me
    Exit Sub
    
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdMailTo_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    Dim sFile As String
    Dim itmX As ListItem
    Dim lRet As Long
    Dim sFileName As String
    Dim sMapiLaunchPath As String
    
    If lvwSavedPackages.SelectedItem Is Nothing Then
        Exit Sub
    End If

    Set itmX = lvwSavedPackages.SelectedItem
    
    sFileName = itmX.Text

    sPath = goUtil.AttachReposPath

    sPath = goUtil.AttachReposPath & sFileName
   
    sMapiLaunchPath = goUtil.gsInstallDir & "\" & "SendMail.exe"
    
    
    If goUtil.utFileExists(sPath) Then
        sMapiLaunchPath = """" & sMapiLaunchPath & """  """ & sPath & """"
        Shell sMapiLaunchPath, vbMaximizedFocus
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdMailTo_Click"
End Sub


Private Sub cmdPrintReport_Click()
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    
    'First be sure control are valid
    
    goUtil.utValidate Me
    txtCommissionPct.Text = Format(txtCommissionPct.Text, "00")
    cmdPrintReport.Enabled = False
    cmdSavePackage.Enabled = False
    If PrintAdjusterReport() Then
        If Not mbUnloadMe Then
            cmdPrintReport.Enabled = True
            cmdSavePackage.Enabled = True
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintReport_Click"
End Sub

Private Function PrintAdjusterReport() As Boolean
    On Error GoTo EH
    Dim sParams As String
    Dim MyAdjReport As Object
    Dim sPercent As String
    Dim sDate As String
    Dim sReportName As String
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'If using Adobe PDF Viewer
    Dim sPDFFilePath As String
    Dim sPrintPreview As String
    Dim bUseAdobeReader As Boolean
    sPrintPreview = GetSetting(goUtil.gsMainAppEXEName, "GENERAL", "PRINT_PREVIEW", "USE_ADOBE")
    
    Select Case UCase(sPrintPreview)
        Case "USE_ADOBE"
            bUseAdobeReader = True
    End Select
    
    sParams = sParams & "pCATID=" & goUtil.gsCurCat & "|"
    sParams = sParams & "pClientCompanyID=" & goUtil.gsCurCar & "|"
    sParams = sParams & "pUSERSID=" & goUtil.gsCurUsersID & "|"
    If chkHideDetails.Value = vbChecked Then
        sParams = sParams & "pHideDetails=" & "True" & "|"
    Else
        sParams = sParams & "pHideDetails=" & "False" & "|"
    End If
    If chkPrintPreview.Value = vbChecked Then
        sParams = sParams & "pbPreview=" & "True" & "|"
    Else
        sParams = sParams & "pbPreview=" & "False" & "|"
    End If
    
    If bUseAdobeReader Then
        sPDFFilePath = goUtil.gsInstallDir & "\TempActiveReport" & goUtil.utGetTickCount & ".pdf"
        sParams = sParams & "psXportPath=" & sPDFFilePath & "|"
        sParams = sParams & "pPDFJPEGQuality=" & "50" & "|"
        sParams = sParams & "pXportType=" & ExportType.ARPdf & "|"
    Else
        sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
    End If
    
    
    
    If IsNumeric(txtCommissionPct.Text) Then
        sPercent = "." & txtCommissionPct.Text
    Else
        sPercent = ".65"
    End If
    sParams = sParams & "pCommissionPercentage=" & sPercent & "|"
    If IsDate(txtSunDaysDate.Text) Then
        sDate = txtSunDaysDate.Text
    Else
        sDate = Format(Now(), "MM/DD/YYYY")
    End If
    sParams = sParams & "pFromSundayDate=" & sDate & "|"

    sReportName = msrptProjectName & "." & msrptClassName
    
   Set oCarList = CreateObject(goUtil.goCurCarList.ClassName)
    
    If bUseAdobeReader Then
        oCarList.GetARReport sReportName, mlrptVersion, sParams
        If goUtil.utFileExists(sPDFFilePath) Then
            If chkPrintPreview.Value = vbChecked Then
                goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, Trim(left(cboAdjusterReports.Text, 20))
            Else
                goUtil.utShellExecute , "PRINT", sPDFFilePath, , , vbNormalFocus, True, True, True, Trim(left(cboAdjusterReports.Text, 20))
            End If
            DoEvents
            Sleep 1000
            goUtil.utDeleteFile sPDFFilePath
            oCarList.CLEANUP
            Set oCarList = Nothing
        End If
    Else

        Set MyAdjReport = oCarList.GetARReport(sReportName, mlrptVersion, sParams)
    
        If mArv Is Nothing Then
            Set mArv = New V2ARViewer.clsARViewer
            mArv.SetUtilObject goUtil
        End If
        
        If Not moForm Is Nothing Then
            Unload moForm
            Set moForm = Nothing
        End If
    
        With mArv
            'Pass in true to have Active reports process on separate thread.
            'This will allow the viewer to load while the report is processing
            'false will force the report to run on single thread
            If chkPrintPreview.Value = vbUnchecked Then
                MyAdjReport.PrintReport False
                Unload MyAdjReport
                Set MyAdjReport = Nothing
            Else
                MyAdjReport.Run False 'True
                .objARvReport = MyAdjReport
                .sRptTitle = msReportName
                .HidePrintButton = False
                .ShowReportOnForm moForm, vbModeless
                Unload .objARvReport
                Set .objARvReport = Nothing
                Unload MyAdjReport
                Set MyAdjReport = Nothing
                oCarList.CLEANUP
                Set oCarList = Nothing
            End If
            
        End With
    End If
    
    PrintAdjusterReport = True
    
    Exit Function
EH:
    
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function PrintAdjusterReport"
End Function

Private Sub cmdSavePackage_Click()
    On Error GoTo EH
    Dim sMess As String
    Dim bUsePassword As Boolean
    Dim sBuildDir As String
    Dim itmX As ListItem
    Dim sFlagText As String
    
    cmdPrintReport.Enabled = False
    cmdSavePackage.Enabled = False
    'If password selected, must validate it
    If chkUsePass.Value = vbChecked Then
        If StrComp(txtPassWord(0).Text, txtPassWord(1).Text, vbBinaryCompare) <> 0 Then
            sMess = "Password not confirmed!"
        ElseIf txtPassWord(0).Text = vbNullString Then
            sMess = "Password can not be blank!"
        End If
        If sMess <> vbNullString Then
            MsgBox sMess, vbExclamation + vbOKOnly, "Invalid Password Entry"
            cmdPrintReport.Enabled = True
            cmdSavePackage.Enabled = True
            Exit Sub
        End If
        bUsePassword = True
    End If

    'Need to create Build Folder to store all the print to file .pdf docs
    'this folder will be used to create the Zip File and then move to
    'the Attach repos folder.
    sBuildDir = goUtil.gsInstallDir & "\BuildSave\"
    If Not goUtil.utFileExists(sBuildDir, True) Then
        goUtil.utMakeDir sBuildDir
    Else
        'Need to be sure nothing exisits in the build dir
        goUtil.utDeleteFile sBuildDir & "*.*"
        Sleep 100
    End If

    'Save this Item to File
    If SaveThisPackageToFile(bUsePassword, sBuildDir) Then
        PopulatelvwSavedPackages
        chkUsePass.Value = vbUnchecked
        MsgBox "Save Succeeded!", vbInformation + vbOKOnly, "Success!"
    End If
    cmdPrintReport.Enabled = True
    cmdSavePackage.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSavePackage_Click"
End Sub

Public Function SaveThisPackageToFile(pbUsePassword As Boolean, psBuildDir As String) As Boolean
    On Error GoTo EH
'    Dim RS As ADODB.Recordset
    Dim oXZip As V2ECKeyBoard.clsXZip
    Dim sSuffixName As String
    Dim sMess As String
    Dim sZipName As String
    Dim sPassWord As String
    Dim sEncryptPassWord As String
    Dim sDestDir As String
    
    'Get Filename from user
    sMess = "Please enter a file name for this item."
    
    sSuffixName = InputBox(sMess, "Enter File Name", Trim(left(cboAdjusterReports.Text, 30)))
    
    sMess = vbNullString
    If Len(sSuffixName) > 30 Then
        sMess = "Name too Big!"
        sSuffixName = vbNullString
    End If
    goUtil.utCleanFileFolderName sSuffixName, False
    
    If Trim(sSuffixName) = vbNullString Then
        MsgBox "File not Saved!" & vbCrLf & vbCrLf & sMess, vbExclamation + vbOKOnly, "Save Aborted"
        GoTo CLEAN_UP
    End If
    
    'Create Zipname
    sZipName = msCatName & "_" & sSuffixName & ".zip"
    
    'The ultimate destination for this file will be
    'the attach repos
    sDestDir = goUtil.AttachReposPath
    
    'Check to see if it already Exists
    If goUtil.utFileExists(sDestDir & sZipName, False) Then
        sMess = "The file """ & sZipName & """ already exists!" & vbCrLf & vbCrLf
        sMess = sMess & "Click ""Yes"" to update this existing file." & vbCrLf
        sMess = sMess & "(Replace exisiting documents and Append new documents)" & vbCrLf & vbCrLf
        sMess = sMess & "Click ""No"" to Abort!"
        If MsgBox(sMess, vbExclamation + vbYesNo, "File Already Exists") = vbNo Then
            GoTo CLEAN_UP
        End If
    End If
    
    If Not PrintPackageItemsToFile(psBuildDir) Then
        sMess = "Package items NOT saved!" & vbCrLf & vbCrLf
        MsgBox sMess, vbExclamation + vbOKOnly, "Problems saving to file"
        GoTo CLEAN_UP
    End If
    'Need to save these created files into 1 zip file
    
    'Create the Zip utility
    Set oXZip = New V2ECKeyBoard.clsXZip
    oXZip.SetUtilObject goUtil
    
    If pbUsePassword Then
        sPassWord = txtPassWord(0).Text
        sEncryptPassWord = goUtil.Encode(sPassWord)
    End If
    If Not oXZip.SaveZIPFiles(psBuildDir, sZipName, "*.*", sEncryptPassWord, sDestDir) Then
        GoTo CLEAN_UP
    End If
    
    SaveThisPackageToFile = True
    
CLEAN_UP:
    Screen.MousePointer = vbDefault
    'cleanup
    Set oXZip = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveThisPackageToFile"
End Function

Private Function PrintPackageItemsToFile(psBuilDir As String) As Boolean
    On Error GoTo EH
    Dim sSaveToFileName As String
    Dim bPrintPackageitems As Boolean
    
    If cboAdjusterReports.ListIndex = -1 Then
        Exit Function
    End If
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    'Max Len of File name is 40 Chars
    sSaveToFileName = Trim(left(cboAdjusterReports.Text, 40))
    goUtil.utCleanFileFolderName sSaveToFileName, False
    sSaveToFileName = sSaveToFileName & ".pdf"
    
    bPrintPackageitems = PrintActiveReportToFile(cboAdjusterReports, psBuilDir, sSaveToFileName, mbXMLExport, mbXMLOnly)
        
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    PrintPackageItemsToFile = bPrintPackageitems
    
    Exit Function
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function PrintPackageItemsToFile"
End Function

Public Function PrintActiveReportToFile(poReportItem As Object, _
                                psSaveToFilePath As String, _
                                psSaveToFileName As String, _
                                Optional pbExportXML As Boolean, _
                                Optional pbExportXMLOnly As Boolean) As Boolean
    On Error GoTo EH
    Dim sParams As String
    Dim sReportName As String
    Dim sReportTitle As String
    Dim srptProjectName As String
    Dim srptClassName As String
    Dim lrptVersion As Long
    Dim sData As String
    Dim saryData() As String
    Dim ocboReport As ComboBox
    Dim itmXReport As ListItem
    Dim MyActReport As ActiveReport
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'If using Adobe PDF Viewer
    Dim sPDFFilePath As String
    'Export to XML FileName
    Dim sXMLFilePath As String
    Dim sXMLFileName As String
    'Extra params
    Dim sPercent As String
    Dim sDate As String
    
    If TypeOf poReportItem Is ComboBox Then
        Set ocboReport = poReportItem
        sData = ocboReport.Text
    ElseIf TypeOf poReportItem Is ListItem Then
        Set itmXReport = poReportItem
        sData = itmXReport.ListSubItems(GuiPackageItemListView.ReportFormat - 1)
    Else
        Exit Function
    End If
    
    If sData = vbNullString Then
        Exit Function
    End If
    
    sReportTitle = Trim(left(sData, 200))
    goUtil.utCleanFileFolderName sReportTitle, False
    sData = Mid(sData, InStr(1, sData, String(200, " "), vbBinaryCompare))
    sData = Trim(sData)
    saryData() = Split(sData, "|", , vbBinaryCompare)
    
    srptProjectName = saryData(0)
    srptClassName = saryData(1)
    lrptVersion = saryData(2)
    
    
    'Build Params List to be passed in to Create Report Object
    'This Object will have list of Report Parameters it requires
    sParams = vbNullString
    sParams = sParams & "pCATID=" & goUtil.gsCurCat & "|"
    sParams = sParams & "pClientCompanyID=" & goUtil.gsCurCar & "|"
    sParams = sParams & "pUSERSID=" & goUtil.gsCurUsersID & "|"
    If chkHideDetails.Value = vbChecked Then
        sParams = sParams & "pHideDetails=" & "True" & "|"
    Else
        sParams = sParams & "pHideDetails=" & "False" & "|"
    End If
    
    If IsNumeric(txtCommissionPct.Text) Then
        sPercent = "." & txtCommissionPct.Text
    Else
        sPercent = ".65"
    End If
    sParams = sParams & "pCommissionPercentage=" & sPercent & "|"
    If IsDate(txtSunDaysDate.Text) Then
        sDate = txtSunDaysDate.Text
    Else
        sDate = Format(Now(), "MM/DD/YYYY")
    End If
    sParams = sParams & "pFromSundayDate=" & sDate & "|"

    'If using Adobe PDF Viewer
    sPDFFilePath = goUtil.gsInstallDir & "\TempActiveReport" & goUtil.utGetTickCount & ".pdf"
    sParams = sParams & "psXportPath=" & sPDFFilePath & "|"
    sParams = sParams & "pPDFJPEGQuality=" & "50" & "|"
    sParams = sParams & "pXportType=" & ExportType.ARPdf & "|"
    
    'Set the report name
    sReportName = srptProjectName & "." & srptClassName

    Set oCarList = CreateObject(goUtil.goCurCarList.ClassName)

    'Add Export XML Parameters here
    If pbExportXML Then
        sParams = sParams & "pbExportXML=True|"
        If pbExportXMLOnly Then
            sParams = sParams & "pbExportXMLOnly=True|"
        End If
    End If
    
    Set MyActReport = oCarList.GetARReport(sReportName, lrptVersion, sParams)
    
    If goUtil.utFileExists(sPDFFilePath) Or (pbExportXML And pbExportXMLOnly) Then
        If pbExportXML Then
            If Not pbExportXMLOnly Then
                goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
                goUtil.utDeleteFile sPDFFilePath
            End If
        Else
            goUtil.utCopyFile sPDFFilePath, psSaveToFilePath & psSaveToFileName
            goUtil.utDeleteFile sPDFFilePath
        End If
        
        If pbExportXML Then
            'change the pdffile path the XML
            sXMLFilePath = sPDFFilePath
            sXMLFilePath = left(sXMLFilePath, InStrRev(sXMLFilePath, ".", , vbBinaryCompare))
            sXMLFilePath = sXMLFilePath & "xml"
            'Change the pdf to XML file path
            sXMLFileName = psSaveToFileName
            sXMLFileName = left(sXMLFileName, InStrRev(sXMLFileName, ".", , vbBinaryCompare))
            sXMLFileName = sXMLFileName & "xml"
            goUtil.utCopyFile sXMLFilePath, psSaveToFilePath & sXMLFileName
            goUtil.utDeleteFile sXMLFilePath
        End If
   End If
    
CLEAN_UP:
    'Cleanup
    Set ocboReport = Nothing
    Set itmXReport = Nothing
    Set MyActReport = Nothing
    
    PrintActiveReportToFile = True
    
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintActiveReportToFile"
End Function

Private Sub cmdSundaysDate_Click()
    On Error GoTo EH
    
    MyGUI.ShowCalendar txtSunDaysDate
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSundaysDate_Click"
End Sub

Private Sub cmdViewFile_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    Dim sFile As String
    Dim itmX As ListItem
    Dim lRet As Long
    Dim sFileName As String
    
    If lvwSavedPackages.SelectedItem Is Nothing Then
        Exit Sub
    End If

    Set itmX = lvwSavedPackages.SelectedItem
    
    sFileName = itmX.Text

    sPath = goUtil.AttachReposPath

    sMyFilter = sMyFilter & "ZIP File" & " (" & sFileName & ")" & SD & sFileName & SD

    sPath = goUtil.utGetPath(App.EXEName, "TempZipDir", "You Can Drag and Drop " & sFileName & " to your email program.", "You Can Drag and Drop " & sFileName & " to your email program.", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If sSelFile <> vbNullString Then
        sPath = goUtil.AttachReposPath & sFileName
    End If
    
    If goUtil.utFileExists(sPath) Then
        lRet = goUtil.utShellExecute(GetDesktopWindow, "OPEN", sPath, vbNullString, App.Path, vbNormalFocus, False, False, True)
    End If
    
    PopulatelvwSavedPackages

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdViewFile_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    mbFormLoading = True
    
    mbXMLExport = CBool(GetSetting(App.EXEName, "GENERAL", "XML_EXPORT", False))
    If mbXMLExport Then
        chkXMLExport.Value = vbChecked
    Else
        chkXMLExport.Value = vbUnchecked
    End If
    
    mbXMLOnly = CBool(GetSetting(App.EXEName, "GENERAL", "XML_ONLY", False))
    If mbXMLOnly Then
        chkXMLOnly.Value = vbChecked
    Else
        chkXMLOnly.Value = vbUnchecked
    End If
    'only enable this if xml is checked
    chkXMLOnly.Enabled = mbXMLExport

    
    'use this to help build the Saved file name
    msCatName = goUtil.gsCurCarDBName & "_" & GetCurCatName
    
    LoadHeaderlvwSavedPackages
    LoadReports
    PopulatelvwSavedPackages
    
    mbFormLoading = False
    mbUnloadMe = False
    Exit Sub
EH:
    mbFormLoading = False
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function GetCurCatName() As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    
    'need to get Correct Assignment type
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    sSQL = "SELECT  [Name] "
    sSQL = sSQL & "FROM     CAT "
    sSQL = sSQL & "WHERE    [CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      [CompanyID] = " & goUtil.gsCurCompany & " "
    
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    RS.MoveFirst
    
    GetCurCatName = goUtil.IsNullIsVbNullString(RS.Fields("Name"))
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurCatName"
End Function


Public Function LoadReports() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim lNewIndex As Long
    Dim lMaxSPVersion As Long
    Dim sData As String
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT MAX(SPVERSION) As SPVERSION "
    sSQL = sSQL & "From SOFTWAREPACKAGE "
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.EOF Then
        Exit Function
    End If
    
    RS.MoveFirst
    lMaxSPVersion = RS.Fields("SPVERSION").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Application A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMaxSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMaxSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Reports' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Adjuster' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    
    cboAdjusterReports.Clear
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            sData = vbNullString
            sData = RS.Fields("Description").Value
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
             cboAdjusterReports.AddItem sData
            lNewIndex = cboAdjusterReports.NewIndex
            cboAdjusterReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            RS.MoveNext
        Loop
    End If
    
    'Need to Add Items from History Table
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  ApplicationHistory A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMaxSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMaxSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Reports' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Adjuster' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "
    
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            sData = vbNullString
            sData = RS.Fields("Description").Value & " - (Previous Version) "
            sData = sData & String(200, " ")
            sData = sData & RS.Fields("ProjectName").Value & "|" & RS.Fields("ClassName").Value & "|" & RS.Fields("Version").Value
             cboAdjusterReports.AddItem sData
            lNewIndex = cboAdjusterReports.NewIndex
            cboAdjusterReports.ItemData(lNewIndex) = RS.Fields("ApplicationID").Value
            RS.MoveNext
        Loop
    End If
    
    txtSunDaysDate.Text = Format(Now(), "MM/DD/YYYY")
    
    
    'cleanup
    
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadReports"
End Function
    

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    Select Case UnloadMode
        Case vbFormControlMenu
            CLEANUP
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True, , , , True
    If Not goUtil.gfrmECTray Is Nothing Then
        goUtil.gfrmECTray.ShowMe False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub lvwSavedPackages_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lvwSavedPackages.SortOrder = lvwAscending Then
        lvwSavedPackages.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lvwSavedPackages.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lvwSavedPackages.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case SavedPackagesListView.DateCreated, SavedPackagesListView.DateLastUpdated
            lvwSavedPackages.SortKey = ColumnHeader.Index
        Case Else
            lvwSavedPackages.SortKey = ColumnHeader.Index - 1
    End Select
    
    lvwSavedPackages.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lvwAvail_ColumnClick"
End Sub


Private Sub lvwSavedPackages_DblClick()
    cmdViewFile_Click
End Sub

Private Sub lvwSavedPackages_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn, KeyCodeConstants.vbKeySpace
            cmdViewFile_Click
        Case KeyCodeConstants.vbKeyDelete
            cmdDelItem_Click
    End Select
End Sub

Private Sub txtCommissionPct_GotFocus()
    goUtil.utSelText txtCommissionPct
End Sub

Private Sub txtCommissionPct_LostFocus()
    On Error GoTo EH
    goUtil.utValidate , txtCommissionPct
    txtCommissionPct.Text = Format(txtCommissionPct.Text, "00")
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub txtCommissionPct_LostFocus"
End Sub

Private Sub txtSunDaysDate_GotFocus()
    goUtil.utSelText txtSunDaysDate
End Sub

Private Sub txtSunDaysDate_LostFocus()
    goUtil.utValidate , txtSunDaysDate
End Sub

Private Sub LoadHeaderlvwSavedPackages()
    On Error GoTo EH
    'set the columnheaders
    With lvwSavedPackages
        .Sorted = True
        .ColumnHeaders.Add , "Name", "Name"
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated"
        .ColumnHeaders.Add , "DateCreated", "Date Created"
        .ColumnHeaders.Add , "DateCreatedSort", "Sort Date Created"
        
        '"Avail WOrd XL Forms"
        .ColumnHeaders.Item(SavedPackagesListView.Name).Width = 5500
        .ColumnHeaders.Item(SavedPackagesListView.Name).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdated).Width = 1500
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdated).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdatedSort).Width = 0  'Hidden
        .ColumnHeaders.Item(SavedPackagesListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateCreated).Width = 1500
        .ColumnHeaders.Item(SavedPackagesListView.DateCreated).Alignment = lvwColumnLeft
        .ColumnHeaders.Item(SavedPackagesListView.DateCreatedSort).Width = 0   'Hidden
        .ColumnHeaders.Item(SavedPackagesListView.DateCreatedSort).Alignment = lvwColumnLeft
       
    End With
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub LoadHeaderlvwSavedPackages"
End Sub



Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    Set moGUI = Nothing
    
    If Not moForm Is Nothing Then
        Unload moForm
        Set moForm = Nothing
    End If
    If Not mArv Is Nothing Then
        mArv.CLEANUP
        Set mArv = Nothing
    End If
    
    CLEANUP = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function
