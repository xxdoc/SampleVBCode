VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form frmProcessData 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebControl (Process Loss Report Data)"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10470
   Icon            =   "frmProcessData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10470
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   32
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Frame framProg 
      Appearance      =   0  'Flat
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.Frame framProcess 
         Caption         =   "Process"
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         Begin VB.TextBox txtProgMessLoss 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1800
            Width           =   4455
         End
         Begin MSComctlLib.ProgressBar ProgBarLoss 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   2160
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.TextBox txtMess 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2220
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame framErrors 
         Caption         =   "Messages"
         Height          =   2535
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   4815
         Begin VB.TextBox txtErrors 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2220
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   240
            Width           =   4575
         End
      End
   End
   Begin VB.Frame framAdjUL 
      Appearance      =   0  'Flat
      Caption         =   "Adjuster Uploads"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   5055
      Begin VB.CheckBox chkXML 
         Caption         =   "XM&L OFF"
         DisabledPicture =   "frmProcessData.frx":0442
         DownPicture     =   "frmProcessData.frx":05FB
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3785
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Value           =   2  'Grayed
         Width           =   1150
      End
      Begin NTService.NTService NTSvcWebControl 
         Left            =   4440
         Top             =   1800
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
         DisplayName     =   "Webcontrol Service Vs 2.0"
         Interactive     =   -1  'True
         ServiceName     =   "V2WebControl"
         StartMode       =   2
      End
      Begin VB.CommandButton cmdShutDownAutoImport 
         Caption         =   "Shut Down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1150
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1150
      End
      Begin VB.Timer Timer_AutoImportReset 
         Enabled         =   0   'False
         Interval        =   50000
         Left            =   3960
         Top             =   1800
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   2880
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcessData.frx":07B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcessData.frx":0C06
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcessData.frx":1058
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer_Status 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   3480
         Top             =   1800
      End
      Begin VB.ListBox lstFTPPaths 
         Appearance      =   0  'Flat
         Height          =   1785
         ItemData        =   "frmProcessData.frx":14AA
         Left            =   120
         List            =   "frmProcessData.frx":14AC
         TabIndex        =   13
         Top             =   1800
         Width           =   4815
      End
      Begin VB.CheckBox chkAutoImport 
         Caption         =   "&Auto Import"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1150
         Left            =   120
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Value           =   2  'Grayed
         Width           =   1150
      End
      Begin VB.CheckBox chkUpdateXMLStatus 
         Caption         =   "Update  XML Status"
         Height          =   435
         Left            =   3785
         TabIndex        =   10
         ToolTipText     =   "Automatically update XML Loss Report Data"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblListFTPPaths 
         Caption         =   "FTP Path List (Upload Paths)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   4125
      End
   End
   Begin VB.Frame framCar 
      Appearance      =   0  'Flat
      Caption         =   "Carrier Loss Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5400
      TabIndex        =   14
      Top             =   3000
      Width           =   4935
      Begin VB.Frame framCbo 
         Height          =   3375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox chkAutoUpdateXMLTrans 
            Caption         =   "Auto Update XML"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2640
            TabIndex        =   30
            ToolTipText     =   "Automatically update XML Loss Report Data"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chkAssignByZIP 
            Caption         =   "Assign By ZIPCODE"
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   840
            Width           =   1935
         End
         Begin MSComctlLib.ImageList imgListClientCompany 
            Left            =   4080
            Top             =   -120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imgListCompany 
            Left            =   3480
            Top             =   -120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin VB.CommandButton cmdProcess 
            Caption         =   "&Process Raw Data"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   31
            Top             =   1365
            Width           =   1935
         End
         Begin VB.CommandButton cmdViewLR 
            Caption         =   " &View Loss Reports"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   28
            Top             =   405
            Width           =   1935
         End
         Begin VB.ComboBox cboCompany 
            Height          =   315
            ItemData        =   "frmProcessData.frx":14AE
            Left            =   120
            List            =   "frmProcessData.frx":14B0
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   465
            Width           =   2205
         End
         Begin VB.CommandButton cmdRawDataPath 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4245
            Picture         =   "frmProcessData.frx":14B2
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Browse"
            Top             =   2295
            Width           =   375
         End
         Begin VB.CommandButton cmdADJFTPPath 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4245
            Picture         =   "frmProcessData.frx":192C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Browse"
            Top             =   2940
            Width           =   375
         End
         Begin VB.TextBox txtRawDataPath 
            Height          =   360
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   2280
            Width           =   4485
         End
         Begin VB.ComboBox cboCar 
            Height          =   315
            ItemData        =   "frmProcessData.frx":1DA6
            Left            =   150
            List            =   "frmProcessData.frx":1DA8
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1065
            Width           =   2205
         End
         Begin VB.ComboBox cboLossFormat 
            Height          =   315
            ItemData        =   "frmProcessData.frx":1DAA
            Left            =   150
            List            =   "frmProcessData.frx":1DAC
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1665
            Width           =   2205
         End
         Begin VB.TextBox txtADJFTPPath 
            Height          =   360
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   2925
            Width           =   4485
         End
         Begin VB.Image imgClientCompany 
            Height          =   300
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   780
            Width           =   615
         End
         Begin VB.Image imgCompany 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblCar 
            Caption         =   "Client Company"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   18
            Top             =   840
            Width           =   1995
         End
         Begin VB.Label lblFormat 
            Caption         =   "Loss Format"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   20
            Top             =   1440
            Width           =   2235
         End
         Begin VB.Label lblRawData 
            Caption         =   "Raw Data Path"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   22
            Top             =   2040
            Width           =   4395
         End
         Begin VB.Label lblADJ_FTP 
            Caption         =   "Adjuster FTP Path"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   25
            Top             =   2685
            Width           =   4515
         End
         Begin VB.Label Label1 
            Caption         =   "Company"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1995
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu barExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuErrorLog 
         Caption         =   "Error &Log"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuRegistry 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuSupport 
         Caption         =   "&Support"
      End
   End
End
Attribute VB_Name = "frmProcessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum PicList
   Idle = 1
   Busy
   Disabled
End Enum

Private WithEvents moLoss As V2ECKeyBoard.clsLossReports
Attribute moLoss.VB_VarHelpID = -1
Private msBaseCarDLLPath As String
Private mbShutDownV2AutoImport As Boolean
Private mbServiceStarted As Boolean
Private mbCleanupLossReports As Boolean
Private msDateLastUpdated As String
Private mbAssignByZip As Boolean
Private mbPendingLossReportsOnly As Boolean
Private mbEditAutoUpdateXMLTrans As Boolean
Private mbXMLUpdatesFlag As Boolean

Private Sub cboCar_click()
    On Error GoTo EH
    Dim sKey As String
    LoadFormats
    txtRawDataPath.Text = vbNullString
    txtADJFTPPath.Text = vbNullString
    EnableViewLoss
    EnableAutoUpdateXML
    sKey = cboCar.ItemData(cboCar.ListIndex)
    On Error Resume Next
    imgClientCompany.Picture = imgListClientCompany.ListImages("""" & sKey & """").Picture
    If Err.Number <> 0 Then
        imgClientCompany.Picture = Nothing
    End If
    On Error GoTo EH
    Exit Sub
EH:
    ShowError Err, "Private Sub cboCar_click", Me
End Sub



Private Sub cboCompany_Click()
    On Error GoTo EH
    Dim sAryCompany() As String
    Dim sKey As String
    
    sAryCompany = Split(cboCompany.Text, "|")
    goUtil.gsCarPrefix = sAryCompany(1)
    
    LoadCarriers
    cboLossFormat.Clear
    txtRawDataPath.Text = vbNullString
    txtADJFTPPath.Text = vbNullString
    EnableViewLoss
    EnableAutoUpdateXML
    sKey = cboCompany.ItemData(cboCompany.ListIndex)
    On Error Resume Next
    imgCompany.Picture = imgListCompany.ListImages("""" & sKey & """").Picture
    If Err.Number <> 0 Then
        imgCompany.Picture = Nothing
    End If
    On Error GoTo EH
    Exit Sub
EH:
    ShowError Err, "Private Sub cboCompany_Click", Me
End Sub


Private Sub cboLossFormat_Click()
    LoadDataPaths
End Sub

Private Sub chkAssignByZIP_Click()
    On Error GoTo EH
    Dim lRet As VBA.VbMsgBoxResult
    Dim sMess As String
    Dim sRet As String
    If chkAssignByZIP.Value = vbChecked Then
        sMess = "Are you sure the selected Client Company is currently assigning ALL its claims by ZIPCODE?"
        lRet = MsgBox(sMess, vbQuestion + vbYesNo, "Assign by ZIPCODE")
        If lRet = vbYes Then
            sRet = InputBox("Enter Password", "Enter PassWord to Assign By ZIPCODE")
            If sRet <> Format(Now(), "DDYYMM") Then
                chkAssignByZIP.Value = vbUnchecked
                Exit Sub
            End If
        Else
            chkAssignByZIP.Value = vbUnchecked
            Exit Sub
        End If
        mbAssignByZip = True
    Else
        mbAssignByZip = False
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub chkAssignByZIP_Click", Me
End Sub

Private Sub chkAutoImport_Click()
    If chkAutoImport.Value = vbChecked Then
        WCService NetContinue
    ElseIf chkAutoImport.Value = vbUnchecked Then
        WCService NetPause
    ElseIf chkAutoImport.Value = vbGrayed Then
        SaveSetting "V2AutoImport", "Msg", "Status", 0
    End If
End Sub

Private Sub chkAutoUpdateXMLTrans_Click()
    On Error GoTo EH
    Dim lRet As VBA.VbMsgBoxResult
    Dim sMess As String
    Dim sRet As String
    Dim lCheckValue As Long
    
    lCheckValue = chkAutoUpdateXMLTrans.Value
    
    If Not mbEditAutoUpdateXMLTrans Then
        sMess = "Are you sure you want to change the " & chkAutoUpdateXMLTrans.Caption & " Option?"
        lRet = MsgBox(sMess, vbQuestion + vbYesNo, chkAutoUpdateXMLTrans.Caption)
        If lRet = vbYes Then
            sRet = InputBox("Enter Password", "Enter PassWord to change the " & chkAutoUpdateXMLTrans.Caption & " Option.")
            If sRet <> Format(Now(), "DDYYMM") Then
                'If don't get password correct change the option back
                If lCheckValue = vbChecked Then
                    lCheckValue = vbUnchecked
                Else
                    lCheckValue = vbChecked
                End If
                mbEditAutoUpdateXMLTrans = True
                chkAutoUpdateXMLTrans.Value = lCheckValue
                mbEditAutoUpdateXMLTrans = False
                If lCheckValue = vbChecked Then
                    sMess = "Invalid Password entered while Checking " & chkAutoUpdateXMLTrans.Caption & vbCrLf & vbCrLf
                Else
                    sMess = "Invalid Password entered while Unchecking " & chkAutoUpdateXMLTrans.Caption & vbCrLf & vbCrLf
                End If
                sMess = sMess & "Invalid Password: " & sRet & vbCrLf
                sMess = sMess & "Company: " & RTrim(left(cboCompany.Text, 50)) & vbCrLf
                sMess = sMess & "Client CO: " & RTrim(left(cboCar.Text, 50)) & vbCrLf
                sMess = sMess & "Loss Format: " & RTrim(left(cboLossFormat.Text, 50)) & vbCrLf
                ErrorLog sMess
                Exit Sub
            Else
                SaveSetting App.EXEName, "Dir", goUtil.gsCarPrefix & "_" & cboCar.Text & "_" & cboLossFormat.Text & "_chkAutoUpdateXMLTrans", chkAutoUpdateXMLTrans.Value
                If lCheckValue = vbChecked Then
                    sMess = "Sucessfully Checked " & chkAutoUpdateXMLTrans.Caption & vbCrLf & vbCrLf
                Else
                    sMess = "Sucessfully Unchecked  " & chkAutoUpdateXMLTrans.Caption & vbCrLf & vbCrLf
                End If
                sMess = sMess & "Company: " & RTrim(left(cboCompany.Text, 50)) & vbCrLf
                sMess = sMess & "Client CO: " & RTrim(left(cboCar.Text, 50)) & vbCrLf
                sMess = sMess & "Loss Format: " & RTrim(left(cboLossFormat.Text, 50)) & vbCrLf
                ErrorLog sMess
            End If
        Else
            If lCheckValue = vbChecked Then
                lCheckValue = vbUnchecked
            Else
                lCheckValue = vbChecked
            End If
            mbEditAutoUpdateXMLTrans = True
            chkAutoUpdateXMLTrans.Value = lCheckValue
            mbEditAutoUpdateXMLTrans = False
            Exit Sub
        End If
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub chkAutoUpdateXMLTrans_Click", Me
End Sub

Private Sub chkUpdateXMLStatus_Click()
    On Error GoTo EH
    Dim lRet As VBA.VbMsgBoxResult
    Dim sMess As String
    Dim sRet As String
    Dim lCheckValue As Long
    
    lCheckValue = chkUpdateXMLStatus.Value
    
    If Not mbEditAutoUpdateXMLTrans Then
        sMess = "Are you sure you want to " & chkUpdateXMLStatus.Caption & "?"
        lRet = MsgBox(sMess, vbQuestion + vbYesNo, chkUpdateXMLStatus.Caption)
        If lRet = vbYes Then
            sRet = InputBox("Enter Password", "Enter PassWord to " & chkUpdateXMLStatus.Caption & ".")
            If sRet <> Format(Now(), "DDYYMM") Then
                'If don't get password correct change the option back
                If lCheckValue = vbChecked Then
                    lCheckValue = vbUnchecked
                Else
                    lCheckValue = vbChecked
                End If
                mbEditAutoUpdateXMLTrans = True
                chkUpdateXMLStatus.Value = lCheckValue
                mbEditAutoUpdateXMLTrans = False
                If lCheckValue = vbChecked Then
                    sMess = "Invalid Password entered while Checking " & chkUpdateXMLStatus.Caption & vbCrLf & vbCrLf
                Else
                    sMess = "Invalid Password entered while Unchecking " & chkUpdateXMLStatus.Caption & vbCrLf & vbCrLf
                End If
                sMess = sMess & "Invalid Password: " & sRet
                ErrorLog sMess
                Exit Sub
            Else
                If lCheckValue = vbChecked Then
                    chkXML.Enabled = True
                    sMess = "Sucessfully Checked " & chkUpdateXMLStatus.Caption & vbCrLf & vbCrLf
                Else
                    chkXML.Enabled = False
                    sMess = "Sucessfully Unchecked  " & chkUpdateXMLStatus.Caption & vbCrLf & vbCrLf
                End If
                ErrorLog sMess
            End If
        Else
            If lCheckValue = vbChecked Then
                lCheckValue = vbUnchecked
                chkXML.Enabled = False
            Else
                lCheckValue = vbChecked
                chkXML.Enabled = True
            End If
            mbEditAutoUpdateXMLTrans = True
            chkUpdateXMLStatus.Value = lCheckValue
            mbEditAutoUpdateXMLTrans = False
            Exit Sub
        End If
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub chkUpdateXMLStatus_Click", Me
End Sub


Private Sub chkXML_Click()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMess As String
    
    If chkXML.Value = vbChecked Then
        SaveSetting "V2AutoImport", "Msg", "XML_UPDATES", True
        chkXML.Caption = "XM&L ON" & vbCrLf
        sMess = "XML ON!"
        EnableFramCBO False
    ElseIf chkXML.Value = vbUnchecked Then
        SaveSetting "V2AutoImport", "Msg", "XML_UPDATES", False
        chkXML.Caption = "XM&L OFF" & vbCrLf
        sMess = "XML OFF!"
        EnableFramCBO True
    End If
    
    ErrorLog sMess
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "Error Number# " & lErrNum & vbCrLf
    sMess = sMess & "Error Desc: " & sErrDesc & vbCrLf
    sMess = sMess & "Private Sub chkXML_Click" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Function EnableFramCBO(pbEnabled As Boolean) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sMess As String
    Dim oControl As Object
    
    For Each oControl In Me.Controls
        If StrComp(oControl.Name, NTSvcWebControl.Name, vbTextCompare) <> 0 _
            And Not TypeOf oControl Is ImageList _
            And Not TypeOf oControl Is Timer _
            And Not TypeOf oControl Is Menu Then
            If StrComp(oControl.Container.Name, framCbo.Name, vbTextCompare) = 0 Then
                If TypeOf oControl Is CommandButton Then
                    oControl.Visible = pbEnabled
                ElseIf TypeOf oControl Is CheckBox Then
                    oControl.Visible = pbEnabled
                Else
                    oControl.Enabled = pbEnabled
                End If
            End If
        End If
    Next
    
    framCbo.Enabled = pbEnabled
    
    EnableFramCBO = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sMess = "Error Number# " & lErrNum & vbCrLf
    sMess = sMess & "Error Desc: " & sErrDesc & vbCrLf
    sMess = sMess & "Private Function EnableFramCBO" & vbCrLf & vbCrLf
    ErrorLog sMess
    ErrorLog sMess
End Function

Private Sub cmdADJFTPPath_Click()
    On Error GoTo EH
    txtADJFTPPath.Text = GetPath("ADJFTPPath", "Browse to " & cboCar.Text & " " & cboLossFormat.Text & " ADJ_FTP Path", "CLICK OPEN TO SAVE PATH", txtADJFTPPath.Text, Me.hWnd)
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdADJFTPPath_Click", Me
End Sub

Private Sub cmdShutDownAutoImport_Click()
    On Error GoTo EH
    If chkAutoImport.Value <> vbGrayed Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Shut Down Auto Import") = vbYes Then
            mbShutDownV2AutoImport = True
        End If
    End If
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdShutDownAutoImport_Click", Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
        ' Just hide form if user presses Close button
        FormWinRegPos Me, True
        SaveSetting "V2Webcontrol", "Msg", "WebControlVisible", False
        Cancel = True
    End If
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Form_QueryUnload" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Sub lstFTPPaths_Click()
    On Error GoTo EH
    
    lstFTPPaths.ToolTipText = lstFTPPaths.Text
    
    Exit Sub
EH:
    ShowError Err, "Private Sub lstFTPPaths_Click", Me
End Sub

Private Sub lstFTPPaths_DblClick()
    On Error GoTo EH
    Dim sFTPPaths As String
    Dim lCount As Long
    
    If lstFTPPaths.ListIndex > -1 Then
        If MsgBox("Remove selected path?", vbYesNo + vbQuestion, "Remove Path") = vbYes Then
            lstFTPPaths.RemoveItem lstFTPPaths.ListIndex
        End If
    End If
    
    For lCount = 0 To lstFTPPaths.ListCount - 1
        sFTPPaths = sFTPPaths & lstFTPPaths.List(lCount) & LIST_ADJFTP_DELIM
    Next
    
    SaveSetting App.EXEName, "Dir", "ListADJFTP", sFTPPaths
    
    Exit Sub
EH:
    ShowError Err, "Private Sub lstFTPPaths_DblClick", Me
End Sub

Private Sub mnuErrorLog_Click()
    On Error GoTo EH
    
    Load frmErrorLog
    frmErrorLog.Show vbModeless
    frmErrorLog.WindowState = vbNormal
    Exit Sub
EH:
    ShowError Err, "Private Sub mnuErrorLog_Click", Me
End Sub

Private Sub mnuExit_Click()
    ' Just hide form if user presses Close button
    FormWinRegPos Me, True
    SaveSetting "V2WebControl", "Msg", "WebControlVisible", False
End Sub

Private Sub mnuRegistry_Click()
    Load frmWebRegSettings
    frmWebRegSettings.Show vbModeless
    frmWebRegSettings.WindowState = vbNormal
End Sub

Private Sub mnuSupport_Click()
    On Error GoTo EH
    Load frmSupport
    frmSupport.framEasyClaim.Caption = "Webcontrol Support"
    frmSupport.Show vbModeless
    Exit Sub
EH:
    ShowError Err, "Private Sub mnuSupport_Click", Me
End Sub

Private Sub moLoss_CleanUpLossReports()
    On Error GoTo EH
    
    mbCleanupLossReports = True
    
    Exit Sub
EH:
    ShowError Err, "Private Sub moLoss_CleanUpLossReports", Me
End Sub

Private Sub moLoss_GetAssignByZip(bAssignByZIP As Boolean)
    On Error GoTo EH
    
    bAssignByZIP = mbAssignByZip
    
    Exit Sub
EH:
    ShowError Err, "Private Sub moLoss_GetAssignByZip", Me
End Sub

Private Sub moLoss_MemoryCleanUpFinished()
    On Error Resume Next
    txtMess.Text = vbNullString
    txtMess.Refresh
    Me.Refresh
End Sub


Private Sub moLoss_PopulateLossReportsRS(RSLossReports As ADODB.Recordset)
    On Error GoTo EH
    Dim sSQL As String
    Dim lUID As Long
    
    lUID = GetUID(, False)
    If lUID = 0 Then
        GoTo CLEAN_UP
    End If
    
    'Need to Get Companies, not client Companies
    'that the DB User Name has access to...
    sSQL = "z_spsGetAssignmentsInfo "
    sSQL = sSQL & "1, "                         '@bHideDeleted
    sSQL = sSQL & lUID & ", "                   '@UID
    sSQL = sSQL & "null, "                      '@AssignmentsID
    sSQL = sSQL & "1, "                         '@bShowAssignmentsInfo
    sSQL = sSQL & "1, "                         '@bShowAssignmentTypeInfo
    sSQL = sSQL & "1, "                         '@bShowCompanyInfo
    sSQL = sSQL & "1, "                         '@bShowCatInfo
    sSQL = sSQL & "1, "                         '@bShowClientCompanyInfo
    sSQL = sSQL & "0, "                         '@bShowFeeScheduleInfo
    sSQL = sSQL & "1, "                         '@bShowClientCompanyCatInfo
    sSQL = sSQL & "1, "                         '@bShowClientCompanyCatSpecInfo
    sSQL = sSQL & "1, "                         '@bShowAdjusterSpecUsersInfo
    sSQL = sSQL & "1, "                         '@bShowAdjusterSpecInfo
    sSQL = sSQL & "1, "                         '@bShowStatusInfo
    sSQL = sSQL & "1, "                         '@bShowTypeOfLossInfo
    sSQL = sSQL & "1, "                         '@bShowRAAdjusterSpecUsersInfo
    sSQL = sSQL & "1, "                         '@bShowRAAdjusterSpecInfo
    sSQL = sSQL & "0, "                         '@bShowBatchesInfo
    sSQL = sSQL & "null, "                      '@OrderBy
    sSQL = sSQL & "null, "                      '@GroupBy
    '@SearchBy
    sSQL = sSQL & "' "
    If msDateLastUpdated <> vbNullString Then
        sSQL = sSQL & "And DateLastUpdated >=Convert(DateTime,''" & msDateLastUpdated & "'') "
    End If
    If mbPendingLossReportsOnly Then
        sSQL = sSQL & "And StatusStatus = ''PENDING'' "
    End If
    'Only show records for the selected Client Company
    sSQL = sSQL & "And ClientCompanyCatSpecID IN    ( "
    sSQL = sSQL & "                                 SELECT  ClientCompanyCatSpecID "
    sSQL = sSQL & "                                 FROM    ClientCompanyCatSpec "
    sSQL = sSQL & "                                 WHERE   ClientCompanyID = " & cboCar.ItemData(cboCar.ListIndex) & " "
    sSQL = sSQL & "                                 ) "
    sSQL = sSQL & "' "
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    'Need this so RecordCount will populate and All fields will populate
    RSLossReports.CursorLocation = adUseClient
    RSLossReports.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    Set RSLossReports.ActiveConnection = Nothing

    
CLEAN_UP:
    Exit Sub
EH:
    ShowError Err, "Private Sub moLoss_PopulateLossReportsRS", Me
End Sub

Private Sub NTSvcWebControl_Continue(Success As Boolean)
    SaveSetting "V2WebControl", "Msg", "Status", PicList.Idle
    Success = True
End Sub

Private Sub NTSvcWebControl_Pause(Success As Boolean)
    SaveSetting "V2WebControl", "Msg", "Status", PicList.Disabled
    Success = True
End Sub

Private Sub NTSvcWebControl_Start(Success As Boolean)
    SaveSetting "V2WebControl", "Msg", "Status", PicList.Idle
    SaveSetting "V2WebControl", "Msg", "WebControlVisible", True
    Success = True
    mbServiceStarted = True
End Sub

Private Sub NTSvcWebControl_Stop()
    Timer_Status.Enabled = False
    Timer_AutoImportReset.Enabled = False
    Set gDB = Nothing
    Set gWS = Nothing
    CloseConnection
    DoEvents
    Unload Me
    SaveSetting "V2WebControl", "Msg", "Status", 0
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer_AutoImportReset_Timer()
    On Error GoTo EH
    Dim iStatus As Integer
    Dim sMess As String
    Dim sCommandLine As String
    Dim sShellBat As String
    
    
    
    On Error Resume Next
    iStatus = GetSetting("V2WebControl", "Msg", "Status", 0)
    If Err.Number > 0 Then
        Err.Clear
        iStatus = 0
        SaveSetting "V2WebControl", "Msg", "Status", iStatus
    End If
    On Error GoTo EH
    If iStatus > 0 Then
        sCommandLine = "RunAsDepOfV2WebControlService"
        '---------------------------DE BUG------------------------------
        
        sShellBat = """" & App.Path & "\V2AutoImport.exe "" " & sCommandLine

        '---------------------------DE BUG------------------------------
        'Need to reset the util object if user logs of session but does not
        'restart server
        If goUtil Is Nothing Then
            SetUtilObject
        End If
        goUtil.utSaveFileData App.Path & "\V2AutoImport.bat", sShellBat
        Shell App.Path & "\V2AutoImport.bat", vbHide
        SaveSetting "V2WebControl", "Msg", "Reset", False
        Timer_AutoImportReset.Enabled = False
        Timer_Status.Enabled = True
    End If
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Timer_AutoImportReset_Timer" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Function LoadAutoUpdateXMLNames() As Variant
    On Error GoTo EH
    Dim oReg As V2ECKeyBoard.clsRegSetting
    Dim vXML As Variant
    Dim lCount As Long
    Dim sXMLKeyName As String
    Dim bchkAutoUpdateXMLTransValue As Boolean
    Dim lNameCount As Long
    Dim saryXMLNames() As String
       
    'HKEY_USERS\.DEFAULT\Software\VB and VBA Program Settings\V2WebControl\Dir
    'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\V2WebControl\Dir
    Set oReg = New V2ECKeyBoard.clsRegSetting
    'Enumerate all the DSN names in the Registry
    vXML = oReg.EnumValues(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & App.EXEName & "\Dir\")
    'Add them to the List
    If DynamicArraySet(vXML) Then
        For lCount = 0 To UBound(vXML, 1)
            If vXML(lCount, 0) <> vbNullString Then
                sXMLKeyName = vXML(lCount, 0)
                If InStr(1, sXMLKeyName, "_chkAutoUpdateXMLTrans", vbTextCompare) > 0 Then
                    bchkAutoUpdateXMLTransValue = GetSetting(App.EXEName, "Dir", sXMLKeyName, False)
                    If bchkAutoUpdateXMLTransValue Then
                        lNameCount = lNameCount + 1
                        ReDim Preserve saryXMLNames(1 To lNameCount)
                        saryXMLNames(lNameCount) = sXMLKeyName
                    End If
                End If
            End If
        Next
    End If
    
    If DynamicArraySet(saryXMLNames) Then
        LoadAutoUpdateXMLNames = saryXMLNames
    Else
        LoadAutoUpdateXMLNames = vbNullString
    End If
    
    'CLeanup
    Set oReg = Nothing
    Exit Function
EH:
    ShowError Err, "Private Function LoadAutoUpdateXMLNames", Me
End Function

Private Sub Timer_Status_Timer()
    On Error GoTo EH
    Dim iStatus As Integer
    Dim bV2WebControlVisible As Boolean
    Dim bResetV2AutoImport As Boolean
    Dim sErrorMess As String
    Dim sMess As String
    Dim sShellBat As String
    Dim sCommandLine As String
    Dim vAutoUpdateXMLNames As Variant
    Dim lPos As Long
    Dim saryXMLNameParts() As String
    Dim sCompanyPrefix As String
    Dim sClientCo As String
    Dim sLossFormat As String
    
    On Error Resume Next
    bV2WebControlVisible = CBool(GetSetting("V2WebControl", "Msg", "WebControlVisible", True))
    If Err.Number > 0 Then
        Err.Clear
        bV2WebControlVisible = True
        SaveSetting "V2WebControl", "Msg", "WebControlVisible", bV2WebControlVisible
    End If
    On Error GoTo EH
    
    If Me.Visible <> bV2WebControlVisible Then
        Me.Visible = bV2WebControlVisible
        If bV2WebControlVisible Then
            On Error Resume Next
            AppActivate Me.Caption
            If Err.Number > 0 Then
                Err.Clear
            End If
            On Error GoTo EH
        End If
    End If
    
    If mbCleanupLossReports Then
        mbCleanupLossReports = False
        If Not moLoss Is Nothing Then
            moLoss.CLEANUP
            Set moLoss = Nothing
        End If
    End If
    
    'If user clicked shutdown button
    If mbShutDownV2AutoImport Then
        SaveSetting "V2WebControl", "Msg", "Status", 0
        iStatus = 0
        mbShutDownV2AutoImport = False
    Else
        On Error Resume Next
        iStatus = GetSetting("V2WebControl", "Msg", "Status", 0)
        If Err.Number > 0 Then
            Err.Clear
            iStatus = 0
            SaveSetting "V2WebControl", "Msg", "Status", iStatus
        End If
        On Error GoTo EH
    End If
    
    'Only reset V2AutoImport if user logged off windows session.
    'V2AutoImport will shutdown when a user loggs off because it is not
    'controlled by the services panel.
    On Error Resume Next
    bResetV2AutoImport = CBool(GetSetting("V2WebControl", "Msg", "Reset", False))
    If Err.Number > 0 Then
        Err.Clear
        bResetV2AutoImport = False
        SaveSetting "V2WebControl", "Msg", "Reset", bResetV2AutoImport
    End If
    On Error GoTo EH
    If bResetV2AutoImport And Not Timer_AutoImportReset.Enabled And iStatus > 0 Then
        If iStatus > 0 Then
            Timer_AutoImportReset.Enabled = True
            Timer_Status.Enabled = False
            Exit Sub
        End If
    End If
    
    'Check for any errors that V2AutoImport is logging and send them to
    'NT service event viewr and error log
    sErrorMess = GetSetting("V2WebControl", "Msg", "ErrorMess", vbNullString)
    If sErrorMess <> vbNullString Then
        ErrorLog sErrorMess
        SaveSetting "V2WebControl", "Msg", "ErrorMess", vbNullString
    End If
    
    'Check for any errors that ECUpdateBatches is logging and send them to
    'NT service event viewr and error log
    sErrorMess = GetSetting("ECUpdateBatches", "Msg", "ErrorMess", vbNullString)
    If sErrorMess <> vbNullString Then
        ErrorLog sErrorMess
        SaveSetting "ECUpdateBatches", "Msg", "ErrorMess", vbNullString
    End If

    Select Case iStatus
        Case PicList.Busy
            If chkAutoImport.Picture <> imgList.ListImages(iStatus).Picture Then
                chkAutoImport.Picture = imgList.ListImages(iStatus).Picture
                chkAutoImport.Value = vbChecked
                chkAutoImport.Caption = "&Auto Import (Busy)"
                framAdjUL.Enabled = False
            End If
        Case PicList.Idle
            If chkAutoImport.Picture <> imgList.ListImages(iStatus).Picture Then
                sCommandLine = "RunAsDepOfV2WebControlService"
                '---------------------------DE BUG------------------------------
                 sShellBat = """" & App.Path & "\V2AutoImport.exe "" " & sCommandLine
                '---------------------------DE BUG------------------------------
                goUtil.utSaveFileData App.Path & "\V2AutoImport.bat", sShellBat
                Shell App.Path & "\V2AutoImport.bat", vbHide
                chkAutoImport.Picture = imgList.ListImages(iStatus).Picture
                chkAutoImport.Value = vbChecked
                chkAutoImport.Caption = "&Auto Import (ON)"
                framAdjUL.Enabled = True
            Else
                'If Idle Check for XML Transactions
                'Set the Check Button
                mbXMLUpdatesFlag = GetSetting("V2AutoImport", "Msg", "XML_UPDATES", False)
                If mbXMLUpdatesFlag Then
                    chkXML.Value = vbChecked
                Else
                    chkXML.Value = vbUnchecked
                End If
                If mbXMLUpdatesFlag Then
                    vAutoUpdateXMLNames = LoadAutoUpdateXMLNames()
                    If IsArray(vAutoUpdateXMLNames) Then
                        For lPos = LBound(vAutoUpdateXMLNames, 1) To UBound(vAutoUpdateXMLNames, 1)
                            saryXMLNameParts() = Split(vAutoUpdateXMLNames(lPos), "_", , vbBinaryCompare)
                            sCompanyPrefix = saryXMLNameParts(0)
                            sClientCo = saryXMLNameParts(1)
                            sLossFormat = saryXMLNameParts(2)
                            DoEvents
                            Sleep 100
                            ProcessStuff True, sCompanyPrefix, sClientCo, sLossFormat
                        Next
                    End If
                End If
            End If
        Case PicList.Disabled
            If chkAutoImport.Picture <> imgList.ListImages(iStatus).Picture Then
                sCommandLine = "RunAsDepOfV2WebControlService"
                '---------------------------DE BUG------------------------------
                sShellBat = """" & App.Path & "\V2AutoImport.exe "" " & sCommandLine
                '---------------------------DE BUG------------------------------
                goUtil.utSaveFileData App.Path & "\V2AutoImport.bat", sShellBat
                Shell App.Path & "\V2AutoImport.bat", vbHide
                chkAutoImport.Picture = imgList.ListImages(iStatus).Picture
                chkAutoImport.Value = vbUnchecked
                chkAutoImport.Caption = "&Auto Import (OFF)"
                framAdjUL.Enabled = True
            End If
        Case Else
            If chkAutoImport.Picture Then
                chkAutoImport.Picture = Nothing
                chkAutoImport.Value = vbGrayed
                chkAutoImport.Caption = "&Auto Import"
                framAdjUL.Enabled = True
            End If
    End Select
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Timer_Status_Timer" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub


Private Sub cmdRawDataPath_Click()
    On Error GoTo EH
    txtRawDataPath.Text = GetPath("RawDataPath", "Browse to " & cboCar.Text & " " & cboLossFormat.Text & " Raw Data", "CLICK OPEN TO SAVE PATH", txtRawDataPath.Text, Me.hWnd)
    Exit Sub
EH:
    ShowError Err, "Private Sub cmdRawDataPath_Click", Me
End Sub

Private Sub cmdExit_Click()
    ' Just hide form if user presses Close button
    FormWinRegPos Me, True
    SaveSetting "V2WebControl", "Msg", "WebControlVisible", False
    If Not moLoss Is Nothing Then
        moLoss.CLEANUP
        Set moLoss = Nothing
    End If
    
End Sub

Private Sub cmdProcess_Click()
    ProcessStuff False
End Sub

Public Sub ProcessStuff(Optional bHiddenProcess As Boolean, _
                        Optional sCompanyPrefix As String, _
                        Optional sClientCo As String, _
                        Optional sLossFormat As String)
    On Error GoTo EH
    Dim sMess As String
    Dim sclsLoss As String
    Dim sAppEXEName As String
    Dim sFormat As String
    Dim sRawDataPath As String
    Dim sFTPPath As String
    Dim lPos As Long
    Dim bFoundCompany As Boolean
    Dim bFoundClientCo As Boolean
    Dim bFoundLossFormat As Boolean
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not moLoss Is Nothing Then
        Exit Sub
    End If
    
    If goUtil Is Nothing Then
        Exit Sub
    End If
    
    If Not bHiddenProcess Then
        'Check for existing messages
        If txtErrors.Text <> vbNullString Then
            sMess = "Do you want to reset your previous messages ? "
            If MsgBox(sMess, vbQuestion + vbYesNo, "Reset Messages Box") = vbYes Then
                txtErrors.Text = vbNullString
                txtErrors.Refresh
            End If
        End If
    Else
        If sCompanyPrefix <> vbNullString And sClientCo <> vbNullString And sLossFormat <> vbNullString Then
            'Select Company via the Company Prefix
            For lPos = 0 To cboCompany.ListCount - 1
                If InStr(1, cboCompany.List(lPos), sCompanyPrefix, vbTextCompare) > 0 Then
                    cboCompany.ListIndex = lPos
                    bFoundCompany = True
                    Exit For
                End If
            Next
            If Not bFoundCompany Then
                Exit Sub
            Else
                Sleep 500
            End If
            
            'Select Client Company
            For lPos = 0 To cboCar.ListCount - 1
                If StrComp(RTrim(left(cboCar.List(lPos), 20)), sClientCo, vbTextCompare) = 0 Then
                    cboCar.ListIndex = lPos
                    bFoundClientCo = True
                    Exit For
                End If
            Next
            If Not bFoundClientCo Then
                Exit Sub
            Else
                Sleep 500
            End If
            
            'Select Loss Format
            For lPos = 0 To cboLossFormat.ListCount - 1
                If StrComp(RTrim(left(cboLossFormat.List(lPos), 20)), sLossFormat, vbTextCompare) = 0 Then
                    cboLossFormat.ListIndex = lPos
                    bFoundLossFormat = True
                    Exit For
                End If
            Next
            If Not bFoundLossFormat Then
                Exit Sub
            Else
                DoEvents
                Sleep 500
            End If
            
        Else
            Exit Sub
        End If
    End If
    
    If goUtil Is Nothing Then
        Exit Sub
    End If
    
    If cboCar.Text = "V2ECKeyBoard" Then
        sAppEXEName = "V2ECKeyBoard"
    Else
        sAppEXEName = goUtil.gsCarPrefix & cboCar.Text
    End If
    sFormat = cboLossFormat.Text
    sRawDataPath = txtRawDataPath.Text
    sFTPPath = txtADJFTPPath.Text
    
    Set moLoss = New V2ECKeyBoard.clsLossReports
    
    moLoss.SetUtilObject goUtil
    moLoss.IgnoreProcessRawDataErrors = False
    'Very Important to set the Application
    moLoss.APPEXEName = sAppEXEName
    
    '1.5.2004 Not Applicable in V2
'    SendAdjusterTable
   
    'Current version Web Control will process ASN and CCMS formats
    'as well, we handle processing but not DB update of Unknown Text only formats
    'ASN
    UpdateProcessMessage sFormat, sRawDataPath, sFTPPath
    msDateLastUpdated = DateAdd("n", -10, Now())
    If Not moLoss.ProcessRawData(sFormat, sRawDataPath, sFTPPath, ProgBarLoss, txtProgMessLoss) Then
        moLoss.CLEANUP
        Set moLoss = Nothing
        GoTo CLEAN_UP
    End If
    If Not bHiddenProcess Then
        'Show Loss Reports
        txtMess.Text = "Showing Loss Reports..." & vbCrLf & "Loss Report Path: " & sFTPPath
        txtMess.Refresh
        Set goUtil.goProgForm = New V2ECKeyBoard.clsProgForm
        goUtil.goProgForm.SetUtilObject goUtil
        With goUtil.goProgForm
            .LoadForm
            .Caption = "Loss Report Progress"
            .framTableText = vbNullString
            .framRecordText = vbNullString
            .framFileText = "Loss Reports"
            .cmdCancelEnable = True
            .ShowForm
        End With
        moLoss.ShowLossReports , goUtil.goProgForm
    Else
        moLoss.CLEANUP
        Set moLoss = Nothing
        GoTo CLEAN_UP
    End If
    
CLEAN_UP:
    txtMess.Text = "Ready..."
    ProgBarLoss.Value = 0
    Me.MousePointer = vbDefault
    msDateLastUpdated = vbNullString
    Exit Sub
EH:
    msDateLastUpdated = vbNullString
    Me.MousePointer = vbDefault
    If Not bHiddenProcess Then
        ShowError Err, "Public Sub ProcessStuff", Me
    Else
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sMess = "Error Number# " & lErrNum & vbCrLf
        sMess = sMess & "Error Desc: " & sErrDesc & vbCrLf
        sMess = sMess & "Public Sub ProcessStuff" & vbCrLf & vbCrLf
        'Dont turn off xml for certain errors
        If lErrNum = -2147467259 Then
            'Dont turn off xml for Deadlock errors
            'let the process refresh and try again.
        Else
            sMess = sMess & "Turning Off XML Updates! " & Now()
            chkXML.Value = vbUnchecked
        End If
        
        ErrorLog sMess
        txtErrors.Text = txtErrors.Text & vbCrLf & vbCrLf & sMess
        txtErrors.Refresh
    End If
    Set moLoss = Nothing
End Sub

Private Sub UpdateProcessMessage(psFOrmat As String, psLossPath As String, psOutPath As String)
    On Error GoTo EH
    Dim sMess As String

    sMess = "Processing " & psFOrmat & " Format." & vbCrLf
    sMess = sMess & "Raw Data Path: " & psLossPath & vbCrLf
    
    txtMess.Text = sMess
    txtMess.Refresh
    
    Exit Sub
EH:
    ShowError Err, "Private Sub UpdateProcessMessage", Me
End Sub

Private Sub cmdViewLR_Click()
    On Error GoTo EH
    Dim sAppEXEName As String
    Dim sFTPPath As String
    Dim oForm As Form
    
    If Not moLoss Is Nothing Then
        Exit Sub
    End If
    
    Set moLoss = New V2ECKeyBoard.clsLossReports
    moLoss.SetUtilObject goUtil
    
    sFTPPath = txtADJFTPPath.Text
    Set goUtil.goProgForm = New V2ECKeyBoard.clsProgForm
    goUtil.goProgForm.SetUtilObject goUtil
    With goUtil.goProgForm
        .LoadForm
        .Caption = "Loss Report Progress"
        .framTableText = vbNullString
        .framRecordText = vbNullString
        .framFileText = "Loss Reports"
        .cmdCancelEnable = True
        .ShowForm
    End With
    
    'Show Loss Reports
    txtMess.Text = "Showing Loss Reports..." & vbCrLf & "Loss Report Path: " & sFTPPath
    txtMess.Refresh
    mbPendingLossReportsOnly = True
    moLoss.ShowLossReports True, goUtil.goProgForm
    mbPendingLossReportsOnly = False
    txtMess.Text = "Ready..."
    Me.MousePointer = vbDefault
    Exit Sub
EH:
    mbPendingLossReportsOnly = False
    Me.MousePointer = vbDefault
    ShowError Err, "Private Sub cmdViewLR_Click", Me
    Set moLoss = Nothing
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim iStatus As Integer
    Dim fValue As String
    Dim Value As String
    Dim sMess As String
    Dim lCount As Long
    
    'NT Service OPtions
    If Command$ = "-installWebcontrol2.0" Then
        If NTSvcWebControl.Install Then
            MsgBox NTSvcWebControl.DisplayName & " installed successfully.", vbInformation
        Else
            MsgBox NTSvcWebControl.DisplayName & " install failed!", vbCritical
        End If
        End
        Exit Sub
    ElseIf Command$ = "-uninstallWebcontrol2.0" Then
        If NTSvcWebControl.Uninstall Then
            MsgBox NTSvcWebControl.DisplayName & " uninstalled successfully.", vbInformation
        Else
            MsgBox NTSvcWebControl.DisplayName & " uninstall failed!", vbCritical
        End If
        End
        Exit Sub
    ElseIf Command$ = "DEBUGME" Then
       'Debug the Service
        NTSvcWebControl.Debug = True
    ElseIf Command$ <> vbNullString Then
        MsgBox NTSvcWebControl.DisplayName & " Invalid command!", vbCritical
        End
        Exit Sub
    End If
    
    ' Enable Pause/Continue. Must be set before
    ' StartService is called or in design mode
    NTSvcWebControl.ControlsAccepted = svcCtrlPauseContinue
    ' connect service to Windows NT services controller
    NTSvcWebControl.StartService
    
    For lCount = 1 To 5
        DoEvents
        Sleep 500
    Next
    
    If Not mbServiceStarted Then
        End
        Exit Sub
    End If
    
RUN_SERVICE:
    
    SetUtilObject
    
    FormWinRegPos Me
    
    Timer_Status.Enabled = True
    
    txtMess.Text = "Ready..."
    
    'Set up the Carrier selection drop downs
    LoadCompanies
    LoadUploadPaths
    
    'Clean up Temp Files
    goUtil.DelTempDirFiles
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "Private Sub Form_Load" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    ErrorLog sMess
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set gDB = Nothing
    Set gWS = Nothing
    CloseConnection
    If Not goUtil Is Nothing Then
        goUtil.CLEANUP
        Set goUtil = Nothing
    End If
    
    Exit Sub
End Sub


Private Sub moLoss_ErrorMess(ByVal Mess As String)
    txtErrors.Text = txtErrors.Text & Mess
    txtErrors.Refresh
    ErrorLog Mess
End Sub

Private Sub moLoss_MemoryCleanUpAlert()
    On Error Resume Next
    txtMess.Text = "Memory Cleanup... Pleas Wait."
    txtMess.Refresh
    Me.Refresh
End Sub

'Private Sub moLoss_UpdateAdjusters(colAdjusters As Collection)
'    On Error GoTo EH
'    Dim sDSN As String
'    Dim sMess As String
'
'    sDSN = GetSetting("V2WebControl", "DSN", "NAME", "ACCESS_2000")
'
'    If sDSN = "ACCESS_2000" Then
'        UpdateAdjusters_ACCESS2000 colAdjusters
'    Else
'        'CODE FOR SQL SERVER
'        UpdateAdjusters_SQLServer colAdjusters
'    End If
'
'    Exit Sub
'EH:
'    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
'    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
'    sMess = sMess & Err.Description & vbCrLf & vbCrLf
'    sMess = sMess & "Private Sub moLoss_UpdateAdjusters" & vbCrLf
'    sMess = sMess & Me.Name & vbCrLf
'    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
'    txtErrors.Text = txtErrors.Text & sMess
'    txtErrors.Refresh
'    ErrorLog sMess
'    Err.Clear
'    Resume Next
'End Sub

'Private Sub UpdateAdjusters_ACCESS2000(colAdjusters As Collection)
'    On Error GoTo EH
'    Dim sSQL As String
'    Dim vADJ As Variant
'    Dim MyADJ As udtAdjuster
'    Dim sMess As String
'
'    SetDB GetSetting("V2WebControl", "Dir", "V2WebControl_SERVER_SHARE", vbNullString) & "\WebControl2k.mdb"
'
'    'Update the adjuster table with the Event from clsLossReports
'    For Each vADJ In colAdjusters
'        MyADJ = vADJ
'        Select Case MyADJ.IsDirty
'            Case IsDirty.AddMe
'                sSQL = "DELETE * FROM Adjuster "
'                sSQL = sSQL & "WHERE Adjuster.Fact = '" & CleanString(MyADJ.Key) & "' "
'                gDB.Execute sSQL
'
'                sSQL = "INSERT INTO Adjuster (CRID, "
'                sSQL = sSQL & "Fact, "
'                sSQL = sSQL & "ADJFName, "
'                sSQL = sSQL & "ADJLName, "
'                sSQL = sSQL & "EMailAddress, "
'                sSQL = sSQL & "ADJSSN, "
'                sSQL = sSQL & "ADJPassword, "
'                sSQL = sSQL & "ADJContactPhone, "
'                sSQL = sSQL & "ADJTeamLeader, "
'                sSQL = sSQL & "ADJUpdateEmail, "
'                sSQL = sSQL & "ADJDateLastUpdated, "
'                sSQL = sSQL & "ADJLicDaysLeft, "
'                sSQL = sSQL & "ADJAppVSInfo )"
'
'                sSQL = sSQL & "VALUES (" & IIf(MyADJ.CRID = vbNullString, "Null", Val(MyADJ.CRID)) & ", "
'                sSQL = sSQL & "'" & IIf(MyADJ.FACT = vbNullString, "Null", CleanString(MyADJ.FACT)) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJFName) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJLName) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.EmailAddress) & "', "
'                sSQL = sSQL & Val(MyADJ.ADJSSN) & ", "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJPassword) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJContactPhone) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJTeamLeader) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJUpdateEmail) & "', "
'                sSQL = sSQL & "#" & MyADJ.ADJDateLastUpdated & "#, "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJLicDaysLeft) & "', "
'                sSQL = sSQL & "'" & Replace(MyADJ.ADJAPPVSInfo, "'", "''") & "' ) "
'
'                gDB.Execute sSQL, dbFailOnError
'                If gDB.RecordsAffected = 0 Then
'                    Err.Raise -999, , "No Records Affected."
'                End If
'
'            Case IsDirty.DeleteMe
'                sSQL = "DELETE * FROM Adjuster "
'                sSQL = sSQL & "WHERE Adjuster.Fact = '" & CleanString(MyADJ.Key) & "' "
'                gDB.Execute sSQL, dbFailOnError
'                If gDB.RecordsAffected = 0 Then
'                    Err.Raise -999, , "No Records Affected."
'                End If
'        End Select
'    Next
'    SendAdjusterTable colAdjusters
'    Exit Sub
'EH:
'    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
'    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
'    sMess = sMess & Err.Description & vbCrLf & vbCrLf
'    sMess = sMess & "DB: " & gDB.Name & vbCrLf
'    sMess = sMess & "Adjuster: """ & MyADJ.FACT & "/" & MyADJ.CRID & vbCrLf
'    sMess = sMess & "Private Sub UpdateAdjusters_ACCESS2000" & vbCrLf
'    sMess = sMess & Me.Name & vbCrLf
'    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
'    txtErrors.Text = txtErrors.Text & sMess
'    txtErrors.Refresh
'    ErrorLog sMess
'    Err.Clear
'    Resume Next
'End Sub

'Private Sub UpdateAdjusters_SQLServer(colAdjusters As Collection)
'    On Error GoTo EH
'    Dim sSQL As String
'    Dim vADJ As Variant
'    Dim MyADJ As udtAdjuster
'    Dim sMess As String
'    Dim sProdDSN As String
'    Dim lRecordsAffected As Long
'
'    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
'    OpenConnection sProdDSN
'
'    'Update the adjuster table with the Event from clsLossReports
'    For Each vADJ In colAdjusters
'        MyADJ = vADJ
'        Select Case MyADJ.IsDirty
'            Case IsDirty.AddMe
'                sSQL = "DELETE FROM Adjuster "
'                sSQL = sSQL & "WHERE Adjuster.Fact = '" & CleanString(MyADJ.Key) & "' "
'                gConn.Execute sSQL
'
'                sSQL = "INSERT INTO Adjuster (CRID, "
'                sSQL = sSQL & "Fact, "
'                sSQL = sSQL & "ADJFName, "
'                sSQL = sSQL & "ADJLName, "
'                sSQL = sSQL & "EMailAddress, "
'                sSQL = sSQL & "ADJSSN, "
'                sSQL = sSQL & "ADJPassword, "
'                sSQL = sSQL & "ADJContactPhone, "
'                sSQL = sSQL & "ADJTeamLeader, "
'                sSQL = sSQL & "ADJUpdateEmail, "
'                sSQL = sSQL & "ADJDateLastUpdated, "
'                sSQL = sSQL & "ADJLicDaysLeft, "
'                sSQL = sSQL & "ADJAppVSInfo )"
'
'                sSQL = sSQL & "VALUES (" & IIf(MyADJ.CRID = vbNullString, "Null", Val(MyADJ.CRID)) & ", "
'                sSQL = sSQL & "'" & IIf(MyADJ.FACT = vbNullString, "Null", CleanString(MyADJ.FACT)) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJFName) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJLName) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.EmailAddress) & "', "
'                sSQL = sSQL & Val(MyADJ.ADJSSN) & ", "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJPassword) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJContactPhone) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJTeamLeader) & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJUpdateEmail) & "', "
'                sSQL = sSQL & "'" & MyADJ.ADJDateLastUpdated & "', "
'                sSQL = sSQL & "'" & CleanString(MyADJ.ADJLicDaysLeft) & "', "
'                sSQL = sSQL & "'" & Replace(MyADJ.ADJAPPVSInfo, "'", "''") & "' ) "
'
'                gConn.Execute sSQL, lRecordsAffected
'                If lRecordsAffected = 0 Then
'                    Err.Raise -999, , "No Records Affected."
'                End If
'
'            Case IsDirty.DeleteMe
'                sSQL = "DELETE FROM Adjuster "
'                sSQL = sSQL & "WHERE Adjuster.Fact = '" & CleanString(MyADJ.Key) & "' "
'                gConn.Execute sSQL, lRecordsAffected
'                If lRecordsAffected = 0 Then
'                    Err.Raise -999, , "No Records Affected."
'                End If
'        End Select
'    Next
'    SendAdjusterTable colAdjusters
'    Exit Sub
'EH:
'    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
'    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
'    sMess = sMess & Err.Description & vbCrLf & vbCrLf
'    sMess = sMess & "DB: " & gConn.DefaultDatabase & vbCrLf
'    sMess = sMess & "Adjuster: """ & MyADJ.FACT & "/" & MyADJ.CRID & vbCrLf
'    sMess = sMess & "Private Sub UpdateAdjusters_SQLServer" & vbCrLf
'    sMess = sMess & Me.Name & vbCrLf
'    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
'    txtErrors.Text = txtErrors.Text & sMess
'    txtErrors.Refresh
'    ErrorLog sMess
'    Err.Clear
'    Resume Next
'End Sub

Private Sub txtADJFTPPath_Change()
    CleanFileName txtADJFTPPath, True
    EnableProcess txtADJFTPPath
End Sub

Private Sub txtADJFTPPath_GotFocus()
    SelText txtADJFTPPath
End Sub

Private Sub txtErrors_GotFocus()
    SelText txtErrors
End Sub

Private Sub txtMess_GotFocus()
    SelText txtMess
End Sub

Private Sub moLoss_UpdateDB(ByVal oLossReport As V2ECKeyBoard.clsCarLR)
    On Error GoTo EH
    Dim sMess As String
    Dim lErrNum As Long
    Dim sErrMess As String
    
    UpdateDB_SQLServer oLossReport

    Exit Sub
EH:
    lErrNum = Err.Number
    sErrMess = Err.Description
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & vbCrLf
    sMess = sMess & sErrMess & vbCrLf & vbCrLf
    sMess = sMess & "Private Sub moLoss_UpdateDB" & vbCrLf
    sMess = sMess & Me.Name & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    txtErrors.Text = txtErrors.Text & sMess
    txtErrors.Refresh
    ErrorLog sMess
    If StrComp(oLossReport.ClassName, "V2ECCarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        oLossReport.Status = lErrNum & "|" & sErrMess
        Err.Clear
    Else
        Err.Clear
    End If
    Resume Next
End Sub

Private Sub UpdateDB_ACCESS2000(ByVal oLossReport As Object)
    On Error GoTo EH
    Dim sSQL As String
    Dim sMess As String
    Dim sSQLError As String
    Dim lVersion As Long
    Dim sAppEXEName As String
    
    lVersion = CLng(App.Major & App.Minor & App.Revision)
    sAppEXEName = App.EXEName
    
    SetDB GetSetting("V2WebControl", "Dir", "V2WebControl_SERVER_SHARE", vbNullString) & "\WebControl2k.mdb"
    sSQL = oLossReport.GetLRSQL(sAppEXEName, lVersion, sMess)
    
    gDB.Execute sSQL, dbFailOnError
    
    If gDB.RecordsAffected = 0 Then
        Err.Raise -999, , "No Records Affected."
    End If
    
    Exit Sub
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & " " & Now & vbCrLf
    sMess = sMess & Err.Description & vbCrLf
    sMess = sMess & "DB: " & gDB.Name & vbCrLf
    sMess = sMess & "Report Format: " & oLossReport.ClassName & vbCrLf
    sMess = sMess & "Could not Add """ & oLossReport.PrnPath & """ into Database." & vbCrLf
    sMess = sMess & "This Loss Report is assigned to: " & oLossReport.Adjuster & "_" & oLossReport.CRID & vbCrLf
    sMess = sMess & "Private Sub UpdateDB_ACCESS2000" & vbCrLf
    sMess = sMess & Me.Name & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    txtErrors.Text = txtErrors.Text & sMess
    txtErrors.Refresh
    ErrorLog sMess
    Err.Clear
    Resume Next
End Sub

Private Sub UpdateDB_SQLServer(ByVal oLossReport As V2ECKeyBoard.clsCarLR)
    On Error GoTo EH
    Dim sSQL As String
    Dim sMess As String
    Dim sErrorMess As String
    Dim sSQLError As String
    Dim lVersion As Long
    Dim sAppEXEName As String
    Dim sProdDSN As String
    Dim lRecordsAffected As Long
    Dim lRet As Long
    
    
    Dim lErrNum As Long
    Dim sErrMess As String
    
    lVersion = CLng(App.Major & App.Minor & App.Revision)
    sAppEXEName = App.EXEName
    
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDSN
    
    'If dealing with Farmers XML need to check
    'for already existing Claim...
    'Pass In sMess to update Policy limits only if this is an update
    'to a claim that has already been started.
    If StrComp(oLossReport.ClassName, "V2ECCarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        'First Check the Collection of Assignments to see if it Exists.
        'This is a lot faster then doing a DB query
        If oLossReport.Status = "AddToExistingLoss" Then
            sMess = "*ADJUSTERUSERNAME*" & "_" & oLossReport.ACID & "_" & "*IBNUMBER*" & "_" & oLossReport.CLIENTNUM
        End If
    End If
    
GET_SQL:
    sSQL = oLossReport.GetLRSQL(sAppEXEName, lVersion, sMess)
    If sSQL = vbNullString Then
        If StrComp(oLossReport.Status, "AddToExistingLoss", vbTextCompare) = 0 Then
            Exit Sub ' bail here
        End If
    End If
    
    gConn.Execute sSQL, lRecordsAffected
    If lRecordsAffected = 0 Then
        Err.Raise -999, , "No Records Affected."
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrMess = Err.Description
    'Check for Farmers XML Unit Adding
    If StrComp(oLossReport.ClassName, "V2ECCarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        'If Database returns this error means need to Add this Unit to existing
        'Assignment ...
        If InStr(1, sErrMess, "V2ECCarFarmers.clsLossXML01|AddToExistingLoss|") > 0 Then
            Err.Clear
            sMess = sErrMess
            GoTo GET_SQL
            Exit Sub
        End If
    End If
    
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & lErrNum & " " & Now & vbCrLf
    sMess = sMess & sErrMess & vbCrLf
    sMess = sMess & "DB: " & gConn.DefaultDatabase & vbCrLf
    sMess = sMess & "Report Format: " & oLossReport.ClassName & vbCrLf
    sMess = sMess & "Could not Add """ & oLossReport.PrnKey & """ into Database." & vbCrLf
    sMess = sMess & "This Loss Report is assigned to: " & oLossReport.ACID & vbCrLf
    sMess = sMess & "Private Sub UpdateDB_SQLServer" & vbCrLf
    sMess = sMess & Me.Name & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    txtErrors.Text = txtErrors.Text & sMess
    txtErrors.Refresh
    ErrorLog sMess
    If StrComp(oLossReport.ClassName, "V2ECCarFarmers.clsLossXML01", vbTextCompare) = 0 Then
        oLossReport.Status = lErrNum & "|" & sErrMess
        Err.Clear
        Exit Sub
    Else
        Err.Clear
    End If
    sErrorMess = "Click ""Yes"" to not prompt for subsequent Errors and continue processing." & vbCrLf & vbCrLf
    sErrorMess = sErrorMess & "Click ""No"" to continue processing as well as continue to prompt subsequent Errors." & vbCrLf & vbCrLf
    sErrorMess = sErrorMess & "Click ""Cancel"" to STOP processing all together."
    If Not moLoss.IgnoreProcessRawDataErrors Then
        lRet = MsgBox(sMess & vbCrLf & vbCrLf & vbCrLf & sErrorMess, vbYesNoCancel, "Error")
        If lRet = vbYes Then
            moLoss.IgnoreProcessRawDataErrors = True
        ElseIf lRet = vbNo Then
             moLoss.IgnoreProcessRawDataErrors = False
        ElseIf lRet = vbCancel Then
            moLoss.IgnoreProcessRawDataErrors = True
            oLossReport.AbortProcessRawData = True
        End If
    End If
    Resume Next
End Sub

Private Function SendAdjusterTable(Optional poAdjCol As Collection) As Boolean
    On Error GoTo EH
    Dim sDSN As String
    Dim sMess As String
    
    sDSN = GetSetting("V2WebControl", "DSN", "NAME", "ACCESS_2000")
    
    If sDSN = "ACCESS_2000" Then
        SendAdjusterTable = SendAdjusterTable_ACCESS2000(poAdjCol)
    Else
        'CODE FOR SQL SERVER
        SendAdjusterTable = SendAdjusterTable_SQLServer(poAdjCol)
    End If
    
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf & vbCrLf
    sMess = sMess & "Private Function SendAdjusterTable" & vbCrLf
    sMess = sMess & Me.Name & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    txtErrors.Text = txtErrors.Text & sMess
    txtErrors.Refresh
    ErrorLog sMess
    Err.Clear
    Resume Next
End Function

Private Function SendAdjusterTable_ACCESS2000(Optional poAdjCol As Collection) As Boolean
    On Error GoTo EH
    Dim RS As DAO.Recordset
    Dim sSQL As String
    Dim MyADJ As V2ECKeyBoard.udtAdjuster
    Dim sMess As String
    
    If poAdjCol Is Nothing Then
        moLoss.ResetAdjusterCol
    Else
        Set poAdjCol = Nothing
        Set poAdjCol = New Collection
    End If
    
    'Need to send in Adjuster Table To be used by
    'CCMS Object (and future objects ?) when resolving Adjuster CRID
    'If it can't be resolved then the adjuster is missing and needs
    'to be added.  The adding of missing adjusters will take place
    'in the moLoss_AddMissingAdjusters event raised by clsLossReports object.
    sSQL = "SELECT A.CRID, "
    sSQL = sSQL & "A.Fact, "
    sSQL = sSQL & "A.ADJFName, "
    sSQL = sSQL & "A.ADJLName, "
    sSQL = sSQL & "A.EMailAddress, "
    sSQL = sSQL & "A.ADJSSN, "
    sSQL = sSQL & "A.ADJPassword, "
    sSQL = sSQL & "A.ADJContactPhone, "
    sSQL = sSQL & "A.ADJTeamLeader, "
    sSQL = sSQL & "A.ADJUpdateEmail, "
    sSQL = sSQL & "A.ADJDateLastUpdated, "
    sSQL = sSQL & "A.ADJLicDaysLeft, "
    sSQL = sSQL & "A.ADJAppVSInfo "
    sSQL = sSQL & "FROM Adjuster A "
    
    SetDB GetSetting("V2WebControl", "Dir", "V2WebControl_SERVER_SHARE", vbNullString) & "\WebControl2k.mdb"
    
    If gDB Is Nothing Then
        Exit Function
    End If
    Set RS = gDB.OpenRecordset(sSQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
    End If
    Do Until RS.EOF
        MyADJ.CRID = IIf(IsNull(RS!CRID), vbNullString, RS!CRID)
        MyADJ.FACT = IIf(IsNull(RS!FACT), vbNullString, RS!FACT)
        MyADJ.ADJFName = IIf(IsNull(RS!ADJFName), vbNullString, RS!ADJFName)
        MyADJ.ADJLName = IIf(IsNull(RS!ADJLName), vbNullString, RS!ADJLName)
        MyADJ.EmailAddress = IIf(IsNull(RS!EmailAddress), vbNullString, RS!EmailAddress)
        MyADJ.ADJSSN = IIf(IsNull(RS!ADJSSN), vbNullString, RS!ADJSSN)
        MyADJ.ADJPassword = IIf(IsNull(RS!ADJPassword), vbNullString, RS!ADJPassword)
        MyADJ.ADJContactPhone = IIf(IsNull(RS!ADJContactPhone), vbNullString, RS!ADJContactPhone)
        MyADJ.ADJTeamLeader = IIf(IsNull(RS!ADJTeamLeader), vbNullString, RS!ADJTeamLeader)
        MyADJ.ADJUpdateEmail = IIf(IsNull(RS!ADJUpdateEmail), vbNullString, RS!ADJUpdateEmail)
        MyADJ.ADJDateLastUpdated = IIf(IsNull(RS!ADJDateLastUpdated), "12:00:00 AM", RS!ADJDateLastUpdated)
        MyADJ.ADJLicDaysLeft = IIf(IsNull(RS!ADJLicDaysLeft), vbNullString, RS!ADJLicDaysLeft)
        MyADJ.ADJAPPVSInfo = IIf(IsNull(RS!ADJAPPVSInfo), vbNullString, RS!ADJAPPVSInfo)
        MyADJ.IsDirty = IsDirty.NoChange
        MyADJ.Key = IIf(IsNull(RS!FACT), vbNullString, RS!FACT)
        
        If MyADJ.CRID = vbNullString Or MyADJ.FACT = vbNullString Then
            Err.Raise -999, , "Invalid CRID or FACT"
            
        Else
            If poAdjCol Is Nothing Then
                moLoss.AddAdjuster MyADJ
            Else
                poAdjCol.Add MyADJ, MyADJ.Key
            End If
            
        End If
        
        RS.MoveNext
    Loop
    Set RS = Nothing
    SendAdjusterTable_ACCESS2000 = True
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf & vbCrLf
    sMess = sMess & "Adjuster: """ & MyADJ.FACT & "/" & MyADJ.CRID & vbCrLf
    sMess = sMess & "Private Function SendAdjusterTable_ACCESS2000" & vbCrLf
    sMess = sMess & Me.Name & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    txtErrors.Text = txtErrors.Text & sMess
    txtErrors.Refresh
    ErrorLog sMess
    Err.Clear
End Function

Private Function SendAdjusterTable_SQLServer(Optional poAdjCol As Collection) As Boolean
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim MyADJ As V2ECKeyBoard.udtAdjuster
    Dim sMess As String
    Dim sProdDSN As String
    Dim lRecordsAffected As Long
    
    If poAdjCol Is Nothing Then
        moLoss.ResetAdjusterCol
    Else
        Set poAdjCol = Nothing
        Set poAdjCol = New Collection
    End If
    
    'Need to send in Adjuster Table To be used by
    'CCMS Object (and future objects ?) when resolving Adjuster CRID
    'If it can't be resolved then the adjuster is missing and needs
    'to be added.  The adding of missing adjusters will take place
    'in the moLoss_AddMissingAdjusters event raised by clsLossReports object.
    sSQL = "SELECT A.CRID, "
    sSQL = sSQL & "A.Fact, "
    sSQL = sSQL & "A.ADJFName, "
    sSQL = sSQL & "A.ADJLName, "
    sSQL = sSQL & "A.EMailAddress, "
    sSQL = sSQL & "A.ADJSSN, "
    sSQL = sSQL & "A.ADJPassword, "
    sSQL = sSQL & "A.ADJContactPhone, "
    sSQL = sSQL & "A.ADJTeamLeader, "
    sSQL = sSQL & "A.ADJUpdateEmail, "
    sSQL = sSQL & "A.ADJDateLastUpdated, "
    sSQL = sSQL & "A.ADJLicDaysLeft, "
    sSQL = sSQL & "A.ADJAppVSInfo "
    sSQL = sSQL & "FROM Adjuster A "
    
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    OpenConnection sProdDSN
    Set RS = New ADODB.Recordset
    
    If gConn Is Nothing Then
        Exit Function
    End If
    RS.Open sSQL, gConn, adOpenStatic, adLockReadOnly
    
    If Not RS.EOF Then
        RS.MoveFirst
    End If
    Do While Not RS.EOF
        MyADJ.CRID = IIf(IsNull(RS!CRID), vbNullString, RS!CRID)
        MyADJ.FACT = IIf(IsNull(RS!FACT), vbNullString, RS!FACT)
        MyADJ.ADJFName = IIf(IsNull(RS!ADJFName), vbNullString, RS!ADJFName)
        MyADJ.ADJLName = IIf(IsNull(RS!ADJLName), vbNullString, RS!ADJLName)
        MyADJ.EmailAddress = IIf(IsNull(RS!EmailAddress), vbNullString, RS!EmailAddress)
        MyADJ.ADJSSN = IIf(IsNull(RS!ADJSSN), vbNullString, RS!ADJSSN)
        MyADJ.ADJPassword = IIf(IsNull(RS!ADJPassword), vbNullString, RS!ADJPassword)
        MyADJ.ADJContactPhone = IIf(IsNull(RS!ADJContactPhone), vbNullString, RS!ADJContactPhone)
        MyADJ.ADJTeamLeader = IIf(IsNull(RS!ADJTeamLeader), vbNullString, RS!ADJTeamLeader)
        MyADJ.ADJUpdateEmail = IIf(IsNull(RS!ADJUpdateEmail), vbNullString, RS!ADJUpdateEmail)
        MyADJ.ADJDateLastUpdated = IIf(IsNull(RS!ADJDateLastUpdated), "12:00:00 AM", RS!ADJDateLastUpdated)
        MyADJ.ADJLicDaysLeft = IIf(IsNull(RS!ADJLicDaysLeft), vbNullString, RS!ADJLicDaysLeft)
        MyADJ.ADJAPPVSInfo = IIf(IsNull(RS!ADJAPPVSInfo), vbNullString, RS!ADJAPPVSInfo)
        MyADJ.IsDirty = IsDirty.NoChange
        MyADJ.Key = IIf(IsNull(RS!FACT), vbNullString, RS!FACT)
        
        
        If MyADJ.CRID = vbNullString Or MyADJ.FACT = vbNullString Then
            Err.Raise -999, , "Invalid CRID or FACT"
            
        Else
            If poAdjCol Is Nothing Then
                moLoss.AddAdjuster MyADJ
            Else
                poAdjCol.Add MyADJ, MyADJ.Key
            End If
            
        End If
        
        RS.MoveNext
    Loop
    Set RS = Nothing
    SendAdjusterTable_SQLServer = True
    Exit Function
EH:
    sMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sMess = sMess & "ERROR # " & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf & vbCrLf
    sMess = sMess & "Adjuster: """ & MyADJ.FACT & "/" & MyADJ.CRID & vbCrLf
    sMess = sMess & "Private Function SendAdjusterTable_SQLServer" & vbCrLf
    sMess = sMess & Me.Name & vbCrLf
    sMess = sMess & "<<<<<<<<<< END ERROR MESSAGE >>>>>>>>>>" & vbCrLf & vbCrLf
    txtErrors.Text = txtErrors.Text & sMess
    txtErrors.Refresh
    ErrorLog sMess
    Err.Clear
End Function

Public Sub LoadCompanies()
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sUserName As String
    Dim sUID As String
    Dim sCompany As String
    Dim sCarPrefix As String
    Dim sProdDSN As String
    Dim sErrorMess As String
    Dim sImagePath As String
    
    cboCompany.Clear
    imgListCompany.ListImages.Clear
    imgCompany.Picture = Nothing
    cboCar.Clear
    imgListClientCompany.ListImages.Clear
    imgClientCompany.Picture = Nothing
    
    sUserName = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    If sUserName = vbNullString Then
        MsgBox "Must Set up Database Login under Settings!", vbExclamation + vbOKOnly, "User Name Not Found"
        Exit Sub
    End If
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    CloseConnection
    If Not OpenConnection(sProdDSN, , sErrorMess) Then
        MsgBox sErrorMess, vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    Set RS = New ADODB.Recordset
    
    sSQL = "z_spsGetCompanyUsersInfo 1,0,null,null,null,null,null,null,null,null,null,'and username=''" & sUserName & "'' ' "
    
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF Then
        RS.MoveFirst
        sUID = RS!USERSID
    Else
        GoTo CLEAN_UP
    End If
    
    RS.Close
    'Need to Get Companies, not client Companies
    'that the DB User Name has access to...
    sSQL = "z_spsGetCompanyInfo "
    sSQL = sSQL & "1, "
    sSQL = sSQL & sUID & ", "
    sSQL = sSQL & "null, "
    sSQL = sSQL & "null, "
    sSQL = sSQL & "null, "
    sSQL = sSQL & "'And IsClientOf Is Null "
    sSQL = sSQL & "And CompanyID IN( "
    sSQL = sSQL & "SELECT CompanyID "
    sSQL = sSQL & "FROM CompanyUsers "
    sSQL = sSQL & "WHERE UsersID = " & sUID & " "
    sSQL = sSQL & "AND Active = 1 "
    sSQL = sSQL & ") ' "
    
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            sCompany = RS!DBName
            sCarPrefix = RS!CarrierPrefix
            cboCompany.AddItem sCompany & String(200, Chr(32)) & "|" & sCarPrefix
            'Set the CompanyID to ItemData for this Added Item
            cboCompany.ItemData(cboCompany.NewIndex) = RS!CompanyID
            sImagePath = GetSetting("V2WebControl", "Dir", "WebSitePath", vbNullString)
            sImagePath = sImagePath & "\Images"
            If goUtil.utFileExists(sImagePath, True) Then
                sImagePath = sImagePath & "\" & RS!LogoImageName
                If goUtil.utFileExists(sImagePath) Then
                    imgListCompany.ListImages.Add , """" & CStr(RS!CompanyID) & """", LoadPicture(sImagePath)
                End If
            End If
            RS.MoveNext
        Loop
    End If
    
CLEAN_UP:
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utShowError App.EXEName, Err, "Private Sub LoadCompanies", Me
End Sub

Private Sub LoadCarriers()
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sUserName As String
    Dim sUID As String
    Dim sClientCompany As String
    Dim sProdDSN As String
    Dim sErrorMess As String
    Dim sImagePath As String
    
    cboCar.Clear
    imgListClientCompany.ListImages.Clear
    imgClientCompany.Picture = Nothing
    
    If FileExists("C:\WINDOWS\SYSTEM\ECS\DLL", True) Then
        msBaseCarDLLPath = "C:\WINDOWS\SYSTEM\ECS\DLL"
    ElseIf FileExists("C:\WINDOWS\SYSTEM32\ECS\DLL", True) Then
        msBaseCarDLLPath = "C:\WINDOWS\SYSTEM32\ECS\DLL"
    ElseIf FileExists("C:\WINNT\SYSTEM32\ECS\DLL", True) Then
        msBaseCarDLLPath = "C:\WINNT\SYSTEM32\ECS\DLL"
    End If

    sUserName = goUtil.utGetECSCryptSetting("V2WebControl", "DBConn", "USERID")
    If sUserName = vbNullString Then
        MsgBox "Must Set up Database Login under Settings!", vbExclamation + vbOKOnly, "User Name Not Found"
        Exit Sub
    End If
    sProdDSN = GetSetting("V2WebControl", "DSN", "NAME", vbNullString)
    CloseConnection
    If Not OpenConnection(sProdDSN, , sErrorMess) Then
        MsgBox sErrorMess, vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    Set RS = New ADODB.Recordset
    
    sSQL = "z_spsGetCompanyUsersInfo 1,0,null,null,null,null,null,null,null,null,null,'and username=''" & sUserName & "'' ' "
    
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF Then
        RS.MoveFirst
        sUID = RS!USERSID
    Else
        GoTo CLEAN_UP
    End If
    
    RS.Close
    'Need to Get Companies, not client Companies
    'that the DB User Name has access to...
    sSQL = "z_spsGetCompanyInfo "
    sSQL = sSQL & "1, "
    sSQL = sSQL & sUID & ", "
    sSQL = sSQL & "null, "
    sSQL = sSQL & "null, "
    sSQL = sSQL & "null, "
    sSQL = sSQL & "'And IsClientOf = " & cboCompany.ItemData(cboCompany.ListIndex) & " "
    sSQL = sSQL & "And CompanyID IN( "
    sSQL = sSQL & "SELECT CompanyID "
    sSQL = sSQL & "FROM CompanyUsers "
    sSQL = sSQL & "WHERE UsersID = " & sUID & " "
    sSQL = sSQL & "AND Active = 1 "
    sSQL = sSQL & ") ' "
    
    RS.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            sClientCompany = RS!DBName
            sClientCompany = Replace(sClientCompany, ".mdb", vbNullString, , , vbTextCompare)
            cboCar.AddItem sClientCompany
            'Set the CompanyID to ItemData for this Added Item
            cboCar.ItemData(cboCar.NewIndex) = RS!CompanyID
            sImagePath = GetSetting("V2WebControl", "Dir", "WebSitePath", vbNullString)
            sImagePath = sImagePath & "\Images"
            If goUtil.utFileExists(sImagePath, True) Then
                sImagePath = sImagePath & "\" & RS!LogoImageName
                If goUtil.utFileExists(sImagePath) Then
                    imgListClientCompany.ListImages.Add , """" & CStr(RS!CompanyID) & """", LoadPicture(sImagePath)
                End If
            End If
            RS.MoveNext
        Loop
    End If
    
CLEAN_UP:
    Set RS = Nothing
    
    Exit Sub
EH:
    ShowError Err, "Private Sub LoadCarriers", Me
End Sub

Private Sub LoadFormats()
    On Error GoTo EH
    Dim oCar As V2ECKeyBoard.clsCarLists
    Dim sCar As String
    Dim vFormat As Variant
    Dim sFormat As String
    
    cboLossFormat.Clear
    
    If cboCar.Text = "V2ECKeyBoard" Then
        cboLossFormat.AddItem "clsLossUnknown"
        Exit Sub
    End If
    sCar = goUtil.gsCarPrefix & cboCar.Text
    If Not FileExists(msBaseCarDLLPath & "\" & sCar & ".dll") Then
        Err.Raise 429, , "ActiveX component can't create object " & msBaseCarDLLPath & "\" & sCar & vbCrLf & "Carrier Object Error" & vbCrLf & "Or no profile exists."
    End If
    
    Set oCar = CreateObject(sCar & ".clsLists")
        
    
    For Each vFormat In oCar.FormatList
        sFormat = vFormat
        cboLossFormat.AddItem sFormat
    Next

    Set oCar = Nothing
    Exit Sub
EH:
    ShowError Err, "Private Sub LoadFormats", Me
End Sub

Public Sub LoadDataPaths()
    On Error GoTo EH
    
    EnableAutoUpdateXML
   
    txtRawDataPath.Text = GetSetting(App.EXEName, "Dir", goUtil.gsCarPrefix & "_" & cboCar.Text & "_" & cboLossFormat.Text & "_txtRawDataPath", vbNullString)
    txtADJFTPPath.Text = GetSetting(App.EXEName, "Dir", goUtil.gsCarPrefix & "_" & cboCar.Text & "_" & cboLossFormat.Text & "_txtADJFTPPath", vbNullString)
    cmdRawDataPath.Enabled = True
    cmdADJFTPPath.Enabled = True
    Exit Sub
EH:
    ShowError Err, "Private Sub LoadRawDataPath", Me
End Sub

Public Sub EnableAutoUpdateXML()
    On Error GoTo EH
    
    'If the Selected Loss Report Format is XML Type...
    'Need to Enable and show the Auto Update XML Check Box
    If InStr(1, cboLossFormat.Text, "clsLossXML", vbTextCompare) > 0 Then
        mbEditAutoUpdateXMLTrans = True
        chkAutoUpdateXMLTrans.Value = GetSetting(App.EXEName, "Dir", goUtil.gsCarPrefix & "_" & cboCar.Text & "_" & cboLossFormat.Text & "_chkAutoUpdateXMLTrans", vbUnchecked)
        chkAutoUpdateXMLTrans.Enabled = True
        mbEditAutoUpdateXMLTrans = False
    Else
        mbEditAutoUpdateXMLTrans = True
        chkAutoUpdateXMLTrans.Value = vbUnchecked
        chkAutoUpdateXMLTrans.Enabled = False
        mbEditAutoUpdateXMLTrans = False
    End If
    
    Exit Sub
EH:
    ShowError Err, "Public Sub EnableAutoUpdateXML", Me
End Sub

Private Sub LoadUploadPaths()
    On Error GoTo EH
    Dim sFTPPaths As String
    Dim vFTPPaths As Variant
    Dim lCount As Long
    
    lstFTPPaths.Clear
    
    sFTPPaths = GetSetting(App.EXEName, "Dir", "ListADJFTP", vbNullString)
    vFTPPaths = Split(sFTPPaths, LIST_ADJFTP_DELIM)
    
    If IsArray(vFTPPaths) And sFTPPaths <> vbNullString Then
        Do Until vFTPPaths(lCount) = vbNullString
            lstFTPPaths.AddItem vFTPPaths(lCount)
            lCount = lCount + 1
        Loop
    End If
    
    
    Exit Sub
EH:
    ShowError Err, "Private Sub LoadUploadPaths", Me
End Sub

Private Sub txtRawDataPath_Change()
    CleanFileName txtRawDataPath, True
    EnableProcess txtRawDataPath
End Sub

Private Sub EnableViewLoss()
    On Error GoTo EH
    If cboCompany.Text <> vbNullString And cboCar.Text <> vbNullString Then
        cmdViewLR.Enabled = True
    Else
        cmdViewLR.Enabled = False
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub EnableViewLoss", Me
End Sub

Private Sub EnableProcess(poTextBox As Object)
    On Error GoTo EH
    Dim sListADJFTP As String
    If (Not FileExists(txtRawDataPath.Text, True)) Or (Not FileExists(txtADJFTPPath.Text, True)) Then
        cmdProcess.Enabled = False
        cmdViewLR.Enabled = False
    Else
        cmdProcess.Enabled = True
        cmdViewLR.Enabled = True
    End If
    
    If FileExists(poTextBox.Text, True) Then
        SaveSetting App.EXEName, "Dir", goUtil.gsCarPrefix & "_" & cboCar.Text & "_" & cboLossFormat.Text & "_" & poTextBox.Name, poTextBox.Text
        'Need to update the List of ADJ FTP sites, Because
        'Auto Import will use this to import the diff upload directories
        'For various Carriers
        If InStr(1, poTextBox.Name, "ADJFTP", vbTextCompare) > 0 Then
            sListADJFTP = GetSetting(App.EXEName, "Dir", "ListADJFTP", vbNullString)
            If InStr(1, sListADJFTP, poTextBox.Text, vbTextCompare) = 0 Then
                sListADJFTP = sListADJFTP & poTextBox.Text & IIf(Right(poTextBox.Text, 1) = "\", "Upload\", "\Upload\") & LIST_ADJFTP_DELIM
                SaveSetting App.EXEName, "Dir", "ListADJFTP", sListADJFTP
                LoadUploadPaths
            End If
        End If
    End If
    
    Exit Sub
EH:
    ShowError Err, "Private Sub EnableProcess", Me
End Sub

Private Sub txtRawDataPath_GotFocus()
    SelText txtRawDataPath
End Sub


