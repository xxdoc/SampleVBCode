VERSION 5.00
Begin VB.Form frmSetUpNewCat 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup New Cat"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetUpNewCat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   25
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame framTaxInfo 
      Appearance      =   0  'Flat
      Caption         =   "Applicable Tax Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   5655
      Begin VB.TextBox txtTaxPercent 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "TaxPercent"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "Tax Percent"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.Frame framCatSite 
      Appearance      =   0  'Flat
      Caption         =   "Site Location/Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   5655
      Begin VB.ComboBox cboStates 
         Height          =   360
         Left            =   2400
         TabIndex        =   18
         Top             =   1080
         Width           =   785
      End
      Begin VB.TextBox txtZIP 
         Height          =   375
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   20
         Top             =   1440
         Width           =   1570
      End
      Begin VB.TextBox txtCity 
         Height          =   375
         Left            =   2400
         MaxLength       =   25
         TabIndex        =   16
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtAddress 
         Height          =   375
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "State"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1185
         Width           =   2775
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "Zip Code"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "City"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "Address"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   5415
      End
   End
   Begin VB.Frame framCatDetails 
      Appearance      =   0  'Flat
      Caption         =   "Catastrophe Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtBillCatCode 
         Height          =   375
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   11
         Top             =   2160
         Width           =   1570
      End
      Begin VB.ComboBox cboCar 
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox cboTOL 
         Height          =   360
         Left            =   2400
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox chkUseXactTOL 
         Caption         =   "Use Xactimate Settings"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtCATCode 
         Height          =   375
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1800
         Width           =   1570
      End
      Begin VB.TextBox txtCATName 
         Height          =   375
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "Billing Cat Code"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "Carrier Cat Code"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "Type of Loss"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1545
         Width           =   5295
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "Carrier"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   825
         Width           =   5295
      End
      Begin VB.Label lblSetUpNewCat 
         Caption         =   "CAT Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmSetUpNewCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cboStates_Change()
    On Error GoTo EH
    Dim lPos As Long
    
    goUtil.utUCText cboStates
    
    'Make sure MAX len FOr state is no greater than 2
    If Len(cboStates.Text) > 20 Then
        lPos = cboStates.SelStart
        cboStates.Text = left(cboStates.Text, 2)
        cboStates.SelStart = lPos
        cboStates.SelLength = 0
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboStates_Change"
End Sub

Private Sub cboStates_LostFocus()
    On Error GoTo EH
    cboStates.Text = Trim(cboStates.Text)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboStates_LostFocus"
End Sub

Private Sub cboTOL_Change()
    On Error GoTo EH
    Dim lPos As Long
    
    goUtil.utUCText cboTOL
    
    'Make sure MAx len for TOL is not greater than 20
    If Len(cboTOL.Text) > 20 Then
        lPos = cboTOL.SelStart
        cboTOL.Text = left(cboTOL.Text, 20)
        cboTOL.SelStart = lPos
        cboTOL.SelLength = 0
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboTOL_Change"
End Sub

Private Sub cboTOL_LostFocus()
    On Error GoTo EH
    cboTOL.Text = Trim(cboTOL.Text)
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cboTOL_LostFocus"
End Sub

Private Sub chkUseXactTOL_Click()
    On Error GoTo EH
    
    SaveSetting App.EXEName, "GENERAL", "CAT_SETUP_USE_XACT_TOL", chkUseXactTOL.Value
    
    If chkUseXactTOL.Value = vbChecked Then
        goUtil.utLoadTOL cboTOL
        goUtil.utLoadStates cboStates
    Else
        cboTOL.Clear
        cboStates.Clear
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkUseXactTOL_Click"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo EH
    
    Unload Me
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
    On Error GoTo EH
    If CreateCAT Then
        Unload Me
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdOK_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me
    
    'BGS 10.10.2002 Put trail "..." on Lables
    goUtil.utSuffixLabels lblSetUpNewCat, 50
    
    PopulateForm
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    Me.Visible = False
    gfrmECTray.ShowMe False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub PopulateForm()
    On Error GoTo EH
    Dim sUseXactTOL As String
    '1. Load the Available Carriers
    'Depending on what Carrieir objects are installed...
    goUtil.utLoadCarriers cboCar
    
    '2. Load the Type Of Loss from Xactimate if possible (Include HERE States too!!!)
    sUseXactTOL = GetSetting(App.EXEName, "GENERAL", "CAT_SETUP_USE_XACT_TOL", chkUseXactTOL.Value)
    chkUseXactTOL.Value = sUseXactTOL
        
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateForm"
End Sub

Private Sub txtAddress_Change()
    goUtil.utUCText txtAddress
End Sub

Private Sub txtAddress_GotFocus()
    goUtil.utSelText txtAddress
End Sub

Private Sub txtAddress_LostFocus()
    txtAddress.Text = Trim(txtAddress.Text)
End Sub

Private Sub txtBillCatCode_Change()
    goUtil.utUCText txtBillCatCode
End Sub

Private Sub txtBillCatCode_GotFocus()
    goUtil.utSelText txtBillCatCode
End Sub

Private Sub txtBillCatCode_LostFocus()
    txtBillCatCode.Text = Trim(txtBillCatCode.Text)
End Sub

Private Sub txtCATCode_Change()
    goUtil.utUCText txtCATCode
End Sub

Private Sub txtCATCode_GotFocus()
    goUtil.utSelText txtCATCode
End Sub

Private Sub txtCATCode_LostFocus()
    txtCATCode.Text = Trim(txtCATCode.Text)
End Sub

Private Sub txtCATName_Change()
    goUtil.utCleanFileFolderName txtCATName
End Sub

Private Sub txtCATName_GotFocus()
    goUtil.utSelText txtCATName
End Sub

Private Sub txtCATName_LostFocus()
    txtCATName.Text = Trim(txtCATName.Text)
End Sub

Private Sub txtCity_Change()
    goUtil.utUCText txtCity
End Sub

Private Sub txtCity_GotFocus()
    goUtil.utSelText txtCity
End Sub

Private Sub txtCity_LostFocus()
    txtCity.Text = Trim(txtCity.Text)
End Sub

Private Sub txtTaxPercent_Change()
    goUtil.utCleanValTextBox txtTaxPercent
End Sub

Private Sub txtTaxPercent_GotFocus()
    goUtil.utSelText txtTaxPercent
End Sub

Private Sub txtTaxPercent_LostFocus()
    txtTaxPercent.Text = Trim(txtTaxPercent.Text)
End Sub

Private Sub txtZIP_GotFocus()
    goUtil.utSelText txtZIP
End Sub

Private Function CreateCAT() As Boolean
    On Error GoTo EH
    Dim dTaxPercent As Double
    Dim colCatPrefs As Collection
    
    'Trim again just incase ALT+Ok shortcut keys to get here
    txtCATName.Text = Trim(txtCATName.Text)
    cboTOL.Text = Trim(cboTOL.Text)
    txtCATCode.Text = Trim(txtCATCode.Text)
    txtBillCatCode.Text = Trim(txtBillCatCode.Text)
    txtAddress.Text = Trim(txtAddress.Text)
    txtCity.Text = Trim(txtCity.Text)
    cboStates.Text = Trim(cboStates.Text)
    txtZIP.Text = Trim(txtZIP.Text)
    txtTaxPercent.Text = Trim(txtTaxPercent.Text)
    
    If Not IsNumeric(txtTaxPercent.Text) Then
        txtTaxPercent.Text = 0
    End If
    dTaxPercent = txtTaxPercent.Text
    
    If Trim(txtCATName.Text) = vbNullString Then
        MsgBox "Enter CAT Name!", vbExclamation, Me.Caption
        Exit Function
    ElseIf Trim(cboCar.Text) = vbNullString Then
        MsgBox "Select a Carrier!", vbExclamation, Me.Caption
        Exit Function
    ElseIf Trim(cboTOL.Text) = vbNullString Then
        MsgBox "Select Type Of Loss!", vbExclamation, Me.Caption
        Exit Function
    ElseIf Trim(txtCATCode.Text) = vbNullString Then
        MsgBox "Enter CAT Code!", vbExclamation, Me.Caption
        Exit Function
    ElseIf Trim(txtAddress.Text) = vbNullString Then
        MsgBox "Enter Site Address!", vbExclamation, Me.Caption
        Exit Function
    ElseIf Trim(txtCity.Text) = vbNullString Then
        MsgBox "Enter Site City!", vbExclamation, Me.Caption
        Exit Function
    ElseIf Trim(cboStates.Text) = vbNullString Then
        MsgBox "Select Site State!", vbExclamation, Me.Caption
        Exit Function
    ElseIf Trim(txtZIP.Text) = vbNullString Then
        MsgBox "Enter Site ZIP Code!", vbExclamation, Me.Caption
        Exit Function
    ElseIf NoTaxInRequiredState(cboStates.Text, dTaxPercent) Then
        Exit Function
    End If
    
    'Set up this collection. it will be used to create CatPRef INI file
    'Have to use ini file instead of registry because the cat can be backed up
    'to floppy or CD
    Set colCatPrefs = New Collection
    colCatPrefs.Add "CAT_NAME=" & txtCATName.Text, "CAT_NAME"
    colCatPrefs.Add "CAT_CARRIER=" & cboCar.Text, "CAT_CARRIER"
    colCatPrefs.Add "CAT_TOL=" & cboTOL.Text, "CAT_TOL"
    colCatPrefs.Add "CAT_CAT_CODE=" & txtCATCode.Text, "CAT_CAT_CODE"
    colCatPrefs.Add "CAT_BILLING_CAT_CODE=" & txtBillCatCode.Text, "CAT_BILLING_CAT_CODE"
    colCatPrefs.Add "SITE_ADDRESS=" & txtAddress.Text, "SITE_ADDRESS"
    colCatPrefs.Add "SITE_CITY=" & txtCity.Text, "SITE_CITY"
    colCatPrefs.Add "SITE_STATE=" & cboStates.Text, "SITE_STATE"
    colCatPrefs.Add "SITE_ZIP=" & txtZIP.Text, "SITE_ZIP"
    colCatPrefs.Add "TAX_TAX_PERCENT=" & txtTaxPercent.Text, "TAX_TAX_PERCENT"

    If goUtil.utValidate(Me) Then
        If goUtil.utCreateCat(App.EXEName, goUtil.gsInstallDir, colCatPrefs) Then
            gfrmECTray.LoadTree
            CreateCAT = True
        End If
    End If
    
    'CLean up
    Set colCatPrefs = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function CreateCAT"
End Function

Private Function NoTaxInRequiredState(psState As String, pdTax As Double) As Boolean
    On Error GoTo EH
    Dim bState As Boolean
    Dim sMess As String
    Dim sTemp As String
    Dim sDefaultTax As String
    
    'Issue 178 4.4.2002 Force Taxes for Texas, New Mexico, and West Virginia
    ' TX, NM, or WV
    'IF this changes in future HAve to update This Reg Setting
    sTemp = GetSetting(App.EXEName, "GENERAL", "REQ_TAX_STATES", "TX|NM|WV")
    sDefaultTax = GetSetting(App.EXEName, "GENERAL", "REQ_TAX_DEFAULT", "8.250")
    
    If InStr(1, sTemp, Trim(psState), vbTextCompare) > 0 Then
        bState = True
    End If
    
    If bState Then
        If pdTax = 0 Then
            sMess = "The state of """ & UCase(psState) & """ requires a Tax Percent!"
            sTemp = InputBox(sMess, "Tax Percent", sDefaultTax & "<--Change default if appropriate")
            sTemp = goUtil.utCleanValString(sTemp)
            If sTemp = vbNullString Then
                sTemp = 0
            End If
            sTemp = Val(sTemp)
            If Len(sTemp) = 4 Then
                sTemp = sTemp & "0"
            End If
            If sTemp = 0 Then
                MsgBox "Invalid Tax percent.", vbExclamation + vbOKOnly, Me.Caption
                NoTaxInRequiredState = True
            Else
                txtTaxPercent.Text = sTemp
            End If
        Else
            GoTo VERIFY_TAX
        End If
    Else
        '8.8.2002 also prompt for tax percent if not a required state
VERIFY_TAX:
        sMess = "Are you sure the Tax Percent is correct for this CAT?"
        sTemp = InputBox(sMess, "Tax Percent", Format(pdTax, "0.000"), Me.left, Me.top)
        sTemp = goUtil.utCleanValString(sTemp)
        If sTemp = vbNullString Then
            sTemp = pdTax
            NoTaxInRequiredState = True
        End If
        sTemp = Val(sTemp)
        If Len(sTemp) = 4 Then
            sTemp = sTemp & "0"
        End If
        If sTemp = 0 Then
            sTemp = "0.000"
        End If
        If sTemp = "0.000" Then
            If bState Then
                MsgBox "Invalid Tax percent.", vbExclamation + vbOKOnly, Me.Caption
                NoTaxInRequiredState = True
            End If
        End If
        txtTaxPercent.Text = sTemp
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function NoTaxInRequiredState"
End Function

Private Sub txtZIP_LostFocus()
    txtZIP.Text = Trim(txtZIP.Text)
End Sub
