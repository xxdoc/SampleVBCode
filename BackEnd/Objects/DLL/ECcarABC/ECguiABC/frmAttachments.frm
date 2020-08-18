VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAttachments 
   AutoRedraw      =   -1  'True
   Caption         =   "Attachments"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Tag             =   "Attachments"
   Begin VB.Frame framCommands 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7245
      TabIndex        =   18
      Top             =   5400
      Width           =   4455
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Sho&w Item List"
         Height          =   855
         Left            =   1200
         MaskColor       =   &H00000000&
         Picture         =   "frmAttachments.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   855
         Left            =   3360
         MaskColor       =   &H00000000&
         Picture         =   "frmAttachments.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2280
         MaskColor       =   &H00000000&
         Picture         =   "frmAttachments.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSpelling 
         Caption         =   "Spellin&g"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Picture         =   "frmAttachments.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framAttachments 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11580
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   495
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   495
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "&Select A&ll"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   495
         Width           =   1100
      End
      Begin VB.CommandButton cmdEditAttchments 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   495
         Width           =   975
      End
      Begin VB.CommandButton cmdlAttachPDFFilePath 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Picture         =   "frmAttachments.frx":0CE0
         TabIndex        =   2
         ToolTipText     =   "Browse"
         Top             =   495
         Width           =   975
      End
      Begin MSComctlLib.ImageList imgAttachStatus 
         Left            =   3600
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAttachments.frx":115A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAttachments.frx":12B4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame framAttachMaint 
         Caption         =   "Attachment Maintenance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   4440
         Width           =   11355
         Begin VB.CommandButton cmdPrintAttachments 
            Caption         =   "&Print Selected"
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
            Left            =   1080
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdRefreshAttachments 
            Caption         =   "&Refresh"
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
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdDelAttachments 
            Caption         =   "&Delete"
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
            Left            =   9720
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkHideDeleted 
            Caption         =   "Sho&w Deleted"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   1100
         End
      End
      Begin VB.TextBox txtSpellMe 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2400
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton cmdDown 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "frmAttachments.frx":16A0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move Selected Photo DOWN"
         Top             =   1095
         Width           =   600
      End
      Begin VB.CommandButton cmdUp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "frmAttachments.frx":1AE2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move Selected Photo UP"
         Top             =   495
         Width           =   600
      End
      Begin VB.CommandButton CmdReNumberSort 
         Caption         =   "&Save Sort Order"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Renumber and Save"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ImageList imgListPhotos 
         Left            =   4200
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstvAttachments 
         Height          =   3615
         Left            =   720
         TabIndex        =   12
         Tag             =   "Enable"
         ToolTipText     =   "Right Click for Menu"
         Top             =   840
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6376
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgAttachStatus"
         ColHdrIcons     =   "imgAttachStatus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtAttachPDFFilePath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   5400
      End
      Begin VB.CheckBox chkDelOrigPDF 
         Caption         =   "Delete Original PDF After Attach:"
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
         Left            =   795
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblAttachPDFFilePath 
         Caption         =   "PDF file attachment path:"
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
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Menu PopUpmnuAttachment 
      Caption         =   "PopUpattachment"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuEditAttachment 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeleteAttachment 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSelectAllAttachment 
         Caption         =   "&Select All"
      End
   End
End
Attribute VB_Name = "frmAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmClaim As frmClaim
Private mbUnloadMe As Boolean
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private mitmXSelected As ListItem 'Currently selected Photo Item
Private mbLoading As Boolean 'Loading Form
Private mbLoadingMe As Boolean 'Just loading data
Private msFindText As String
Private mlLastFindIndex As Long
Private mbRenumberSort As Boolean

Public Property Let itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    If Not mitmXSelected Is Nothing Then
        cmdEditAttchments.Enabled = True
    Else
        cmdEditAttchments.Enabled = False
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Let itmXSelected"
End Property
Public Property Set itmXSelected(pitmX As ListItem)
    On Error GoTo EH
    Set mitmXSelected = pitmX
    If Not mitmXSelected Is Nothing Then
        cmdEditAttchments.Enabled = True
    Else
        cmdEditAttchments.Enabled = False
    End If
    Exit Property
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Property Set itmXSelected"
End Property
Public Property Get itmXSelected() As ListItem
    Set itmXSelected = mitmXSelected
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

Public Property Let AssignmentsID(psAssignmentsID As String)
    msAssignmentsID = psAssignmentsID
End Property
Public Property Get AssignmentsID() As String
    AssignmentsID = msAssignmentsID
End Property



Public Property Let UnloadMe(pbFlag As Boolean)
    mbUnloadMe = pbFlag
End Property
Public Property Get UnloadMe() As Boolean
    UnloadMe = mbUnloadMe
End Property

Public Property Let MyfrmClaim(pofrmClaim As Object)
    Set mfrmClaim = pofrmClaim
End Property
Public Property Set MyfrmClaim(pofrmClaim As Object)
    Set mfrmClaim = pofrmClaim
End Property
Public Property Get MyfrmClaim() As Object
    Set MyfrmClaim = mfrmClaim
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub chkDelOrigPDF_Click()
    On Error GoTo EH
    If mbLoading Then
        Exit Sub
    End If
    Dim bDelOrigPDF As Boolean
    
    If chkDelOrigPDF.Value = vbChecked Then
        bDelOrigPDF = True
    Else
        bDelOrigPDF = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "DELETE_ORIG_PDF", bDelOrigPDF
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkDelOrigPDF_Click"
End Sub

Private Sub chkHideDeleted_Click()
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    
    If chkHideDeleted.Value = vbChecked Then
        chkHideDeleted.Caption = "&Hide Deleted"
        bHideDeleted = True
    Else
        chkHideDeleted.Caption = "Sho&w Deleted"
        bHideDeleted = False
    End If
    
    SaveSetting App.EXEName, "GENERAL", "HIDE_DELETED", bHideDeleted
    If Not mbLoading Then
        LoadMe
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub chkHideDeleted_Click"
End Sub


Private Sub cmdDelAttachments_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sAttachID As String
    Dim vAttachID As Variant
    Dim colAttachID As Collection
    
    
    If lstvAttachments.ListItems.Count > 0 Then
        If MsgBox("Are you sure ?", vbYesNo, "DELETE SELECTED ATTACHMENT ITEMS") = vbYes Then
            lstvAttachments.Visible = False
            Set colAttachID = New Collection
            For Each itmX In lstvAttachments.ListItems
                If itmX.Selected Then
                    colAttachID.Add itmX.SubItems(GuiAttachListView.ID - 1), itmX.SubItems(GuiAttachListView.ID - 1)
                End If
            Next
            For Each vAttachID In colAttachID
                sAttachID = vAttachID
                If DeleteAttachItem(sAttachID) Then
                    lstvAttachments.ListItems.Remove ("""" & sAttachID & """")
                End If
            Next
        End If
    End If
    
    lstvAttachments.Visible = True
    Set itmX = Nothing
    Set colAttachID = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdDelAttachments_Click"
End Sub

Private Sub cmdDown_Click()
    goUtil.utMoveListItem lstvAttachments, MoveDown
End Sub


Private Sub cmdEditAttchments_Click()
    On Error GoTo EH
    
    cmdEditAttchments.Enabled = False
    
    EditAttachment
    
    cmdEditAttchments.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdEditAttchments_Click"
End Sub

Private Sub cmdExit_Click()
    On Error GoTo EH
    
    mbUnloadMe = True
    Me.Visible = False
    mfrmClaim.Timer_UnloadForm.Enabled = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdFind_Click()
    On Error GoTo EH
    If lstvAttachments.ListItems.Count > 0 Then
        mlLastFindIndex = 0
        goUtil.utFindListItem Me, lstvAttachments, msFindText, mlLastFindIndex
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFind_Click"
End Sub

Private Sub cmdFindNext_Click()
    On Error GoTo EH
    
    If msFindText <> vbNullString And lstvAttachments.ListItems.Count > 0 Then
        goUtil.utFindListItem Me, lstvAttachments, msFindText, mlLastFindIndex
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdFindNext_Click"
End Sub

Private Sub cmdlAttachPDFFilePath_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    Dim sKBFileSize As String
    Dim sMess As String
    
    sMyFilter = sMyFilter & "PDF Document File" & " (*." & "pdf" & ")" & SD & "*." & "pdf" & SD
   
    
    sPath = goUtil.utGetPath(App.EXEName, "PDFDocumentFile", "Browse to the PDF Document File you want to attach", "", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If goUtil.utFileExists(sPath & sSelFile) Then
        If StrComp(sPath, goUtil.AttachReposPath, vbTextCompare) = 0 Then
            MsgBox "Can't use this directory for attaching files!", vbExclamation + vbOKOnly, "INVALID DIRECTORY!"
            txtAttachPDFFilePath.Text = vbNullString
            Exit Sub
        End If
        '4.13.2005 BGS  Issue319  frmAttachments Allowing too big pdf file size to be attached
        'If file size is over 3MB give error message
        
        If Not UnderMaxFileSize(sPath & sSelFile, sKBFileSize) Then
            sMess = sSelFile & " (" & sKBFileSize & ") exceeds the maximum file size! " & vbCrLf
            sMess = sMess & "The maximum file size allowed is 3MB (approximately 3,000 KB)." & vbCrLf & vbCrLf
            sMess = sMess & "Reduce the DPI (dots per inch) Pixel quality for scanned documents." & vbCrLf
            sMess = sMess & "Use Gray Scale instead of color for black and white documents." & vbCrLf
            sMess = sMess & "Separate documents into different attachments instead of lumping them all in one." & vbCrLf
            
            MsgBox sMess, vbExclamation + vbOKOnly, "File Too Big!"
            txtAttachPDFFilePath.Text = vbNullString
            Exit Sub
        End If
        
        txtAttachPDFFilePath.Text = sPath & sSelFile
        If AttachPDFFile() Then
            EditAttachment
        End If
    Else
        txtAttachPDFFilePath.Text = vbNullString
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdlAttachPDFFilePath_Click"
End Sub

Private Function UnderMaxFileSize(psFilePath As String, psKBFileSize As String) As Boolean
    On Error GoTo EH
    Dim oFI As V2ECKeyBoard.clsFileVersion
    Dim FI As V2ECKeyBoard.FILE_INFORMATION
    
    Set oFI = New V2ECKeyBoard.clsFileVersion
    
    FI = oFI.GetFileInformation(psFilePath)
    
    
    psKBFileSize = Format(CStr(FI.nFileSize / 1000), "###,###,###.###") & " KB"
    
    If FI.nFileSize > 3000000 Then
        UnderMaxFileSize = False
    Else
        UnderMaxFileSize = True
    End If
    
    Set oFI = Nothing
    
  Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UnderMaxFileSize"
End Function

Private Sub cmdPrintAttachments_Click()
    On Error GoTo EH
    
    'check to see if this claim is currenlty unloading
    'if it is don' allow this event to occur
    If mbUnloadMe Then
        Exit Sub
    End If
    
    cmdPrintAttachments.Enabled = False
    If PrintPDFAttachItemsFromLVW(lstvAttachments) Then
        If Not mbUnloadMe Then
            cmdPrintAttachments.Enabled = True
        End If
    End If

    Exit Sub
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintAttachments_Click"
End Sub

Public Function PrintPDFAttachItemsFromLVW(poListView As Object, Optional pbSelectedOnly As Boolean = True) As Boolean
    On Error GoTo EH
    Dim sCaption As String
    Dim sIBNUM As String
    Dim sCLIENTNUM As String
    Dim sPDFFileName As String
    Dim sPDFFilePath As String
    Dim itmX As ListItem
    Dim oListView As ListView
    Dim RS As ADODB.Recordset
    
    If TypeOf poListView Is ListView Then
        Set oListView = poListView
    Else
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    'GEt ref to Assignments Record
    
    Set RS = mfrmClaim.adoRSAssignments
    
    'Loop through the Selected Items and Send them to Adobe Viewer
    'Adjuster Must have Adobe Viewer Installed
    
    For Each itmX In oListView.ListItems
        If itmX.Selected Then
            sIBNUM = goUtil.IsNullIsVbNullString(RS.Fields("IBNUM"))
            sCLIENTNUM = goUtil.IsNullIsVbNullString(RS.Fields("CLIENTNUM"))
            sCaption = "Attachment - " & itmX.SubItems(GuiAttachListView.AttachName - 1) ' & Chr(160) & " "
            sCaption = sCaption & "(" & sIBNUM & "_" & sCLIENTNUM & ")"
            
            sPDFFileName = itmX.SubItems(GuiAttachListView.Attachment - 1)
            sPDFFilePath = goUtil.gsInstallDir & "\AttachRepos\" & sPDFFileName
            'Need to shell the PDF Loss Report to Adobe Reader
            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sCaption
            
        End If
    Next

    Screen.MousePointer = vbNormal
    PrintPDFAttachItemsFromLVW = True
    
    'clean up
    Set RS = Nothing
    Set itmX = Nothing
    Set oListView = Nothing
    
    Exit Function
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintPDFAttachItemsFromLVW"
End Function


Private Sub cmdPrintList_Click()
    On Error GoTo EH
    goUtil.utPrintListView App.EXEName, lstvAttachments, "Attachments"
     Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPrintList_Click"
End Sub

Private Sub cmdRefreshAttachments_Click()
    On Error GoTo EH
    cmdRefreshAttachments.Enabled = False
    Screen.MousePointer = vbHourglass
    RefreshAttachments
    Screen.MousePointer = vbDefault
    cmdRefreshAttachments.Enabled = True
Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRefreshAttachments_Click"
End Sub

Public Sub RefreshAttachments()
    LoadMe
End Sub

Private Sub CmdReNumberSort_Click()
    On Error GoTo EH
    Dim bHideDeleted As Boolean
    Dim sMess As String
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    
    'Can't save sort while able to view deleted records...
    'so give message box indicating this
    sMess = "Can't Save Sort while ""hide deleted records on all screens"" is unchecked!" & vbCrLf & vbCrLf
    sMess = sMess & "You can check this item under the Fee Schedule."
    If Not bHideDeleted And Not mbUnloadMe Then
        MsgBox sMess, vbExclamation + vbOKOnly, "Can't Save Sort!"
        Exit Sub
    End If
    
    If CmdReNumberSort.Enabled Then
        Screen.MousePointer = vbHourglass
        CmdReNumberSort.Enabled = False
        ReNumberAttachSort
        CmdReNumberSort.Enabled = True
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub CmdReNumberSort_Click"
End Sub

Public Sub ReNumberAttachSort()
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX  As MSComctlLib.ListItem
    Dim lCount As Long
    Dim udtAttach As GuiAttachItem
    
    Set oListView = lstvAttachments
    lCount = 0
    mbRenumberSort = True
    For Each itmX In oListView.ListItems
        lCount = lCount + 1
        With udtAttach
            .IDAssignments = itmX.SubItems(GuiAttachListView.IDAssignments - 1)
            .ID = itmX.SubItems(GuiAttachListView.ID - 1)
            .SortOrder = lCount
            .UpLoadMe = "True"
            .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
            .UpdateByUserID = goUtil.gsCurUsersID
        End With
        If EditAttachmentItem(udtAttach) Then
            itmX.Text = lCount
            itmX.SubItems(GuiAttachListView.SortOrder - 1) = lCount
        Else
            Exit For
        End If
    Next
    
    LoadMe
    
    Set itmX = Nothing
    Set oListView = Nothing
     mbRenumberSort = False
    Exit Sub
EH:
     mbRenumberSort = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub Public Sub ReNumberAttachSort"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH
    If SaveMe Then
        mfrmClaim.RefreshMe
        cmdSave.Enabled = False
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cmdSelAll_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    For Each itmX In lstvAttachments.ListItems
        itmX.Selected = True
    Next
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSelAll_Click"
End Sub

Private Sub cmdSpelling_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    Dim sText As String
    Dim saryText() As String
    Dim lPos As Long
    Dim udtAttach As GuiAttachItem
    
    txtSpellMe.Text = vbNullString
    
    For Each itmX In lstvAttachments.ListItems
        sText = sText & itmX.SubItems(GuiAttachListView.Description - 1) & vbCrLf
    Next
    'take off the last VBCRLF
    If sText <> vbNullString Then
        sText = left(sText, InStrRev(sText, vbCrLf, , vbBinaryCompare) - 1)
    Else
        Exit Sub
    End If
    
    'Set the Spelling text box
    txtSpellMe.Text = sText
    
    cmdSpelling.Enabled = False
    goUtil.utLoadSP
    goUtil.goSP.CheckSP txtSpellMe
    
    'Now Get the Corrected Text into Array
    sText = txtSpellMe.Text
    saryText() = Split(sText, vbCrLf, , vbBinaryCompare)
    
    'check the spelling against the List view...
    'if any changes then need to save those changes to the db
    For lPos = LBound(saryText, 1) To UBound(saryText, 1)
        sText = saryText(lPos)
        Set itmX = lstvAttachments.ListItems(lPos + 1)
        
        If StrComp(sText, itmX.SubItems(GuiAttachListView.Description - 1), vbTextCompare) <> 0 Then
            With udtAttach
                .RTAttachmentsID = itmX.SubItems(GuiAttachListView.RTAttachmentsID - 1)
                .AssignmentsID = itmX.SubItems(GuiAttachListView.AssignmentsID - 1)
                .ID = itmX.SubItems(GuiAttachListView.ID - 1)
                .IDAssignments = itmX.SubItems(GuiAttachListView.IDAssignments - 1)
                .AttachDate = itmX.Text
                .SortOrder = itmX.SubItems(GuiAttachListView.SortOrder - 1)
                'Set the text to the corrected text
                itmX.SubItems(GuiAttachListView.Description - 1) = sText
                .Description = sText
                .AttachName = itmX.SubItems(GuiAttachListView.AttachName - 1)
                .Attachment = itmX.SubItems(GuiAttachListView.Attachment - 1)
                .DownloadAttachment = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.DownloadAttachment - 1))
                .UploadAttachment = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.UploadAttachment - 1))
                .IsDeleted = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.IsDeleted - 1))
                .DownLoadMe = goUtil.GetFlagFromText(itmX.SubItems(GuiAttachListView.DownLoadMe - 1))
                .UpLoadMe = "True"
                .AdminComments = itmX.SubItems(GuiAttachListView.AdminComments - 1)
                .DateLastUpdated = Format(Now(), "MM/DD/YYYY HH:MM:SS")
                itmX.SubItems(GuiAttachListView.DateLastUpdated - 1) = .DateLastUpdated
                itmX.SubItems(GuiAttachListView.DateLastUpdatedSort - 1) = Format(.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
                .UpdateByUserID = goUtil.gsCurUsersID
            End With
            EditAttachmentItem udtAttach
        End If
    Next
    
    cmdSpelling.Enabled = True
    
    'cleanup
    Set itmX = Nothing
    txtSpellMe.Text = vbNullString
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSpelling_Click"
End Sub

Private Sub cmdUp_Click()
    goUtil.utMoveListItem lstvAttachments, MoveUp
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyF11 'maximize window / Normalize window
            If Me.WindowState = VBRUN.FormWindowStateConstants.vbMaximized Then
                Me.WindowState = VBRUN.FormWindowStateConstants.vbNormal
            Else
                Me.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
            End If
            
        Case Else
            If Not mfrmClaim Is Nothing Then
                mfrmClaim.Form_KeyDown KeyCode, Shift
            End If
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim bDelOrigPDF As Boolean
    
    mbLoading = True
    
    'Set Check box for Deleting original PDF File after Attach
    bDelOrigPDF = CBool(GetSetting(App.EXEName, "GENERAL", "DELETE_ORIG_PDF", False))
    If bDelOrigPDF Then
        chkDelOrigPDF.Value = vbChecked
    Else
        chkDelOrigPDF.Value = vbUnchecked
    End If
    
    'Set the Form Icon
    Me.Icon = mfrmClaim.optClaim(GuiClaimOptions.opt07_Attachments).Picture
    
    LoadHeaderlstvAttach
    
    Screen.MousePointer = vbHourglass
    LoadMe
    Screen.MousePointer = vbDefault
    
    CheckStatus
    
    cmdSave.Enabled = False
    mbLoading = False
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Sub LoadHeaderlstvAttach()
    On Error GoTo EH
    Dim bGridOn As Boolean
    Dim bHideDeleted As Boolean
    Dim bHideUploadFlags As Boolean
    
    bHideDeleted = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True))
    bHideUploadFlags = CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_UPLOAD_FLAGS", True))
    
    'set the columnheaders
    With lstvAttachments
        .ColumnHeaders.Add , "AttachDate", "Date"
        .ColumnHeaders.Add , "AttachDateSort", "Sort Date" ' Hidden
        .ColumnHeaders.Add , "SortOrder", "Sort Order"
        .ColumnHeaders.Add , "SortOrderSort", "Sort Order Sort" ' Hidden"
        .ColumnHeaders.Add , "AttachName", "Name"
        .ColumnHeaders.Add , "Description", "Description"
        .ColumnHeaders.Add , "Attachment", "File Name"
        .ColumnHeaders.Add , "UpLoadAttachment", "UpLoad Attachment" ' Hidden
        .ColumnHeaders.Add , "IsDeleted", "Is Deleted" ' Hidden
        .ColumnHeaders.Add , "UpLoadMe", "UpLoad Me" ' Hidden
        .ColumnHeaders.Add , "DateLastUpdated", "Date Last Updated"
        .ColumnHeaders.Add , "DateLastUpdatedSort", "Sort Date Last Updated" ' hidden
        .ColumnHeaders.Add , "AdminComments", "Admin Comments" ' Hidden
        .ColumnHeaders.Add , "ID", "ID" 'Hidden
        .ColumnHeaders.Add , "IDAssignments", "IDAssignments" 'Hidden
        .ColumnHeaders.Add , "RTAttachmentsID", "RTAttachmentsID" ' Hidden
        .ColumnHeaders.Add , "AssignmentsID", "AssignmentsID"  ' Hidden
        .ColumnHeaders.Add , "DownloadAttachment", "DownloadAttachment" 'Hidden
        .ColumnHeaders.Add , "DownLoadMe", "DownLoadMe"  ' hidden
        .ColumnHeaders.Add , "UpdateByUserID", "UpdateByUserID"  ' Hidden
        
        .Sorted = False
        .SortOrder = lvwAscending
        
        'AttachDate
        .ColumnHeaders.Item(GuiAttachListView.AttachDate).Width = 1335
        .ColumnHeaders.Item(GuiAttachListView.AttachDate).Alignment = lvwColumnLeft
        'ActDateSort
        .ColumnHeaders.Item(GuiAttachListView.AttachDateSort).Width = 0 ' Hidden
        .ColumnHeaders.Item(GuiAttachListView.AttachDateSort).Alignment = lvwColumnLeft
        'SortOrder
        .ColumnHeaders.Item(GuiAttachListView.SortOrder).Width = 1230
        .ColumnHeaders.Item(GuiAttachListView.SortOrder).Alignment = lvwColumnLeft
        'SortOrderSort
        .ColumnHeaders.Item(GuiAttachListView.SortOrderSort).Width = 0 'Hidden
        .ColumnHeaders.Item(GuiAttachListView.SortOrderSort).Alignment = lvwColumnLeft
        'SortOrder
        .ColumnHeaders.Item(GuiAttachListView.SortOrder).Width = 1230
        .ColumnHeaders.Item(GuiAttachListView.SortOrder).Alignment = lvwColumnLeft
        'AttachName
        .ColumnHeaders.Item(GuiAttachListView.AttachName).Width = 5000
        .ColumnHeaders.Item(GuiAttachListView.AttachName).Alignment = lvwColumnLeft
        'Description
        .ColumnHeaders.Item(GuiAttachListView.Description).Width = 5000
        .ColumnHeaders.Item(GuiAttachListView.Description).Alignment = lvwColumnLeft
        'Attachment
        .ColumnHeaders.Item(GuiAttachListView.Attachment).Width = 2000 '0
        .ColumnHeaders.Item(GuiAttachListView.Attachment).Alignment = lvwColumnLeft
        'UpLoadAttachment
        .ColumnHeaders.Item(GuiAttachListView.UploadAttachment).Width = 0 'Hidden 400
        .ColumnHeaders.Item(GuiAttachListView.UploadAttachment).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiAttachListView.UploadAttachment).Icon = GuiAttachStatusList.UpLoadMe
        'Is Deleted
'        If bHideDeleted Then
'            .ColumnHeaders.Item(GuiAttachListView.IsDeleted).Width = 0 'Hidden 400
'        Else
            .ColumnHeaders.Item(GuiAttachListView.IsDeleted).Width = 400
'        End If
        .ColumnHeaders.Item(GuiAttachListView.IsDeleted).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiAttachListView.IsDeleted).Icon = GuiAttachStatusList.IsDeleted
        'UpLoad Me
'        If bHideUploadFlags Then
'            .ColumnHeaders.Item(GuiAttachListView.UpLoadMe).Width = 0 'Hidden 400
'        Else
            .ColumnHeaders.Item(GuiAttachListView.UpLoadMe).Width = 400
'        End If
        .ColumnHeaders.Item(GuiAttachListView.UpLoadMe).Alignment = lvwColumnCenter
        .ColumnHeaders.Item(GuiAttachListView.UpLoadMe).Icon = GuiAttachStatusList.UpLoadMe
        'DateLastUpdated
        .ColumnHeaders.Item(GuiAttachListView.DateLastUpdated).Width = 2200
        .ColumnHeaders.Item(GuiAttachListView.DateLastUpdated).Alignment = lvwColumnLeft
        'DateLastUpdatedSort
        .ColumnHeaders.Item(GuiAttachListView.DateLastUpdatedSort).Width = 0  'Hidden Sort Column
        .ColumnHeaders.Item(GuiAttachListView.DateLastUpdatedSort).Alignment = lvwColumnLeft
        'AdminComments
        .ColumnHeaders.Item(GuiAttachListView.AdminComments).Width = 0 'Hidden 10000
        .ColumnHeaders.Item(GuiAttachListView.AdminComments).Alignment = lvwColumnLeft
        'ID
        .ColumnHeaders.Item(GuiAttachListView.ID).Width = 0  'hidden
        .ColumnHeaders.Item(GuiAttachListView.ID).Alignment = lvwColumnLeft
        'IDAssignments
        .ColumnHeaders.Item(GuiAttachListView.IDAssignments).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiAttachListView.IDAssignments).Alignment = lvwColumnLeft
        'RTAttachmentsID
        .ColumnHeaders.Item(GuiAttachListView.RTAttachmentsID).Width = 0   'Hidden
        .ColumnHeaders.Item(GuiAttachListView.RTAttachmentsID).Alignment = lvwColumnLeft
        'AssignmentsID
        .ColumnHeaders.Item(GuiAttachListView.AssignmentsID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiAttachListView.AssignmentsID).Alignment = lvwColumnLeft
        'DownloadAttachment
        .ColumnHeaders.Item(GuiAttachListView.DownloadAttachment).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiAttachListView.DownloadAttachment).Alignment = lvwColumnLeft
        'DownLoadMe
        .ColumnHeaders.Item(GuiAttachListView.DownLoadMe).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiAttachListView.DownLoadMe).Alignment = lvwColumnLeft
        'UpdateByUserID
        .ColumnHeaders.Item(GuiAttachListView.UpdateByUserID).Width = 0  'Hidden
        .ColumnHeaders.Item(GuiAttachListView.UpdateByUserID).Alignment = lvwColumnLeft
    End With
    
    bGridOn = CBool(GetSetting(App.EXEName, "GENERAL", "GRID_ON", False))
    lstvAttachments.GridLines = bGridOn
    
    If bHideDeleted Then
        chkHideDeleted.Value = vbChecked
    Else
        chkHideDeleted.Value = vbUnchecked
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub LoadHeaderlstvAttach"
End Sub

Public Function LoadMe() As Boolean
    On Error GoTo EH
    
    mbLoadingMe = True
    
    If Not mfrmClaim.SetadoRSRTAttachments(msAssignmentsID) Then
        Exit Function
    End If
    
    PopulatelstvAttach
    
    LoadMe = True
    mbLoadingMe = False
    Exit Function
EH:
    mbLoadingMe = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function LoadMe"
End Function

Public Function SaveMe() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim iCurrentStatus As V2ECKeyBoard.AssgnStatus
    Dim sSQL As String
    'If Close date is Set then be sure all the other dates are set tooooo
    Dim bCloseDateIsSet As Boolean
    ' Vars
    
    'validate all the fields on this form
    goUtil.utValidate Me
    
'    'Check for Drop Down Items Not Selected that should be
'    If cboAssignmentType.ListIndex = -1 Then
'        sMess = sMess & "Assignment Type not selected !" & vbCrLf
'    End If
'    If cboCatCode.ListIndex = -1 Then
'        sMess = sMess & "Cat Code not selected !" & vbCrLf
'    End If
'    If cboACID.ListIndex = -1 Then
'        sMess = sMess & "ACID not selected !" & vbCrLf
'    End If
'    If cboACIDDisplay.ListIndex = -1 Then
'        sMess = sMess & "ACID Display not selected !" & vbCrLf
'    End If
'    If cboMAState.ListIndex = -1 Then
'        sMess = sMess & "Mailing State not selected !" & vbCrLf
'    End If
'    If cboPAState.ListIndex = -1 Then
'        sMess = sMess & "Property State not selected !" & vbCrLf
'    End If
'    If cboTypeOfLoss.ListIndex = -1 Then
'        sMess = sMess & "Type Of Loss not selected !" & vbCrLf
'    End If
'
'    'DATES !!!!
'    'Close Date
'    If IsDate(txtCloseDate.Text) Then
'        'Check for Close date but no other dates filled out
'        bCloseDateIsSet = True
'        sCloseDate = "#" & Format(txtCloseDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        sCloseDate = "null"
'    End If
'
'    'Loss Date
'    If IsDate(txtLossDate.Text) Then
'        sLossDate = "#" & Format(txtLossDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        'Check for Close date but no other dates filled out
'        If bCloseDateIsSet Then
'            sMess = sMess & "Loss Date is not set!" & vbCrLf
'        End If
'        sLossDate = "null"
'    End If
'    'Assigned Date
'    If IsDate(txtAssignedDate.Text) Then
'        sAssignedDate = "#" & Format(txtAssignedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Assigned Date is not set!" & vbCrLf
'        End If
'        sAssignedDate = "null"
'    End If
'    'Received Date
'    If IsDate(txtReceivedDate.Text) Then
'        sReceivedDate = "#" & Format(txtReceivedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Received Date is not set!" & vbCrLf
'        End If
'        sReceivedDate = "null"
'    End If
'    'Contact Date
'    If IsDate(txtContactDate.Text) Then
'        sContactDate = "#" & Format(txtContactDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Contact Date is not set!" & vbCrLf
'        End If
'        sContactDate = "null"
'    End If
'    'Inspected Date
'    If IsDate(txtInspectedDate.Text) Then
'        sInspectedDate = "#" & Format(txtInspectedDate.Text, "MM/DD/YYYY") & "#"
'    Else
'        If bCloseDateIsSet Then
'            sMess = sMess & "Inspected Date is not set!" & vbCrLf
'        End If
'        sInspectedDate = "null"
'    End If
'
'    If sMess <> vbNullString Then
'        sMess = "Could not save " & Me.Caption & vbCrLf & vbCrLf & sMess
'        MsgBox sMess, vbExclamation + vbOKOnly, "Could Not Save Claim Information."
'        Exit Function
'    End If
'
'    'Use this to check new values to be inserted
'    'against the current values in this recordset
'    Set MyadoRSAssignments = mfrmClaim.adoRSAssignments
'
'    'set the Assignemtn vars
'    sAssignmentsID = msAssignmentsID
'
'    sID = msAssignmentsID
'
'    sAssignmentTypeID = cboAssignmentType.ItemData(cboAssignmentType.ListIndex)
'
'    sClientCompanyCatSpecID = cboCatCode.ItemData(cboCatCode.ListIndex)
'
'    sAdjusterSpecID = cboACID.ItemData(cboACID.ListIndex)
'
'    sAdjusterSpecIDDisplay = cboACIDDisplay.ItemData(cboACIDDisplay.ListIndex)
'
'    sSPVersion = "[SPVersion]"
'    'IBNUM
'    sIBNUM = "'" & goUtil.utCleanSQLString(UCase(txtIBNUM.Text)) & "'"
'    'CLIENTNUM
'    sCLIENTNUM = "'" & goUtil.utCleanSQLString(UCase(txtCLIENTNUM.Text)) & "'"
'    'Policy Number
'    sPolicyNo = "'" & goUtil.utCleanSQLString(UCase(txtPolicyNo.Text)) & "'"
'    'Policty Description
'    sPolicyDescription = "'" & goUtil.utCleanSQLString(UCase(txtPolicyDescription.Text)) & "'"
'    'Insured
'    sInsured = "'" & goUtil.utCleanSQLString(UCase(txtInsured.Text)) & "'"
'
'    'Mailing Address
'    'Street
'    sMAStreet = UCase(txtMAStreet.Text)
'    'City
'    sMACity = UCase(txtMACity.Text)
'    'State
'    sMAState = left(UCase(cboMAState.Text), 2)
'    'Zip
'    sMAZIP = txtMAZIP.Text
'    'Zip4
'    sMAZIP4 = txtMAZIP4.Text
'    'Other Post Code
'    sMAOtherPostCode = UCase(txtMAOtherPostCode.Text)
'    'Build entire Address
'    If sMAZIP = "00000" & sMAZIP4 = "0000" Then
'        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, sMAOtherPostCode)) & "'"
'    Else
'        sMailingAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sMAStreet, sMACity, sMAState, Format(sMAZIP, "00000") & "-" & Format(sMAZIP4, "0000"))) & "'"
'    End If
'    'Street
'    sMAStreet = "'" & goUtil.utCleanSQLString(UCase(txtMAStreet.Text)) & "'"
'    'City
'    sMACity = "'" & goUtil.utCleanSQLString(UCase(txtMACity.Text)) & "'"
'    'State
'    sMAState = "'" & goUtil.utCleanSQLString(left(UCase(cboMAState.Text), 2)) & "'"
'    'Zip
'    sMAZIP = txtMAZIP.Text
'    'Zip4
'    sMAZIP4 = txtMAZIP4.Text
'    'Other Post Code
'    sMAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtMAOtherPostCode.Text)) & "'"
'    'End Mailing Address
'
'    'Property Address
'    'Street
'    sPAStreet = UCase(txtPAStreet.Text)
'    'City
'    sPACity = UCase(txtPACity.Text)
'    'State
'    sPAState = left(UCase(cboPAState.Text), 2)
'    'Zip
'    sPAZIP = txtPAZIP.Text
'    'Zip4
'    sPAZIP4 = txtPAZIP4.Text
'    'other PostCode
'    sPAOtherPostCode = UCase(txtPAOtherPostCode.Text)
'    'Build entire Address
'    If sPAZIP = "00000" & sPAZIP4 = "0000" Then
'        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, sPAOtherPostCode)) & "'"
'    Else
'        sPropertyAddress = "'" & goUtil.utCleanSQLString(goUtil.utJoinToAddress(sPAStreet, sPACity, sPAState, Format(sPAZIP, "00000") & "-" & Format(sPAZIP4, "0000"))) & "'"
'    End If
'    'Street
'    sPAStreet = "'" & goUtil.utCleanSQLString(UCase(txtPAStreet.Text)) & "'"
'    'City
'    sPACity = "'" & goUtil.utCleanSQLString(UCase(txtPACity.Text)) & "'"
'    'State
'    sPAState = "'" & goUtil.utCleanSQLString(left(UCase(cboPAState.Text), 2)) & "'"
'    'Zip
'    sPAZIP = txtPAZIP.Text
'    'Zip4
'    sPAZIP4 = txtPAZIP4.Text
'    'Other Post Code
'    sPAOtherPostCode = "'" & goUtil.utCleanSQLString(UCase(txtPAOtherPostCode.Text)) & "'"
'    'End Property Address
'
'    'Home Phone
'    sHomePhone = "'" & goUtil.utCleanSQLString(UCase(txtHomePhone.Text)) & "'"
'    'Business Phone
'    sBusinessPhone = "'" & goUtil.utCleanSQLString(UCase(txtBusinessPhone.Text)) & "'"
'    'Mortgage name
'    sMortgageeName = "'" & goUtil.utCleanSQLString(UCase(txtMortgageeName.Text)) & "'"
'    'Agent No
'    sAgentNo = "'" & goUtil.utCleanSQLString(UCase(txtAgentNo.Text)) & "'"
'    'Reported By
'    sReportedBy = "'" & goUtil.utCleanSQLString(UCase(txtReportedBy.Text)) & "'"
'    'Reported by Phone
'    sReportedByPhone = "'" & goUtil.utCleanSQLString(UCase(txtReportedByPhone.Text)) & "'"
'    'Deductible
'    sDeductible = txtDeductible.Text
'
'    sAppDedClassTypeIDOrder = "[AppDedClassTypeIDOrder]"
'
'    'if the Loss report was changed to TEXT then need to update these vars
'    'otherwise they remain the same!
'    '(Attaching a PDF Loss Report already updates Assignments table See --> Private Sub cmdAttachPDFLossReport_Click)
'    If StrComp(cboAssignmentLossReportFormat.Text, "TEXT", vbTextCompare) = 0 Then
'        sLRFormat = "'TEXT'"
'        sLossReport = "'" & goUtil.utCleanSQLString(txtLossReport.Text) & "'"
'
'        sDownLoadLossReport = "[DownLoadLossReport]"
'
'        sUpLoadLossReport = "True"
'    Else
'        sLRFormat = "[LRFormat]"
'
'        sLossReport = "[LossReport]"
'
'        sDownLoadLossReport = "[DownLoadLossReport]"
'
'        sUpLoadLossReport = "[UpLoadLossReport]"
'    End If
'    'Type Of Loss
'    sTypeOfLossID = cboTypeOfLoss.ItemData(cboTypeOfLoss.ListIndex)
'
'    sXactTypeOfLoss = "[XactTypeOfLoss]"
'
'    sSentToXact = "[SentToXact]"
'
'    sReassigned = "[Reassigned]"
'
'    sDateReassigned = "[DateReassigned]"
'
'
'    'STATUS ID !
'    'Check for Closed Date
'    'Change Status ID
'    If IsDate(txtCloseDate.Text) Then
'        sStatusID = CStr(V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED)
'    Else
'        'Check to see if the Current Status is Closed if it Is Need to
'        'Change the Status to NEW
'        iCurrentStatus = MyadoRSAssignments.Fields("StatusID").Value
'        Select Case iCurrentStatus
'            Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED
'                sStatusID = V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_NEW
'            Case Else
'                sStatusID = "[StatusID]"
'        End Select
'
'    End If
'
'    sRAAdjusterSpecID = "[RAAdjusterSpecID]"
'
'    sIsLocked = "[IsLocked]"
'
'    sIsDeleted = "[IsDeleted]"
'
'    sDownLoadMe = "[DownLoadMe]"
'
'    sUpLoadMe = "True"
'
'    sDownLoadAll = "[DownLoadAll]"
'
'    sUpLoadAll = "[UpLoadAll]"
'    'Admin Comments
'    sAdminComments = "'" & goUtil.utCleanSQLString(txtAdminComments.Text) & "'"
'
'    sMiscDelimSettings = "[MiscDelimSettings]"
'
'    sDateLastUpdated = "#" & Format(Now(), "MM/DD/YYYY") & "#"
'
'    sUID = goUtil.gsCurUsersID
'
'
'    sSQL = "Update Assignments Set "
'    sSQL = sSQL & "[AssignmentsID] = " & sAssignmentsID & ", "
'    sSQL = sSQL & "[ID] = " & sID & ", "
'    sSQL = sSQL & "[AssignmentTypeID] = " & sAssignmentTypeID & ", "
'    sSQL = sSQL & "[ClientCompanyCatSpecID] = " & sClientCompanyCatSpecID & ", "
'    sSQL = sSQL & "[AdjusterSpecID] = " & sAdjusterSpecID & ", "
'    sSQL = sSQL & "[AdjusterSpecIDDisplay] =" & sAdjusterSpecIDDisplay & ", "
'    sSQL = sSQL & "[SPVersion] = " & sSPVersion & ", "
'    sSQL = sSQL & "[IBNUM] = " & sIBNUM & ", "
'    sSQL = sSQL & "[CLIENTNUM] = " & sCLIENTNUM & ", "
'    sSQL = sSQL & "[PolicyNo] = " & sPolicyNo & ", "
'    sSQL = sSQL & "[PolicyDescription] = " & sPolicyDescription & ", "
'    sSQL = sSQL & "[Insured] = " & sInsured & ", "
'    sSQL = sSQL & "[MailingAddress] = " & sMailingAddress & ", "
'    sSQL = sSQL & "[MAStreet] = " & sMAStreet & ", "
'    sSQL = sSQL & "[MACity] = " & sMACity & ", "
'    sSQL = sSQL & "[MAState] = " & sMAState & ", "
'    sSQL = sSQL & "[MAZIP] = " & sMAZIP & ", "
'    sSQL = sSQL & "[MAZIP4] = " & sMAZIP4 & ", "
'    sSQL = sSQL & "[MAOtherPostCode] = " & sMAOtherPostCode & ", "
'    sSQL = sSQL & "[HomePhone]  = " & sHomePhone & ", "
'    sSQL = sSQL & "[BusinessPhone] = " & sBusinessPhone & ", "
'    sSQL = sSQL & "[PropertyAddress] = " & sPropertyAddress & ", "
'    sSQL = sSQL & "[PAStreet]  = " & sPAStreet & ", "
'    sSQL = sSQL & "[PACity]  = " & sPACity & ", "
'    sSQL = sSQL & "[PAState] = " & sPAState & ", "
'    sSQL = sSQL & "[PAZIP]  = " & sPAZIP & ", "
'    sSQL = sSQL & "[PAZIP4] = " & sPAZIP4 & ", "
'    sSQL = sSQL & "[PAOtherPostCode]  = " & sPAOtherPostCode & ", "
'    sSQL = sSQL & "[MortgageeName]  = " & sMortgageeName & ", "
'    sSQL = sSQL & "[AgentNo]  = " & sAgentNo & ", "
'    sSQL = sSQL & "[ReportedBy] = " & sReportedBy & ", "
'    sSQL = sSQL & "[ReportedByPhone] = " & sReportedByPhone & ", "
'    sSQL = sSQL & "[Deductible]  = " & sDeductible & ", "
'    sSQL = sSQL & "[AppDedClassTypeIDOrder] = " & sAppDedClassTypeIDOrder & ", "
'    sSQL = sSQL & "[LRFormat]  = " & sLRFormat & ", "
'    sSQL = sSQL & "[LossReport] = " & sLossReport & ", "
'    sSQL = sSQL & "[DownLoadLossReport] = " & sDownLoadLossReport & ", "
'    sSQL = sSQL & "[UpLoadLossReport] = " & sUpLoadLossReport & ", "
'    sSQL = sSQL & "[StatusID]  = " & sStatusID & ", "
'    sSQL = sSQL & "[TypeOfLossID] = " & sTypeOfLossID & ", "
'    sSQL = sSQL & "[XactTypeOfLoss] = " & sXactTypeOfLoss & ", "
'    sSQL = sSQL & "[SentToXact] = " & sSentToXact & ", "
'    sSQL = sSQL & "[LossDate] = " & sLossDate & ", "
'    sSQL = sSQL & "[AssignedDate] = " & sAssignedDate & ", "
'    sSQL = sSQL & "[ReceivedDate] = " & sReceivedDate & ", "
'    sSQL = sSQL & "[ContactDate] = " & sContactDate & ", "
'    sSQL = sSQL & "[InspectedDate] = " & sInspectedDate & ", "
'    sSQL = sSQL & "[CloseDate]  = " & sCloseDate & ", "
'    sSQL = sSQL & "[Reassigned]  = " & sReassigned & ", "
'    sSQL = sSQL & "[DateReassigned] = " & sDateReassigned & ", "
'    sSQL = sSQL & "[RAAdjusterSpecID] = " & sRAAdjusterSpecID & ", "
'    sSQL = sSQL & "[IsLocked] = " & sIsLocked & ", "
'    sSQL = sSQL & "[IsDeleted] = " & sIsDeleted & ", "
'    sSQL = sSQL & "[DownLoadMe] = " & sDownLoadMe & ", "
'    sSQL = sSQL & "[UpLoadMe] = " & sUpLoadMe & ", "
'    sSQL = sSQL & "[DownLoadAll] = " & sDownLoadAll & ", "
'    sSQL = sSQL & "[UpLoadAll] = " & sUpLoadAll & ", "
'    sSQL = sSQL & "[AdminComments] = " & sAdminComments & ", "
'    sSQL = sSQL & "[MiscDelimSettings] = " & sMiscDelimSettings & ", "
'    sSQL = sSQL & "[DateLastUpdated] = " & sDateLastUpdated & ", "
'    sSQL = sSQL & "[UpdateByUserID] = " & sUID & " "
'    sSQL = sSQL & "WHERE AssignmentsID = " & sAssignmentsID & " "
'
'
'    Set oConn = New ADODB.Connection
'    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
'
'    oConn.Execute sSQL
    
    cmdSave.Enabled = False
    SaveMe = True
    
    'cleanup
    Set oConn = Nothing
'    Set MyadoRSAssignments = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveMe"
End Function

Public Function CheckStatus() As Boolean
    On Error GoTo EH
'    Dim lTabsPos As Long
'    Dim oFrame As Control
'    Dim MyFrame As Frame
'    Dim oControl As Control
'    Dim MyTextBox As TextBox
'    Dim MycmdButton As CommandButton
'    Dim sFrameName As String
'
'     'If this claim is closed only certain things can be edited
'    If mfrmClaim.MyStatus = iAssignmentsStatus_CLOSED Then
'        For lTabsPos = 1 To TSClaimInfo.Tabs.Count
'            Select Case UCase(TSClaimInfo.Tabs(lTabsPos).Tag)
'                Case UCase(framSpecifics.Name), _
'                        UCase(framInsuredInfo.Name), _
'                        UCase(framPolicyLimits.Name)
'                    sFrameName = TSClaimInfo.Tabs(lTabsPos).Tag
'                    For Each oFrame In Me.Controls
'                        If TypeOf oFrame Is Frame Then
'                            Set MyFrame = oFrame
'                            If StrComp(MyFrame.Name, sFrameName, vbTextCompare) = 0 Then
'                                MyFrame.Enabled = False
'                                Exit For
'                            End If
'                        End If
'                    Next
'                Case UCase(framDates.Name)
'                    'Need to disable all dates except the closedate
'                    For Each oControl In Me.Controls
'                        If TypeOf oControl Is TextBox Then
'                            Set MyTextBox = oControl
'                            If StrComp(MyTextBox.Tag, "Date", vbTextCompare) = 0 Then
'                                If StrComp(MyTextBox.Name, txtCloseDate.Name, vbTextCompare) <> 0 Then
'                                    MyTextBox.Enabled = False
'                                End If
'                            End If
'                        ElseIf TypeOf oControl Is CommandButton Then
'                            Set MycmdButton = oControl
'                            If StrComp(MycmdButton.Tag, "Date", vbTextCompare) = 0 Then
'                                If StrComp(MycmdButton.Name, cmdCloseDate.Name, vbTextCompare) <> 0 Then
'                                    MycmdButton.Enabled = False
'                                End If
'                            End If
'                        End If
'                    Next
'                Case UCase(framLossReport.Name)
'                    'Need to disable all control except the closedate
'                    For Each oControl In Me.Controls
'                        If TypeOf oControl Is CommandButton Then
'                            Set MycmdButton = oControl
'                            If StrComp(MycmdButton.Name, cmdViewPDFLossReport.Name, vbTextCompare) = 0 Then
'                                MycmdButton.Enabled = True
'                            End If
'                        Else
'                            If (Not TypeOf oControl Is ImageList) And (Not TypeOf oControl Is TabStrip) Then
'                                If oControl.Container.Name = framLossReport.Name Then
'                                    oControl.Enabled = False
'                                End If
'                            End If
'                        End If
'                    Next
'            End Select
'        Next
'    End If
    
    CheckStatus = True
    
    'cleanup
'    Set oFrame = Nothing
'    Set MyFrame = Nothing
'    Set oControl = Nothing
'    Set MyTextBox = Nothing
'    Set MycmdButton = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CheckStatus"
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    SaveSetting App.EXEName, "GENERAL", "HIDE_DELETED", "True"
    
    Select Case UnloadMode
        Case vbFormControlMenu
            Cancel = True
            mbUnloadMe = True
            Me.Visible = False
            mfrmClaim.Timer_UnloadForm.Enabled = True
        Case Else
            CmdReNumberSort_Click
            CLEANUP
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    ReSizeMe
End Sub

Public Sub ReSizeMe()
    On Error Resume Next
    Dim sNavScreenPos As String
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = False
    End If
    
    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        Me.Width = Screen.Width - 10
        Me.Height = Screen.Height - (10 + mfrmClaim.Height + goUtil.utGetTaskbarHeight)
        Me.top = mfrmClaim.top + mfrmClaim.Height
        Me.left = 10
    End If
    
    'RePos Controls
    'Width and Lefts
    framAttachments.Width = Me.Width - 360
    lstvAttachments.Width = framAttachments.Width - 825
    framAttachMaint.Width = framAttachments.Width - 225
    txtAttachPDFFilePath.Width = framAttachments.Width - 6180
    chkHideDeleted.left = framAttachments.Width - 3060
    cmdDelAttachments.left = framAttachments.Width - 1860
    
    'framCommands
    framCommands.left = Me.Width - 4695
    
    'Heights and Tops
    framAttachments.Height = Me.Height - 1815
    lstvAttachments.Height = framAttachments.Height - 1680
    framAttachMaint.top = framAttachments.Height - 855
    
    'framCommands
    framCommands.top = Me.Height - 1710
    
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
    End If
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    If cmdSave.Enabled Then
        SaveMe
        If Not mfrmClaim Is Nothing Then
            mfrmClaim.RefreshMe
        End If
    End If
    Set mfrmClaim = Nothing
    Set MyGUI = Nothing
    Set mitmXSelected = Nothing

    
    CLEANUP = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub lstvAttachments_Click()
    On Error GoTo EH
    'Set the selected Photo
    itmXSelected = lstvAttachments.SelectedItem
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvAttachments_Click"
End Sub

Private Sub lstvAttachments_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo EH
    Dim sMess As String
    If lstvAttachments.SortOrder = lvwAscending Then
        lstvAttachments.SortOrder = lvwDescending
        sMess = "Descending"
    Else
        lstvAttachments.SortOrder = lvwAscending
        sMess = "Ascending"
    End If
    
    'Set the Tool Tip
    lstvAttachments.ToolTipText = "Sort By " & ColumnHeader.Text & " " & sMess
    
    'Need to see if the Column clicked was a NON text sort
    'Like Date or number.  If a non textual column was clicked
    'need to use the next column as the sort since this next column
    'will be hidden and contain a sort friendly format.
    'Sort Key is Base 0 where CoulmnHeader is not
    Select Case ColumnHeader.Index
        Case GuiAttachListView.DateLastUpdated, GuiAttachListView.SortOrder
            lstvAttachments.SortKey = ColumnHeader.Index
        Case Else
            lstvAttachments.SortKey = ColumnHeader.Index - 1
    End Select
    
    lstvAttachments.Sorted = True
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvAttachments_ColumnClick"
End Sub

Private Sub lstvAttachments_DblClick()
    On Error GoTo EH
    'Set the selected claim
    
    itmXSelected = lstvAttachments.SelectedItem
    If Not lstvAttachments.SelectedItem Is Nothing Then
        EditAttachment
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvAttachments_DblClick"
End Sub

Public Function EditAttachment() As Boolean
    On Error GoTo EH
    Dim oListView As MSComctlLib.ListView
    Dim itmX As MSComctlLib.ListItem

    If lstvAttachments.ListItems.Count = 0 Then
        Exit Function
    Else
        Set oListView = lstvAttachments
    End If

    Set itmX = oListView.SelectedItem
    
    With EditAttach
        .MyAttachments = Me
        .MyfrmClaim = Me.MyfrmClaim
        .AssignmentsID = itmX.SubItems(GuiAttachListView.IDAssignments - 1)
        .AttachID = itmX.SubItems(GuiAttachListView.ID - 1)
         Load EditAttach
        .Caption = "Edit Attachment"
        .txtAttachDate.Text = itmX.Text
        .txtAttachName.Text = itmX.SubItems(GuiAttachListView.AttachName - 1)
        .txtDescription.Text = itmX.SubItems(GuiAttachListView.Description - 1)
        .cmdSave.Enabled = True
        .LoadDescriptionList
        .Show vbModal
    End With
   

    Unload EditAttach
    Set EditAttach = Nothing
    If lstvAttachments.Visible Then
        lstvAttachments.SetFocus
    End If
    
    EditAttachment = True
    
    Set oListView = Nothing
    Set itmX = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditAttachment"
    Unload EditAttach
End Function

Private Sub lstvAttachments_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo EH
    itmXSelected = lstvAttachments.SelectedItem
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvAttachments_ItemClick"
End Sub

Private Sub lstvAttachments_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyReturn
            EditAttachment
        Case vbKeyDelete
            cmdDelAttachments_Click
    End Select
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvAttachments_KeyDown"
End Sub

Private Sub lstvAttachments_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    If Button = vbRightButton Then
        PopupMenu PopUpmnuAttachment
    End If
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub lstvAttachments_MouseUp"
End Sub

Private Sub mnuDeleteAttachment_Click()
    On Error GoTo EH
    Dim itmX As MSComctlLib.ListItem
    Dim sAttachID As String
    
    Set itmX = lstvAttachments.SelectedItem
    
    If Not itmX Is Nothing Then
        sAttachID = itmX.SubItems(GuiAttachListView.ID - 1)
        If MsgBox("Are you sure you want to delete this Attachment Item?", vbYesNo, "DELETE SELECTED ITEM") = vbYes Then
            If DeleteAttachItem(sAttachID) Then
                lstvAttachments.ListItems.Remove ("""" & sAttachID & """")
            End If
        End If
        lstvAttachments.SetFocus
    End If
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuDeleteAttachment_Click"
End Sub

Private Sub mnuEditAttachment_Click()
    On Error GoTo EH
    
    EditAttachment
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuEditAttachment_Click"
End Sub

Private Sub mnuSelectAllAttachment_Click()
    On Error GoTo EH
    Dim itmX As ListItem
    
    For Each itmX In lstvAttachments.ListItems
        itmX.Selected = True
    Next
    
    Set itmX = Nothing
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub mnuSelectAllAttachment_Click"
End Sub

Private Sub txtAttachPDFFilePath_GotFocus()
    goUtil.utSelText txtAttachPDFFilePath
End Sub


Private Sub PopulatelstvAttach()
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim RS As ADODB.Recordset
    
    'Clear the List view
    Set oListView = lstvAttachments

    oListView.Visible = False
    oListView.ListItems.Clear

    Set RS = mfrmClaim.adoRSRTAttachments

    If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF
            '1. AttachDate
            If Not IsNull(RS.Fields("AttachDate").Value) Then
                If IsDate(RS.Fields("AttachDate").Value) Then
                    Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", Format(goUtil.IsNullIsVbNullString(RS.Fields("AttachDate")), "MM/DD/YYYY"))
                    itmX.SubItems(GuiAttachListView.AttachDateSort - 1) = Format(RS.Fields("AttachDate").Value, "YYYY/MM/DD")
                Else
                    Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", vbNullString)
                    itmX.SubItems(GuiAttachListView.AttachDateSort - 1) = vbNullString
                End If
            Else
                Set itmX = oListView.ListItems.Add(, """" & goUtil.IsNullIsVbNullString(RS.Fields("ID")) & """", vbNullString)
                itmX.SubItems(GuiAttachListView.AttachDateSort - 1) = vbNullString
            End If
            
            '2. Sort Order
            itmX.SubItems(GuiAttachListView.SortOrder - 1) = goUtil.IsNullIsVbNullString(RS.Fields("SortOrder"))
            'Sort Order Sort
            itmX.SubItems(GuiAttachListView.SortOrderSort - 1) = goUtil.utNumInTextSortFormat(goUtil.IsNullIsVbNullString(RS.Fields("SortOrder")))

            '3. AttachName
            itmX.SubItems(GuiAttachListView.AttachName - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AttachName"))
            
            '4. Description
            itmX.SubItems(GuiAttachListView.Description - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Description"))
            
            '5. Attachment
            itmX.SubItems(GuiAttachListView.Attachment - 1) = goUtil.IsNullIsVbNullString(RS.Fields("Attachment"))
            'If the actual file is missing then need to show Deleted icon
            sTemp = goUtil.AttachReposPath
            sTemp = sTemp & goUtil.IsNullIsVbNullString(RS.Fields("Attachment"))
            If Not goUtil.utFileExists(sTemp) Then
                iMyIcon = GuiAttachStatusList.IsDeleted
            Else
                iMyIcon = Empty
            End If
            itmX.ListSubItems(GuiAttachListView.Attachment - 1).ReportIcon = iMyIcon
            
            '6. UpLoadAttachment
            If CBool(RS.Fields("UpLoadAttachment")) Then
                iMyIcon = GuiAttachStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadAttachment"))
            itmX.SubItems(GuiAttachListView.UploadAttachment - 1) = sFlagText
            itmX.ListSubItems(GuiAttachListView.UploadAttachment - 1).ReportIcon = iMyIcon
            
            '7. Is Deleted
            If CBool(RS.Fields("IsDeleted")) Then
                iMyIcon = GuiAttachStatusList.IsDeleted
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("IsDeleted"))
            itmX.SubItems(GuiAttachListView.IsDeleted - 1) = sFlagText
            itmX.ListSubItems(GuiAttachListView.IsDeleted - 1).ReportIcon = iMyIcon
            
            '8. UpLoad Me
            If CBool(RS.Fields("UpLoadMe")) Then
                iMyIcon = GuiAttachStatusList.UpLoadMe
            Else
                iMyIcon = Empty
            End If
            sFlagText = goUtil.GetFlagText(RS.Fields("UpLoadMe"))
            itmX.SubItems(GuiAttachListView.UpLoadMe - 1) = sFlagText
            itmX.ListSubItems(GuiAttachListView.UpLoadMe - 1).ReportIcon = iMyIcon
            
            '9. DateLastUpdated
            If Not IsNull(RS.Fields("DateLastUpdated").Value) Then
                If IsDate(RS.Fields("DateLastUpdated").Value) Then
                    itmX.SubItems(GuiAttachListView.DateLastUpdated - 1) = Format(RS.Fields("DateLastUpdated").Value, "MM/DD/YYYY HH:MM:SS")
                    itmX.SubItems(GuiAttachListView.DateLastUpdatedSort - 1) = Format(RS.Fields("DateLastUpdated").Value, "YYYY/MM/DD HH:MM:SS")
                Else
                    itmX.SubItems(GuiAttachListView.DateLastUpdated - 1) = vbNullString
                    itmX.SubItems(GuiAttachListView.DateLastUpdatedSort - 1) = vbNullString
                End If
            Else
                itmX.SubItems(GuiAttachListView.DateLastUpdated - 1) = vbNullString
                itmX.SubItems(GuiAttachListView.DateLastUpdatedSort - 1) = vbNullString
            End If
            
            '10. AdminComments
            itmX.SubItems(GuiAttachListView.AdminComments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AdminComments"))
            
            '11. ID hidden
            itmX.SubItems(GuiAttachListView.ID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("ID"))
            
            '12. IDAssignments hidden
            itmX.SubItems(GuiAttachListView.IDAssignments - 1) = goUtil.IsNullIsVbNullString(RS.Fields("IDAssignments"))
            
            '13. RTAttachmentsID hidden
            itmX.SubItems(GuiAttachListView.RTAttachmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("RTAttachmentsID"))
            
            '14. AssignmentsID hidden
            itmX.SubItems(GuiAttachListView.AssignmentsID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("AssignmentsID"))
            
            '15. DownloadAttachment hidden
            itmX.SubItems(GuiAttachListView.DownloadAttachment - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownloadAttachment"))
            
            '16. DownLoadMe hidden
            itmX.SubItems(GuiAttachListView.DownLoadMe - 1) = goUtil.IsNullIsVbNullString(RS.Fields("DownLoadMe"))
            
            '17. UpdateByUserID hidden
            itmX.SubItems(GuiAttachListView.UpdateByUserID - 1) = goUtil.IsNullIsVbNullString(RS.Fields("UpdateByUserID"))
            
            itmX.Selected = False

            RS.MoveNext
        Loop
    End If
    oListView.Visible = True
    'Cleanup
    Set itmX = Nothing
    Set RS = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulatelstvAttach"
    oListView.Visible = True
End Sub

Private Function AttachPDFFile() As Boolean
    On Error GoTo EH
    Dim sPDFAttachPath As String
    Dim sIBNUM As String
    Dim sYYMMDDHHMMSS As String
    Dim sUsersID As String
    Dim sLRFormat As String
    Dim sPDFFileName As String
    Dim sNewAttachPath As String
    Dim sAttachName As String
    Dim sMess As String
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    Dim RS As ADODB.Recordset
    Dim sSortOrder As String
    Dim itmX As ListItem
    Dim bOverWriteAttachment As Boolean
    Dim sPreviousItem As String
    Dim lCount As Long
    Dim sNow As String

    
    Screen.MousePointer = vbHourglass
    
    'see if there is an Attachment Selected and ask if want to OverWrite
    'that attachment
    
    'loop through items looking for the very first selected item
    If lstvAttachments.ListItems.Count > 0 Then
        For lCount = 1 To lstvAttachments.ListItems.Count
            If lstvAttachments.ListItems(lCount).Selected Then
                Set itmX = lstvAttachments.ListItems(lCount)
                If itmX.Key = lstvAttachments.SelectedItem.Key Then
                    Exit For
                Else
                    Set itmX = Nothing
                End If
            End If
        Next
    End If
    
    If Not itmX Is Nothing Then
        sMess = "Do you want overwrite the selected attachment?..." & vbCrLf
        sMess = sMess & "[" & itmX.SubItems(GuiAttachListView.AttachName - 1) & "] - " & itmX.SubItems(GuiAttachListView.Description - 1) & vbCrLf
        sMess = sMess & "Or do you want to add this as a new attachment?" & vbCrLf & vbCrLf
        sMess = sMess & "Click ""YES"" to overwite the selected attachment." & vbCrLf
        sMess = sMess & "Click ""NO"" to add this attachment as new."
        If MsgBox(sMess, vbQuestion + vbYesNo, "OVERWRITE SELECTED ATTACHMENT ?") = vbYes Then
            bOverWriteAttachment = True
        End If
    End If
    
    
    'Build the PdfAttachment File name
    sIBNUM = mfrmClaim.adoRSAssignments.Fields("IBNUM").Value
    sYYMMDDHHMMSS = Format(Now(), "YYMMDDHHMMSS")
    sUsersID = goUtil.gsCurUsersID
    
    'create the PDF FIle name (Example '"FRE21220_040709090032_1.pdf"
    'Note those pdf attachments uploaded via website will look like this ...
    'FRE21220_040709090032_1@3545969.pdf
    'the @#########  is Token id created by Cold Fusion...
    'Flags that it was created On the Web Site vs Easy Claim client
    'Easy Claim DOES NOT USE THE TOKEN ID !!!
    sPDFFileName = sIBNUM & "_" & sYYMMDDHHMMSS & "_" & sUsersID & ".pdf"
    
    'Get the File path to the raw file that needs to be attached
    sPDFAttachPath = txtAttachPDFFilePath.Text
    
    'get the Attachname from the actual file name
    sAttachName = Mid(sPDFAttachPath, InStrRev(sPDFAttachPath, "\") + 1)
    sAttachName = left(sAttachName, InStrRev(sAttachName, ".") - 1)
    sAttachName = left(sAttachName, 50) ' max of 50 chars
    
    If Not goUtil.utFileExists(sPDFAttachPath) Then
        Screen.MousePointer = vbNormal
        MsgBox "Can't find " & sPDFAttachPath, vbCritical + vbOKOnly, "Error reading file"
        Exit Function
    End If
    
    'The Attachment path will be in AttachRepos under install dir
    'see if it exists, build it if it does not
    sNewAttachPath = goUtil.gsInstallDir & "\AttachRepos\"
    If Not goUtil.utFileExists(sNewAttachPath, True) Then
        goUtil.utMakeDir sNewAttachPath
    End If
    
    'Add the file to sNewAttachPath
    sNewAttachPath = sNewAttachPath & sPDFFileName
    
    'Copy the file over and then update the DB
    sMess = goUtil.utCopyFile(sPDFAttachPath, sNewAttachPath)
    
    If Not sMess = vbNullString Then
        Screen.MousePointer = vbNormal
        MsgBox "Error Attaching file!" & vbCrLf & vbCrLf & sMess, vbCritical + vbOKOnly, "Error"
        Exit Function
    End If
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sNow = Now()
    'If overwriting the currently selected Attachment
    If bOverWriteAttachment Then
         'Need to get some things from the currently selected Attachment
        Set itmX = lstvAttachments.SelectedItem
        
        'Set the ID for the attachment to be overwritten
        sID = itmX.SubItems(GuiAttachListView.ID - 1)
        
        sSQL = "UPDATE RTAttachments SET "
        sSQL = sSQL & "[AttachDate] = #" & Format(Now(), "MM/DD/YYYY") & "# , "
        sSQL = sSQL & "[AttachName] = '" & goUtil.utCleanSQLString(sAttachName) & "', "
        sSQL = sSQL & "[Attachment] = '" & sPDFFileName & "', "
        sSQL = sSQL & "[UpLoadAttachment] = True, "
        sSQL = sSQL & "[UpLoadMe] = True, "
        sSQL = sSQL & "[DateLastUpdated] = #" & sNow & "# , "
        sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
        sSQL = sSQL & "WHERE [ID] = " & sID & " "
    Else
        'If adding new then
        sID = goUtil.GetAccessDBUID("ID", "RTAttachments")
        
        'Need to get the Max Sort
        sSQL = "SELECT   MAX([SortOrder]) + 1 As SortOrder "
        sSQL = sSQL & "FROM     RTAttachments "
        sSQL = sSQL & "WHERE    [IDAssignments] = " & msAssignmentsID & " "
        
        Set RS = New ADODB.Recordset
        'Use Disconnected Record Set on asUseClient Cusor ONLY !
        RS.CursorLocation = adUseClient
        RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
        
        If RS.RecordCount = 1 Then
            sSortOrder = goUtil.IsNullIsVbNullString(RS.Fields("SortOrder"))
            If sSortOrder = vbNullString Or sSortOrder = "0" Then
                sSortOrder = "1"
            End If
        Else
            sSortOrder = "1"
        End If
        sSQL = "INSERT INTO RTAttachments "
        sSQL = sSQL & "( "
        sSQL = sSQL & "[RTAttachmentsID], "
        sSQL = sSQL & "[AssignmentsID], "
        sSQL = sSQL & "[ID], "
        sSQL = sSQL & "[IDAssignments], "
        sSQL = sSQL & "[AttachDate], "
        sSQL = sSQL & "[SortOrder], "
        sSQL = sSQL & "[Description], "
        sSQL = sSQL & "[AttachName], "
        sSQL = sSQL & "[Attachment], "
        sSQL = sSQL & "[DownloadAttachment], "
        sSQL = sSQL & "[UpLoadAttachment], "
        sSQL = sSQL & "[IsDeleted], "
        sSQL = sSQL & "[DownLoadMe], "
        sSQL = sSQL & "[UpLoadMe], "
        sSQL = sSQL & "[AdminComments], "
        sSQL = sSQL & "[DateLastUpdated], "
        sSQL = sSQL & "[UpdateByUserID] "
        sSQL = sSQL & ") "
        sSQL = sSQL & "SELECT "
        sSQL = sSQL & sID & " As [RTAttachmentsID], "
        sSQL = sSQL & msAssignmentsID & " As [AssignmentsID], "
        sSQL = sSQL & sID & " As [ID], "
        sSQL = sSQL & msAssignmentsID & " As [IDAssignments], "
        sSQL = sSQL & "#" & Format(Now(), "MM/DD/YYYY") & "# As [AttachDate], "
        sSQL = sSQL & sSortOrder & " As [SortOrder], "
        sSQL = sSQL & "'' As [Description], "
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(sAttachName) & "' As [AttachName], "
        sSQL = sSQL & "'" & sPDFFileName & "' As [Attachment], "
        sSQL = sSQL & "False As [DownloadAttachment], "
        sSQL = sSQL & "True As [UpLoadAttachment], "
        sSQL = sSQL & "False As [IsDeleted], "
        sSQL = sSQL & "False As [DownLoadMe], "
        sSQL = sSQL & "True As [UpLoadMe], "
        sSQL = sSQL & "'' As [AdminComments], "
        sSQL = sSQL & "#" & sNow & "# As [DateLastUpdated], "
        sSQL = sSQL & goUtil.gsCurUsersID & " As [UpdateByUserID]"
    End If

    oConn.Execute sSQL
    
    'If just overwrote the currenlty selctged attachment itme then need to
    'Get rid of the old one.
    If bOverWriteAttachment Then
        sPreviousItem = goUtil.AttachReposPath & itmX.SubItems(GuiAttachListView.Attachment - 1)
        goUtil.utDeleteFile (sPreviousItem)
    End If
    
    'IF the Delete Original PDF After Attach check box
    'is checked then remove the original PDF
    If chkDelOrigPDF.Value = vbChecked Then
        goUtil.utDeleteFile sPDFAttachPath
    End If
    
    Sleep 500
    RefreshAttachments
    Screen.MousePointer = vbNormal
    
    'Now Selected the item just updated
    sSQL = "SELECT [ID] "
    sSQL = sSQL & "FROM RTAttachments "
    sSQL = sSQL & "WHERE [DateLastUpdated] = #" & sNow & "# "
    sSQL = sSQL & "AND [AssignmentsID] = " & msAssignmentsID & " "
    sSQL = sSQL & "AND [Attachment] = '" & sPDFFileName & "' "
    
    Set RS = New ADODB.Recordset
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        sID = RS.Fields("ID").Value
        With lstvAttachments
            For lCount = 1 To .ListItems.Count
                Set itmX = .ListItems(lCount)
                With itmX
                    If itmX.SubItems(GuiAttachListView.ID - 1) = sID Then
                        itmX.Selected = True
                        itmX.EnsureVisible
                        Exit For
                    End If
                End With
            Next
        End With
    End If
    
    AttachPDFFile = True
    'cleanup
    Set oConn = Nothing
    Set RS = Nothing
    Set itmX = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function AttachPDFFile"
End Function

Public Function EditAttachmentItem(pudtAttach As GuiAttachItem) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    'packaged Items Update
    Dim RSPackageItem As ADODB.Recordset
    Dim sPackageItemID As String
    Dim sRTAttachmentsID As String
    Dim sReportFormat As String
    Dim sAttachName As String
    
    sSQL = "UPDATE RTAttachments Set "
    If Not mbRenumberSort Then
        sSQL = sSQL & "[RTAttachmentsID] = " & pudtAttach.RTAttachmentsID & ", "
        sSQL = sSQL & "[AssignmentsID] = " & pudtAttach.AssignmentsID & ", "
        sSQL = sSQL & "[ID] = " & pudtAttach.ID & ", "
        sSQL = sSQL & "[IDAssignments] = " & pudtAttach.IDAssignments & ", "
        sSQL = sSQL & "[AttachDate] = #" & pudtAttach.AttachDate & "#, "
    End If
        sSQL = sSQL & "[SortOrder] = " & pudtAttach.SortOrder & ", "
    If Not mbRenumberSort Then
        sSQL = sSQL & "[Description] = '" & goUtil.utCleanSQLString(pudtAttach.Description) & "', "
        sSQL = sSQL & "[AttachName] = '" & goUtil.utCleanSQLString(pudtAttach.AttachName) & "', "
        sSQL = sSQL & "[Attachment] = '" & goUtil.utCleanSQLString(pudtAttach.Attachment) & "', "
        sSQL = sSQL & "[DownloadAttachment] = " & pudtAttach.DownloadAttachment & ", "
        sSQL = sSQL & "[UpLoadAttachment] = " & pudtAttach.UploadAttachment & ", "
        sSQL = sSQL & "[IsDeleted] = " & pudtAttach.IsDeleted & ", "
        sSQL = sSQL & "[DownLoadMe] = " & pudtAttach.DownLoadMe & ", "
    End If
    sSQL = sSQL & "[UpLoadMe] = " & pudtAttach.UpLoadMe & ", "
    If Not mbRenumberSort Then
        sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtAttach.AdminComments) & "', "
    End If
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtAttach.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtAttach.UpdateByUserID & " "
    sSQL = sSQL & "WHERE [IDAssignments] = " & pudtAttach.IDAssignments & " "
    sSQL = sSQL & "AND [ID] = " & pudtAttach.ID & " "
    If mbRenumberSort Then
        sSQL = sSQL & "AND [SortOrder] <> " & pudtAttach.SortOrder & " "
    End If

    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    If mbRenumberSort Then
        GoTo CLEAN_UP
    End If
    '---------------------------Package Item Update------------------
    'if so need to update the package item if it also happens to be in there
    sRTAttachmentsID = pudtAttach.RTAttachmentsID
    If sRTAttachmentsID = vbNullString Then
        GoTo CLEAN_UP
    End If
    
    mfrmClaim.SetadoRSPackageItem msAssignmentsID, vbNullString, , , sRTAttachmentsID
    Set RSPackageItem = mfrmClaim.adoRSPackageItem
    If RSPackageItem.RecordCount > 0 Then
        Do Until RSPackageItem.EOF
            sPackageItemID = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("PackageItemID"))
            sAttachName = goUtil.utCleanSQLString(pudtAttach.Attachment)
            sReportFormat = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("ReportFormat"))
            sReportFormat = left(sReportFormat, Len(sReportFormat) - 100)
            sReportFormat = sReportFormat & String(100 - Len(sAttachName), Chr(32)) & sAttachName
            sSQL = "UPDATE PackageItem SET "
            sSQL = sSQL & "[RTAttachmentsID] = " & pudtAttach.RTAttachmentsID & ", "
            sSQL = sSQL & "[IDRTAttachments] = " & pudtAttach.ID & ", "
            sSQL = sSQL & "[ReportFormat] = '" & goUtil.utCleanSQLString(sReportFormat) & "', "
            sSQL = sSQL & "[Name] = '" & goUtil.utCleanSQLString(pudtAttach.AttachName) & "', "
            sSQL = sSQL & "[Description] = '" & goUtil.utCleanSQLString(pudtAttach.Description) & "', "
            sSQL = sSQL & "[AttachmentName] = '" & goUtil.utCleanSQLString(sAttachName) & "', "
            sSQL = sSQL & "[UploadMe] = True, "
            sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
            sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
            sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
            sSQL = sSQL & "AND [PackageItemID] = " & sPackageItemID & " "
            oConn.Execute sSQL
            Sleep 100
            RSPackageItem.MoveNext
        Loop
    End If
    '---------------------------package Item Update^^^^^^^^^^^^^^^^^^^^^
CLEAN_UP:
    EditAttachmentItem = True
    'Clean up
    Set oConn = Nothing
    Exit Function
    
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditAttachmentItem"
End Function

Public Function DeleteAttachItem(psID As String) As Boolean
    On Error GoTo EH
    Dim sSQL As String
    Dim sPath As String
    Dim bUpdateAsDeletedOnly As Boolean
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    'packaged Items Update
    Dim RSPackageItem As ADODB.Recordset
    Dim bIsDeleted As Boolean
    Dim sPackageItemID As String
    Dim sRTAttachmentsID As String
    
    'Need to remove the actual PDF files as well because they are
    'not needed anymore. only if this Record has never been uploaded
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name


    sSQL = "SELECT A.[Attachment] "
    sSQL = sSQL & "FROM RTAttachments A "
    sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
    'Only allow actual deletion of PDF File that have never been uploaded
    'The Main Table Indentity will be negative number if this is true.
    sSQL = sSQL & "AND (A.[RTAttachmentsID] Is Null Or A.[RTAttachmentsID] < 0)  "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        If Not IsNull(RS!Attachment) Then
            'Remove PDF File
            sPath = goUtil.AttachReposPath & RS!Attachment
            goUtil.utDeleteFile sPath
        End If
    Else
        bUpdateAsDeletedOnly = True
    End If
    
    
    '---------------------------Package Item Update------------------
    Set RS = New ADODB.Recordset
    
    sSQL = "SELECT [RTAttachmentsID] "
    sSQL = sSQL & "FROM RTAttachments "
    sSQL = sSQL & "WHERE [ID] = " & psID & " "
    sSQL = sSQL & "AND [IDAssignments] = " & msAssignmentsID & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        If Not IsNull(RS.Fields("RTAttachmentsID")) Then
            sRTAttachmentsID = RS.Fields("RTAttachmentsID")
        End If
    End If
    '---------------------------package Item Update^^^^^^^^^^^^^^^^^^^^^

    
    If bUpdateAsDeletedOnly Then
        sSQL = "UPDATE RTAttachments SET "
        sSQL = sSQL & "[IsDeleted] = IIF([IsDeleted], False, True), "
        sSQL = sSQL & "[UpLoadMe] = True, "
        sSQL = sSQL & "[DateLastUpdated] = #" & Format(Now(), "MM/DD/YYYY HH:MM:SS") & "#, "
        sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
        sSQL = sSQL & "WHERE [ID] = " & psID & " "
        sSQL = sSQL & "AND [IDAssignments] = " & msAssignmentsID & " "
    Else
        sSQL = "DELETE * FROM RTAttachments A "
        sSQL = sSQL & "WHERE A.[ID] = " & psID & " "
        sSQL = sSQL & "AND A.[IDAssignments] = " & msAssignmentsID & " "
    End If

    oConn.Execute sSQL
    
    '---------------------------Package Item Update------------------
    'if so need to update the package item if it also happens to be in there and is not already  flagged as deleted
    If sRTAttachmentsID = vbNullString Then
        GoTo CLEAN_UP
    End If
    
    mfrmClaim.SetadoRSPackageItem msAssignmentsID, vbNullString, , , sRTAttachmentsID
    Set RSPackageItem = mfrmClaim.adoRSPackageItem
    If RSPackageItem.RecordCount > 0 Then
        Do Until RSPackageItem.EOF
            bIsDeleted = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("IsDeleted"))
            If Not bIsDeleted Then
                sPackageItemID = goUtil.IsNullIsVbNullString(RSPackageItem.Fields("PackageItemID"))
                If bUpdateAsDeletedOnly Then
                    sSQL = "UPDATE PackageItem SET "
                    sSQL = sSQL & "[IsDeleted] = True, "
                    sSQL = sSQL & "[UploadMe] = True, "
                    sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
                    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
                    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
                    sSQL = sSQL & "AND [PackageItemID] = " & sPackageItemID & " "
                    'BGS 6.13.2005 Important!!!
                    'Do not Allow Delete from Package Table if it's Sent Date
                    'has been set.
                    sSQL = sSQL & "AND [SentDate] Is Null  "
                Else
                    sSQL = "DELETE * FROM PackageItem "
                    sSQL = sSQL & "WHERE [AssignmentsID] = " & msAssignmentsID & " "
                    sSQL = sSQL & "AND [PackageItemID] = " & sPackageItemID & " "
                End If
                oConn.Execute sSQL
                Sleep 100
            End If
            RSPackageItem.MoveNext
        Loop
    End If
    '---------------------------package Item Update^^^^^^^^^^^^^^^^^^^^^
    
    DeleteAttachItem = True
    'clean up
CLEAN_UP:
    Set RS = Nothing
    Set RSPackageItem = Nothing
    Set oConn = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function DeleteAttachItem"
End Function

