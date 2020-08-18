VERSION 5.00
Begin VB.Form frmSPCustomDic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Dictionary"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomDic.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   2820
      Width           =   915
   End
   Begin VB.Frame framCommands 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3340
      Left            =   3120
      TabIndex        =   1
      Top             =   80
      Width           =   2925
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1095
         Width           =   915
      End
      Begin VB.ListBox lstDelWords 
         Height          =   1980
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1095
         Width           =   1665
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtWord 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   30
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblSel 
         Caption         =   "Selected Word(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   855
         Width           =   2055
      End
      Begin VB.Label lblNew 
         Caption         =   "New Word"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   15
         Width           =   1935
      End
   End
   Begin VB.ListBox lstWords 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   45
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   40
      Width           =   3060
   End
End
Attribute VB_Name = "frmSPCustomDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moSP As clsSP
Private Const W_COMMAND As Long = 6150
Private Const W_INIT As Long = 3200
Private mbInit As Boolean

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let SP(poSP As clsSP)
    Set moSP = poSP
End Property
Public Property Set SP(poSP As clsSP)
    Set moSP = poSP
End Property

Private Sub cmdAdd_Click()
    On Error GoTo EH
    Dim sText As String
    sText = txtWord.Text
    
    'Can't allow spaces in words because that is what our
    'spell check is separating each word out.
    If InStr(1, sText, Chr(32)) > 0 Then
        sText = Replace(sText, Chr(32), vbNullString)
        txtWord.Text = sText
        txtWord.SetFocus
        Exit Sub
    End If
    
    'if its got stuff add it
    If sText <> vbNullString Then
        lstWords.AddItem txtWord.Text
        txtWord.Text = vbNullString
    End If
    txtWord.SetFocus
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdAdd_Click"
End Sub

Private Sub cmdDone_Click()
    SaveCusDic
    'After done saveing cusdictionary need to open it again.
    moSP.LoadDictionaries
    
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo EH
    
    If lstDelWords.ListCount > 0 Then
        If MsgBox("Are you sure you want to permanently remove the selected word(s)?", vbYesNo, "Remove From Custom Dictionary") = vbYes Then
            RemoveWords
        End If
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdRemove_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    LoadWords
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    goUtil.utFormWinRegPos goUtil.gsMainAppEXEName, Me, True
    
    If Not goUtil.gfrmECTray Is Nothing Then
        goUtil.gfrmECTray.ShowMe False
    End If
    If Not moSP Is Nothing Then
        moSP.CLEANUP
    End If
    CLEANUP
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub lstWords_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdateWordSelection
End Sub

Private Sub lstWords_LostFocus()
    UpdateWordSelection
End Sub

Private Sub lstWords_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateWordSelection
End Sub

Private Sub txtWord_GotFocus()
    On Error GoTo EH
    SelText txtWord
    Exit Sub
EH:
    Err.Clear
End Sub

Private Function LoadWords() As Boolean
    On Error GoTo EH
    Dim varyWords As Variant
    Dim lCount As Long
    
    'Need to close dictionaries because we will be making
    'changes to the custom dictionary
    moSP.CloseDictionaries
    
    If FileExists(moSP.CusDictionaryPath) Then
        varyWords = Split(GetFileData(moSP.CusDictionaryPath), vbCrLf)
    End If
    
    lstWords.Clear
    
    If IsArray(varyWords) Then
        For lCount = LBound(varyWords, 1) To UBound(varyWords, 1)
            'Skip the very first one is a null string
            If varyWords(lCount) <> vbNullString Then
                lstWords.AddItem varyWords(lCount)
            End If
        Next
    End If
    
    LoadWords = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function LoadWords"
End Function

Private Function UpdateWordSelection() As Boolean
    On Error GoTo EH
    Dim lCount As Long
    
    lstDelWords.Clear
    
    For lCount = 0 To lstWords.ListCount - 1
        If lstWords.Selected(lCount) Then
            lstDelWords.AddItem lstWords.List(lCount)
        End If
    Next
    UpdateWordSelection = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UpdateWordSelection"
End Function

Private Function RemoveWords() As Boolean
    On Error GoTo EH
    Dim lCount As Long
    Dim lCount2 As Long
    
    For lCount = 0 To lstDelWords.ListCount - 1
        For lCount2 = 0 To lstWords.ListCount - 1
            If lstWords.List(lCount2) = lstDelWords.List(lCount) Then
                lstWords.RemoveItem (lCount2)
            End If
            'keep looping just incase they have duplicates they want
            'to get rid of. But we also do not want to exceed the
            'the index count
            If lCount2 + 1 > lstWords.ListCount - 1 Then
                Exit For
            End If
        Next
    Next
    
    'Now that we have finished removing from lstwords we can clear.
    lstDelWords.Clear
    RemoveWords = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function RemoveWords"
End Function

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    If Not moSP Is Nothing Then
        Set moSP = Nothing
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CleanUp"
End Function

Public Function SaveCusDic() As Boolean
    On Error GoTo EH
    Dim sWords As String
    Dim lCount As Long
      
    For lCount = 0 To lstWords.ListCount - 1
        sWords = sWords & lstWords.List(lCount) & vbCrLf
    Next
    
    'BGS here is where we can Kill the existing custome dic.
    'And replace it with updated data.
    If FileExists(moSP.CusDictionaryPath) Then
        SetAttr moSP.CusDictionaryPath, vbNormal
        Kill moSP.CusDictionaryPath
    End If
    
    SaveFileData moSP.CusDictionaryPath, sWords
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SaveCusDic"
End Function
