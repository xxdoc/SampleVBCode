VERSION 5.00
Object = "{9E3C8244-FB0C-11D1-AA5E-008048E292F1}#1.0#0"; "PolSpell.ocx"
Begin VB.Form frmSP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer SPTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   5040
   End
   Begin VB.TextBox txtView 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5540
      HideSelection   =   0   'False
      Left            =   40
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   40
      Width           =   7720
   End
   Begin SPELLCHECKERLib.SpellChecker SP 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   4260
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvaryText As Variant
Private msText As String
Private mbCancel As Boolean

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Let TextArray(pvaryText As Variant)
    mvaryText = pvaryText
End Property
Public Property Get TextArray() As Variant
    TextArray = mvaryText
End Property

Public Property Let Text(psText As String)
    msText = psText
    txtView.Text = msText
End Property
Public Property Get Text() As String
    Text = msText
End Property

Public Sub CheckSP(psText As String)
    On Error GoTo EH
    Dim iRet As Integer
    Dim sRet As String
    sRet = SP.CheckText(psText, iRet)
    If iRet = 0 Then
        psText = sRet
        mbCancel = False
    Else
        mbCancel = True
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub CheckSP"
End Sub

Public Function OpenDictionary(psDicPath As String, psCusDicPath As String) As Boolean
    OpenDictionary = SP.OpenDictionary(psDicPath, psCusDicPath)
    Exit Function
EH:
    OpenDictionary = False
End Function

Public Function CloseDictionary() As Boolean
    CloseDictionary = True
    SP.CloseDictionary
    Exit Function
EH:
    CloseDictionary = False
End Function


Private Sub SPTimer_Timer()
    On Error GoTo EH
    SPTimer.Enabled = False
    CheckEachWord
    'BGS 3.5.2002 after done checking need to hide this form
    Me.Hide
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SPTimer_Timer"
End Sub

Public Sub CheckEachWord()
    On Error GoTo EH
    Dim lCount As Long
    Dim sCheckWord As String
    Dim sText As String
    Dim lStart As Long
    Dim lSelStart As Long
    Dim lSelLen As Long
    sText = msText
    
    lStart = 1
    SP.IgnoreWordsInUppercase = False
    For lCount = LBound(mvaryText, 1) To UBound(mvaryText, 1)
        'As we check each word, we will also highlight
        'that word in the Text View using SelStart and SelLength
        
        lSelStart = InStr(lStart, sText, mvaryText(lCount), vbBinaryCompare) - 1
        lSelLen = Len(mvaryText(lCount))
        txtView.SelStart = lSelStart
        txtView.SelLength = lSelLen
        
        sCheckWord = mvaryText(lCount)
        CheckSP sCheckWord
        If mbCancel Then
            Exit For
        End If
        sText = Replace(sText, mvaryText(lCount), sCheckWord, , 1)
        txtView.Text = sText
        mvaryText(lCount) = sCheckWord
        
        'Need to figure out where to start looking
        lStart = lSelStart + Len(sCheckWord)
        '162 BGS 3.7.2002 Spell Check ECKEYBOARD ERROR # 5
        'Need to be sure that lStart is at least 1.  Blank spaces leading
        'text can cause lStart to be 0
        If lStart < 1 Then
            lStart = 1
        End If
    Next
   
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub CheckEachWord"
End Sub
    

