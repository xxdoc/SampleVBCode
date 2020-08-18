VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmSP As frmSP
Private mfrmSPCusDic As frmSPCustomDic
Private msDicPath As String
Private msCusDicPath As String

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsSP"
End Property

Public Property Let DictionaryPath(psPath As String)
    msDicPath = psPath
End Property
Public Property Get DictionaryPath() As String
    DictionaryPath = msDicPath
End Property

Public Property Let CusDictionaryPath(psPath As String)
    msCusDicPath = psPath
End Property
Public Property Get CusDictionaryPath() As String
    CusDictionaryPath = msCusDicPath
End Property

Private Sub Class_Initialize()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    Set mfrmSP = New frmSP
    Load mfrmSP
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub Class_Initialize" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Sub CheckSP(poText As Object)
    On Error GoTo EH
    Dim txtBox As TextBox
    Dim varyText As Variant
    Dim sText As String
    Dim sCheckWord As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'BGS 3.5.2002 need to split the Text object text into
    'an array of words to be checked.
    sText = poText.Text
    If sText <> vbNullString Then
        varyText = Split(sText, Chr(32), , vbBinaryCompare)
    End If
    
    If IsArray(varyText) Then
        With mfrmSP
            .Text = sText
            .TextArray = varyText
            .SPTimer.Enabled = True
            .txtView.SelStart = 0
            .txtView.SelLength = 0
            .Show vbModal
            'BGS now that we have checked each word we can join the array
            'and update the Text Object text.
            sText = .Text
            varyText = .TextArray
        End With
        sText = Join(varyText, Chr(32))
        poText.Text = sText
        MsgBox "Spell Check Complete!", vbInformation + vbOKOnly, "Spelling"
    End If
        
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub CheckSP" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not mfrmSP Is Nothing Then
        mfrmSP.CloseDictionary
        Unload mfrmSP
        Set mfrmSP = Nothing
    End If
    
    If Not mfrmSPCusDic Is Nothing Then
        mfrmSPCusDic.CLEANUP
        Unload mfrmSPCusDic
        Set mfrmSPCusDic = Nothing
    End If
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function CleanUp" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function LoadDictionaries() As Boolean
    On Error GoTo EH
    If mfrmSP Is Nothing Then
        Set mfrmSP = New frmSP
        Load mfrmSP
    End If
    LoadDictionaries = mfrmSP.OpenDictionary(msDicPath, msCusDicPath)
    Exit Function
EH:
    LoadDictionaries = False
End Function

Public Function CloseDictionaries() As Boolean
    On Error GoTo EH
    CloseDictionaries = mfrmSP.CloseDictionary
    Exit Function
EH:
    CloseDictionaries = False
End Function

Public Sub ShowCustomDictionary(Optional piMode As Integer = vbModal, _
                                Optional oMnuItem As Object)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If Not mfrmSPCusDic Is Nothing Then
        mfrmSPCusDic.Show
        mfrmSPCusDic.WindowState = vbNormal
        Exit Sub
    End If
    
    Set mfrmSPCusDic = New frmSPCustomDic
    mfrmSPCusDic.SP = Me
    Load mfrmSPCusDic
    mfrmSPCusDic.Show piMode
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Sub ShowCustomDictionary" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub
