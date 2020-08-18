VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmProg As frmProgress

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsProgForm"
End Property

Public Property Let Object(poProgForm As Object)
    Set mfrmProg = poProgForm
End Property
Public Property Get Object() As Object
   Set Object = mfrmProg
End Property
    
'mfrmProg.PBarFile
Public Property Let PBarFile(pProgBar As Object)
    Set mfrmProg.PBarFile.Object = pProgBar
End Property
Public Property Get PBarFile() As Object
    If Not mfrmProg Is Nothing Then
        Set PBarFile = mfrmProg.PBarFile.Object
    End If
End Property

'mfrmProg.PBarTable
Public Property Let PBarTable(pProgBar As Object)
    Set mfrmProg.PBarTable.Object = pProgBar
End Property
Public Property Get PBarTable() As Object
    Set PBarTable = mfrmProg.PBarTable.Object
End Property

'mfrmProg.PBarRecord
Public Property Let PBarRecord(pProgBar As Object)
    Set mfrmProg.PBarRecord.Object = pProgBar
End Property
Public Property Get PBarRecord() As Object
    Set PBarRecord = mfrmProg.PBarRecord.Object
End Property

'mfrmProg.framMain
Public Property Let framMainText(poText As String)
    mfrmProg.framMain.Caption = poText
End Property
Public Property Get framMainText() As String
    framMainText = mfrmProg.framMain.Caption
End Property

'mfrmProg.framTable
Public Property Let framTableText(poText As String)
    If framTableText <> vbNullString And poText = vbNullString Then
        HideFrame "framTable"
    ElseIf framTableText = vbNullString And poText <> vbNullString Then
        ShowFrame "framTable"
    End If
    mfrmProg.framTable.Caption = poText
End Property
Public Property Get framTableText() As String
    framTableText = mfrmProg.framTable.Caption
End Property

'mfrmProg.framRecord
Public Property Let framRecordText(poText As String)
    If framRecordText <> vbNullString And poText = vbNullString Then
        HideFrame "framRecord"
    ElseIf framRecordText = vbNullString And poText <> vbNullString Then
        ShowFrame "framRecord"
    End If
    mfrmProg.framRecord.Caption = poText
End Property
Public Property Get framRecordText() As String
    framRecordText = mfrmProg.framRecord.Caption
End Property

'mfrmProg.framFile
Public Property Let framFileText(poText As String)
    If framFileText <> vbNullString And poText = vbNullString Then
        HideFrame "framFile"
    ElseIf framFileText = vbNullString And poText <> vbNullString Then
        ShowFrame "framFile"
    End If
    mfrmProg.framFile.Caption = poText
End Property
Public Property Get framFileText() As String
    framFileText = mfrmProg.framFile.Caption
End Property

'mfrmProg.lblField
Public Property Let lblFieldText(poText As String)
    mfrmProg.lblField.Caption = poText
End Property
Public Property Get lblFieldText() As String
    lblFieldText = mfrmProg.lblField.Caption
End Property

'mfrmProg.lblFile
Public Property Let lblFileText(poText As String)
    mfrmProg.lblFile.Caption = poText
End Property
Public Property Get lblFileText() As String
    lblFileText = mfrmProg.lblFile.Caption
End Property

'mfrmProg.lblTable
Public Property Let lblTableText(poText As String)
    mfrmProg.lblTable.Caption = poText
End Property
Public Property Get lblTableText() As String
    lblTableText = mfrmProg.lblTable.Caption
End Property

'mfrmProg.Caption
Public Property Let Caption(poText As String)
    mfrmProg.Caption = poText
End Property
Public Property Get Caption() As String
    Caption = mfrmProg.Caption
End Property

'mfrmProg.cmdCancel
Public Property Let cmdCancelEnable(pbFlag As Boolean)
    mfrmProg.cmdCancel.Enabled = pbFlag
End Property
Public Property Get cmdCancelEnable() As Boolean
    cmdCancelEnable = mfrmProg.cmdCancel.Enabled
End Property

Public Property Let CancelMe(pbFlag As Boolean)
    mfrmProg.CancelMe = pbFlag
End Property
Public Property Get CancelMe() As Boolean
    CancelMe = mfrmProg.CancelMe
End Property

Public Property Let left(pLeft As Long)
    mfrmProg.left = pLeft
End Property
Public Property Get left() As Long
    left = mfrmProg.left
End Property

Public Property Let top(pTop As Long)
    mfrmProg.top = pTop
End Property
Public Property Get top() As Long
    top = mfrmProg.top
End Property

Public Property Let Height(pHeight As Long)
    mfrmProg.Height = pHeight
End Property
Public Property Get Height() As Long
    Height = mfrmProg.Height
End Property

Public Property Let Width(pWidth As Long)
    mfrmProg.Width = pWidth
End Property
Public Property Get Width() As Long
    Width = mfrmProg.Width
End Property
    
Public Sub SetFocus()
    On Error GoTo EH
    If Not mfrmProg Is Nothing Then
        If mfrmProg.Visible = False Then
            mfrmProg.Visible = True
            HangALeft
        End If
         mfrmProg.SetFocus
    End If
    Exit Sub
EH:
    Err.Clear
    Resume Next
End Sub

Public Sub Refresh()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
     mfrmProg.Refresh
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub Refresh"
End Sub

Public Sub RefreshMe()
    On Error GoTo EH
    Dim MyControl As Control
    Dim lErrNum As Long
    Dim sErrDesc As String
    

    For Each MyControl In mfrmProg.Controls
        If MyControl.Tag = "Refresh" Then
            MyControl.Refresh
        End If
    Next
    

    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub RefreshMe"
End Sub

Public Sub LoadForm()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mfrmProg Is Nothing Then
        If Not mfrmProg Is Nothing Then
            Unload mfrmProg
            Set mfrmProg = Nothing
        End If
        Set mfrmProg = New V2ECKeyBoard.frmProgress
        Load mfrmProg
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub LoadForm"
End Sub

Public Sub ShowForm(Optional pbDoNotHide As Boolean = False)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If Not pbDoNotHide Then
        On Error Resume Next
        mfrmProg.Show vbModeless
        If Err.Number > 0 Then
            Err.Clear
        End If
        On Error GoTo EH
    End If
    
    HangALeft
    
    goUtil.utShowFormIgnoreModality mfrmProg, pbDoNotHide
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub ShowForm"
End Sub


Public Function ShowFrame(psFrameName As String) As Boolean
    On Error GoTo EH
    Dim oControl As Control
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    For Each oControl In mfrmProg.Controls
        If TypeOf oControl Is Frame Then
            If StrComp(oControl.Name, psFrameName, vbTextCompare) = 0 Then
                oControl.Enabled = True
                ResizeFrames
                Exit Function
            End If
        End If
    Next
    
    ShowFrame = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    ShowFrame = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ShowFrame"
End Function

Public Function HideFrame(psFrameName As String) As Boolean
    On Error GoTo EH
    Dim oControl As Control
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    For Each oControl In mfrmProg.Controls
        If TypeOf oControl Is Frame Then
            If StrComp(oControl.Name, psFrameName, vbTextCompare) = 0 Then
                oControl.Enabled = False
                ResizeFrames
                Exit Function
            End If
        End If
    Next
    
    HideFrame = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    HideFrame = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function HideFrame"
End Function

Private Sub ResizeFrames()
    On Error GoTo EH
    Dim lFrameCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    With mfrmProg
        If .framTable.Enabled Then
            lFrameCount = lFrameCount + 1
            .framTable.Visible = True
            .framRecord.top = 1320
            .framFile.top = 2400
        Else
            .framTable.Visible = False
            .framRecord.top = 240
            .framFile.top = 1320
        End If
        If .framRecord.Enabled Then
            lFrameCount = lFrameCount + 1
            .framRecord.Visible = True
        Else
            .framRecord.Visible = False
            .framFile.top = 240
        End If
        If .framFile.Enabled Then
            lFrameCount = lFrameCount + 1
            .framFile.Visible = True
        Else
            .framFile.Visible = False
        End If
        
        Select Case lFrameCount
            Case 1
                .Height = 2320
            Case 2
                .Height = 3380
            Case 3
                .Height = 4440
        End Select
        'Reposition the Command button and Main fram Height
        .cmdCancel.top = .Height - 960
        .framMain.Height = .Height - 465
    End With
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ResizeFrames"
End Sub

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
    
    
    If Not mfrmProg Is Nothing Then
        Unload mfrmProg
        Set mfrmProg = Nothing
    End If
    
    CLEANUP = True
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    CLEANUP = False
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function CLEANUP"
End Function

Private Sub HangALeft()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    'Check the Left Against ECTRAY
    If Not goUtil.gfrmECTray Is Nothing Then
        If goUtil.gfrmECTray.Visible Then
            If Not mfrmProg Is Nothing Then
                If goUtil.gfrmECTray.left = 0 Then
                    mfrmProg.left = goUtil.gfrmECTray.Width + 40
                Else
                    mfrmProg.left = goUtil.gfrmECTray.left - mfrmProg.Width
                End If
                DoEvents
            End If
        End If
    End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub HangALeft"
End Sub
