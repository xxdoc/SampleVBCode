VERSION 5.00
Begin VB.UserControl ecsTime 
   Alignable       =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "ecsTime.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   1935
   Begin VB.VScrollBar scrTime 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Value           =   15000
      Width           =   255
   End
   Begin VB.TextBox txtTime 
      Height          =   495
      HideSelection   =   0   'False
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "ecsTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Appearance
    Flat = 0
    ThreeD = 1
End Enum
Private Const NONUM As Integer = -1
Private mscrValue As Integer
Private mbHour As Boolean
Private mbMin As Boolean
Private mbAMPM As Boolean
'Store Property Values
Private mFont As Font
Private mBackColor As OLE_COLOR
Private mForeColor As OLE_COLOR
Private mText As String
Private mAppearance As Appearance
Private mTabStop As Boolean

Public Event time24HR(ps24HR As String)
Public Event timeAMPM(psAMPM As String)

Public Property Get Appearance() As Appearance
    Appearance = mAppearance
End Property

Public Property Let Appearance(ByVal pAppearance As Appearance)
    txtTime.Appearance = pAppearance
    mAppearance = pAppearance
    PropertyChanged "Appearance"
End Property

Public Property Get Text() As String
   Text = mText
End Property

Public Property Let Text(ByVal psTime As String)
    txtTime.Text = Format(psTime, "HH:MM AMPM")
    mText = txtTime.Text
    PropertyChanged "Text"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtTime.BackColor = New_BackColor
    mBackColor = txtTime.BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtTime.ForeColor = New_ForeColor
    mForeColor = txtTime.ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
    Set Font = mFont
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtTime.Font = New_Font
    Set mFont = txtTime.Font
    PropertyChanged "Font"
   
End Property

Public Property Let TabStop(pbFlag As Boolean)
    txtTime.TabStop = pbFlag
    mTabStop = pbFlag
    PropertyChanged "TabStop"
End Property

Public Property Get TabStop() As Boolean
    TabStop = mTabStop
End Property

Private Sub UserControl_GotFocus()
    txtTime.SetFocus
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Font", mFont
        .WriteProperty "BackColor", mBackColor
        .WriteProperty "ForeColor", mForeColor
        .WriteProperty "Text", mText
        .WriteProperty "Appearance", mAppearance
        .WriteProperty "TabStop", mTabStop
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set mFont = .ReadProperty("Font", txtTime.Font)
        mBackColor = .ReadProperty("BackColor", txtTime.BackColor)
        mForeColor = .ReadProperty("ForeColor", txtTime.ForeColor)
        mText = .ReadProperty("Text", txtTime.Text)
        mAppearance = .ReadProperty("Appearance", txtTime.Appearance)
        mTabStop = .ReadProperty("TabStop", txtTime.TabStop)
    End With
    Set txtTime.Font = mFont
    txtTime.BackColor = mBackColor
    txtTime.ForeColor = mForeColor
    txtTime.Text = mText
    txtTime.Appearance = mAppearance
    txtTime.TabStop = mTabStop
End Sub


Public Property Let ecsTime(psTime As Date)
    txtTime.Text = Format(psTime, "HH:MM AMPM")
End Property

Public Property Get timeAMPM() As String
    timeAMPM = txtTime.Text
End Property
Public Property Get time24HR() As String
    time24HR = Format(txtTime.Text, "HH:MM")
End Property

Private Sub txtTime_Change()
    RaiseEvent time24HR(Format(txtTime.Text, "HH:MM"))
    RaiseEvent timeAMPM(txtTime.Text)
    mText = txtTime.Text
End Sub

Private Sub txtTime_Click()
    SelTime
End Sub

Private Sub txtTime_GotFocus()
    SelTime
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    On Error GoTo EH
    Dim iKey As Integer
    Dim sTime As String
    Dim iSelStart As Integer
    Dim iSelLen As Integer
    
    sTime = txtTime.Text
    
    iKey = GetVbKey(KeyAscii)
    
    If iKey = NONUM Then
        Exit Sub
    End If
    With txtTime
        Select Case .SelStart
            Case 0
                If iKey <> 0 Then
                    If iKey = 1 Then
                        Mid(sTime, 1, 1) = 1
                    Else
                        Mid(sTime, 1, 1) = 0
                    End If
                Else
                    Mid(sTime, 1, 1) = 0
                End If
                iSelStart = 1
                iSelLen = 1
            Case 1
                If Left(sTime, 1) = 0 Then
                    If iKey > 0 Then
                        Mid(sTime, 2, 1) = iKey
                    Else
                        Exit Sub
                    End If
                Else
                    If iKey <= 2 Then
                        Mid(sTime, 2, 1) = iKey
                    Else
                        Exit Sub
                    End If
                End If
                iSelStart = 3
                iSelLen = 2
            Case 3
                If iKey <= 5 Then
                    Mid(sTime, 4, 1) = iKey
                Else
                    Exit Sub
                End If
                iSelStart = 4
                iSelLen = 2
            Case 4
                Mid(sTime, 5, 1) = iKey
                iSelStart = 6
                iSelLen = 2
            Case 6
                If Mid(sTime, 7, 2) = "PM" Then
                    Mid(sTime, 7, 2) = "AM"
                Else
                    Mid(sTime, 7, 2) = "PM"
                End If
                iSelStart = 6
                iSelLen = 2
            Case 7
                iSelStart = 7
                iSelLen = 1
            
        End Select
        .Text = sTime
        .SelStart = iSelStart
        .SelLength = iSelLen
    End With
    Exit Sub
EH:
    MsgBox Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub


Private Sub scrTime_Change()
    Dim bIncrease As Boolean
    Dim iHour As Integer
    Dim iMin As Integer
    Dim sAMPM As String
    Dim sTime As String
    Dim iPos As Integer
    sTime = txtTime.Text
    
    SelTime
    iPos = txtTime.SelStart
    If scrTime.Value < mscrValue Then
        bIncrease = True
    Else
        bIncrease = False
    End If
    mscrValue = scrTime.Value
    If mbHour Then
        iHour = Left(sTime, 2)
        If bIncrease Then
            Select Case iHour
                Case Is <= 11
                    iHour = iHour + 1
                Case Is = 12
                    iHour = 1
            End Select
        Else
            Select Case iHour
                Case Is > 0
                    iHour = iHour - 1
                Case Is = 0
                    iHour = 12
            End Select
        End If
        Mid(sTime, 1, 2) = Format(iHour, "0#")
    ElseIf mbMin Then
        iMin = Mid(txtTime.Text, 4, 2)
        If bIncrease Then
            Select Case iMin
                Case Is <= 58
                    iMin = iMin + 1
                Case Is = 59
                    iMin = 0
            End Select
        Else
            Select Case iMin
                Case Is > 0
                    iMin = iMin - 1
                Case Is = 0
                    iMin = 59
            End Select
        End If
        Mid(sTime, 4, 2) = Format(iMin, "0#")
    ElseIf mbAMPM Then
        sAMPM = Mid(txtTime.Text, 7, 2)
        If sAMPM = "PM" Then
            sAMPM = "AM"
        Else
            sAMPM = "PM"
        End If
        Mid(sTime, 7, 2) = sAMPM
    End If
    
    txtTime.Text = sTime
    txtTime.SelStart = iPos
    SelTime

    If mscrValue < 100 Or mscrValue > 32600 Then
        scrTime.Value = 15000
        mscrValue = 15000
    End If
    
End Sub


Private Sub UserControl_Initialize()
    txtTime.Text = Format(Now, "HH:MM AMPM")
    mText = txtTime.Text
    Set mFont = txtTime.Font
    mForeColor = txtTime.ForeColor
    mBackColor = txtTime.BackColor
    mAppearance = txtTime.Appearance
    mTabStop = txtTime.TabStop
    mscrValue = 15000
End Sub


Private Sub UserControl_Resize()
     On Error Resume Next
     scrTime.Height = UserControl.Height
     scrTime.Left = UserControl.Width - 260
     txtTime.Height = UserControl.Height
     txtTime.Width = UserControl.Width - 245
End Sub

Private Function GetVbKey(piKey As Integer)
    Select Case piKey
        Case vbKey1
            GetVbKey = 1
        Case vbKey2
            GetVbKey = 2
        Case vbKey3
            GetVbKey = 3
        Case vbKey4
            GetVbKey = 4
        Case vbKey5
            GetVbKey = 5
        Case vbKey6
            GetVbKey = 6
        Case vbKey7
            GetVbKey = 7
        Case vbKey8
            GetVbKey = 8
        Case vbKey9
            GetVbKey = 9
        Case vbKey0, vbKeyA, vbKeyP, 97, 112
            GetVbKey = 0
        Case Else
            GetVbKey = NONUM
    End Select
End Function

Private Sub SelTime()
    With txtTime
        Select Case .SelStart
            Case 0, 1, 2
                mbHour = True
                mbMin = False
                mbAMPM = False
                .SelStart = 0
                .SelLength = 2
            Case 3, 4, 5
                mbMin = True
                mbHour = False
                mbAMPM = False
                .SelStart = 3
                .SelLength = 2
            Case 6, 7, 8
                mbAMPM = True
                mbHour = False
                mbMin = False
                .SelStart = 6
                .SelLength = 2
        End Select
    End With
End Sub



