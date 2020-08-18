VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5385
   Begin MSACAL.Calendar Calendar 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _Version        =   524288
      _ExtentX        =   9128
      _ExtentY        =   6800
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   7
      Day             =   13
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdtCurrentDate As Date
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type


Public Property Let CurrentDate(pdtDate As Date)
    mdtCurrentDate = pdtDate
End Property
Public Property Get CurrentDate() As Date
    CurrentDate = mdtCurrentDate
End Property

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub Calendar_DblClick()
    On Error GoTo EH
    
    mdtCurrentDate = Calendar.Value
    Me.Visible = False
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Calendar_DblClick"
End Sub


Private Sub Calendar_KeyPress(KeyAscii As Integer)
    On Error GoTo EH
    
    Select Case KeyAscii
        Case KeyCodeConstants.vbKeyEscape
            mdtCurrentDate = NULL_DATE
            Me.Visible = False
        Case KeyCodeConstants.vbKeyReturn
            mdtCurrentDate = Calendar.Value
            Me.Visible = False
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Calendar_KeyUp"
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    
    If mdtCurrentDate <> NULL_DATE Then
        Calendar.Value = mdtCurrentDate
    Else
        Calendar.Value = Now()
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub Form_Paint()
    On Error GoTo EH
    
    goUtil.utAlwaysOnTop Me, True
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    
    Select Case UnloadMode
        Case vbFormControlMenu
            mdtCurrentDate = NULL_DATE
            Cancel = True
            Me.Visible = False
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Public Sub SetTopAndLeft()
    On Error GoTo EH
    
    Dim mouse As POINTAPI
    GetCursorPos mouse
    
    Me.top = goUtil.ConvertPixelsToTwips(mouse.Y)
    Me.left = goUtil.ConvertPixelsToTwips(mouse.X)
    'check to see if the top of the calendar is hidden below screen
    If Me.top + Me.Height > Screen.Height Then
        Me.top = Screen.Height - (Me.Height + goUtil.utGetTaskbarHeight)
    End If
    'check to see if the left is hidden to the right of the screen
    If Me.left + Me.Width > Screen.Width Then
        'include taskbar height incase the taskbar is on the right of the screen!
        Me.left = Screen.Width - (Me.Width + goUtil.utGetTaskbarHeight)
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub SetTopAndLeft"
End Sub
