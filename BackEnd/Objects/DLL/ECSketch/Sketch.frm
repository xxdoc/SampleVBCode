VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSketch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7950
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   12480
   ControlBox      =   0   'False
   Icon            =   "Sketch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Sketch.frx":0442
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   832
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   752
      ImageHeight     =   472
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sketch.frx":E4CF
            Key             =   "grey"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sketch.frx":3BC21
            Key             =   "white"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFillColor 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   11640
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picLineColor 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   11640
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   11760
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkFilled 
      Height          =   495
      Left            =   11640
      Picture         =   "Sketch.frx":3E91E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton radLineType 
      Height          =   495
      Index           =   1
      Left            =   11640
      Picture         =   "Sketch.frx":41435
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.OptionButton radLineType 
      Height          =   495
      Index           =   0
      Left            =   11640
      Picture         =   "Sketch.frx":43EB7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   11640
      Picture         =   "Sketch.frx":467BB
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox lblInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Height          =   495
      Left            =   11640
      Picture         =   "Sketch.frx":49336
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   11640
      Picture         =   "Sketch.frx":4BEEF
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton cmdUndo 
      Height          =   495
      Left            =   11640
      Picture         =   "Sketch.frx":4F017
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdLabel 
      Height          =   495
      Left            =   11640
      Picture         =   "Sketch.frx":51BC6
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Label"
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox picTest 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7170
      Left            =   360
      MousePointer    =   2  'Cross
      Picture         =   "Sketch.frx":54789
      ScaleHeight     =   474
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   749
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   11295
   End
   Begin VB.ListBox List1 
      Height          =   7275
      ItemData        =   "Sketch.frx":81ECB
      Left            =   12600
      List            =   "Sketch.frx":81ECD
      MousePointer    =   4  'Icon
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   9960
      TabIndex        =   13
      Top             =   7560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line Width"
      Height          =   255
      Left            =   9120
      TabIndex        =   19
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label lblFill 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fill Color"
      Height          =   255
      Left            =   11640
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblLineColor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Line Color"
      Height          =   375
      Left            =   11640
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblNotify 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left Click where you want your label placed - ESC to cancel"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   7560
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "frmSketch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const JPEGFILE As String = "c:\temp\ECSKETCH.jpg"

Private mousex, mousey          As Integer
Private moSketch As clsECSKETCH

Private m_Image     As cImage
Private m_Jpeg      As cJpeg
Private m_FileName  As String

Public Property Let mySketch(poSketch As clsECSKETCH)
    Set moSketch = poSketch
End Property
Public Property Set mySketch(poSketch As clsECSKETCH)
    Set moSketch = poSketch
End Property
Public Property Get mySketch() As clsECSKETCH
    Set mySketch = moSketch
End Property

Private Sub chkFilled_Click()
    If chkFilled.Value = vbChecked Then
        lblFill.visible = True
        picFillColor.visible = True
        mFilled = True
    Else
        lblFill.visible = False
        picFillColor.visible = False
        mFilled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    moSketch.Cancel = True
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    Dim a
    a = MsgBox("Are you sure you want to clear this sketch?", vbOKCancel)
    If a <> 2 Then
        picTest.Cls
        miPtCnt = 0
        List1.Clear
        cntr = 0
        iHISTORYCOUNT = 0
    End If
End Sub

Private Sub cmdLabel_Click()
Static hasClicked As Boolean
Dim info() As String
Dim i As Integer

    If hasClicked Then
        Exit Sub
    End If
    SetButtons False
    hasClicked = True
    lblInfo.Text = ""
    frmLabel.Show (vbModal)
    
    If frmLabel.Cancel Then
        lblInfo.Text = vbNullString
    Else
        lblInfo.Text = frmLabel.txtLabel.Text
    End If
    
    Unload frmLabel
    Set frmLabel = Nothing
    mbBalloonMode = True
    If lblInfo.Text <> "" Then
        lblNotify.visible = True
        Do Until mbBalloonClick Or mbEscapeBalloon
            DoEvents
            Sleep 100 'prevent locking the computer up
        Loop
    End If
    
    If lblInfo.Text <> "" And mbBalloonClick Then
        info = Split(lblInfo.Text, vbCrLf, , vbBinaryCompare)
        For i = LBound(info, 1) To UBound(info, 1)
            If info(i) = vbNullString Then
                info(i) = " "
            End If
        Next
        ReDim Preserve info(0 To 9) As String
        Call Balloon_InfoBox("left", mousex, mousey, &H80FFFF, 0, 0, 0, info(0), info(1), info(2), info(3), info(4), info(5), info(6), info(7), info(8), info(9))
        List1.List(iHISTORYCOUNT) = iHISTORYCOUNT & "). LABEL"

        mHistory(iHISTORYCOUNT).ch000_Type = "LABEL"
        mHistory(iHISTORYCOUNT).ch001_x1 = mousex
        mHistory(iHISTORYCOUNT).ch004_y1 = mousey
        mHistory(iHISTORYCOUNT).ch007_content = lblInfo.Text
        iHISTORYCOUNT = iHISTORYCOUNT + 1
        ReDim Preserve mHistory(iHISTORYCOUNT)
        
    End If
    
    lblNotify.visible = False
    mbEscapeBalloon = False
    mbBalloonMode = False
    mbBalloonClick = False
    hasClicked = False
    SetButtons True
End Sub

Private Sub cmdSave_Click()
    Dim wddx As Variant
    moSketch.Save = True
    If UBound(mHistory, 1) > 0 Then
        ReDim Preserve mHistory(UBound(mHistory) - 1)
    Else
        moSketch.WddxXml = vbNullString
        mHistory(0).ch000_Type = vbNullString
     
    End If
    If mHistory(0).ch000_Type <> vbNullString Then
       
        'drop mHistory array into moSketch
        moSketch.myCurves = mHistory
        'serialize line/box/arc/label data
        moSketch.SerializeToWddxPacket (True)
        
        'get mHistory ready for another possible line
        ReDim Preserve mHistory(UBound(mHistory) + 1)
        
        
          'Save the pictest picture viewer to jpg
            Set m_Jpeg = New cJpeg
            Call Notify("Saving Sketch...One Moment", True)
            DoEvents
            'Lose the grid background to save file space
            Call LoadPic("white")
            
            Call RedrawPic(False)
            'Sample the cImage by hDC
            m_Jpeg.SampleHDC picTest.hDC, picTest.Width, picTest.Height
            
            'BGS 12.13.2004
            'Since the temp JPEG depends upon the user having a C:\Temp
            'Directory, need to be sure that the directory actually exists...
            If Not DirectoryExists("C:\Temp") Then
                CreatePath "C:\Temp"
            End If
            
           'Delete file if it exists
            RidFile JPEGFILE
    
           'Save the JPG file
            m_Jpeg.SaveFile JPEGFILE
            moSketch.myJPGPath = JPEGFILE
            'load back the grid
            Call LoadPic("grey")
            Call RedrawPic(False)
            Call Notify("", False)
            
        Set m_Image = Nothing
        Set m_Jpeg = Nothing
    Else
        moSketch.WddxXml = vbNullString
        
    End If
    Me.Hide
End Sub

Private Sub cmdUndo_Click()
    Call RedrawPic(True)
End Sub

Private Sub Form_Activate()


Dim sText As String

    Me.DrawWidth = 3
    
    'sText = "PUT TEXT HERE"
    'Call DrawChars(picTest, sText, 50, vbBlue, 10, 10, True, True, vbRed, False)

    Me.DrawWidth = 1

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And mbBalloonMode Then
        mbEscapeBalloon = True
        lblNotify.visible = False
    End If
    
End Sub

Private Sub Form_Load()
Dim wddx As Variant
Dim b As Variant
    'new blank drawing set line/arc/balloon counter to 0
    iHISTORYCOUNT = 0
    ReDim Preserve mHistory(iHISTORYCOUNT + 1)
    SetmMyFrmSketch Me
    mLineWidth = 1
    ForeColor = 0
    FillColor = 16777215
    picLineColor.BackColor = ForeColor
    picFillColor.BackColor = FillColor
    If Trim(moSketch.WddxXml) <> vbNullString Then
        moSketch.DeSerializeWddxPacket (True)
        mHistory = moSketch.myCurves
        iHISTORYCOUNT = UBound(mHistory, 1) + 1
        ReDim Preserve mHistory(0 To iHISTORYCOUNT)
        RedrawPic (False)
        'restore draw mode to last one used
        'Bgs 10.22.2004
        If mHistory(iHISTORYCOUNT - 1).ch008_clr = vbNullString Then
            ForeColor = &H80000012 'Black
        Else
            ForeColor = mHistory(iHISTORYCOUNT - 1).ch008_clr
        End If
        '/Bgs 10.22.2004
        If mHistory(iHISTORYCOUNT - 1).ch009_fillclr = "" Then
           FillColor = 16777215 'white
        Else
            FillColor = mHistory(iHISTORYCOUNT - 1).ch009_fillclr
        End If
        
        picLineColor.BackColor = ForeColor
        picFillColor.AutoRedraw = True
        picFillColor.BackColor = FillColor
        
        If mHistory(iHISTORYCOUNT - 1).ch000_Type <> "LABEL" Then
            Slider1.Value = mHistory(iHISTORYCOUNT - 1).ch007_content
        End If
                
        Select Case mHistory(iHISTORYCOUNT - 1).ch000_Type
            Case "ARC"
                radLineType_Click (0)
                radLineType(0).Value = True
            Case "LINE"
                radLineType_Click (0)
                radLineType(0).Value = True
            Case "BOX"
                radLineType_Click (1)
                chkFilled.Value = vbUnchecked
                radLineType(1).Value = True
            Case "BOXF"
                radLineType_Click (1)
                chkFilled.Value = vbChecked
                radLineType(1).Value = True
        End Select
        
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set moSketch = Nothing
   
   SetmMyFrmSketch Nothing
End Sub

Private Sub picFillColor_Click()
    CD.ShowColor
    picFillColor.BackColor = CD.Color
    FillColor = CD.Color
End Sub

Private Sub picLineColor_Click()
    CD.ShowColor
    picLineColor.BackColor = CD.Color
    ForeColor = CD.Color

End Sub

Private Sub Notify(MSG As String, visible As Boolean)
    lblNotify.visible = visible
    lblNotify.Caption = MSG
End Sub
Private Sub picTest_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'Note: Notice that I'm testing for a circumference of less than 25600
'       pixels (gdPi * (mArc.dRadius * 2) < 25600) before drawing the
'       arc. For some reason (unknown to me) VB has a problem drawing
'       an arc with a circumference greater than 25,600 pixels.

If mbBalloonMode Then
    mousex = x: mousey = y
    Exit Sub
End If

Dim FScale  As Single
Dim fCirc   As Single
    picTest.DrawWidth = mLineWidth
    FScale = picTest.ScaleX(1, picTest.ScaleMode, vbPixels)
    cntr = List1.ListCount
    If Button = vbLeftButton Then
    
        Select Case miPtCnt
            Case 0
                SetButtons (False)
                If mBox And Not mFilled Then
                    Call Notify("Right click to set the opposite edge of your box", True)
                ElseIf mBox And mFilled Then
                    Call Notify("Right click to set the opposite edge of your filled box", True)
                Else
                    Call Notify("Left click = Set Next point in arc : Right Click = Make line", True)
                End If
                    
                mbshowing = False
                'Set the new point.
                mPts(0).x = x
                mPts(0).y = y

                miPtCnt = miPtCnt + 1
                
                mHistory(iHISTORYCOUNT).ch001_x1 = x
                mHistory(iHISTORYCOUNT).ch004_y1 = y
                'cntr = cntr + 1
            Case 1
                Call Notify("Left click = end this arc", True)
                'Erase the temp line.
                picTest.DrawMode = vbInvert  '2 inverted draws = erase.
                picTest.Line (mPts(0).x, mPts(0).y)-(x, y), ForeColor
                picTest.DrawMode = vbCopyPen
                               
                'Set the new point.
                mPts(1).x = x
                mPts(1).y = y
                mptTemp = mPts(1)
                
                'Draw a line between 1st 2 points and mouse position.
                picTest.DrawMode = vbInvert  'Use invert so line can be erased.
                picTest.Line (mPts(0).x, mPts(0).y)-(mptTemp.x, mptTemp.y), ForeColor
                picTest.DrawMode = vbCopyPen
                
                miPtCnt = miPtCnt + 1

                mHistory(iHISTORYCOUNT).ch002_x2 = x
                mHistory(iHISTORYCOUNT).ch005_y2 = y
                'cntr = cntr + 1
            Case 2
                Call Notify("", False)
                'Calculate the arc and draw it.
                mArc = CalcArc(mPts(0), mPts(1), mptTemp)
                fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, ForeColor, .dRadsStart, .dRadsEnd
                    End With
                Else
                    Call Drawit(vbInvert, "LINE", ForeColor, mPts(0).x, mPts(0).y, mPts(1).x, mPts(1).y)
                    Call Drawit(vbInvert, "LINE", ForeColor, mPts(1).x, mPts(1).y, mPts(2).x, mPts(2).y)
                End If

                picTest.DrawMode = vbCopyPen

                'Set the new point.
                mPts(2).x = x
                mPts(2).y = y
                
                'Calculate the arc and draw it in green.
                mArc = CalcArc(mPts(0), mPts(1), mPts(2))
                fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, ForeColor, .dRadsStart, .dRadsEnd
                    End With
                Else
                    Call Drawit(vbCopyPen, "LINE", ForeColor, mPts(0).x, mPts(0).y, mPts(1).x, mPts(1).y)
                    Call Drawit(vbCopyPen, "LINE", ForeColor, mPts(1).x, mPts(1).y, mPts(2).x, mPts(2).y)
                End If

                miPtCnt = 0
                List1.List(cntr) = iHISTORYCOUNT & "). ARC"
                
                mHistory(iHISTORYCOUNT).ch003_x3 = x
                mHistory(iHISTORYCOUNT).ch006_y3 = y
                mHistory(iHISTORYCOUNT).ch000_Type = "ARC"
                mHistory(iHISTORYCOUNT).ch008_clr = ForeColor
                mHistory(iHISTORYCOUNT).ch007_content = mLineWidth
                iHISTORYCOUNT = iHISTORYCOUNT + 1
                ReDim Preserve mHistory(iHISTORYCOUNT)
                
                cntr = cntr + 1
               
                 SetButtons True
        End Select
        
    ElseIf Button = vbRightButton Then
        'end the line currently being drawn
        Call Notify("", False)
        If miPtCnt > 1 Then
            List1.List(cntr) = iHISTORYCOUNT & "). ARC"
            mHistory(iHISTORYCOUNT).ch003_x3 = x
            mHistory(iHISTORYCOUNT).ch006_y3 = y
            mHistory(iHISTORYCOUNT).ch000_Type = "ARC"
            mHistory(iHISTORYCOUNT).ch008_clr = ForeColor
            mHistory(iHISTORYCOUNT).ch009_fillclr = FillColor
            mHistory(iHISTORYCOUNT).ch007_content = mLineWidth
            picTest.DrawMode = vbCopyPen

                'Set the new point.
                mPts(2).x = x
                mPts(2).y = y
                
                'Calculate the arc and draw it in green.
                mArc = CalcArc(mPts(0), mPts(1), mPts(2))
                fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, ForeColor, .dRadsStart, .dRadsEnd
                    End With
                Else
                    Call Drawit(vbCopyPen, "LINE", ForeColor, mPts(0).x, mPts(0).y, mPts(1).x, mPts(1).y)
                    Call Drawit(vbCopyPen, "LINE", ForeColor, mPts(1).x, mPts(1).y, mPts(2).x, mPts(2).y)
                End If
            iHISTORYCOUNT = iHISTORYCOUNT + 1
            ReDim Preserve mHistory(iHISTORYCOUNT)
            SetButtons True
            cntr = cntr + 1
        ElseIf miPtCnt = 1 Then
            If mBox Then
                If mFilled Then
                  mHistory(iHISTORYCOUNT).ch000_Type = "BOXF"
                  List1.List(cntr) = iHISTORYCOUNT & "). Filled BOX"
                Else
                  mHistory(iHISTORYCOUNT).ch000_Type = "BOX"
                  List1.List(cntr) = iHISTORYCOUNT & "). BOX"
                End If
            Else
                mHistory(iHISTORYCOUNT).ch000_Type = "LINE"
                List1.List(cntr) = iHISTORYCOUNT & "). LINE"
            End If
            
            mHistory(iHISTORYCOUNT).ch002_x2 = x
            mHistory(iHISTORYCOUNT).ch005_y2 = y
            
            mHistory(iHISTORYCOUNT).ch008_clr = ForeColor
            mHistory(iHISTORYCOUNT).ch009_fillclr = FillColor
            mHistory(iHISTORYCOUNT).ch007_content = mLineWidth
            'erase temp line
            Call Drawit(vbInvert, "LINE", ForeColor, mHistory(iHISTORYCOUNT).ch001_x1, mHistory(iHISTORYCOUNT).ch004_y1, mHistory(iHISTORYCOUNT).ch002_x2, mHistory(iHISTORYCOUNT).ch005_y2)
            If mBox Then
                If mFilled Then
                    Call Drawit(vbCopyPen, "BOXF", ForeColor, mHistory(iHISTORYCOUNT).ch001_x1, mHistory(iHISTORYCOUNT).ch004_y1, mHistory(iHISTORYCOUNT).ch002_x2, mHistory(iHISTORYCOUNT).ch005_y2, FillColor)
                Else
                    Call Drawit(vbCopyPen, "BOX", ForeColor, mHistory(iHISTORYCOUNT).ch001_x1, mHistory(iHISTORYCOUNT).ch004_y1, mHistory(iHISTORYCOUNT).ch002_x2, mHistory(iHISTORYCOUNT).ch005_y2)
                End If
            Else
                Call Drawit(vbCopyPen, "LINE", ForeColor, mHistory(iHISTORYCOUNT).ch001_x1, mHistory(iHISTORYCOUNT).ch004_y1, mHistory(iHISTORYCOUNT).ch002_x2, mHistory(iHISTORYCOUNT).ch005_y2)
            End If
            iHISTORYCOUNT = iHISTORYCOUNT + 1
            ReDim Preserve mHistory(iHISTORYCOUNT)
            SetButtons True
        End If
        
        miPtCnt = 0
        
        
    End If
    
End Sub
Private Function Drawit(Mode As Variant, LineType As Variant, clr As Variant, X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant, Optional fclr As Variant, Optional x3 As Variant, Optional y3 As Variant)
Dim FScale  As Single
Dim fCirc   As Single
    'this drawing function handles all drawing (rubber band or permanent) except labels
    picTest.DrawMode = Mode
    
    Select Case LineType
        Case "LINE"
            picTest.Line (X1, Y1)-(X2, Y2), clr
        Case "ARC"
                'Calculate the arc and draw it
                mArc = CalcArc(mPts(0), mPts(1), mPts(2))
                fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, clr, .dRadsStart, .dRadsEnd
                    End With
                Else
                    picTest.Line (mPts(0).x, mPts(0).y)-(mPts(1).x, mPts(1).y), clr
                    picTest.Line (mPts(1).x, mPts(1).y)-(mPts(2).x, mPts(2).y), clr
                End If
        Case "BOX"
            picTest.FillStyle = vbFSTransparent    ' Set FillStyle to transparent.
            picTest.Line (X1, Y1)-(X2, Y2), clr, B
        Case "BOXF"
            picTest.FillStyle = vbSolid    ' Set FillStyle to solid.
            picTest.FillColor = fclr
            picTest.Line (X1, Y1)-(X2, Y2), clr, B
    End Select
End Function

Private Sub picTest_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'Note: Notice that I'm testing for a circumference of less than 25600
'       pixels (gdPi * (mArc.dRadius * 2) < 25600) before drawing the
'       arc. For some reason (unknown to me) VB has a problem drawing
'       an arc with a circumference greater than 25,600 pixels.

Dim FScale  As Single
Dim fCirc   As Single

    FScale = picTest.ScaleX(1, picTest.ScaleMode, vbPixels)
    
    Select Case miPtCnt
        Case 1
            picTest.DrawMode = vbInvert  'Use invert so line can be erased.
            
            If mbshowing Then
                'Erase the temp line.
                picTest.Line (mPts(0).x, mPts(0).y)-(mptTemp.x, mptTemp.y), 0
                mbshowing = False
            End If
            
            mptTemp.x = x
            mptTemp.y = y
            
            'Draw a line from point to mouse
            
            picTest.Line (mPts(0).x, mPts(0).y)-(mptTemp.x, mptTemp.y), ForeColor
            picTest.DrawMode = vbCopyPen
            
            mbshowing = True
        Case 2
            'Erase the temp line.
            picTest.DrawMode = vbInvert  '2 inverted draws = erase.
            
            If mbshowing Then
                'Calculate the arc and draw it.
                mArc = CalcArc(mPts(0), mPts(1), mptTemp)
                fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, ForeColor, .dRadsStart, .dRadsEnd
                    End With
                Else
                    picTest.Line (mPts(0).x, mPts(0).y)-(mPts(1).x, mPts(1).y), ForeColor
                    picTest.Line (mPts(1).x, mPts(1).y)-(mptTemp.x, mptTemp.y), ForeColor
                End If
                mbshowing = False
            End If
            
            'Set the new point.
            mptTemp.x = x
            mptTemp.y = y
            
            'Calculate the arc and draw it.
            mArc = CalcArc(mPts(0), mPts(1), mptTemp)
            fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
            If mArc.bValidArc And fCirc < 25600 Then
                With mArc
                    picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, ForeColor, .dRadsStart, .dRadsEnd
                End With
            Else
                picTest.Line (mPts(0).x, mPts(0).y)-(mPts(1).x, mPts(1).y), ForeColor
                picTest.Line (mPts(1).x, mPts(1).y)-(mptTemp.x, mptTemp.y), ForeColor
            End If
            
            picTest.DrawMode = vbCopyPen
            mbshowing = True
           
    End Select
    Text1.Text = x
    Text2.Text = y
   
End Sub

Private Sub SetButtons(status As Boolean)

    cmdUndo.Enabled = status
    cmdCancel.Enabled = status
    cmdLabel.Enabled = status
    cmdClear.Enabled = status
    cmdSave.Enabled = status

End Sub

Private Sub picTest_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mbBalloonMode Then
        If Button = vbLeftButton Then
            mbBalloonClick = True
        End If
    End If
    
End Sub

Private Sub radLineType_Click(Index As Integer)
    If Index = 1 Then
        chkFilled.visible = True
        mBox = True
        If chkFilled.Value = vbChecked Then
            lblFill.visible = True
            picFillColor.visible = True
        Else
            lblFill.visible = False
            picFillColor.visible = False
        End If
    Else
        chkFilled.visible = False
        mBox = False
        lblFill.visible = False
        picFillColor.visible = False
    End If

End Sub

Private Sub RedrawPic(undo As Boolean)
Dim fCirc   As Single
Dim i, b As Integer
Dim info() As String
Dim FScale  As Single
    
    If undo Then
        picTest.Cls
        List1.Clear
        'remove the last item in the list & redraw
        'draw arc here
        picTest.DrawMode = vbCopyPen
              
        For i = 0 To iHISTORYCOUNT - 2
            'Set the new point.
            mPts(0).x = mHistory(i).ch001_x1
            mPts(0).y = mHistory(i).ch004_y1
            mPts(1).x = mHistory(i).ch002_x2
            mPts(1).y = mHistory(i).ch005_y2
            mPts(2).x = mHistory(i).ch003_x3
            mPts(2).y = mHistory(i).ch006_y3
            Select Case mHistory(i).ch000_Type
                Case "ARC"
                    'Calculate the arc and draw it.
                    picTest.DrawWidth = mHistory(i).ch007_content
                    mArc = CalcArc(mPts(0), mPts(1), mPts(2))
                    fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
                    If mArc.bValidArc And fCirc < 25600 Then
                        With mArc
                            picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, mHistory(i).ch008_clr, .dRadsStart, .dRadsEnd
                        End With
                    End If
                Case "LINE"
                    picTest.DrawWidth = mHistory(i).ch007_content
                    Call Drawit(vbCopyPen, "LINE", mHistory(i).ch008_clr, mHistory(i).ch001_x1, mHistory(i).ch004_y1, mHistory(i).ch002_x2, mHistory(i).ch005_y2)
                Case "LABEL"
                    info = Split(mHistory(i).ch007_content, vbCrLf, , vbBinaryCompare)
                    For b = LBound(info, 1) To UBound(info, 1)
                        If info(b) = vbNullString Then
                            info(b) = " "
                        End If
                    Next
                    ReDim Preserve info(0 To 9) As String
                    Call Balloon_InfoBox("left", mHistory(i).ch001_x1, mHistory(i).ch004_y1, &H80FFFF, 0, 0, 0, info(0), info(1), info(2), info(3), info(4), info(5), info(6), info(7), info(8), info(9))
                Case "BOX"
                    picTest.DrawWidth = mHistory(i).ch007_content
                    Call Drawit(vbCopyPen, "BOX", mHistory(i).ch008_clr, mHistory(i).ch001_x1, mHistory(i).ch004_y1, mHistory(i).ch002_x2, mHistory(i).ch005_y2, mHistory(i).ch009_fillclr)
                Case "BOXF"
                    picTest.DrawWidth = mHistory(i).ch007_content
                    Call Drawit(vbCopyPen, "BOXF", mHistory(i).ch008_clr, mHistory(i).ch001_x1, mHistory(i).ch004_y1, mHistory(i).ch002_x2, mHistory(i).ch005_y2, mHistory(i).ch009_fillclr)
            End Select
            miPtCnt = 0
           
        List1.List(i) = i & "). " & mHistory(i).ch000_Type
        Next
        iHISTORYCOUNT = iHISTORYCOUNT - 1
        If iHISTORYCOUNT < 0 Then iHISTORYCOUNT = 0
        ReDim Preserve mHistory(0 To i)
    Else

        'Set picTest.Picture = Nothing
        List1.Clear
        'remove the last item in the list & redraw
        'draw arc here
        picTest.DrawMode = vbCopyPen
              
        For i = 0 To iHISTORYCOUNT - 1
            'Set the new point.
            mPts(0).x = mHistory(i).ch001_x1
            mPts(0).y = mHistory(i).ch004_y1
            mPts(1).x = mHistory(i).ch002_x2
            mPts(1).y = mHistory(i).ch005_y2
            mPts(2).x = mHistory(i).ch003_x3
            mPts(2).y = mHistory(i).ch006_y3
            
            Select Case mHistory(i).ch000_Type
                Case "ARC"
                    picTest.DrawWidth = mHistory(i).ch007_content
                    'Calculate the arc and draw it.
                    mArc = CalcArc(mPts(0), mPts(1), mPts(2))
                    fCirc = (gdPi * (mArc.dRadius * 2)) * FScale 'Circumference in pixels.
                    If mArc.bValidArc And fCirc < 25600 Then
                        With mArc
                            picTest.Circle (.ptCenter.x, .ptCenter.y), .dRadius, mHistory(i).ch008_clr, .dRadsStart, .dRadsEnd
                        End With
                    End If
                Case "LINE"
                    picTest.DrawWidth = mHistory(i).ch007_content
                    Call Drawit(vbCopyPen, "LINE", mHistory(i).ch008_clr, mHistory(i).ch001_x1, mHistory(i).ch004_y1, mHistory(i).ch002_x2, mHistory(i).ch005_y2)
                Case "LABEL"
                    mHistory(i).ch007_content = Replace(mHistory(i).ch007_content, Chr(10), vbCrLf, , , vbBinaryCompare)
                    info = Split(mHistory(i).ch007_content, vbCrLf, , vbBinaryCompare)
                    For b = LBound(info, 1) To UBound(info, 1)
                        If info(b) = vbNullString Then
                            info(b) = " "
                        End If
                    Next
                    ReDim Preserve info(0 To 9) As String
                    Call Balloon_InfoBox("left", mHistory(i).ch001_x1, mHistory(i).ch004_y1, &H80FFFF, 0, 0, 0, info(0), info(1), info(2), info(3), info(4), info(5), info(6), info(7), info(8), info(9))
                Case "BOX"
                    picTest.DrawWidth = mHistory(i).ch007_content
                    Call Drawit(vbCopyPen, "BOX", mHistory(i).ch008_clr, mHistory(i).ch001_x1, mHistory(i).ch004_y1, mHistory(i).ch002_x2, mHistory(i).ch005_y2, mHistory(i).ch009_fillclr)
                Case "BOXF"
                    picTest.DrawWidth = mHistory(i).ch007_content
                    Call Drawit(vbCopyPen, "BOXF", mHistory(i).ch008_clr, mHistory(i).ch001_x1, mHistory(i).ch004_y1, mHistory(i).ch002_x2, mHistory(i).ch005_y2, mHistory(i).ch009_fillclr)
            End Select
            miPtCnt = 0
           
        List1.List(i) = i & "). " & mHistory(i).ch000_Type
        Next

    End If

End Sub

Private Sub LoadPic(filename As String)
Dim mypic As StdPicture
 On Error Resume Next
        Set mypic = ImageList1.ListImages(filename).Picture
        If Err.Number = 0 Then
            Set m_Image = New cImage
            m_Image.CopyStdPicture mypic
        Else
            MsgBox "Can not load picture file" & vbCrLf & """" & filename & """", vbExclamation, "File Load Error"
        End If
        Set mypic = Nothing
PaintImage m_Image
End Sub
Private Sub PaintImage(TheImage As cImage)
    If ObjPtr(TheImage) = 0 Then
        picTest.Cls
    Else
        
        TheImage.PaintHDC picTest.hDC, 0, 0
        picTest.Refresh
    End If
End Sub

Private Sub Slider1_Change()
    mLineWidth = Slider1.Value
End Sub

