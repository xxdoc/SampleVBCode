Attribute VB_Name = "basBalloon"
Private gMyFrmSketch As Object
Public endx, endy, infoctr As Integer
Dim liSwitch As Variant
Dim liX As Variant
Dim liLen As Variant
Dim liTemp As Variant
Dim finum(10) As String

Public Sub SetmMyFrmSketch(poObject As Object)
    Set gMyFrmSketch = poObject
End Sub


Public Function Balloon_InfoBox(orientation As String, startx, starty, bcolor, bpointox, bpointoy, bpointcolor As Integer, Optional Info1, Optional Info2, Optional Info3, Optional Info4, Optional Info5, Optional Info6, Optional Info7, Optional Info8, Optional Info9, Optional Info10 As String)

'ORIENTATION
'-----------
'TOP
'RIGHT
'BOTTOM
'LEFT
'STARTX,STARTY : Box coordinates(TopX,TopY)
'BCOLOR : Fill color of the box
'BPOINTOX,BPOINTOY : Coordinates of where the balloon pointer points
'BPOINTCOLOR : Color of the balloon pointer
'INFO1,INFO2,INFO3,INFO4,INFO5,INFO6,INFO7,INFO8,INFO9 ,INFO10: Data to put in balloon
infoctr = 0

  If Info1 <> "" Then infoctr = infoctr + 1
  If Info2 <> "" Then infoctr = infoctr + 1
  If Info3 <> "" Then infoctr = infoctr + 1
  If Info4 <> "" Then infoctr = infoctr + 1
  If Info5 <> "" Then infoctr = infoctr + 1
  If Info6 <> "" Then infoctr = infoctr + 1
  If Info7 <> "" Then infoctr = infoctr + 1
  If Info8 <> "" Then infoctr = infoctr + 1
  If Info9 <> "" Then infoctr = infoctr + 1
  If Info10 <> "" Then infoctr = infoctr + 1
  
  Info1 = Trim(Info1)
  Info2 = Trim(Info2)
  Info3 = Trim(Info3)
  Info4 = Trim(Info4)
  Info5 = Trim(Info5)
  Info6 = Trim(Info6)
  Info7 = Trim(Info7)
  Info8 = Trim(Info8)
  Info9 = Trim(Info9)
  Info10 = Trim(Info10)
 
  finum(1) = Info1
  finum(2) = Info2
  finum(3) = Info3
  finum(4) = Info4
  finum(5) = Info5
  finum(6) = Info6
  finum(7) = Info7
  finum(8) = Info8
  finum(9) = Info9
  finum(10) = Info10
  
 'Bubble sort to find out which of the ten DATA is longer
     liLen = 10
     Do
      liSwitch = 0
      For liX = 1 To liLen - 1
        If Len(finum(liX)) > Len(finum(liX + 1)) Then
            liTemp = finum(liX)
            finum(liX) = finum(liX + 1)
            finum(liX + 1) = liTemp
            liSwitch = 1
        End If
      Next liX
      liLen = liLen - 1
     Loop Until liSwitch = 0
  
  
 
 
 endx = startx + Len(finum(10)) * 8 'Switch(Len(finum(10)) >= 1 Or Len(finum(10)) <= 5, 9, _
                                        Len(finum(10)) >= 6 Or Len(finum(10)) <= 10, 8, _
                                        Len(finum(10)) >= 11 Or Len(finum(10)) <= 20, 7, _
                                        True, 6)
                                            
 endy = starty + 8 * (infoctr * 2)
 
 Balloon_InfoBox = 0
 'If (orientation = "TOP") And (bpointoy >= starty) Then Balloon_InfoBox = 1: Beep: Exit Function
 'If (orientation = "RIGHT") And (bpointox <= endx) Then Balloon_InfoBox = 2: Beep: Exit Function
 'If (orientation = "BOTTOM") And (bpointoy <= endy) Then Balloon_InfoBox = 3: Beep: Exit Function
 'If (orientation = "LEFT") And (bpointox >= startx) Then Balloon_InfoBox = 4: Beep: Exit Function
 
   
   'Border
   gMyFrmSketch.picTest.Line (startx - 1, starty - 1)-(endx + 1, endy + 1), &H0&, BF
   
   'Main Box
   gMyFrmSketch.picTest.Line (startx, starty)-(endx, endy), bcolor, BF
   
   'PUT TEXT
   If Info1 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 3
     gMyFrmSketch.picTest.Print Info1
   End If
   
   If Info2 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 18
     gMyFrmSketch.picTest.Print Info2
   End If
   
   If Info3 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 33
     gMyFrmSketch.picTest.Print Info3
   End If
   
   If Info4 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 48
     gMyFrmSketch.picTest.Print Info4
   End If
   
   If Info5 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 63
     gMyFrmSketch.picTest.Print Info5
   End If
   
   If Info6 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 78
     gMyFrmSketch.picTest.Print Info6
   End If
   
   If Info7 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 93
     gMyFrmSketch.picTest.Print Info7
   End If
   
   If Info8 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 108
     gMyFrmSketch.picTest.Print Info8
   End If
   
   If Info9 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 123
     gMyFrmSketch.picTest.Print Info9
   End If
   
   If Info10 <> "" Then
     gMyFrmSketch.picTest.CurrentX = startx + 4
     gMyFrmSketch.picTest.CurrentY = starty + 138
     gMyFrmSketch.picTest.Print Info10
   End If
   
'   If orientation = "BOTTOM" Then
'    'Pointer Lines
'    gMyFrmSketch.picTest.Line (startx + ((endx - startx) / 2) - 10, endy + 2)-(bpointox, bpointoy), bpointcolor
'    gMyFrmSketch.picTest.Line (startx + ((endx - startx) / 2) + 10, endy + 2)-(bpointox, bpointoy), bpointcolor
'   End If
'
'  If orientation = "TOP" Then
'    'Pointer Lines
'    gMyFrmSketch.picTest.Line (startx + ((endx - startx) / 2) - 10, starty - 2)-(bpointox, bpointoy), bpointcolor
'    gMyFrmSketch.picTest.Line (startx + ((endx - startx) / 2) + 10, starty - 2)-(bpointox, bpointoy), bpointcolor
'  End If
  
  'If orientation = "LEFT" Then
    'Pointer Lines
    'gMyFrmSketch.picTest.Line (startx - 2, (starty + ((endy - starty) / 2)) - 10)-(bpointox, bpointoy), bpointcolor
    'gMyFrmSketch.picTest.Line (startx - 2, (starty + ((endy - starty) / 2)) + 10)-(bpointox, bpointoy), bpointcolor
  'End If
  
  'If orientation = "RIGHT" Then
    'Pointer Lines
    'gMyFrmSketch.picTest.Line (endx + 2, (starty + ((endy - starty) / 2)) - 10)-(bpointox, bpointoy), bpointcolor
    'gMyFrmSketch.picTest.Line (endx + 2, (starty + ((endy - starty) / 2)) + 10)-(bpointox, bpointoy), bpointcolor
  'End If
  
End Function


Private Sub Command1_Click()
  res = Balloon_InfoBox("BOTTOM", 10, 30, &H80FFFF, 70, 250, 0, "0.02", "Nugget", "", "", "", "", "", "", "", "")
  res1 = Balloon_InfoBox("TOP", 300, 150, &H80FFFF, 360, 110, 0, "10.24524", "3.335", "8.4446", "", "", "", "", "", "", "")
  res1 = Balloon_InfoBox("RIGHT", 150, 350, &H80FFFF, 450, 310, 0, "12.224", "13.335", "82.4446", "4.26656", "", "", "", "", "", "")
  res1 = Balloon_InfoBox("LEFT", 550, 350, &H80FFFF, 450, 400, 0, "12.224", "13.335", "82.4446", "4.25424", "8.43434", "23.3232", "", "", "", "")
End Sub



