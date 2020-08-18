Attribute VB_Name = "basGeometry"
Option Explicit

Public Type PointDbl    'Point structure (in Doubles)
    X   As Double       'X-coordinate of point.
    Y   As Double       'Y-coordinate of point.
End Type

Public Type LineDbl     'Line structure (in Doubles)
    ptStart As PointDbl 'Starting point (X, Y) on line.
    ptEnd   As PointDbl 'Ending point (X, Y) on line.
End Type

Public Type ArcStruct
    bValidArc   As Boolean  'Is this a valid arc.
    ptStart     As PointDbl 'Starting point.
    PtMid       As PointDbl 'Mid point.
    ptEnd       As PointDbl 'Ending point.
    ptCenter    As PointDbl 'Center point.
    dRadius     As Double   'Radius.
    dRadsStart  As Double   'Starting angle in radians.
    dRadsMid    As Double   'Mid angle in radians.
    dRadsEnd    As Double   'Ending angle in radians.
End Type



Public Const gdPi   As Double = 3.14159265358979    'Pi
Public Function PointOnLine(ptStart As PointDbl, ptEnd As PointDbl, ByVal dDistance As Double) As PointDbl

'Returns a point on a line at dDistance from ptStart.
'This point need not be between ptStart and ptEnd.

Dim dDX     As Single
Dim dDY     As Single
Dim dLen    As Single
Dim dPct    As Single
    
    If dDistance > 1000000 Then
        dDistance = 1000000
    End If
        
    dLen = Distance(ptStart, ptEnd)
    
    If dLen > 0 Then
        dDX = ptEnd.X - ptStart.X
        dDY = ptEnd.Y - ptStart.Y
        dPct = Div(dDistance, dLen)
        PointOnLine.X = ptStart.X + (dDX * dPct)
        PointOnLine.Y = ptStart.Y + (dDY * dPct)
    Else
        PointOnLine.X = ptStart.X
        PointOnLine.Y = ptStart.Y
    End If
    
End Function
Public Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    
' Divides 2 numbers avoiding a "Division by zero" error.

    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

Public Function CalcArc(Pt1 As PointDbl, Pt2 As PointDbl, Pt3 As PointDbl) As ArcStruct

'Calculates all data needed to draw an arc from 3 points.
'Returns an ArcStruct structure. (see declares section)

'Example Syntax:
'Dim Arc1 As ArcStruct
'   Arc1 = CalcArc(Pt1, Pt2, Pt3)
'   If Arc1.bValidArc Then
'       With Arc1
'           picBox.Circle(.ptCenter.X, .ptCenter.Y), .dRadius, lAnyColor, .dRadsStart, .dRadsEnd
'       End With
'   End If

Dim iIdx(1)     As Integer
Dim dRads(3)    As Double
Dim uLine(3)    As LineDbl
Dim ptCenter    As PointDbl

    'Setup 2 lines using the 3 points.
    uLine(0).ptStart = Pt1
    uLine(0).ptEnd = Pt2
    uLine(1).ptStart = Pt2
    uLine(1).ptEnd = Pt3
    
    'Create a perpendicular line from the
    'centers of each of the two lines.
    uLine(2) = PerpLineCenter(uLine(0))
    uLine(3) = PerpLineCenter(uLine(1))
    
    'If the perp lines don't intersect then the 3 points
    'are on a straight line and cannot be an arc.
    If LineIntersect(uLine(2), uLine(3), ptCenter) <> -1 Then
        
        'If the perp lines intersect then it forms an arc.
        'Setup 3 lines from the center; 1 line to each outer point.
        uLine(0).ptStart = ptCenter
        uLine(0).ptEnd = Pt1
        uLine(1).ptStart = ptCenter
        uLine(1).ptEnd = Pt2
        uLine(2).ptStart = ptCenter
        uLine(2).ptEnd = Pt3
        dRads(0) = LineAngleRadians(uLine(0))
        dRads(1) = LineAngleRadians(uLine(1))
        dRads(2) = LineAngleRadians(uLine(2))
        
        'An arc is always drawn counter-clockwise, so order the points.
        If Not IsBetween(dRads(1), dRads(0), dRads(2), False) Then
            'dRads(1) is not between dRads(0) and dRads(2),
            'so the arc must wrap around the 0° mark. This means the
            'greater of dRads(0) and dRads(2) is the start point.
            If dRads(2) > dRads(0) Then 'Reversed, so swap points.
                dRads(3) = dRads(0)
                uLine(3) = uLine(0)
                dRads(0) = dRads(2)
                uLine(0) = uLine(2)
                dRads(2) = dRads(3)
                uLine(2) = uLine(3)
            End If
        Else
            'No wrap around, so the lessor of dRads(0)
            'and dRads(2) is the start point.
            If dRads(2) < dRads(0) Then 'Reversed, so swap points.
                dRads(3) = dRads(0)
                uLine(3) = uLine(0)
                dRads(0) = dRads(2)
                uLine(0) = uLine(2)
                dRads(2) = dRads(3)
                uLine(2) = uLine(3)
            End If
        End If
            
        'Now that the points and angles are all in order, return the data.
        With CalcArc
            .bValidArc = True
            .ptStart = uLine(0).ptEnd
            .PtMid = uLine(1).ptEnd
            .ptEnd = uLine(2).ptEnd
            .ptCenter = ptCenter
            .dRadius = Distance(.ptCenter, .ptStart)
            .dRadsStart = dRads(0)
            .dRadsMid = dRads(1)
            .dRadsEnd = dRads(2)
        End With
    
    Else
        'Straight line; Set bValidArc to False.
        CalcArc.bValidArc = False
        
    End If
    
End Function

Public Function Distance(ptStart As PointDbl, ptEnd As PointDbl) As Double

'Calculates the distance between 2 points.

    'Standard hypotenuse equation (c = Sqr(a^2 + b^2))
    Distance = Sqr(((ptEnd.X - ptStart.X) ^ 2) + ((ptEnd.Y - ptStart.Y) ^ 2))
    
End Function


Public Function IsBetween(ByVal vTestData As Variant, ByVal vLowerBound As Variant, ByVal vUpperBound As Variant, Optional ByVal bInclusive As Boolean = True) As Boolean

'Returns True if vTestData is between vLowerBound and vUpperBound.
'bInclusive = Are the bounds included in the test?

Dim vTemp   As Variant

    If vLowerBound = vUpperBound Then
        Exit Function   'Returns false if upper and lower bounds are equal.
    Else
        If vLowerBound > vUpperBound Then
            'If bounds are reversed, swap them.
            vTemp = vLowerBound
            vLowerBound = vUpperBound
            vUpperBound = vTemp
        End If
        If bInclusive Then
            'If bounds are included in test (use >= and <=).
            IsBetween = (vTestData >= vLowerBound) And (vTestData <= vUpperBound)
        Else
            'If bounds are not included in test (use > and <).
            IsBetween = (vTestData > vLowerBound) And (vTestData < vUpperBound)
        End If
    End If
    
End Function

Public Function LineAngleDegrees(Line1 As LineDbl) As Double

'Returns the angle of a line in degrees (see LineAngleRadians).

    LineAngleDegrees = RadiansToDegrees(LineAngleRadians(Line1))
    
End Function


Public Function LineAngleRadians(Line1 As LineDbl) As Double

'Calculates the angle(in radians) of a line from ptStart to ptEnd.

Dim dDeltaX As Double
Dim dDeltaY As Double
Dim dAngle  As Double

    dDeltaX = Line1.ptEnd.X - Line1.ptStart.X
    dDeltaY = Line1.ptEnd.Y - Line1.ptStart.Y
    
    If dDeltaX = 0 Then      'Vertical
        If dDeltaY < 0 Then
            dAngle = gdPi / 2
        Else
            dAngle = gdPi * 1.5
        End If
    
    ElseIf dDeltaY = 0 Then  'Horizontal
        If dDeltaX >= 0 Then
            dAngle = 0
        Else
            dAngle = gdPi
        End If
    
    Else    'Angled
        'Note: ++ = positive X, positive Y; +- = positive X, negative Y; etc.
        'On a true coordinate plane, Y increases as it move upward.
        'In VB coordinates, Y is reversed. It increases as it moves downward.
        
        'Calc for true Upper Right Quadrant (++) (For VB this is +-)
        dAngle = Atn(Abs(dDeltaY / dDeltaX))        'VB Upper Right (+-)
        
        'Correct for other 3 quadrants in VB coordinates (Reversed Y)
        If dDeltaX >= 0 And dDeltaY >= 0 Then       'VB Lower Right (++)
            dAngle = (gdPi * 2) - dAngle
            
        ElseIf dDeltaX < 0 And dDeltaY >= 0 Then    'VB Lower Left (-+)
            dAngle = gdPi + dAngle
            
        ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left (--)
            dAngle = gdPi - dAngle
            
        End If
        
    End If
    
    LineAngleRadians = dAngle
    
End Function

Public Function PerpLineCenter(Line1 As LineDbl) As LineDbl

'Returns a line perpendicular (90°) to Line1 using
'the center of Line1 as the first point.

Dim dDeltaX As Double
Dim dDeltaY As Double
Dim Line2   As LineDbl

    Line2.ptStart.X = (Line1.ptStart.X + Line1.ptEnd.X) / 2#
    Line2.ptStart.Y = (Line1.ptStart.Y + Line1.ptEnd.Y) / 2#
    dDeltaX = Line2.ptStart.X - Line1.ptStart.X
    dDeltaY = Line2.ptStart.Y - Line1.ptStart.Y
    Line2.ptEnd.X = Line2.ptStart.X + -dDeltaY
    Line2.ptEnd.Y = Line2.ptStart.Y + dDeltaX
    
    PerpLineCenter = Line2
    
End Function

Function LineIntersect(Line1 As LineDbl, Line2 As LineDbl, ptIntersect As PointDbl) As Integer

'Calculate the intersection point of any two given non-parallel lines.
'
'Returns:  -1 = lines are parallel (no intersection).
'           0 = Neither line contains the intersect point between its points.**
'           1 = Line1 contains the intersect point between its points.**
'           2 = Line2 contains the intersect point between its points.**
'           3 = Both Lines contain the intersect point between their points.**
'           ** Lines Do intersect; Also fills in the ptIntersect point.
'
'BTW:       There are 18 lines of pure code, 25 lines of pure comments and 6
'           mixed lines in this function, just in case you were wondering. (:oþ}

Dim bIntersect  As Boolean
Dim iReturn     As Integer
Dim dDenom      As Double
Dim dPctDelta1  As Double
Dim dPctDelta2  As Double
Dim Delta(2)    As PointDbl

        'Calculate the Deltas (distance of X2 - X1 or Y2 - Y1 of any 2 points)
        Delta(0).X = Line1.ptStart.X - Line2.ptStart.X   'Line1-Line2.ptStart X-Cross-Delta
        Delta(0).Y = Line1.ptStart.Y - Line2.ptStart.Y   'Line1-Line2.ptStart Y-Cross-Delta
        Delta(1).X = Line1.ptEnd.X - Line1.ptStart.X   'Line1 X-Delta
        Delta(1).Y = Line1.ptEnd.Y - Line1.ptStart.Y   'Line1 Y-Delta
        Delta(2).X = Line2.ptEnd.X - Line2.ptStart.X   'Line2 X-Delta
        Delta(2).Y = Line2.ptEnd.Y - Line2.ptStart.Y   'Line2 Y-Delta
        
        'Calculate the denominator (zero = parallel (no intersection))
        'Formula: (L2Dy * L1Dx) - (L2Dx * L1Dy)
        iReturn = -1
        dDenom = (Delta(2).Y * Delta(1).X) - (Delta(2).X * Delta(1).Y)
        bIntersect = (dDenom <> 0)
        
        If bIntersect Then
            'The lines will intersect somewhere.
            'Solve for both lines using the Cross-Deltas (Delta(0))
            
            'This yields percentage (0.1 = 10%; 1 = 100%) of the distance
            'between ptStart and ptEnd, of the opposite line, where the line used
            'in the calculation will cross it.
            '0 = ptStart direct hit; 1 = ptEnd direct hit; 0.5 = Centered between Pts; etc.
            'If < 0 or > 1 then the lines still intersect, just not between the points.
            
            'Solve for Line1 where Line2 will cross it.
            dPctDelta1 = ((Delta(2).X * Delta(0).Y) - (Delta(2).Y * Delta(0).X)) / dDenom
            
            'Solve for Line2 where Line1 will cross it.
            dPctDelta2 = ((Delta(1).X * Delta(0).Y) - (Delta(1).Y * Delta(0).X)) / dDenom
        
            'Check for absolute intersection. If the percentage is not between
            '0 and 1 then the lines will not intersect between their points.
            'Returns 0, 1, 2 or 3.
            iReturn = IIf(IsBetween(dPctDelta1, 0#, 1#), 1, 0) _
                Or IIf(IsBetween(dPctDelta2, 0#, 1#), 2, 0)
            
            'Calculate point of intersection on Line1 and fill ptIntersect.
            ptIntersect.X = Line1.ptStart.X + (dPctDelta1 * Delta(1).X)
            ptIntersect.Y = Line1.ptStart.Y + (dPctDelta1 * Delta(1).Y)
        
        End If
        
        'Return the results.
        LineIntersect = iReturn
        
End Function

Public Function RadiansToDegrees(ByVal dRadians As Double) As Double

'Converts Radians to Degrees.

    RadiansToDegrees = dRadians * (180# / gdPi)
    
End Function

Public Function DegreesToRadians(ByVal dDegrees As Double) As Double

'Converts Degrees to Radians.

    DegreesToRadians = dDegrees * (gdPi / 180#)
    
End Function

