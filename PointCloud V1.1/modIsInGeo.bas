Attribute VB_Name = "modIsInGeom"
Option Explicit

Public Function IsInCircle(XP As Long, YP As Long, _
                           X1 As Long, Y1 As Long, _
                           X2 As Long, Y2 As Long, _
                           X3 As Long, Y3 As Long) As Boolean

'Return TRUE if the point (xp,yp) lies inside the circumcircle
'made up by points (x1,y1) (x2,y2) (x3,y3)
'The circumcircle centre is returned in (xc,yc) and the radius r
'NOTE: A point on the edge is inside the circumcircle
     
    Dim eps     As Double
    Dim m1      As Double
    Dim m2      As Double
    Dim mx1     As Double
    Dim mx2     As Double
    Dim my1     As Double
    Dim my2     As Double
    Dim dx      As Double
    Dim dy      As Double
    Dim rsqr    As Double
    Dim drsqr   As Double
    Dim CenterX As Double
    Dim CenterY As Double

    eps = 0.0001
    IsInCircle = False
      
    If Abs(Y1 - Y2) < eps And Abs(Y2 - Y3) < eps Then Exit Function
    
    If Abs(Y2 - Y1) < eps Then
        m2 = -(X3 - X2) / (Y3 - Y2)
        mx2 = (X2 + X3) * 0.5
        my2 = (Y2 + Y3) * 0.5
        CenterX = (X2 + X1) * 0.5
        CenterY = m2 * (CenterX - mx2) + my2
    ElseIf Abs(Y3 - Y2) < eps Then
        m1 = -(X2 - X1) / (Y2 - Y1)
        mx1 = (X1 + X2) * 0.5
        my1 = (Y1 + Y2) * 0.5
        CenterX = (X3 + X2) * 0.5
        CenterY = m1 * (CenterX - mx1) + my1
    Else
        m1 = -(X2 - X1) / (Y2 - Y1)
        m2 = -(X3 - X2) / (Y3 - Y2)
        mx1 = (X1 + X2) * 0.5
        mx2 = (X2 + X3) * 0.5
        my1 = (Y1 + Y2) * 0.5
        my2 = (Y2 + Y3) * 0.5
        CenterX = (m1 * mx1 - m2 * mx2 + my2 - my1) / (m1 - m2)
        CenterY = m1 * (CenterX - mx1) + my1
    End If
    dx = X2 - CenterX
    dy = Y2 - CenterY
    rsqr = dx * dx + dy * dy
    dx = XP - CenterX
    dy = YP - CenterY
    drsqr = dx * dx + dy * dy
    If drsqr <= rsqr Then IsInCircle = True
        
End Function

Public Function IsInTriangle(XP As Long, YP As Long, _
                             X1 As Long, Y1 As Long, _
                             X2 As Long, Y2 As Long, _
                             X3 As Long, Y3 As Long) As Boolean

    Dim Val1    As Single
    Dim Val2    As Single
    Dim Val3    As Single
    
    Val1 = (X1 - XP) * (Y2 - YP) - (X2 - XP) * (Y1 - YP)
    Val2 = (X2 - XP) * (Y3 - YP) - (X3 - XP) * (Y2 - YP)
    Val3 = (X3 - XP) * (Y1 - YP) - (X1 - XP) * (Y3 - YP)
    
    If (Val1 > 0 And Val2 > 0 And Val3 > 0) Or _
       (Val1 < 0 And Val2 < 0 And Val3 < 0) Then IsInTriangle = True

End Function

Public Function IsInCanvas(abc As FACE) As Boolean

    Dim FaceA As Boolean
    Dim FaceB As Boolean
    Dim FaceC As Boolean
    
    With abc
        If Dots1.Dots(.A).Screen.X > 0 And Dots1.Dots(.A).Screen.X < cWidth And _
           Dots1.Dots(.A).Screen.Y > 0 And Dots1.Dots(.A).Screen.Y < cHeight Then _
            FaceA = True
        If Dots1.Dots(.B).Screen.X > 0 And Dots1.Dots(.B).Screen.X < cWidth And _
           Dots1.Dots(.B).Screen.Y > 0 And Dots1.Dots(.B).Screen.Y < cHeight Then _
            FaceB = True
        If Dots1.Dots(.C).Screen.X > 0 And Dots1.Dots(.C).Screen.X < cWidth And _
           Dots1.Dots(.C).Screen.Y > 0 And Dots1.Dots(.C).Screen.Y < cHeight Then _
            FaceC = True
        If FaceA Or FaceB Or FaceC Then IsInCanvas = True
    End With

End Function
