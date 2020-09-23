VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPolygonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'This is a very useful class to store polygon's Dots and
'check if the polygon is convex or not and if a point is inside
'the polygon or not
'By Raul Fragoso on 14-March-2002
'****************************************************************
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const PI As Double = 3.14159265358979
Private Polygon() As POINTAPI

' Return True if the point is in the polygon.
Public Function IsInPolygon(ByVal X As Single, ByVal Y As Single) As Boolean

    Dim pt          As Integer
    Dim total_angle As Single

    ' Get the angle between the point and the
    ' first and last Dots.
    total_angle = GetAngle(Polygon(UBound(Polygon)).X, Polygon(UBound(Polygon)).Y, X, Y, Polygon(1).X, Polygon(1).Y)

    ' Add the angles from the point to each other
    ' pair of Dots.
    For pt = 1 To UBound(Polygon) - 1
        total_angle = total_angle + GetAngle(Polygon(pt).X, Polygon(pt).Y, _
                                              X, Y, Polygon(pt + 1).X, Polygon(pt + 1).Y)
    Next pt

    ' The total angle should be 2 * PI or -2 * PI if
    ' the point is in the polygon and close to zero
    ' if the point is outside the polygon.
    IsInPolygon = (Abs(total_angle) > 0.000001)
End Function

' Return the angle ABC.
' Return a value between PI and -PI.
' Note that the value is the opposite of what you might
' expect because Y coordinates increase downward.
Private Function GetAngle(ByVal Ax As Single, ByVal Ay As Single, ByVal Bx As Single, ByVal By As Single, ByVal CX As Single, ByVal CY As Single) As Single
    
    Dim dot_product     As Single
    Dim cross_product   As Single

    dot_product = (Ax - Bx) * (CX - Bx) + (Ay - By) * (CY - By)     ' Calculate the dot product.
    cross_product = (Ax - Bx) * (CY - By) - (Ay - By) * (CX - Bx)   ' Calculate the Z coordinate of the cross product.
    GetAngle = ATan2(cross_product, dot_product)                    ' Calculate the angle.

End Function

' Return the angle with tangent opp/hyp. The returned
' value is between PI and -PI.
Private Function ATan2(ByVal opp As Single, ByVal adj As Single) As Single
    
    Dim angle As Single

    ' Get the basic angle.
    If Abs(adj) < 0.0001 Then
        angle = PI / 2
    Else
        angle = Abs(Atn(opp / adj))
    End If

    ' See if we are in quadrant 2 or 3.
    If adj < 0 Then
        ' angle > PI/2 or angle < -PI/2.
        angle = PI - angle
    End If

    ' See if we are in quadrant 3 or 4.
    If opp < 0 Then
        angle = -angle
    End If

    ' Return the result.
    ATan2 = angle
    
End Function

Public Sub AddVertex(X As Long, Y As Long)
    
    ReDim Preserve Polygon(UBound(Polygon) + 1)
    Polygon(UBound(Polygon)).X = X
    Polygon(UBound(Polygon)).Y = Y
    
End Sub

Public Function GetVertexX(idx As Long) As Long
    
    GetVertexX = Polygon(idx).X

End Function

Public Function GetVertexY(idx As Long) As Long
    
    GetVertexY = Polygon(idx).Y

End Function

Public Function VertexCount() As Long

    VertexCount = UBound(Polygon)

End Function

Public Sub ClearDots()

   ReDim Polygon(0)

End Sub

Private Sub Class_Initialize()
    
    ReDim Polygon(0)

End Sub
