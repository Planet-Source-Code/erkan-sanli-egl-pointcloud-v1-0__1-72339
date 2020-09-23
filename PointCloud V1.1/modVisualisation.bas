Attribute VB_Name = "modVisualisation"
Option Explicit
Option Base 1

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Sub Render(hDC As Long)
    
    DoEvents
    Call BackBuffer.Paint(CanBuffer.hDC)

    Select Case VStyle
        Case Wireframe: Call RenderWireframe
        Case Facet:     Call RenderFacet
        Case Smooth:    Call RenderSmooth
    End Select
    
    If ShowMeshBorder Then Call DrawBorder
    If ShowBox Then Call DrawBox
    If ShowHideDot Then Call RenderObjDots
    
    Select Case SelType
        Case Rectangular:   Call DrawRect
        Case Polygonal:     Call DrawPolygon
    End Select
    
    Call CanBuffer.Paint(hDC)

End Sub

Public Sub RenderObjDots()
    
    Dim idx As Long
    Dim X1 As Single
    Dim Y1 As Single
    
    With Dots1
        For idx = 1 To .NumDot
            If .Dots(idx).Visible Then
                X1 = .Dots(idx).Screen.X
                Y1 = .Dots(idx).Screen.Y
                If X1 > 0 And X1 < cWidth - 1 And Y1 > 0 And Y1 < cHeight - 1 Then
                    If ClipFar Then
                        If .Center.VectorT.Z < .Dots(idx).VectorT.Z Then Call DrawDot(X1, Y1, .Dots(idx).Selected)
                    Else
                        Call DrawDot(X1, Y1, .Dots(idx).Selected)
                    End If
                End If
            End If
        Next idx
    End With

End Sub

'Public Sub RenderDot(Dots() As DOT)
'
'    Dim idx As Long
'    Dim X1 As Single
'    Dim Y1 As Single
'
'    For idx = 0 To UBound(Dots)
'        If Dots(idx).Visible Then
'            X1 = Dots(idx).Screen.X
'            Y1 = Dots(idx).Screen.Y
'            If X1 > 0 And X1 < cWidth - 1 And Y1 > 0 And Y1 < cHeight - 1 Then
'                If ClipFar Then
'                    If Dots1.Center.VectorT.Z < Dots(idx).VectorT.Z Then Call DrawDot(X1, Y1, Dots(idx).Selected)
'                Else
'                    Call DrawDot(X1, Y1, Dots(idx).Selected)
'                End If
'            End If
'        End If
'    Next idx
'
'End Sub

Private Sub RenderWireframe()

    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim PenSelect   As Long
    Dim pAPI        As POINTAPI
    
    For idxMesh = 1 To Mesh1.NumMeshs
        For idxFace = 1 To Mesh1.Meshs(idxMesh).NumFaces
            With Mesh1.Meshs(idxMesh).Faces(idxFace)
                If EditMF > DeleteMesh Then
                    If IsInTriangle(CLng(LastX), CLng(LastY), _
                                    Dots1.Dots(.A).Screen.X, Dots1.Dots(.A).Screen.Y, _
                                    Dots1.Dots(.B).Screen.X, Dots1.Dots(.B).Screen.Y, _
                                    Dots1.Dots(.C).Screen.X, Dots1.Dots(.C).Screen.Y) Then
                        PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, ColorRGBToLong(Color.rgbSelFace)))
                        SelectedMeshIndex = idxMesh
                        SelectedFaceIndex = idxFace
                    Else
                        PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, Color.lWireframe))
                    End If
                Else
                    PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, Color.lWireframe))
                End If
                MoveToEx CanBuffer.hDC, Dots1.Dots(.A).Screen.X, Dots1.Dots(.A).Screen.Y, pAPI
                LineTo CanBuffer.hDC, Dots1.Dots(.B).Screen.X, Dots1.Dots(.B).Screen.Y
                LineTo CanBuffer.hDC, Dots1.Dots(.C).Screen.X, Dots1.Dots(.C).Screen.Y
                LineTo CanBuffer.hDC, Dots1.Dots(.A).Screen.X, Dots1.Dots(.A).Screen.Y
            End With
            Call DeleteObject(PenSelect)
        Next idxFace
    Next

End Sub

Private Sub RenderFacet()

    Dim idxMesh         As Integer
    Dim idxFace         As Long
    Dim idxFaceV        As Long
    Dim PenSelect       As Long
    Dim BrushSelect     As Long
    Dim pAPI(1 To 3)    As POINTAPI

    If SortVisibleFaces < 1 Then Exit Sub
    
    For idxFace = 1 To UBound(Mesh1.FaceV)
        idxFaceV = Mesh1.FaceV(idxFace).idxFace
        idxMesh = Mesh1.FaceV(idxFace).idxMesh
        If IsInCanvas(Mesh1.Meshs(idxMesh).Faces(idxFaceV)) Then
            With Mesh1.Meshs(idxMesh).Faces(idxFaceV)
                If EditMF > DeleteMesh Then
                    If IsInTriangle(CLng(LastX), CLng(LastY), _
                                    Dots1.Dots(.A).Screen.X, Dots1.Dots(.A).Screen.Y, _
                                    Dots1.Dots(.B).Screen.X, Dots1.Dots(.B).Screen.Y, _
                                    Dots1.Dots(.C).Screen.X, Dots1.Dots(.C).Screen.Y) Then
                        ColorRGB = Color.rgbSelFace
                        SelectedMeshIndex = idxMesh
                        SelectedFaceIndex = idxFaceV
                    Else
                        ColorRGB = Mesh1.Meshs(idxMesh).Faces(idxFaceV).Color
                        ColorRGB = ColorSca(ColorRGB, ColorShade(idxMesh, idxFaceV))
                        PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, ColorRGBToLong(ColorPlus(ColorRGB, -20))))
                    End If
                Else
                    ColorRGB = Mesh1.Meshs(idxMesh).Faces(idxFaceV).Color
                    ColorRGB = ColorSca(ColorRGB, ColorShade(idxMesh, idxFaceV))
                    PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, ColorRGBToLong(ColorPlus(ColorRGB, -20))))
                End If
                ColorLong = ColorRGBToLong(ColorRGB)
                BrushSelect = SelectObject(CanBuffer.hDC, CreateSolidBrush(ColorLong))
                pAPI(1) = Dots1.Dots(.A).Screen
                pAPI(2) = Dots1.Dots(.B).Screen
                pAPI(3) = Dots1.Dots(.C).Screen
            End With
            Call Polygon(CanBuffer.hDC, pAPI(1), 3)
            DeleteObject BrushSelect
            DeleteObject PenSelect
        End If
    Next

End Sub

Private Sub RenderSmooth()

    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim idxFaceV    As Long

    If SortVisibleFaces < 1 Then Exit Sub
    ReDim triVert(UBound(Mesh1.FaceV) * 3)   ' 3= A,B,C of face
    For idxFace = 1 To UBound(Mesh1.FaceV)
        idxFaceV = Mesh1.FaceV(idxFace).idxFace
        idxMesh = Mesh1.FaceV(idxFace).idxMesh
        If IsInCanvas(Mesh1.Meshs(idxMesh).Faces(idxFaceV)) Then
            ColorRGB = Mesh1.Meshs(idxMesh).Faces(idxFaceV).Color
            ColorRGB = ColorSca(ColorRGB, ColorShade(idxMesh, idxFaceV))
            Call DrawGradientTriangle(idxMesh, idxFaceV)
        End If
    Next
    Erase triVert

End Sub

Private Sub DrawDot(X As Single, Y As Single, Selected As Boolean)
        
    ColorLong = IIf(Selected, Color.lSelDots, Color.lObjDots)
    CanBuffer.SetPixel X, Y, ColorLong
    If BigDot Then
        CanBuffer.SetPixel X + 1, Y, ColorLong
        CanBuffer.SetPixel X, Y + 1, ColorLong
        CanBuffer.SetPixel X + 1, Y + 1, ColorLong
    End If

End Sub


Public Sub DrawRect()

'           +Y
'           ^
'           3--------2
'           |        |
'           |        |
'           |        |
'           0--------1 > +X
    
    Dim p As POINTAPI
    Dim PenSelect As Long
    
    PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, Color.lSelGeo))
    MoveToEx CanBuffer.hDC, Rect1.X, Rect1.Y, p
    LineTo CanBuffer.hDC, Rect2.X, Rect1.Y
    LineTo CanBuffer.hDC, Rect2.X, Rect2.Y
    LineTo CanBuffer.hDC, Rect1.X, Rect2.Y
    LineTo CanBuffer.hDC, Rect1.X, Rect1.Y
    Call DeleteObject(PenSelect)

End Sub

Public Sub DrawPolygon()

'            _______
'           /       |
'          /        |
'         /      ___|
'        |      |
'        |______|
    
    Dim p As POINTAPI
    Dim PenSelect As Long
    Dim idx As Long

    With Geometry
        If .VertexCount > 0 Then
            PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, Color.lSelGeo))
            MoveToEx CanBuffer.hDC, .GetVertexX(1), .GetVertexY(1), p
            For idx = 2 To .VertexCount
                LineTo CanBuffer.hDC, .GetVertexX(idx), .GetVertexY(idx)
            Next
            If StartPolygon Then
                LineTo CanBuffer.hDC, CLng(LastX), CLng(LastY)
            Else
                LineTo CanBuffer.hDC, .GetVertexX(1), .GetVertexY(1)
            End If
            Call DeleteObject(PenSelect)
        End If
    End With
    
End Sub

Public Sub DrawBox()

'           +Y
'           ^
'           4--------3
'          /|       /|
'         / |      / |
'        8--+-----7  |
'        |  |     |  |
'        |  1-----+--2 >+X
'        | /      | /
'        |/       |/
'        5--------6
'       /
'      +Z

    Dim p           As POINTAPI
    Dim PenSelect   As Long
    Dim X1          As Single
    Dim Y1          As Single
    
    PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, Color.lBox))
    With Dots1
        MoveToEx CanBuffer.hDC, .Box(1).Screen.X, .Box(1).Screen.Y, p
        LineTo CanBuffer.hDC, .Box(2).Screen.X, .Box(2).Screen.Y
        LineTo CanBuffer.hDC, .Box(3).Screen.X, .Box(3).Screen.Y
        LineTo CanBuffer.hDC, .Box(4).Screen.X, .Box(4).Screen.Y
        LineTo CanBuffer.hDC, .Box(1).Screen.X, .Box(1).Screen.Y
        LineTo CanBuffer.hDC, .Box(5).Screen.X, .Box(5).Screen.Y
        LineTo CanBuffer.hDC, .Box(6).Screen.X, .Box(6).Screen.Y
        LineTo CanBuffer.hDC, .Box(7).Screen.X, .Box(7).Screen.Y
        LineTo CanBuffer.hDC, .Box(8).Screen.X, .Box(8).Screen.Y
        LineTo CanBuffer.hDC, .Box(5).Screen.X, .Box(5).Screen.Y
        MoveToEx CanBuffer.hDC, .Box(2).Screen.X, .Box(2).Screen.Y, p
        LineTo CanBuffer.hDC, .Box(6).Screen.X, .Box(6).Screen.Y
        MoveToEx CanBuffer.hDC, .Box(3).Screen.X, .Box(3).Screen.Y, p
        LineTo CanBuffer.hDC, .Box(7).Screen.X, .Box(7).Screen.Y
        MoveToEx CanBuffer.hDC, .Box(4).Screen.X, .Box(4).Screen.Y, p
        LineTo CanBuffer.hDC, .Box(8).Screen.X, .Box(8).Screen.Y
        Call DeleteObject(PenSelect)
    End With
    
End Sub

Public Sub CalculateBox(Dots() As DOT, Box() As DOT, Center As DOT)

'           +Y
'           ^
'           4--------3
'          /|       /|
'         / |      / |
'        8--+-----7  |
'        |  |     |  |
'        |  1-----+--2 >+X
'        | /      | /
'        |/       |/
'        5--------6
'       /
'      +Z

    Dim idx         As Long
    Dim MaxVertex   As VECTOR4
    Dim MinVertex   As VECTOR4
    
    With Mesh1
        MaxVertex = Dots(1).Vector
        MinVertex = Dots(1).Vector
        MaxVertex.W = 1
        MinVertex.W = 1
        For idx = 2 To UBound(Dots)
            If MaxVertex.X < Dots(idx).Vector.X Then MaxVertex.X = Dots(idx).Vector.X
            If MaxVertex.Y < Dots(idx).Vector.Y Then MaxVertex.Y = Dots(idx).Vector.Y
            If MaxVertex.Z < Dots(idx).Vector.Z Then MaxVertex.Z = Dots(idx).Vector.Z
            If MinVertex.X > Dots(idx).Vector.X Then MinVertex.X = Dots(idx).Vector.X
            If MinVertex.Y > Dots(idx).Vector.Y Then MinVertex.Y = Dots(idx).Vector.Y
            If MinVertex.Z > Dots(idx).Vector.Z Then MinVertex.Z = Dots(idx).Vector.Z
        Next
        Box(1).Vector = MinVertex
        Box(2).Vector = MinVertex:  Box(2).Vector.X = MaxVertex.X
        Box(3).Vector = MaxVertex:  Box(3).Vector.Z = MinVertex.Z
        Box(4).Vector = MinVertex:  Box(4).Vector.Y = MaxVertex.Y
        Box(5).Vector = MinVertex:  Box(5).Vector.Z = MaxVertex.Z
        Box(6).Vector = MaxVertex:  Box(6).Vector.Y = MinVertex.Y
        Box(7).Vector = MaxVertex
        Box(8).Vector = MaxVertex:  Box(8).Vector.X = MinVertex.X
        Center.Vector.X = (MaxVertex.X + MinVertex.X) / 2
        Center.Vector.Y = (MaxVertex.Y + MinVertex.Y) / 2
        Center.Vector.Z = (MaxVertex.Z + MinVertex.Z) / 2
        Center.Vector.W = 1
    End With
        
End Sub

Private Sub DrawBorder()

    Dim pAPI        As POINTAPI
    Dim PenSelect   As Long
    Dim idx         As Long
    
    If Mesh1.NumMeshs = 0 Then Exit Sub
    With Mesh1.Meshs(Mesh1.NumMeshs)
        PenSelect = SelectObject(CanBuffer.hDC, CreatePen(0, 1, vbYellow))
        For idx = 1 To UBound(.BorderEdges)
            If .BorderEdges(idx).Used = 0 Then
                MoveToEx CanBuffer.hDC, Dots1.Dots(.BorderEdges(idx).Start).Screen.X, Dots1.Dots(.BorderEdges(idx).Start).Screen.Y, pAPI
                LineTo CanBuffer.hDC, Dots1.Dots(.BorderEdges(idx).End).Screen.X, Dots1.Dots(.BorderEdges(idx).End).Screen.Y
            End If
        Next
    End With
    Call DeleteObject(PenSelect)

End Sub

Public Function Ratio(ByVal Value1 As Single, ByVal Value2 As Single) As Single

    If Value2 = 0 Then
        Ratio = 0
    Else
        Ratio = Value1 / Value2
    End If

End Function

