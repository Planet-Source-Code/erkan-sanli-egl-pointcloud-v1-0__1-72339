Attribute VB_Name = "modGradientAPI"
Option Explicit

'Const GRADIENT_FILL_RECT_H      As Long = &H0
Const GRADIENT_FILL_RECT_V      As Long = &H1
Const GRADIENT_FILL_TRIANGLE    As Long = &H2

Private Type TRIVERTEX
    X           As Long
    Y           As Long
    Red         As Integer
    Green       As Integer
    Blue        As Integer
    Alpha       As Integer
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1     As Long
    Vertex2     As Long
    Vertex3     As Long
End Type

Private Type GRADIENT_RECT
    UpperLeft   As Long
    LowerRight  As Long
End Type


Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" _
    (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
    pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
    
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
    (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
    pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Public triVert() As TRIVERTEX

Public Sub DrawGradientTriangle(idxMesh As Integer, idxFace As Long)

    Dim vert(2) As TRIVERTEX
    Dim gTri    As GRADIENT_TRIANGLE

    With Mesh1.Meshs(idxMesh).Faces(idxFace)
        If triVert(.A).Alpha = 0 Then
            triVert(.A).X = Dots1.Dots(.A).Screen.X
            triVert(.A).Y = Dots1.Dots(.A).Screen.Y
            triVert(.A).Red = ColorRGB.rgbRed
            triVert(.A).Green = ColorRGB.rgbGreen
            triVert(.A).Blue = ColorRGB.rgbBlue
            triVert(.A).Alpha = 1
        End If
        If triVert(.B).Alpha = 0 Then
            triVert(.B).X = Dots1.Dots(.B).Screen.X
            triVert(.B).Y = Dots1.Dots(.B).Screen.Y
            triVert(.B).Red = ColorRGB.rgbRed
            triVert(.B).Green = ColorRGB.rgbGreen
            triVert(.B).Blue = ColorRGB.rgbBlue
            triVert(.B).Alpha = 1
        End If
        If triVert(.C).Alpha = 0 Then
            triVert(.C).X = Dots1.Dots(.C).Screen.X
            triVert(.C).Y = Dots1.Dots(.C).Screen.Y
            triVert(.C).Red = ColorRGB.rgbRed
            triVert(.C).Green = ColorRGB.rgbGreen
            triVert(.C).Blue = ColorRGB.rgbBlue
            triVert(.C).Alpha = 1
        End If

        vert(0).X = triVert(.A).X
        vert(0).Y = triVert(.A).Y
        vert(0).Red = Val("&h" & Hex(triVert(.A).Red) & "00")
        vert(0).Green = Val("&h" & Hex(triVert(.A).Green) & "00")
        vert(0).Blue = Val("&h" & Hex(triVert(.A).Blue) & "00")
    
        vert(1).X = triVert(.B).X
        vert(1).Y = triVert(.B).Y
        vert(1).Red = Val("&h" & Hex(triVert(.B).Red) & "00")
        vert(1).Green = Val("&h" & Hex(triVert(.B).Green) & "00")
        vert(1).Blue = Val("&h" & Hex(triVert(.B).Blue) & "00")
    
        vert(2).X = triVert(.C).X
        vert(2).Y = triVert(.C).Y
        vert(2).Red = Val("&h" & Hex(triVert(.C).Red) & "00")
        vert(2).Green = Val("&h" & Hex(triVert(.C).Green) & "00")
        vert(2).Blue = Val("&h" & Hex(triVert(.C).Blue) & "00")
    End With
    
    gTri.Vertex1 = 0
    gTri.Vertex2 = 1
    gTri.Vertex3 = 2
    
    Call GradientFillTriangle(CanBuffer.hDC, vert(0), 3, gTri, 1, GRADIENT_FILL_TRIANGLE)

End Sub

Public Sub DrawGradientRectangle()

    Dim vert(1) As TRIVERTEX
    Dim gRect   As GRADIENT_RECT
    
    With vert(0)
        .X = 0
        .Y = 0
        .Red = Val("&h" & Hex(Color.rgbBack1.rgbRed) & "00")
        .Green = Val("&h" & Hex(Color.rgbBack1.rgbGreen) & "00")
        .Blue = Val("&h" & Hex(Color.rgbBack1.rgbBlue) & "00")
        .Alpha = 0&
    End With

    With vert(1)
        .X = cWidth
        .Y = cHeight
        .Red = Val("&h" & Hex(Color.rgbBack2.rgbRed) & "00")
        .Green = Val("&h" & Hex(Color.rgbBack2.rgbGreen) & "00")
        .Blue = Val("&h" & Hex(Color.rgbBack2.rgbBlue) & "00")
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    Call GradientFillRect(BackBuffer.hDC, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)

End Sub

