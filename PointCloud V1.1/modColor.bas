Attribute VB_Name = "modColor"
Option Explicit

Public ColorRGB      As RGBQUAD
Public ColorLong     As Long

Public Function ColorSet(Red As Byte, Green As Byte, Blue As Byte) As RGBQUAD

    ColorSet.rgbRed = Red
    ColorSet.rgbGreen = Green
    ColorSet.rgbBlue = Blue

End Function

Public Function ColorRGBToLong(C As RGBQUAD) As Long

    ColorRGBToLong = RGB(C.rgbRed, C.rgbGreen, C.rgbBlue)
    
End Function


Public Function ColorLimits(ByVal iColor As Integer) As Byte

    If iColor < 0 Then
        ColorLimits = 0
    ElseIf iColor > 255 Then
        ColorLimits = 255
    Else
        ColorLimits = CByte(iColor)
    End If
    
End Function

Public Function ColorSca(C As RGBQUAD, S As Single) As RGBQUAD

    ColorSca.rgbRed = ColorLimits(C.rgbRed * S)
    ColorSca.rgbGreen = ColorLimits(C.rgbGreen * S)
    ColorSca.rgbBlue = ColorLimits(C.rgbBlue * S)

End Function

Public Function ColorShade(idxMesh As Integer, idxFace As Long) As Single

    ColorShade = Abs((DotProduct(Mesh1.Meshs(idxMesh).NormalsT(idxFace), Light1.Normal))) * InvScl
    
End Function

Public Function ColorPlus(C1 As RGBQUAD, V As Single) As RGBQUAD

    ColorPlus.rgbRed = ColorLimits(C1.rgbRed + V)
    ColorPlus.rgbGreen = ColorLimits(C1.rgbGreen + V)
    ColorPlus.rgbBlue = ColorLimits(C1.rgbBlue + V)

End Function

