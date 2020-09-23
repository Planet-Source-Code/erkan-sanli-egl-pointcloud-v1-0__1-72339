Attribute VB_Name = "modMatrix"
Option Explicit

Public Const sPIDiv180 As Single = 0.017453!
Public Const sPIDiv360 As Single = 0.008726!
Public Const s180DivPI As Single = 57.29578!
Public Const s360DivPI As Single = 114.5916!

Public Type MATRIX ' 3x4
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
End Type

Public matOutput   As MATRIX

Public Function Out() As MATRIX
    
    Dim N As VECTOR4
    Dim U As VECTOR4
    Dim V As VECTOR4
    Dim CosX As Single
    Dim SinX As Single
    Dim CosY As Single
    Dim SinY As Single
    Dim CosZ As Single
    Dim SinZ As Single
    Dim F1 As Single
    Dim F2 As Single
    Dim F3 As Single
    Dim F4 As Single
    Dim F5 As Single
    Dim F6 As Single
    Dim F7 As Single
    Dim F8 As Single
   
    With Camera1
        N = VectorNormalize(VectorSubtract(.LookAtPoint, .WorldPosition))
        U = VectorNormalize(CrossProduct(.VUP, N))
        V = CrossProduct(N, U)
    End With
    
    With Position
        With .Rot
            CosX = Cos(.X * sPIDiv180)
            SinX = Sin(.X * sPIDiv180)
            CosY = Cos(.Y * sPIDiv180)
            SinY = Sin(.Y * sPIDiv180)
            CosZ = Cos(.Z * sPIDiv180)
            SinZ = Sin(.Z * sPIDiv180)
        End With

        F1 = CosZ * -CosY + SinZ * -SinX * -SinY
        F2 = SinZ * -CosX
        F3 = CosZ * SinY + SinZ * -SinX * -CosY
        F4 = -SinZ * -CosY + CosZ * -SinX * -SinY
        F5 = CosZ * -CosX
        F6 = -SinZ * SinY + CosZ * -SinX * -CosY
        F7 = -CosX * -SinY
        F8 = -CosX * -CosY
        
        Out.rc11 = .Sca * (F1 * U.X + F2 * U.Y + F3 * U.Z)
        Out.rc12 = .Sca * (F4 * U.X + F5 * U.Y + F6 * U.Z)
        Out.rc13 = .Sca * (F7 * U.X + SinX * U.Y + F8 * U.Z)
        Out.rc14 = .Tra.X * U.X + .Tra.Y * U.Y + .Tra.Z * U.Z

        Out.rc21 = .Sca * (F1 * V.X + F2 * V.Y + F3 * V.Z)
        Out.rc22 = .Sca * (F4 * V.X + F5 * V.Y + F6 * V.Z)
        Out.rc23 = .Sca * (F7 * V.X + SinX * V.Y + F8 * V.Z)
        Out.rc24 = .Tra.X * V.X + .Tra.Y * V.Y + .Tra.Z * V.Z

        Out.rc31 = .Sca * (F1 * N.X + F2 * N.Y + F3 * N.Z)
        Out.rc32 = .Sca * (F4 * N.X + F5 * N.Y + F6 * N.Z)
        Out.rc33 = .Sca * (F7 * N.X + SinX * N.Y + F8 * N.Z)
        Out.rc34 = .Tra.X * N.X + .Tra.Y * N.Y + .Tra.Z * N.Z

    End With

End Function

'Public Function ConvertFOVtoZoom(FOV As Single) As Single
'
'    ConvertFOVtoZoom = 1 / Tan(FOV * sPIDiv360)
'
'End Function
'
'Public Function ConvertZoomtoFOV(Zoom As Single) As Single
'
'    ConvertZoomtoFOV = s360DivPI * Atn(1 / Zoom)
'
'End Function

Public Function MatrixMultiplyVector(m1 As MATRIX, V1 As VECTOR4) As VECTOR4
            
    MatrixMultiplyVector.X = (m1.rc11 * V1.X) + (m1.rc12 * V1.Y) + (m1.rc13 * V1.Z) + (m1.rc14 * V1.W)
    MatrixMultiplyVector.Y = (m1.rc21 * V1.X) + (m1.rc22 * V1.Y) + (m1.rc23 * V1.Z) + (m1.rc24 * V1.W)
    MatrixMultiplyVector.Z = (m1.rc31 * V1.X) + (m1.rc32 * V1.Y) + (m1.rc33 * V1.Z) + (m1.rc34 * V1.W)
    MatrixMultiplyVector.W = V1.W
    
End Function

Public Function VectorSet(X As Single, Y As Single, Z As Single) As VECTOR4

    VectorSet.X = X
    VectorSet.Y = Y
    VectorSet.Z = Z
    VectorSet.W = 1
    
End Function

Public Function VectorSubtract(V1 As VECTOR4, V2 As VECTOR4) As VECTOR4

    VectorSubtract.X = V1.X - V2.X
    VectorSubtract.Y = V1.Y - V2.Y
    VectorSubtract.Z = V1.Z - V2.Z
    VectorSubtract.W = 1
End Function

Public Function VectorMultiply(V1 As VECTOR4, V2 As VECTOR4) As VECTOR4

    VectorMultiply.X = V1.X * V2.X
    VectorMultiply.Y = V1.Y * V2.Y
    VectorMultiply.Z = V1.Z * V2.Z
    VectorMultiply.W = 1
End Function

Public Function VectorAddition(V1 As VECTOR4, V2 As VECTOR4) As VECTOR4

    VectorAddition.X = V1.X + V2.X
    VectorAddition.Y = V1.Y + V2.Y
    VectorAddition.Z = V1.Z + V2.Z
    
End Function

Public Function VectorSca(V As VECTOR4, s As Single) As VECTOR4

    VectorSca.X = V.X * s
    VectorSca.Y = V.Y * s
    VectorSca.Z = V.Z * s
    VectorSca.W = 1

End Function

Public Function VectorNormalize(V As VECTOR4) As VECTOR4
    
    Dim sngLength As Single
    
    sngLength = VectorLength(V)
    If sngLength = 0 Then sngLength = 1
    VectorNormalize.X = V.X / sngLength
    VectorNormalize.Y = V.Y / sngLength
    VectorNormalize.Z = V.Z / sngLength
    
End Function
Public Function VectorLength(V As VECTOR4) As Single
    
    VectorLength = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)

End Function

Public Function DotProduct(V1 As VECTOR4, V2 As VECTOR4) As Single

    DotProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z '+ V1.W * V2.W
    
End Function

Public Function CrossProduct(V1 As VECTOR4, V2 As VECTOR4) As VECTOR4
    
    CrossProduct.X = V1.Y * V2.Z - V1.Z * V2.Y
    CrossProduct.Y = V1.Z * V2.X - V1.X * V2.Z
    CrossProduct.Z = V1.X * V2.Y - V1.Y * V2.X

End Function

Public Function CalculateNormal(vecA As VECTOR4, vecB As VECTOR4, vecC As VECTOR4) As VECTOR4

    CalculateNormal = VectorNormalize( _
                        CrossProduct( _
                            VectorSubtract(vecB, vecA), _
                            VectorSubtract(vecB, vecC)))

End Function


