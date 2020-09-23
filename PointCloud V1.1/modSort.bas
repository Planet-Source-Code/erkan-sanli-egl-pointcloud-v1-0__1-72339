Attribute VB_Name = "modSort"
Option Explicit
Option Base 1

Public Function SortVisibleFaces() As Long
    
    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim idxFaceV    As Long
    Dim addFace     As Boolean
        
    With Mesh1
        idxFaceV = 1
        Erase .FaceV
        If ShowBackFace Then
            For idxMesh = 1 To .NumMeshs
                For idxFace = 1 To .Meshs(idxMesh).NumFaces
                    If .Meshs(idxMesh).NormalsT(idxFace).Z < 0 Then
                        .Meshs(idxMesh).Faces(idxFace).Color = Color.rgbFaces
                    Else
                        .Meshs(idxMesh).Faces(idxFace).Color = Color.rgbBackFace
                    End If
                    Call AddFaceV(idxMesh, idxFace, idxFaceV)
                    idxFaceV = idxFaceV + 1
                Next
            Next
        Else
            For idxMesh = 1 To .NumMeshs
                For idxFace = 1 To .Meshs(idxMesh).NumFaces
                    If .Meshs(idxMesh).NormalsT(idxFace).Z < 0 Then
                        .Meshs(idxMesh).Faces(idxFace).Color = Color.rgbFaces
                        Call AddFaceV(idxMesh, idxFace, idxFaceV)
                        idxFaceV = idxFaceV + 1
                    End If
                Next
            Next
        End If
    End With
    idxFaceV = idxFaceV - 1
    If idxFaceV > 1 Then SortFaces 1, idxFaceV
    SortVisibleFaces = idxFaceV
    
End Function

Private Sub SortFaces(ByVal First As Long, ByVal Last As Long)

    Dim FirstIdx  As Long
    Dim MidIdx As Long
    Dim LastIdx  As Long
    Dim MidVal As Single
    Dim TempOrder  As ORDER

    If (First < Last) Then
        With Mesh1
            MidIdx = (First + Last) \ 2
            MidVal = .FaceV(MidIdx).ZValue
            FirstIdx = First
            LastIdx = Last
            Do
                Do While .FaceV(FirstIdx).ZValue < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While .FaceV(LastIdx).ZValue > MidVal
                    LastIdx = LastIdx - 1
                Loop
                If (FirstIdx <= LastIdx) Then
                    TempOrder = .FaceV(LastIdx)
                    .FaceV(LastIdx) = .FaceV(FirstIdx)
                    .FaceV(FirstIdx) = TempOrder
                    FirstIdx = FirstIdx + 1
                    LastIdx = LastIdx - 1
                End If
            Loop Until FirstIdx > LastIdx

            If (LastIdx <= MidIdx) Then
                SortFaces First, LastIdx
                SortFaces FirstIdx, Last
            Else
                SortFaces FirstIdx, Last
                SortFaces First, LastIdx
            End If
        End With
    End If

End Sub

Private Sub AddFaceV(idxMesh As Integer, idxFace As Long, idxFaceV As Long)
    
    With Mesh1
        ReDim Preserve .FaceV(idxFaceV)
        .FaceV(idxFaceV).idxMesh = idxMesh
        .FaceV(idxFaceV).idxFace = idxFace
        .FaceV(idxFaceV).ZValue = _
            Dots1.Dots(.Meshs(idxMesh).Faces(idxFace).A).VectorT.Z + _
            Dots1.Dots(.Meshs(idxMesh).Faces(idxFace).B).VectorT.Z + _
            Dots1.Dots(.Meshs(idxMesh).Faces(idxFace).C).VectorT.Z
    End With
                        

End Sub
