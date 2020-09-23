VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOBJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Wavefront OBJ File parser
'
Dim hFile       As Long
Private Lines() As String

Public Sub WriteOBJ(FileName As String)
    
    Dim strTemp     As String
    Dim idx         As Long
    Dim idxMesh     As Integer
    Dim idxFace     As Long
    Dim TotalNumFace As Long
    Dim strX        As String
    Dim strY        As String
    Dim strZ        As String
        
    On Error GoTo err
    
    If FileExist(FileName) Then Kill FileName
    
    strTemp = "#" & vbCrLf & "# Create : EGL Point Cloud V1.0" & vbCrLf & "#" & vbCrLf
    
    hFile = FreeFile
    Open FileName For Binary As hFile

'Description
        Put #hFile, , strTemp

'Vertices
        For idx = 1 To Mesh1.NumVertices
            strX = GetVal(Mesh1.Vertices(idx).X)
            strY = GetVal(Mesh1.Vertices(idx).Y)
            strZ = GetVal(Mesh1.Vertices(idx).Z)
            
            
            
            
            strTemp = "v  " & strX & " " & strY & " " & strZ & vbCrLf
            Put #hFile, , strTemp
        Next
        strTemp = "# " & CStr(Mesh1.NumVertices) & " vertices" & vbCrLf & vbCrLf & "g Object" & vbCrLf
        Put #hFile, , strTemp
        
'Faces
        For idxMesh = 1 To Mesh1.NumMeshs
            For idxFace = 1 To Mesh1.Meshs(idxMesh).NumFaces
                With Mesh1.Meshs(idxMesh).Faces(idxFace)
                    strTemp = "f " & CStr(.A) & " " & CStr(.B) & " " & CStr(.C) & vbCrLf
                    Put #hFile, , strTemp
                    TotalNumFace = TotalNumFace + 1
                End With
            Next
        Next
        strTemp = "# " & CStr(TotalNumFace) & " faces" & vbCrLf & vbCrLf & "g " & vbCrLf
        Put #hFile, , strTemp
        
    Close #hFile
    Exit Sub
err:
    Close #hFile
End Sub

Private Function GetVal(Val As Single) As String
    
    Val = Round(Val, 6)
    GetVal = CStr(Val)
    GetVal = Replace(GetVal, ",", ".")
    
End Function

Public Sub ReadOBJ(FileName As String)
    
    Dim strData As String
    Dim idx     As Long
    Dim dx      As Single
    Dim dy      As Single
'    Dim dmax    As Single
    Dim char    As String
    
    On Error Resume Next
    
    'Reset
    LoadComplete = False
    Erase Lines
    Mesh1.NumMeshs = 0
    Mesh1.NumVertices = 0
    Erase Mesh1.Meshs
    Erase Mesh1.Vertices
    
    hFile = FreeFile
    Open FileName For Input As #hFile
        strData = Input(LOF(1) - 1, #hFile)
    Close #hFile
    Lines = Split(strData, vbLf)
    With Mesh1
        Erase Dots1.Dots
        Erase .Vertices
        .NumMeshs = 1
        ReDim .Meshs(.NumMeshs)
        
        For idx = 0 To UBound(Lines)
            char = Left(Lines(idx), 1)
            Select Case char
                Case "v"
                    .NumVertices = .NumVertices + 1
                    Dots1.NumDot = .NumVertices
                    ReDim Preserve Mesh1.Vertices(1 To .NumVertices)
                    ReDim Preserve Dots1.Dots(1 To .NumVertices)
                    .Vertices(.NumVertices) = GetVectorValue(Lines(idx))
                    Dots1.Dots(.NumVertices).Vector = .Vertices(.NumVertices)
                    Dots1.Dots(.NumVertices).Visible = True
                Case "f":
                    With .Meshs(1)
                        .NumFaces = .NumFaces + 1
                        ReDim Preserve .Faces(.NumFaces)
                        ReDim Preserve .Normals(.NumFaces)
                        .Faces(.NumFaces) = GetFaceValue(Lines(idx))
                        .Normals(.NumFaces) = _
                            CalculateNormal(Dots1.Dots(.Faces(.NumFaces).A).Vector, _
                                            Dots1.Dots(.Faces(.NumFaces).B).Vector, _
                                            Dots1.Dots(.Faces(.NumFaces).C).Vector)
                    End With
            End Select
        Next
        .Meshs(1).NormalsT = .Meshs(1).Normals
    End With
    With Dots1
        Call CalculateBox(.Dots, .Box, .Center)
        
'move center
        For idx = 1 To .NumDot
            .Dots(idx).Vector = VectorSubtract(.Dots(idx).Vector, .Center.Vector)
        Next idx
        
        For idx = 1 To 8
            .Box(idx).Vector = VectorSubtract(.Box(idx).Vector, .Center.Vector)
        Next idx
        
        .Center.Vector = VectorSet(0, 0, 0)
        
'scale screen
        dx = .Box(7).Vector.X - .Box(1).Vector.X
        dy = .Box(7).Vector.Y - .Box(1).Vector.Y
        If dx > dy Then
            MaxH = dx
        Else
            MaxH = dy
        End If
        
        Position.Sca = (cHeight / MaxH) * 0.9 ' 0.9 bigness 90%
        For idx = 1 To .NumDot
            .Dots(idx).Vector = VectorSca(.Dots(idx).Vector, Position.Sca)
        Next idx
        For idx = 1 To 8
            .Box(idx).Vector = VectorSca(.Box(idx).Vector, Position.Sca)
        Next idx
        .ClpZ = (.Box(7).Vector.Z - .Box(1).Vector.Z) \ 100
        
    End With
    Call ResetMeshParameters
    Call ResetCameraParameters
    Call ResetLightParameters
    LoadComplete = True

End Sub

Private Function GetVectorValue(Line As String) As VECTOR4

    Dim Value As String
    Dim Segments() As String
    
    Segments = Split(Line, " ")
    
'X Value
    Value = Segments(UBound(Segments) - 2)
    GetVectorValue.X = CSng(Replace(Value, ".", ","))
'Y Value
    Value = Segments(UBound(Segments) - 1)
    GetVectorValue.Y = CSng(Replace(Value, ".", ","))
'Z Value
    Value = Segments(UBound(Segments))
    GetVectorValue.Z = CSng(Replace(Value, ".", ","))
'W Value
    GetVectorValue.W = 1

End Function

Private Function GetFaceValue(Line As String) As FACE

    Dim Value As String
    Dim Segments() As String
    
    Segments = Split(Line, " ")
    
'A Value
    Value = Segments(UBound(Segments) - 2)
    GetFaceValue.A = CLng(Value)
'B Value
    Value = Segments(UBound(Segments) - 1)
    GetFaceValue.B = CLng(Value)
'C Value
    Value = Segments(UBound(Segments))
    GetFaceValue.C = CLng(Value)

End Function

