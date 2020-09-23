Attribute VB_Name = "modDelaunay"
'Credit to Paul Bourke (pbourke@swin.edu.au) for the original Fortran 77 Program :))
'Conversion by EluZioN (EluZioN@casesladder.com)
'Revision by Erkan Sanli (July 2009)

Option Explicit
Option Base 1

Dim Faces()     As FACE

Public Sub Triangulate()
    
'Takes as input NumVert Dots in arrays screen()
'Returned is a list of NumTri triangular faces in the array
'Faces(). These Facess are arranged in clockwise order.

'    Dim NumVert                 As Long
    Dim NumTri                  As Integer
    Dim NumEdge                 As Long
    Dim idxVert                 As Integer
    Dim idxEdge                 As Integer
    Dim idxEdge2                As Integer
    Dim Edges()                 As EDGE
    Dim xmin                    As Long
    Dim xmax                    As Long
    Dim ymin                    As Long
    Dim ymax                    As Long
    Dim dx                      As Double
    Dim dy                      As Double
    Dim dmax                    As Double
    Dim xmid                    As Long
    Dim ymid                    As Long
    Dim idx                     As Long
    Dim idx2                    As Long
    
    On Error Resume Next
    
    With Dots1
        ReDim Preserve .SelDots(.NumSelDot + 3)
        ReDim Faces(.NumSelDot * 4)
        ReDim Edges(UBound(Faces) * 3)
        
        xmin = .SelDots(1).Screen.X
        ymin = .SelDots(1).Screen.Y
        xmax = xmin
        ymax = ymin
        For idxVert = 2 To .NumSelDot
            If .SelDots(idxVert).Screen.X < xmin Then xmin = .SelDots(idxVert).Screen.X
            If .SelDots(idxVert).Screen.X > xmax Then xmax = .SelDots(idxVert).Screen.X
            If .SelDots(idxVert).Screen.Y < ymin Then ymin = .SelDots(idxVert).Screen.Y
            If .SelDots(idxVert).Screen.Y > ymax Then ymax = .SelDots(idxVert).Screen.Y
        Next idxVert
        dx = xmax - xmin
        dy = ymax - ymin
        If dx > dy Then
            dmax = dx
        Else
            dmax = dy
        End If
        xmid = (xmax + xmin) * 0.5
        ymid = (ymax + ymin) * 0.5
'________________________________________________________________________

'Set up the Super Face
'This is a Faces which encompasses all the sample points.
'The superFaces coordinates are added to the end of the
'screen list. The SuperFace is the first Face in
'the Faces list.
        NumTri = 1
        .SelDots(.NumSelDot + 1).Screen.X = xmid - 2 * dmax
        .SelDots(.NumSelDot + 1).Screen.Y = ymid - dmax
        .SelDots(.NumSelDot + 2).Screen.X = xmid
        .SelDots(.NumSelDot + 2).Screen.Y = ymid + 2 * dmax
        .SelDots(.NumSelDot + 3).Screen.X = xmid + 2 * dmax
        .SelDots(.NumSelDot + 3).Screen.Y = ymid - dmax
        Faces(NumTri).A = .NumSelDot + 1
        Faces(NumTri).B = .NumSelDot + 2
        Faces(NumTri).C = .NumSelDot + 3

'Include each point one at a time into the existing mesh
        For idxVert = 1 To .NumSelDot
            'Set up the edge buffer.If the point lies inside the circumcircle
            'then the three edges of that Faces are added to the edge buffer.
            NumEdge = 0
            idxEdge = 0
            Do
                idxEdge = idxEdge + 1
                If IsInCircle( _
                            .SelDots(idxVert).Screen.X, .SelDots(idxVert).Screen.Y, _
                            .SelDots(Faces(idxEdge).A).Screen.X, .SelDots(Faces(idxEdge).A).Screen.Y, _
                            .SelDots(Faces(idxEdge).B).Screen.X, .SelDots(Faces(idxEdge).B).Screen.Y, _
                            .SelDots(Faces(idxEdge).C).Screen.X, .SelDots(Faces(idxEdge).C).Screen.Y) Then
                    Edges(NumEdge + 1).Start = Faces(idxEdge).A
                    Edges(NumEdge + 1).End = Faces(idxEdge).B
                    Edges(NumEdge + 2).Start = Faces(idxEdge).B
                    Edges(NumEdge + 2).End = Faces(idxEdge).C
                    Edges(NumEdge + 3).Start = Faces(idxEdge).C
                    Edges(NumEdge + 3).End = Faces(idxEdge).A
                    Faces(idxEdge) = Faces(NumTri)
                    NumEdge = NumEdge + 3
                    idxEdge = idxEdge - 1
                    NumTri = NumTri - 1
                End If
            Loop While idxEdge < NumTri

'Tag multiple edges
'Note: if all Facess are specified anticlockwise then all
'interior edges are opposite pointing in direction.
            For idxEdge = 1 To NumEdge - 1
                If Not Edges(idxEdge).Start = 0 And _
                   Not Edges(idxEdge).End = 0 Then
                    For idxEdge2 = idxEdge + 1 To NumEdge
                        If Not Edges(idxEdge2).Start = 0 And _
                           Not Edges(idxEdge2).End = 0 Then
                            If Edges(idxEdge).Start = Edges(idxEdge2).End And _
                               Edges(idxEdge2).Start = Edges(idxEdge).End Then
                                Edges(idxEdge).Start = 0
                                Edges(idxEdge).End = 0
                                Edges(idxEdge2).Start = 0
                                Edges(idxEdge2).End = 0
                            End If
                       End If
                    Next idxEdge2
                End If
            Next idxEdge
                       
'Form new Facess for the current point
'Skipping over any tagged edges.
'All edges are arranged in clockwise order.
            For idxEdge = 1 To NumEdge
                If Not Edges(idxEdge).Start = 0 And Not Edges(idxEdge).End = 0 Then
                    NumTri = NumTri + 1
                    Faces(NumTri).A = Edges(idxEdge).Start
                    Faces(NumTri).B = Edges(idxEdge).End
                    Faces(NumTri).C = idxVert
                End If
            Next idxEdge
        Next idxVert
    End With
        
'Remove Facess with superFaces Dots
'These are Facess which have a screen number greater than NumVert
    idxVert = 0
    Do
        idxVert = idxVert + 1
        If Faces(idxVert).A > Dots1.NumSelDot Or _
           Faces(idxVert).B > Dots1.NumSelDot Or _
           Faces(idxVert).C > Dots1.NumSelDot Then
            Faces(idxVert).A = Faces(NumTri).A
            Faces(idxVert).B = Faces(NumTri).B
            Faces(idxVert).C = Faces(NumTri).C
            idxVert = idxVert - 1
            NumTri = NumTri - 1
        End If
    Loop While idxVert < NumTri
'______________________________________________________________________________________
'Transfer created faces to mesh.faces
    If NumTri = 0 Then Exit Sub
    
    Mesh1.NumMeshs = Mesh1.NumMeshs + 1
    ReDim Preserve Mesh1.Meshs(Mesh1.NumMeshs)
    With Mesh1.Meshs(Mesh1.NumMeshs)
        
        .NumFaces = CLng(NumTri)
        ReDim .Faces(1 To .NumFaces)
        ReDim .Normals(1 To .NumFaces)
        ReDim .NormalsT(1 To .NumFaces)
        ReDim .BorderEdges(1)
        
        For idx = 1 To .NumFaces
            With .Faces(idx)
            .A = Dots1.SelDots(Faces(idx).B).Index
            .B = Dots1.SelDots(Faces(idx).A).Index
            .C = Dots1.SelDots(Faces(idx).C).Index
            Mesh1.Meshs(Mesh1.NumMeshs).Normals(idx) = _
                CalculateNormal(Dots1.Dots(.A).Vector, Dots1.Dots(.B).Vector, Dots1.Dots(.C).Vector)
            Dots1.Dots(.A).Selected = False
            Dots1.Dots(.B).Selected = False
            Dots1.Dots(.C).Selected = False
            Call SearchEdge(.A, .B)
            Call SearchEdge(.B, .C)
            Call SearchEdge(.C, .A)
            End With
        Next
        ReDim Preserve .BorderEdges(UBound(.BorderEdges) - 1)
    End With
    Erase Faces
    Erase Edges
    Unsaved = True
    
End Sub

Public Sub SearchEdge(SEdge As Long, EEdge As Long)

    Dim idx         As Long
    Dim Coincident  As Boolean
    
    With Mesh1.Meshs(Mesh1.NumMeshs)
        For idx = 1 To UBound(.BorderEdges)
            With .BorderEdges(idx)
                If SEdge = .Start And EEdge = .End Then
                    .Used = .Used + 1: Coincident = True: Exit For
                ElseIf EEdge = .Start And SEdge = .End Then
                    .Used = .Used + 1: Coincident = True: Exit For
                End If
            End With
        Next

        If Not Coincident Then
            .BorderEdges(UBound(.BorderEdges)).Start = SEdge
            .BorderEdges(UBound(.BorderEdges)).End = EEdge
            ReDim Preserve .BorderEdges(UBound(.BorderEdges) + 1)
        End If
    End With

End Sub
