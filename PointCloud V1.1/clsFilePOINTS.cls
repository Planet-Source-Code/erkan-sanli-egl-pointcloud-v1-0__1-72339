VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilePOINTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PrevNum As Byte = 99

Dim hFile           As Long
Private DelimChar   As String
Private Lines()     As String

Public Function InputList(FileName As String, lstIn As ListBox)

    Dim idx As Byte
    lstIn.Clear
    ReDim Lines(PrevNum)
    hFile = FreeFile
    Open FileName For Binary As #hFile
        For idx = 0 To PrevNum
            Lines(idx) = ReadString
            lstIn.AddItem Lines(idx)
        Next
    Close #hFile

End Function

Private Function ReadString() As String

    Dim TempChar    As String
    Dim TempString  As String
    
    TempChar = StrConv(InputB(1, #hFile), vbUnicode)
    While TempChar <> vbLf
        TempString = TempString & TempChar
        TempChar = StrConv(InputB(1, #hFile), vbUnicode)
    Wend
    ReadString = Left$(TempString, Len(TempString) - 1)

End Function

Public Sub RefreshOutputList(lstOut As ListBox)
    
    Dim idx As Byte
    Dim Lines2()     As String
   
    If UBound(Lines) < 1 Then Exit Sub
    Lines2 = Lines
    lstOut.Clear
    For idx = 0 To PrevNum
        lstOut.AddItem Parse(Lines2(idx))
    Next
        
End Sub

Private Function Parse(inLine As String, Optional Preview As Boolean = True) As String
    
    
    Dim charpos     As Long
    Dim test        As String
    Dim FormatLine  As String
    Dim SplitLine() As String
    Dim NumSplitLine As Byte
    
    On Error Resume Next
'Replace delimiter chars to comma
    If delim.tTab Then inLine = Replace(inLine, vbTab, "|")
    If delim.tSemicolon Then inLine = Replace(inLine, ";", "|")
    If delim.tComma Then inLine = Replace(inLine, ",", "|")
    If delim.tSpace Then inLine = Replace(inLine, " ", "|")
    If delim.tOther Then inLine = Replace(inLine, delim.tDelimChar, "|")
'Eliminate double comma
    Do
        charpos = InStr(1, inLine, "||")
        inLine = Replace(inLine, "||", "|")
    Loop Until charpos = 0
    
'Delete first and end char comma
    test = Left$(inLine, 1)
    If test = "|" Then inLine = Right$(inLine, Len(inLine) - 1)
    test = Right$(inLine, 1)
    If test = "|" Then inLine = Left$(inLine, Len(inLine) - 1)
    
    SplitLine = Split(inLine, "|")
    NumSplitLine = UBound(SplitLine)
    
'Format
    If Preview Then
        Select Case delim.tFormat
            Case 0  '"X , Y , Z"
                If NumSplitLine = 2 Then
                    inLine = "X: " & SplitLine(0) & "   Y: " & SplitLine(1) & "   Z: " & SplitLine(2)
                Else
                    inLine = "Error line"
                End If
            Case 1  '"Number , X , Y , Z"
                If NumSplitLine = 3 Then
                    inLine = "N: " & SplitLine(0) & "   X: " & SplitLine(1) & "   Y: " & SplitLine(2) & "   Z: " & SplitLine(3)
                Else
                    inLine = "Error line"
                End If
            Case 2  '"X , Y , Z , Description"
                If NumSplitLine = 3 Then
                    inLine = "X: " & SplitLine(0) & "   Y: " & SplitLine(1) & "   Z: " & SplitLine(2) & "   D: " & SplitLine(3)
                Else
                    inLine = "Error line"
                End If
            Case 3  '"Number , X , Y , Z , Description"
                If NumSplitLine = 4 Then
                    inLine = "N: " & SplitLine(0) & "   X: " & SplitLine(1) & "   Y: " & SplitLine(2) & "   Z: " & SplitLine(3) & "   D: " & SplitLine(4)
                Else
                    inLine = "Error line"
                End If
        End Select
    Else
        Select Case delim.tFormat
            Case 0, 2 '"X , Y , Z" or "X , Y , Z , Description"
                    inLine = SplitLine(0) & "|" & SplitLine(1) & "|" & SplitLine(2)
            Case 1, 3 '"Number , X , Y , Z" or "Number , X , Y , Z , Description"
                    inLine = SplitLine(1) & "|" & SplitLine(2) & "|" & SplitLine(3)
        End Select
    End If
    Parse = inLine

End Function

Public Sub InputAll(FileName As String)
    
    Dim strData As String
    Dim idx     As Long
    Dim dx      As Single
    Dim dy      As Single
'    Dim dmax    As Single
    
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
    
    For idx = 0 To UBound(Lines)
        Lines(idx) = Parse(Lines(idx), False)
    Next

    With Dots1
        .NumDot = UBound(Lines) + 1
        Mesh1.NumVertices = .NumDot
        ReDim .Dots(1 To .NumDot)
        ReDim Mesh1.Vertices(1 To .NumDot)
        
        For idx = 1 To .NumDot
            Mesh1.Vertices(idx) = GetVectorValue(Lines(idx - 1))
            .Dots(idx).Vector = Mesh1.Vertices(idx)
            .Dots(idx).Visible = True
        Next idx
        Erase Lines
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
'    Call ResetCameraParameters
'    Call ResetLightParameters
    LoadComplete = True

End Sub

Private Function GetVectorValue(Line As String) As VECTOR4

    Dim Value As String
    Dim Segments() As String
    
    Segments = Split(Line, "|")
    
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

Private Sub Class_Initialize()
    ReDim Lines(0)
End Sub

