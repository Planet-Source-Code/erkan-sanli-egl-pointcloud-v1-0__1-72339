Attribute VB_Name = "modKeyboard"
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const Step = 5

Public Sub UpdateParameters()

'    If Action Then Exit Sub
    
    If m_blnWheelPresent And SelType = NoSelect Then Call WheelAction(frmMain.picCanvas.hWnd)
    
    Dim T   As Single
    Dim R   As Single
    Dim s   As Single
    

    With Position
        R = 5
        s = 0.25
        T = .Sca * 2.5

'R: Mesh Rotate
        If State(vbKeyR) Then
            
            If State(vbKeyUp) Then Call ActRot(.Rot.X, -R)
            If State(vbKeyDown) Then Call ActRot(.Rot.X, R)
            If State(vbKeyLeft) Then Call ActRot(.Rot.Y, R)
            If State(vbKeyRight) Then Call ActRot(.Rot.Y, -R)
            If State(vbKeyPageUp) Then Call ActRot(.Rot.Z, -R)
            If State(vbKeyPageDown) Then Call ActRot(.Rot.Z, R)
            
'S: Mesh Scale
        ElseIf State(vbKeyS) Then
        
            If State(vbKeyPageDown) Or State(vbKeyLeft) Or State(vbKeyDown) Then Call ActSca(.Sca, -s)
            If State(vbKeyPageUp) Or State(vbKeyRight) Or State(vbKeyUp) Then Call ActSca(.Sca, s)
            
'T: Mesh Translate (Move)
        ElseIf State(vbKeyT) Then
            If State(vbKeyRight) Then Call ActTra(.Tra.X, T)
            If State(vbKeyLeft) Then Call ActTra(.Tra.X, -T)
            If State(vbKeyUp) Then Call ActTra(.Tra.Y, T)
            If State(vbKeyDown) Then Call ActTra(.Tra.Y, -T)
            If State(vbKeyPageDown) Then Call ActTra(.Tra.Z, T)
            If State(vbKeyPageUp) Then Call ActTra(.Tra.Z, -T)
        End If
    End With

'_______________________________Light_____________________________________
    
'L: Light Translate (Relative Camera1.worldposition,NO WORLD COORDINATE SYSTEM)
    With Light1.Position
        If State(vbKeyL) Then
            If State(vbKeyLeft) Then .X = .X + R
            If State(vbKeyRight) Then .X = .X - R
            If State(vbKeyUp) Then .Y = .Y - R
            If State(vbKeyDown) Then .Y = .Y + R
            If State(vbKeyPageUp) Then .Z = .Z - R
            If State(vbKeyPageDown) Then .Z = .Z + R
            Light1.Normal = VectorNormalize(Light1.Position)
'            Light1.NormalT = MatrixMultiplyVector(matOutput, Light1.Normal)
        End If
    End With
'______________________________Other__________________________________
        
    If State(vbKeyC) Then Call ResetCameraParameters
    If State(vbKeyX) Then Call ResetMeshParameters
    If State(vbKeyV) Then Call ResetLightParameters

'Clipping
    If State(vbKeyAdd) And ClipFar Then Call frmMain.cmdClipZ_Click(0)
    If State(vbKeySubtract) And ClipFar Then Call frmMain.cmdClipZ_Click(1)
    If State(vbKeyLeft) And ClipFar Then Call frmMain.cmdClipZ_Click(0)
    If State(vbKeyRight) And ClipFar Then Call frmMain.cmdClipZ_Click(1)
        
    matOutput = Out

End Sub

Public Sub ResetCameraParameters()
    
    With Camera1
        .ClipNear = 0
        .ClipFar = 500
        .Zoom = 1
        .FOV = ConvertZoomtoFOV(.Zoom)
        .WorldPosition = VectorSet(0, 0, -500)
        .LookAtPoint = VectorSet(0, 0, 0)
        .VUP = VectorSet(0, 1, 0)
    End With
    
End Sub

Public Sub ResetMeshParameters()
    
    Dim idx As Long

    With Position
        .Rot = VectorSet(0, 0, 0)
        .Tra = VectorSet(0, 0, 0)
        .Sca = 1
    End With

End Sub

Public Sub ResetLightParameters()
    
    With Light1
        .Position = VectorSet(0, 0, -500)
        .Normal = VectorNormalize(Light1.Position)
    End With

End Sub

Private Function State(key As Long) As Boolean
 
    Dim lngKeyState As Integer
    
    lngKeyState = GetKeyState(key)
    State = IIf((lngKeyState And &H8000), True, False)

End Function

Public Sub ActRot(Val As Single, Step As Single)
    
    Val = Val + Step
    Val = Val Mod 360

End Sub

Public Sub ActSca(Val As Single, Step As Single)
    
    Val = Val + Step
    If Val < 0.05 Then Val = 0.05
    InvScl = 1 / Val

End Sub

Public Sub ActTra(Val As Single, Step As Single)
    
    Val = Val + Step

End Sub

