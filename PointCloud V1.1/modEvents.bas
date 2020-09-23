Attribute VB_Name = "modEvents"
Option Explicit

'Mouse
Private Const SM_MOUSEWHEELPRESENT  As Long = 75
Private Const WM_MOUSEWHEEL         As Integer = &H20A

Private Type MSG
    hWnd        As Long
    message     As Long
    wParam      As Long
    lParam      As Long
    time        As Long
    pt          As POINTAPI
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long

Public m_blnWheelPresent    As Boolean

'Keyboard
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const StepRT = 5
Private Const StepS = 0.25

Public Sub RefreshEvents(hWnd As Long)
    
    Dim lResult         As Long
    Dim tMouseCords     As POINTAPI
    Dim lCurrentHwnd    As Long
    Dim tMSG            As MSG
    Dim iDir            As Single

'MouseWheel
    If m_blnWheelPresent And SelType = NoSelect Then
        lResult = GetCursorPos(tMouseCords)
        lCurrentHwnd = WindowFromPoint(tMouseCords.X, tMouseCords.Y)
        If lCurrentHwnd = hWnd Then
            lResult = GetMessage(tMSG, frmMain.hWnd, 0, 0)
            lResult = TranslateMessage(tMSG)
            lResult = DispatchMessage(tMSG)
            If tMSG.message = WM_MOUSEWHEEL Then
                iDir = Sgn(tMSG.wParam \ &H7FFF) * 0.25  '0.25 = scale step
                Call ActSca(Position.Sca, iDir)
            End If
        End If
    End If

'Keyboard
     

    With Position

'R: Mesh Rotate
        If State(vbKeyR) Then
            
            If State(vbKeyUp) Then Call ActRot(.Rot.X, -StepRT)
            If State(vbKeyDown) Then Call ActRot(.Rot.X, StepRT)
            If State(vbKeyLeft) Then Call ActRot(.Rot.Y, StepRT)
            If State(vbKeyRight) Then Call ActRot(.Rot.Y, -StepRT)
            If State(vbKeyPageUp) Then Call ActRot(.Rot.Z, -StepRT)
            If State(vbKeyPageDown) Then Call ActRot(.Rot.Z, StepRT)
            
'S: Mesh Scale
        ElseIf State(vbKeyS) Then
        
            If State(vbKeyPageDown) Or State(vbKeyLeft) Or State(vbKeyDown) Then Call ActSca(.Sca, -StepS)
            If State(vbKeyPageUp) Or State(vbKeyRight) Or State(vbKeyUp) Then Call ActSca(.Sca, StepS)
            
'T: Mesh Translate (Move)
        ElseIf State(vbKeyT) Then
            If State(vbKeyRight) Then Call ActTra(.Tra.X, StepRT)
            If State(vbKeyLeft) Then Call ActTra(.Tra.X, -StepRT)
            If State(vbKeyUp) Then Call ActTra(.Tra.Y, StepRT)
            If State(vbKeyDown) Then Call ActTra(.Tra.Y, -StepRT)
            If State(vbKeyPageDown) Then Call ActTra(.Tra.Z, StepRT)
            If State(vbKeyPageUp) Then Call ActTra(.Tra.Z, -StepRT)
        End If
    End With

'______________________________Other__________________________________
        
    If State(vbKeyX) Then Call ResetMeshParameters

'Clipping
    If State(vbKeyAdd) And ClipFar Then Call frmMain.cmdClipZ_Click(0)
    If State(vbKeySubtract) And ClipFar Then Call frmMain.cmdClipZ_Click(1)
    If State(vbKeyLeft) And ClipFar Then Call frmMain.cmdClipZ_Click(0)
    If State(vbKeyRight) And ClipFar Then Call frmMain.cmdClipZ_Click(1)
        
    matOutput = Out

End Sub

Public Sub ResetCameraParameters()
    
    With Camera1
        .Zoom = 1
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
        InvScl = 1
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

Public Sub MouseInit()
    
    m_blnWheelPresent = GetSystemMetrics(SM_MOUSEWHEELPRESENT)

End Sub

