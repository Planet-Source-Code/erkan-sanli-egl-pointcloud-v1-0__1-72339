Attribute VB_Name = "modMouse"
Option Explicit

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

Public Sub EventWheel(hClient As Long)

    Dim lResult         As Long
    Dim tMouseCords     As POINTAPI
    Dim lCurrentHwnd    As Long
    Dim iDir            As Single
    Dim m_tMSG          As MSG
        
    lResult = GetCursorPos(tMouseCords)
    lCurrentHwnd = WindowFromPoint(tMouseCords.X, tMouseCords.Y)
    If lCurrentHwnd = hClient Then
        lResult = GetMessage(m_tMSG, frmMain.hWnd, 0, 0)
        lResult = TranslateMessage(m_tMSG)
        lResult = DispatchMessage(m_tMSG)
        If m_tMSG.message = WM_MOUSEWHEEL Then
            iDir = Sgn(m_tMSG.wParam \ &H7FFF) * 0.25  '0.25 = scale step
            Call ActSca(Position.Sca, iDir)
        End If
    End If
    
End Sub

Public Sub MouseInit()
    
    m_blnWheelPresent = GetSystemMetrics(SM_MOUSEWHEELPRESENT)

End Sub

