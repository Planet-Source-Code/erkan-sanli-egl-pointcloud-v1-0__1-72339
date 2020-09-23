VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DIB_RGB_COLORS As Long = 0

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
End Type

Private Type SAFEARRAYBOUND
    cElements       As Long
    lLbound         As Long
End Type

Private Type SAFEARRAY2D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    Bounds(1)       As SAFEARRAYBOUND
End Type

Private Type POINTAPI
    X               As Long
    Y               As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Function VarPtrArray Lib "MSVBVM60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public hDC       As Long
Public hDIB      As Long
Public Width     As Long
Public Height    As Long
Private m_SizeImage As Long
Private m_hOldDIB   As Long
Private m_lpBits    As Long
Private m_Data()    As Long
Private sa          As SAFEARRAY2D
Private m_hBrush    As Long

Public Sub Create(NewWidth As Long, NewHeight As Long, Optional Orientation As Boolean = False)
    
    On Error GoTo ErrorCreate
    
    Dim bi  As BITMAPINFO

    Width = NewWidth
    Height = NewHeight
    m_SizeImage = NewWidth * NewHeight * 4
    
    With bi.bmiHeader
        .biSize = Len(bi)
        .biWidth = Width
        .biHeight = IIf(Orientation, Height, -Height)
        .biPlanes = 1
        .biBitCount = 32
        .biSizeImage = m_SizeImage
    End With
    
    hDC = CreateCompatibleDC(0)
    If (hDC <> 0) Then
        hDIB = CreateDIBSection(hDC, bi, DIB_RGB_COLORS, m_lpBits, 0, 0)
        If (hDIB <> 0) Then
            m_hOldDIB = SelectObject(hDC, hDIB)
            With sa
                .cbElements = 4
                .cDims = 2
                .Bounds(0).lLbound = 0
                .Bounds(0).cElements = Height
                .Bounds(1).lLbound = 0
                .Bounds(1).cElements = Width
                .pvData = m_lpBits
            End With
            Call CopyMemory(ByVal VarPtrArray(m_Data()), VarPtr(sa), 4)
        Else
            Call Delete
        End If
    End If
    Exit Sub
    
ErrorCreate:
    MsgBox "Error: clsDIB > Create"
End Sub

Public Sub Clear()
    
    On Error GoTo ErrorClear
    
    Call ZeroMemory(m_Data(0, 0), m_SizeImage)
    Exit Sub

ErrorClear:
    MsgBox "Error: clsDIB > Clear"
End Sub

Public Sub Delete()
    
    On Error GoTo ErrorDelete
    
    If (hDC <> 0) Then
        If (hDIB <> 0) Then
            Call CopyMemory(ByVal VarPtrArray(m_Data()), 0&, 4)
            Call SelectObject(hDC, m_hOldDIB)
            Call DeleteObject(hDIB)
        End If
        Call DeleteDC(hDC)
    End If
    hDC = 0
    hDIB = 0
    m_hOldDIB = 0
    m_lpBits = 0
    Exit Sub
    
ErrorDelete:
    MsgBox "Error: clsDIB > Delete"
End Sub

Public Sub Paint(ByVal hDestDC As Long)
    BitBlt hDestDC, 0, 0, Width, Height, hDC, 0, 0, vbSrcCopy
End Sub

Public Function GetPixel(X As Single, Y As Single) As Long
    On Error Resume Next
    GetPixel = m_Data(X, Y)
End Function

Public Sub SetPixel(X As Single, Y As Single, lColor As Long)
    m_Data(X, Y) = lColor
End Sub

Private Sub Class_Initialize()
    hDIB = 0
End Sub

Private Sub Class_Terminate()
    Call Delete
End Sub

