VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileEPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' EPJ File parser (EGL Point Cloud Job file)
' Creator   : EGL Point Cloud V1
'             This is a my format.
'             This version, load a single object
' Author    : Erkan Þanlý July2009
' Copyright : Free code
' Version   : 1.0

Private Type Header
    Signature       As String * 3   '"EPJ"
    Version         As Byte
End Type

Dim fHeader         As Header

Public Sub WriteEPJ(FileName As String)
    
    Dim hFile       As Long
    
    On Error GoTo err
    If FileExist(FileName) Then Kill FileName
    
    fHeader.Signature = "EPJ"
    fHeader.Version = 1
        
    hFile = FreeFile
    Open FileName For Binary As hFile
        Put #hFile, , fHeader
        Put #hFile, , Dots1
        Put #hFile, , Mesh1
    Close #hFile
err:
    Close #hFile
End Sub

Public Sub ReadEPJ(FileName As String)
    
    Dim hFile   As Long
    
    On Error Resume Next
    
    LoadComplete = False
    hFile = FreeFile
    Open FileName For Binary As hFile

        Get #hFile, , fHeader
        If fHeader.Signature <> "EPJ" And fHeader.Version = 1 Then 'Check signature
            MsgBox "Wrong file format or version" & vbNewLine & "Abort loading"
            LoadComplete = False
            Exit Sub
        End If
        Get #hFile, , Dots1
        Get #hFile, , Mesh1
    Close #hFile
    Call ResetMeshParameters
    Call ResetCameraParameters
    Call ResetLightParameters
    LoadComplete = True

End Sub
