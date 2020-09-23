Attribute VB_Name = "modFile"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function FileExist(FileName As String) As Boolean
    
    Dim s As String
    
    On Error GoTo err
    s = Dir(FileName)
    If Len(s) Then FileExist = True
    Exit Function
err:
    FileExist = False
    
End Function

Public Function GetFileNameEx(strFilePath As String) As String
   
    If Len(strFilePath) Then GetFileNameEx = Dir(strFilePath)
    
End Function






