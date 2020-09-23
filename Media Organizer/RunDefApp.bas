Attribute VB_Name = "RunDefApp"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Public Function RunFile(Path As String, SenderForm As Form)
If Path <> "" Then
    ShellExecute SenderForm.hwnd, vbNullString, Path, vbNullString, "C:\", SW_SHOWNORMAL
End If
End Function

'***This module is used to run/execute files that can be run/executed***'
'***Path is the full path to the file to be run/executed***'
'****Senderform is the calling window***'
