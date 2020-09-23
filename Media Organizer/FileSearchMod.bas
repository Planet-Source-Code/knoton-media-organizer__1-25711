Attribute VB_Name = "FileSearchMod"
'***This module is used to find files, its path and the size of the file***'

Option Explicit
Public ReturnPath() As String       'Holds an array of paths to later on be saved in db
Public ReturnFileName() As String   'Holds an array of Filenames to later on be saved in db
Public ReturnFileSize() As String   'Holds an array of Filesize to later on be saved in db
Public noOfFiles As Long            'Holds the number of files found matching SearchStr
Public objInst As Object            'Used to get filesize
Public strExt() As String
'***FileFinding API***'
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
    OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function

Public Sub FindFiles(Path As String, SearchStr As String)
Dim FileName As String
Dim DirName As String
Dim dirNames() As String
Dim nDir As Integer
Dim i As Integer
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim g As Integer

Set objInst = CreateObject("WindowsInstaller.Installer")

If Right(Path, 1) <> "\" Then Path = Path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(Path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
    Do While Cont
    DirName = StripNulls(WFD.cFileName)
    If (DirName <> ".") And (DirName <> "..") Then
        ' Check for directory
        If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
            If DirName <> "RECYCLER" Then 'Do not list subdirs/files in recycle bin
                dirNames(nDir) = DirName
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
    End If
    Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
    Loop
    Cont = FindClose(hSearch)
End If
' Walk through this directory.
hSearch = FindFirstFile(Path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
    While Cont
        FileName = StripNulls(WFD.cFileName)
        If (FileName <> ".") And (FileName <> "..") Then
            If Not GetFileAttributes(Path & FileName) And FILE_ATTRIBUTE_DIRECTORY Then
                If strExt(0) = "*" Then
                    On Error GoTo ErrHandler
                    noOfFiles = noOfFiles + 1
                    ReDim Preserve ReturnPath(noOfFiles)
                    ReDim Preserve ReturnFileName(noOfFiles)
                    ReDim Preserve ReturnFileSize(noOfFiles)
                    ReturnPath(noOfFiles) = Path
                    ReturnFileName(noOfFiles) = FileName
                    ReturnFileSize(noOfFiles) = ByteConvert(objInst.filesize(Path & FileName))
                Else
                    For g = 0 To Form1.lstselExt.ListCount - 1
                        If InStrRev(FileName, "." & strExt(g), -1) Then
                            On Error GoTo ErrHandler
                            noOfFiles = noOfFiles + 1
                            ReDim Preserve ReturnPath(noOfFiles)
                            ReDim Preserve ReturnFileName(noOfFiles)
                            ReDim Preserve ReturnFileSize(noOfFiles)
                            ReturnPath(noOfFiles) = Path
                            ReturnFileName(noOfFiles) = FileName
                            ReturnFileSize(noOfFiles) = ByteConvert(objInst.filesize(Path & FileName))
                            Exit For
                        End If
                    Next g
                End If
            End If
        End If
        Cont = FindNextFile(hSearch, WFD)
    Wend
    Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
    For i = 0 To nDir - 1
        Call FindFiles(Path & dirNames(i) & "\", SearchStr)
    Next i
End If
Set objInst = Nothing
ErrHandler:
If Err.Number = -2147467259 Then
    ReturnFileSize(noOfFiles) = "?"
    On Error GoTo 0
    Resume Next
ElseIf Err.Number <> 0 Then
    MsgBox Err.Number & " " & Err.Description
End If
End Sub

Public Function ByteConvert(bytes As Double) As String
'Convert bytes to kb,mb,gb
If bytes < 1024 Then ByteConvert = FormatNumber(bytes, 0, 0, 0, True) & " b"
If bytes >= 1024 And bytes < 1048576 Then ByteConvert = FormatNumber((bytes / 1024), 2, 0, 0, True) & " KB"
If bytes >= 1048576 And bytes < 1073741824 Then ByteConvert = FormatNumber(((bytes / 1024) / 1024), 2, 0, 0, True) & " MB"
If bytes >= 1073741824 And bytes < 1099511627776# Then ByteConvert = FormatNumber((((bytes / 1024) / 1024) / 1024), 2, 0, 0, True) & " GB"
End Function

