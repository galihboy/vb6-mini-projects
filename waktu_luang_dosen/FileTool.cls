VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2277
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As _
    Long
Const MAX_PATH = 260

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


Public Function GetFileName(ByVal FilePathName As String, Optional WithExtension As Boolean = False) As String
Attribute GetFileName.VB_HelpID = 2278
' ritorna il nome del file estratto da <FilePathName>
' ritorna "" se non lo trova
Dim a As Long, b As Long
Dim n As String

a = InStrRev(FilePathName, "\")
b = InStrRev(FilePathName, ".")
If b = 0 Or WithExtension Then b = Len(FilePathName) + 1
'If b > a + 1 And b > 0 Then
    GetFileName = Mid$(FilePathName, a + 1, b - a - 1)
'End If

End Function

Public Function DirExist(ByVal DirPath) As Boolean
Attribute DirExist.VB_HelpID = 2279
' true se la directory esiste
DirExist = IIf(DirPath > "", Dir(DirPath, vbDirectory) > "", False)
End Function

Public Function FileExist(ByVal PathFileName As String) As Boolean
Attribute FileExist.VB_HelpID = 2280
On Error GoTo xERR
FileExist = IIf(PathFileName > "", Dir(PathFileName) > "", False)
Exit Function
xERR:
End Function


Public Function ChangeFileExtension(ByVal FileName As String, ByVal Extension As String, _
    Optional ByVal AddIfMissing As Boolean) As String
Attribute ChangeFileExtension.VB_HelpID = 2281
    Dim i As Long
    Extension = Trim(Extension)
    Ensure Extension > "", Me, "ChangeFileExtension", "Empty extension"
    If Left(Extension, 1) = "." Then Extension = Mid(Extension, 2)
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case "."
                ' we've found an extension, so replace it
                ChangeFileExtension = Left$(FileName, i) & Extension
                Exit Function
            Case ":", "\"
                Exit For
        End Select
    Next
    
    ' there is no extension
    If AddIfMissing Then
        ChangeFileExtension = FileName & "." & Extension
    Else
        ChangeFileExtension = FileName
    End If
End Function

Public Function ExtractDirectoryPath(ByVal FullPath As String) As String
Attribute ExtractDirectoryPath.VB_HelpID = 2282

'INPUT: File FullPath
'RETURNS:  Directory only
'EXAMPLE:
'DirectoryFromFile("C:\Program Files\My Program\MyData.txt")
'Returns "C:\Program Files\My Program"

Dim sAns As String
sAns = Trim(FullPath)
If Len(sAns) = 0 Then Exit Function
If InStr(sAns, "\") = 0 Then Exit Function

If Right(sAns, 1) = "\" Then
    ExtractDirectoryPath = sAns
    Exit Function
End If

Do Until Right(sAns, 1) = "\"

    sAns = Left(sAns, Len(sAns) - 1)
Loop

ExtractDirectoryPath = sAns

End Function

Public Function IsValidFileName(ByVal FileName As String) As Boolean
Attribute IsValidFileName.VB_HelpID = 2283
Dim i As Long
Dim ch As String * 1
Dim Chars As String
Dim vbLT As New LangTool
Chars = "_!£$%&()=?'^ìéè*[]}{òç@à°#-§+|ù " & Chr(34)

For i = 1 To Len(FileName)
    ch = Mid$(LCase(FileName), i, 1)
    If Not vbLT.Between(ch, "a", "z") Then
        If Not vbLT.Between(ch, "0", "9") Then
            If Not vbLT.Between(ch, "0", "9") Then
                If InStr(1, Chars, ch) = 0 Then Exit Function
            End If
        End If
    End If
Next i

IsValidFileName = Len(Trim(FileName)) > 0
                
End Function

Public Function GetExtension(ByVal FullFilePath As String) As String
Attribute GetExtension.VB_HelpID = 2284
' ritorna, se esiste, l'estensione del file

Dim p As Long
If Len(FullFilePath) > 0 Then
    p = InStrRev(FullFilePath, ".")
    If p > 0 Then
        If p < Len(FullFilePath) Then
            GetExtension = Mid(FullFilePath, p + 1)
        End If
    End If
End If
        
End Function


' Returns a one-based string array containing the files (or directories)
' located in the specified path
'
' the Path argument can contain wildcards, e.g. "C:\*.doc")

Function FilesToArray(ByVal path As String, Optional ByVal IncludeDirs As _
    Boolean) As String()
Attribute FilesToArray.VB_HelpID = 2285
    Dim lRet As Long
    Dim handle As Long
    Dim FindData As WIN32_FIND_DATA
    Dim FileName As String
    Dim fileCount As Long
    Dim ok As Boolean
    Dim res() As String
    Const IncRes As Long = 256
    ' start the searching, exit if no file matches the spec
    handle = FindFirstFile(path, FindData)
    If handle < 0 Then
        FilesToArray = res
        Exit Function
    End If
    
    Do
        ' get this entry's name
        FileName = Left$(FindData.cFileName, InStr(FindData.cFileName, _
            vbNullChar) - 1)
        
        If (FindData.dwFileAttributes And vbDirectory) = 0 Then
            ' this is a file
            ok = Not IncludeDirs
        ElseIf FileName <> "." And FileName <> ".." Then
            ' this is a directory, but not a ./.. entry
            ok = IncludeDirs
        Else
            ' this is a ./.. entry
            ok = False
        End If
            
        If ok Then
            ' add this entry to the result
            If fileCount = 0 Then ReDim res(0)
            res(fileCount) = FileName
            fileCount = fileCount + 1
            If fileCount > UBound(res) Then
                ' make room in the array if necessary
                ReDim Preserve res(0 To fileCount + IncRes) As String
            End If

        End If
        ' read the next file, returns zero when there are no more files
        lRet = FindNextFile(handle, FindData)
    Loop While lRet

    ' stop enumeration
    FindClose handle
    
    ' discard unused array items and return to caller
    ReDim Preserve res(0 To fileCount - 1) As String
    FilesToArray = res
End Function

' Returns a collection holding all the filenames that
' match a given filespec and search attributes.

Function GetFiles(filespec As String, Optional Attributes As VbFileAttribute) _
    As Collection
Attribute GetFiles.VB_HelpID = 2286
    Dim FileName As String
    
    Set GetFiles = New Collection
    
    ' start the search
    FileName = Dir$(filespec, Attributes)
    
    Do While Len(FileName)
        ' we've found a new file
        GetFiles.Add FileName, FileName
        ' get ready for the next iteration
        FileName = Dir$
    Loop
End Function

' Returns a collection holding all the subdirectories in a path
' that match search attributes (optionally it returns the entire path).

Function GetDirectories(path As String, Optional Attributes As VbFileAttribute, _
    Optional IncludePath As Boolean) As Collection
Attribute GetDirectories.VB_HelpID = 2287
    Dim dirname As String
    Dim path2 As String
    
    ' initialize the result
    Set GetDirectories = New Collection
    
    ' build the path name with a trailing backslash
    path2 = path
    If Right$(path2, 1) <> "\" Then path2 = path2 & "\"
    
    ' start the search
    dirname = Dir$(path2 & "*.*", vbDirectory Or Attributes)
    
    Do While Len(dirname)
        If dirname = "." Or dirname = ".." Then
            ' exclude the "." and ".." entries
        ElseIf (GetAttr(path2 & dirname) And vbDirectory) = 0 Then
            ' ignore regular files
        Else
            ' this is a directory
            ' include the path if requested
            If IncludePath Then dirname = path2 & dirname
            GetDirectories.Add dirname, dirname
        End If
        ' get next string
        dirname = Dir$
    Loop
  
End Function


Public Function ReadFile(ByVal FilePathName As String) As String
Attribute ReadFile.VB_HelpID = 2288
' returns then whole file in a string, if present
Ensure FileExist(FilePathName), Me, "ReadFile", "File '" & FilePathName & "' doesn't exist."
Dim NF As Long

NF = FreeFile
If NF > 0 Then
    Open FilePathName For Binary As #NF
    ReadFile = Space$(LOF(NF))
    Get #NF, , ReadFile
    Close #NF
End If

End Function

Public Function WriteFile(ByRef WholeFile As String, _
    ByVal FilePathName As String, Optional ByVal OverWrite As Boolean = True) As Boolean
Attribute WriteFile.VB_HelpID = 2289
' returns true if the file has been written
Dim NF As Long
Dim path As String

NF = FreeFile
If NF > 0 Then
    path = ExtractDirectoryPath(FilePathName)
    Ensure DirExist(path), Me, "WriteFile", "Path '" & path & "' doesn't exist"
    If OverWrite Then
        If FileExist(FilePathName) Then Kill FilePathName
    End If
    Open FilePathName For Binary Access Write As #NF
        Put #NF, , WholeFile
    Close #NF
    WriteFile = True
End If

End Function

