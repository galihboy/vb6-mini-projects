Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GetPrivateProfileSection Lib "kernel32.dll" Alias _
"GetPrivateProfileSectionA" (ByVal lpAppName As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias _
"GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As _
     Integer, ByVal lParam As Any) As Long

'constants for searching the ListBox
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const LB_FINDSTRING = &H18F

'Public Const nmFile As String = "dosen.txt"
Public Const nmFileINI As String = "dosen.ini"
Public Const nmFileRUANG As String = "ruang.ini"
Public Const nmFilePESAN As String = "pesan.txt"

Public alamatFile As String
Public alamatFileRuang As String
Public alamatFileINI As String
Public jmlDosen As Integer

Public Type arrData
    sKunci As String
    sNilai As Integer
End Type

'function to get find an item in the Listbox
Public Function GetListBoxIndex(HWND As Long, SearchKey As String, Optional FindExactMatch As Boolean = True) As Long

    If FindExactMatch Then
        GetListBoxIndex = SendMessage(HWND, LB_FINDSTRINGEXACT, -1, ByVal SearchKey)
    Else
        GetListBoxIndex = SendMessage(HWND, LB_FINDSTRING, -1, ByVal SearchKey)
    End If

End Function

Public Sub TambahData(strAlamatFile As String, strData As String)
    Dim iFileNo As Integer
    iFileNo = FreeFile
    
    Open strAlamatFile For Append As #iFileNo
        Print #iFileNo, strData
        'Print #iFileNo, strData
    Close #iFileNo
End Sub

'Private Function CountINIKeys(ByRef sINIPath As String) As Long
'    Dim Sections() As String
'    Dim Keys() As String
'    Dim lNumSections As Long
'    Dim lLoopSections As Long
'    Dim lNumKeys As Long
'    Dim lCountkeys As Long
 
' Get the number of all the sections within the INI file
'    Sections() = GetINISectionNames(sINIPath, lNumSections)
'    If (lNumSections < 1) Then 'If there are no section exit
'        Exit Function
'    End If
 
'    lCountkeys = 0
'    For lLoopSections = 0 To lNumSections - 1
'        If (Len(Sections(LoopSections))) Then
'            Keys() = GetINIKeyNames(sINIPath, Sections(lLoopSections), lNumKeys)
'            lCountkeys = lCountkeys + lNumKeys
'        End If
    
'    Next lLoopSections
    
'    CountINIKeys = lCountkeys
    
'End Function
 
'Private Function GetINIKeyNames(ByRef sFile As String, _
'ByRef sSection As String, ByRef lNumOfKeys As Long) As String()
'    Dim StrBuf As String
'    Dim BufLen As Long
'    Dim RetVal As Long
    
'    BufLen = 16
    
'    Do
'        BufLen = BufLen * 2
'        StrBuf = Space$(BufLen)
'        RetVal = GetPrivateProfileSection(sSection, StrBuf, BufLen, sFile)
'    Loop While RetVal = BufLen - 2
    
'    If (RetVal) Then
'        GetININames = Split(Left$(StrBuf, RetVal - 1), vbNullChar)
'        lNumOfKeys = UBound(GetININames) + 1
'    End If
    
'End Function
 
Public Function GetINISectionNames( _
ByRef sFile As String, ByRef lNumOfSections As Long) As String()
    Dim StrBuf As String
    Dim BufLen As Long
    Dim retval As Long
 
    BufLen = 16
 
    Do
        BufLen = BufLen * 2
        StrBuf = Space$(BufLen)
        retval = GetPrivateProfileSectionNames(StrBuf, BufLen, sFile)
    Loop While retval = BufLen - 2
 
    If (retval) Then
        GetINISectionNames = Split(Left$(StrBuf, retval - 1), vbNullChar)
        lNumOfSections = UBound(GetINISectionNames) + 1
    End If
    
End Function

