VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iniTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2099
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mIniFileName As String

Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Property Let FileName(ByVal INIFileName As String)
Attribute FileName.VB_HelpID = 2101
'On Error Resume Next
    Dim NF As Long
    If Dir(INIFileName, vbNormal) = "" Then
      NF = FreeFile
      Open INIFileName For Output As #NF
      Close #NF
    End If
    mIniFileName = INIFileName
End Property


Public Property Get FileName() As String
'On Error Resume Next
    FileName = mIniFileName
End Property



Public Function GetValue(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As String) As String
Attribute GetValue.VB_HelpID = 2102
  On Error GoTo xERR
  Dim Value As String, retval As String, X As Integer
  retval = String$(255, 0)
  X = GetPrivateProfileString(Section, Key, DefaultValue, retval, Len(retval), mIniFileName)
  GetValue = Trim(Left(retval, X))
Exit Function
xERR:
  GetValue = DefaultValue
End Function

Public Function WriteValue(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
Attribute WriteValue.VB_HelpID = 2103
  On Error GoTo xERR
  Dim X As Integer
  X = WritePrivateProfileString(Section, Key, Value, mIniFileName)
  If X <> 0 Then WriteValue = True
  Exit Function
xERR:
End Function

Public Function GetAllSections() As Collection
Attribute GetAllSections.VB_HelpID = 2104
  Dim Value As String, retval As String, X As Integer
  Dim s() As String, i As Integer
  retval = String$(255, 0)
  X = GetPrivateProfileString(vbNullString, "", "", retval, Len(retval), mIniFileName)
  Value = Trim(Left(retval, X))
  s = Split(Value, Chr(0))
  Set GetAllSections = New Collection
  With GetAllSections
    For i = LBound(s) To UBound(s)
      If s(i) <> "" Then .Add s(i)
    Next
  End With
End Function

Public Function GetAllKeys(ByVal Section As String) As Collection
Attribute GetAllKeys.VB_HelpID = 2105
  Dim Value As String, retval As String, X As Integer
  Dim s() As String, i As Integer
  retval = String$(255, 0)
  X = GetPrivateProfileString(Section, vbNullString, "", retval, Len(retval), mIniFileName)
  Value = Trim(Left(retval, X))
  s = Split(Value, Chr(0))
  Set GetAllKeys = New Collection
  With GetAllKeys
    For i = LBound(s) To UBound(s)
      If s(i) <> "" Then .Add s(i)
    Next
  End With
End Function

Public Function DeleteSection(ByVal Section As String) As Boolean
Attribute DeleteSection.VB_HelpID = 2106
  On Error GoTo xERR
  Dim X As Integer
  'X = WritePrivateProfileString(Section, vbNullString, "", mIniFileName)
  X = WritePrivateProfileString(Section, 0&, 0&, mIniFileName)
  If X <> 0 Then DeleteSection = True
  Exit Function
xERR:
End Function

Public Function DeleteKey(ByVal Section As String, ByVal Key As String) As Boolean
Attribute DeleteKey.VB_HelpID = 2107
  On Error GoTo xERR
  Dim X As Integer
  X = WritePrivateProfileString(Section, Key, vbNullString, mIniFileName)
  If X <> 0 Then DeleteKey = True
  Exit Function
xERR:
End Function

