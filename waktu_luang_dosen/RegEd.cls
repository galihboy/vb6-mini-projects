VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2114
Option Explicit

Public Enum RegEdConsts
    rgHKEY_CLASSES_ROOT = &H80000000
    rgHKEY_CURRENT_USER = &H80000001
    rgHKEY_LOCAL_MACHINE = &H80000002
    rgHKEY_USERS = &H80000003
    rgHKEY_PERFORMANCE_DATA = &H80000004
    rgHKEY_CURRENT_CONFIG = &H80000005
    rgHKEY_DYN_DATA = &H80000006
End Enum

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, _
    ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, _
    phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long


Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_SUCCESS = 0&

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4

Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))
Const REG_OPENED_EXISTING_KEY = &H2
Const KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
                           ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const MAXWin9xLength As Long = 255
Private vbLT As New LangTool

'
' Create a registry key, then close it
' Returns True if the key already existed, False if it was created

Public Function CreateRegistryKey(ByVal hKey As RegEdConsts, ByVal KeyName As String) As Boolean
Attribute CreateRegistryKey.VB_HelpID = 2115
Dim handle As Long, disposition As Long
        
If RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, 0, handle, disposition) Then
    'Err.Raise 1001, , "Unable to create the registry key"
    CreateRegistryKey = False
Else
    ' Return True if the key already existed.
    CreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
    ' Close the key.
    RegCloseKey handle
    CreateRegistryKey = True
End If

End Function

' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

' Return True if a Registry key exists

Public Function CheckRegistryKey(ByVal hKey As RegEdConsts, ByVal KeyName As String) As Long
Attribute CheckRegistryKey.VB_HelpID = 2116
Dim handle As Long
Dim Ret As Long
' Try to open the key

Ret = RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle)
Select Case Ret
    Case 0:
        ' The key exists
        CheckRegistryKey = CLng(True)
        ' Close it before exiting
        RegCloseKey handle
    Case 5:
        ' access denied
        CheckRegistryKey = 5
    Case Else:
        ' access denied
        CheckRegistryKey = Ret
End Select

End Function

' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Enumerate registry keys under a given key
'
' returns a collection of strings

Public Function EnumRegistryKeys(ByVal hKey As RegEdConsts, ByVal KeyName As String) As Collection
Attribute EnumRegistryKeys.VB_HelpID = 2117
Dim handle As Long
Dim Length As Long
Dim Index As Long
Dim subkeyName As String

' initialize the result collection
Set EnumRegistryKeys = New Collection

' Open the key, exit if not found
If Len(KeyName) Then
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    ' in all case the subsequent functions use hKey
    hKey = handle
End If

Do
    ' this is the max length for a key name
    Length = 260
    subkeyName = Space$(Length)
    ' get the N-th key, exit the loop if not found
    If RegEnumKey(hKey, Index, subkeyName, Length) Then Exit Do
    
    ' add to the result collection
    subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
    EnumRegistryKeys.Add subkeyName, subkeyName
    ' prepare to query for next key
    Index = Index + 1
Loop

' Close the key, if it was actually opened
If handle Then RegCloseKey handle
        
End Function

Public Function SeekRegistryValue(ByVal hKey As RegEdConsts, ByVal KeyName As String, _
    ByVal EntryName As String, ByRef EntryValue) As Boolean
Attribute SeekRegistryValue.VB_HelpID = 2118
Dim re() As RegistryEntry
Dim reLen As Long
Dim i As Long

re = EnumRegistryValues(hKey, KeyName)

If Not vbLT.IsEmptyArray(re) Then
    reLen = vbLT.ArraySize(re)
    vbLT.ResetVar EntryValue
    If reLen > 0 Then
        For i = 0 To reLen - 1
            If re(i).EntryName = EntryName Then
                EntryValue = re(i).EntryValue
                SeekRegistryValue = True
                Exit Function
            End If
        Next i
    End If
End If

End Function
Public Function EnumRegistryValues(ByVal hKey As RegEdConsts, ByVal KeyName As String) As RegistryEntry()
Attribute EnumRegistryValues.VB_HelpID = 2119
' ritorna un collezione di coppie (nome,valore) dove
' valore è del tipo relativo a quello rappresentato nel registro
Dim handle As Long
Dim Length As Long
Dim Index As Long
Dim subkeyName As String, res As Long
Dim ValName As String, lenValName As Long
Const ValLen As Long = 1024
Dim ValType As Long
Dim DataBuffer() As Byte, lenDataBuffer As Long
Const DataBufferLen As Long = 4096
Dim byteArrayItem() As Byte
Dim stringItem As String
Dim longItem As Long
Dim Item As Variant
Dim i As Long, s As String
Dim retval() As RegistryEntry, rEntry As RegistryEntry

' Open the key, exit if not found
If Len(KeyName) Then
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    ' in all case the subsequent functions use hKey
    hKey = handle
End If

Do
    ValName = Space$(ValLen)
    lenValName = ValLen
    ReDim DataBuffer(DataBufferLen)
    lenDataBuffer = DataBufferLen
    res = RegEnumValue(hKey, Index, ValName, lenValName, 0&, ValType, DataBuffer(0), lenDataBuffer)
    If res = ERROR_SUCCESS Then
        
        vbLT.ResetVar Item
        Erase byteArrayItem
        stringItem = ""
        longItem = 0
        
        Select Case ValType
            Case REG_BINARY ' 3 ==> ritorna un array di byte
                If lenDataBuffer > 0 Then
                    ReDim byteArrayItem(lenDataBuffer - 1)
                    For i = 0 To lenDataBuffer - 1
                        byteArrayItem(i) = DataBuffer(i)
                    Next i
                End If
                Item = byteArrayItem
                
            Case REG_DWORD ' 4 ==> ritorna un long
                If lenDataBuffer > 0 Then
                    s = Space$(lenDataBuffer)
                    longItem = 0
                    For i = 1 To lenDataBuffer
                        On Error Resume Next
                            longItem = longItem + 256 ^ (i - 1) * DataBuffer(i - 1)
                        On Error GoTo 0
                    Next i
                End If
                Item = longItem
                
            Case Else
                If lenDataBuffer > 0 Then
                    stringItem = Space$(lenDataBuffer - 1)
                    For i = 1 To lenDataBuffer - 1
                        Mid$(stringItem, i, 1) = Chr$(DataBuffer(i - 1))
                    Next i
                End If
                Item = stringItem
                
        End Select
                
        ReDim Preserve retval(Index)
        Set rEntry = New RegistryEntry
        rEntry.EntryName = Left(ValName, lenValName)
        rEntry.EntryValue = Item
        Set retval(Index) = rEntry
        Index = Index + 1
    End If
Loop While (res = ERROR_SUCCESS)
If res = ERROR_NO_MORE_ITEMS Then 'ok
   EnumRegistryValues = retval
End If
    
End Function

Public Function ReadEntryValue(ByVal hKey As RegEdConsts, ByVal KeyName As String, _
            ByVal EntryName As String, ByRef EntryValue As Variant) As Boolean
Attribute ReadEntryValue.VB_HelpID = 2120
        ' cerca nella chiave specificata na voce; se la trova ritorna true
        ' ed aggiorna il valore di EntryValue
        Dim res() As RegistryEntry
        Dim VBex As New LangTool, i As Long
        Dim v As Variant, b() As Byte
        On Error GoTo xERR

    res = EnumRegistryValues(hKey, KeyName)
    If Not VBex.IsEmptyArray(res) Then
            For i = LBound(res) To UBound(res)
            If Trim(LCase(res(i).EntryName)) = Trim(LCase(EntryName)) Then
                v = res(i).EntryValue
                ReadEntryValue = True
                    Select Case VarType(EntryValue)
                    Case vbString:  EntryValue = CStr(v)
                    Case vbBoolean: EntryValue = CBool(v)
                    Case vbLong:    EntryValue = CLng(v)
                    Case vbInteger: EntryValue = CInt(v)
                    Case vbSingle:    EntryValue = CSng(v)
                    Case vbDouble:    EntryValue = CDbl(v)
                    Case vbCurrency:    EntryValue = CCur(v)
                    Case vbArray + vbByte: EntryValue = v
                    Case Else: Err.Raise 1001, , "Unsupported value type :" & TypeName(v)
                    ReadEntryValue = False
                    End Select
                    'EntryValue = v
                    Exit Function
                End If
        Next i
        End If
Exit Function
xERR:
    LogPrint "RegistryTool.ReadEntryValue Err#" & Err.Number & " at line " & Erl & "; " & Err.Description
    LogPrint "RegistryTool.ReadEntryValue > KeyName = " & KeyName & "; EntryName =" & EntryName
    'Throw Me, "ReadEntryValue", Err.Description
    ReadEntryValue = False
End Function


' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Delete a registry key
'
' Under Windows NT it doesn't work if the key contains subkeys

Public Function DeleteRegistryKey(ByVal hKey As RegEdConsts, ByVal KeyName As String) As Boolean
Attribute DeleteRegistryKey.VB_HelpID = 2121
    DeleteRegistryKey = (RegDeleteKey(hKey, KeyName) = 0)
End Function


' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Delete a registry value
'
' Return True if successful, False if the value hasn't been found

Public Function DeleteRegistryValue(ByVal hKey As RegEdConsts, ByVal KeyName As String, ByVal ValueName As String) As Boolean
Attribute DeleteRegistryValue.VB_HelpID = 2122
Dim handle As Long
Dim Ret As Long

' Open the key, exit if not found
If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Function
Err.Clear
Ret = RegDeleteValue(handle, ValueName)
' Delete the value (returns 0 if success)
'Debug.Print Ret, Err.LastDllError
DeleteRegistryValue = (Ret = 0)
' Close the handle
RegCloseKey handle

End Function



' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Write or Create a Registry value
' returns True if successful
'
' Use KeyName = "" for the default value
'
' Value can be an integer value (REG_DWORD), a string (REG_SZ)
' or an array of binary (REG_BINARY). Raises an error otherwise.

Public Function SetRegistryValue(ByVal hKey As RegEdConsts, ByVal KeyName As String, ByVal ValueName As String, Value As Variant) As Boolean
Attribute SetRegistryValue.VB_HelpID = 2123
Dim handle As Long
Dim lngValue As Long
Dim StrValue As String
Dim binValue() As Byte
Dim Length As Long
Dim retval As Long

' Open the key, exit if not found
If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) <> 0 Then Exit Function

retval = -1

' three cases, according to the data type in Value
Select Case VarType(Value)
    Case vbInteger, vbLong
        lngValue = CLng(Value)
        retval = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
    Case vbString, vbBoolean
        StrValue = CStr(Value)
        retval = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal StrValue, _
            Len(StrValue))
    Case vbArray + vbByte
        binValue = Value
        Length = UBound(binValue) - LBound(binValue) + 1
        retval = RegSetValueEx(handle, ValueName, 0&, REG_BINARY, _
            binValue(LBound(binValue)), Length)
    Case Else
        RegCloseKey handle
        Err.Raise 1001, , "Unsupported value type :" & TypeName(Value)
End Select

' Close the key and signal success
RegCloseKey handle
' signal success if the value was written correctly
SetRegistryValue = (retval = 0)

End Function


' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

Public Function EraseRegistryTree(ByVal hKey As RegEdConsts, ByVal KeyName As String, ByRef MaxLevelsErased As Long, Optional ByRef TotKeys As Long = 0, Optional ByRef DelKeys As Long = 0) As Boolean
Attribute EraseRegistryTree.VB_HelpID = 2124
    ' elimina una chiave e tutte le sue sottochiavi
    ' ritorna true se tutte la chiave e le sottochivi  sono state eliminate
    ' totKeys = numero chiavi navigate, DelKeys = numero chiavi cancellate
    ' MaxLevelErased = numero massimo di livello delle sottochiavi
    
    Dim regEntries As Collection
    Dim rT As New RegistryTool, rEntry As Variant
    Dim MaxLE As Long, ActLevel As Long, PassLevel As Long
    Static notFirstTime As Boolean, Ret As Boolean
    Dim ActLevelDelCount As Long
        
    If Not notFirstTime Then
        If Not CheckRegistryKey(hKey, ByVal KeyName) Then
            EraseRegistryTree = False
            Exit Function
        End If
        notFirstTime = True
        MaxLevelsErased = 0 ' init first time
        TotKeys = 1
        DelKeys = 0
    End If
    
    Set regEntries = rT.EnumRegistryKeys(hKey, KeyName)
        
    If regEntries.Count = 0 Then ' it's a leaf
        If CheckRegistryKey(hKey, KeyName) Then
            Ret = rT.DeleteRegistryKey(hKey, KeyName)
            If Ret Then DelKeys = DelKeys + 1
            EraseRegistryTree = Ret
        End If
        Exit Function
    Else
        TotKeys = TotKeys + regEntries.Count
        ActLevel = MaxLevelsErased + 1
        MaxLE = ActLevel
        ActLevelDelCount = 0
        For Each rEntry In regEntries
            PassLevel = ActLevel
            Ret = EraseRegistryTree(hKey, KeyName & "\" & rEntry, PassLevel, TotKeys, DelKeys)
            If Ret Then ActLevelDelCount = ActLevelDelCount + 1
            If (PassLevel > MaxLE) And Ret Then MaxLE = PassLevel
        Next
        MaxLevelsErased = MaxLE
        Ret = rT.DeleteRegistryKey(hKey, KeyName)
        If Ret Then DelKeys = DelKeys + 1
        If ActLevel = 1 Then
            EraseRegistryTree = (TotKeys = DelKeys)
            notFirstTime = False
        Else
            EraseRegistryTree = (ActLevelDelCount = regEntries.Count)
        End If
    End If
    
End Function

Public Function SeekRegistryLeafs(ByVal hKey As RegEdConsts, ByVal KeyName As String, _
    Optional ByVal MaxEntries As Long = -1) As Collection
Attribute SeekRegistryLeafs.VB_HelpID = 2125
    ' ritorna una collection di stringhe
Dim RegKey As Variant, SeekKeys As Collection, SeekKey As Variant
Dim EnumKeys As Collection
Dim RetValue As New Collection
Dim Key As String, IsLeaf As Boolean, AddIT As Boolean
Dim TotKeys As New Collection
Static KeysFound As Long
Static Level As Long

Level = Level + 1
Set EnumKeys = EnumRegistryKeys(hKey, KeyName)

If EnumKeys Is Nothing Then
    ' è una foglia
    IsLeaf = True
Else
    IsLeaf = (EnumKeys.Count = 0)
End If

If IsLeaf Then
    AddIT = IIf(MaxEntries < 0, True, (KeysFound < MaxEntries))
    If AddIT Then
        TotKeys.Add KeyName
        KeysFound = KeysFound + 1
    End If
Else
    For Each RegKey In EnumKeys
        Key = vbLT.AppendToPath(KeyName, RegKey)
        If MaxEntries < 0 Then
            Set SeekKeys = SeekRegistryLeafs(hKey, Key)
        Else
            Set SeekKeys = SeekRegistryLeafs(hKey, Key, MaxEntries)
        End If

        If SeekKeys.Count > 0 Then
            For Each SeekKey In SeekKeys
                TotKeys.Add SeekKey
            Next
        End If
    Next
End If
Set SeekRegistryLeafs = TotKeys

Level = Level - 1
If Level = 0 Then KeysFound = 0

End Function

Public Function FlushRegistryChanges(ByVal hKey As RegEdConsts) As Boolean
Attribute FlushRegistryChanges.VB_HelpID = 2126
Dim t As Single
t = Timer
FlushRegistryChanges = (RegFlushKey(hKey) = 0)
t = Timer - t: Debug.Print "Registry flushed in " & Format(t, "#,##0.000") & " secs"
End Function
