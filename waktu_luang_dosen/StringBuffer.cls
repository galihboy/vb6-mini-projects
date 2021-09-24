VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2127
Option Explicit

' STRINGBUFFER CLASS
' 19.8.2001 BY RICCARDO IOTTI
' updates
' 07.10.2001
'   - corretto errore che si verificava nell'append quando la stringa da aggiungere
'     era più lunga del blockSize
'   - aggiornata la gestione dell'allocazione dinamica di memoria per la stringa; ora
'     il blocksize rappresenta il "quanto di allocazione"

Private pStr As String
Private pBlockSize As Long
Private pBlocksCount As Long
Private pAddedBlocksCountSTAT As Long
Private pLength As Long

Private Initialized As Boolean

Public Sub Init(Optional ByVal BlockLen As Long = 65536)
Attribute Init.VB_HelpID = 2128
Ensure Not Initialized, Me, "Init", "StringBuffer has been alreday initialized"

pLength = 0
pBlockSize = BlockLen
pBlocksCount = 0
pStr = ""
Initialized = True

End Sub

Private Sub AutoInit()

pLength = 0
pBlockSize = 65536
pBlocksCount = 0
pStr = ""
Initialized = True

End Sub


Public Property Get Value() As String
Attribute Value.VB_HelpID = 2129
Attribute Value.VB_UserMemId = 0
If Not Initialized Then AutoInit
Value = Left$(pStr, pLength)
End Property

Public Property Let Value(ByVal NewValue As String)
Dim L As Long
If Not Initialized Then AutoInit
L = Len(NewValue)
SetSize L
Mid$(pStr, 1, L) = NewValue

End Property

Public Function Append(ByVal StrToAppend As String) As StringBuffer
Attribute Append.VB_HelpID = 2130
Dim L As Long, p As Long
If Not Initialized Then AutoInit

L = Len(StrToAppend)
p = pLength
UpdateSize pLength + L
Mid$(pStr, p + 1, L) = StrToAppend

Set Append = Me

End Function

Public Property Get Length() As Long
Attribute Length.VB_HelpID = 2131
Length = pLength
End Property


Public Function GetSubString(ByVal StartPos As Long, Optional ByVal Length As Long = -1) As StringBuffer
Attribute GetSubString.VB_HelpID = 2132
If Not Initialized Then AutoInit

Dim mSB As New StringBuffer, mS As String

Ensure StartPos > 0 And StartPos <= pLength, Me, "GetSubString", "Invalid start position"

If Length < 0 Then
    mS = Mid$(Left$(pStr, pLength), StartPos)
Else
    Ensure StartPos + Length - 1 < pLength, Me, "GetSubString", "Invalid length"
    mS = Mid$(pStr, StartPos, Length)
End If
mSB.Value = mS
Set GetSubString = mSB

End Function

Public Function SetSubString(ByVal StartPos As Long, ByVal Length As Long, ByVal ValueStr As String) As StringBuffer
Attribute SetSubString.VB_HelpID = 2133
If Not Initialized Then AutoInit

Dim mSB As New StringBuffer, mS As String

Ensure StartPos > 0 And StartPos <= pLength, Me, "SetSubString", "Invalid start position"
Ensure Length >= 0, Me, "SetSubString", "Invalid length"

mSB.Value = Left$(pStr, StartPos - 1) & ValueStr & Mid$(pStr, StartPos + Length, pLength - StartPos - Length + 1)
Set SetSubString = mSB

End Function


Public Property Get CharAt(ByVal Position As Long) As String
Attribute CharAt.VB_HelpID = 2134
If Not Initialized Then AutoInit
Ensure Position > 0 And Position <= pLength, Me, "CharAt", "Invalid char position"
CharAt = Mid$(pStr, Position, 1)
End Property

Public Property Let CharAt(ByVal Position As Long, ByVal CharStr As String)
If Not Initialized Then AutoInit
Ensure Len(CharStr) = 1, Me, "CharAt", "Invalid char length"
Ensure Position > 0 And Position <= pLength, Me, "CharAt", "Invalid char position"
Mid$(pStr, Position, 1) = CharStr
End Property

Public Function LeftStr(ByVal Length As Long) As StringBuffer
Attribute LeftStr.VB_HelpID = 2135
If Not Initialized Then AutoInit

Dim mSB As New StringBuffer

Ensure Length >= 0 And Length <= pLength, Me, "LeftStr", "Invalid length"
mSB.Value = Left$(pStr, Length)
Set LeftStr = mSB
End Function

Public Function RightStr(ByVal Length As Long) As StringBuffer
Attribute RightStr.VB_HelpID = 2136
If Not Initialized Then AutoInit
Dim mSB As New StringBuffer
Ensure Length >= 0 And Length <= pLength, Me, "RightStr", "Invalid length"
mSB.Value = Right$(pStr, Length)
Set RightStr = mSB
End Function

Public Function RightTrim() As StringBuffer
Attribute RightTrim.VB_HelpID = 2137
If Not Initialized Then AutoInit
Dim mSB As New StringBuffer
mSB.Value = RTrim$(Left$(pStr, pLength))
Set RightTrim = mSB
End Function

Public Function LeftTrim() As StringBuffer
Attribute LeftTrim.VB_HelpID = 2138
If Not Initialized Then AutoInit
Dim mSB As New StringBuffer
mSB.Value = LTrim$(pStr)
Set LeftTrim = mSB
End Function

Public Function FullTrim() As StringBuffer
Attribute FullTrim.VB_HelpID = 2139
If Not Initialized Then AutoInit
Dim mSB As New StringBuffer
mSB.Value = Trim$(pStr)
Set FullTrim = mSB
End Function

Public Function Equals(ByVal CmpStr) As Boolean
Attribute Equals.VB_HelpID = 2140
Dim L As Long
If Not Initialized Then AutoInit
L = Len(CmpStr)
If L = pLength Then
    Equals = (CmpStr = Left$(pStr, pLength))
End If
End Function

Public Function EqualsIgnoreCase(ByVal CmpStr) As Boolean
Attribute EqualsIgnoreCase.VB_HelpID = 2141
If Not Initialized Then AutoInit
Dim L As Long
L = Len(CmpStr)
If L = pLength Then
    EqualsIgnoreCase = (LCase$(CmpStr) = LCase$(Left$(pStr, pLength)))
End If
End Function

Public Function LowCase() As StringBuffer
Attribute LowCase.VB_HelpID = 2142
If Not Initialized Then AutoInit
Dim mSB As New StringBuffer
mSB.Value = LCase$(Left$(pStr, pLength))
Set LowCase = mSB
End Function

Public Function UpCase() As StringBuffer
Attribute UpCase.VB_HelpID = 2143
If Not Initialized Then AutoInit
Dim mSB As New StringBuffer
mSB.Value = UCase$(Left$(pStr, pLength))
Set UpCase = mSB
End Function

Public Sub Clear()
Attribute Clear.VB_HelpID = 2144
Dim bs As Long
bs = pBlockSize
Initialized = False
Init bs
End Sub

Public Function Reverse() As StringBuffer
Attribute Reverse.VB_HelpID = 2145
' inverte la stringa
Dim mSB As New StringBuffer
If Not Initialized Then AutoInit

If pLength > 0 Then mSB.Value = StrReverse(Left$(pStr, pLength))

Set Reverse = mSB

End Function


Public Function InString(ByVal StartPosition As Long, _
    ByVal SeekStr As String, Optional ByVal CmpMethod As VbCompareMethod = vbBinaryCompare) As Long
Attribute InString.VB_HelpID = 2146
If Not Initialized Then AutoInit
Dim L As Long
Ensure StartPosition > 0 And StartPosition <= pLength, Me, "InString", "Wrong start position"
L = Len(SeekStr)
Ensure L > 0, Me, "InString", "Empty seek character"
If L + StartPosition - 1 <= pLength Then
    InString = InStr(StartPosition, Left$(pStr, pLength), SeekStr, CmpMethod)
End If

End Function
    
    
Public Function InStringReverse(ByVal StartPosition As Long, _
    ByVal SeekStr As String, Optional ByVal CmpMethod As VbCompareMethod = vbBinaryCompare) As Long
Attribute InStringReverse.VB_HelpID = 2147
Dim L As Long
If Not Initialized Then AutoInit
' ho uniformato la sintassi s InString; cambia solo che lavora sulla stringa rovesciata
Ensure StartPosition > 0 And StartPosition <= pLength, Me, "InString", "Wrong start position"
L = Len(SeekStr)
Ensure L > 0, Me, "InStringReverse", "Empty seek character"
If L + StartPosition - 1 <= pLength Then
    InStringReverse = InStrRev(Left$(pStr, pLength), SeekStr, 1 + pLength - StartPosition, CmpMethod)
End If

End Function

Public Function ReplaceStr(ByVal StrSeek As String, ByVal StrReplace As String, _
    Optional ByVal start As Long = -1, Optional ByVal Count As Long = 1, _
    Optional ByVal CmpMethod As VbCompareMethod = vbBinaryCompare) As StringBuffer
Attribute ReplaceStr.VB_HelpID = 2148
    Dim lSeek As Long, lReplace As Long
    If Not Initialized Then AutoInit
    Dim mSB As New StringBuffer
    
    lSeek = Len(StrSeek)
    lReplace = Len(StrReplace)
    
    Ensure lSeek > 0 And lSeek <= pLength, Me, "ReplaceStr", "Invalid SeekString length"
    If pLength > 0 Then
        mSB.Value = Replace(Left$(pStr, pLength), StrSeek, StrReplace, start, Count, CmpMethod)
    End If
    Set ReplaceStr = mSB
End Function

Private Sub SetSize(ByVal Length As Long)

pLength = Length
pBlocksCount = Int(Length / pBlockSize) + 1
pStr = String$(pBlocksCount * pBlockSize, Chr$(0))

End Sub

Private Sub UpdateSize(ByVal Length As Long)
' mantiene sempre il contenuto
Dim RealSize As Long
Dim BlocksToAdd As Long, NewBlockCount As Long

pLength = Length
RealSize = pBlocksCount * pBlockSize

If pLength > RealSize Then
    NewBlockCount = Int(Length / pBlockSize) + 1
    BlocksToAdd = NewBlockCount - pBlocksCount
    If pAddedBlocksCountSTAT <= 0 Then
        pAddedBlocksCountSTAT = BlocksToAdd
    Else
        If BlocksToAdd >= pAddedBlocksCountSTAT Then
            pAddedBlocksCountSTAT = BlocksToAdd
        Else
            pAddedBlocksCountSTAT = pAddedBlocksCountSTAT - 1
            BlocksToAdd = pAddedBlocksCountSTAT
        End If
    End If
    Ensure BlocksToAdd > 0, Me, "UpdateSize", "Unexpected condition: unable to allocate StringBuffer blocks"
    pBlocksCount = NewBlockCount
    RealSize = pBlocksCount * pBlockSize
    pStr = pStr & String$(pBlockSize * BlocksToAdd, Chr$(0))
End If

End Sub

