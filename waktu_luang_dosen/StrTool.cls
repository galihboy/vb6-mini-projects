VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StrTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2333
Option Explicit

Private vbLT As New LangTool, vbTK As New Tokenizer

' ----------------------------------------------------------------------------
' Procedure : StrEquality
' DateTime  : 10/12/2001 16.51
' Author    : Administrator
' Purpose   : Compare two strings and return the equality level
' Usage :
'   dim Phrase as string, ListTitles () as string
'   ... loads ListTitles and Phrase ...
'   debug.print StrEquality("pippo","pluto paperino topolino minnie...")
' Use LengthCompare = false if you are seeking a word in a long documet, and
' documet lenght is not relevant, but the fact it contains similar words
' Use PositionWeight=0 if don't care about different order in words mathing.
'---------------------------------------------------------------------------------------
Public Function StrEquality(ByVal Str1 As String, ByVal Str2 As String, _
    Optional ByVal Delimiter As String = " ", Optional ByVal LengthCompare As Boolean, _
    Optional PositionWeight As Single = 0.25, _
    Optional CmpMethod As VbCompareMethod = vbTextCompare) As Double
Attribute StrEquality.VB_HelpID = 2334
    
Ensure Len(Delimiter) = 1, Me, "StrEquality", "Invalid delimiter"
    
Dim v1() As String, v2() As String, Match As Variant
Dim ValidLen As Long, Threshold As Double
Dim TotChars As Long
Dim Weight As Double
Dim CmpRes As Double, htCmpBest As New HashTable
Dim CmpBestLen As Long, CmpKey As String
Dim i As Long, k As Long, v As Variant

Ensure PositionWeight >= 0#, Me, "SteEquality", "Invalid position weight"

v1 = vbTK.Tokenize(Str1, Delimiter)
Ensure Not vbLT.IsEmptyArray(v1), Me, "StrEquality", "Empty parameter Str1"
    
v2 = vbTK.Tokenize(Str2, Delimiter)
Ensure Not vbLT.IsEmptyArray(v2), Me, "StrEquality", "Empty parameter Str2"

If LengthCompare Then
    ValidLen = vbLT.MaxValue(Array(vbLT.ArraySize(v1), vbLT.ArraySize(v2)))
Else
    ValidLen = vbLT.MinValue(Array(vbLT.ArraySize(v1), vbLT.ArraySize(v2)))
End If

With htCmpBest
    For i = 0 To UBound(v1)
        For k = 0 To UBound(v2)
            CmpRes = TokensEquality(v2(k), v1(i), CmpMethod) * (1 / (1 + PositionWeight * Abs(i - k)))
            If v1(i) > v2(k) Then
                CmpKey = v1(i) & v2(k)
            Else
                CmpKey = v2(k) & v1(i)
            End If
            If .Exists(CmpKey) Then
                If CmpRes > .Item(CmpKey) Then
                    .Item(CmpKey) = CmpRes
                End If
            Else
                .Item(CmpKey) = CmpRes
            End If
        Next k
    Next i
    
    Match = .Values
    
    vbLT.QuickSort Match, , True
    ValidLen = vbLT.MinValue(Array(ValidLen, vbLT.ArraySize(Match)))
    Threshold = Match(ValidLen - 1)
    For Each v In .Keys
        If .Item(v) >= Threshold Then
            TotChars = TotChars + Len(v)
            Weight = Weight + Len(v) * .Item(v)
        End If
    Next
End With
StrEquality = Weight / TotChars
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : TokensEquality
' DateTime  : 10/12/2001 19.27
' Author    : Administrator
' Purpose   : compares two tokens and returns a value between 0,1
'   0 means completely different
'   1 means perfecly equal
' Usage :
'---------------------------------------------------------------------------------------
Public Function TokensEquality(ByVal Token1 As String, _
    ByVal Token2 As String, ByVal CmpMethod As VbCompareMethod) As Double
Attribute TokensEquality.VB_HelpID = 2335
Dim W1 As Long, W2 As Long, t As String
Dim i As Long, k As Long

W1 = Len(Token1): W2 = Len(Token2)
Ensure W1 > 0 And W2 > 0, Me, "TokensEquality", "Cannot compare empty tokens"
If W2 > W1 Then
    t = Token1
    Token1 = Token2
    Token2 = t
    W1 = Len(Token1): W2 = Len(Token2)
End If
For i = 1 To W2
    If StrComp(Mid$(Token1, i, 1), Mid$(Token2, i, 1)) = 0 Then
        k = k + W2 - i + 1
    End If
Next i

TokensEquality = k / (W1 * (W1 + 1) / 2)
End Function

'---------------------------------------------------------------------------------------
' Procedure : StrHash
' DateTime  : 02/01/2002 17.47
' Description : retrieves an hash code for a string from 0 to RangeUpper-1
'---------------------------------------------------------------------------------------
Public Function StrHash(StrValue As String, Optional ByVal RangeUpper As Long = 32768) As Long
Attribute StrHash.VB_HelpID = 2336
Dim c As Long, i As Long, l As Long, k As Long, n As Long
l = Len(StrValue)
For i = 1 To l
    c = AscW(Mid$(StrValue, i, 1))
    k = (k + c) Mod RangeUpper
    n = (n + c Mod i) Mod l
Next i
StrHash = (k + n) Mod RangeUpper
End Function


