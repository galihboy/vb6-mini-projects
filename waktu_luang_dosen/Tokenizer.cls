VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2304
'---------------------------------------------------------------------------------------
' Module    : Modulo di classe:Tokenizer
' DateTime  : 05/12/2001 10.32
' Author    : Administrator
' Purpose   : Implements Tokenize functions
'---------------------------------------------------------------------------------------
Option Explicit

' tokenizer tool
Private vbLT As New LangTool

'---------------------------------------------------------------------------------------
' Procedure : Tokenize
' DateTime  : 05/12/2001 10.33
' Author    : Administrator
' Purpose   : Splits a string in Tokens, unsing a separator char; can skip sequences
'               of separatos chars
' Usage :
'   dim tNK as new Tokenizer
'   dim Tokens() as string
'   Tokens = tNK.Tokenize(SourceStr," ")
'---------------------------------------------------------------------------------------

Public Function Tokenize(ByVal Source As String, _
    ByVal CharSep As String, Optional ByVal SkipSepCharsSequences As Boolean = True) As String()
Attribute Tokenize.VB_HelpID = 2305

Dim i As Long, k As Long
Dim lenSource As Long
Dim NewToken As String, TokensCount As Long
Dim BlocksCount As Long
Const BlockSize As Long = 256
Dim retval() As String

Ensure Len(CharSep) = 1, Me, "Tokenizer", "Invalid separator char"
Ensure Len(Source) > 0, Me, "Tokenizer", "Cannot tokenize an empty string"

lenSource = Len(Source)

Do
    i = i + 1
Loop While (Mid$(Source, i, 1) = CharSep) And SkipSepCharsSequences

While i <= lenSource
    k = InStr(i, Source, CharSep)
    If k = 0 Then
        ReDim Preserve retval(0 To TokensCount)
        retval(TokensCount) = Mid(Source, i)
        Tokenize = retval
        Exit Function
    Else
        NewToken = Mid$(Source, i, k - i)
        If TokensCount >= (BlockSize * BlocksCount) Then
            BlocksCount = BlocksCount + 1
            ReDim Preserve retval(0 To BlocksCount * BlockSize)
        End If
        retval(TokensCount) = NewToken
        TokensCount = TokensCount + 1
        i = k
        Do
            i = i + 1
        Loop While (Mid$(Source, i, 1) = CharSep) And SkipSepCharsSequences
    End If
Wend
If TokensCount > 0 Then ReDim Preserve retval(0 To TokensCount - 1)
Tokenize = retval

End Function

'---------------------------------------------------------------------------------------
' Procedure : JoinTokens
' DateTime  : 05/12/2001 10.35
' Author    : Administrator
' Purpose   : Joins a Subset of a string array in one string; needs a zero
'               or one based array
' Usage :
'   dim s as string, tNK as new Tokenizer
'   dim Tokens() as string
'   Tokens = tNK.Tokenize(TextName.text)
'   s = JoinTokens(Tokens,,ubound(Tokens)-1,"_")
'---------------------------------------------------------------------------------------

Public Function JoinTokens(Tokens() As String, Optional ByVal StartFrom As Long = -1, _
    Optional ByVal EndTo As Long = -1, Optional ByVal Delimiter As String = " ") As String
Attribute JoinTokens.VB_HelpID = 2306
Dim lB As Long, uB As Long, Size As Long
Dim jStart As Long, jEnd As Long
Dim SB As New StringBuffer, i As Long
Ensure Not vbLT.IsEmptyArray(Tokens), Me, "JoinTokens", "Empty Array"

lB = LBound(Tokens)
uB = UBound(Tokens)
Size = uB - lB + 1
If StartFrom < 0 Then
    jStart = lB
Else
    jStart = StartFrom
End If
    
If EndTo < 0 Then
    jEnd = uB
Else
    jEnd = EndTo
End If
Ensure jStart <= jEnd, Me, "JoinTokens", "Invalid range"
Ensure jStart >= lB, Me, "JoinTokens", "Invalid start index"
Ensure jEnd <= uB, Me, "JoinTokens", "Invalid end index"
For i = jStart To jEnd
    SB.Append Tokens(i)
    If i < jEnd Then SB.Append Delimiter
Next i

JoinTokens = SB.Value

End Function

'-----------
