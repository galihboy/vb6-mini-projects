VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Escaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2300
'---------------------------------------------------------------------------------------
' Module    : Modulo di classe:Escaper
' DateTime  : 05/12/2001 10.23
' Author    : Administrator
' Purpose   : Manage generic character escaping
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum EscMethod
    EscCharsOfSequence
    EscCharsOutOfSequence
End Enum

Private htEC As New HashTable
Private pExclude As EscMethod
Private pEscChar As String * 1
Private pInitDone As Boolean

Private Const HexLen As Long = 4

'---------------------------------------------------------------------------------------
' Procedure : Init
' DateTime  : 05/12/2001 10.24
' Author    : Administrator
' Purpose   : The object constructor; specifies wich chars needs to be escaped
'               you can specify chars to be escaped or chars not to be escaped
'               as you define in EscMode parameter; Esc Char must be escaped
' Usage :
'   Dim eSC as new Escaper
'   eSC.Init "0123456789",EscCharsOfSequence
'---------------------------------------------------------------------------------------

Public Sub Init(ByVal CharsSequence As String, _
    Optional ByVal EscMode As EscMethod = EscCharsOfSequence, _
    Optional ByVal EscChar As String = "\")
Attribute Init.VB_HelpID = 2301
Dim l As Long, i As Long
Dim ch As String * 1

l = Len(CharsSequence)
Ensure l > 0, Me, "Init", "Escape chars sequence is empty"
Ensure Len(EscChar) = 1, Me, "Init", "Invalid escape char"

' il carattere di escape è tra quelli esclusi
If EscMode = EscCharsOutOfSequence Then
    Ensure InStr(1, CharsSequence, EscChar) = 0, Me, "Init", "Need to escape escape character"
End If

htEC.RemoveAll
htEC.SetSize (l)

For i = 1 To l
    ch = Mid$(CharsSequence, i, 1)
    If htEC.Exists(ch) Then
        Warn Me, "Init", "Char '" & ch & "' is already present in sequence"
    Else
        htEC.Add ch, i
    End If
Next i

If EscMode = EscCharsOfSequence Then
    If Not htEC.Exists(EscChar) Then htEC.Add EscChar, i + 1
End If

pExclude = EscMode
pEscChar = EscChar
pInitDone = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EscapeSTR
' DateTime  : 05/12/2001 10.28
' Author    : Administrator
' Purpose   : Escapes a String
' Usage :
'   s = eSC.EscapeSTR ("pippo2001")
'---------------------------------------------------------------------------------------

Public Function EscapeSTR(ByVal Str As String) As String
Attribute EscapeSTR.VB_HelpID = 2302
Ensure pInitDone, Me, "EscapeSTR", "Escaper not initialized"
Dim i As Long, LStr As Long
Dim k As Long, lToken As Long
Dim OutStr As String, OutPos As Long, OutLen As Long
Dim ch As String * 1, Token As String
LStr = Len(Str)

If LStr > 0 Then
    OutLen = LStr * 2
    OutStr = Space$(OutLen)
    OutPos = 1
    For i = 1 To LStr
        ch = Mid$(Str, i, 1)
        Token = ch
        If pExclude = EscCharsOfSequence Then
            If htEC.Exists(ch) Then Token = pEscChar & FHex(AscW(ch), HexLen)
        Else
            If Not htEC.Exists(ch) Then Token = pEscChar & FHex(AscW(ch), HexLen)
        End If
        lToken = Len(Token)
        k = OutPos
        OutPos = OutPos + lToken
        If OutPos > OutLen Then
            OutStr = OutStr & Space$(lToken)
            OutLen = OutLen + lToken
        End If ' ensure to have enough space
        Mid$(OutStr, k, lToken) = Token
    Next i
    EscapeSTR = Left(OutStr, OutPos - 1)
End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : UnEscapeSTR
' DateTime  : 05/12/2001 10.29
' Author    : Administrator
' Purpose   : Is the inverse function of EscapeSTR
' Usage :
'   s = "pippo2001"
'   debug.print eSC.UnEscapeSTR( eSC.EscapeSTR (s))
'---------------------------------------------------------------------------------------
Public Function UnEscapeSTR(ByVal Str As String) As String
Attribute UnEscapeSTR.VB_HelpID = 2303
Ensure pInitDone, Me, "EscapeSTR", "Escaper not initialized"
Dim Pos As Long
Dim HexCode As String
Dim LStr As Long

Pos = 1
LStr = Len(Str)
Do
    Pos = InStr(Pos, Str, pEscChar)
    If Pos > 0 Then
        If Pos + HexLen <= LStr Then
            HexCode = Mid(Str, Pos + 1, HexLen)
            Str = Left(Str, Pos - 1) & Replace(Str, pEscChar & HexCode, ChrW(CLng("&H" & HexCode)), Pos, 1)
            LStr = LStr - HexLen
            Pos = Pos + 1
        Else
            Pos = 0
        End If
    End If
Loop Until Pos = 0

UnEscapeSTR = Str

End Function

Private Function FHex(ByVal Number As Long, Optional ByVal TotChars As Long = 4) As String
Dim h As String
h = Hex(Number)
FHex = UCase(String(TotChars - Len(h), "0") & h)
End Function


