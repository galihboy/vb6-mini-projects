VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleEncrypter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2049
Option Explicit

Public Function XorKey(Text As String, ByVal Key As String) As String
Attribute XorKey.VB_HelpID = 2050
Dim i As Long, TextLen As Long
Dim KeyLen As Long
Const Mask16 As Long = &HFFFF
Dim vKey() As Long

TextLen = Len(Text)
KeyLen = Len(Key)
Ensure KeyLen > 0, Me, "Transform", "Empty key"

If TextLen > 0 Then
    ReDim vKey(KeyLen - 1)
    For i = 0 To KeyLen - 1
        vKey(i) = AscW(Mid$(Key, i + 1, 1))
    Next i
    XorKey = String$(TextLen, Chr$(0))
    For i = 1 To TextLen
        Mid$(XorKey, i, 1) = ChrW$(Mask16 And (AscW(Mid$(Text, i, 1)) Xor vKey(i Mod KeyLen)))
    Next i
End If

End Function


