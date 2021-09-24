VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2034
' ------------------------------------------------------------------------
'  The CBITARRAY class
'
'  simiulates an array of Boolean values
'  saves memory by packing one element in one bit
'
' IMPORTANT: you make make ITEM the default member for this class
'            do this from inside the Tools | Procedure Attributes dialog box
'
' Usage:
'    Dim bitArr As New CBitArray
'    bitArr.Init(10000)         ' 10,000 boolean elements
'    ' assign a value
'    bitArr(10) = True
'    ' read it back
'    Print bitArr(10)
'
' ------------------------------------------------------------------------

Option Explicit

Dim mask(0 To 7) As Integer
Dim Values() As Byte

Public Sub Init(ByVal numEls As Long)
Attribute Init.VB_HelpID = 2035
    ' redimension the internal array
    ' MUST BE THE FIRST METHOD CALLED FOR THIS CLASS
    ReDim Values(numEls \ 8) As Byte
End Sub

Private Sub Class_Initialize()
    ' initialize the mask() array
    Dim i As Integer
    mask(0) = 1
    For i = 1 To 7
        mask(i) = mask(i - 1) * 2
    Next
End Sub

' for smoother syntax, you should make ITEM the default member for this class

Property Get Item(ByVal Index As Long) As Boolean
Attribute Item.VB_HelpID = 2036
    ' retrieve an array item
    Item = Values(Index \ 8) And mask(Index And 7)
End Property

Property Let Item(ByVal Index As Long, ByVal new_Item As Boolean)
    Dim ndx As Long, M As Byte
    ' cache the mask into a local variable
    M = mask(Index And 7)
    ndx = Index \ 8
    ' set or reset only the relevant bit
    Values(ndx) = (Values(ndx) And Not M) Or (new_Item And M)
End Property


