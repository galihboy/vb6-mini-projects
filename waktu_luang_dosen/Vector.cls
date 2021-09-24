VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2228
Attribute VB_Description = "Implements a Vector class that wraps a vb array; exposes Seeking, Sorting and many other functionalities."
' ##MODULE_DESCRIPTION Implements a Vector class that wraps a vb array; _
exposes Seeking, Sorting and many other functionalities.

'---------------------------------------------------------------------------------------
' Module    : Vector
' DateTime  : 03/12/2001 13.21
' Author    : Administrator
' Purpose   : Implements a Vector class that wraps a vb array;
'               exposes Seeking and Sorting functionality;
'               You can set AutoGrow to 0 if if want to work with a fixed array,
'               or to a positive value ( default is 256 ); in this case the array
'               will grow automatically. Notes the difference betweeen
'               UsedLength ad Length: the first is the maximum  lenght of the
'               non empty values sequence, the second is the length of the real array.
'               These two properties will be equal if the array is compacted.
'               Example:
'
'               Dim vStr as new Vector
'               Dim StrArray() as string
'               with vStr
'                   .Init vtString
'                  do: loop until (.AddItem (...) = 1000)
'                   .Compact
'                   .Sort
'                   StrArray = .Converter.CStrArray(.value)
'                   .Clear
'               end with
'
'
'---------------------------------------------------------------------------------------
Option Explicit

Private v As Variant
Private pItemsType As enVectorType
Private pInitialized As Boolean
Private pAutoGrowSize As Long
Private pUpIndex As Long
Private pLowIndex As Long
Private pUpLowIndexOn As Boolean
Private pFragmentation As Long
Private vbLT As New LangTool

Public Enum enVectorType
    vtInteger = 2
    vtLong = 3
    vtSingle = 4
    vtDouble = 5
    vtCurrency = 6
    vtDate = 7
    vtString = 8
    vtObject = 9
    vtBoolean = 11
    vtVariant = 12
    vtDataObject = 13
    vtDecimal = 14
    vtByte = 17
    vtUserDefinedType = 36
    vtArray = 8192
End Enum

Public Enum enGrowMode
    vtGrowUp
    vtGrowDown
    vtGrowBoth
End Enum

' ##ENUMERATION_MEMBER_DESCRIPTION vtUndefined Undefined array ( not dimensioned )
' ##ENUMERATION_MEMBER_DESCRIPTION vtClean Empty array ( all items are empty )
' ##ENUMERATION_MEMBER_DESCRIPTION vtUnFilledFragmented The array is not filled, and the central block has some _
empty items inside.
' ##ENUMERATION_MEMBER_DESCRIPTION vtUnFilledCompacted The array is not filled, but the central block has no _
empty items inside.
' ##ENUMERATION_MEMBER_DESCRIPTION vtFilledFragmented The array is filled, but there are some empty items.
' ##ENUMERATION_MEMBER_DESCRIPTION vtFilledCompacted The array is filled with no empty items.


Public Enum enVectorState
    vtUndefined
    vtClean
    vtUnFilledFragmented
    vtUnFilledCompacted
    vtFilledFragmented
    vtFilledCompacted
End Enum

' ##VARIABLE_DESCRIPTION Converter Exposes methods to convert a generic variant array in a typed array.
Public Converter As New ArrayCnv
Attribute Converter.VB_VarHelpID = 2229
Attribute Converter.VB_VarDescription = "Exposes methods to convert a generic variant array in a typed array."

'---------------------------------------------------------------------------------------
' Procedure : Grow
' DateTime  : 03/12/2001 13.30
' Author    : Administrator
' Purpose   : Increses the size of the Vector
'---------------------------------------------------------------------------------------
'
Public Sub Grow(ByVal Size, Optional ByVal GrowMode As enGrowMode = vtGrowUp)
Attribute Grow.VB_Description = "Enlarges the array"
Attribute Grow.VB_HelpID = 2230
'##PARAMETER_DESCRIPTION Size The number of elements to append
'##PARAMETER_DESCRIPTION GrowMode You can grow up, down or in both directions
'##BLOCK_DESCRIPTION Enlarges the array

Ensure pInitialized, Me, "Grow", "Vector not initilized"
Ensure Size > 0, Me, "Grow", "Size must be positive"
Ensure vbLT.SequenceSeek(GrowMode, vbTextCompare, vtGrowUp, vtGrowDown, vtGrowBoth) >= 0, _
    Me, "Grow", "Invalid GrowMode: " & GrowMode
Dim lB As Long, uB As Long
Dim NewV As Variant, i As Long
If EmptyVector Then
    uB = 0
    lB = 0
Else
    lB = LBound(v)
    uB = UBound(v)
End If

Select Case GrowMode
    Case vtGrowUp:
        ReDim NewV(lB To uB + Size)
    Case vtGrowDown:
        ReDim NewV(lB - Size To uB)
    Case vtGrowBoth:
        ReDim NewV(lB - Size To uB + Size)
    Case Else:
        Throw Me, "Grow", "Unexpected condition: wrong GrowMode"
End Select
If CopyVector(NewV) Then
    SetVector NewV
Else
    Throw Me, "Grow", "Unable to grow"
End If
End Sub

Private Function CopyVector(ByRef NewVector As Variant) As Boolean
Dim lB As Long, uB As Long, i As Long
If Not EmptyVector Then
    lB = LBound(v)
    uB = UBound(v)

    For i = lB To uB
        If Not IsEmpty(v(i)) Then
            If pItemsType = vtObject Then
                Set NewVector(i) = vbLT.VarTypeCnv(v(i), pItemsType)
            Else
                NewVector(i) = vbLT.VarTypeCnv(v(i), pItemsType)
            End If
        End If
    Next i
    CopyVector = True
End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : State
' DateTime  : 03/12/2001 13.32
' Author    : Administrator
' Purpose   : Retrieve the Vector state;
'               - Undefined if the vector is not initialized or empty
'               - UnFilledFragmented if the Vector is not compacted
'               - UnFilledCompacted if the Vector is not filled but there a
'                   filled sequence of not empty values
'               - FilledFragmented if the vector is compacted but there are some
'                   empty values
'               - FilledCompacted if the vector is filled of not empty values
'---------------------------------------------------------------------------------------
Public Property Get State() As enVectorState
Attribute State.VB_Description = "Retrieves the state of the array"
Attribute State.VB_HelpID = 2231
'##BLOCK_DESCRIPTION Retrieves the state of the array
Ensure pInitialized, Me, "State", "Vector not initilized"
If EmptyVector Then
    State = vtUndefined
Else
    If pUpLowIndexOn Then
        If (pUpIndex = UBound(v)) And (pLowIndex = LBound(v)) Then
            If pFragmentation > 0 Then
                State = vtFilledFragmented
            Else
                State = vtFilledCompacted
            End If
        Else
            If pFragmentation > 0 Then
                State = vtUnFilledFragmented
            Else
                State = vtUnFilledCompacted
            End If
        End If
    Else
        State = vtClean
    End If
End If

End Property

'---------------------------------------------------------------------------------------
' Procedure : AutoGrowSize
' DateTime  : 03/12/2001 15.22
' Author    : Administrator
' Purpose   : The size of the minimum grow amount
'---------------------------------------------------------------------------------------
Public Property Get AutoGrowSize() As Long
Attribute AutoGrowSize.VB_Description = "Sets the automatic grow size of the vector. If zero, the array if fixed size.Sets the automatic grow size of the vector. If zero, the array if fixed size."
Attribute AutoGrowSize.VB_HelpID = 2232
Ensure pInitialized, Me, "Get AutoGrowSize", "Vector not initilized"
AutoGrowSize = pAutoGrowSize
End Property

Public Property Let AutoGrowSize(ByVal NewGrowSize As Long)
'##BLOCK_DESCRIPTION Sets the automatic grow size of the vector. _
If zero, the array if fixed size.
Ensure pInitialized, Me, "Let AutoGrowSize", "Vector not initilized"
Ensure NewGrowSize > 0, Me, "Let AutoGrowSize", "Invalid Grow Size"

pAutoGrowSize = NewGrowSize

End Property


'---------------------------------------------------------------------------------------
' Procedure : ItemsType
' DateTime  : 03/12/2001 13.36
' Author    : Administrator
' Purpose   : Specifies the Type of the Vectors items
'---------------------------------------------------------------------------------------
Public Property Get ItemsType() As enVectorType
Attribute ItemsType.VB_Description = "The type of the vector"
Attribute ItemsType.VB_HelpID = 2233
'##BLOCK_DESCRIPTION The type of the vector
Ensure pInitialized, Me, "Get ItemsType", "Vector not initilized"
ItemsType = pItemsType
End Property

'---------------------------------------------------------------------------------------
' Procedure : ItemsTypeName
' DateTime  : 17/12/2001 12.36
' Description : Specifies the Type name of the Vectors items ( eg. "Array of variant" )
'---------------------------------------------------------------------------------------
Public Property Get ItemsTypeName() As String
Attribute ItemsTypeName.VB_Description = "The name of the type ( eg. "
Attribute ItemsTypeName.VB_HelpID = 2234
'##BLOCK_DESCRIPTION The name of the type ( eg. "Array of variant" )
Ensure pInitialized, Me, "Get ItemsType", "Vector not initilized"
ItemsTypeName = vbLT.VarConstTypeName(pItemsType)
End Property

'---------------------------------------------------------------------------------------
' Procedure : Item
' DateTime  : 03/12/2001 15.22
' Author    : Administrator
' Purpose   : Returns The Vector item of the required index
'---------------------------------------------------------------------------------------

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_Description = "Returns The Vector item of the required index: this is the default property.  "
Attribute Item.VB_HelpID = 2235
Attribute Item.VB_UserMemId = 0
'##BLOCK_DESCRIPTION Returns The Vector item of the required index: this is the _
default property.
Ensure pInitialized, Me, "Get Item", "Vector not initilized"
Ensure vbLT.Between(Index, LowBound, UpperBound), Me, "Get Item", "Index out of range"
If pItemsType = vtObject Then
    Set Item = v(Index)
Else
    Item = v(Index)
End If

End Property

'---------------------------------------------------------------------------------------
' Procedure : AddItem
' DateTime  : 17/12/2001 12.39
' Description : Sets the current item and increases the counter
'---------------------------------------------------------------------------------------
Public Function AddItem(ByVal Value As Variant) As Long
Attribute AddItem.VB_Description = "Sets the current item and increases the counter"
Attribute AddItem.VB_HelpID = 2236
'##BLOCK_DESCRIPTION Sets the current item and increases the counter
With Me
    .Item(.UsedLength) = Value
    AddItem = .UsedLength
End With
End Function

Public Property Let Item(ByVal Index As Long, ByVal Value As Variant)
Ensure pInitialized, Me, "Let Item", "Vector not initilized"
Dim lB As Long, uB As Long, k As Long, j As Long, xValue As Variant
lB = LBound(v): uB = UBound(v)
Ensure (pAutoGrowSize > 0) Or vbLT.Between(Index, lB, uB), Me, "Let Item", "Index out of range"
Ensure Not IsObject(Value), Me, "Let Index", "Object not managed by Vector"
Ensure Not IsArray(Value), Me, "Let Index", "Array Items not managed by Vector"
Ensure VarType(Value) = pItemsType, Me, "Let Item", "Invalid Type"

If Index < lB Then j = Index - lB
If Index > uB Then j = Index - uB
If j <> 0 Then
    k = Fix(Sgn(j) + j / pAutoGrowSize)
    If j < 0 Then
        ReDim Preserve v(lB + k * pAutoGrowSize To uB)
    Else
        ReDim Preserve v(lB To uB + k * pAutoGrowSize)
    End If
End If
j = Index - pLowIndex
k = Index - pUpIndex

If pUpLowIndexOn Then
    If k > 0 Then
        pUpIndex = Index
        pFragmentation = pFragmentation + IIf(k > 1, k, 0)
    End If
    If j < 0 Then
        pLowIndex = Index
        pFragmentation = pFragmentation + IIf(Abs(j) > 1, Abs(j), 0)
    End If
Else
    pUpLowIndexOn = True
    pUpIndex = Index
    pLowIndex = Index
    pFragmentation = 0
End If
If pItemsType = vtObject Then
    Set v(Index) = vbLT.VarTypeCnv(Value, pItemsType)
Else
    v(Index) = vbLT.VarTypeCnv(Value, pItemsType)
End If
End Property

Public Property Get LowBound() As Long
Attribute LowBound.VB_Description = "The real lower bound"
Attribute LowBound.VB_HelpID = 2237
'##BLOCK_DESCRIPTION The real lower bound
Ensure pInitialized, Me, "Get LowBound", "Vector not initilized"
LowBound = LBound(v)
End Property

Public Property Get UpperBound() As Long
Attribute UpperBound.VB_Description = "The real Upper bound"
Attribute UpperBound.VB_HelpID = 2238
'##BLOCK_DESCRIPTION The real Upper bound
Ensure pInitialized, Me, "Get UpperBound", "Vector not initilized"
UpperBound = UBound(v)
End Property

'---------------------------------------------------------------------------------------
' Procedure : Value
' DateTime  : 03/12/2001 13.37
' Author    : Administrator
' Purpose   : Is the array, passed as a Variant
'---------------------------------------------------------------------------------------
Public Property Get Value() As Variant
Attribute Value.VB_Description = "Is the array, passed as a Variant  "
Attribute Value.VB_HelpID = 2239
'##BLOCK_DESCRIPTION Is the array, passed as a Variant
Ensure pInitialized, Me, "Get Value", "Vector not initilized"
Value = v
End Property

Public Property Let Value(ByVal NewVector As Variant)
Ensure pInitialized, Me, "Let Value", "Vector not initilized"
SetVector NewVector
End Property

Private Function SetVector(NewVector As Variant) As Boolean
Ensure IsArray(NewVector), Me, "AnalizeVector", "Needs a Vector"
Ensure Not vbLT.IsEmptyArray(NewVector), Me, "AnalizeVector", "Empty array passed"

Dim i As Long, UsedZone As Boolean, LastUsed
Dim LowerUsed As Long, UpperUsed As Long
Dim Fragmentation As Long
For i = LBound(NewVector) To UBound(NewVector)
    Ensure vbLT.SequenceSeek(VarType(NewVector(i)), vbTextCompare, pItemsType, vbEmpty) >= 0, _
        Me, "Let Value", "Invalid Value array type"
    If Not IsEmpty(NewVector(i)) Then
        If UsedZone Then
            UpperUsed = i
        Else
            UsedZone = True
            LowerUsed = i
            UpperUsed = i
        End If
    Else
        If UsedZone Then Fragmentation = Fragmentation + 1
    End If
Next i

pUpLowIndexOn = UsedZone
pLowIndex = LowerUsed
pUpIndex = UpperUsed
pFragmentation = Fragmentation
v = NewVector

End Function

'---------------------------------------------------------------------------------------
' Procedure : Clear
' DateTime  : 03/12/2001 15.22
' Author    : Administrator
' Purpose   : Erases the Vector content
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
Attribute Clear.VB_Description = "Erases the Vector content"
Attribute Clear.VB_HelpID = 2240
'##BLOCK_DESCRIPTION Erases the Vector content
Ensure pInitialized, Me, "Clear", "Vector not initilized"
On Error Resume Next
pUpIndex = 0
pLowIndex = 0
pUpLowIndexOn = False
pFragmentation = 0
ReDim v(0)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SeekValue
' DateTime  : 03/12/2001 13.38
' Author    : Administrator
' Purpose   : Returns true if the value is present in the Vector
'---------------------------------------------------------------------------------------
Public Function SeekValue(ByVal Value As Variant, Optional ByVal FromIndex, _
    Optional ByVal ToIndex, Optional CmpMethod As VbCompareMethod = vbTextCompare) As Boolean
Attribute SeekValue.VB_Description = "Returns true if the value is present in the Vector"
Attribute SeekValue.VB_HelpID = 2241
'##BLOCK_DESCRIPTION Returns true if the value is present in the Vector
Ensure pInitialized, Me, "SeekValueIndex", "Vector not initilized"
Ensure Not EmptyVector, Me, "SeekValue", "Empty array"
If IsMissing(FromIndex) Then FromIndex = LBound(v)
If IsMissing(ToIndex) Then ToIndex = UBound(v)
Ensure FromIndex <= ToIndex, Me, "SeekValue", "Invalid index range"
Dim i As Long, sVal As String
sVal = CStr(Value)

For i = FromIndex To ToIndex
    If StrComp(CStr(v(i)), sVal, CmpMethod) = 0 Then
        SeekValue = True
        Exit Function
    End If
Next i

End Function

'---------------------------------------------------------------------------------------
' Procedure : SeekValueIndexArray
' DateTime  : 03/12/2001 13.38
' Author    : Administrator
' Purpose   : Returns the array of the indexes sequence of the found values;
'               EG:
'               If vLong.SeekValueIndexArray(129).Length=0 then ' not found
'---------------------------------------------------------------------------------------
Public Function SeekValueIndexArray(ByVal Value As Variant, Optional ByVal FromIndex, _
    Optional ByVal ToIndex, Optional CmpMethod As VbCompareMethod = vbTextCompare) As Vector
Attribute SeekValueIndexArray.VB_Description = "Returns the array of the indexes sequence of the found values; EG: If vLong.SeekValueIndexArray(129).Length=0 then ' not found"
Attribute SeekValueIndexArray.VB_HelpID = 2242
'##BLOCK_DESCRIPTION Returns the array of the indexes sequence of the found values; _
EG: If vLong.SeekValueIndexArray(129).Length=0 then ' not found

' ritorna il vettore di tutte le occorrenze trovate di
Ensure pInitialized, Me, "SeekValueIndex", "Vector not initilized"
Ensure Not EmptyVector, Me, "SeekValue", "Empty array"
If IsMissing(FromIndex) Then FromIndex = LBound(v)
If IsMissing(ToIndex) Then ToIndex = UBound(v)
Ensure FromIndex <= ToIndex, Me, "SeekValue", "Invalid index range"
Dim i As Long, rVect As New Vector, sVal As String

sVal = CStr(Value)
With rVect
    .Init vtLong
    For i = FromIndex To ToIndex
        If StrComp(CStr(v(i)), sVal, CmpMethod) = 0 Then
            .AddItem i
        End If
    Next i
    If .UsedLength > 0 Then .Compact
End With
Set SeekValueIndexArray = rVect

End Function

'---------------------------------------------------------------------------------------
' Procedure : EmptyVector
' DateTime  : 03/12/2001 15.21
' Author    : Administrator
' Purpose   : Returns True if the Vector is emtpy
'---------------------------------------------------------------------------------------
'
Private Function EmptyVector() As Boolean
Attribute EmptyVector.VB_Description = "Returns True if the Vector is emtpy"
'##BLOCK_DESCRIPTION Returns True if the Vector is emtpy
Ensure pInitialized, Me, "Get IsEmpty", "Vector not initilized"
On Error GoTo Success
Dim i As Long
i = LBound(v)
i = UBound(v)
EmptyVector = False
Exit Function
Success:
EmptyVector = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : Sort
' DateTime  : 03/12/2001 13.40
' Author    : Administrator
' Purpose   : Sort the Vector; you can do in ascendig or descending mode;
'               You can exclude Empty values to be sorted, if the array
'               has no fragementation, from being considered.
'---------------------------------------------------------------------------------------
Public Sub Sort(Optional ByVal Descending As Boolean, _
    Optional ByVal CompactVector As Boolean = True)
Attribute Sort.VB_Description = "Sort the Vector; you can do in ascendig or descending mode; You can exclude Empty values to be sorted, if the array has no fragementation, from being considered."
Attribute Sort.VB_HelpID = 2243
'##BLOCK_DESCRIPTION Sort the Vector; you can do in ascendig or descending mode; _
You can exclude Empty values to be sorted, if the array _
has no fragementation, from being considered.
Ensure pInitialized, Me, "Sort", "Vector not initilized"
Ensure Not EmptyVector, Me, "Sort", "Empty array"
Ensure Sortable, Me, "Sort", "Cannot sort this vector: type is " & vbLT.VarConstTypeName(pItemsType)
If CompactVector Then Compact
QuickSort v, , Descending
End Sub

Private Sub QuickSort(VariantArray As Variant, Optional LastIndex As Variant, _
    Optional Descending As Boolean)

    Dim Value As Variant, temp As Variant
    Dim sp As Integer
    Dim leftStk(128) As Long, rightStk(128) As Long
    Dim leftNdx As Long, rightNdx As Long
    Dim i As Long, j As Long

    ' account for optional arguments
    If IsMissing(LastIndex) Then LastIndex = UBound(VariantArray)
    ' init pointers
    leftNdx = LBound(VariantArray)
    rightNdx = LastIndex
    ' init stack
    sp = 1
    leftStk(sp) = leftNdx
    rightStk(sp) = rightNdx

    Do
        If rightNdx > leftNdx Then
            Value = VariantArray(rightNdx)
            i = leftNdx - 1
            j = rightNdx
            ' find the pivot item
            If Descending Then
                Do
                    Do: i = i + 1: Loop Until VariantArray(i) <= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or VariantArray(j) >= Value
                    temp = VariantArray(i)
                    VariantArray(i) = VariantArray(j)
                    VariantArray(j) = temp
                Loop Until j <= i
            Else
                Do
                    Do: i = i + 1: Loop Until VariantArray(i) >= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or VariantArray(j) <= Value
                    temp = VariantArray(i)
                    VariantArray(i) = VariantArray(j)
                    VariantArray(j) = temp
                Loop Until j <= i
            End If
            ' swap found items
            temp = VariantArray(j)
            VariantArray(j) = VariantArray(i)
            VariantArray(i) = VariantArray(rightNdx)
            VariantArray(rightNdx) = temp
            ' push on the stack the pair of pointers that differ most
            sp = sp + 1
            If (i - leftNdx) > (rightNdx - i) Then
                leftStk(sp) = leftNdx
                rightStk(sp) = i - 1
                leftNdx = i + 1
            Else
                leftStk(sp) = i + 1
                rightStk(sp) = rightNdx
                rightNdx = i - 1
            End If
        Else
            ' pop a new pair of pointers off the stacks
            leftNdx = leftStk(sp)
            rightNdx = rightStk(sp)
            sp = sp - 1
            If sp = 0 Then Exit Do
        End If
    Loop
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Init
' DateTime  : 03/12/2001 13.41
' Author    : Administrator
' Purpose   : Initilizes the Vector; you can define the base (0 is default)
'               and the Type of the array and the AutoGrow size
'---------------------------------------------------------------------------------------
Public Sub Init(ByVal VectType As enVectorType, _
    Optional ByVal ArrayLength As Long, Optional ByVal BaseIndex As Long = 0, _
    Optional ByVal AutoGrowSize As Long = 256)
Attribute Init.VB_Description = "Initilizes the Vector; you can define the base (0 is default) and the Type of the array and the AutoGrow size"
Attribute Init.VB_HelpID = 2370
'##BLOCK_DESCRIPTION Initilizes the Vector; you can define the base (0 is default) _
and the Type of the array and the AutoGrow size
Ensure vbLT.SequenceSeek(VectType, vbTextCompare, vtInteger, vtLong, vtSingle, _
    vtDouble, vtCurrency, vtDate, vtString, vtObject, vtBoolean, vtVariant, vtDataObject, _
    vtDecimal, vtByte, vtUserDefinedType, vtArray) >= 0, Me, "Init", "Invalid Vector type"
    
Ensure ArrayLength >= 0, Me, "SetSize", "Invalid Array Length"
Ensure AutoGrowSize >= 0, Me, "Init", "Invalid AutoGrowsize"
Ensure Not pInitialized, Me, "Init", "Array is already initialized"
pAutoGrowSize = AutoGrowSize
pItemsType = VectType
If ArrayLength > 0 Then
    ReDim v(BaseIndex To BaseIndex + ArrayLength - 1)
Else
    ReDim v(0)
End If
pInitialized = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Length
' DateTime  : 03/12/2001 13.43
' Author    : Administrator
' Purpose   : The real length of the Array
'---------------------------------------------------------------------------------------
Public Property Get Length() As Long
Attribute Length.VB_Description = "The real length of the Array"
Attribute Length.VB_HelpID = 2245
'##BLOCK_DESCRIPTION The real length of the Array
Ensure pInitialized, Me, "Get Lenght", "Vector not initilized"
Length = UpperBound - LowBound + 1
End Property

'---------------------------------------------------------------------------------------
' Procedure : UsedLength
' DateTime  : 03/12/2001 13.43
' Author    : Administrator
' Purpose   : The maximum difference from the indexes of not empty values
'               plus one
'---------------------------------------------------------------------------------------
Public Property Get UsedLength() As Long
Attribute UsedLength.VB_Description = "The maximum difference from the indexes of not empty values plus one"
Attribute UsedLength.VB_HelpID = 2246
'##BLOCK_DESCRIPTION The maximum difference from the indexes of not empty values _
plus one
Ensure pInitialized, Me, "Get Lenght", "Vector not initilized"

If pUpLowIndexOn Then UsedLength = pUpIndex - pLowIndex + 1

End Property

'---------------------------------------------------------------------------------------
' Procedure : UsedUpperIndex
' DateTime  : 03/12/2001 13.44
' Author    : Administrator
' Purpose   : The maximum index used
'---------------------------------------------------------------------------------------
Public Property Get UsedUpperIndex() As Long
Attribute UsedUpperIndex.VB_Description = "The maximum index used"
Attribute UsedUpperIndex.VB_HelpID = 2247
'##BLOCK_DESCRIPTION The maximum index used
Ensure pInitialized, Me, "Get UpIndex", "Vector not initilized"
Ensure pUpLowIndexOn, Me, "Get UpIndex", "UpIndex not avaible"
UsedUpperIndex = pUpIndex
End Property

'---------------------------------------------------------------------------------------
' Procedure : UsedLowerIndex
' DateTime  : 03/12/2001 13.44
' Author    : Administrator
' Purpose   : The minimum index used
'---------------------------------------------------------------------------------------
Public Property Get UsedLowerIndex() As Long
Attribute UsedLowerIndex.VB_Description = "The minimum index used"
Attribute UsedLowerIndex.VB_HelpID = 2248
'##BLOCK_DESCRIPTION The minimum index used
Ensure pInitialized, Me, "Get LowIndex", "Vector not initilized"
Ensure pUpLowIndexOn, Me, "Get LowIndex", "LowIndex not avaible"
UsedLowerIndex = pLowIndex
End Property

'---------------------------------------------------------------------------------------
' Procedure : Compact
' DateTime  : 03/12/2001 13.44
' Author    : Administrator
' Purpose   : Resizes the array to remove empty values at the top and the bottom
'               of The UsedUpperIndex and UsedLowerIndex
'---------------------------------------------------------------------------------------
Public Sub Compact()
Attribute Compact.VB_Description = "Resizes the array to remove empty values at the top and the bottom of The UsedUpperIndex and UsedLowerIndex"
Attribute Compact.VB_HelpID = 2249
'##BLOCK_DESCRIPTION Resizes the array to remove empty values at the top and the bottom _
of The UsedUpperIndex and UsedLowerIndex
Ensure pInitialized, Me, "Compact", "Vector not initilized"
Ensure pUpLowIndexOn, Me, "Compact", "UpIndex and LowIndex not avaible: cannot compact vector"
Dim TmpVect As Variant, i As Long
ReDim TmpVect(pLowIndex To pUpIndex)

For i = pLowIndex To pUpIndex
    If pItemsType = vtObject Then
        Set TmpVect(i) = v(i)
    Else
        TmpVect(i) = v(i)
    End If
Next i
v = TmpVect
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Sortable
' DateTime  : 17/12/2001 12.16
' Description : True if items type allow sorting
'---------------------------------------------------------------------------------------
Public Property Get Sortable() As Boolean
Attribute Sortable.VB_Description = "True if items type allow sorting"
Attribute Sortable.VB_HelpID = 2250
'##BLOCK_DESCRIPTION True if items type allow sorting
Sortable = vbLT.SequenceSeek(pItemsType, vbTextCompare, _
    vtInteger, vtLong, vtSingle, vtDouble, vtCurrency, _
    vtDate, vtString, vtBoolean, vtVariant, vtDecimal, vtByte) >= 0
End Property

'---------------------------------------------------------------------------------------
' Procedure : Append
' DateTime  : 19/12/2001 18.09
' Description : appends a vector the existing one
'---------------------------------------------------------------------------------------
Public Sub Append(ByVal AppendVector As Vector)
Attribute Append.VB_Description = "Appends a vector to the existing one, starting from upper index."
Attribute Append.VB_HelpID = 2251
'##BLOCK_DESCRIPTION Appends a vector to the existing one, starting from upper index.
Ensure Not AppendVector Is Nothing, Me, "Append", "Null vector object: cannot append"
Dim i As Long, Base As Long

With AppendVector
    If .UsedLength > 0 Then
        Base = .UsedLowerIndex
        For i = 0 To .UsedLength - 1
            Me.AddItem .Item(i + Base)
        Next i
    End If
End With
End Sub
