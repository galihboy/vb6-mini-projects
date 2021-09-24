VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LangTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2149
Option Explicit

Private vbGC As New GlobalConverter

Public Function IsArray(vArray As Variant) As Boolean
Attribute IsArray.VB_HelpID = 2150

IsArray = (VarType(vArray) And vbArray) = vbArray

End Function

Public Function IsEmptyArray(vArray As Variant) As Boolean
Attribute IsEmptyArray.VB_Description = "Returns true if the variant passed is an empty array"
Attribute IsEmptyArray.VB_HelpID = 2151
'##BLOCK_DESCRIPTION Returns true if the variant passed is an empty array
IsEmptyArray = (ArraySize(vArray) = 0)
End Function

Public Function ArraySize(vArray As Variant, Optional ByVal Dimension As Long = 1) As Long
Attribute ArraySize.VB_Description = "Returns the number of elements of the specified dimension of the array."
Attribute ArraySize.VB_HelpID = 2152
'##PARAMETER_DESCRIPTION Dimension Sets the dimension of the array of to retrieve the size. If omitted _
 1 is assumed.
'##BLOCK_DESCRIPTION Returns the number of elements of the specified dimension of the array.
On Error GoTo ArrayEmpty
If Not IsArray(vArray) Then Throw TypeName(Me), "ArraySize", "The Variant received is not an array"
ArraySize = 1 + UBound(vArray, Dimension) - LBound(vArray, Dimension)
ArrayEmpty:
End Function

Public Function MaxValue(vArray As Variant) As Variant
Attribute MaxValue.VB_Description = "Returns the maximum value in the array passed."
Attribute MaxValue.VB_HelpID = 2153
'##BLOCK_DESCRIPTION Returns the maximum value in the array passed.
'##PARAMETER_DESCRIPTION vArray An array - use Array() function.
Dim Max As Variant
Dim i As Long, vt As VariantTypeConstants

Ensure IsArray(vArray), Me, "MaxValue", "MaxValue requires an array as parameter"
Ensure ArraySize(vArray) > 0, Me, "MaxValue", "Empty array"

For i = 0 To UBound(vArray)
    If i = 0 Then
        Max = vArray(0)
        vt = VarType(vArray(0))
    Else
        If vArray(i) > Max Then Max = vArray(i)
        Ensure VarType(vArray(i)) = vt, Me, "MaxValue", "Array items must be of the same type"
    End If
Next i

MaxValue = Max

End Function

Public Function MinValue(vArray As Variant) As Variant
Attribute MinValue.VB_Description = "Returns the minimum value in the array passed."
Attribute MinValue.VB_HelpID = 2154
'##BLOCK_DESCRIPTION Returns the minimum value in the array passed.
'##PARAMETER_DESCRIPTION vArray An array - use Array() function.
Dim Min As Variant
Dim i As Long, vt As VariantTypeConstants

Ensure IsArray(vArray), Me, "MinValue", "MinValue requires an array as parameter"
Ensure ArraySize(vArray) > 0, Me, "MinValue", "Empty array"

For i = 0 To UBound(vArray)
    If i = 0 Then
        Min = vArray(0)
        vt = VarType(vArray(0))
    Else
        Ensure VarType(vArray(i)) = vt, Me, "MinValue", "Array items must be of the same type"
        If vArray(i) < Min Then Min = vArray(i)
    End If
Next i

MinValue = Min

End Function
Public Function ArrayToPath(ParamArray PathTokens()) As String
Attribute ArrayToPath.VB_Description = "Converts the passed paramArray in a path string, using backslash as a path separator."
Attribute ArrayToPath.VB_HelpID = 2155
'##BLOCK_DESCRIPTION Converts the passed paramArray in a path string, using backslash as _
 a path separator.
Dim i As Long
Const Slash = "\"

If UBound(PathTokens) >= 0 Then
    ArrayToPath = PathTokens(0)
    For i = 1 To UBound(PathTokens)
        ArrayToPath = AppendToPath(ArrayToPath, PathTokens(i), Slash)
    Next i
End If
End Function

Public Function SequenceSeek(ByVal SeekValue As Variant, ByVal CmpMethod As VbCompareMethod, ParamArray Sequence()) As Long
Attribute SequenceSeek.VB_Description = "Returns 0 if seekValue is not found, otherwise the forst position found of seekValue in Sequence"
Attribute SequenceSeek.VB_HelpID = 2156
'##BLOCK_DESCRIPTION Returns 0 if seekValue is not found, otherwise the forst position found of _
seekValue in Sequence
Dim i As Long, k As Long
Dim f As Long
k = UBound(Sequence)
f = -1
Ensure k > 0, Me, "SequenceSeek", "Empty sequence"
For i = 0 To k
    If StrComp(CStr(SeekValue), CStr(Sequence(i)), CmpMethod) = 0 Then
        f = i
        Exit For
    End If
Next i
SequenceSeek = f
End Function

Public Function ArraySeek(ByVal SeekValue As Variant, ArrayValue As Variant, Optional ByVal CmpMethod As VbCompareMethod = vbTextCompare) As Long
Attribute ArraySeek.VB_HelpID = 2157
' returns LBound-1 if seekValue is not found, otherwise the first position found of
' seekValue in Sequence
Dim i As Long, uB As Long
Dim f As Long, lB As Long
Ensure IsArray(ArrayValue), Me, "ArraySeek", "Not a valid array"
Ensure Not IsEmptyArray(ArrayValue), Me, "ArraySeek", "Empty Array"

uB = UBound(ArrayValue)
lB = LBound(ArrayValue)

f = lB - 1

For i = lB To uB
    If StrComp(CStr(SeekValue), CStr(ArrayValue(i)), CmpMethod) = 0 Then
        f = i
        Exit For
    End If
Next i
ArraySeek = f
End Function
Public Function AppendToPath(ByVal path, ByVal PathToAdd, Optional Slash As String = "\") As String
Attribute AppendToPath.VB_HelpID = 2158
' aggiunge percorsi o file a percorsi
' ritorna sempre un path senza slash finale
Dim p As String, a As String

a = Trim(PathToAdd)
p = Trim(path)

If Len(a) > 0 Then
    If Left(a, 1) = Slash Then
        a = Mid(a, 2)
    End If
Else
    AppendToPath = p
    Exit Function
End If

If Len(p) > 0 Then
    If Right(p, 1) <> Slash Then
        p = p & Slash
    End If
Else
    AppendToPath = a
    Exit Function
End If

AppendToPath = p & a
If Len(AppendToPath) > 0 Then
    If Right(AppendToPath, 1) = Slash Then
        AppendToPath = Mid(AppendToPath, 1, Len(AppendToPath) - 1)
    End If
End If

End Function

Public Function Between(ByVal Value, ByVal Smaller As Variant, ByVal Greater As Variant, Optional ByVal Inclusive As Boolean = True) As Boolean
Attribute Between.VB_HelpID = 2159
If Inclusive Then
    Between = (Smaller <= Value) And (Value <= Greater)
Else
    Between = (Smaller < Value) And (Value < Greater)
End If
End Function

Public Function StrOccursCount(ByRef StrWhere As String, _
    ByVal StrWhat As String, Optional ByVal CmpMethod As VbCompareMethod = vbTextCompare) As Long
Attribute StrOccursCount.VB_HelpID = 2160
' ritorna il numero di occorrenze trovate di StrWhat in StrWhere
Dim i As Long
Dim Pos As String
Dim L As Long

L = Len(StrWhat)

Ensure L > 0, Me, "StrOccursCount", "Unable to seek emtpty strings"

If StrWhere > "" Then
    Pos = 1
    Do
        Pos = InStr(Pos, StrWhere, StrWhat, CmpMethod)
        If Pos > 0 Then
            i = i + 1
            Pos = Pos + L
        End If
    Loop Until Pos = 0
End If

StrOccursCount = i
End Function

Public Function IsInteger(ByVal Value As Variant) As Boolean
Attribute IsInteger.VB_HelpID = 2161
    Dim s As String
    Dim i As Long, k As Long
    Dim IsNegative As Boolean
    Dim Skip As Boolean
    
    s = Trim$(CStr(Value))

    For i = 1 To Len(s)
        k = Asc(Mid(s, i, 1))
        Skip = False
        If i = 1 Then
            If (k = 45) Then
                IsNegative = True
                Skip = True
            End If
            If (k = 43) Then
                IsNegative = False
                Skip = True
            End If
        End If
        If Not Skip Then
            If Not (k >= 48 And k <= 57) Then Exit Function
        End If
    Next i
    
    IsInteger = True

End Function

Public Function QuoteSplit(ByVal SplitString As String, _
    Optional ByVal Delimiter As String = ",") As String()
Attribute QuoteSplit.VB_HelpID = 2162
' splitta una stringa in tokens separati dal delimitatore specificato
' è sensibile alla presenza di stringhe nella SpliString, per cui i delimitatori
' presenti dentro a "..." non vengono considerati validi
Dim SLen As Long, nTokens As Long
Dim MaskStr() As Byte
Dim i As Long, k As Long, j As Long, z As Long
Dim retval() As String

SLen = Len(SplitString)
If SLen > 0 Then
    ReDim MaskStr(1 To SLen)
    ' RICERCA DOPPI APICI
    i = 0
    k = 0
    Do
        i = InStr(i + 1, SplitString, Chr(34))
        If k = 0 Then
            k = i
        Else
            If i = 0 Then ' fino alla fine della stringa
                z = SLen
            Else
                z = i
            End If
            For j = k To z
                MaskStr(j) = 1
            Next j
            k = 0
        End If
    Loop Until i = 0
    
    ' splittaggio vero e proprio
    ReDim retval(0 To SLen - 1) ' dimesioni massime possibili
    nTokens = 0
    i = 0
    k = 1
    Do
        i = InStr(i + 1, SplitString, Delimiter)
        If i > 0 Then
            If MaskStr(i) = 0 Then ' delimitatore valido
                retval(nTokens) = Mid$(SplitString, k, i - k)
                nTokens = nTokens + 1
                k = i + 1
            End If
        Else
            If k > 0 Then ' fino alla fine - ultimo token
                retval(nTokens) = Mid$(SplitString, k)
                nTokens = nTokens + 1
                k = 0
            End If
        End If
    Loop Until i = 0
    ReDim Preserve retval(0 To nTokens - 1)
    
    QuoteSplit = retval
End If

End Function

Public Function SeekCollectionItem(ByVal Value As Variant, Coll As Collection) As Boolean
Attribute SeekCollectionItem.VB_HelpID = 2163
' returns true if the value searched is present in the collection
Dim v As Variant

For Each v In Coll
    If v = Value Then
        SeekCollectionItem = True
        Exit Function
    End If
Next
End Function

Public Sub ResetVar(ByRef Var)
Attribute ResetVar.VB_HelpID = 2164
' resets a variable with the appropriate value
' depending of its type

If IsObject(Var) Then
    Set Var = Nothing
Else
    If IsArray(Var) Then
        Erase Var
    Else
        Var = Empty
    End If
End If

End Sub

Public Function VarTypeCnv(ByVal Value As Variant, ByVal TypeConst As VbVarType) As Variant
Attribute VarTypeCnv.VB_HelpID = 2165
Dim rVal As Variant
Select Case TypeConst
    Case vbByte: rVal = CByte(Value)
    Case vbDecimal: rVal = CDec(Value)
    Case vbCurrency: rVal = CCur(Value)
    Case vbInteger: rVal = CInt(Value)
    Case vbLong: rVal = CLng(Value)
    Case vbSingle: rVal = CSng(Value)
    Case vbDouble: rVal = CDbl(Value)
    Case vbDate: rVal = CDate(Value)
    Case vbString: rVal = CStr(Value)
    Case vbBoolean: rVal = CBool(Value)
    Case vbObject:
        Ensure IsObject(Value), Me, "VarTypeCnv", "Cannot convert to an object"
        Set rVal = Value
    'Case vbDataObject:
    '    Ensure VarType(Value) = vbDataObject, Me, "VarTypeCnv", "Cannot convert to an object"
    '    rVal = Value
    Case vbVariant:
        rVal = CVar(Value)
    Case vbUserDefinedType:
        Ensure VarType(Value) = vbUserDefinedType, Me, "VarTypeCnv", "Cannot convert to an object"
        rVal = Value
    Case vbNull: rVal = Null
    Case vbEmpty: rVal = Empty
    Case vbError: rVal = CVErr(Value)
    Case vbArray: rVal = Value
    Case Else: Throw Me, "VarTypeCnv", "Unsupported TypeConst: " & TypeConst
End Select
If IsObject(rVal) Then
    Set VarTypeCnv = rVal
Else
    VarTypeCnv = rVal
End If
End Function

Public Function VarTypeConst(ByVal TypeName As String) As VbVarType
Attribute VarTypeConst.VB_HelpID = 2166
Dim rVal As VbVarType
Select Case LCase(Trim(TypeName))
    Case "byte": rVal = vbByte
    Case "decimal": rVal = vbDecimal
    Case "currency": rVal = vbCurrency
    Case "integer": rVal = vbInteger
    Case "long": rVal = vbLong
    Case "single": rVal = vbSingle
    Case "double": rVal = vbDouble
    Case "date": rVal = vbDate
    Case "string": rVal = vbString
    Case "boolean": rVal = vbBoolean
    Case "dataobject": rVal = vbDataObject
    Case "object": rVal = vbObject
    Case "variant": rVal = vbVariant
    Case "userdefinedtype", "udt": rVal = vbUserDefinedType
    Case "null": rVal = vbNull
    Case "error": rVal = vbError
    Case "empty": rVal = vbEmpty
    Case "array": rVal = vbArray
    Case Else: Throw Me, "VarTypeConst", "Unsupported TypeName " & TypeName
End Select
VarTypeConst = rVal

End Function

Public Function VarConstTypeName(ByVal cVarType As VbVarType) As String
Attribute VarConstTypeName.VB_HelpID = 2167
Dim rVal As String, rArr As String
If cVarType >= vbArray Then
    If cVarType = vbArray Then
        VarConstTypeName = "Array"
        Exit Function
    Else
        rArr = "Array of "
    End If
End If
Select Case cVarType
    Case vbBoolean: rVal = "Boolean"
    Case vbByte: rVal = "Byte"
    Case vbCurrency: rVal = "Currency"
    Case vbDate: rVal = "Date"
    Case vbArray: rVal = "Array"
    Case vbDataObject: rVal = "DataObject"
    Case vbDecimal: rVal = "Decimal"
    Case vbDouble: rVal = "Double"
    Case vbEmpty: rVal = "Empty"
    Case vbError:  rVal = "Error"
    Case vbInteger: rVal = "Integer"
    Case vbLong: rVal = "Long"
    Case vbNull: rVal = "Null"
    Case vbObject: rVal = "Object"
    Case vbSingle: rVal = "Single"
    Case vbString: rVal = "String"
    Case vbUserDefinedType: rVal = "UserDefinedType"
    Case vbVariant: rVal = "Variant"
    Case Else: Throw Me, "VarConstTypeName", "Unsupported TypeConst : " & cVarType
End Select
VarConstTypeName = rArr & rVal

End Function


Public Function FormatLen(ByVal Value As String, ByVal CharToAppend As String, _
    ByVal FinalLen As Long, Optional ByVal AppendOnLeft As Boolean = True) As String
Attribute FormatLen.VB_HelpID = 2168
' formats value adding 'chartoappend' char, on the left or on the right
' of value. Final value len is 'FinalLen'
' EX: formatLen("321","0",8)==>"00000321"; Formaten("ABC","#",5,false)==>"ABC##"
Dim vlen As Long, i As Long
Ensure Len(CharToAppend) = 1, Me, "FormatLen", "Invalid parameter: 'CharToAppend'"
vlen = Len(Value)
Ensure vlen <= FinalLen, Me, "FormatLen", "Final length must be greater ar equal to value length"
i = FinalLen - vlen
If AppendOnLeft Then
    FormatLen = String$(i, CharToAppend) & Value
Else
    FormatLen = Value & String$(i, CharToAppend)
End If
End Function

' Bubble Sort an array of any type
' BubbleSort is especially convenient with small arrays (1,000
' items or fewer) or with arrays that are already almost sorted
'
' NUMELS is the index of the last item to be sorted, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub BubbleSortS(arr() As Single, Optional numEls As Variant,
'  '     Optional descending As Boolean)
'   Dim value As Single

Public Sub BubbleSort(VariantArray As Variant, Optional LastIndex As Variant, _
    Optional Descending As Boolean)
Attribute BubbleSort.VB_HelpID = 2169

    Dim Value As Variant
    Dim Index As Long
    Dim firstItem As Long
    Dim indexLimit As Long, lastSwap As Long

    ' account for optional arguments
    If IsMissing(LastIndex) Then LastIndex = UBound(VariantArray)
    firstItem = LBound(VariantArray)
    lastSwap = LastIndex

    Do
        indexLimit = lastSwap - 1
        lastSwap = 0
        For Index = firstItem To indexLimit
            Value = VariantArray(Index)
            If (Value > VariantArray(Index + 1)) Xor Descending Then
                ' if the items are not in order, swap them
                VariantArray(Index) = VariantArray(Index + 1)
                VariantArray(Index + 1) = Value
                lastSwap = Index
            End If
        Next
    Loop While lastSwap
End Sub

' QuickSort an array of any type
' QuickSort is especially convenient with large arrays (>1,000
' items) that contains items in random order. Its performance
' quickly degrades if the array is already almost sorted. (There are
' variations of the QuickSort algorithm that work good with
' nearly-sorted arrays, though, but this routine doesn't use them.)
'
' LastIndex is the index of the last item to be sorted, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub QuickSortS(VariantArray() As Single, Optional LastIndex As Variant,
'  '     Optional Ascending As Boolean)
'   Dim value As Single, temp As Single

Public Sub QuickSort(VariantArray As Variant, Optional LastIndex As Variant, _
    Optional Ascending As Boolean = True)
Attribute QuickSort.VB_HelpID = 2170

    Dim Value As Variant, temp As Variant
    Dim sp As Integer
    Dim leftStk(32) As Long, rightStk(32) As Long
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
            If Ascending Then
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
' Procedure : CheckNumber
' DateTime  : 06/12/2001 15.07
' Author    : Administrator
' Purpose   : Analize a number
' Usage :
'   Dim Num as cNumber
'   set Num = CheckNumber("-401371.9601")
'
'   See cNumber Class for number description
'   Set specifyPositive if you need a plus sign in the Standard Number
'---------------------------------------------------------------------------------------
Public Function CheckNumber(ByVal Number As String, _
    Optional ByVal DecimalSep As String = ".", _
    Optional ByVal ThousandSep As String = ",", _
    Optional ByVal SpecifyPositive As Boolean = False) As cNumber
Attribute CheckNumber.VB_HelpID = 2171
Dim FractPart As String
Dim IntPart As String
Dim FractPos As Long
Dim Negative As Boolean, ThSepCount As Long
Const Signs As String = "+-"
Const Digits As String = "0123456789"
Const IntGroupSize As Integer = 3
Dim ch As String * 1
Dim i As Long, k As Long, retval As enNumberType
Dim Num As New cNumber

retval = enNaN

On Error GoTo Finish
    
    Number = Trim(Number)
    
    Ensure Len(Number) > 0, Me, "RetVal", "Null number"
    
    Ensure SequenceSeek(DecimalSep, vbTextCompare, ".", ",") >= 0, Me, "RetVal", "Invalid decimal separator :" & DecimalSep
    
    Ensure SequenceSeek(ThousandSep, vbTextCompare, ".", ",") >= 0, Me, "RetVal", "Invalid Thousand separator :" & ThousandSep
    
    Ensure (ThousandSep <> DecimalSep) And (ThousandSep > "" And DecimalSep > ""), Me, "RetVal", "DecimalSeparator must be different form Thousand Separator"
    
    ' determinazione segno
    ch = Left(Number, 1)
    If InStr(1, Signs, ch) > 0 Then ' sign before
        Negative = (ch = "-")
        Number = Trim(Mid(Number, 2))
    Else ' sign after
        ch = Right(Number, 1)
        If InStr(1, Signs, ch) > 0 Then
            Negative = (ch = "-")
            Number = Trim(Left(Number, Len(Number) - 1))
        End If
    End If
    
    ' suddivisione parte intera e parte decimale
    
    FractPos = InStrRev(Number, DecimalSep)
    If FractPos = 0 Then
        FractPart = ""
        IntPart = Number
    Else
        FractPart = Mid(Number, FractPos + 1)
        IntPart = Left(Number, FractPos - 1)
    End If
    
    If (FractPart = "" And IntPart = "") Then
        retval = enNaN
        GoTo Finish
    End If
    
    ' VERIFICA ZERO
    If FractPart = "" Then
        If IsCharSequence(IntPart, "0") Then
            retval = enZero
            GoTo Finish
        End If
    Else
        If IsCharSequence(IntPart, "0") And IsCharSequence(FractPart, "0") Then
            retval = enZero
            GoTo Finish
        End If
    End If
    
    ' verifica parte decimale
    If FractPart > "" Then
        If Not IsCharSequence(FractPart, Digits) Then
            retval = enNaN
            GoTo Finish
        End If
    End If
    ' verifica parte intera
    If IntPart > "" Then
        k = Len(IntPart)
        i = k
        While i > 0
            i = InStrRev(IntPart, ThousandSep, i)
            If i > 0 Then
                If ((k - i + 1) Mod (IntGroupSize + 1)) <> 0 Then
                    retval = enNaN
                    GoTo Finish
                Else
                    ThSepCount = ThSepCount + 1
                    i = i - 1
                End If
            End If
        Wend
        If ThSepCount <> Int(k / (IntGroupSize + 1)) Then
            If ThSepCount > 0 Then
                retval = enNaN
                GoTo Finish
            End If
        End If
        If Not IsCharSequence(IntPart, Digits & IIf(ThSepCount > 0, ThousandSep, "")) Then
            retval = enNaN
            GoTo Finish
        End If
    End If
    If FractPart = "" Then
        If IntPart > "" Then
            If ThSepCount > 0 Then
                If Negative Then
                    retval = enNegIntegerThSep
                Else
                    retval = enPosIntegerThSep
                End If
            Else
                If Negative Then
                    retval = enNegIntegerNoThSep
                Else
                    retval = enPosIntegerNoThSep
                End If
            End If
        Else
            Throw Me, "RetVal", "Unexcpected condition"
        End If
    Else
        If ThSepCount > 0 Then
            If Negative Then
                retval = enNegDecimalThSep
            Else
                retval = enPosDecimalThSep
            End If
        Else
            If Negative Then
                retval = enNegDecimalNoThSep
            Else
                retval = enPosDecimalNoThSep
            End If
        End If
    End If
    
    IntPart = Replace(IntPart, ThousandSep, "")
    
Finish:
    Err.Clear
    On Error Resume Next
    With Num
        .Negative = IIf(retval = enNaN, Empty, Negative)
        .NumberType = retval
        .StdFullNumber = IIf(retval = enNaN, Empty, _
            IIf(Negative, "-", IIf(SpecifyPositive, "+", "")) & IntPart & IIf(FractPart > "", ".", "") & FractPart)
        .IntegerPart = IIf(retval = enNaN, Empty, IntPart)
        .DecimalPart = IIf(retval = enNaN, Empty, FractPart)
        .ThousandSeparated = IIf(retval = enNaN, Empty, (ThSepCount > 0))
        If IntPart = "" Then IntPart = "0"
        If FractPart = "" Then FractPart = "0"
        .DoubleValue = IIf(Negative, -1, 1) * (CDbl(IntPart) + CDbl(FractPart) * 10 ^ (-Len(FractPart)))
        .DecimalValue = IIf(Negative, -1, 1) * (CDec(CDec(IntPart) + CDec(FractPart) * 10 ^ (-Len(FractPart))))
    End With
    
    Set CheckNumber = Num

End Function

Public Function IsCharSequence(ByVal StrValue As String, ByVal CharSequence As String) As Boolean
Attribute IsCharSequence.VB_HelpID = 2172
Dim i As Long, L As Long
' verifica se tutti i caratteri di StrValue appartengono alla sequenza CharSequence
i = 1
L = Len(StrValue)
If L = 0 Then Exit Function
Ensure CharSequence > "", Me, "IsCharSequence", "Empty CharSequence"
Do
    If InStr(1, CharSequence, Mid(StrValue, i, 1)) = 0 Then Exit Function
    i = i + 1
Loop Until (i > L)
IsCharSequence = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : CNum
' DateTime  : 17/12/2001 12.26
' Description : Creates a standard number, storing it in a string
'---------------------------------------------------------------------------------------
Public Function CNum(ByVal NumberValue As Variant, Optional ByVal AllowNull As Boolean = True) As String
Attribute CNum.VB_HelpID = 2173
Dim cN As cNumber
If NumberValue > "" Then
    Set cN = CheckNumber(NumberValue)
    If cN.NumberType <> enNaN Then CNum = cN.StdFullNumber
Else
    If AllowNull Then
        Exit Function
    Else
        Throw Me, "CNum", "Null Number"
    End If
End If
    
End Function

Public Function FormatNumber(ByVal NumberValue As Variant, Optional ByVal DecSeparator = ".", _
    Optional ByVal ThSeparator As String = ",", Optional ByVal TotDecimals As String = -1, _
    Optional ByVal SignPositive As Boolean = False)
Attribute FormatNumber.VB_HelpID = 2380
Dim cN As cNumber, s As String, k As Long, i As Long
Dim IntPart As String, DecPart As String, ch As String
Dim Sign As String

Set cN = CheckNumber(vbGC.GStr(NumberValue), , , SignPositive)
Ensure Len(DecSeparator) = 1, Me, "FormatNumber", "Invalid decimal separator :" & DecSeparator

IntPart = cN.IntegerPart
If cN.Negative Then
    Sign = "-"
Else
    If SignPositive Then Sign = "+"
End If

If ThSeparator > "" Then
    k = Len(IntPart)
    For i = 1 To k
        ch = Mid$(IntPart, k - i + 1, 1)
        s = s & ch
        If (i Mod 3 = 0) Then s = s & ThSeparator
    Next i
    IntPart = StrReverse(s)
End If
DecPart = cN.DecimalPart
If TotDecimals >= 0 Then
    DecPart = Left(DecPart, TotDecimals)
    If Len(DecPart) < TotDecimals Then DecPart = FormatLen(DecPart, "0", TotDecimals, False)
End If
FormatNumber = Sign & IntPart & IIf(DecPart > "", DecSeparator & DecPart, "")

End Function

Public Function CreateVector(ByVal VectorType As enVectorType, Elements As Variant) As Vector
Attribute CreateVector.VB_Description = "Create a vector object."
Attribute CreateVector.VB_HelpID = 2174
'##BLOCK_HELP_DESCRIPTION Create a vector object.
'##BLOCK_DESCRIPTION Rerieves a vector object containing the passed elements vector. You _
 must specify a vector type.
Dim v As New Vector
v.Init VectorType
v.Value = Elements
Set CreateVector = v
End Function


