VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "langX"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2347
'---------------------------------------------------------------------------------------
' Module    : Modulo di classe[langX]
' DateTime  : 20/12/2001 00.17
' Purpose   : Extends globally VB language with new operators and functions
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : IncVar
' DateTime  : 19/12/2001 17.20
' Description : IncVar(x) is the same as the C operator ++x
'---------------------------------------------------------------------------------------
Public Function IncVar(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
Attribute IncVar.VB_HelpID = 2348
Value = Value + Amount
IncVar = Value
End Function

'---------------------------------------------------------------------------------------
' Procedure : DecVar
' DateTime  : 19/12/2001 17.21
' Description : DecVar(x) is the same as the C operator --x
'---------------------------------------------------------------------------------------
Public Function DecVar(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
Attribute DecVar.VB_HelpID = 2349
Value = Value - Amount
DecVar = Value
End Function

'---------------------------------------------------------------------------------------
' Procedure : VarInc
' DateTime  : 19/12/2001 17.21
' Description : VarInc(x) is the same as the C operator x++
'---------------------------------------------------------------------------------------
Public Function VarInc(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
Attribute VarInc.VB_HelpID = 2350
VarInc = Value
Value = Value + Amount
End Function

'---------------------------------------------------------------------------------------
' Procedure : VarDec
' DateTime  : 19/12/2001 17.22
' Description : VarDec(x) is the same as the C operator x--
'---------------------------------------------------------------------------------------
Public Function VarDec(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
Attribute VarDec.VB_HelpID = 2351
VarDec = Value
Value = Value - Amount
End Function

'---------------------------------------------------------------------------------------
' Procedure : LetVar
' DateTime  : 21/12/2001 18.11
' Description : assigns a value and returns it
'---------------------------------------------------------------------------------------
Public Function LetVar(ByRef Variable As Variant, ByVal Value As Variant) As Variant
Attribute LetVar.VB_HelpID = 2352
Variable = Value
LetVar = Variable
End Function

'---------------------------------------------------------------------------------------
' Procedure : VarLet
' DateTime  : 21/12/2001 18.11
' Description : returns the variable value and, after that, assigns new value
'---------------------------------------------------------------------------------------
Public Function VarLet(ByRef Variable As Variant, ByVal Value As Variant) As Variant
Attribute VarLet.VB_HelpID = 2353
VarLet = Variable
Variable = Value
End Function

Public Function Max(ParamArray Values()) As Variant
Attribute Max.VB_Description = "Returns the maximum value in the array passed."
Attribute Max.VB_HelpID = 2389
'##BLOCK_DESCRIPTION Returns the maximum value in the array passed.
Dim i As Long, vt As VariantTypeConstants

Ensure UBound(Values) > 0, Me, "Max", "Empty array"

For i = 0 To UBound(Values)
    If i = 0 Then
        Max = Values(0)
        vt = VarType(Values(0))
    Else
        If Values(i) > Max Then Max = Values(i)
        Ensure VarType(Values(i)) = vt, Me, "Max", "Array items must be of the same type"
    End If
Next i

End Function

Public Function Min(ParamArray Values()) As Variant
Attribute Min.VB_Description = "Returns the minimum value in the array passed."
Attribute Min.VB_HelpID = 2390
'##BLOCK_DESCRIPTION Returns the minimum value in the array passed.
Dim i As Long, vt As VariantTypeConstants

Ensure UBound(Values) > 0, Me, "Min", "Empty array"

For i = 0 To UBound(Values)
    If i = 0 Then
        Min = Values(0)
        vt = VarType(Values(0))
    Else
        Ensure VarType(Values(i)) = vt, Me, "Min", "Array items must be of the same type"
        If Values(i) < Min Then Min = Values(i)
    End If
Next i

End Function

Public Function Xpp(ByRef IntValue As Long) As Long
Attribute Xpp.VB_HelpID = 2391
Xpp = IntValue
IntValue = IntValue + 1
End Function

Public Function ppX(ByRef IntValue As Long) As Long
Attribute ppX.VB_HelpID = 2392
IntValue = IntValue + 1
ppX = IntValue
End Function

Public Function Xmm(ByRef IntValue As Long) As Long
Attribute Xmm.VB_HelpID = 2393
Xmm = IntValue
IntValue = IntValue - 1
End Function

Public Function mmX(ByRef IntValue As Long) As Long
Attribute mmX.VB_HelpID = 2394
IntValue = IntValue - 1
mmX = IntValue
End Function

