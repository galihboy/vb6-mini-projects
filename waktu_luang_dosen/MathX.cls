VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2195
Option Explicit

Public Function Log10(ByVal Number As Variant) As Double
Attribute Log10.VB_HelpID = 2196
Dim Num As Double
Const Lg10 As Double = 2.30258509299405
Num = CDbl(Number)
Log10 = Log(Num) / Lg10
End Function

Public Function LogX(ByVal Number As Variant, ByVal X As Variant) As Double
Attribute LogX.VB_HelpID = 2197
Dim Num As Double
Dim nBase As Double

Num = CDbl(Number)
nBase = CDbl(X)
LogX = Log(Num) / Log(nBase)

End Function

Public Function Ceil(ByVal Number As Variant) As Double
Attribute Ceil.VB_HelpID = 2198
Dim Num As Double
Num = CDbl(Number)
Ceil = Int(Num + 1)
End Function

Public Function Floor(ByVal Number As Variant) As Double
Attribute Floor.VB_HelpID = 2199
Dim Num As Double
Num = CDbl(Number)
Floor = Int(Num)
End Function

Public Function AbsCeil(ByVal Number As Variant, Optional ByVal KeepSign As Boolean) As Double
Attribute AbsCeil.VB_HelpID = 2200
Dim Num As Double, Sign As Integer
Num = CDbl(Number)
Sign = Sgn(Num)
AbsCeil = IIf(KeepSign, Sign, 1) * (Int(Abs(Num) + 1))
End Function

Public Function AbsFloor(ByVal Number As Variant, Optional ByVal KeepSign As Boolean) As Double
Attribute AbsFloor.VB_HelpID = 2201
Dim Num As Double, Sign As Integer
Num = CDbl(Number)
Sign = Sgn(Num)
AbsFloor = IIf(KeepSign, Sign, 1) * (Int(Abs(Num)))
End Function

' arc sine
' error if value is outside the range [-1,1]

Public Function ASin(Value As Double) As Double
Attribute ASin.VB_HelpID = 2202
    If Abs(Value) <> 1 Then
        ASin = Atn(Value / Sqr(1 - Value * Value))
    Else
        ASin = 1.5707963267949 * Sgn(Value)
    End If
End Function

' arc cosine
' error if NUMBER is outside the range [-1,1]

Public Function ACos(ByVal Number As Double) As Double
Attribute ACos.VB_HelpID = 2203
    If Abs(Number) <> 1 Then
        ACos = 1.5707963267949 - Atn(Number / Sqr(1 - Number * Number))
    ElseIf Number = -1 Then
        ACos = 3.14159265358979
    End If
    'elseif number=1 --> Acos=0 (implicit)
End Function

' arc cotangent
' error if NUMBER is zero

Public Function ACot(Value As Double) As Double
Attribute ACot.VB_HelpID = 2204
    ACot = Atn(1 / Value)
End Function

' arc secant
' error if value is inside the range [-1,1]

Public Function ASec(Value As Double) As Double
Attribute ASec.VB_HelpID = 2205
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1 Then
        ASec = 1.5707963267949 - Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ASec = 3.14159265358979 * Sgn(Value)
    End If
End Function

' arc cosecant
' error if value is inside the range [-1,1]

Public Function ACsc(Value As Double) As Double
Attribute ACsc.VB_HelpID = 2206
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ACsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ACsc = 1.5707963267949 * Sgn(Value)
    End If
End Function

' hyperbolic sine

Public Function SinH(Value As Double) As Double
Attribute SinH.VB_HelpID = 2207
    Dim temp As Double
    temp = Exp(Value)
    SinH = (temp - 1 / temp) / 2
End Function

' hyperbolic cosine

Public Function CosH(Value As Double) As Double
Attribute CosH.VB_HelpID = 2208
    Dim temp As Double
    temp = Exp(Value)
    CosH = (temp + 1 / temp) / 2
End Function

' hyperbolic tangent

Public Function TanH(Value As Double) As Double
Attribute TanH.VB_HelpID = 2209
    Dim temp As Double
    temp = Exp(Value)
    TanH = (temp - 1 / temp) / (temp + 1 / temp)
End Function

' hyperbolic cotangent

Public Function CotH(Value As Double) As Double
Attribute CotH.VB_HelpID = 2210
    Dim temp As Double
    temp = Exp(Value)
    CotH = (temp + 1 / temp) / (temp - 1 / temp)
End Function

' hyperbolic secant

Public Function SecH(Value As Double) As Double
Attribute SecH.VB_HelpID = 2211
    Dim temp As Double
    temp = Exp(Value)
    SecH = 2 / (temp + 1 / temp)
End Function

' hyperbolic cosecant

Public Function CscH(Value As Double) As Double
Attribute CscH.VB_HelpID = 2212
    Dim temp As Double
    temp = Exp(Value)
    CscH = 2 / (temp - 1 / temp)
End Function


' hyperbolic arc sine

Public Function ASinH(Value As Double) As Double
Attribute ASinH.VB_HelpID = 2213
    ASinH = Log(Value + Sqr(Value * Value + 1))
End Function

' hyperbolic arc cosine
' error if NUMBER is inside the range [-1,1]

Public Function ACosH(Value As Double) As Double
Attribute ACosH.VB_HelpID = 2214
    ACosH = Log(Value + Sqr(Value * Value - 1))
End Function

' hyperbolic arc tangent
' error if value is zero

Public Function ATanH(Value As Double) As Double
Attribute ATanH.VB_HelpID = 2215
    ATanH = Log((1 / Value + 1) / (1 / Value - 1)) / 2
End Function

' hyperbolic arc cotangent

Public Function ACotH(Value As Double) As Double
Attribute ACotH.VB_HelpID = 2216
    ACotH = Log((Value + 1) / (Value - 1)) / 2
End Function

' hyperbolic arc secant
' error if value is outside the range [-1,1]

Public Function ASecH(Value As Double) As Double
Attribute ASecH.VB_HelpID = 2217
    ASecH = Log((Sqr(1 - Value * Value) + 1) / Value)
End Function

' hyperbolic arc cosecant

Public Function ACscH(Value As Double) As Double
Attribute ACscH.VB_HelpID = 2218
    ACscH = Log((Sgn(Value) * Sqr(1 + Value * Value) + 1) / Value)
End Function

