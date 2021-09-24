VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2307
Option Explicit

' Global Converter
' Conversion untility for exchanging data
' among differet region setted windows platforms

' (c) 2001 by Riccardo Iotti

Private Const gTRUE As String = "TRUE"
Private Const gFALSE As String = "FALSE"
Private Const gDECSEP As String = "."
Private Const gKSep As String = ","
Private Const gDateFormat As Long = jDateGGMM
Private Const gDateLongYear As Boolean = True
Private Const gDateSep As String = "/"
Private UsedDecSep As String

Private vbLT As New LangTool

Public Function GBool(ByVal Value As Variant) As Boolean
Attribute GBool.VB_HelpID = 2308
Dim b As String
b = Trim(UCase(Value))
Select Case b
    Case gTRUE
        GBool = True
    Case gFALSE
        GBool = False
    Case Else
        Err.Raise 13 ' type mismatch
End Select

End Function

Public Function GStr(ByVal Value As Variant) As String
Attribute GStr.VB_HelpID = 2309
Dim Number As String, jD As New jDate
Select Case VarType(Value)
    Case vbCurrency, vbDecimal, vbDouble, vbSingle:
        GStr = Replace(CStr(Value), UsedDecSep, gDECSEP)
    Case vbDate
        jD.DateSeparator = gDateSep
        jD.LongYear = gDateLongYear
        jD.DateFormat = gDateFormat
        jD.SetDate Year(Value), Month(Value), Day(Value)
        GStr = jD.GetDate
    Case vbBoolean
        GStr = IIf(Value, gTRUE, gFALSE)
    Case vbObject
        If TypeName(Value) = "jDate" Then
            Set jD = Value
            GStr = jD.GetDate
        End If
    Case Else
        GStr = CStr(Value)
End Select
    
    
End Function

Public Function GInt(ByVal Value As Variant) As Integer
Attribute GInt.VB_HelpID = 2310
GInt = CInt(Value)
End Function

Public Function GLng(ByVal Value As Variant) As Long
Attribute GLng.VB_HelpID = 2311
GLng = CLng(Value)
End Function

Public Function GSng(ByVal Value As Variant) As Single
Attribute GSng.VB_HelpID = 2312
Dim Num As cNumber
Select Case VarType(Value)
    Case vbString
        Set Num = vbLT.CheckNumber(Value, gDECSEP, gKSep)
        GSng = CSng(Num.DoubleValue)
    Case Else
        GSng = CSng(Value)
End Select
        
End Function

Public Function GDbl(ByVal Value As Variant) As Double
Attribute GDbl.VB_HelpID = 2313
Dim Num As cNumber
Select Case VarType(Value)
    Case vbString
        Set Num = vbLT.CheckNumber(Value, gDECSEP, gKSep)
        GDbl = CDbl(Num.DoubleValue)
    Case Else
        GDbl = CDbl(Value)
End Select
End Function

Public Function GCur(ByVal Value As Variant) As Currency
Attribute GCur.VB_HelpID = 2314
Dim Num As cNumber
Select Case VarType(Value)
    Case vbString
        Set Num = vbLT.CheckNumber(Value, gDECSEP, gKSep)
        GCur = CCur(Num.DoubleValue)
    Case Else
        GCur = CCur(Value)
End Select
End Function

Public Function GByte(ByVal Value As Variant) As Byte
Attribute GByte.VB_HelpID = 2315
GByte = CByte(Value)
End Function

Public Function GDate(ByVal Value As Variant) As jDate
Attribute GDate.VB_HelpID = 2316
Dim jD As New jDate
With jD
    .DateSeparator = gDateSep
    .DateFormat = gDateFormat
    .LongYear = gDateLongYear
End With
Select Case VarType(Value)
    Case vbDate
        jD.SetDate Year(Value), Month(Value), Day(Value)
        Set GDate = jD
    Case vbString
        jD.SetDateString Value
        Set GDate = jD
    Case vbLong
        jD.SetDateLong (Value)
        Set GDate = jD
    Case Else
        Err.Raise 13 ' TYPEMISMATCH
End Select
        
End Function

Public Function GDec(ByVal Value As String) As Variant
Attribute GDec.VB_HelpID = 2317
Dim Num As cNumber
Select Case VarType(Value)
    Case vbString
        Set Num = vbLT.CheckNumber(Value, gDECSEP, gKSep)
        GDec = CDec(Num.DoubleValue)
    Case Else
        GDec = CDec(Value)
End Select
End Function

Private Sub Class_Initialize()
Dim s As Single
s = 1.1
UsedDecSep = IIf(InStr(1, CStr(s), ".") > 0, ".", ",")
End Sub

Public Function VarTypeCnv(ByVal Value As Variant, ByVal TypeConst As VbVarType) As Variant
Attribute VarTypeCnv.VB_HelpID = 2318
Dim rVal As Variant
Select Case TypeConst
    Case vbByte: rVal = GByte(Value)
    Case vbDecimal: rVal = GDec(Value)
    Case vbCurrency: rVal = GCur(Value)
    Case vbInteger: rVal = GInt(Value)
    Case vbLong: rVal = GLng(Value)
    Case vbSingle: rVal = GSng(Value)
    Case vbDouble: rVal = GDbl(Value)
    Case vbDate: rVal = GDate(Value)
    Case vbString: rVal = GStr(Value)
    Case vbBoolean: rVal = GBool(Value)
    Case Else: Throw Me, "VarTypeCnv", "Unsupported TypeConst: " & TypeConst
End Select
VarTypeCnv = rVal
End Function



