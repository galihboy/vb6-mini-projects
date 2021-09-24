VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdNumberConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2219
Option Explicit

Private pKSep As String
Private pDecSep As String

Private vbLT As New LangTool
Private Initialized As Boolean
Private cN As cNumber

Public Sub Init(Optional ByVal ThousandSeparator As String = ",", Optional ByVal DecimalSeparator As String = ".")
Attribute Init.VB_HelpID = 2220

Ensure vbLT.SequenceSeek(ThousandSeparator, vbTextCompare, ",", ".") >= 0, Me, "Init", "Invalid Thousand Separator : " & ThousandSeparator
Ensure vbLT.SequenceSeek(DecimalSeparator, vbTextCompare, ",", ".") >= 0, Me, "Init", "Invalid Decimal Separator" & DecimalSeparator
Ensure ThousandSeparator <> DecimalSeparator, Me, "Init", "Thousand and decimal separators must differs"

pKSep = ThousandSeparator
pDecSep = DecimalSeparator
Initialized = True

End Sub

Public Function CnvNum2Str(ByVal Value As Variant) As String
Attribute CnvNum2Str.VB_HelpID = 2221
Ensure Initialized, Me, "CnvNum2Str", "Class not initialized"
Set cN = vbLT.CheckNumber(Value, pDecSep, pKSep)
Ensure cN.NumberType <> enNaN, Me, "CnvNum2Str", Value & " is not a number"
CnvNum2Str = cN.StdFullNumber
End Function

Public Function CnvNum2Dec(ByVal Value As Variant) As Variant
Attribute CnvNum2Dec.VB_HelpID = 2222
Ensure Initialized, Me, "CnvNum2Dec", "Class not initialized"
Set cN = vbLT.CheckNumber(Value, pDecSep, pKSep)
Ensure cN.NumberType <> enNaN, Me, "CnvNum2Dec", Value & " is not a number"
CnvNum2Dec = cN.DecimalValue
End Function

Public Function CnvNum2Dbl(ByVal Value As Variant) As Double
Attribute CnvNum2Dbl.VB_HelpID = 2223
Ensure Initialized, Me, "CnvNum2Dbl", "Class not initialized"
Set cN = vbLT.CheckNumber(Value, pDecSep, pKSep)
Ensure cN.NumberType <> enNaN, Me, "CnvNum2Dbl", Value & " is not a number"
CnvNum2Dbl = cN.DoubleValue
End Function

Public Function CnvNum2Sng(ByVal Value As Variant) As Single
Attribute CnvNum2Sng.VB_HelpID = 2224
Ensure Initialized, Me, "CnvNum2Sng", "Class not initialized"
Set cN = vbLT.CheckNumber(Value, pDecSep, pKSep)
Ensure cN.NumberType <> enNaN, Me, "CnvNum2Sng", Value & " is not a number"
CnvNum2Sng = CSng(cN.DoubleValue)
End Function

Public Function CnvNum2Int(ByVal Value As Variant) As Integer
Attribute CnvNum2Int.VB_HelpID = 2225
Ensure Initialized, Me, "CnvNum2Int", "Class not initialized"
Set cN = vbLT.CheckNumber(Value, pDecSep, pKSep)
Ensure cN.NumberType <> enNaN, Me, "CnvNum2Int", Value & " is not a number"
CnvNum2Int = CInt(cN.DecimalValue)
End Function

Public Function CnvNum2Lng(ByVal Value As Variant) As Long
Attribute CnvNum2Lng.VB_HelpID = 2226
Ensure Initialized, Me, "CnvNum2Lng", "Class not initialized"
Set cN = vbLT.CheckNumber(Value, pDecSep, pKSep)
Ensure cN.NumberType <> enNaN, Me, "CnvNum2Lng", Value & " is not a number"
CnvNum2Lng = CLng(cN.DecimalValue)
End Function

Public Function CnvNum2Cur(ByVal Value As Variant) As Currency
Attribute CnvNum2Cur.VB_HelpID = 2227
Ensure Initialized, Me, "CnvNum2Cur", "Class not initialized"
Set cN = vbLT.CheckNumber(Value, pDecSep, pKSep)
Ensure cN.NumberType <> enNaN, Me, "CnvNum2Cur", Value & " is not a number"
CnvNum2Cur = CCur(cN.DecimalValue)
End Function

