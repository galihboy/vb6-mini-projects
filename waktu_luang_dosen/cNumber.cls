VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2175
'---------------------------------------------------------------------------------------
' Module    : Modulo di classe:cNumber
' DateTime  : 05/12/2001 10.13
' Author    : Administrator
' Purpose   : Describes a generic number, with format Informations
' Usage :
'   It's returned by CheckNumber function in LangTool - see CheckNumber
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum enNumberType
    enNaN
    enZero
    enPosIntegerNoThSep
    enPosIntegerThSep
    enPosDecimalNoThSep
    enPosDecimalThSep
    enNegIntegerNoThSep
    enNegIntegerThSep
    enNegDecimalNoThSep
    enNegDecimalThSep
End Enum

Public NumberType As enNumberType
Attribute NumberType.VB_VarHelpID = 2176
Public Negative As Boolean
Attribute Negative.VB_VarHelpID = 2177
Public IntegerPart As String
Attribute IntegerPart.VB_VarHelpID = 2178
Public DecimalPart As String
Attribute DecimalPart.VB_VarHelpID = 2179
Public StdFullNumber As String
Attribute StdFullNumber.VB_VarHelpID = 2180
Public ThousandSeparated As Boolean
Attribute ThousandSeparated.VB_VarHelpID = 2181
Public DoubleValue As Double
Attribute DoubleValue.VB_VarHelpID = 2182
Public DecimalValue As Variant
Attribute DecimalValue.VB_VarHelpID = 2183

Private Sub Class_Initialize()
NumberType = enNaN
End Sub
