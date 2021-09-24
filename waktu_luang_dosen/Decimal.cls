VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDecimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2298
'---------------------------------------------------------------------------------------
' Module    : cDecimal
' DateTime  : 05/12/2001 10.10
' Author    : Administrator
' Purpose   : Wrapper for Decimal Type
' Usage:
'   Dim Num as new cDecimal
'   Num = "1283000292380.98277388129"
'---------------------------------------------------------------------------------------
Option Explicit

Private pDec As Variant

' It is the default property
Public Property Get Value() As Variant
Attribute Value.VB_HelpID = 2299
Attribute Value.VB_UserMemId = 0
Value = pDec
End Property

Public Property Let Value(ByVal DecimalValue As Variant)
pDec = CDec(DecimalValue)
End Property
