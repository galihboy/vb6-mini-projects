VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayCnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2252
Attribute VB_Description = "Exposes many methods to convert an array from one type to another"
'---------------------------------------------------------------------------------------
' Module    : ArrayCnv
' DateTime  : 03/12/2001 13.19
' Author    : Administrator
' Purpose   : Exposes many methods to convert an array from one type to another
'---------------------------------------------------------------------------------------
' ##MODULE_DESCRIPTION Exposes many methods to convert an array from one type to another
Option Explicit

Private vbLT As New LangTool

Public Function CLngArray(ArrayValue As Variant) As Long()
Attribute CLngArray.VB_Description = "Converts an array to a long array"
Attribute CLngArray.VB_HelpID = 2253
'##BLOCK_DESCRIPTION Converts an array to a long array
Dim vRet() As Long, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CLngArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CLngArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CLng(ArrayValue(i))
Next i
On Error GoTo 0
CLngArray = vRet
Exit Function
Failure:
End Function

Public Function CIntArray(ArrayValue As Variant) As Integer()
Attribute CIntArray.VB_Description = "Converts an array to a integer array"
Attribute CIntArray.VB_HelpID = 2254
'##BLOCK_DESCRIPTION Converts an array to a integer array
Dim vRet() As Integer, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CByteArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CByteArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CInt(ArrayValue(i))
Next i
On Error GoTo 0
CIntArray = vRet
Exit Function
Failure:
End Function

Public Function CByteArray(ArrayValue As Variant) As Byte()
Attribute CByteArray.VB_Description = "Converts an array to a byte array"
Attribute CByteArray.VB_HelpID = 2255
'##BLOCK_DESCRIPTION Converts an array to a byte array
Dim vRet() As Byte, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CByteArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CByteArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CByte(ArrayValue(i))
Next i
On Error GoTo 0
CByteArray = vRet
Exit Function
Failure:
End Function

Public Function CDecArray(ArrayValue As Variant) As Variant()
Attribute CDecArray.VB_Description = "Converts an array to a decimal array"
Attribute CDecArray.VB_HelpID = 2256
'##BLOCK_DESCRIPTION Converts an array to a decimal array
Dim vRet() As Variant, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CDecArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CDecArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CDec(ArrayValue(i))
Next i
On Error GoTo 0
CDecArray = vRet
Exit Function
Failure:
End Function

Public Function CCurArray(ArrayValue As Variant) As Currency()
Attribute CCurArray.VB_Description = "Converts an array to a currency array"
Attribute CCurArray.VB_HelpID = 2257
'##BLOCK_DESCRIPTION Converts an array to a currency array
Dim vRet() As Currency, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CCurArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CCurArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CCur(ArrayValue(i))
Next i
On Error GoTo 0
CCurArray = vRet
Exit Function
Failure:
End Function

Public Function CSngArray(ArrayValue As Variant) As Single()
Attribute CSngArray.VB_Description = "Converts an array to a single array"
Attribute CSngArray.VB_HelpID = 2258
'##BLOCK_DESCRIPTION Converts an array to a single array
Dim vRet() As Single, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CSngArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CSngArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CSng(ArrayValue(i))
Next i
On Error GoTo 0
CSngArray = vRet
Exit Function
Failure:
End Function

Public Function CDblArray(ArrayValue As Variant) As Double()
Attribute CDblArray.VB_Description = "Converts an array to a double array"
Attribute CDblArray.VB_HelpID = 2259
'##BLOCK_DESCRIPTION Converts an array to a double array
Dim vRet() As Double, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CDblArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CDblArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CDbl(ArrayValue(i))
Next i
On Error GoTo 0
CDblArray = vRet
Exit Function
Failure:
End Function

Public Function CDateArray(ArrayValue As Variant) As Date()
Attribute CDateArray.VB_Description = "Converts an array to a date array"
Attribute CDateArray.VB_HelpID = 2260
'##BLOCK_DESCRIPTION Converts an array to a date array
Dim vRet() As Date, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CDateArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CDateArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CDate(ArrayValue(i))
Next i
On Error GoTo 0
CDateArray = vRet
Exit Function
Failure:
End Function

Public Function CBoolArray(ArrayValue As Variant) As Boolean()
Attribute CBoolArray.VB_Description = "Converts an array to a boolean array"
Attribute CBoolArray.VB_HelpID = 2261
'##BLOCK_DESCRIPTION Converts an array to a boolean array
Dim vRet() As Boolean, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CBoolArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CBoolArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CBool(ArrayValue(i))
Next i
On Error GoTo 0
CBoolArray = vRet
Exit Function
Failure:
End Function

Public Function CStrArray(ArrayValue As Variant) As String()
Attribute CStrArray.VB_Description = "Converts an array to a string array"
Attribute CStrArray.VB_HelpID = 2262
'##BLOCK_DESCRIPTION Converts an array to a string array
Dim vRet() As String, i As Long, l As Long, u As Long
Ensure IsArray(ArrayValue), Me, "CStrArray", "Parameter is not an array"
Ensure Not vbLT.IsEmptyArray(ArrayValue), Me, "CStrArray", "Array is empty"
l = LBound(ArrayValue)
u = UBound(ArrayValue)
ReDim vRet(l To u)
On Error GoTo Failure
For i = l To u
    vRet(i) = CStr(ArrayValue(i))
Next i
On Error GoTo 0
CStrArray = vRet
Exit Function
Failure:
End Function


