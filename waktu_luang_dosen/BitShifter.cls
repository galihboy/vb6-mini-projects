VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitShifter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2290
Option Explicit

' Rotate a Long to the left the specified number of times
Public Function RotateLeft(ByVal Value As Long, ByVal times As Long) As Long
Attribute RotateLeft.VB_HelpID = 2291
    Dim i As Long, signBits As Long
    
    ' no need to rotate more times than required
    times = times Mod 32
    ' return the number if it's a multiple of 32
    If times = 0 Then RotateLeft = Value: Exit Function
    
    For i = 1 To times
        ' remember the 2 most significant bits
        signBits = Value And &HC0000000
        ' clear those bit and shift to the left by one position
        Value = (Value And &H3FFFFFFF) * 2
        ' if the number was negative, then add 1
        ' if bit 30 was set, then set the sign bit
        Value = Value Or ((signBits < 0) And 1) Or (CBool(signBits And _
            &H40000000) And &H80000000)
    Next
    RotateLeft = Value
End Function

' Raise 2 to a power
' the exponent must be in the range [0,31]
Public Function Power2(ByVal exponent As Long) As Long
Attribute Power2.VB_HelpID = 2292
    Static res(0 To 31) As Long
    Dim i As Long
    
    ' rule out errors
    If exponent < 0 Or exponent > 31 Then Err.Raise 5
    
    ' initialize the array at the first call
    If res(0) = 0 Then
        res(0) = 1
        For i = 1 To 30
            res(i) = res(i - 1) * 2
        Next
        ' this is a special case
        res(31) = &H80000000
    End If
    
    ' return the result
    Power2 = res(exponent)
        
End Function

' Rotate an Integer to the left the specified number of times
Public Function RotateLeftI(ByVal Value As Integer, ByVal times As Long) As Integer
Attribute RotateLeftI.VB_HelpID = 2293
    Dim i As Long, signBits As Integer
    
    ' no need to rotate more times than required
    times = times Mod 16
    ' return the number if it's a multiple of 16
    If times = 0 Then RotateLeftI = Value: Exit Function
    
    For i = 1 To times
        ' remember the 2 most significant bits
        signBits = Value And &HC000
        ' clear those bit and shift to the left by one position
        Value = (Value And &H3FFF) * 2
        ' if the number was negative, then add 1
        ' if bit 30 was set, then set the sign bit
        Value = Value Or ((signBits < 0) And 1) Or (CBool(signBits And &H4000) _
            And &H8000)
    Next
    RotateLeftI = Value
End Function

' Rotate a Long to the right the specified number of times
Public Function RotateRight(ByVal Value As Long, ByVal times As Long) As Long
Attribute RotateRight.VB_HelpID = 2294
    Dim i As Long, signBits As Long
    
    ' no need to rotate more times than required
    times = times Mod 32
    ' return the number if it's a multiple of 32
    If times = 0 Then RotateRight = Value: Exit Function
    
    For i = 1 To times
        ' remember the sign bit and bit 0
        signBits = Value And &H80000001
        ' clear those bits and shift to the right by one position
        Value = (Value And &H7FFFFFFE) \ 2
        ' if the number was negative, then re-insert the bit
        ' if bit 0 was set, then set the sign bit
        Value = Value Or ((signBits < 0) And &H40000000) Or (CBool(signBits And _
            1) And &H80000000)
    Next
    RotateRight = Value
End Function

' Rotate an Integer to the right the specified number of times
Public Function RotateRightI(ByVal Value As Integer, ByVal times As Long) As Integer
Attribute RotateRightI.VB_HelpID = 2295
    Dim i As Long, signBits As Integer
    
    ' no need to rotate more times than required
    times = times Mod 16
    ' return the number if it's a multiple of 16
    If times = 0 Then RotateRightI = Value: Exit Function
    
    For i = 1 To times
        ' remember the sign bit and bit 0
        signBits = Value And &H8001
        ' clear those bits and shift to the right by one position
        Value = (Value And &H7FFE) \ 2
        ' if the number was negative, then re-insert the bit
        ' if bit 0 was set, then set the sign bit
        Value = Value Or ((signBits < 0) And &H4000) Or (CBool(signBits And 1) _
            And &H8000)
    Next
    RotateRightI = Value
End Function

' Shift to the left of the specified number of times
Public Function ShiftLeft(ByVal Value As Long, ByVal times As Long) As Long
Attribute ShiftLeft.VB_HelpID = 2296
    ' we need to create a mask of 1's corresponding to the
    ' times in VALUE that will be retained in the result
    Dim mask As Long, signBit As Long
    
    ' return zero if too many times
    If times >= 32 Then Exit Function
    ' return the value if zero times
    If times = 0 Then ShiftLeft = Value: Exit Function
    
    ' this extracts the bit in Value that will become the sign bit
    mask = Power2(31 - times)
    ' this calculates the sign bit of the result
    signBit = CBool(Value And mask) And &H80000000
    ' this clears all the most significant times,
    ' that would be lost anyway, and also clears the sign bit
    Value = Value And (mask - 1)
    ' do the shift to the left, without risking an overflow
    ' and then add the sign bit
    ShiftLeft = (Value * Power2(times)) Or signBit
End Function

' Shift to the right of the specified number of times
Public Function ShiftRight(ByVal Value As Long, ByVal times As Long) As Long
Attribute ShiftRight.VB_HelpID = 2297
    ' we need to create a mask of 1's corresponding to the
    ' digits in VALUE that will be retained in the result
    Dim mask As Long, signBit As Long
    
    ' return zero if too many times
    If times >= 32 Then Exit Function
    ' return the value if zero times
    If times = 0 Then ShiftRight = Value: Exit Function
    
    ' evaluate the sign bit in advance
    signBit = (Value < 0) And Power2(31 - times)
    ' create a mask with 1's for the digits that will be preserved
    If times < 31 Then
        ' if times=31 then the mask is zero
        mask = Not (Power2(times) - 1)
    End If
    ' clear all the digits that will be discarded, and
    ' also clear the sign bit
    Value = (Value And &H7FFFFFFF) And mask
    ' do the shift, without any problem, and add the sign bit
    ShiftRight = (Value \ Power2(times)) Or signBit
End Function


