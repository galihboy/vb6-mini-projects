VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2037
' --------------------------------------------------
'
'  The CSTACK class
'
' Usage:
'    Dim st As New CStack
'
'    ' push two values on the stack
'    st.Push 1234
'    st.Push 4567
'    ' display number of elements in the stack
'    Debug.Print "Count = " & st.Count
'    ' peek at the element on top of stack
'    Debug.Print "Peek = " & st.Peek
'    ' print the elements on the stack while popping them
'    Do While st.Count
'        Debug.Print "TOS item = " & st.Pop
'    Loop
'
'---------------------------------------------------


' this is the collection that holds the values
Dim colValues As New Collection

' add a new value onto the stack
Sub Push(Value As Variant)
Attribute Push.VB_HelpID = 2038
    colValues.Add Value
End Sub

' Pop a value off the stack - raise error if stack is emtpy
Function Pop() As Variant
Attribute Pop.VB_HelpID = 2039
    Pop = colValues.Item(colValues.Count)
    colValues.Remove colValues.Count
End Function

' Return the value on top of the stack, without popping it
' raise error if stack is empty
Function Peek() As Variant
Attribute Peek.VB_HelpID = 2040
    Peek = colValues.Item(colValues.Count)
End Function

' Return the number of values in the stack
Function Count() As Long
Attribute Count.VB_HelpID = 2041
    Count = colValues.Count
End Function

