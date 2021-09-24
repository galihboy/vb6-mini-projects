VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2108
' --------------------------------------------------
'
' The CQUEUE class
'    Dim qu As New CQueue
'
'    ' enqueue to items to the queue
'    qu.Enqueue 1234
'    qu.Enqueue 5678
'    ' display number of elements in the queue
'    Debug.Print "Count = " & qu.Count
'    ' item at the element about to be read
'    Debug.Print "Item = " & qu.Item
'    ' print the elements in the queue while removing them
'    Do While qu.Count
'        Debug.Print "next item = " & qu.Dequeue
'    Loop
'--------------------------------------------------

' this is the collection that holds the values
Dim colValues As New Collection

' add a new value to the queue
Sub Enqueue(Value As Variant)
Attribute Enqueue.VB_HelpID = 2109
    colValues.Add Value
End Sub

' read the next value from the queue and remove it
' raises and error if the queue is empty
Function Dequeue() As Variant
Attribute Dequeue.VB_HelpID = 2110
    Ensure colValues.Count > 0, Me, "Dequeue", "Empty queue"
    If IsObject(colValues.Item(1)) Then
        Set Dequeue = colValues.Item(1)
    Else
        Dequeue = colValues.Item(1)
    End If
    colValues.Remove 1
End Function

' returns value for required index; if index is omitted
' returns first value to be read
' no removing actions are performed
Function Item(Optional ByVal Index As Long) As Variant
Attribute Item.VB_HelpID = 2111
    Ensure colValues.Count > 0, Me, "Item", "Empty queue"
    If Index > 0 And Index <= colValues.Count Then
        If IsObject(colValues.Item(Index)) Then
            Set Item = colValues.Item(Index)
        Else
            Item = colValues.Item(Index)
        End If
    Else
        If Index = 0 Then
            If IsObject(colValues.Item(1)) Then
                Set Item = colValues.Item(1)
            Else
                Item = colValues.Item(1)
            End If
        Else
            Throw Me, "Item", "Invalid index"
        End If
    End If
        
End Function

' Return the number of values in the queue
Function Count() As Long
Attribute Count.VB_HelpID = 2112
    Count = colValues.Count
End Function

' clear the queue
Sub Clear()
Attribute Clear.VB_HelpID = 2113
    Set colValues = New Collection
End Sub

