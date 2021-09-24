VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2016
' -------------------------------------------------
' LINKED LIST class module
'
' This class implements a linked list structure, where you can store
' values (appending them after the last element or inserting them at
' given indexes), remove them, and visit them using recordset-like
' methods such as MoveFirst, MoveNext, and Move(n)
'
' NOTE: make Item the default member for this class
'       you can do so from inside the Tools | Procedure Attributes dialog box
'
' Usage:
'   Dim ll As New LinkedList
'   ' optionally size the inner array
'   ll.SetSize 10000
'   ' add new elements (can optionally take Before or After element)
'   ' returns the index where the element has been stored
'   newIndex = ll.Add newValue
'
'   ' loop over all elements
'   ll.MoveFirst
'   Do Until ll.EOL
'       ' display or process the current element
'       Print ll.Item
'       ' move to next element
'   Loop
'
'   See remarks for the FIND method for details about performing searches
'   on the elements in the linked list
'
'-------------------------------------------------

Option Explicit

' these are used by the Find method
Public Enum FindConditionConstants
    fccEqual
    fccNotEqual
    fccLess
    fccLessEqual
    fccGreater
    fccGreaterEqual
    fccInStr
End Enum

' initial size of the list
Const DEFAULT_INITIALSIZE = 100
' how many items are allocated each time
Const DEFAULT_CHUNKSIZE = 100

Private Type ListType
    Value As Variant
    prevNdx As Long    ' -1 if the element is free
    nextNdx As Long
End Type

' the actual list
Dim List() As ListType
' number of items in the list
Dim m_Count As Long

' index of first/last item in the list
Private FirstNdx As Long
Private LastNdx As Long
' index of first free item in the list
Private FreeNdx As Long

' chucnk size
Private m_ChunkSize As Long

' index to the current element
Private m_CurrIndex As Long
' current EOL status of the list
' (is valid only when m_CurrIndex = 0)
Private m_EOL As Boolean     ' if False, then BOL is true

' the index of the current element

Property Get CurrIndex() As Long
Attribute CurrIndex.VB_HelpID = 2017
    CurrIndex = m_CurrIndex
End Property

Property Let CurrIndex(ByVal NewValue As Long)
    m_CurrIndex = NewValue
End Property

' move to the first element

Sub MoveFirst()
Attribute MoveFirst.VB_HelpID = 2018
    m_CurrIndex = FirstNdx
End Sub

' move to the last element

Sub MoveLast()
Attribute MoveLast.VB_HelpID = 2019
    m_CurrIndex = LastNdx
End Sub

' move the the previous element

Sub MovePrevious()
Attribute MovePrevious.VB_HelpID = 2020
    ' this code works also when m_CurrIndex = 0
    m_CurrIndex = List(m_CurrIndex).prevNdx
    ' in case we move too much
    m_EOL = False
End Sub

' move to the next element

Sub MoveNext()
Attribute MoveNext.VB_HelpID = 2021
    ' this code works also when m_CurrIndex = 0
    m_CurrIndex = List(m_CurrIndex).nextNdx
    ' in case we move too much
    m_EOL = True
End Sub

' move to the Nth element

Sub Move(ByVal Index As Long)
Attribute Move.VB_HelpID = 2022
    CheckIndex Index
    ' if there were no error, then update the current Index
    m_CurrIndex = Index
End Sub

' Return true if we are at the beginning of list

Property Get BOL() As Boolean
Attribute BOL.VB_HelpID = 2023
    BOL = (m_Count = 0) Or (m_CurrIndex = 0 And m_EOL = False)
End Property

' Return true if we are at the end of list

Property Get EOL() As Boolean
Attribute EOL.VB_HelpID = 2024
    EOL = (m_Count = 0) Or (m_CurrIndex = 0 And m_EOL = True)
End Property

' An item of the list (read-write)
' if the argument is omitted it retrieves the current item

Property Get Item(Optional ByVal Index As Long) As Variant
Attribute Item.VB_HelpID = 2025
    If Index = 0 Then Index = m_CurrIndex
    ' check that the index point to a valid, non-free element
    CheckIndex Index
    ' two cases: the value is/isn't an object
    If IsObject(List(Index).Value) Then
        Set Item = List(Index).Value
    Else
        Item = List(Index).Value
    End If
End Property

Property Let Item(Optional ByVal Index As Long, NewValue As Variant)
    If Index = 0 Then Index = m_CurrIndex
    ' check that this is a valid, non-free item
    CheckIndex Index
    ' modify the value in the list
    List(Index).Value = NewValue
End Property

Property Set Item(Optional ByVal Index As Long, NewValue As Object)
    If Index = 0 Then Index = m_CurrIndex
    ' check that this is a valid, non-free item
    CheckIndex Index
    ' modify the value in the list
    Set List(Index).Value = NewValue
End Property

' return True if the list is empty

Property Get IsEmpty() As Boolean
Attribute IsEmpty.VB_HelpID = 2026
    IsEmpty = (m_Count = 0)
End Property

' the number of elements in the list

Property Get Count() As Long
Attribute Count.VB_HelpID = 2027
    Count = m_Count
End Property


' insert a new item before/after a given element
' if both arguments are omitted it is appended to the end of the list
' a negative value for Before means "at the beginning of the list"
' and works also when the list is empty
'
' returns the index of the new value

Function Add(Value As Variant, Optional ByVal Before As Long, _
    Optional ByVal After As Long) As Long
Attribute Add.VB_HelpID = 2028
    Dim ndx As Long
    Dim nextFreeNdx As Long
    
    ' enlarge the list if necessary
    If FreeNdx = 0 Then ExpandList m_ChunkSize
    ' use the first free slot
    ndx = FreeNdx
    nextFreeNdx = List(ndx).nextNdx
    
    ' a special value for Before
    If Before < 0 Then Before = FirstNdx
    
    If Before > 0 Then
        ' check that this is a valid, non-free item
        CheckIndex Before
        ' "Before" item becomes this item's next element
        List(ndx).nextNdx = Before
        ' "Before"'s previous element becomes this item's previous element
        List(ndx).prevNdx = List(Before).prevNdx
        ' "Before's" previous element should point to this item
        List(Before).prevNdx = ndx
        
        If Before = FirstNdx Then
            ' "Before" was the first item in the list
            FirstNdx = ndx
        Else
            ' else, another item's next element points to this item
            List(List(ndx).prevNdx).nextNdx = ndx
        End If
    
    ElseIf After > 0 Then
        ' check that this is a valid, non-free item
        CheckIndex After
        ' "After" item becomes this item's previous element
        List(ndx).prevNdx = After
        ' "After" item's next element becomes this item's next element
        List(ndx).nextNdx = List(After).nextNdx
        ' "After"'s next element should point to this item
        List(After).nextNdx = ndx
        
        If After = LastNdx Then
            ' "After" was the last item in the list
            LastNdx = ndx
        Else
            ' else, another item's previous element points to this item
            List(List(ndx).nextNdx).prevNdx = ndx
        End If
        
    Else
        ' append at the end of the list
        If LastNdx Then
            ' this item becomes the "next" item of the
            ' item that was at the end of the list
            List(LastNdx).nextNdx = ndx
            List(ndx).prevNdx = LastNdx
        Else
            ' this is the first and only item in the list
            FirstNdx = ndx
            ' signal that this item isn't free any longer
            List(ndx).prevNdx = 0
        End If
        ' in all cases this becomes the last item in the list
        LastNdx = ndx
        List(ndx).nextNdx = 0
    End If
    
    ' actually store the new value
    If IsObject(Value) Then
        Set List(ndx).Value = Value
    Else
        List(ndx).Value = Value
    End If
    ' remember we have a new item
    m_Count = m_Count + 1
    ' FreeNdx must point to the first free slot
    FreeNdx = nextFreeNdx
    
    ' return the index of the element just added
    Add = ndx
    
End Function

' remove an item given its index
' if the index is omitted it removes the current item
'
' if the item is also the current item, then its subsequent
' element becomes the current item (if it was the last element
' then the EOL condition becomes True)

Sub Remove(Optional ByVal Index As Long)
Attribute Remove.VB_HelpID = 2029
    If Index = 0 Then Index = m_CurrIndex
    ' check that this is a valid, non-free item
    CheckIndex Index
    
    If Index = FirstNdx Then
        ' the item being removed is the first of the list
        FirstNdx = List(Index).nextNdx
        List(FirstNdx).prevNdx = 0
    ElseIf Index = LastNdx Then
        ' the item being removed is the last of the list
        ' but the list contains at least another item
        LastNdx = List(Index).prevNdx
        List(LastNdx).nextNdx = 0
    Else
        ' the item is in the middle of the list
        List(List(Index).prevNdx).nextNdx = List(Index).nextNdx
        List(List(Index).nextNdx).prevNdx = List(Index).prevNdx
    End If

    ' clear this item's value
    List(Index).Value = Empty
    ' remember we have one element less
    m_Count = m_Count - 1
    
    ' if this was the current item, update m_CurrIndex
    If Index = m_CurrIndex Then
        m_CurrIndex = List(Index).nextNdx
        ' if it was the last element of the list
        If m_CurrIndex = 0 Then m_EOL = True
    End If
    ' put it at the beginning of the free list
    List(Index).nextNdx = FreeNdx
    ' mark it as free
    List(Index).prevNdx = -1
    FreeNdx = Index
End Sub

' remove all items
' this method also resets any SetSize setting

Sub RemoveAll()
Attribute RemoveAll.VB_HelpID = 2030
    ' it simply restarts from the very beginning
    Class_Initialize
End Sub

' search a value in the list
'
' STARTINDEX is the index of the element from where to
' start the search - use ZERO or omitted to start from
' the current element, use -1 to start from first/last element
' if DESCENDING is True then it does a reverse search
'
' returns the index of the found element, or zero if not found
' the element also becomes the current element

Function Find(Value As Variant, Optional Condition As FindConditionConstants = _
    fccEqual, Optional ByVal StartIndex As Long, Optional ByVal Descending As _
    Boolean) As Long
Attribute Find.VB_HelpID = 2031
    Dim isObj As Boolean
    
    ' provide reasonable defaults
    If StartIndex = 0 Then
        StartIndex = m_CurrIndex
    ElseIf StartIndex < 0 Then
        If Not Descending Then
            StartIndex = FirstNdx
        Else
            StartIndex = LastNdx
        End If
    Else
        ' check that this index is valid
        CheckIndex StartIndex
        ' start from the next or previous element
        If Not Descending Then
            StartIndex = List(StartIndex).nextNdx
        Else
            StartIndex = List(StartIndex).prevNdx
        End If
    End If
    
    ' evaluate this once and for all
    isObj = IsObject(Value)
    
    ' two loops, depending on value being an object or not
    Do While StartIndex
        If isObj Then
            ' do the comparison only if the element is also an object
            If IsObject(List(StartIndex).Value) Then
                If Value Is List(StartIndex).Value Then
                    ' exit if we're looking for equality
                    If Condition <> fccNotEqual Then Exit Do
                Else
                    ' exit if we're looking for inequality
                    If Condition = fccNotEqual Then Exit Do
                End If
            End If
        Else
            ' do the comparison only if the element isn't an object
            If Not IsObject(List(StartIndex).Value) Then
                Select Case Condition
                    Case fccNotEqual
                        If List(StartIndex).Value <> Value Then Exit Do
                    Case fccLess
                        If List(StartIndex).Value < Value Then Exit Do
                    Case fccLessEqual
                        If List(StartIndex).Value <= Value Then Exit Do
                    Case fccGreater
                        If List(StartIndex).Value > Value Then Exit Do
                    Case fccGreaterEqual
                        If List(StartIndex).Value >= Value Then Exit Do
                    Case fccInStr
                        If InStr(List(StartIndex).Value, Value) Then Exit Do
                    Case Else
                        ' equality is the default test
                        If List(StartIndex).Value = Value Then Exit Do
                End Select
            End If
        End If
        ' skip to the next or previous item
        If Not Descending Then
            StartIndex = List(StartIndex).nextNdx
        Else
            StartIndex = List(StartIndex).prevNdx
        End If
    Loop
        
    ' make the item the current item and return its index
    m_CurrIndex = StartIndex
    Find = StartIndex
        
End Function

' modify the list size and growth factor

Sub SetSize(ByVal numEls As Long, Optional ByVal ChunkSize As Long)
Attribute SetSize.VB_HelpID = 2032
    ' raise an error if invalid arguments
    If numEls <= 0 Or ChunkSize < 0 Then Err.Raise 5
    
    If numEls > UBound(List) Then
        ' expand the list
        ExpandList numEls - UBound(List)
    End If
End Sub

'--------------------------------------------
' Private procedures
'--------------------------------------------

Private Sub Class_Initialize()
    ' initialize the linked list of free slots
    ReDim List(0) As ListType
    ExpandList DEFAULT_INITIALSIZE
    m_ChunkSize = DEFAULT_CHUNKSIZE
End Sub

' check that there is at least one free slot

Private Sub ExpandList(ByVal numEls As Long)
    Dim ndx As Long, newFreeNdx As Long
    
    ' this will be the first free slot
    newFreeNdx = UBound(List) + 1
    ' expand the list
    ReDim Preserve List(UBound(List) + numEls) As ListType
    ' initialize the links for free items
    ' (note that the nextNdx value for the last element isn't correct
    For ndx = FreeNdx To UBound(List)
        List(ndx).nextNdx = ndx + 1
        List(ndx).prevNdx = -1
    Next
    ' append the existing chain of free slots
    ' (fixes the invalid value stored previously)
    List(UBound(List)).nextNdx = FreeNdx
    ' this is the new start of the free list
    FreeNdx = newFreeNdx
End Sub

' check whether an index is valid

Sub CheckIndex(ByVal Index As Long)
Attribute CheckIndex.VB_HelpID = 2033
    If Index <= 0 Or Index > UBound(List) Then Err.Raise 5
    If List(Index).prevNdx < 0 Then Err.Raise 5
End Sub

