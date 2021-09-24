VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2263
Option Explicit

Public Function InstanceOf(ObjectVariable As Variant, ByVal ClassType As String) As Boolean
Attribute InstanceOf.VB_HelpID = 2264

Ensure IsObject(ObjectVariable), Me, "InstanceOf", "Not a object variable parameter"
Ensure Not ObjectVariable Is Nothing, Me, "InstanceOf", "Null object passed"

Dim vt As String
vt = Trim(UCase(TypeName(ObjectVariable)))
InstanceOf = (UCase(Trim(ClassType)) = vt)

End Function


' non funziona.....
'Public Function ImplInterface(ByVal InterfaceName As String, ByVal Obj As Object) As Boolean
'Dim TLI As New TLIApplication
'Dim tpInfo As TypeInfo, tpInterface As InterfaceInfo
'On Error GoTo Failure
'Set tpInfo = TLI.ClassInfoFromObject(Obj)
'For Each tpInterface In tpInfo.Interfaces
'    If StrComp(InterfaceName, tpInterface.Name, vbTextCompare) = 0 Then
'        ImplInterface = True
'        Exit Function
'    End If
'Next
'Exit Function
'Failure:
'    Err.Raise ErrException, Me, "ImplInterface failure: " & Err.Description & " at line " & Erl
'
'End Function

