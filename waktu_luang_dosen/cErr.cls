VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2337
Option Explicit

Public Description As String
Attribute Description.VB_VarHelpID = 2338
Public Number As Long
Attribute Number.VB_VarHelpID = 2339
Public Source As String
Attribute Source.VB_VarHelpID = 2340
Public HelpFile As String
Attribute HelpFile.VB_VarHelpID = 2341
Public HelpContext As String
Attribute HelpContext.VB_VarHelpID = 2342
Public LastDllError As Long
Attribute LastDllError.VB_VarHelpID = 2343

Public Sub Clear()
Attribute Clear.VB_HelpID = 2344
Description = Empty
Number = Empty
HelpFile = Empty
HelpContext = Empty
LastDllError = Empty
Source = Empty
End Sub

' predefinita
Public Property Get IsError() As Boolean
Attribute IsError.VB_HelpID = 2345
Attribute IsError.VB_UserMemId = 0
IsError = (Number <> 0) Or (LastDllError <> 0)
End Property

Public Sub SetByError(ByVal Error As ErrObject)
Attribute SetByError.VB_HelpID = 2346
With Error
    Number = .Number
    Description = .Description
    Source = .Source
    LastDllError = .LastDllError
    HelpFile = .HelpFile
    HelpContext = .HelpContext
End With
End Sub
