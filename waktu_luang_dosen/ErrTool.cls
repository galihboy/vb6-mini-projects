VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrTool"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2265
'---------------------------------------------------------------------------------------
' Module    : ErrTool
' DateTime  : 03/12/2001 13.04
' Author    : Administrator
' Purpose   : ErrTool implements some Error Managements Routines that you
'               can use to improve VB Error Handling capabilities
'   Need cErr class as a ErrObject Wrapper
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum ErrType
    ErrException = -999999999
    ErrReflection = -999999998
    ErrAbort = -999999997
End Enum

Private LastError As New cErr

'---------------------------------------------------------------------------------------
' Procedure : Ensure
' DateTime  : 03/12/2001 13.01
' Author    : Administrator
' Purpose   : Ensures that a conditions is satisfied, otherwise raises an expection
'               Ensure x<>0,Me,"SumAndDivide","Division by zero"
'               is the same as
'               if not(x<>0) then err.raise ErrException,"SumAndDivide","Division by zero"
'---------------------------------------------------------------------------------------

Public Sub Ensure(ByVal Condition As Boolean, ByRef SourceModule, _
    ByVal SourceMethodName As String, ByVal OtherwiseDescription As String, _
    Optional ByVal ErrorType As ErrType = ErrException)
Attribute Ensure.VB_HelpID = 2266
    
If Not Condition Then
    
    Err.Raise ErrorType, GetModuleName(SourceModule) & "." & SourceMethodName, OtherwiseDescription
    
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Throw
' DateTime  : 03/12/2001 13.07
' Author    : Administrator
' Purpose   : Raises an Exception Error;
'               Throw me,"MySub","Error description"
'               is the same as
'               err.raise ErrException,typename(me) & "." & "MySub","Error description"
'---------------------------------------------------------------------------------------
Public Sub Throw(ByRef SourceModule, ByVal SourceMethodName As String, _
    ByVal ExceptionDescription As String, Optional ByVal ErrorType As ErrType = ErrException)
Attribute Throw.VB_HelpID = 2267
    
    Err.Raise ErrorType, GetModuleName(SourceModule) & "." & SourceMethodName, ExceptionDescription
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ErrReflect
' DateTime  : 05/12/2001 11.04
' Author    : Administrator
' Purpose   : Reflect an error as the same standard error, not as an exception
' Usage : if IsStdError(Err) then ErrReflect(Err) ...
'         or see ErrCatch
'---------------------------------------------------------------------------------------
Public Sub ErrReflect(Error As Variant)
Attribute ErrReflect.VB_HelpID = 2268
Ensure IsObject(Error), Me, "ErrReflect", "ErrReflect needs as ErrObject or a cErr object"
Select Case LCase(TypeName(Error))
    Case "errobject"
        Err.Raise Error.Number, Error.Source, Error.Description, Error.HelpFile, Error.HelpContext
    Case "cerr"
        Err.Raise Error.Number, Error.Source, Error.Description, Error.HelpFile, Error.HelpContext
    Case Else
        Throw Me, "ErrReflect", "ErrReflect needs as ErrObject or a cErr object"
End Select
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ErrCatch
' DateTime  : 12/12/2001 12.35
' Author    : Administrator
' Purpose   : Captures an error
' Usage : used with errCatched
' EG: simple statement and short execution blocks
' On Error Resume Next
'     CheckElement Me ' could raise an error
'     ErrCatch Err    ' captures the error
' On Error GoTo 0     ' clear err object
' If ErrCatched Then  ' ErrCatched has 'IsError' as a predefined property
'     pMinOccurs = OldValue ' restores old value
'     ErrReflect ErrCatched ' reflect last error caught
' End If
' EG: using multiple ErrCatch as error checks
'On Error Resume Next
'    Dim i As Long, RC As RECT
'    Dim k As Long
'    For i = 1 To App.ThreadID
'        GetWindowRect Me.HWND, RC
'        ProcessRect RC
'        If ErrCatch(Err) Then
'            msgbox ...
'            ErrReflect ErrCatched
'        End If
'        If RC.Top > 100 Then
'            ...
'        End If
'        If ErrCatch(Err) Then
'            msgbox ErrCatched.Description, ...
'            Exit Sub
'        End If
'        ...
'    Next i
'On Error GoTo 0
'---------------------------------------------------------------------------------------
Public Function ErrCatch(ByVal Error As ErrObject) As Boolean
Attribute ErrCatch.VB_HelpID = 2269
Dim NewErr As New cErr
NewErr.SetByError Error
Set LastError = NewErr
ErrCatch = NewErr.IsError
End Function

'---------------------------------------------------------------------------------------
' Procedure : ErrCatched
' DateTime  : 12/12/2001 12.38
' Author    : Administrator
' Purpose   : Returns last error catched with ErrCatch
' Usage : see ErrCatch
'---------------------------------------------------------------------------------------
Public Function ErrCatched() As cErr
Attribute ErrCatched.VB_HelpID = 2270
Set ErrCatched = LastError
End Function


'---------------------------------------------------------------------------------------
' Procedure : Warn
' DateTime  : 05/12/2001 10.16
' Author    : Administrator
' Purpose   : Shows a message box with a warning message
' Usage : Warn me,"SubMain","Invalid handle"
'---------------------------------------------------------------------------------------

Public Sub Warn(ByRef SourceModule, ByVal SourceMethodName As String, _
    ByVal MessageDescription As String)
Attribute Warn.VB_HelpID = 2271

MsgBox "Warning raised from " & GetModuleName(SourceModule) & "." & SourceMethodName & vbCrLf & _
    MessageDescription, vbInformation, "Warning..."
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsStdError
' DateTime  : 05/12/2001 10.19
' Author    : Administrator
' Purpose   : Verifies if an error is a standard error
' Usage : if IsStdError(Err) then ...
'---------------------------------------------------------------------------------------

Public Function IsStdError(ByVal Error As ErrObject)
Attribute IsStdError.VB_HelpID = 2272
Dim ErrNumber
ErrNumber = Error.Number
IsStdError = Not (IsException(ErrNumber) Or IsReflection(ErrNumber) Or IsAbort(ErrNumber))
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsException
' DateTime  : 05/12/2001 10.19
' Author    : Administrator
' Purpose   : Verifies if an error is an Exception
' Usage : if IsException(Err) then ....
'---------------------------------------------------------------------------------------

Public Function IsException(ByVal Error As ErrObject) As Boolean
Attribute IsException.VB_HelpID = 2273
Dim ErrNumber
ErrNumber = Error.Number
IsException = (ErrNumber = ErrException)
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsReflection
' DateTime  : 05/12/2001 10.20
' Author    : Administrator
' Purpose   : Verifies if an error is a reflected error
' Usage : If IsReflection(Err) then ...
'---------------------------------------------------------------------------------------

Public Function IsReflection(ByVal Error As ErrObject) As Boolean
Attribute IsReflection.VB_HelpID = 2274
Dim ErrNumber
ErrNumber = Error.Number
IsReflection = (ErrNumber = ErrReflection)
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsAbort
' DateTime  : 05/12/2001 10.20
' Author    : Administrator
' Purpose   : Verifies if an error is an abort error
' Usage : if IsAbort(Err) then end
'---------------------------------------------------------------------------------------
Public Function IsAbort(ByVal Error As ErrObject) As Boolean
Attribute IsAbort.VB_HelpID = 2275
Dim ErrNumber
ErrNumber = Error.Number
IsAbort (ErrNumber = ErrAbort)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ErrTypeName
' DateTime  : 05/12/2001 10.21
' Author    : Administrator
' Purpose   : Returns the Name Type of the error
' Usage : MsgBox ErrTypeName(Err) &" error has occurred..." ...
'---------------------------------------------------------------------------------------
Public Function ErrTypeName(ByVal Error As ErrObject) As String
Attribute ErrTypeName.VB_HelpID = 2276
Dim ErrNumber
ErrNumber = Error.Number
Select Case ErrNumber
    Case ErrException:      ErrTypeName = "Exception"
    Case ErrReflection:     ErrTypeName = "Reflection"
    Case ErrAbort:          ErrTypeName = "Abort"
    Case Else:              ErrTypeName = "Standard"
End Select
End Function

