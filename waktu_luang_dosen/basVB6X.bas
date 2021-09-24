Attribute VB_Name = "basVB6X"
Option Explicit

' VB6X - Estensione VB6

' gestione eccezioni
' abilitare il logging

Public Enum ERR_Type
    ERR_Exception = -999999999
    ERR_Reflection = -999999998
    ERR_Abort = -999999997
End Enum

Public Sub Ensure(ByVal Condition As Boolean, ByRef SourceModule, _
    ByVal SourceMethodName As String, ByVal OtherwiseDescription As String, _
    Optional ByVal ErrType As ERR_Type = ERR_Exception)
    
If Not Condition Then
    
    Err.Raise ErrType, GetModuleName(SourceModule) & "." & SourceMethodName, OtherwiseDescription
    
End If

End Sub

Public Sub Throw(ByRef SourceModule, ByVal SourceMethodName As String, _
    ByVal ExceptionDescription As String, Optional ByVal ErrType As ERR_Type = ERR_Exception)
    
    Err.Raise ErrType, GetModuleName(SourceModule) & "." & SourceMethodName, ExceptionDescription
    
End Sub

Public Sub Warn(ByRef SourceModule, ByVal SourceMethodName As String, _
    ByVal MessageDescription As String)

MsgBox "Warning raised from " & GetModuleName(SourceModule) & "." & SourceMethodName & vbCrLf & _
    MessageDescription, vbInformation, "Warning..."
    
End Sub


Public Function GetModuleName(ByRef SourceModule) As String

    Select Case VarType(SourceModule)
        Case vbObject:
        GetModuleName = TypeName(SourceModule)
    Case vbString:
        GetModuleName = SourceModule
    Case Else:
        GetModuleName = "<UndefinedModule>"
End Select

End Function

Public Sub Catch()
' cattura gli errori
End Sub

Public Sub LogPrint(ByVal msg As String)
' Fa log nella cartella dell'exe
Static NF As Long
Static LogFile As String, s As String, n As Long
Dim cSecs As String

#If Not LogON Then
    Exit Sub
#End If
    
Const EXT = ".LOG"
If LogFile = "" Then
    LogFile = App.path & App.EXEName & Format(Day(Date), "#0") _
    & Format(Month(Date), "#0") _
    & Format(Year(Date), "#0000")
End If

If NF = 0 Then
    Do
        NF = FreeFile
    Loop Until NF > 0
    s = LogFile
    
    While Dir(s & EXT) > ""
        n = n + 1
        s = LogFile & "(" & n & ")"
    Wend
    If n > 0 Then
        LogFile = LogFile & "(" & n & ")"
    End If
    
    Open LogFile & EXT For Output Access Write As #NF
    Debug.Print Now & " >>> logging on " & LogFile & EXT
End If
cSecs = CStr(Int(100 * (Timer - Int(Timer))))
Print #NF, Now & ":" & cSecs & " > " & msg

End Sub
