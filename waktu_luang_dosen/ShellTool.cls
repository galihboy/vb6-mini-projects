VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2002
Attribute VB_Description = "This class adds some improvements to original VB6 shell command."
Option Explicit
' ##MODULE_DESCRIPTION This class adds some improvements to original VB6 shell command.

' > WaitOnProgram ha come primo paramerto l'ID del programma lanciato
' ottenuto dalla shell; se waitDead = true tutti gli eventi sono
' bloccati fino alla fine del programma lanciato dalla shell
' se waitDead = false allora gli eventi vengono eseguiti
' ExecCmd lancia un programma e ne attende il termine bloccando gli
' eventi
' TODO: impostare un timeout

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE As Long = &H100000
Private Const PROCESS_ALL_ACCESS  As Long = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Private Const INFINITE  As Long = &HFFFFFFFF         ' Infinite timeout
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const STATUS_PENDING As Long = &H103
Private Const STILL_ACTIVE As Long = STATUS_PENDING

' -- -- --
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&


Public Function ExecCmd(CommandLine As String) As Long
Attribute ExecCmd.VB_Description = "Execute an external program and waits for its execution. Returns the exit code of the process"
Attribute ExecCmd.VB_HelpID = 2003
'##PARAMETER_DESCRIPTION CommandLine The full path of the program ( eg: notepad.exe )
'##BLOCK_DESCRIPTION Execute an external program and waits for its execution. _
Returns the exit code of the process
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim Ret As Long

' Initialize the STARTUPINFO structure:
start.cb = Len(start)

' Start the shelled application:
Ret = CreateProcessA(0&, CommandLine, 0&, 0&, 1&, _
   NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)


' Wait for the shelled application to finish:
   Ret = WaitForSingleObject(proc.hProcess, INFINITE)
   Call GetExitCodeProcess(proc.hProcess, Ret&)
   Call CloseHandle(proc.hThread)
   Call CloseHandle(proc.hProcess)
   ExecCmd = Ret
End Function

Public Function WaitOnProgram(ByVal idProg As Long, Optional ByVal WaitDead As Boolean) As Long
Attribute WaitOnProgram.VB_Description = "Wait for an external program launched by a shell call and waits for its execution."
Attribute WaitOnProgram.VB_HelpID = 2004
'##PARAMETER_DESCRIPTION idProg The program id returned by a shell call.
'##PARAMETER_DESCRIPTION WaitDead If True, no events will be processed until _
program terimation
'##BLOCK_DESCRIPTION Wait for an external program launched by a shell _
call and waits for its execution.

Dim cRead As Long, iExit As Long, hProg As Long    ' Get process handle
Dim iResult As Long

hProg = OpenProcess(PROCESS_ALL_ACCESS, False, idProg)

If WaitDead Then        ' Stop dead until process terminates
    iResult = WaitForSingleObject(hProg, INFINITE)
    If iResult = WAIT_FAILED Then Err.Raise Err.LastDllError        ' Get the return value
    GetExitCodeProcess hProg, iExit
Else        ' Get the return value
    GetExitCodeProcess hProg, iExit        ' Wait, but allow painting and other processing
    Do While iExit = STILL_ACTIVE
        DoEvents
        GetExitCodeProcess hProg, iExit
    Loop
End If
CloseHandle hProg
WaitOnProgram = iExit
End Function


