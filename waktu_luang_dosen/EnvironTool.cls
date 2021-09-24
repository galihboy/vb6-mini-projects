VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnvironTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2005
Option Explicit
' 22.8.2001 - questa classe sta diventando obsoleta
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Any) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Any) As Long

Private PerformanceFrequency As Currency

Public Property Get WindowsDirectory() As String
Attribute WindowsDirectory.VB_HelpID = 2006
WindowsDirectory = GetWinDir
End Property

Public Property Get SystemDirectory() As String
Attribute SystemDirectory.VB_HelpID = 2007
SystemDirectory = GetSysDir
End Property

Public Function HiTimer() As Double
Attribute HiTimer.VB_HelpID = 2008
Dim Ret As Long
Dim pc As Currency
If PerformanceFrequency = 0 Then
    Ret = QueryPerformanceFrequency(PerformanceFrequency)
End If
If PerformanceFrequency > 0 Then
    Ret = QueryPerformanceCounter(pc)
    Debug.Print pc, PerformanceFrequency
    If Ret > 0 Then HiTimer = pc / PerformanceFrequency
End If
End Function

Private Function GetWinDir() As String
Dim sBuf As String
Dim cSize As Long
Dim retval As Long
sBuf = String(255, 0)
cSize = 255

' Get Windows Directory.
retval = GetWindowsDirectoryA(sBuf, cSize)
If retval > 0 Then
    sBuf = Left(sBuf, retval)
    GetWinDir = sBuf
End If

End Function

Private Function GetSysDir() As String
Dim sBuf As String
Dim cSize As Long
Dim retval As Long
sBuf = String(255, 0)
cSize = 255

' Get System Directory.
If retval > 0 Then
    retval = GetSystemDirectoryA(sBuf, cSize)
    sBuf = Left(sBuf, retval)
End If
End Function

