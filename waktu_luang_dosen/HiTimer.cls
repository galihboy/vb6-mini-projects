VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HiTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2042
'-------------------------------------------
' PerformanceTimer class module
'-------------------------------------------

' Use this class to profile your code and any other operation
' typically with a precision greater than 1 millionth of a second
'
' As soon as you create an object, the timer starts
' but you can also start it explicitly with StartTimer
' Stop the timer and retrieve timing with StopTimer, or
' get the timing without stopping the timer with ElapsedTime
'
' The TotalTime property returns the number of seconds the
' timer has been active, so you can use it to sum up partial
' timings, after swithing the timer on and off
' The FormatTime is similar to elapsed time, but returns
' the time as a formatted string with desired precision
'
' Example:
'         Dim pc As New PerformanceCounter
'         pc.StartTimer
'         ' ...
'         ' put here the code you want to benchmark
'         ' ...
'         ' print elapsed time, but don't stop the timer
'         Debug.Print pc.ElapsedTime
'         ' ...
'         ' so something else here
'         ' ...
'         ' print elapsed time and stop the timer
'         Debug.Print pc.StopTimer
'         ' ...
'         ' prepare another benchmark here
'         ' ...
'         ' start the benchmark, without resetting total time
'         pc.StartTimer
'         ' ...
'         ' put here the code you want to benchmark
'         ' ...
'         ' print elapsed as a formatted string
'         Debug.Print pc.FormatTime("Second benchmark ### secs.", 4)
'         ' print total time
'         Debug.Print pc.TotalTime
'

Option Explicit

Private Declare Function QueryPerformanceFrequencyAny Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (lpFrequency As Any) As Long
Private Declare Function QueryPerformanceCounterAny Lib "kernel32" Alias _
    "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long

' the frequency for this computer
Dim frequency As Currency
Dim startTime As Currency
Dim endTime As Currency
Dim totTime As Currency

' Start the timer
'
' if argument is True, it also resets the
' internal total time counter

Sub StartTimer(Optional ByVal ResetTotalTime As Boolean)
Attribute StartTimer.VB_HelpID = 2043
    ' get the current value of the counter
    QueryPerformanceCounterAny startTime
    ' reset total time counter if requested
    If ResetTotalTime Then totTime = 0
End Sub

' stop the timer
'
' returns the time elapsed since StartTimer

Function StopTimer() As Double
Attribute StopTimer.VB_HelpID = 2044
    ' get the elapsed time
    StopTimer = ElapsedTime
    ' update the total time counter
    totTime = totTime + (endTime - startTime)
    ' reset starting time
    startTime = 0
End Function

' return the elapsed time in seconds since StartTimer
' without stopping the timer

Property Get ElapsedTime() As Double
Attribute ElapsedTime.VB_HelpID = 2045
    ' exit if StartTimer hasn't been called since
    ' the previous call to StopTimer
    If startTime = 0 Then Exit Property
    
    ' get the current value of the counter
    QueryPerformanceCounterAny endTime
    ' return the elapsed time in seconds
    ElapsedTime = (endTime - startTime) / frequency
End Property

' return the total time in seconds

Property Get TotalTime() As Double
Attribute TotalTime.VB_HelpID = 2046
    If startTime = 0 Then
        ' StopTimer has been called
        ' so totTime is correctly updated
        TotalTime = totTime / frequency
    Else
        TotalTime = (totTime + (endTime - startTime)) / frequency
    End If
End Property

' return a time value as a formatted string
' if second argument is omitted, it uses ElapsedTime
'
' return it as a formatted string with
' specified number of decimal - use ### in the string
' as a placeholder for the elapsed time
'   e.g.  Print GetTimeMsg("Elapsed ### secs.", , 4)
'
' NOTE: this function is slightly less precise than
'       GetTime, because arguments are passed

Property Get FormatTime(msg As String, Optional seconds As Double = -1, _
    Optional ByVal decDigits As Integer = 7) As String
Attribute FormatTime.VB_HelpID = 2047
    ' get the elapsed time if not passed as an argument
    If seconds < 0 Then seconds = ElapsedTime()
    ' build the result string
    FormatTime = Replace(msg, "###", CStr(Round(seconds, decDigits)))
End Property

' return the timer precision in seconds

Property Get Precision() As Double
Attribute Precision.VB_HelpID = 2048
    ' frequency must be scaled up by 10E4
    Precision = 1 / (frequency * 10000#)
End Property

' evaluate the frequency once and for all
' when this object is created

Private Sub Class_Initialize()
    ' raise error if API functions aren't supported
    If QueryPerformanceFrequencyAny(frequency) = 0 Then
        Err.Raise 1001, , "This system doesn't support high-res timing"
    End If
    'Debug.Print "HITIMER: freq="; frequency & "; realOne = " & GetRealFrequency(0.5)
    ' get start time as well
    StartTimer
End Sub



