VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrnTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2009
Option Explicit

Private Declare Function SetPrinter Lib "winspool.drv" _
Alias "SetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal Command As Long) As Long

Private Declare Function OpenPrinter Lib "winspool.drv" _
Alias "OpenPrinterA" _
(ByVal pPrinterName As String, _
phPrinter As Long, _
pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function GetPrinter Lib "winspool.drv" _
Alias "GetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal cbBuf As Long, _
pcbNeeded As Long) As Long

Private Declare Function lstrcpy Lib "kernel32" _
Alias "lstrcpyA" _
(ByVal lpString1 As String, _
ByVal lpString2 As Any) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" _
(ByVal hPrinter As Long) As Long

Private Declare Function AddPort Lib "winspool.drv" Alias "AddPortA" (ByVal pName As String, ByVal HWND As Long, ByVal pMonitorName As String) As Long
Private Declare Function DeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal HWND As Long, ByVal pPortName As String) As Long

' constants for DEVMODE structure
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Private Const PRINTER_ATTRIBUTE_DEFAULT = 4
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type PRINTER_DEFAULTS
        pDatatype As Long
        pDevMode As Long
        DesiredAccess As Long
End Type

Public Function WinSetPrinter(ByVal PrinterName As String, ByVal NewPrnName As String, ByVal NewPort As String) As String
Attribute WinSetPrinter.VB_HelpID = 2010
' ritorna "" in caso di corretto funzionamento
' altrimenti una descrizione dell'errore
' per win2000 e win98
Dim handle As Long          'handle to printer
Dim pd As PRINTER_DEFAULTS, X As Long, i As Long
Dim need As Long            ' bytes needed
Dim LastError As Long
Dim bArrayPort() As Byte
Dim bArrayPrnName() As Byte
Dim NewPortLen As Long, NewPrnNameLen As Long

NewPortLen = Len(NewPort)
NewPrnNameLen = Len(NewPrnName)
' determine which printer was selected
' none - exit
If PrinterName = "" Then
    WinSetPrinter = "No printer name avaible."
    Exit Function
End If

' set the PRINTER_DEFAULTS members
pd.pDatatype = 0&
pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

X = OpenPrinter(PrinterName, handle, pd)
' Get a handle to the printer
If X = 0 Then
    'error handler code goes here
    WinSetPrinter = "Unable to open the specified printer"
    Exit Function
End If

' Make an initial call to GetPrinter, requesting Level 2
' (PRINTER_INFO_2) information, to determine how many bytes
' you need
X = GetPrinter(handle, 2, ByVal 0&, 0, need)
If need < 4 Then
    WinSetPrinter = "Unable to get the printer buffer"
    Exit Function
End If
' don't want to check Err.LastDllError here - it's supposed
' to fail
' with a 122 - ERROR_INSUFFICIENT_BUFFER
' redim t as large as you need
ReDim t((need \ 4)) As Long
X = GetPrinter(handle, 2, t(0), need, need)
' and call GetPrinter for keepers this time
If X = 0 Then
    ' failed the GetPrinter
    WinSetPrinter = "Unable to get the printer structure"
    Exit Function
End If
' set the new port name
'ReDim bArrayPort(0 To NewPortLen) As Byte
'For i = 0 To NewPortLen - 1
'    bArrayPort(i) = Asc(Mid(NewPort, i + 1, 1))
'Next i
'bArrayPort(NewPortLen) = 0
bArrayPort = ByteZtStr(NewPort)
t(3) = VarPtr(bArrayPort(0))
' set new printer name
'ReDim bArrayPrnName(0 To NewPrnNameLen) As Byte
'For i = 0 To NewPrnNameLen - 1
'    bArrayPrnName(i) = Asc(Mid(NewPrnName, i + 1, 1))
'Next i
'bArrayPrnName(NewPrnNameLen) = 0
bArrayPrnName = ByteZtStr(NewPrnName)
t(1) = VarPtr(bArrayPrnName(0))

' call SetPrinter to set these changes
X = SetPrinter(handle, 2, t(0), 0&)

If X = False Then
    MsgBox "SetPrinterFailed. Error code: " & Err.LastDllError
    Exit Function
End If

' and close the handle
ClosePrinter (handle)

End Function

Public Function AddPrnPort(ByVal Port As String, ByVal Monitor As String, ByVal HWND As Long)
Attribute AddPrnPort.VB_HelpID = 2011
Dim bPort() As Byte
Dim bMonitor() As Byte
bPort = ByteZtStr(Port)
bMonitor = ByteZtStr(Monitor)
AddPrnPort = Not (AddPort(bPort(0), HWND, bMonitor(0)) = 0)
End Function

Public Function DelPrnPort(ByVal Port As String, ByVal Monitor As String, ByVal HWND As Long)
Attribute DelPrnPort.VB_HelpID = 2012
Dim bPort() As Byte
Dim bMonitor() As Byte
bPort = ByteZtStr(Port)
bMonitor = ByteZtStr(Monitor)
DelPrnPort = Not (DeletePort(bMonitor(0), HWND, bPort(0)) = 0)
End Function

Private Function ByteZtStr(ByVal Str) As Byte()
Dim bArr() As Byte
Dim L As Long, i As Long
L = Len(Str)

ReDim bArr(0 To L) As Byte
For i = 0 To L - 1
    bArr(i) = Asc(Mid(Str, i + 1, 1))
Next i
bArr(L) = 0
ByteZtStr = bArr
End Function
