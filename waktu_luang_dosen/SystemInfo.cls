VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2069
Option Explicit

Public Enum DRIVE_TYPE
    DRIVE_DOESNT_EXIST = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5
    DRIVE_RAMDISK = 6
End Enum

Public Enum OS_VERSION
    OS_WINDOWS_UNKNOWN
    OS_WINDOWS_3X
    OS_WINDOWS_95
    OS_WINDOWS_98
    OS_WINDOWS_NT3X
    OS_WINDOWS_NT40
    OS_WINDOWS_2000
End Enum


Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const BITSPIXEL = 12
Private Const PLANES = 14

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private pUdtOSVersion As OSVERSIONINFOEX
Private pUdtMemStatus As MEMORYSTATUS

Private plMajorVersion  As Long
Private plMinorVersion As Long
Private plPlatformID As Long

Private psComputerName As String
Private plLastDllError As Long

Private Function FreeBytesOnDisk(Drive As String) As Long
    On Error Resume Next
    plLastDllError = 0
    
    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lFreeBytes As Double
    
    
    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
    lFreeBytes = lBytesPerCluster * lFreeClusters
    FreeBytesOnDisk = lFreeBytes
    plLastDllError = Err.LastDllError
End Function

Public Function WindowsDir() As String
Attribute WindowsDir.VB_HelpID = 2070
    Dim temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    temp = String$(MAX_LENGTH, 0)
    Ret = GetWindowsDirectory(temp, MAX_LENGTH)
    temp = Left$(temp, Ret)
    If temp <> "" And Right$(temp, 1) <> "\" Then
        WindowsDir = temp & "\"
    Else
        WindowsDir = temp
    End If
End Function

Public Function SystemDir() As String
Attribute SystemDir.VB_HelpID = 2071
    Dim temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    temp = String$(MAX_LENGTH, 0)
    Ret = GetSystemDirectory(temp, MAX_LENGTH)
    temp = Left$(temp, Ret)
    If temp <> "" And Right$(temp, 1) <> "\" Then
        SystemDir = temp & "\"
    Else
        SystemDir = temp
    End If
End Function

Private Function TotalBytesOnDisk(Drive As String) As Double
    On Error Resume Next
    plLastDllError = 0
    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lTotalBytes As Double

    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
    'dblAns = (Bytes / 1024) / 1024
    TotalBytesOnDisk = lBytesPerCluster * lTotalClusters
    If TotalBytesOnDisk = 0 Then
        TotalBytesOnDisk = ((lBytesPerCluster / 1024) / 1024) * lTotalClusters
    End If
    plLastDllError = Err.LastDllError
End Function

Public Function FreeMegaBytesOnDisk(Optional Drive As String = "C:\") As Double
Attribute FreeMegaBytesOnDisk.VB_HelpID = 2072
    'some time in the future disk may be to large to calculate
    'like this so resume next on any errors
    On Error Resume Next

    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long

    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lFreeBytes As Double

    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"

    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector

    FreeMegaBytesOnDisk = ((lBytesPerCluster / 1024) / 1024) * lFreeClusters
    FreeMegaBytesOnDisk = Format(FreeMegaBytesOnDisk, "###,###,##0.00")
End Function

Public Function TotalMegaBytesOnDisk(Optional Drive As String = "C:\") As Double
Attribute TotalMegaBytesOnDisk.VB_HelpID = 2073
    'some time in the future disk may be to large to calculate
    'like this so resume next on any errors
    On Error Resume Next

    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lTotalBytes As Double
    
    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    
    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
    
    TotalMegaBytesOnDisk = ((lBytesPerCluster / 1024) / 1024) * lTotalClusters
    TotalMegaBytesOnDisk = Format(TotalMegaBytesOnDisk, "###,###,##0.00")
End Function

Public Property Get HasSoundCard() As Boolean
Attribute HasSoundCard.VB_HelpID = 2074
    plLastDllError = 0
    HasSoundCard = (waveOutGetNumDevs > 0)
    plLastDllError = Err.LastDllError
End Property

Public Property Get ComputerName() As String
Attribute ComputerName.VB_HelpID = 2075
    Dim sBuffer As String
    Dim lAns As Long

    plLastDllError = 0
    sBuffer = Space$(255)
    lAns = GetComputerName(sBuffer, 255)
    If lAns <> 0 Then
        'read from beginning of string to null-terminator
        ComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        plLastDllError = Err.LastDllError
    End If
End Property

Public Property Get CurrentUser() As String
Attribute CurrentUser.VB_HelpID = 2076
    Dim L As Long
    Dim sUser As String

    plLastDllError = 0
    sUser = Space(255)
    L = GetUserName(sUser, 255)
    'strip null terminator
    If L <> 0 Then
        CurrentUser = Left(sUser, InStr(sUser, Chr(0)) - 1)
    Else
        plLastDllError = Err.LastDllError
    End If
End Property

Private Sub Class_Initialize()
    pUdtOSVersion.dwOSVersionInfoSize = Len(pUdtOSVersion)
    GetVersionEx pUdtOSVersion
    plMajorVersion = pUdtOSVersion.dwMajorVersion
    plMinorVersion = pUdtOSVersion.dwMinorVersion
    plPlatformID = pUdtOSVersion.dwPlatformId
End Sub
Public Function VolumeName(Optional Drive As String = "C:\")
Attribute VolumeName.VB_HelpID = 2077
    Dim sBuffer As String

    plLastDllError = 0
    sBuffer = Space$(255)
    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    If GetVolumeInformation(Drive, sBuffer, Len(sBuffer), 0, 0, 0, Space$(255), 255) = 0 Then
        plLastDllError = Err.LastDllError
    Else
        VolumeName = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
    End If
End Function

Public Property Get OSVersion() As OS_VERSION
Attribute OSVersion.VB_HelpID = 2078
    On Error GoTo ErrorHandler

    plLastDllError = 0
    Select Case plMajorVersion
        Case 5: OSVersion = OS_WINDOWS_2000 'UNTESTED
        Case 4
            If plPlatformID = VER_PLATFORM_WIN32_NT Then
                OSVersion = OS_WINDOWS_NT40
            Else
                OSVersion = IIf(plMinorVersion = 0, OS_WINDOWS_95, OS_WINDOWS_98)
            End If
        Case 3
            If plPlatformID = VER_PLATFORM_WIN32s Then
                OSVersion = OS_WINDOWS_3X
            ElseIf plPlatformID = VER_PLATFORM_WIN32_NT Then
                OSVersion = OS_WINDOWS_NT40
            End If
        Case Else:  OSVersion = OS_WINDOWS_UNKNOWN
    End Select
Exit Property
ErrorHandler:
    OSVersion = OS_WINDOWS_UNKNOWN
    plLastDllError = Err.LastDllError
End Property

Public Function DriveType(Drive As String) As DRIVE_TYPE
Attribute DriveType.VB_HelpID = 2079
    'fix bad parameter values
    plLastDllError = 0
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    DriveType = GetDriveType(Drive)
    plLastDllError = Err.LastDllError
End Function

Public Function AvailablePhysicalMemory() As Double
Attribute AvailablePhysicalMemory.VB_HelpID = 2080
    'Return Value in Megabytes
    Dim dblAns As Double
    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    AvailablePhysicalMemory = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function TotalPhysicalMemory() As Double
Attribute TotalPhysicalMemory.VB_HelpID = 2081
    'Return Value in Megabytes
    Dim dblAns As Double
    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPhys
    TotalPhysicalMemory = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function PercentMemoryFree() As Double
Attribute PercentMemoryFree.VB_HelpID = 2082
    PercentMemoryFree = Format(AvailableMemory / TotalMemory * 100, "0#")
End Function

Public Function AvailablePageFile() As Double
Attribute AvailablePageFile.VB_HelpID = 2083
    'Return Value in Megabytes
    Dim dblAns As Double

    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPageFile
    AvailablePageFile = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function PageFileSize() As Double
Attribute PageFileSize.VB_HelpID = 2084
    'Return Value in Megabytes
    Dim dblAns As Double

    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPageFile
    PageFileSize = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function AvailableMemory() As Double
Attribute AvailableMemory.VB_HelpID = 2085
    'Return Value in Megabytes
     AvailableMemory = AvailablePhysicalMemory + AvailablePageFile
End Function

Public Function TotalMemory() As Double
Attribute TotalMemory.VB_HelpID = 2086
    'Return Value in Megabytes
    TotalMemory = PageFileSize + TotalPhysicalMemory
End Function

Private Function BytesToMegabytes(bytes As Double) As Double
  Dim dblAns As Double
  dblAns = (bytes / 1024) / 1024
  BytesToMegabytes = Format(dblAns, "###,###,##0.00")
End Function

Public Property Get MaxScreenColors(ByVal hDC As Long) As Double
Attribute MaxScreenColors.VB_HelpID = 2087
    'Returns the maximum number of colors supported
    'by the system - e.g.,  256, 16,777,216
    Dim lngBits As Long
    Dim lngPlanes As Long
    Dim lwndHandle As Long
    Dim dblAns As Double

    plLastDllError = 0
    lwndHandle = hDC
    'bits per pixel
    lngBits = GetDeviceCaps(lwndHandle, BITSPIXEL)
    'number of color planes
    lngPlanes = GetDeviceCaps(lwndHandle, PLANES)
    'maximum colors available
    MaxScreenColors = (2 ^ (lngBits * lngPlanes))
    plLastDllError = Err.LastDllError
End Property

Public Property Get ScreenPixelWidth() As Integer
Attribute ScreenPixelWidth.VB_HelpID = 2088
    plLastDllError = 0
    ScreenPixelWidth = Screen.Width \ Screen.TwipsPerPixelX
End Property

Public Property Get ScreenPixelHeight() As Integer
Attribute ScreenPixelHeight.VB_HelpID = 2089
    plLastDllError = 0
    ScreenPixelHeight = Screen.Height \ Screen.TwipsPerPixelY
End Property

Public Property Get ScreenResolution() As String
Attribute ScreenResolution.VB_HelpID = 2090
    plLastDllError = 0
    ScreenResolution = ScreenPixelWidth & " X " & ScreenPixelHeight
End Property

Public Property Get SystemErrorCode() As Long
Attribute SystemErrorCode.VB_HelpID = 2091
    SystemErrorCode = plLastDllError
End Property
