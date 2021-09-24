Attribute VB_Name = "modXPStyle2"
' ____________________________________________________________________________________________________________
'|                                                                                                            |
'|                    In the name of Allah, the Merciful, the Compassionate.                                  |
'|                                ~~~~~~~~~~~~~~~~~~~~~~~~~~~                                                 |
'|                                    »”„ «··Â «·—Õ„‰ «·—ÕÌ„                                                  |
'|                                          ~~~~~~~~                                                          |
'|                                                                                                            |
'|                                      XPStyle Module                                                        |
'|                                       Version 2.00                                                         |
'|                                                                                                            |
'|                      * - This module was written by:                                                       |
'|                      -------------------------------                                                       |
'|                                              Voodoo Attack!!                                               |
'|                                     voodooattack@hotmail.com                                               |
'|                                                                                                            |
'|                  If you have any questions, feedback, thoughts or anything to share..                      |
'|                                  Please e-mail me immediately! :D                                          |
'|                                                                                                            |
'|____________________________________________________________________________________________________________|
'
'
'   Brief Note:
'   -----------
'
'       For people whom don't like to read much..
'       Just call the XPStyle function from "Sub Main()" or "Sub Form_Initialize()"
'       to activate VisualStyles..
'
'   .......................................................................................
'
' 1 - Overview:
' -------------
'
'   XPStyle is a module designed to give your program the ultimate feeling of comfort
'   When running under the Windows XP(or later) enviroment..
'
'   Windows XP uses a technique named "Visual Styles" to improve the GUI, it grants smooth
'   controls, improved graphics and more..
'
'   But, unfortunately, Visual Basic does not support this capability..
'   Finally Microsoft has provided the solution for this issue, a Manifest file..
'   (Search MSDN for more information)
'
'   A manifest file is a small XML file, That contains information about your program
'   It'll tell windows to skin your program once it starts, and switch the skinning task
'   to the system, so your program will appear properly under an OS that supports
'   visual styles..
'
'   But another proplem appeared, a manifest file should have the name "[myapp.exe].manifest"
'   and it MUST be in the same folder as the exe..
'
'   So you have to include the manifest file in your setup package..
'   And what if your program is a portable program(users may destribute it along)..
'   And what if it is a self-extractor, or a small utility..
'
'   That's why XPStyle exists..
'
'___________________________________________________________________________________________
'
' 2 - How does XPSyle work?
' -------------------------
'
'   XPStyle will write the manifest file immediately, and then, it will restart your
'   program with the same command-line arguments (if there's any).
'
'   (Windows will start skinning your program as long as the manifest file exist when
'   your program is launched, after that.. it'll be of no use.)
'
'   There is also an auto-hide manifest argument that will Hide the manifest file
'   immediately, so it will not be visible at all..
'
'   You may disable the auto-hide option if your program is running from a removable volume
'   (floppy disk, CD-ROM or a Network drive) to speed up the program loading..
'
'   (see the XPStyle function help-comments)
'
'   .......................................................................................
'
'   Version 2 of XPStyle is also inforced with extended theming support:
'
'       - You can now enable/disable theming while in run-time.
'       - You can draw the background of the TabStrip control on any picture box
'         If you use pictureboxes as containers for the TabStrip Control.
'       - You can use WinXP notifications on textboxes, MS RichEDIT, comboboxes or any
'         Control that supports this capability.
'
'
'       - More features will be included within the next release...!
'
'___________________________________________________________________________________________

'Sub Main()
'    'this should be the sequence of the "Sub Main" procedure..
'
'    XPStyle
'    frmMain.Show
'
'End Sub

Option Explicit

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (init As InitCommonControlsExType) As Boolean

Private Declare Function ActivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hwnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
Private Declare Function DeactivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hwnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
Private Declare Function IsThemeActiveXP Lib "uxtheme.dll" Alias "IsThemeActive" () As Boolean
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (hTheme As Long) As Long
Private Declare Function EnableThemeDialogTexture Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, Optional hrgnUpdate As Long, Optional fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Boolean

Private Const ETDT_DISABLE      As Long = &H1
Private Const ETDT_ENABLE       As Long = &H2

Private Const RDW_UPDATENOW     As Long = &H100

Private Const ICC_USEREX_CLASSES As Long = &H200

Private Const ECM_FIRST         As Long = &H1500
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)

Private m_bIsManifestActive     As Boolean
Private bIsVbRunning            As Boolean

Private Type InitCommonControlsExType
    dwSize  As Long     'size of this structure
    dwICC   As Long     'flags indicating which classes to be initialized
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type BALLOONTIP
    cbStruct As Long
    pszTitle As String
    pszText As String
    tIcon As Long
End Type

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

Enum BalloonTipIconConstants
    balNone = 0
    balExcalmation = 1
    balInformation = 2
    balCritical = 3
End Enum

Private Function InitCommonControls() As Boolean
    Dim InitCC As InitCommonControlsExType
    
    With InitCC
        .dwSize = Len(InitCC)
        .dwICC = ICC_USEREX_CLASSES
    End With
    
    InitCommonControls = InitCommonControlsEx(InitCC)         'initialize the common controls
End Function


Private Function CheckVB() As Boolean
    bIsVbRunning = True
    CheckVB = True
End Function


Private Function GetWindowTheme(hwnd As Long, Optional PartID As String) As Long
    'this will retrive the current hTheme used by the window..
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    Dim hTheme As Long
    If PartID = "" Then PartID = "Window"
    hTheme = OpenThemeData(hwnd, StrPtr(PartID))
    CloseThemeData hTheme
    GetWindowTheme = hTheme
    
End Function

Private Function GetWinVersion() As String
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function

Private Function AddDirSep(Path As String) As String
    Dim DirSep As String
    
    If PathIsNetworkPath(Path) = True Then
        DirSep = "/"
    Else
        DirSep = "\"
    End If
    
    If Right(Trim(Path), Len(DirSep)) <> DirSep Then
        AddDirSep = Trim(Path) & DirSep
    Else
        AddDirSep = Path
    End If
    
End Function


Function HideTextBalloonTip(Control As Control) As Boolean
    
    Dim hwnd As Long
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    Select Case UCase(TypeName(Control))
        
        Case "TEXTBOX"
            hwnd = Control.hwnd
        Case "RICHTEXTBOX"
            hwnd = Control.hwnd
        Case "COMBOBOX"
            If (Control.Style = 0 Or 1) Then
                Dim Cbo As COMBOBOXINFO
                Cbo.cbSize = Len(Cbo)
                Call GetComboBoxInfo(Control.hwnd, Cbo)
                hwnd = Cbo.hwndEdit
            Else
                Exit Function
            End If
        Case Else
            hwnd = Control.hwnd
    End Select
    
    HideTextBalloonTip = SendMessage(hwnd, EM_HIDEBALLOONTIP, 0&, 0&)

End Function

Function IsThemingSupported() As Boolean

    Dim hLib As Long                    'module handle..
    hLib = LoadLibrary("uxtheme.dll")   'retrive the module handle.
    Call FreeLibrary(hLib)              'unload the dll
    IsThemingSupported = CBool(hLib)    'if the return value = 0 then
                                        'the dll does not exist,
                                        'otherwise, the dll is there..
End Function





Function IsXPThemed(hwnd As Long) As Boolean
    
    'check if the object is using a visual style..
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    
    Dim hTheme As Long
        
    hTheme = OpenThemeData(hwnd, StrPtr("Window"))  'do the theme test
    
    Call CloseThemeData(hTheme)                     'close the theme data handle
    
    IsXPThemed = CBool(hTheme)                      'if zero, return False.. else return true..
    
    
End Function


Public Function ShowTextBalloonTip(Control As Control, Prompt As String, Optional Title As String, Optional TitleIcon As BalloonTipIconConstants) As Boolean
    
    'This function will show an EDIT balloon tip..
    'this function will only apply to a normal text box, a richtext box or a combobox
    'with syle 0 or 1...
    'any other controls passed to this function will return false (as i know!)
    
    Dim Bal As BALLOONTIP
    Dim hwnd As Long
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function

    Select Case UCase(TypeName(Control))
        Case "COMBOBOX"
            If (Control.Style = 0 Or Control.Style = 1) Then
                Dim Cbo As COMBOBOXINFO
                Cbo.cbSize = Len(Cbo)
                Call GetComboBoxInfo(Control.hwnd, Cbo)
                hwnd = Cbo.hwndEdit
            Else
                Exit Function
            End If
        Case "TEXTBOX"
            hwnd = Control.hwnd
        Case "RICHTEXTBOX"
            hwnd = Control.hwnd
        Case Else
            hwnd = Control.hwnd
    End Select
    
    With Bal
        .cbStruct = Len(Bal)                    'set the structure size
        .pszTitle = StrConv(Title, vbUnicode)   'convert the title into unicode format..
        .pszText = StrConv(Prompt, vbUnicode)   'convert the prompt into unicode format..
        .tIcon = TitleIcon                      'set the title icon
    End With
    
    'show the balloon tip..
    
    ShowTextBalloonTip = SendMessage(hwnd, EM_SHOWBALLOONTIP, 0&, Bal)
    
    
End Function

Function ToggleVisualStyles(Frm As Form, Enable As Boolean, Optional ToggleFormBorder As Boolean = True)
    
    'Enable/diable a form theming ..

    On Error GoTo ErrorHandler
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    
    Dim fControls   As String   'This is the forbidden controls (controls with no .hWnd/cannot be skinned)
                                'i didn't use an array cause i found out that it's too slow..
    
    Dim Ctl         As Control
    Dim sTypeName   As String
           
    For Each Ctl In Frm.Controls                            'loop through the controls collection of the form..
        sTypeName = UCase(TypeName(Ctl))                    'Get the typename of the control.
        If InStr(1, fControls, sTypeName) = 0 Then          'look for the control type name in the forbidden controls list, if found, do nothing..
            Select Case Enable                              'activate/deactivate theming
                Case True:  Call EnableXPLook(Ctl)
                Case False: Call DisableXPLook(Ctl)
            End Select
            If TypeName(Ctl) = "PICTUREBOX" Then Ctl.Refresh    'refresh any pictureboxes in the form..
        End If
    Next
    
    If ToggleFormBorder = True Then
        Select Case Enable                                      'activate/deactivate the form theming..
            Case True
                Call EnableXPLook(Frm)
            Case False
                Call DisableXPLook(Frm): Call DisableXPDlgBackground(Frm)
        End Select
    End If
    
    Set Ctl = Nothing       'erase the ctl variable from memory..

    Frm.Refresh             'refresh the form
    
   'Debug.Print fControls
    Exit Function

ErrorHandler:                                   'This is the error handling section...

    If Err.Number = 438 Then                    'object doesn't have a ".hWnd" property..
        'Debug.Print "Error: The Object '" & Ctl.Name & "' doesn't have a '.hwnd' property.."
        fControls = sTypeName & "," & fControls 'add this typename into the forbidden list..
        Resume Next                             'skip the line where the error happened, and proceed to the next line..
    Else                                        'unexpected error..
        Err.Raise Err.Number                    'show the error..
    End If
End Function


Function EnableXPLook(ByRef Object As Object) As Boolean
    'this function will draw the object using windows xp visual styles..
    'note: the object MUST have a handle
    
    On Error GoTo ErrHandler:

    Dim wRECT   As RECT
    
    GetWindowRect Object.hwnd, wRECT   'retrive the object region.
        
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    ActivateWindowTheme (Object.hwnd) 'try to enable theming
    
    If IsXPThemed(Object.hwnd) <> 0 Then
        'ok
        EnableXPLook = True
    Else
        'error
        GoTo ErrHandler
    End If
    
    Call RedrawWindow(Object.hwnd, wRECT, , RDW_UPDATENOW) 'refresh the object
   
    Exit Function
ErrHandler:
    EnableXPLook = False
    Exit Function
End Function

Function DisableXPLook(ByRef Object As Object) As Boolean
    'this function will disable the object's visual style..
    'note: the object MUST have a handle
    'same as the EnableXPLook function..
    
    Dim wRECT As RECT
    
    On Error GoTo ErrHandler:
    
    GetWindowRect Object.hwnd, wRECT
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    DeactivateWindowTheme (Object.hwnd)
    
    If IsXPThemed(Object.hwnd) = False Then
        DisableXPLook = True
    Else
        GoTo ErrHandler
    End If
    
    Call RedrawWindow(Object.hwnd, wRECT, , RDW_UPDATENOW)
    
    Exit Function
    
ErrHandler:
    DisableXPLook = False
    Exit Function
End Function

Function DrawTabBackground(oPictureBox As Object, Optional sTab As Object)
    
    On Error Resume Next
    'Draw a TabStrip control's background texture in a picture box..
    'this is a good example on how to draw controls using "uxtheme.dll" API calls..
    
    Dim hTheme          As Long         'The theme handle
    Dim dRECT           As RECT         'The drawing Region
    Dim tabHwnd         As Long
    Const TAB_BODY      As Integer = 10 'this is the PartID of the tabstrip background..
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    tabHwnd = sTab.hwnd
    
    If tabHwnd <> 0 Then
        If IsXPThemed(sTab.hwnd) = False Then oPictureBox.Cls: Exit Function  'if the frame theming is disabled, Clear the picture box and exit
    End If
    
    oPictureBox.Cls
    oPictureBox.AutoRedraw = False
    
    'copy the picturebox measurements into the RECT object
    
    dRECT.Left = 0
    dRECT.Top = 0
    dRECT.Right = oPictureBox.ScaleX(oPictureBox.Width, oPictureBox.ScaleMode, vbPixels)
    dRECT.Bottom = oPictureBox.ScaleY(oPictureBox.Height, oPictureBox.ScaleMode, vbPixels)

    hTheme = OpenThemeData(oPictureBox.hwnd, StrPtr("TAB"))      'Retrive the handle of the current theme being used.
    
    If hTheme <> 0 Then
        Call DrawThemeBackground(hTheme, oPictureBox.hDC, TAB_BODY, 0, dRECT, dRECT) 'draw the tab background on the picture box
    Else
        oPictureBox.Cls
    End If
    
    oPictureBox.AutoRedraw = True
    
    CloseThemeData hTheme           'close the theme data handle..
    
End Function


Sub EnableXPDlgBackground(Frm As Form)
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Sub
    If IsThemeActiveXP() = False Then Exit Sub
    Call EnableThemeDialogTexture(Frm.hwnd, ETDT_ENABLE)
End Sub

Sub DisableXPDlgBackground(Form As Form)
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Sub
    If IsThemeActiveXP() = False Then Exit Sub
    Call EnableThemeDialogTexture(Form.hwnd, ETDT_DISABLE)
End Sub

Public Function IsVBRunning() As Boolean
    
    'this function will check if vb is running..
    'I made it public cause you maight need it in your project..
    
    Debug.Assert (CheckVB) = True   '-  this works only if vb is running..
                                    '   the original purpose of the "assert"
                                    '   function is to make sure that a function(or variable)
                                    '   will return the specified value .. else, the
                                    '   program will pause..
                                    
                                    '-  what we do here is:
                                    '   call another function, "CheckVB"
                                    '   through the "assert" function
                                    '   that will set the value of the
                                    '   "bIsVbRunning" to true if called..
                                    '   ("Assert" will only call the function
                                    '   when VB is running "Debug mode"..)
                                    
                                    '   I Hope this would be useful somehow.. :D

    IsVBRunning = bIsVbRunning
    bIsVbRunning = False
    
End Function


Private Function IsWindowsXP() As Boolean
    If Val(Trim(GetWinVersion)) >= 5.01 Then
        IsWindowsXP = True
    End If
End Function




Private Function vb5Replace(Expression As String, Find As String, ReplaceWith As String, Optional Start As Long = 1, Optional Count As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    
On Error GoTo ExitProcedure

    'I wrote this function to add the "replace" capability to vb5
    'so you can now use XPStyle module in vb5
    'NOTE: IF YOU USE VB6, YOU CAN COMMENT THIS FUNCTION..
    
    Dim iFind As Long
    Dim nextStart As Long
    Dim sCount As Long
    
    iFind = InStr(Start, Expression, Find, Compare)
    
    nextStart = Start
    
    If iFind = -1 Then
        vb5Replace = Expression
        Exit Function
    Else
        Do
            If sCount >= Count And Count <> -1 Then
                Exit Do
            End If
            iFind = InStr(nextStart, Expression, Find, Compare)
            If iFind = 0 Then Exit Do
            sCount = sCount + 1
            Expression = Left(Expression, iFind - 1) & ReplaceWith & Mid(Expression, iFind + Len(Find))
            If nextStart < Len(Expression) Then
                nextStart = iFind + Len(ReplaceWith) + 1
            Else
                Exit Do
            End If
            
        Loop
    End If

ExitProcedure:
    
    vb5Replace = Expression
    
End Function

' ______________________________________________________________________________________________________________
'|                                                                                                              |
'| This is the main function of this module..                                                                   |
'| This function will write the program manifest, restart the program(optional),                                |
'| delete the manifest(optional too)..                                                                          |
'|______________________________________________________________________________________________________________|
'|                                                                                                              |
'|   [ Parameter ]  +   [ Description ]                                                                         |
'|______________________________________________________________________________________________________________|
'|                  |                                                                                           |
'|   [Autorestart] -|    [True]: The program will automatically restart.  (the visual styles will be activated) |
'|                  |   [False]: The program will continue normally.      (will not be skinned till restarted)  |
'|                  |                                                                                           |
'|                  |           If set to [True], CreateNew will be set to [False] by default.                  |
'| -----------------|------------------------------------------------------------------------------------------ |
'|                  |                                                                                           |
'|   [Autohide   ] -|    [True]: The manifest file will not be visible.                                         |
'|                  |   [False]: The manifest file will remain.                                                 |
'|                  |                                                                                           |
'|                  |           If set to [True], CreateNew will be set to [False] by default.                  |
'| -----------------|------------------------------------------------------------------------------------------ |
'|                  |                                                                                           |
'|   [CreateNew  ] -|    [True]: Write a fresh manifest file.                                                   |
'|                  |   [False]: Nothing.                                                                       |
'|                  |                                                                                           |
'|                  |           Only for use with Autorestart=[False] Or Autohide=[False]                       |
'|__________________l___________________________________________________________________________________________|
'

Function XPStyle(Optional AutoRestart As Boolean = True, Optional Autohide As Boolean = True, Optional CreateNew As Boolean = False) As Boolean
    
    If IsWindowsXP = False Or IsVBRunning Or IsThemingSupported = False Then Exit Function
    If IsThemeActiveXP = False Then Exit Function
    
    Const IsVB6 As Boolean = True   'change to false if you are using vb5
    
    On Error Resume Next
    
    Dim XML             As String
    Dim ManifestCheck   As String
    Dim strManifest     As String
    Dim FreeFileNo      As Integer
    
    If AutoRestart = True Or Autohide = True Then CreateNew = False
    
    '(put the XML into a string)
    XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & "<assembly " & vbCrLf & "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & "   manifestVersion=""1.0"">" & vbCrLf & "<assemblyIdentity " & vbCrLf & "    processorArchitecture=""x86"" " & vbCrLf & "    version=""EXEVERSION""" & vbCrLf & "    type=""win32""" & vbCrLf & "    name=""COMPANYNAME.EXENAME""/>" & vbCrLf & "    <description>EXEDESCRIPTION</description>" & vbCrLf & "    <dependency>" & vbCrLf & "    <dependentAssembly>" & vbCrLf & "    <assemblyIdentity" & vbCrLf & "         type=""win32""" & vbCrLf & "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & "         version=""6.0.0.0""" & vbCrLf & "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & "         language=""*""" & vbCrLf & "         processorArchitecture=""x86""/>" & vbCrLf & "    </dependentAssembly>" & vbCrLf & "    </dependency>" & vbCrLf & "</assembly>" & vbCrLf & "")
    
    strManifest = AddDirSep(App.Path) & App.EXEName & ".exe.manifest"        'retrive the name of the manifest file
    
    ManifestCheck = Dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive) 'look for the manifest file.
    
    If ManifestCheck = "" Or CreateNew = True Then           'if not found.. or the "CreateNew" parameter is set to true, make a new one..
        
        If IsVB6 = True Then                                 'see if vb is version 6 or not!
            'VISUAL BASIC  6
            XML = Replace(XML, "EXENAME", App.EXEName & ".exe")             'Replace the string "EXENAME" with the program's exe file name.
            XML = Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0") 'Replace the "EXEVERSION" string.
            XML = Replace(XML, "EXEDESCRIPTION", App.FileDescription)       'Replace the app DESCRIPTION.
            XML = Replace(XML, "COMPANYNAME", App.CompanyName)
        Else
            'VISUAL BASIC  5
            XML = vb5Replace(XML, "EXENAME", App.EXEName & ".exe")
            XML = vb5Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0")
            XML = vb5Replace(XML, "EXEDESCRIPTION", App.FileDescription)
            XML = vb5Replace(XML, "COMPANYNAME", App.CompanyName)
        End If
        
        FreeFileNo = FreeFile                           'get the next free file
        If ManifestCheck <> "" Then                     'this should be the "CreateNew" mode..
            SetAttr strManifest, vbNormal
            Kill (strManifest)
        End If
        
        Open strManifest For Binary As #(FreeFileNo)    'open the file in binary mode
            Put #(FreeFileNo), , XML                    'use "put" to write to the file.. note that "put" (binary mode) is much faster than "print"(output mode)..
        Close #(FreeFileNo)                             'close the file.
        SetAttr strManifest, vbHidden + vbSystem        'set the file attributes to "Hidden, System"
        XPStyle = False                                 'return false.. this means that the application is not yet using visual styles..
        
        If AutoRestart = True Then                      'if in automode (default), the program will restart.
        
            Shell App.Path & "\" & App.EXEName & ".exe" & _
            Space(1) & Command$, vbNormalFocus
                                            
                                                        'restart the program and bypass command line parameters (if any)..
            End                                         'end the session.
        End If
        
    Else                                                'the manifest file exists.
    
        
        If Autohide = True Then                         'if the autohide mode is enabled (default), then we should delete the file..
                                                        'in normal conditions, the manifest file will not appear at all ;)
            SetAttr strManifest, vbNormal
            Kill (strManifest)
        End If
        
        XPStyle = True
        
    End If
        m_bIsManifestActive = XPStyle
End Function

Public Property Get IsManifestActive() As Boolean
    
    IsManifestActive = m_bIsManifestActive

End Property

