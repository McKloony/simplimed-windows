Attribute VB_Name = "basWindow"
Option Explicit

Private FM As VB.Form
Private FS As VB.Form
Private Const SND_ASYNC = &H1 'Sounds
Private Const SND_NODEFAULT = &H2
Private Const SND_ALIAS = &H10000

Private Const MM_TEXT = 1
Private Const WM_DROPFILES = &H233
Private Const GWL_WNDPROC As Long = -4&
Private Const WM_GETMINMAXINFO As Long = &H24&
Private Const WM_DESTROY As Long = &H2&
Private Const WM_NCLBUTTONDOWN As Long = &HA1&
Private Const WM_NCMOUSEMOVE As Long = &HA0&
Private Const WM_DISPLAYCHANGE As Long = &H7E
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_SYSCOMMAND As Long = &H112&
Private Const SC_CLOSE As Long = &HF060&

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const WAIT_TIMEOUT = &H102
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SRCCOPY = &HCC0020
Private Const VER_PLATFORM_WIN32s As Long = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT As Long = 2&
Private Const VER_SUITE_PERSONAL As Long = &H200&

Private Const MB_ABORTRETRYIGNORE = &H2&
Private Const MB_APPLMODAL = &H0&
Private Const MB_COMPOSITE = &H2
Private Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Private Const MB_DEFBUTTON1 = &H0&
Private Const MB_DEFBUTTON2 = &H100&
Private Const MB_DEFBUTTON3 = &H200&
Private Const MB_DEFMASK = &HF00&
Private Const MB_ICONASTERISK = &H40&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONHAND = &H10&
Private Const MB_ICONMASK = &HF0&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_MISCMASK = &HC000&
Private Const MB_MODEMASK = &H3000&
Private Const MB_NOFOCUS = &H8000&
Private Const MB_OK = &H0&
Private Const MB_OKCANCEL = &H1&
Private Const MB_PRECOMPOSED = &H1
Private Const MB_RETRYCANCEL = &H5&
Private Const MB_SETFOREGROUND = &H10000
Private Const MB_SYSTEMMODAL = &H1000&
Private Const MB_TASKMODAL = &H2000&
Private Const MB_TYPEMASK = &HF&
Private Const MB_USEGLYPHCHARS = &H4
Private Const MB_YESNO = &H4&
Private Const MB_YESNOCANCEL = &H3&

Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7

Private Const CB_SETCURSEL = &H14E
Private Const CB_FINDSTRING = &H14C&
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157

Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_COPYRETURNORG = &H4
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1

Private Const ILC_COLOR = &H0
Private Const ILC_MASK = &H1
Private Const ILC_COLOR4 = &H4
Private Const ILC_COLOR8 = &H8
Private Const ILC_COLOR16 = &H10
Private Const ILC_COLOR24 = &H18
Private Const ILC_COLOR32 = &H20
Private Const ILD_NORMAL = 0

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Public Const Dial1 = MB_ICONQUESTION Or MB_YESNO Or MB_SYSTEMMODAL
Public Const Dial2 = MB_ICONASTERISK Or MB_SYSTEMMODAL
Public Const Dial3 = MB_ICONEXCLAMATION Or MB_SYSTEMMODAL
Public Const Dial4 = MB_ICONHAND Or MB_SYSTEMMODAL

Private TiHw1 As Long
Private TiHw2 As Long
Private TiHw3 As Long
Private TiHw4 As Long
Private TiHw5 As Long
Private TiHw6 As Long
Private TiHw7 As Long
Private TiHw8 As Long
Private TiZal As Long
Private WProc As Long
Private MProc As Long
Private mVorh As Boolean
Private mVers As String
Private SizPa As SIZEPAR

Private Const TOKEN_READ& = &H20008
Private Const TOKENELEVATION = 8
Private Const TOKENELEVATIONTYPE = 18

Private Enum TOKEN_ELEVATION_TYPE
    TokenElevationTypeDefault = 1
    TokenElevationTypeFull
    TokenElevationTypeLimited
End Enum

Public GlMou As Boolean

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public Type SIZEPAR
    xMin As Long
    yMin As Long
    xMax As Long
    yMax As Long
End Type

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
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Service Pack
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Service Pack
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type TEXTMETRIC
    tmHeight As Integer
    tmAscent As Integer
    tmDescent As Integer
    tmInternalLeading As Integer
    tmExternalLeading As Integer
    tmAveCharWidth As Integer
    tmMaxCharWidth As Integer
    tmWeight As Integer
    tmItalic As String * 1
    tmUnderlined As String * 1
    tmStruckOut As String * 1
    tmFirstChar As String * 1
    tmLastChar As String * 1
    tmDefaultChar As String * 1
    tmBreakChar As String * 1
    tmPitchAndFamily As String * 1
    tmCharSet As String * 1
    tmOverhang As Integer
    tmDigitizedAspectX As Integer
    tmDigitizedAspectY As Integer
End Type

Private Const SPI_GETACCESSTIMEOUT = &H3C
Private Const SPI_GETACTIVEWINDOWTRACKING = &H1000
Private Const SPI_GETACTIVEWNDTRKTIMEOUT = &H2002
Private Const SPI_GETACTIVEWNDTRKZORDER = &H100C
Private Const SPI_GETANIMATION = &H48
Private Const SPI_GETBEEP = &H1
Private Const SPI_GETBORDER = &H5
Private Const SPI_GETCOMBOBOXANIMATION = &H1004
Private Const SPI_GETDEFAULTINPUTLANG = &H59
Private Const SPI_GETDRAGFULLWINDOWS = &H26
Private Const SPI_GETFASTTASKSWITCH = &H23
Private Const SPI_GETFILTERKEYS = &H32
Private Const SPI_GETFOREGROUNDFLASHCOUNT = &H2004
Private Const SPI_GETFOREGROUNDLOCKTIMEOUT = &H2000
Private Const SPI_GETGRADIENTCAPTIONS = &H1008
Private Const SPI_GETGRIDGRANULARITY = &H12
Private Const SPI_GETHIGHCONTRAST = &H42
Private Const SPI_GETHOTTRACKING = &H100E
Private Const SPI_GETICONMETRICS = &H2D
Private Const SPI_GETICONTITLELOGFONT = &H1F
Private Const SPI_GETICONTITLEWRAP = &H19
Private Const SPI_GETKEYBOARDDELAY = &H16
Private Const SPI_GETKEYBOARDPREF = &H44
Private Const SPI_GETKEYBOARDSPEED = &HA
Private Const SPI_GETLISTBOXSMOOTHSCROLLING = &H1006
Private Const SPI_GETLOWPOWERACTIVE = &H53
Private Const SPI_GETLOWPOWERTIMEOUT = &H4F
Private Const SPI_GETMENUANIMATION = &H1002
Private Const SPI_GETMENUDROPALIGNMENT = &H1B
Private Const SPI_GETMENUUNDERLINES = &H100A
Private Const SPI_GETMINIMIZEDMETRICS = &H2B
Private Const SPI_GETMOUSE = &H3
Private Const SPI_GETMOUSEHOVERHEIGHT = &H64
Private Const SPI_GETMOUSEHOVERTIME = &H66
Private Const SPI_GETMOUSEHOVERWIDTH = &H62
Private Const SPI_GETMOUSEKEYS = &H36
Private Const SPI_GETMOUSESPEED = &H70
Private Const SPI_GETMOUSETRAILS = &H5E
Private Const SPI_GETNONCLIENTMETRICS = &H29
Private Const SPI_GETPOWEROFFACTIVE = &H54
Private Const SPI_GETPOWEROFFTIMEOUT = &H50
Private Const SPI_GETSCREENREADER = &H46
Private Const SPI_GETSCREENSAVEACTIVE = &H10
Private Const SPI_GETSCREENSAVERRUNNING = &H72
Private Const SPI_GETSCREENSAVETIMEOUT = &HE
Private Const SPI_GETSERIALKEYS = &H3E
Private Const SPI_GETSHOWIMEUI = &H6E
Private Const SPI_GETSHOWSOUNDS = &H38
Private Const SPI_GETSOUNDSENTRY = &H40
Private Const SPI_GETSTICKYKEYS = &H3A
Private Const SPI_GETTOGGLEKEYS = &H34
Private Const SPI_GETWHEELSCROLLLINES = &H68
Private Const SPI_GETWINDOWSEXTENSION = &H5C
Private Const SPI_GETWORKAREA = &H30
Private Const SPI_ICONHORIZONTALSPACING = &HD
Private Const SPI_ICONVERTICALSPACING = &H18
Private Const SPI_LANGDRIVER = &HC
Private Const SPI_SCREENSAVERRUNNING = &H61
Private Const SPI_SETACCESSTIMEOUT = &H3D
Private Const SPI_SETACTIVEWINDOWTRACKING = &H1001
Private Const SPI_SETACTIVEWNDTRKTIMEOUT = &H2003
Private Const SPI_SETACTIVEWNDTRKZORDER = &H100D
Private Const SPI_SETANIMATION = &H49
Private Const SPI_SETBEEP = &H2
Private Const SPI_SETBORDER = &H6
Private Const SPI_SETCOMBOBOXANIMATION = &H1005
Private Const SPI_SETCURSORS = &H57
Private Const SPI_SETDEFAULTINPUTLANG = &H5A
Private Const SPI_SETDESKPATTERN = &H15
Private Const SPI_SETDESKWALLPAPER = &H14
Private Const SPI_SETDOUBLECLICKTIME = &H20
Private Const SPI_SETDOUBLECLKHEIGHT = &H1E
Private Const SPI_SETDOUBLECLKWIDTH = &H1D
Private Const SPI_SETDRAGFULLWINDOWS = &H25
Private Const SPI_SETDRAGHEIGHT = &H4D
Private Const SPI_SETDRAGWIDTH = &H4C
Private Const SPI_SETFASTTASKSWITCH = &H24
Private Const SPI_SETFILTERKEYS = &H33
Private Const SPI_SETFOREGROUNDFLASHCOUNT = &H2005
Private Const SPI_SETFOREGROUNDLOCKTIMEOUT = &H2001
Private Const SPI_SETGRADIENTCAPTIONS = &H1009
Private Const SPI_SETGRIDGRANULARITY = &H13
Private Const SPI_SETHANDHELD = &H4E
Private Const SPI_SETHIGHCONTRAST = &H43
Private Const SPI_SETHOTTRACKING = &H100F
Private Const SPI_SETICONMETRICS = &H2E
Private Const SPI_SETICONS = &H58
Private Const SPI_SETICONTITLELOGFONT = &H22
Private Const SPI_SETICONTITLEWRAP = &H1A
Private Const SPI_SETKEYBOARDDELAY = &H17
Private Const SPI_SETKEYBOARDPREF = &H45
Private Const SPI_SETKEYBOARDSPEED = &HB
Private Const SPI_SETLANGTOGGLE = &H5B
Private Const SPI_SETLISTBOXSMOOTHSCROLLING = &H1007
Private Const SPI_SETLOWPOWERACTIVE = &H55
Private Const SPI_SETLOWPOWERTIMEOUT = &H51
Private Const SPI_SETMENUANIMATION = &H1003
Private Const SPI_SETMENUDROPALIGNMENT = &H1C
Private Const SPI_SETMENUUNDERLINES = &H100B
Private Const SPI_SETMINIMIZEDMETRICS = &H2C
Private Const SPI_SETMOUSE = &H4
Private Const SPI_SETMOUSEBUTTONSWAP = &H21
Private Const SPI_SETMOUSEHOVERHEIGHT = &H65
Private Const SPI_SETMOUSEHOVERTIME = &H67
Private Const SPI_SETMOUSEHOVERWIDTH = &H63
Private Const SPI_SETMOUSEKEYS = &H37
Private Const SPI_SETMOUSESPEED = &H71
Private Const SPI_SETMOUSETRAILS = &H5D
Private Const SPI_SETNONCLIENTMETRICS = &H2A
Private Const SPI_SETPENWINDOWS = &H31
Private Const SPI_SETPOWEROFFACTIVE = &H56
Private Const SPI_SETPOWEROFFTIMEOUT = &H52
Private Const SPI_SETSCREENREADER = &H47
Private Const SPI_SETSCREENSAVEACTIVE = &H11
Private Const SPI_SETSCREENSAVERRUNNING = &H61
Private Const SPI_SETSCREENSAVETIMEOUT = &HF
Private Const SPI_SETSERIALKEYS = &H3F
Private Const SPI_SETSHOWIMEUI = &H6F
Private Const SPI_SETSHOWSOUNDS = &H39
Private Const SPI_SETSOUNDSENTRY = &H41
Private Const SPI_SETSTICKYKEYS = &H3B
Private Const SPI_SETTOGGLEKEYS = &H35
Private Const SPI_SETWHEELSCROLLLINES = &H69
Private Const SPI_SETWORKAREA = &H2F

Private Const SPI_SETFONTSMOOTHING = &H4B '(2000)
Private Const SPI_GETFONTSMOOTHING = &H4A '(2000)
Private Const SPI_GETFONTSMOOTHINGTYPE = 8202 '(XP)
Private Const SPI_SETFONTSMOOTHINGTYPE = 8203 '(XP)
Private Const SPI_GETFONTSMOOTHINGCONTRAST = 8204 '(XP)
Private Const SPI_SETFONTSMOOTHINGCONTRAST = 8205 '(XP)
Private Const SPI_GETCLEARTYPE = 4168 '(VISTA)
Private Const SPI_SETCLEARTYPE = 4169 '(VISTA)

Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDCHANGE = SPIF_SENDWININICHANGE

Private Const FE_FONTSMOOTHINGORIENTATIONBGR = 0
Private Const FE_FONTSMOOTHINGORIENTATIONRGB = 1
Private Const FE_FONTSMOOTHINGSTANDARD = 1
Private Const FE_FONTSMOOTHINGCLEARTYPE = 2
Private Const FE_FONTSMOOTHINGDOCKING = 32768

'lfCharSet Konstanten
Private Const ANSI_CHARSET = 0          'Ansi Zeichensatz
Private Const ARABIC_CHARSET = 178      'Arabisch (NT/2000)
Private Const BALTIC_CHARSET = 186      'Baltisch (Win 9x)
Private Const CHINESEBIG5_CHARSET = 136 'Chinesisch
Private Const DEFAULT_CHARSET = 1       'Standard
Private Const EASTEUROPE_CHARSET = 238  'Osteuropäisch (Win 9x)
Private Const GB2312_CHARSET = 134      'Englisch
Private Const GREEK_CHARSET = 161       'Griechisch (Win 9x)
Private Const HANGEUL_CHARSET = 129     'Handgeul
Private Const HEBREW_CHARSET = 177      'Hebräisch (NT/2000)
Private Const JOHAB_CHARSET = 130       'Johab (Win 9x)
Private Const MAC_CHARSET = 77          'Mac (Win 9x)
Private Const OEM_CHARSET = 255         'OEM
Private Const RUSSIAN_CHARSET = 204     'Russisch (Win 9x)
Private Const SHIFTJIS_CHARSET = 128    'ShiftJis
Private Const SYMBOL_CHARSET = 2        'Symbolisch
Private Const THAI_CHARSET = 222        'Thailändisch (NT/2000)
Private Const TURKISH_CHARSET = 162     'Türkisch (Win 9x)

Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Const SECURITY_BUILTIN_DOMAIN_RID As Long = &H20
Private Const DOMAIN_ALIAS_RID_ADMINS As Long = &H220

Const GRADIENT_FILL_RECT_H = &H0
Const GRADIENT_FILL_RECT_V = &H1
Const GRADIENT_FILL_TRIANGLE = &H2
Const GRADIENT_FILL_OP_FLAG = &HFF

Private Type LOGFONT2
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type ENUMLOGFONTEX
    elfLogFont As LOGFONT2
    elfFullName(LF_FULLFACESIZE) As Byte
    elfStyle(LF_FACESIZE) As Byte
    elfScript(LF_FACESIZE) As Byte
End Type

Private Type TRIVERTEX
    PosiX As Long
    PosiY As Long
    FaRed As Integer
    FaGre As Integer
    FaBlu As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Private FnLet As String
Private FntNr As Integer

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)
Private Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
Private Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Any, lpSource As Any, ByVal length As Long)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsUserAnAdmin Lib "shell32" () As Long
Private Declare Function DragQueryPoint Lib "shell32.dll" (ByVal HDROP As Long, lpPoint As POINTAPI) As Long
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function GetVersionEx1 Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetVersionEx2 Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function SendNotifyMessage Lib "user32.dll" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBmp As Long, ByVal nStartScan As Long, ByVal cScanLines As Long, lpvBits As Any, lpbm As BITMAPINFO, ByVal fuColorUse As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetMapMode Lib "gdi32.dll" (ByVal wHdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal wHdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fown As Long, lplpvObj As IPicture) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SystemParametersInfoA Lib "user32.dll" (ByVal uAction As Long, ByVal uiParam As Long, ByRef pvParam As Long, ByVal fuWinIni As Long) As Long 'ByRef pvParam As Long
Private Declare Function EnumFontFamiliesEx Lib "gdi32.dll" Alias "EnumFontFamiliesExA" (ByVal hDC As Long, lpLogFont As LOGFONT2, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As Any, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function CheckTokenMembership Lib "advapi32.dll" (ByVal hToken As Long, ByVal pSidToCheck As Long, pbIsMember As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInfo As Any, ByVal TokenInfoLen As Long, ReturnLen As Long) As Long

Private clFil As clsFile

Private AkHun As Integer
Private AkTau As Integer
Private AkMil As Integer

Private Sub ArraySor(ByRef SoAry() As String, Optional ByVal ArSta As Long, Optional ByVal ArEnd As Long)
On Error Resume Next

Dim AktZa As Long
Dim AktWe As Long
Dim RnIdx As Long
Dim SoLow As Long
Dim SoHoh As Long
Dim Parti As String

SoLow = IIf(ArSta = 0, LBound(SoAry), ArSta)
SoHoh = IIf(ArEnd = 0, UBound(SoAry), ArEnd)

If SoLow < SoHoh Then
   If SoHoh - SoLow = 1 Then
      If UCase(SoAry(SoLow)) > UCase(SoAry(SoHoh)) Then
         ArraySwp SoAry(SoLow), SoAry(SoHoh)
      End If
   Else
      'Einen zufälligen Ausgangspunkt generieren
      RnIdx = Rnd() * (SoHoh - SoLow) + SoLow
      ArraySwp SoAry(SoHoh), SoAry(RnIdx)
      Parti = UCase(SoAry(SoHoh))
      Do
         'Von beiden Seiten auf den Ausgangspunkt "zugehen"
         AktZa = SoLow: AktWe = SoHoh
         Do While (AktZa < AktWe) And (UCase(SoAry(AktZa)) <= Parti)
            AktZa = AktZa + 1
         Loop
         Do While (AktWe > AktZa) And (UCase(SoAry(AktWe)) >= Parti)
            AktWe = AktWe - 1
         Loop

         'Wenn der Ausgangspunkt noch nicht erreicht ist, sind 2 Elemente auf
         'beiden Seiten funktionsunfähig, deswegen werden sie vertauscht
         If AktZa < AktWe Then
            ArraySwp SoAry(AktZa), SoAry(AktWe)
         End If
      Loop While AktZa < AktWe

      'Den Ausgangspunkt zu seinem richtigen Platz im Array führen
      ArraySwp SoAry(AktZa), SoAry(SoHoh)

      'Die ArraySor-Routine rekursiv nochmals aufrufen
      If (AktZa - SoLow) < (SoHoh - AktZa) Then
         ArraySor SoAry, SoLow, AktZa - 1
         ArraySor SoAry, AktZa + 1, SoHoh
      Else
         ArraySor SoAry, AktZa + 1, SoHoh
         ArraySor SoAry, SoLow, AktZa - 1
      End If
   End If
End If

End Sub
Private Sub ArraySwp(ArErs As String, ArZwe As String)
On Error Resume Next
   
Dim TmpVa As String

TmpVa = ArErs

ArErs = ArZwe

ArZwe = TmpVa
   
End Sub
Private Function CCol(ByVal Col As Byte) As Integer

If Col > &H7F Then
     CCol = (Col * &H100&) - &H10000
Else
     CCol = Col * &H100&
End If

End Function
Public Sub FontLis(ByVal FohDC As Long)
On Error GoTo WiErr

Dim LogFi As LOGFONT2

ReDim Preserve FnAry(0)

LogFi.lfCharSet = DEFAULT_CHARSET
DoEvents

EnumFontFamiliesEx FohDC, LogFi, AddressOf FontSor, 0&, 0&
DoEvents

ArraySor FnAry
DoEvents

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " FontLis " & Err.Number
Resume Next

End Sub
Private Function FontSor(ByRef LpElf As ENUMLOGFONTEX, ByVal LpIdx As Long, ByVal FnTyp As Long, ByVal lParam As Long) As Long
On Error Resume Next
    
Dim FntNa As String
Dim RetWe As Boolean

RetWe = FontKon(LpElf.elfLogFont.lfFaceName, FntNa)

If Not FnLet = FntNa Then
    If InStr(1, FntNa, "@", 1) = 0 Then
        FntNr = FntNr + 1
        ReDim Preserve FnAry(FntNr)
        FnAry(FntNr) = FntNa
    End If
End If

FnLet = FntNa

FontSor = 1

End Function
Private Function FontKon(ByAry() As Byte, OuStr As String) As Boolean
On Error Resume Next

Dim AktZa As Long

OuStr = vbNullString

For AktZa = 0 To UBound(ByAry)
    If ByAry(AktZa) = 0 Then Exit For
    OuStr = OuStr & Chr(ByAry(AktZa))
Next AktZa

FontKon = True

End Function
Public Sub TimEnde(ByVal TimZa As Integer)
On Error Resume Next

Select Case TimZa
Case 1: KillTimer 0, TiHw1
        GlTi1 = False 'Timer Ein/Ausschalten
Case 2: KillTimer 0, TiHw2
        GlTi2 = False 'Timer Ein/Ausschalten
Case 3: KillTimer 0, TiHw3
        GlTi3 = False 'Timer Ein/Ausschalten
Case 4: KillTimer 0, TiHw4
        GlTi4 = False 'Timer Ein/Ausschalten
Case 5: KillTimer 0, TiHw5
        GlTi5 = False 'Timer Ein/Ausschalten
Case 6: KillTimer 0, TiHw6
        GlTi6 = False 'Timer Ein/Ausschalten
Case 7: KillTimer 0, TiHw7
        GlTi7 = False 'Timer Ein/Ausschalten
Case 8: KillTimer 0, TiHw8
        GlTi8 = False 'Timer Ein/Ausschalten
End Select

End Sub
Public Sub TimInit(ByVal TimZa As Integer, ByVal Intvl As Long)
On Error GoTo WiErr

Select Case TimZa
Case 1: TiHw1 = SetTimer(0, 0, Intvl * 1000, AddressOf STeAl)
        GlTi1 = True 'Timer Ein/Ausschalten
Case 2: TiHw2 = SetTimer(0, 0, Intvl * 1000, AddressOf AChip)
        GlTi2 = True 'Timer Ein/Ausschalten
Case 3: TiHw3 = SetTimer(0, 0, Intvl * 1000, AddressOf STeAk)
        GlTi3 = True 'Timer Ein/Ausschalten
Case 4: TiHw4 = SetTimer(0, 0, Intvl * 1000, AddressOf SSplS1) 'Splashscreen schließen
        GlTi4 = True 'Timer Ein/Ausschalten
Case 5: TiHw5 = SetTimer(0, 0, Intvl * 1000, AddressOf SSplS2) 'Splashscreen schließen
        GlTi5 = True 'Timer Ein/Ausschalten
Case 6: TiHw6 = SetTimer(0, 0, Intvl * 1000, AddressOf SBiAu)
        GlTi6 = True 'Timer Ein/Ausschalten
Case 7: TiHw7 = SetTimer(0, 0, Intvl * 1000, AddressOf SBiIa)
        GlTi7 = True 'Timer Ein/Ausschalten
Case 8: TiHw8 = SetTimer(0, 0, Intvl * 1000, AddressOf SInPaD) 'Grid-Fix
        GlTi8 = True 'Timer Ein/Ausschalten
End Select

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " TimInit " & Err.Number
Exit Sub

End Sub
Public Function WindowAdEle() As Integer
On Error GoTo WiErr

Dim hTok As Long
Dim TLen As Long
Dim TokenIsElevated As Long
Dim TEType As TOKEN_ELEVATION_TYPE

If OpenProcessToken(GetCurrentProcess, TOKEN_READ, hTok) = 0 Then
    WindowAdEle = 2 'Error: Couldn't open the process token
    Exit Function
End If

If GetTokenInformation(hTok, TOKENELEVATION, TokenIsElevated, Len(TokenIsElevated), TLen) = 0 Then
    WindowAdEle = 3 'Error: Couldn't retrieve the elevation right of the current process token
    CloseHandle hTok
    Exit Function
End If

If TokenIsElevated Then
    If GetTokenInformation(hTok, TOKENELEVATIONTYPE, TEType, Len(TEType), TLen) = 0 Then
        WindowAdEle = 4 'Error: Couldn't retrieve the elevation token class
    ElseIf TEType = TokenElevationTypeFull Or TEType = TokenElevationTypeDefault Then
        WindowAdEle = 1
    End If
End If

CloseHandle hTok

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowAdEle " & Err.Number
Exit Function

End Function
Public Function WindowAdmin() As Long
On Error GoTo WiErr

Dim RetWe As Long

RetWe = IsUserAnAdmin()

WindowAdmin = RetWe

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowAdmin " & Err.Number
Exit Function

End Function
Public Function WindowAdRig() As Boolean
On Error GoTo WiErr

Dim uAuthNt As SID_IDENTIFIER_AUTHORITY
Dim pSidAdmins As Long
Dim lResult As Long

uAuthNt.Value(5) = 5

If AllocateAndInitializeSid(uAuthNt, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, pSidAdmins) <> 0 Then
    If CheckTokenMembership(0, pSidAdmins, lResult) <> 0 Then
        WindowAdRig = (lResult <> 0)
    End If
    Call FreeSid(pSidAdmins)
End If

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowAdRig " & Err.Number
Exit Function

End Function
Public Function WindowArry(ByRef TmAry() As String, ByVal IdxZa As Integer, ByVal IdxWe As Long, ByVal TmStr As String) As Boolean
On Error GoTo WiErr

Dim AktZa As Integer
Dim GesZa As Integer
Dim Vorha As Boolean

GesZa = UBound(TmAry)

If GesZa > 0 Then
    For AktZa = 1 To GesZa
        If IdxWe > 0 Then
            If TmAry(AktZa, IdxZa) = IdxWe Then
                Vorha = True
                Exit For
            End If
        Else
            If TmAry(AktZa, IdxZa) Like TmStr Then
                Vorha = True
                Exit For
            End If
        End If
    Next AktZa
End If

WindowArry = Vorha

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowArry " & Err.Number
Exit Function

End Function
Public Function WindowBitm(ByVal Width As Long, ByVal Height As Long, ByVal ColorStart As Long, ByVal ColorEnd As Long, ByVal LeftToRight As Boolean) As StdPicture
On Error GoTo WiErr

Dim hwnDc As Long
Dim hwDib As Long
Dim TmpDC As Long
Dim hBmNe As Long
Dim hBmOl As Long
Dim AktZa As Long
Dim AktCo As Long
Dim CoSt1 As Double
Dim CoSt2 As Double
Dim CoSt3 As Double
Dim TmCo1 As Double
Dim TmCo2 As Double
Dim TmCo3 As Double
Dim TmPic As IPicture
Dim PicDe As PictDesc
Dim IdPic(3) As Long
Dim ColoS(3) As Byte
Dim ColoE(3) As Byte
Dim Bitma() As RGBQUAD
Dim BmInf As BITMAPINFO
        
hwnDc = GetDC(GetDesktopWindow)
TmpDC = CreateCompatibleDC(hwnDc)
hBmNe = CreateCompatibleBitmap(hwnDc, Width, Height)
hBmOl = SelectObject(TmpDC, hBmNe)
 
CopyMemory1 ColoS(0), ColorStart, 4
CopyMemory1 ColoE(0), ColorEnd, 4
CoSt1 = WindowStep(ColoS(0), ColoE(0), Width)
CoSt2 = WindowStep(ColoS(1), ColoE(1), Width)
CoSt3 = WindowStep(ColoS(2), ColoE(2), Width)
TmCo1 = ColoS(0)
TmCo2 = ColoS(1)
TmCo3 = ColoS(2)
   
If Not LeftToRight Then
    ReDim Bitma(Height - 1, Width - 1)
    For AktZa = 0 To Width - 1
        TmCo1 = TmCo1 + CoSt1
        TmCo2 = TmCo2 + CoSt2
        TmCo3 = TmCo3 + CoSt3
        For AktCo = 0 To Height - 1
            Bitma(AktCo, AktZa).rgbRed = TmCo1
            Bitma(AktCo, AktZa).rgbGreen = TmCo2
            Bitma(AktCo, AktZa).rgbBlue = TmCo3
        Next AktCo
    Next AktZa
Else
    ReDim Bitma(Width - 1, Height - 1)
    For AktZa = 0 To Width - 1
        TmCo1 = TmCo1 + CoSt1
        TmCo2 = TmCo2 + CoSt2
        TmCo3 = TmCo3 + CoSt3
        For AktCo = 0 To Height - 1
            Bitma(AktZa, AktCo).rgbRed = TmCo1
            Bitma(AktZa, AktCo).rgbGreen = TmCo2
            Bitma(AktZa, AktCo).rgbBlue = TmCo3
        Next AktCo
    Next AktZa
End If
   
With BmInf.bmiHeader
    .biSize = Len(BmInf.bmiHeader)
    .biWidth = Width
    .biHeight = Height
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = Width * Height * 4
End With
SetDIBits TmpDC, hBmNe, 0, Height, Bitma(0, 0), BmInf, 0
   
Call DeleteObject(hwDib)
Call SelectObject(TmpDC, hBmOl)
DeleteDC TmpDC
ReleaseDC GetDesktopWindow, hwnDc
   
IdPic(0) = &H7BF80980
IdPic(1) = &H101ABF32
IdPic(2) = &HAA00BB8B
IdPic(3) = &HAB0C3000
   
With PicDe
    .cbSizeofStruct = Len(PicDe)
    .hImage = hBmNe
    .picType = 1
End With

If OleCreatePictureIndirect(PicDe, IdPic(0), Abs(True), TmPic) = 0 Then Set WindowBitm = TmPic

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowBitm " & Err.Number
Resume Next

End Function
Public Sub WindowClas1(mHwnd As Long)
On Error GoTo WiErr

MProc = SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf WindowMroc)
    
Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowClas1 " & Err.Number
Resume Next
        
End Sub
Public Sub WindowClas2(mHwnd As Long)
On Error GoTo WiErr

MProc = SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf WindowProc2)
    
Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowClas2 " & Err.Number
Resume Next
        
End Sub
Public Function WindowClear(ByVal WeSet As Boolean) As Long
On Error GoTo WiErr
'ClearType Auslesen / Setzen

Dim RetWe As Long
Dim TesWe As Long

If WeSet = True Then
    RetWe = SystemParametersInfoA(SPI_SETFONTSMOOTHING, FE_FONTSMOOTHINGSTANDARD, 0&, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
    RetWe = SystemParametersInfoA(SPI_SETFONTSMOOTHINGTYPE, 0&, FE_FONTSMOOTHINGCLEARTYPE, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
    RetWe = SystemParametersInfoA(SPI_SETCLEARTYPE, 0&, FE_FONTSMOOTHINGCLEARTYPE, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
Else
    RetWe = SystemParametersInfoA(SPI_GETFONTSMOOTHINGTYPE, 0&, TesWe, 0&)
End If

WindowClear = TesWe

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowClear " & Err.Number
Resume Next
        
End Function
Public Sub WindowClose()
On Error GoTo WiErr
'Programm beenden

Dim RetWe As Long

Set FM = frmMain

RetWe = SendNotifyMessage(FM.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0)

Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowClose " & Err.Number
Resume Next

End Sub
Sub WindowCmb(ByVal CmbIt As VB.ComboBox, ByVal DroHo As Long)
On Error GoTo WiErr

Dim DroLa As Long

DroLa = Screen.Height / Screen.TwipsPerPixelY
If DroHo > DroLa / 2 Then DroHo = DroLa / 2

With CmbIt
    MoveWindow .hwnd, .Left / Screen.TwipsPerPixelX, .Top / Screen.TwipsPerPixelY, .Width / Screen.TwipsPerPixelY, DroHo, 1
End With
    
Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowCmb " & Err.Number
Resume Next

End Sub
Sub WindowCmL(ByVal CmbIt As VB.ComboBox)
On Error GoTo WiErr

PostMessage CmbIt.hwnd, CB_SHOWDROPDOWN, 0, 0

Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowCmL " & Err.Number
Resume Next

End Sub
Sub WindowCmX(ByVal CmbIt As XtremeSuiteControls.ComboBox, ByVal DroHo As Long)
On Error GoTo WiErr

Dim DroLa As Long

DroLa = Screen.Height / Screen.TwipsPerPixelY
If DroHo > DroLa / 2 Then DroHo = DroLa / 2

With CmbIt
    MoveWindow .hwnd, .Left / Screen.TwipsPerPixelX, .Top / Screen.TwipsPerPixelY, .Width / Screen.TwipsPerPixelY, DroHo, 1
End With
    
Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowCmX " & Err.Number
Resume Next

End Sub
Public Sub WindowEnSu(mHwnd As Long)
On Error Resume Next

Dim RetWe As Long

RetWe = SetWindowLong(mHwnd, GWL_WNDPROC, MProc)
    
End Sub


Public Sub WindowEml(ByVal EmEmp As String, Optional ByVal EmBet As String, Optional ByVal EmTex As String)
On Error GoTo WiErr

Dim RetWe As Long
Dim EmBuf As String

If EmBet = vbNullString Then EmBet = Chr$(32)
If EmTex = vbNullString Then EmTex = Chr$(32)

If InStr(1, EmEmp, "bcc", vbTextCompare) > 1 Then
    EmBuf = EmEmp & "Subject=" & EmBet & "&Body=" & EmTex
Else
    EmBuf = EmEmp & "?Subject=" & EmBet & "&Body=" & EmTex
End If

RetWe = ShellExecute(0&, "open", EmBuf, vbNullString, vbNullString, SW_SHOW)

Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowEml " & Err.Number
Resume Next

End Sub

Public Function WindowFont() As Boolean
On Error GoTo WiErr

Dim wHdc As Long
Dim hwnd As Long
Dim MpMod As Long
Dim TxMat As TEXTMETRIC

hwnd = GetDesktopWindow()

wHdc = GetWindowDC(hwnd)

If wHdc Then
    MpMod = SetMapMode(wHdc, MM_TEXT)
    GetTextMetrics wHdc, TxMat
    MpMod = SetMapMode(wHdc, MpMod)
    ReleaseDC hwnd, wHdc
    If (TxMat.tmHeight > 16) Then
        WindowFont = False
    Else
        WindowFont = True
    End If
End If

Exit Function

WiErr:
If GlDbg = True Then SErLog Err.Description & " WindowFont " & Err.Number
Resume Next
   
End Function
Public Function WindowIcon(FiNam As String, Optional ByVal Flag As Boolean = False, Optional IcoGr As Long) As Long
On Error GoTo WiErr

If Flag = True Then
    WindowIcon = LoadImage(App.hInstance, FiNam, IMAGE_ICON, IcoGr, IcoGr, LR_LOADFROMFILE)
Else
    WindowIcon = LoadImage(App.hInstance, FiNam, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
End If

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowIcon " & Err.Number
Resume Next

End Function
Public Sub WindowInfo(ByVal FM As Form, ByVal AnwNam As String)

Dim RetWe As Long

RetWe = ShellAbout(FM.hwnd, AnwNam, "© 1997-2007 SimpliMed Praxissoftware" & vbCrLf & "(Tel: 0700 / 8888 4422) - Version " & App.Major & "." & App.Minor & "." & App.Revision, FM.Icon)

End Sub
Public Sub WindowIni1(FeHwn As Long, FeSta As Long, FoSiz As SIZEPAR)
On Error GoTo WiErr

SizPa = FoSiz

WProc = SetWindowLong(FeHwn, GWL_WNDPROC, AddressOf WindowProc1)

Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowInit1 " & Err.Number
Resume Next

End Sub
Public Function WindowLoad(ByVal FoNam As String) As Boolean
On Error Resume Next

WindowLoad = False

For Each FS In Forms
    If LCase$(FS.Name) = LCase$(FoNam) Then
        WindowLoad = True
        Exit For
    End If
Next FS

End Function
Public Function WindowMess(ByVal MeStr As String, ByVal Dialo As Long, ByVal TiStr As String, ByVal hwnd As Long) As Long
On Error GoTo WiErr

Dim RetWe As Long

RetWe = MessageBox(hwnd, MeStr, TiStr, Dialo)

WindowMess = RetWe

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowMess " & Err.Number
Resume Next
  
End Function
Private Function WindowMroc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo WiErr

Select Case Msg
Case WM_MOUSEWHEEL:
    GlMou = True
Case WM_LBUTTONDOWN:
    GlMou = False
End Select

WindowMroc = CallWindowProc(MProc, hwnd, Msg, wParam, lParam)
        
Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowMroc " & Err.Number
Resume Next
        
End Function
Public Sub WindowMut()
On Error GoTo WiErr

Dim RetWe As Long

RetWe = CreateMutex(0&, 0&, "SimpliMed23")

Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowMut " & Err.Number
Resume Next

End Sub
Public Sub WindowPicSav(obCon As Object, ByVal FiNam As String)
On Error GoTo WiErr

'Inhalt einer Form/PictureBox als Bild speichern
Dim PiBi6 As VB.PictureBox
Dim FrmOldScaleMode As Integer
Dim PicOldScaleMode As Integer
Dim OldAutoRedraw As Boolean

Set FM = frmMain
Set PiBi6 = FM.picBild6

'Scale-Mode auf Pixel setzen
FrmOldScaleMode = FM.ScaleMode
FM.ScaleMode = vbPixels

' Eigenschaften der 2.PictureBox, die als Zwischenspeicher dient
With PiBi6
    .BorderStyle = 0
    .AutoRedraw = True
    .ScaleMode = vbPixels
    
    '2.PictureBox über das Container-Objekt legen
    With obCon
        PicOldScaleMode = .ScaleMode
        .ScaleMode = vbPixels
        PiBi6.Move 0, 0, .ScaleWidth, .ScaleHeight
        FM.ScaleMode = FrmOldScaleMode
        OldAutoRedraw = .AutoRedraw
    
        'Inhalt des Containers in die 2. PictureBox kopieren
        .AutoRedraw = False
        BitBlt PiBi6.hDC, 0, 0, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, vbSrcCopy
        .AutoRedraw = OldAutoRedraw
        .ScaleMode = PicOldScaleMode
    End With
    
    'Inhalt der 2. PictureBox als Bitmap speichern
    SavePicture .Image, FiNam
    .Cls
    .AutoRedraw = False
End With

Exit Sub

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowPicSav " & Err.Number
Resume Next

End Sub
Private Function WindowProc1(ByVal mHwnd As Long, ByVal FoMsg As Long, ByVal wParm As Long, ByVal lParm As Long, ByVal WiSta As Long) As Long
On Error GoTo WiErr
    
Dim RetWe As Long
Dim MaxIn As MINMAXINFO

If FoMsg = WM_GETMINMAXINFO Then
    CopyMemory1 MaxIn, lParm, Len(MaxIn)
    
    MaxIn.ptMaxPosition.x = 0
    MaxIn.ptMaxPosition.y = 0
    MaxIn.ptMaxSize.x = Screen.Width / Screen.TwipsPerPixelX
    MaxIn.ptMaxSize.y = Screen.Height / Screen.TwipsPerPixelY
    
    MaxIn.ptMinTrackSize.x = SizPa.xMin
    MaxIn.ptMinTrackSize.y = SizPa.yMin
    MaxIn.ptMaxTrackSize.x = SizPa.xMax
    MaxIn.ptMaxTrackSize.y = SizPa.yMax
    
    CopyMemory2 lParm, MaxIn, Len(MaxIn)
    RetWe = DefWindowProc(mHwnd, FoMsg, wParm, lParm)
Else
    RetWe = CallWindowProc(WProc, mHwnd, FoMsg, wParm, lParm)
End If

WindowProc1 = RetWe
    
Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowProc1 " & Err.Number
Resume Next
    
End Function
Private Function WindowProc2(ByVal mHwnd As Long, ByVal HwMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

If HwMsg = WM_DESTROY Then
    Call WindowEnSu(mHwnd)
ElseIf HwMsg = WM_DISPLAYCHANGE Then
    If GlTi8 = False Then TimInit 8, 2
ElseIf HwMsg = WM_NCLBUTTONDOWN Then
    'Call frmMain.FLeMo
ElseIf HwMsg = WM_NCMOUSEMOVE Then
    'Call frmMain.FMoUp(Trim(wParam))
End If

WindowProc2 = CallWindowProc(MProc, mHwnd, HwMsg, wParam, lParam)
    
End Function

Public Function WindowRGB(Farbe As Long) As tRGB
On Error GoTo WiErr

WindowRGB.rot = (Farbe And &HFF&)
WindowRGB.grün = (Farbe And &HFF00&) \ 256
WindowRGB.blau = (Farbe And &HFF0000) \ 65536

Exit Function

WiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "WindowRGB " & Err.Number
Resume Next

End Function
Public Function WinRound(ByVal EinBet As Double, ByVal KomSt As Integer) As Double
On Error Resume Next

Dim TenPower As Variant

TenPower = CDec(10 ^ KomSt)

WinRound = CDbl(Sgn(EinBet) * Int(CDec(0.5) + Abs(EinBet) * TenPower) / TenPower)

End Function
Public Sub WinSound(ByVal SoIdx As Integer)
On Error Resume Next
'Spielt einen Systemsound

Dim SoStr As String
Dim RetWe As Long
  
Select Case SoIdx
Case 0: SoStr = "SystemQuestion"
Case 1: SoStr = "SystemExclaimation"
Case 2: SoStr = "SystemHand"
Case 3: SoStr = "Maximize"
Case 4: SoStr = "MenuCommand"
Case 5: SoStr = "MenuPopup"
Case 6: SoStr = "Minimize"
Case 7: SoStr = "MailBeep"
Case 8: SoStr = "Open"
Case 9: SoStr = "Close"
Case 10: SoStr = "AppGPFault"
Case 11: SoStr = ".Default"
Case 12: SoStr = "SystemAsterisk"
Case 13: SoStr = "RestoreUp"
Case 14: SoStr = "RestoreDown"
Case 15: SoStr = "SystemExit"
Case 16: SoStr = "SystemStart"
End Select
  
RetWe = PlaySound(SoStr, 0&, SND_ALIAS Or SND_ASYNC Or SND_NODEFAULT)

End Sub
Public Sub WinUnZip(ByVal ZipNa As String, ByVal OrdNa As String)
On Error Resume Next

Dim sh As Object
Dim fSource As Object
Dim fTarget As Object

Set sh = CreateObject("Shell.Application")

Set fSource = sh.Namespace((ZipNa))
Set fTarget = sh.Namespace((OrdNa))

fTarget.CopyHere fSource.Items

End Sub

Private Function WindowRun(ByVal PID As Long) As Boolean
On Error Resume Next

Dim hProcess As Long

hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, PID)
WindowRun = hProcess <> 0
If WindowRun Then CloseHandle hProcess

End Function
Public Sub WindowSleep(ByVal SecMi As Long)
On Error Resume Next

Sleep SecMi

End Sub

Public Function WindowStart(ByVal FiNam As String, Optional ByVal WindowStyle As VBA.VbAppWinStyle = vbMaximizedFocus, Optional ByVal Warten As Boolean = False, Optional ByVal Synchron As Boolean = True) As Double
On Error Resume Next

Dim ProcInfo As PROCESS_INFORMATION
Dim StartInfo As STARTUPINFO

With StartInfo
    .cb = Len(StartInfo)
    .dwFlags = STARTF_USESHOWWINDOW
    .wShowWindow = WindowStyle
End With

If CreateProcessA(0&, FiNam, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, StartInfo, ProcInfo) = 0 Then Err.Raise vbObjectError + 1, "modShellWait", "Nicht gefunden " & FiNam

If Warten = True Then
    If Synchron = True Then
        WaitForSingleObject ProcInfo.hProcess, INFINITE
    Else
        While WaitForSingleObject(ProcInfo.hProcess, 0) = WAIT_TIMEOUT
            DoEvents
        Wend
    End If
Else

WindowStart = ProcInfo.dwProcessId

End If

CloseHandle ProcInfo.hProcess
CloseHandle ProcInfo.hThread
  
End Function
Private Function WindowStep(ByVal ColorStart As Long, ByVal ColorEnd As Long, ByVal Steps As Long) As Double
On Error Resume Next

If ColorStart > ColorEnd Then
    WindowStep = ColorStart - ColorEnd
    If WindowStep <> 0 Then
        WindowStep = -(WindowStep / Steps)
    End If
Else
    WindowStep = ColorEnd - ColorStart
    If WindowStep <> 0 Then
        WindowStep = WindowStep / Steps
    End If
End If

End Function
Public Sub WindowUnHok(mHwnd As Long)
On Error Resume Next

Dim RetWe As Long

RetWe = SetWindowLong(mHwnd, GWL_WNDPROC, WProc)

End Sub
Public Sub WindowVerl()
'Erzeugt einen Farbverlauf auf einer PictureBox oder einer Form
On Error Resume Next

Dim PiBi7 As VB.PictureBox

Dim FoRct As GRADIENT_RECT
Dim FoVtx(0 To 1) As TRIVERTEX

'Set PiBi7 = frmMain.picBild7

With FoVtx(0)
    .PosiX = 0
    .PosiY = 0
    .FaRed = CCol(255)
    .FaGre = CCol(255)
    .FaBlu = CCol(255)
    .Alpha = 0
End With

With FoVtx(1)
    .PosiX = PiBi7.ScaleWidth
    .PosiY = PiBi7.ScaleHeight
    .FaRed = CCol(255)
    .FaGre = 0
    .FaBlu = CCol(255)
    .Alpha = 0
End With

FoRct.UpperLeft = 0
FoRct.LowerRight = 1

GradientFillRect PiBi7.hDC, FoVtx(0), 2, FoRct, 1, GRADIENT_FILL_RECT_V

End Sub
Public Function WindowVers() As String
'Gibt die Windows Version zurück
On Error GoTo WiErr
    
Dim OsVer As OSVERSIONINFOEX
Dim OsInf As OSVERSIONINFO

If mVorh Then
    WindowVers = mVers
    Exit Function
End If

OsInf.dwOSVersionInfoSize = Len(OsInf)

If GetVersionEx1(OsInf) = 0 Then
    mVers = "WIN_32s"
    Exit Function
End If
    
With OsInf
    Select Case .dwPlatformId
    Case VER_PLATFORM_WIN32s
        mVers = "WIN_32s"
    Case VER_PLATFORM_WIN32_WINDOWS
        Select Case .dwMinorVersion
            Case 0: mVers = "WIN_95"
            Case 10: mVers = "WIN_98"
            Case 90: mVers = "WIN_ME"
        End Select
    Case VER_PLATFORM_WIN32_NT
        Select Case .dwMajorVersion
            Case 3: mVers = "WIN_NT_3x"
            Case 4: mVers = "WIN_NT_4x"
            Case 5:
                Select Case .dwMinorVersion
                Case 0: mVers = "WIN_2K"
                Case 1:
                    OsVer.dwOSVersionInfoSize = Len(OsVer)
                    If GetVersionEx2(OsVer) = 0 Then
                        mVers = "WIN_XP_PROF"
                        Exit Function
                    End If
                    If (OsVer.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL Then
                        mVers = "WIN_XP_PROF_HOME"
                    Else
                        mVers = "WIN_XP_PROF"
                    End If
                Case 2: mVers = "WIN_2003"
                End Select
        End Select
    End Select
End With

WindowVers = mVers
mVorh = True

Exit Function

WiErr:
If GlDbg = True Then SErLog Err.Description & " WindowVers " & Err.Number
Resume Next

End Function
Public Function WinMax(a As Long, b As Long) As Long
    WinMax = IIf(a > b, a, b)
End Function
Public Function WinMin(a As Long, b As Long) As Long
    WinMin = IIf(a < b, a, b)
End Function
Private Function WinZaEi(ZaEin) As String
On Error Resume Next

Select Case ZaEin
Case 1: WinZaEi = "ein"
Case 2: WinZaEi = "zwei"
Case 3: WinZaEi = "drei"
Case 4: WinZaEi = "vier"
Case 5: WinZaEi = "fünf"
Case 6: WinZaEi = "sechs"
Case 7: WinZaEi = "sieben"
Case 8: WinZaEi = "acht"
Case 9: WinZaEi = "neun"
End Select

End Function
Public Function WinZaUm(AkZal As Long) As String
    
Dim TmTau As String
Dim TmMil As String
Dim TmHun As String

Select Case AkZal
Case Is > 999999999
    WinZaUm = "<Zahl ist zu groß>"
Case Is > 1000000
    AkMil = CInt(Left(CStr(AkZal), Len(CStr(AkZal)) - 6))
    TmTau = Left(CStr(AkZal), Len(CStr(AkZal)) - 3)
    AkTau = CInt(Right(TmTau, 3))
    TmTau = ""
    AkHun = CInt(Right(CStr(AkZal), 3))
    
    'Zehnmillionen ermitteln
    TmMil = WinZaZe(CInt(Right(CStr(AkMil), 2)))
    If TmMil = "eins" Then 'Sonderfall "hundertEIN tausen"
        TmMil = "einemillion"
    Else
        TmMil = TmMil + "millionen"
    End If
    'Hundertmillionen ermitteln
    If AkMil > 99 Then
        TmMil = WinZaEi(CInt(Left(CStr(AkMil), 1))) + "hundert" + TmMil
    End If
    
    'Zehntausender ermitteln
    TmTau = WinZaZe(CInt(Right(CStr(AkTau), 2)))
    If TmTau = "eins" Then 'Sonderfall "hundertEIN tausen"
        TmTau = "eintausend"
    Else
        TmTau = TmTau + "tausend"
    End If
    'Hunderttausender ermitteln
    If AkTau > 99 Then
        TmTau = WinZaEi(CInt(Left(CStr(AkTau), 1))) + _
         "hundert" + TmTau
    End If
    'Zehner ermitteln
    TmHun = TmHun + _
     WinZaZe(CInt(Right(CStr(AkHun), 2)))
    'Hunderter ermitteln
    If AkHun > 99 Then
        TmHun = WinZaEi(CInt(Left(CStr(AkHun), 1))) + "hundert" + TmHun
    End If
    'Zusammensetzen
    WinZaUm = TmMil + TmTau + TmHun
    
Case Is > 1000
    AkTau = CInt(Left(CStr(AkZal), Len(CStr(AkZal)) - 3))
    AkHun = CInt(Right(CStr(AkZal), 3))
    
    'Zehntausender ermitteln
    TmTau = WinZaZe(CInt(Right(CStr(AkTau), 2)))
    If TmTau = "eins" Then 'Sonderfall "hundertEIN tausen"
        TmTau = Left(TmTau, Len(TmTau) - 1) + "tausend"
    Else
        TmTau = TmTau + "tausend"
    End If
    'Hunderttausender ermitteln
    If AkTau > 99 Then
        TmTau = WinZaEi(CInt(Left(CStr(AkTau), 1))) + "hundert" + TmTau
    End If
    'Zehner ermitteln
    TmHun = TmHun + _
     WinZaZe(CInt(Right(CStr(AkHun), 2)))
    'Hunderter ermitteln
    If AkHun > 99 Then
        TmHun = WinZaEi(CInt(Left(CStr(AkHun), 1))) + "hundert" + TmHun
    End If
    'Zusammensetzen
    WinZaUm = TmTau + TmHun
    
Case Is > 0
    'Zehner ermitteln
    WinZaUm = WinZaZe(CInt(Right(CStr(AkZal), 2)))
    'Hunderter ermitteln
    If AkZal > 99 Then
        WinZaUm = WinZaEi(CInt(Left(CStr(AkZal), 1))) + "hundert" + WinZaUm
    End If
End Select

End Function

Private Function WinZaZe(ZaZen) As String
On Error Resume Next

Dim TmEin As String

Select Case ZaZen
Case Is < 20 'bis "zwanzig" Sonderfälle!

    Select Case ZaZen
    Case 1: WinZaZe = "eins"
    Case 2: WinZaZe = "zwei"
    Case 3: WinZaZe = "drei"
    Case 4: WinZaZe = "vier"
    Case 5: WinZaZe = "fünf"
    Case 6: WinZaZe = "sechs"
    Case 7: WinZaZe = "sieben"
    Case 8: WinZaZe = "acht"
    Case 9: WinZaZe = "neun"
    Case 10: WinZaZe = "zehn"
    Case 11: WinZaZe = "elf"
    Case 12: WinZaZe = "zwölf"
    Case 13: WinZaZe = "dreizehn"
    Case 14: WinZaZe = "vierzehn"
    Case 15: WinZaZe = "fünfzehn"
    Case 16: WinZaZe = "sechzehn"
    Case 17: WinZaZe = "siebzehn"
    Case 18: WinZaZe = "achtzehn"
    Case 19: WinZaZe = "neunzehn"
    End Select
    
Case Else 'größer zwanzig nur zehner ermitteln und einer aus andrer function
    
    TmEin = WinZaEi(CInt(Right(CStr(ZaZen), 1)))
    
    If TmEin <> "" Then
        TmEin = TmEin + "und"
    End If
    
    Select Case (CInt(Left(CStr(ZaZen), 1)) * 10)
    Case 20: WinZaZe = TmEin + "zwanzig"
    Case 30: WinZaZe = TmEin + "dreißig"
    Case 40: WinZaZe = TmEin + "vierzig"
    Case 50: WinZaZe = TmEin + "fünfzig"
    Case 60: WinZaZe = TmEin + "sechzig"
    Case 70: WinZaZe = TmEin + "siebzig"
    Case 80: WinZaZe = TmEin + "achtzig"
    Case 90: WinZaZe = TmEin + "neunzig"
    End Select
    
End Select

End Function
Public Sub WinZip(ByVal ZipNa As String, ByVal FiNam As String)
On Error Resume Next

Dim ObShe As Object
Dim fSource As Object
Dim fTarget As Object
Dim iSource As Object
Dim ObItm As Object
Dim ObOrd As Object
Dim AktZa As Long

Set ObShe = CreateObject("Shell.Application")

Set fTarget = ObShe.Namespace((ZipNa))
If fTarget Is Nothing Then
    WinZpEr ZipNa
    Set fTarget = ObShe.Namespace((ZipNa))
End If

Dim OrdNa As String
Dim ZipDa As String

OrdNa = Left(FiNam, InStrRev(FiNam, "\"))
ZipDa = Mid(FiNam, InStrRev(FiNam, "\") + 1)

Set fSource = ObShe.Namespace((OrdNa))
For AktZa = 0 To fSource.Items.Count - 1
    If fSource.Items.Item((AktZa)).Name = ZipDa Then
        Set ObItm = fSource.Items.Item((AktZa))
        Exit For
    End If
Next AktZa

fTarget.CopyHere ObItm

End Sub
Private Sub WinZpEr(ByVal ZipNa As String)
On Error Resume Next

Dim fileNo As Integer
Dim ZIPFileEOCD(22) As Byte

ZIPFileEOCD(0) = Val("&H50")
ZIPFileEOCD(1) = Val("&H4b")
ZIPFileEOCD(2) = Val("&H05")
ZIPFileEOCD(3) = Val("&H06")

fileNo = FreeFile
Open ZipNa For Binary Access Write As #fileNo
Put #fileNo, , ZIPFileEOCD
Close #fileNo

End Sub

Public Function SZipp(ByVal ZipFil As String, ByVal SrcPfa As String, Optional ByVal DelSrc As Boolean = False, Optional ByVal UnZip As Boolean = False, Optional ByVal VerPa As String) As Boolean
On Error GoTo ErrHdl

Dim RetWe As Long
Dim ExPfa As String
Dim PaStr As String
Dim CmdSt As String

Dim PrInfo As PROCESS_INFORMATION
Dim StInfo As STARTUPINFO

SZipp = False

' Initialize file helper for existence checks
Set clFil = New clsFile

ExPfa = App.Path & "\SimpliZip.exe"

' Verify SimpliZip.exe exists in application directory
If clFil.FilVor(ExPfa) = False Then
    SPopu "Fehlende Programmdatei", "Das Programm SimpliZip.exe konnte nicht gefunden werden.", IC48_Information
    Set clFil = Nothing
    Exit Function
End If

' Strip trailing backslash for FindFirstFile compatibility
' (preserve root paths like "C:\" which need the backslash)
If Len(SrcPfa) > 3 Then
    If Right$(SrcPfa, 1) = "\" Then
        SrcPfa = Left$(SrcPfa, Len(SrcPfa) - 1)
    End If
End If

' Verify source file or folder exists
If clFil.FilVor(SrcPfa) = False Then
    SPopu "Fehlende Quelldatei", "Die Quelldateien konnten nicht gefunden werden.", IC48_Information
    Set clFil = Nothing
    Exit Function
End If

' Diagnostic logging
If GlLog = True Then
    SLogi "SZipp: SimpliZip gefunden: " & ExPfa
    SLogi "  ZIP Datei: " & ZipFil
    SLogi "  Quelle: " & SrcPfa
    SLogi "  Quelle loeschen: " & CStr(DelSrc)
End If

' Build command line parameters: "zip-file" "source-path" [-delete] [-unzip]
' Note: -delete must NOT be in quotes
PaStr = """" & ZipFil & """ """ & SrcPfa & """"

If DelSrc = True Then
    PaStr = PaStr & " -delete"
End If

If UnZip = True Then
    PaStr = PaStr & " -unzip"
End If

If VerPa <> vbNullString Then
    PaStr = PaStr & " -pw " & VerPa
End If

If GlLog = True Then SLogi "  Parameter: " & PaStr

' Full command line for CreateProcessA
CmdSt = """" & ExPfa & """ " & PaStr

' Configure startup: hidden window
With StInfo
    .cb = Len(StInfo)
    .dwFlags = STARTF_USESHOWWINDOW
    .wShowWindow = SW_HIDE
End With

' Start SimpliZip.exe via CreateProcessA
If CreateProcessA(0&, CmdSt, 0&, 0&, 0&, NORMAL_PRIORITY_CLASS, 0&, 0&, StInfo, PrInfo) = 0 Then
    If GlLog = True Then SLogi "SZipp: CreateProcess fehlgeschlagen"
    Set clFil = Nothing
    Exit Function
End If

If GlLog = True Then SLogi "SZipp: Prozess gestartet, warte auf Beendigung..."

' Wait for process completion (non-blocking UI via DoEvents)
Do While WaitForSingleObject(PrInfo.hProcess, 100) = WAIT_TIMEOUT
    DoEvents
Loop

' Get exit code
GetExitCodeProcess PrInfo.hProcess, RetWe

' Clean up process handles
CloseHandle PrInfo.hProcess
CloseHandle PrInfo.hThread

If RetWe = 0 Then
    SZipp = True
    If GlLog = True Then SLogi "SZipp: Erfolgreich beendet (ExitCode: 0)"
Else
    If GlLog = True Then SLogi "SZipp: Beendet mit ExitCode: " & RetWe
End If

Set clFil = Nothing
Exit Function

ErrHdl:
If GlLog = True Then SLogi "SZipp Fehler: " & Err.Number & " - " & Err.Description
SZipp = False
If Not PrInfo.hProcess = 0 Then CloseHandle PrInfo.hProcess
If Not PrInfo.hThread = 0 Then CloseHandle PrInfo.hThread
Set clFil = Nothing

End Function
