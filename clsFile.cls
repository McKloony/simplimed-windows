VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private mHwnd As Long   'Fensterhandle
Private mDrv As String  'Laufwerksbuchstabe mit \
Private mPfa As String  'Kompletter Pfad ohne Dateiname ohne \
Private mOrd As String  'Speicherordner ohne \
Private mNam As String  'Dateiname mit Extension
Private mNao As String  'Dateiname ohne Extension
Private mExt As String  'Dateiextension
Private mSta As String  'Startverzeichnis für Dateidialog
Private mStr As String  'Filterstring für Dateidialog
Private mTit As String  'Titel für Dialoge
Private mDat As String  'Dateninhalt für Schreiboperationen
Private mPar As String
Private mLoe As String  'Löschstring, immer kompletter Pfad + Dateiname, immer mit vbNullChar trennen + Ende mit vbNullChar
Private mCop As String  'Kopierstring, immer kompletter Pfad + Dateiname, immer mit vbNullChar trennen + Ende mit vbNullChar, Zielpfad mit ; trennen
Private Farb() As Byte

Private Const CSIDL_FLAG_CREATE As Long = &H8000
Private Const CSIDL_ADMINTOOLS As Long = &H30
Private Const CSIDL_DESKTOPC = &H0                          'Desktop
Private Const CSIDL_ALTSTARTUP As Long = &H1D               'Alternativer Autostart (DBCS)
Private Const CSIDL_APPDATA As Long = &H1A                  'Anwendungsdaten
Private Const CSIDL_BITBUCKET As Long = &HA                 'Papierkorb
Private Const CSIDL_CDBURN_AREA As Long = &H3B              'CD Burning
Private Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F        'Für alle Benutzer (All Users)
Private Const CSIDL_COMMON_ALTSTARTUP As Long = &H1D        'Alternativer Autostart (DBCS, "All users"-Profil unter NT4)
Private Const CSIDL_COMMON_APPDATA As Long = &H23           'Anwendungsdaten
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19  'Desktop
Private Const CSIDL_COMMON_DOCUMENTS As Long = &H2E         'Dokumente
Private Const CSIDL_COMMON_FAVORITES As Long = &H1F         'Favoriten
Private Const CSIDL_COMMON_MUSIC As Long = &H35             'Gemeinsame Musik
Private Const CSIDL_COMMON_PICTURES As Long = &H36          'Gemeinsame Bilder
Private Const CSIDL_COMMON_PROGRAMS As Long = &H17          'Programme-Ordner im Startmenü
Private Const CSIDL_COMMON_STARTMENU As Long = &H16         'Starmenü
Private Const CSIDL_COMMON_STARTUP As Long = &H18           'Autostart
Private Const CSIDL_COMMON_TEMPLATES As Long = &H2D         'Vorlagen
Private Const CSIDL_COMMON_VIDEO As Long = &H37             'Gemeinsame Videos
Private Const CSIDL_CONTROLS As Long = &H3                  'Systemsteuerung
Private Const CSIDL_COOKIES As Long = &H21                  'Cookies
Private Const CSIDL_DESKTOP As Long = &H0                   'Desktop
Private Const CSIDL_DESKTOPDIRECTORY As Long = &H10         'Desktopverzeichnis
Private Const CSIDL_DRIVES As Long = &H11                   'Treiber
Private Const CSIDL_FAVORITES As Long = &H6                 'Favoriten
Private Const CSIDL_FONTS As Long = &H14                    'Schriftarten
Private Const CSIDL_HISTORY As Long = &H22                  'Verlauf
Private Const CSIDL_INTERNET As Long = &H1                  'Internet Cache (Internet Explorer)
Private Const CSIDL_INTERNET_CACHE As Long = &H20           'Temporäre Internetdateien
Private Const CSIDL_LOCAL_APPDATA As Long = &H1C            'Anwendungsdaten (lokal)
Private Const CSIDL_MYDOCUMENTS As Long = &HC               'Meine Dokumente
Private Const CSIDL_MYMUSIC As Long = &HD                   'Eigene Musik
Private Const CSIDL_MYPICTURES As Long = &H27               'Eigene Bilder
Private Const CSIDL_MYVIDEO As Long = &HE                   'Eigene Videos
Private Const CSIDL_NETHOOD As Long = &H13                  'Netzwerkumgebung
Private Const CSIDL_NETWORK As Long = &H12                  'Netzwerk
Private Const CSIDL_PERSONAL As Long = &H5                  'Eigene Dateien
Private Const CSIDL_PRINTERS As Long = &H4                  'Drucker
Private Const CSIDL_PRINTHOOD As Long = &H1B                'Druckerumgebung
Private Const CSIDL_PROFILE As Long = &H28                  'Profil
Private Const CSIDL_PROGRAM_FILES As Long = &H26            'Programme
Private Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B     'Gemeinsamme Dateien
Private Const CSIDL_PROGRAMS As Long = &H2                  'Programme (im Startmenü)
Private Const CSIDL_RECENT As Long = &H8                    'Zuletzt verwendete Dokumente
Private Const CSIDL_SENDTO As Long = &H9                    'Senden An
Private Const CSIDL_STARTMENU As Long = &HB                 'Startmenü
Private Const CSIDL_STARTUP As Long = &H7                   'Autostart
Private Const CSIDL_SYSTEM As Long = &H25                   'System (bzw. System32)
Private Const CSIDL_TEMPLATES As Long = &H15                'Vorlagen
Private Const CSIDL_WINDOWS As Long = &H24                  'Windows

Private Const FO_MOVE = &H1                 ' Verschieben
Private Const FO_COPY = &H2                 ' Kopieren
Private Const FO_DELETE = &H3               ' Löschen (in den Papierkorb)
Private Const FO_RENAME = &H4               ' Umbenennen
Private Const FOF_MULTIDESTFILES = &H1      ' pTo gibt für jedes Element in pFrom einen Zielpfad an
Private Const FOF_CONFIRMMOUSE = &H2        ' Bisher nicht implementiert
Private Const FOF_SILENT = &H4              ' Keine Dialoge anzeigen
Private Const FOF_RENAMEONCOLLISION = &H8   ' Bei Namenskollisionen automatisch umbenennen
Private Const FOF_NOCONFIRMATION = &H10     ' Keine Bestätigungen durch den Benutzer - alle Fragen werden mit "Ja" beantwortet
Private Const FOF_WANTMAPPINGHANDLE = &H20  ' Operation soll in hNameMaps geloggt werden (hier unberücksichtigt)
Private Const FOF_ALLOWUNDO = &H40          ' Falls möglich werden Undo-Informationen gesichert
Private Const FOF_FILESONLY = &H80          ' Bei Angabe von WildCards (*.*) sollen nur Dateien - nicht Verzeichnisse -  berücksichtigt werden
Private Const FOF_SIMPLEPROGRESS = &H100    ' Einfacher Fortschrittsbalken - keine Dateinamen anzeigen
Private Const FOF_NOCONFIRMMKDIR = &H200    ' Erzeugung von Verzeichnissen nicht bestätigen lassen

Private Const DRIVE_CDROM = 5           'CD-ROM-Laufwerk
Private Const DRIVE_FIXED = 3           'Festplatte
Private Const DRIVE_RAMDISK = 6         'RAM-Disklaufwerk
Private Const DRIVE_REMOTE = 4          'Netzlaufwerk
Private Const DRIVE_REMOVABLE = 2       'Wechselbar, z.B. Diskette oder ZIP-Drive

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Private Const SHERB_NORMAL = &H0
Private Const SHERB_NOCONFIRMATION = &H1
Private Const SHERB_NOPROGRESSUI = &H2
Private Const SHERB_NOSOUND = &H4

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const INVALID_HANDLE_VALUE = -1

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const STILL_ACTIVE = &H103
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_USENEWUI = &H50

Private Const Large = &H100
Private Const Small = &H101

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

Private Const SHFD_CAPACITY_DEFAULT = 0 ' Standard-Kapazität des Laufwerks
Private Const SHFD_CAPACITY_360 = 3     ' 360 kB (nur 5 1/4"-Laufwerke)
Private Const SHFD_CAPACITY_720 = 5     ' 720 kB (nur 3.5"-Laufwerke)
Private Const SHFD_FORMAT_QUICK = 1     ' Quickformat
Private Const SHFD_FORMAT_FULL = 0      ' Vollständig formatieren
Private Const SHFD_FORMAT_SYSONLY = 2   ' Systemdateien kopieren
 
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
 
Private Const CREATE_ALWAYS = 2
Private Const CREATE_NEW = 1
Private Const OPEN_ALWAYS = 4
Private Const OPEN_EXISTING = 3
Private Const TRUNCATE_EXISTING = 5
 
' CreateFile dwFlagsAndAttributes
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20 ' Archiv-Datei
Private Const FILE_ATTRIBUTE_HIDDEN = &H2 ' Versteckt
Private Const FILE_ATTRIBUTE_NORMAL = &H80 ' Normal
Private Const FILE_ATTRIBUTE_READONLY = &H1 ' Schreibgeschützt
Private Const FILE_ATTRIBUTE_SYSTEM = &H4 ' Systemdatei
Private Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800&
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10&
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100&
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000 ' Erlaubt Case-Sensitive-Dateinamen
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_WRITE_THROUGH = &H80000000

Private Const MAX_PATH As Long = 259
Private Const CP_UTF8 As Long = 65001

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SETSELECTION As Long = &H467
Private Const WM_USER As Long = &H400

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type BrowseCallbackParams
    hwnd As Long
    InitialFolder As String
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFilename As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function SearchTree Lib "imagehlp.dll" Alias "SearchTreeForFile" (ByVal lpRootPath As String, ByVal lpInputPathName As String, ByVal lpOutputPathBuffer As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, lpFilePart As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal lngDrive As Long, ByVal lngCapacity As Long, ByVal lngFormatType As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFilename As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ShellExExInfo As SHELLEXECUTEINFO) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "OLE32.dll" (ByVal pv As Long)
Private Sub Class_Terminate()

Set FM = Nothing

mDrv = vbNullString
mPfa = vbNullString
mOrd = vbNullString
mNam = vbNullString
mNao = vbNullString
mExt = vbNullString
mSta = vbNullString
mStr = vbNullString
mDat = vbNullString
mLoe = vbNullString
mCop = vbNullString

End Sub

Public Property Let DaCop(ByVal Flag As String)
Attribute DaCop.VB_Description = "Kopierstring, immer kompletter Pfad + Dateiname, immer mit vbNullChar trennen + Ende mit vbNullChar, Zielpfad mit ; trennen"
    mCop = Flag
End Property

Public Property Get DaCop() As String
    DaCop = mCop
End Property

Public Property Let DaDrv(ByVal Flag As String)
Attribute DaDrv.VB_Description = "Laufwerksbuchstabe mit \\"
    mDrv = Flag
End Property

Public Property Get DaDrv() As String
    DaDrv = mDrv
End Property

Public Property Let DaExt(ByVal Flag As String)
Attribute DaExt.VB_Description = "Dateiextension"
    mExt = Flag
End Property

Public Property Get DaExt() As String
    DaExt = mExt
End Property

Public Property Let DaLoe(ByVal Flag As String)
Attribute DaLoe.VB_Description = "Löschstring, immer kompletter Pfad + Dateiname, immer mit vbNullChar trennen + Ende mit vbNullChar"
    mLoe = Flag
End Property

Public Property Get DaLoe() As String
    DaLoe = mLoe
End Property

Public Property Let DaNam(ByVal Flag As String)
Attribute DaNam.VB_Description = "Dateiname mit Extension"
    mNam = Flag
End Property

Public Property Get DaNam() As String
    DaNam = mNam
End Property

Public Property Let DaNaO(ByVal Flag As String)
Attribute DaNaO.VB_Description = "Dateiname ohne Extension"
    mNao = Flag
End Property

Public Property Get DaNaO() As String
    DaNaO = mNao
End Property

Public Property Let DaOrd(ByVal Flag As String)
Attribute DaOrd.VB_Description = "Speicherordner ohne \\"
    mOrd = Flag
End Property

Public Property Get DaOrd() As String
    DaOrd = mOrd
End Property

Public Property Let DaPfa(ByVal Flag As String)
Attribute DaPfa.VB_Description = "Kompletter Pfad ohne Dateiname ohne \\"
    mPfa = Flag
End Property

Public Property Get DaPfa() As String
    DaPfa = mPfa
End Property
Public Property Let DaPar(ByVal Flag As String)
    mPar = Flag
End Property
Public Property Get DaPar() As String
    DaPar = mPar
End Property

Public Property Let DaStr(ByVal Flag As String)
Attribute DaStr.VB_Description = "Filestring für Dateidialog"
    mStr = Flag
End Property

Public Property Get DaStr() As String
    DaStr = mStr
End Property
Public Property Let DaTit(ByVal Flag As String)
    mTit = Flag
End Property

Public Property Get DaTit() As String
    DaTit = mTit
End Property

Public Function FilActi(ByVal TskId As Long) As Boolean
Attribute FilActi.VB_Description = "Ist Anwendung aktiv? Vorher Task-ID ermitteln"
On Error Resume Next
'Ist Anwendung aktiv? Vorher Task-ID ermitteln

Dim RetWe As Long
Dim FHwnd As Long
Dim ExCod As Long

FHwnd = OpenProcess(PROCESS_ALL_ACCESS, False, TskId)

RetWe = GetExitCodeProcess(FHwnd, ExCod)
RetWe = CloseHandle(FHwnd)

FilActi = IIf(ExCod = STILL_ACTIVE, True, False)

End Function

Public Function FilAnw(ByVal FiNa As String) As String
Attribute FilAnw.VB_Description = "Liefert die passende Anwendung zu einer Datei, wird über FilPfa ausgegeben"
On Error Resume Next
'Liefert die passende Anwendung zu einer Datei, wird über FilPfa ausgegeben

Dim RetWe As Long
Dim TmStr As String * 256

RetWe = FindExecutable(FiNa, App.Path, TmStr)

If RetWe > 32 Then
        FilAnw = Left$(TmStr, InStr(TmStr, vbNullChar) - 1)
    Else
        FilAnw = vbNullString
    End If
End Function

Public Function FilAusf() As Long
Attribute FilAusf.VB_Description = "Führt das angegebene Programm aus, benötigt DaPfa und DaNam"
On Error Resume Next
'Führt das angegebene Programm aus, benötigt DaPfa und DaNam

Dim RetWe As Long

RetWe = ShellExecute(mHwnd, "open", mNam & vbNullString, mPfa & vbNullString, mPar & vbNullString, SW_NORMAL)

FilAusf = RetWe

End Function

Private Function FilBySt(ByAry() As Byte) As String
On Error Resume Next
'Wandelt Bytefolge in String um

Dim ByStr As String
Dim PoStr As String

ByStr = StrConv(ByAry, vbUnicode)
PoStr = InStr(ByStr, Chr(0))
If PoStr > 0 Then
    ByStr = Left(ByStr, PoStr - 1)
End If
FilBySt = ByStr
 
End Function
Public Function FilDow(UrlNa As String, FiNam As String) As Boolean
On Error Resume Next

Dim RetWe As Long

RetWe = URLDownloadToFile(0, UrlNa, FiNam, 0, 0)

If RetWe = 0 Then FilDow = True
    
End Function
Public Function FilCnRe(ByRef FileName As String) As String
On Error Resume Next
'Convertiert von UTF-8

Dim FileNumber As Integer
Dim Buffer() As Byte

FileNumber = FreeFile
Open FileName For Binary Access Read As #FileNumber
ReDim Buffer(0 To LOF(FileNumber) - 1)
Get #FileNumber, , Buffer
Close #FileNumber
FilCnRe = FilCoRe(Buffer)
    
End Function
Public Sub FilCnWr(ByRef FileName As String, ByRef Contents As String)
On Error Resume Next
'Convertiert nach UTF-8

Dim FileNumber As Integer
Dim Buffer() As Byte

Buffer = FilCoWr(Contents)
FileNumber = FreeFile
Open FileName For Binary Access Write As #FileNumber
Put #FileNumber, , Buffer
Close #FileNumber
    
End Sub
Public Function FilCol(ByVal Farb As Long) As Long
Attribute FilCol.VB_Description = "Zeigt den Farbdialog an"
On Error Resume Next
'Zeigt den Farbdialog an

Dim RetWe As Long
Dim GetCo As CHOOSECOLOR
Dim Custcolor(16) As Long

With GetCo
    .lStructSize = Len(GetCo)
    .hwndOwner = mHwnd
    .hInstance = 0
    .lpCustColors = StrConv(Farb, vbUnicode)
    .flags = 0
End With

RetWe = ChooseColorAPI(GetCo)

If RetWe <> 0 Then
    FilCol = GetCo.rgbResult
    Farb = StrConv(GetCo.lpCustColors, vbFromUnicode)
Else
    FilCol = -1
End If

End Function

Public Function FilCop(ByVal KoTyp As Integer) As Boolean
Attribute FilCop.VB_Description = "Kopiert, Verschiebt eine Datei oder benennt diese um"
On Error Resume Next
'Kopiert, Verschiebt eine Datei oder benennt diese um

Dim RetWe As Long
Dim DaStr As String
Dim PfStr As String
Dim Lange As Integer
Dim StaZe As Integer
Dim Posit As Integer
Dim Obstr As SHFILEOPSTRUCT

FilCop = False
StaZe = 1
Lange = Len(mCop)
Do
Posit = InStr(StaZe, mCop, ";", 1)
If Posit > 0 Then
    StaZe = Posit + 1
End If
Loop Until Posit = 0
If StaZe > 1 Then
    PfStr = Mid$(mCop, StaZe, Lange - (StaZe - 1)) & "\"
    DaStr = Left$(mCop, StaZe - 2)
Else
    PfStr = vbNullString
    DaStr = vbNullString
End If

If Not DaStr = vbNullString And Not PfStr = vbNullString Then
    With Obstr
        Select Case KoTyp
        Case 1: .wFunc = FO_COPY
        Case 2: .wFunc = FO_MOVE
        Case 3: .wFunc = FO_RENAME
        End Select
        .pFrom = DaStr & vbNullChar & vbNullChar
        .pTo = PfStr & vbNullChar & vbNullChar
        .fFlags = FOF_NOCONFIRMATION Or FOF_RENAMEONCOLLISION Or FOF_NOCONFIRMMKDIR
    End With
    RetWe = SHFileOperation(Obstr)
    If RetWe = 0 Then
        FilCop = True
    End If
End If

End Function
Private Function FilCoWr(ByRef Source As String) As Byte()
On Error Resume Next
'Convertiert nach UTF-8
    
Dim length As Long
Dim Pointer As Long
Dim SIZE As Long
Dim Buffer() As Byte

length = Len(Source)
Pointer = StrPtr(Source)
SIZE = WideCharToMultiByte(CP_UTF8, 0, Pointer, length, 0, 0, 0, 0)
ReDim Buffer(0 To SIZE - 1)

WideCharToMultiByte CP_UTF8, 0, Pointer, length, VarPtr(Buffer(0)), SIZE, 0, 0

FilCoWr = Buffer
    
End Function
Public Function FilDel(Optional Recyc As Boolean = False) As Boolean
'Datei oder Order Löschen

Dim SHFileOp As SHFILEOPSTRUCT

If Right(mLoe, 1) = "\" Then
    mLoe = Left(mLoe, Len(mLoe) - 1)
End If

With SHFileOp
    .wFunc = FO_DELETE
    .pFrom = mLoe
    .fFlags = FOF_NOCONFIRMATION Or FOF_FILESONLY
    If Recyc Then .fFlags = .fFlags Or FOF_ALLOWUNDO
End With
FilDel = (SHFileOperation(SHFileOp) = 0)

End Function
Public Function FilDir(ByVal PfaNa As String) As Boolean
On Error Resume Next
'Prüft, ob Verzeichnis vorhanden ist

Dim RetWe As Long

RetWe = MakeSureDirectoryPathExists(PfaNa)

If RetWe = 0 Then
    FilDir = False
Else
    FilDir = True
End If

End Function

Private Function FilCoRe(ByRef Source() As Byte) As String
On Error Resume Next
'Convertiert von UTF-8

Dim SIZE As Long
Dim Pointer As Long
Dim length As Long
Dim Buffer As String

SIZE = UBound(Source) - LBound(Source) + 1
Pointer = VarPtr(Source(LBound(Source)))
length = MultiByteToWideChar(CP_UTF8, 0, Pointer, SIZE, 0, 0)
Buffer = Space$(length)
MultiByteToWideChar CP_UTF8, 0, Pointer, SIZE, StrPtr(Buffer), length
FilCoRe = Buffer
    
End Function
Public Function FilDrv(ByVal Flag As Integer) As String
Attribute FilDrv.VB_Description = "Liefert das jeweils erste Laufwerk seiner Art"
On Error Resume Next
'Liefert das jeweils erste Laufwerk seiner Art

Dim DrvTy As Long
Dim StrSt As Long
Dim RetWe As Long
Dim AkDrv As String
Dim ReStr As String
Dim DrStr As String * 150

FilDrv = vbNullString
StrSt = 1
RetWe = GetLogicalDriveStrings(150, DrStr)

If RetWe <> 0 Then
    AkDrv = Mid$(DrStr, StrSt, 3)
    Do
    DrvTy = GetDriveType(AkDrv)
    Select Case Flag
    Case 1: GoSub TyDis
    Case 2: GoSub TyFes
    Case 3: GoSub TyZip
    Case 4: GoSub TyCDR
    End Select
    If Not ReStr = vbNullString Then
        FilDrv = AkDrv
        Exit Do
    End If
    StrSt = StrSt + 4
    AkDrv = Mid$(DrStr, StrSt, 3)
    Loop While (Mid$(DrStr, StrSt, 1) <> vbNullChar)
End If

Exit Function

TyDis:
If DrvTy = DRIVE_REMOVABLE Then
    If AkDrv = "a:\" Then
        ReStr = AkDrv
    Else
        ReStr = vbNullString
    End If
Else
    ReStr = vbNullString
End If
Return

TyFes:
If DrvTy = DRIVE_FIXED Then
    If Not AkDrv = "c:\" Then
        ReStr = AkDrv
    Else
        ReStr = vbNullString
    End If
Else
    ReStr = vbNullString
End If
Return

TyZip:
If DrvTy = DRIVE_REMOVABLE Then
    If Not AkDrv = "a:\" And Not AkDrv = "b:\" Then
        ReStr = AkDrv
    Else
        ReStr = vbNullString
    End If
Else
    ReStr = vbNullString
End If
Return

TyCDR:
If DrvTy = DRIVE_CDROM Then
    ReStr = AkDrv
Else
    ReStr = vbNullString
End If
Return

End Function

Public Function FilFld() As String
Attribute FilFld.VB_Description = "Zeigt Ordnerauswahl"
On Error Resume Next
'Zeigt Ordnerauswahl
  

  


End Function

Public Function FilFor()
Attribute FilFor.VB_Description = "Formatiert ein Laufwerk"
On Error Resume Next
'Formatiert ein Laufwerk

Dim RetWe As Long

RetWe = SHFormatDrive(mHwnd, 0, SHFD_CAPACITY_DEFAULT, SHFD_FORMAT_FULL)

End Function
Public Function FilFree(ByVal DrvNa As String) As Single
On Error Resume Next
'Liefert freien Speicherplatz

Dim FreBy As Currency
Dim Kompl As Currency
Dim FrSpa As Currency
Dim StrTm As String

If Len(DrvNa) = 1 Then DrvNa = DrvNa + ":\"

If GetDiskFreeSpaceEx(DrvNa, FreBy, Kompl, FrSpa) = 0 Then FrSpa = -1

StrTm = CStr(FrSpa)

If InStr(StrTm, ",") > 0 Then StrTm = Left$(StrTm, InStr(StrTm, ",") - 1) + Mid$(StrTm, InStr(StrTm, ",") + 1)

FilFree = Format$(CSng(StrTm) / 1000000000, "#,##0.00")

End Function
Public Sub FilInfo(ByVal hwnd As Long, ByVal FiNam As String)
On Error Resume Next
'Liefert Informationsdialog zu einer Datei, benötigt DaPfa und DaNam

Dim FILEINFO As SHELLEXECUTEINFO

With FILEINFO
    .cbSize = Len(FILEINFO)
    .fMask = SEE_MASK_FLAG_NO_UI Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_NOCLOSEPROCESS
    .hwnd = hwnd
    .lpVerb = "properties"
    .lpFile = FiNam
End With

ShellExecuteEx FILEINFO

End Sub
Public Function FilLis(ByVal StVer As String, ByVal SuMus As String, ByRef DiNam() As String) As Integer
On Error Resume Next
'Auflisten aller Dateien in einem Ordner optional aller Unterordner

Dim RetWe As Long
Dim FiNam As String
Dim TmpSt As String
Dim AktZa As Integer
Dim WinDat As WIN32_FIND_DATA

ReDim DiNam(0)

If Right(StVer, 1) <> "\" Then StVer = StVer & "\"

RetWe = FindFirstFile(StVer & SuMus, WinDat)

If RetWe = 0 Then
    FilLis = 0
    Exit Function
End If

Do
TmpSt = Left(WinDat.cFileName, InStr(WinDat.cFileName, Chr(0)) - 1)
If TmpSt <> vbNullString Then
    FiNam = UCase(Left$(TmpSt, 1)) & Mid$(TmpSt, 2, Len(TmpSt) - 1)
    If (WinDat.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
        'If (FiNam <> ".") And (FiNam <> "..") Then
        '    AktZa = AktZa + 1
        '    ReDim Preserve DiNam(AktZa)
        '    DiNam(AktZa) = FiNam
        'End If
    Else
        AktZa = AktZa + 1
        ReDim Preserve DiNam(AktZa)
        DiNam(AktZa) = FiNam
    End If
End If
Loop While FindNextFile(RetWe, WinDat)

RetWe = FindClose(RetWe)

FilLis = AktZa

FiSor LBound(DiNam), UBound(DiNam), DiNam

End Function
Public Function FilLoe() As Boolean
Attribute FilLoe.VB_Description = "Löscht eine oder mehrer Dateien (auch * erlaubt), benötigt DaLoe"
On Error Resume Next
'Löscht eine oder mehrer Dateien (auch * erlaubt), benötigt DaLoe

Dim RetWe As Long
Dim Obstr As SHFILEOPSTRUCT

FilLoe = False
If Not IsNull(mLoe) Then
    With Obstr
        .wFunc = FO_DELETE
        .pFrom = mLoe & vbNullChar & vbNullChar
        .fFlags = FOF_SILENT + FOF_NOCONFIRMATION
    End With
    RetWe = SHFileOperation(Obstr)
    If RetWe = 0 Then
        FilLoe = True
    End If
End If

End Function

Public Function FilOpn() As String
Attribute FilOpn.VB_Description = "Zeigt den Öffnen Dialog, benötigt hWnd, DaStr und StaVe"
On Error Resume Next
'Zeigt den Öffnen Dialog, benötigt hWnd, DaStr und StaVe

Dim RetWe As Long
Dim Posit As Long
Dim FiStr As String
Dim FiNam As String
Dim FiCol As Collection
Dim OpnFl As OPENFILENAME

Set FiCol = New Collection

' Initialisieren des OPENFILENAME-Structs
With OpnFl
    .lStructSize = Len(OpnFl)
    .hwndOwner = mHwnd
    .lpstrFilter = mStr
    .lpstrCustomFilter = vbNullString
    .lpstrFile = String$(4096, 0) ' Puffer für Dateinamen
    .nMaxFile = Len(OpnFl.lpstrFile)
    .lpstrTitle = mTit
    .lpstrInitialDir = mSta
    .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
End With

RetWe = GetOpenFileName(OpnFl)

If RetWe <> 0 Then
    FiStr = Left$(OpnFl.lpstrFile, InStr(OpnFl.lpstrFile, vbNullChar) - 1)

    Posit = InStr(OpnFl.lpstrFile, vbNullChar) + 1
    
    Do While Posit < Len(OpnFl.lpstrFile)
        Dim TmStr As String
        TmStr = Mid$(OpnFl.lpstrFile, Posit, InStr(Posit, OpnFl.lpstrFile, vbNullChar) - Posit)
        If Len(TmStr) = 0 Then Exit Do
        
        FiCol.Add FiStr & "\" & TmStr
        Posit = Posit + Len(TmStr) + 1
    Loop

    If FiCol.Count = 0 Then
        FiCol.Add FiStr
    End If
        
    Dim FiLis As Variant
    For Each FiLis In FiCol
        If FiNam = vbNullString Then
            FiNam = FiLis
        Else
            FiNam = FiNam & Chr$(59) & FiLis
        End If
    Next FiLis
    
    FilOpn = FiNam
Else
    FilOpn = vbNullString
End If

End Function

Public Function FilOrd(ByVal StVer As String, ByRef OrdNa() As String, ByRef OrdPf() As String) As Integer
On Error Resume Next
'Auflisten aller Dateien in einem Ordner optional aller Unterordner

Dim RetWe As Long
Dim FiNam As String
Dim AktZa As Integer
Dim WinDat As WIN32_FIND_DATA

ReDim OrdNa(0)

If Right(StVer, 1) <> "\" Then StVer = StVer & "\"

RetWe = FindFirstFile(StVer & "*.*", WinDat)

If RetWe = 0 Then
    FilOrd = 0
    Exit Function
End If

Do
FiNam = Left(WinDat.cFileName, InStr(WinDat.cFileName, Chr(0)) - 1)
If (WinDat.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
    If (FiNam <> ".") And (FiNam <> "..") Then
        AktZa = AktZa + 1
        ReDim Preserve OrdNa(AktZa)
        ReDim Preserve OrdPf(AktZa)
        OrdNa(AktZa) = FiNam
        OrdPf(AktZa) = StVer & FiNam
    End If
End If
Loop While FindNextFile(RetWe, WinDat)

FindClose RetWe

FilOrd = AktZa

End Function
Public Sub FilPfa(ByVal FiNa As String)
Attribute FilPfa.VB_Description = "Liefer den Pfad, Namen und Endung"
On Error Resume Next
'Liefer den Pfad, Namen und Endung

Dim Posi As Long
Dim NePo As Long
Dim DaNa As String

If Mid$(FiNa, 2, 1) = ":" Then
    mDrv = Left(FiNa, 3)
    DaNa = Right(FiNa, Len(FiNa) - 3)
Else
    mDrv = vbNullString
End If

Posi = InStrRev(FiNa, "\")
If Posi = 0 Then
    mPfa = FiNa
    mNam = vbNullString
    mExt = vbNullString
    mOrd = vbNullString
Else
    mPfa = Left$(FiNa, Posi - 1)
    NePo = InStrRev(mPfa, "\")
    If NePo > 0 Then
        mOrd = Right$(mPfa, Len(mPfa) - NePo)
    Else
        mOrd = mPfa
    End If
    DaNa = Right$(FiNa, Len(FiNa) - Posi)
    Posi = InStrRev(DaNa, ".")
    If Posi = 0 Then
        mNam = vbNullString
        mNao = vbNullString
        mExt = vbNullString
    Else
        mNam = DaNa
        mNao = Left$(DaNa, Posi - 1)
        mExt = Right$(DaNa, Len(DaNa) - Posi)
    End If
End If

End Sub
Public Function FilReIn(ByVal FiNam As String) As String
On Error Resume Next
'Liest die Dateiinformationen

Dim hFile As Long
Dim ZeStr As String
Dim CTime As FILETIME
Dim STime As SYSTEMTIME
Dim Dummi As FILETIME

hFile = CreateFile(FiNam, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0&)

If hFile = -1 Then Exit Function

GetFileTime hFile, CTime, Dummi, Dummi
  
FileTimeToLocalFileTime CTime, CTime

FileTimeToSystemTime CTime, STime

ZeStr = Format$(STime.wYear, "0000") & "." & Format$(STime.wMonth, "00") & "." & Format$(STime.wDay, "00") & " - " & Format$(STime.wHour, "00") & ":" & Format$(STime.wMinute, "00") & ":" & Format$(STime.wSecond, "00")

FilReIn = ZeStr

CloseHandle hFile

End Function
Public Function FilReBy() As Byte()
Attribute FilReBy.VB_Description = "Ließt Bytefolge aus einer Datei, benötigt DaPfa und DaNam"
On Error Resume Next
'Liest Bytefolge aus einer Datei, benötigt DaPfa und DaNam

Dim RetWe As Long
Dim FilWe As Long
Dim ByRea As Long
Dim ByFil As Long
Dim ByAry() As Byte

ByFil = FileLen(mPfa & "\" & mNam)
ReDim ByAry(ByFil) As Byte
RetWe = CreateFile(mPfa & "\" & mNam, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

If RetWe <> INVALID_HANDLE_VALUE Then
   FilWe = ReadFile(RetWe, ByAry(0), ByFil, ByRea, 0)
   FilReBy = ByAry
   CloseHandle RetWe
End If

End Function

Public Function FilReSt() As String
Attribute FilReSt.VB_Description = "Ließt String aus einer Datei, benötigt DaPfa und DaNam"
On Error Resume Next
'Liest String aus einer Datei, benötigt DaPfa und DaNam

Dim RetWe As Long
Dim FilWe As Long
Dim ByRea As Long
Dim ByFil As Long
Dim ByStr As String
Dim ByAry() As Byte

ByFil = FileLen(mPfa & "\" & mNam)
ReDim ByAry(ByFil) As Byte
RetWe = CreateFile(mPfa & "\" & mNam, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

If RetWe <> INVALID_HANDLE_VALUE Then
    FilWe = ReadFile(RetWe, ByAry(0), ByFil, ByRea, 0)
    ByStr = FilBySt(ByAry)
    FilReSt = ByStr
    CloseHandle RetWe
End If

End Function

Public Function FilSav() As String
Attribute FilSav.VB_Description = "Zeigt den Speichern Dialog, benötigt hWnd, DaStr, StaVe und DaNam"
On Error Resume Next
'Zeigt den Speichern Dialog, benötigt hWnd, DaStr, StaVe und DaNam

Dim RetWe As Long
Dim OpnFl As OPENFILENAME

With OpnFl
    .lStructSize = Len(OpnFl)
    .hwndOwner = mHwnd
    .nFilterIndex = 1
    If mNam = vbNullString Then
        .lpstrFile = String$(4096, 0) 'Leerer Puffer
    Else
        .lpstrFile = mNam & String$(4096 - Len(mNam), 0) 'Puffer mit vorgeschlagenem Dateinamen füllen
    End If
    .nMaxFile = Len(OpnFl.lpstrFile)
    .lpstrInitialDir = mSta
    .lpstrFilter = mStr
    .lpstrTitle = mTit
    .flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
    .lpstrDefExt = mExt
End With

RetWe = GetSaveFileName(OpnFl)

If RetWe <> 0 Then
    FilSav = Left(OpnFl.lpstrFile, InStr(OpnFl.lpstrFile, vbNullChar) - 1)
Else
    FilSav = vbNullString
End If

End Function
Public Function FilStart() As Long
On Error GoTo ErrHdl
'Startet eine Anwendung mit Datei und Parameter

Dim RetWe As Long
Dim ParStr As String

' Parameter String bauen: Pfade mit Leerzeichen in Anfuehrungszeichen
' Each argument that may contain spaces must be wrapped in quotes
If Len(mPfa) > 0 Then
    ParStr = """" & mPfa & """"
Else
    ParStr = vbNullString
End If

If Len(mPar) > 0 Then
    ParStr = ParStr & Space$(1) & """" & mPar & """"
End If

RetWe = ShellExecute(mHwnd, "open", mNam, ParStr, vbNullString, SW_NORMAL)

' ShellExecute returns > 32 on success, <= 32 indicates error
If RetWe <= 32 Then
    If GlLog = True Then SLogi "FilStart: ShellExecute failed, code=" & CStr(RetWe) & ", exe=" & mNam & ", par=" & ParStr
End If

FilStart = RetWe

Exit Function

ErrHdl:
If GlLog = True Then SLogi "FilStart: Error " & Err.Number & " - " & Err.Description & ", exe=" & mNam
FilStart = 0

End Function

Private Sub FiSor(ByVal LB As Long, ByVal UB As Long, ByRef DiNam() As String)
On Error Resume Next

Dim P1 As Long
Dim P2 As Long
Dim ReStr As String
Dim TmpSt As String

P1 = LB
P2 = UB
ReStr = DiNam((P1 + P2) / 2)

Do
    Do While (DiNam(P1) < ReStr)
        P1 = P1 + 1
    Loop

    Do While (DiNam(P2) > ReStr)
        P2 = P2 - 1
    Loop

    If P1 <= P2 Then
        TmpSt = DiNam(P1)
        DiNam(P1) = DiNam(P2)
        DiNam(P2) = TmpSt
        
        P1 = P1 + 1
        P2 = P2 - 1
    End If
Loop Until (P1 > P2)

If LB < P2 Then Call FiSor(LB, P2, DiNam())
If P1 < UB Then Call FiSor(P1, UB, DiNam())

End Sub


Public Function FilSys(ByVal Flag As Integer) As String
Attribute FilSys.VB_Description = "Liefert bestimmte Systempfade"
On Error Resume Next
'Liefert bestimmte Systempfade

Dim RetWe As Long
Dim Lange As Long
Dim RetSt As String
Dim SysPf As String * 512

Select Case Flag
Case 1: RetWe = GetSystemDirectory(SysPf, 512)
Case 2: RetWe = GetWindowsDirectory(SysPf, 512)
Case 3: RetWe = GetTempPath(512, SysPf)
End Select

If RetWe > 0 Then
    RetSt = Left$(SysPf, InStr(SysPf, vbNullChar) - 1)
    Lange = Len(RetSt)
    If Flag = 3 Then
        FilSys = Left$(RetSt, Lange - 1)
    Else
        FilSys = RetSt
    End If
Else
    FilSys = vbNullString
End If

End Function

Public Function FilTmp() As String
Attribute FilTmp.VB_Description = "Zeigt die nächste freie Temp-Datei"
On Error Resume Next
'Zeigt die nächste freie Temp-Datei

Dim RetWe As Long
Dim TmpPf As String * 512
Dim TmpNa As String * 576

RetWe = GetTempPath(512, TmpPf)
If (RetWe > 0 And RetWe < 512) Then
   RetWe = GetTempFileName(TmpPf, "tmp", 0, TmpNa)
   If RetWe <> 0 Then
      FilTmp = Left$(TmpNa, InStr(TmpNa, vbNullChar) - 1)
    Else
        FilTmp = CurDir$ & "\Temp.tmp"
   End If
Else
    FilTmp = CurDir$ & "\Temp.tmp"
End If

End Function
Public Sub FilTray()
Attribute FilTray.VB_Description = "Löscht den Papierkorn"
On Error Resume Next
'Löscht den Papierkorn

Dim RetWe As Long

RetWe = SHEmptyRecycleBin(0&, vbNullString, SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND)

End Sub
Public Function FilTyp(AkDrv As String) As String
On Error Resume Next
'Gibt des Datenträgertyp zurück

Dim DrvTy As Long

DrvTy = GetDriveType(AkDrv)

Select Case DrvTy
Case DRIVE_CDROM: FilTyp = "CD"
Case DRIVE_REMOTE: FilTyp = "NE"
Case DRIVE_RAMDISK: FilTyp = "RA"
Case DRIVE_REMOVABLE: FilTyp = "WE"
Case DRIVE_FIXED: FilTyp = "FE"
Case Else: FilTyp = "FE"
End Select

End Function
Public Function FilVer(Flag As Integer) As String
Attribute FilVer.VB_Description = "Liefert Spezielle Verzeichnisse"
On Error Resume Next
'Liefert Spezielle Verzeichnisse

Dim RetWe As Long
Dim PfaNa As String
Dim ItmLs As ITEMIDLIST

Select Case Flag
Case 1: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_DESKTOPC, ItmLs)              'Desktop
Case 2: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_PERSONAL, ItmLs)              'Eigene Dateien
Case 3: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_COMMON_DOCUMENTS, ItmLs)      'Dokumente
Case 4: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_MYDOCUMENTS, ItmLs)           'Meine Dokumente
Case 5: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_DESKTOPDIRECTORY, ItmLs)      'Desktopverzeichnis
Case 6: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_NETWORK, ItmLs)               'Netzwerk
Case 7: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_NETHOOD, ItmLs)               'Netzwerkumgebung
Case 8: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_FAVORITES, ItmLs)             'Favoriten
Case 9: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_RECENT, ItmLs)                'Zuletzt verwendete Dokumente
Case 10: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_APPDATA, ItmLs)              'Anwendungsdaten
Case 11: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_COMMON_APPDATA, ItmLs)       'Anwendungsdaten (für Alle)
Case 12: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_LOCAL_APPDATA, ItmLs)        'Anwendungsdaten (Benutzer)
Case 13: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_COMMON_FAVORITES, ItmLs)     'Favoriten
Case 14: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_COMMON_PICTURES, ItmLs)      'Gemeinsame Bilder
Case 15: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_COMMON_PROGRAMS, ItmLs)      'Programme-Ordner im Startmenü
Case 16: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_COMMON_STARTMENU, ItmLs)     'Starmenü
Case 17: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_COMMON_STARTUP, ItmLs)       'Autostart
Case 18: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_HISTORY, ItmLs)              'Verlauf
Case 19: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_MYPICTURES, ItmLs)           'Eigene Bilder
Case 20: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_PROGRAM_FILES_COMMON, ItmLs) 'Gemeinsamme Dateien
Case 21: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_DESKTOP, ItmLs)              'Desktop
Case 22: RetWe = SHGetSpecialFolderLocation(mHwnd, CSIDL_PROGRAM_FILES, ItmLs)        'Programme
End Select

If RetWe = 0 Then
    PfaNa = Space$(255)
    If SHGetPathFromIDList(ItmLs.mkid.cb, PfaNa) <> 0 Then
        FilVer = Left$(PfaNa, InStr(PfaNa, vbNullChar) - 1)
    Else
        FilVer = vbNullString
    End If
Else
    FilVer = vbNullString
End If

End Function
Public Function FilVor(FiNam As String) As Boolean
On Error Resume Next
'Ersatz für die Dir$() Finktion

Dim Gefun As Long
Dim WinDat As WIN32_FIND_DATA

Gefun = FindFirstFile(FiNam, WinDat)

If Gefun <> INVALID_HANDLE_VALUE Then
    FilVor = True
Else
    FilVor = False
End If

End Function

Public Function FilWeBy(FiDat() As Byte, Optional NoOve As Boolean = False) As Boolean
Attribute FilWeBy.VB_Description = "Schteibt Bytefolge in eine Datei, benötigt DaPfa, DaNam"
On Error Resume Next
'Schteibt Bytefolge in eine Datei, benötigt DaPfa, DaNam

Dim RetWe As Long
Dim FilWe As Long
Dim ByFer As Long
Dim ByWri As Long

If NoOve = True And Dir(mPfa & "\" & mNam) <> vbNullString Then Exit Function
ByWri = UBound(FiDat) - LBound(FiDat)
RetWe = CreateFile(mPfa & "\" & mNam, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

If RetWe <> INVALID_HANDLE_VALUE Then
    FilWe = WriteFile(RetWe, FiDat(0), ByWri, ByFer, 0) <> 0
    If FilWe <> 0 Then
        FilWe = FlushFileBuffers(RetWe)
        FilWe = CloseHandle(RetWe)
    End If
End If

End Function

Public Function FilWrSt(Optional NoOve As Boolean = False) As Boolean
Attribute FilWrSt.VB_Description = "Schteibt String in eine Datei, benötigt DaPfa, DaNam und StrDa"
On Error Resume Next
'Schteibt String in eine Datei, benötigt DaPfa, DaNam und StrDa

Dim RetWe As Long
Dim FilWe As Long
Dim ByFer As Long
Dim ByWri As Long

If NoOve = True And Dir(mPfa & "\" & mNam) <> vbNullString Then Exit Function
ByWri = Len(StrDa)
RetWe = CreateFile(mPfa & "\" & mNam, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

If RetWe <> INVALID_HANDLE_VALUE Then
    FilWe = WriteFile(RetWe, ByVal StrDa, ByWri, ByFer, 0) <> 0
    If FilWe <> 0 Then
        FilWe = FlushFileBuffers(RetWe)
        FilWe = CloseHandle(RetWe)
    End If
End If

End Function
Public Property Let hwnd(ByVal Flag As Long)
    mHwnd = Flag
End Property

Public Property Get hwnd() As Long
    hwnd = mHwnd
End Property

Public Property Let StaVe(ByVal Flag As String)
Attribute StaVe.VB_Description = "Startverzeichnis für Dateidialog"
    mSta = Flag
End Property

Public Property Get StaVe() As String
    StaVe = mSta
End Property

Public Property Let StrDa(ByVal Flag As String)
Attribute StrDa.VB_Description = "Dateninhalt für Schreiboperationen"
    mDat = Flag
End Property

Public Property Get StrDa() As String
    StrDa = mDat
End Property
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If uMsg = BFFM_INITIALIZED Then
        ' Setzt den Startordner
        Call SendMessage(hwnd, BFFM_SETSELECTION, True, ByVal lpData)
    End If
    BrowseCallbackProc = 0
End Function
Private Function GetAddressOf(func As Long) As Long
    GetAddressOf = func
End Function
