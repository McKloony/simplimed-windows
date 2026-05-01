VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnwend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As VB.Form

Private mHwn As Long   'Fensterhandle
Private mNam As String 'Name der Anwendung

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 260
Private Const TH32CS_SNAPPROCESS = &H2&
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const WS_VISIBLE = &H10000000
Private Const WS_BORDER = &H800000
Private Const WM_CLOSE = &H10

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5
Private Const GWL_STYLE = (-16)

Private Type PROCESSENTRY32
    dwSize As Long                 ' Größe der Struktur in Bytes
    cntUsage As Long               ' Anzahl der Referenzen auf den Prozess
    th32ProcessID As Long          ' Prozess-ID
    th32DefaultHeapID As Long      ' ID des Heaps für den Prozess
    th32ModuleID As Long           ' Modul-ID (nur für ToolHelp-Funktionen)
    cntThreads As Long             ' Anzahl der Threads des Prozesses
    th32ParentProcessID As Long    ' ID des Prozesses, der den Prozess gestartet hat
    pcPriClassBase As Long         ' Basisprioritaet des Prozesses
    dwFlags As Long                ' (Reserviert)
    szExeFile As String * MAX_PATH ' Pfad und Dateiname
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal wIndx As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Property Let AnHan(Flag As Long)
Attribute AnHan.VB_Description = "Fensterhandle"
    mHwn = Flag
End Property

Public Property Get AnHan() As Long
    AnHan = mHwn
End Property

Public Property Let AnNam(ByVal Flag As String)
Attribute AnNam.VB_Description = "Name der gesuchten Anwendung"
    mNam = Flag
End Property

Public Property Get AnNam() As String
    AnNam = mNam
End Property

Private Function AnwFil(ByVal FeHan As Long) As Long
On Error Resume Next
'Filtert die Fenstertexte

Dim RetWe As Long
Dim Style As Long
Dim FeTit As String

Style = GetWindowLong(FeHan, GWL_STYLE)
Style = Style And (WS_VISIBLE Or WS_BORDER)
RetWe = GetWindowTextLength(FeHan) + 1
FeTit = Space$(RetWe)
RetWe = GetWindowText(FeHan, FeTit, RetWe)
FeTit = Left$(FeTit, Len(FeTit) - 1)

If (FeTit <> vbNullString) And Style = (WS_VISIBLE Or WS_BORDER) Then
    If FeTit = mNam Then
        AnwFil = FeHan
    End If
End If

End Function
Public Function AnwHa() As Long
Attribute AnwHa.VB_Description = "Gibt das Fensterhandle der gesuchten Anwendung zurück. Benötigt ein Starthandle"
On Error Resume Next
'Findes das Fensterhandle zum Anwendungsnamen

Dim FeHan As Long
Dim RetHa As Long

FeHan = GetWindow(mHwn, GW_HWNDFIRST)

Do
RetHa = AnwFil(FeHan)
FeHan = GetWindow(FeHan, GW_HWNDNEXT)
    If RetHa <> 0 Then
        AnwHa = RetHa
        Exit Do
    End If
Loop Until FeHan = 0

End Function

Public Function AnwID(ByVal CoStr As String) As Long
On Error Resume Next

Dim Posi1 As Long

Posi1 = InStr(CoStr, vbTab)

If Posi1 > 0 Then AnwID = CLng(Left$(CoStr, Posi1 - 1))

End Function
Private Function AnwIn(ByVal mHwnd As Long, Optional ByVal SuStr As String) As Long
On Error Resume Next

Dim PaFen As Long
Dim TskId As Long
Dim RetWe As Long
Dim Style As Long
Dim FeCap As String

'Darstellung des Fensters
Style = GetWindowLong(mHwnd, GWL_STYLE)
Style = Style And (WS_VISIBLE Or WS_BORDER)
     
'FeCap des Fenster auslesen
RetWe = GetWindowTextLength(mHwnd) + 1
FeCap = Space$(RetWe)
RetWe = GetWindowText(mHwnd, FeCap, RetWe)
FeCap = Left$(FeCap, Len(FeCap) - 1)

'In Abhängigkeit der Optionen die Ausgabe erstellen
If (FeCap <> vbNullString) And (Style = (WS_VISIBLE Or WS_BORDER)) Then
    PaFen = mHwnd
    Do
    PaFen = GetParent(PaFen)
    Loop Until PaFen = 0
    
    'TskId Id ermitteln
    RetWe = GetWindowThreadProcessId(mHwnd, TskId)
                
    If InStr(1, FeCap, SuStr, 1) > 0 Then
        AnwIn = TskId
    Else
        AnwIn = 0
    End If
End If

End Function
Public Function AnwLo(ByVal AnwID As Long) As Boolean
On Error Resume Next
'Entfernt die Anwendung aus dem Speicher

Dim hAnwe As Long

hAnwe = OpenProcess(PROCESS_TERMINATE, False, AnwID)

If hAnwe <> 0 Then
    AnwLo = TerminateProcess(hAnwe, 1&) <> 0
    CloseHandle hAnwe
End If
  
End Function

Public Function AnwNa(ByVal CoStr As String) As String
On Error Resume Next

Dim Posi1 As Long
  
Posi1 = InStr(CoStr, vbTab)

If Posi1 > 0 And Len(CoStr) > Posi1 Then AnwNa = Mid$(CoStr, Posi1 + 1)

End Function

Private Function AnwNT() As Boolean
On Error Resume Next

Dim OSInfo As OSVERSIONINFO

With OSInfo
    .dwOSVersionInfoSize = Len(OSInfo)
    .szCSDVersion = Space$(128)
    GetVersionEx OSInfo
    AnwNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT)
End With
  
End Function

Public Function AnwSu(Optional ByVal FeClo As Boolean) As Long
Attribute AnwSu.VB_Description = "Sucht eine bestimmte Anwendung, benötigt AnNam"
On Error Resume Next
'Sucht eine bestimmte Anwendung, benötigt AnNam

Set FM = frmMain

Dim mHwnd As Long
Dim RetWe As Long
Dim FeTit As String
    
mHwnd = GetWindow(FM.hwnd, GW_HWNDFIRST)

Do
RetWe = GetWindowTextLength(mHwnd) + 1
FeTit = Space$(RetWe)
RetWe = GetWindowText(mHwnd, FeTit, RetWe)
FeTit = Left$(FeTit, Len(FeTit) - 1)
If InStr(1, FeTit, AnNam, 1) > 0 Then
    AnwSu = mHwnd
    If FeClo = True Then
        SendMessage mHwnd, WM_CLOSE, ByVal 0&, ByVal 0&
    End If
    Exit Do
End If
mHwnd = GetWindow(mHwnd, GW_HWNDNEXT)
Loop Until mHwnd = 0
  
End Function
Public Sub AnwTe(ByVal SuStr As String)
On Error Resume Next
'Sucht und terminiert eine Bestimmte Anwendung

Dim mHwnd As Long
Dim TskFn As Long
Dim TskId As Long
Dim RetWe As Long

Set FM = frmMain

'Auch der Desktop ist ein Fenster
mHwnd = GetDesktopWindow
TskFn = AnwIn(mHwnd)

'Einstieg
mHwnd = GetWindow(FM.hwnd, GW_HWNDFIRST)

'Alle vorhandenen Fenster abklappern
Do
TskFn = AnwIn(mHwnd, SuStr)
If TskFn > 0 Then
    TskId = OpenProcess(PROCESS_TERMINATE, 0&, TskFn)
    RetWe = TerminateProcess(TskId, 1&)
    RetWe = CloseHandle(TskId)
End If
mHwnd = GetWindow(mHwnd, GW_HWNDNEXT)
Loop Until mHwnd = 0
    
End Sub
Private Function AnwTr(ByVal s As String) As String
On Error Resume Next

Dim Posi1 As Long
  
Posi1 = InStr(s, vbNullChar)
  
If Posi1 > 0 Then
    AnwTr = Left$(s, Posi1 - 1)
Else
    AnwTr = s
End If
   
End Function
