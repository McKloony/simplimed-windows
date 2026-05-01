VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private mRate As Byte   'Transparenzrate
Private mHwnd As Long   'Fensterhandle
Private mFeLin As Long  'Fenster X Pos
Private mFeObn As Long  'Fenster y Pos
Private mFeBre As Long  'Fensterbreite
Private mFeHoh As Long  'Fensterhöhe
Private mFeSta As Long  'Fensterstatus
Private RetWe As Long

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

Private Const SM_CXSCREEN = 0        ' Width of screen
Private Const SM_CYSCREEN = 1        ' Height of screen
Private Const SM_CXFULLSCREEN = 16   ' Width of window client area
Private Const SM_CYFULLSCREEN = 17   ' Height of window client area
Private Const SM_CYMENU = 15         ' Height of menu
Private Const SM_CYCAPTION = 4       ' Height of caption or title
Private Const SM_CXFRAME = 32        ' Width of window frame
Private Const SM_CYFRAME = 33        ' Height of window frame
Private Const SM_CXHSCROLL = 21      ' Width of arrow bitmap on horizontal scroll bar
Private Const SM_CYHSCROLL = 3       ' Height of arrow bitmap on horizontal scroll bar
Private Const SM_CXVSCROLL = 2       ' Width of arrow bitmap on vertical scroll bar
Private Const SM_CYVSCROLL = 20      ' Height of arrow bitmap on vertical scroll bar
Private Const SM_CXSIZE = 30         ' Width of bitmaps in title bar
Private Const SM_CYSIZE = 31         ' Height of bitmaps in title bar
Private Const SM_CXCURSOR = 13       ' Width of cursor
Private Const SM_CYCURSOR = 14       ' Height of cursor
Private Const SM_CXBORDER = 5        ' Width of window frame that cannot be sized
Private Const SM_CYBORDER = 6        ' Height of window frame that cannot be sized
Private Const SM_CXDOUBLECLICK = 36  ' Width of rectangle around the location of the first click. The second click must occur in the same rectangular location.
Private Const SM_CYDOUBLECLICK = 37  ' Height of rectangle around the location of the first click. The second click must occur in the same rectangular location.
Private Const SM_CXDLGFRAME = 7      ' Width of dialog frame window
Private Const SM_CYDLGFRAME = 8      ' Height of dialog frame window
Private Const SM_CXICON = 11         ' Width of icon
Private Const SM_CYICON = 12         ' Height of icon
Private Const SM_CXICONSPACING = 38  ' Width of rectangles the system uses to position tiled icons
Private Const SM_CYICONSPACING = 39  ' Height of rectangles the system uses to position tiled icons
Private Const SM_CXMIN = 28          ' Minimum width of window
Private Const SM_CYMIN = 29          ' Minimum height of window
Private Const SM_CXMINTRACK = 34     ' Minimum tracking width of window
Private Const SM_CYMINTRACK = 35     ' Minimum tracking height of window
Private Const SM_CXHTHUMB = 10       ' Width of scroll box (thumb) on horizontal scroll bar
Private Const SM_CYVTHUMB = 9        ' Width of scroll box (thumb) on vertical scroll bar
Private Const SM_DBCSENABLED = 42    ' Returns a non-zero if the current Windows version uses double-byte characters, otherwise returns zero
Private Const SM_DEBUG = 22          ' Returns non-zero if the Windows version is a debugging version
Private Const SM_MENUDROPALIGNMENT = 40 ' Alignment of pop-up menus. If zero,
Private Const SM_MOUSEPRESENT = 19   ' Non-zero if mouse hardware is installed
Private Const SM_PENWINDOWS = 41     ' Handle of Pen Windows dynamic link library if Pen Windows is installed
Private Const SM_SWAPBUTTON = 23     ' Non-zero if the left and right mouse buttons are swapped

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const ABM_GETTASKBARPOS = &H5
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const GWL_STYLE = -16
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const SC_SCREENSAVE = &HF140&
Private Const SC_CLOSE As Long = &HF060&
Private Const SPI_SETSCREENSAVEACTIVE = &H11
Private Const SPI_GETWORKAREA As Long = 48&
Private Const WM_SETREDRAW = &HB
Private Const WM_SYSCOMMAND = &H112&
Private Const WS_CAPTION = &HC00000
Private Const WS_EX_LAYERED = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const GWL_EXSTYLE = (-20)
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1
Private Const AC_SRC_NO_PREMULT_ALPHA = &H1
Private Const AC_SRC_NO_ALPHA = &H2
Private Const AC_DST_NO_PREMULT_ALPHA = &H10
Private Const AC_DST_NO_ALPHA = &H20
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal SPI_Action As Long, ByVal uiParam As Long, ByRef pvParam As Any, ByVal fWinIni As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendNotifyMessage Lib "user32.dll" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Sub Class_Terminate()

Set FM = Nothing

End Sub

Public Property Let FeBre(ByVal Flag As Long)
Attribute FeBre.VB_Description = "Fensterbreite"
    mFeBre = Flag
End Property

Public Property Get FeBre() As Long
    FeBre = mFeBre
End Property

Public Property Let FeHoh(ByVal Flag As Long)
Attribute FeHoh.VB_Description = "Fensterhöhe"
    mFeHoh = Flag
End Property

Public Property Get FeHoh() As Long
    FeHoh = mFeHoh
End Property

Public Property Let FeLin(ByVal Flag As Long)
Attribute FeLin.VB_Description = "Fenster X Pos"
    mFeLin = Flag
End Property

Public Property Get FeLin() As Long
    FeLin = mFeLin
End Property

Public Sub FenBli()
Attribute FenBli.VB_Description = "Läßt Titelleiste einmal blinken"
'Läßt Titelleiste einmal blinken

RetWe = FlashWindow(mHwnd, -1)

End Sub

Public Sub FenCap()
Attribute FenCap.VB_Description = "Entfernt die Titelzeile eines Formulars"
'Entfernt die Titelzeile eines Formulars

Dim AltWe As Long
Dim NeuWe As Long
Dim DifWe As Integer
Dim FeRec As RECT

AltWe = GetWindowLong(mHwnd, GWL_STYLE)

NeuWe = AltWe And Not WS_CAPTION
AltWe = SetWindowLong(mHwnd, GWL_STYLE, NeuWe)
DifWe = GetSystemMetrics(SM_CYCAPTION)

FeRec.Top = FeRec.Top - DifWe

End Sub

Public Sub FenCli()
Attribute FenCli.VB_Description = "Liefert aktuelle Koordinaten des Clientbereiches"
'Liefert aktuelle Koordinaten des Clientbereiches

Dim PaHwn As Long
Dim PaRec As RECT

PaHwn = GetParent(mHwnd)

RetWe = GetClientRect(PaHwn, PaRec)

With PaRec
    mFeLin = .Left
    mFeObn = .Top
    mFeBre = .Right - .Left
    mFeHoh = .Bottom - .Top
End With

End Sub

Public Sub FenDsk(Flag As Integer)
Attribute FenDsk.VB_Description = "Friert den gesamten Desktop ein 1:Desktop 2:Fenster 3:Auftauen"
On Error GoTo WiErr
'Friert den gesamten Desktop ein

Select Case Flag
Case 1: RetWe = LockWindowUpdate(GetDesktopWindow)
Case 2: RetWe = LockWindowUpdate(mHwnd)
Case 3: RetWe = LockWindowUpdate(0&)
End Select

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " FenDsk " & Err.Number
Resume Next

End Sub
Public Function FenFin(ByVal SuStr As String) As Long
On Error Resume Next
'Findt den wWnd eines fensters

RetWe = FindWindow(SuStr, vbNullString)

If RetWe <> 0 Then
    FenFin = RetWe
Else
     FenFin = 0
End If

End Function
Public Sub FenFix()
Attribute FenFix.VB_Description = "Deaktiviert Größenänderung bei Fenstern"
'Deaktiviert Größenänderung bei Fenstern

Dim SysMe As Long
      
SysMe = GetSystemMenu(mHwnd, False)

RetWe = RemoveMenu(SysMe, &HF010&, &H0&)
   
End Sub
Public Function FenGro(Flag As Integer) As Long
On Error Resume Next

Dim WBLin As Single
Dim WBObe As Single
Dim WBBre As Single
Dim WBHoh As Single
Dim FeRec As RECT
  
RetWe = SystemParametersInfo(SPI_GETWORKAREA, 0&, FeRec, 0&)

WBObe = FeRec.Top
WBLin = FeRec.Left
WBHoh = FeRec.Bottom - FeRec.Top
WBBre = FeRec.Right - FeRec.Left

Select Case Flag
Case 1: FenGro = CLng(WBLin)
Case 2: FenGro = CLng(WBObe)
Case 3: FenGro = CLng(WBBre)
Case 4: FenGro = CLng(WBHoh)
End Select

End Function
Public Sub FenInf(ByVal ApNam As String, ByVal FiNam As String, ByVal Versi As String)
Attribute FenInf.VB_Description = "Zeigt Infodialog für ein Programm"
'Zeigt Infodialog für ein Programm

RetWe = ShellAbout(mHwnd, ApNam, FiNam & vbCrLf & Versi, FM.Icon)

End Sub
Public Sub FenMax()
'Entfernt die Min- und Max Buttons
On Error Resume Next

Dim MenWe As Long

MenWe = GetSystemMenu(mHwnd, False)

Call RemoveMenu(MenWe, 4, MF_BYPOSITION)
Call RemoveMenu(MenWe, 3, MF_BYPOSITION)

End Sub

Public Function FenMet(Flag As Integer) As Long
Attribute FenMet.VB_Description = "1=Auflösun-X 2=Auflösung-Y 3=Clientbereich-X 4=Clientbereich-Y 5=Clientbereich ohne Taskbar 6=VertScroll-X 7=VertScroll-Y 8=HorzScroll-X 9=HorzScroll-Y"
'Gibt bestimmte Auflösungsparameter zurück

Dim TaPos As APPBARDATA

RetWe = SHAppBarMessage(ABM_GETTASKBARPOS, TaPos)

Select Case Flag
Case 1: FenMet = GetSystemMetrics(SM_CXSCREEN)
Case 2: FenMet = GetSystemMetrics(SM_CYSCREEN)
Case 3: FenMet = GetSystemMetrics(SM_CXFULLSCREEN)
Case 4: FenMet = GetSystemMetrics(SM_CYFULLSCREEN)
Case 5: FenMet = GetSystemMetrics(SM_CYSCREEN) - (TaPos.rc.Bottom - TaPos.rc.Top)
Case 6: FenMet = GetSystemMetrics(SM_CXVSCROLL)
Case 7: FenMet = GetSystemMetrics(SM_CYVSCROLL)
Case 8: FenMet = GetSystemMetrics(SM_CXHSCROLL)
Case 9: FenMet = GetSystemMetrics(SM_CYHSCROLL)
End Select

End Function
Public Sub FenMin()
'Entfernt die Min- und Max Buttons

Dim MenWe As Long

MenWe = GetWindowLong(mHwnd, GWL_STYLE)

RetWe = SetWindowLong(mHwnd, GWL_STYLE, MenWe And Not WS_MINIMIZEBOX And Not WS_MAXIMIZEBOX)

End Sub
Public Sub FenMov()
Attribute FenMov.VB_Description = "Plaziert das Fenster ein einer bestimmten Position"
'Plaziert das Fenster ein einer bestimmten Position

Select Case mFeSta
Case 0: If IsZoomed(mHwnd) = 1 Then ShowWindow mHwnd, SW_NORMAL
        If IsIconic(mHwnd) = 1 Then ShowWindow mHwnd, SW_NORMAL
        MoveWindow mHwnd, mFeLin, mFeObn, mFeBre, mFeHoh, True
Case 1: ShowWindow mHwnd, SW_MAXIMIZE
Case 2: ShowWindow mHwnd, SW_MINIMIZE
End Select

End Sub

Public Sub FenSav()
Attribute FenSav.VB_Description = "Liefert aktuelle Koordinaten des Fensters"
'Liefert aktuelle Koordinaten des Fensters

Dim FeRec As RECT

If IsIconic(mHwnd) = 1 Then
    mFeSta = 2
ElseIf IsZoomed(mHwnd) = 1 Then
    mFeSta = 1
Else
    mFeSta = 0
    RetWe = GetWindowRect(mHwnd, FeRec)
    With FeRec
        mFeLin = .Left
        mFeObn = .Top
        mFeBre = .Right - .Left
        mFeHoh = .Bottom - .Top
    End With
End If

End Sub

Public Sub FenSp1()
Attribute FenSp1.VB_Description = "Sperrt Bildschiraktualisierung"
'Sperrt Bildschiraktualisierung

RetWe = SendMessage(mHwnd, WM_SETREDRAW, False, 0&)

End Sub

Public Sub FenSp2()
Attribute FenSp2.VB_Description = "Hebt Aktualisierungssperre wieder auf"
'Hebt Aktualisierungssperre wieder auf

Dim ClRe As RECT

RetWe = SendMessage(mHwnd, WM_SETREDRAW, True, 0&)
RetWe = GetClientRect(mHwnd, ClRe)
RetWe = RedrawWindow(mHwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)

End Sub

Public Sub FenSrc(Flag As Boolean)
Attribute FenSrc.VB_Description = "Setzt den Status des Screensavers"
'Setzt den Status des Screensavers

If Flag Then
    RetWe = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, True, 0, 0)
Else
    RetWe = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, False, 0, 0)
End If

End Sub
Public Sub FenSys()
Attribute FenSys.VB_Description = "Entfernt das Systemmenü"
'Entfernt das Systemmenü

Dim MenWe As Long
    
RetWe = GetSystemMenu(mHwnd, 0)

If RetWe <> 0 Then
    MenWe = GetMenuItemCount(RetWe)
    If MenWe <> 0 Then
        RemoveMenu RetWe, MenWe - 1, MF_BYPOSITION Or MF_REMOVE
        RemoveMenu RetWe, MenWe - 2, MF_BYPOSITION Or MF_REMOVE
        DrawMenuBar mHwnd
    End If
End If

End Sub

Public Sub FenTit(ByVal FeStr As String)
On Error Resume Next
'Setzte Caption einex Fensters

RetWe = SetWindowText(mHwnd, FeStr)

End Sub
Public Sub FenTra()
On Error Resume Next
'Fenster transparent machen

RetWe = GetWindowLong(mHwnd, GWL_EXSTYLE)

If mRate < 255 Then
    RetWe = RetWe Or WS_EX_LAYERED
    SetWindowLong mHwnd, GWL_EXSTYLE, RetWe
    SetLayeredWindowAttributes mHwnd, 0, mRate, LWA_ALPHA
Else
    RetWe = RetWe Xor WS_EX_LAYERED
    SetWindowLong mHwnd, GWL_EXSTYLE, RetWe
End If

End Sub

Public Sub FenVor()
Attribute FenVor.VB_Description = "Setzt Fenster in den Fordergrund"
'Setzt Fenster in den Fordergrund
  
SetWindowPos mHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

End Sub
Public Sub FenZur()
Attribute FenZur.VB_Description = "Setzt das fenster in den Hintergrund"
'Setzt Fenster in den Hintergrund
  
SetWindowPos mHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

End Sub
Public Sub FenZus(Flag As Integer)
Attribute FenZus.VB_Description = "Setzt das Fenster in einen der drei Zustände"
'Setzt das Fenster in einen der drei Zustände

Select Case Flag
Case 1: ShowWindow mHwnd, SW_MAXIMIZE
Case 2: ShowWindow mHwnd, SW_MINIMIZE
Case 3: ShowWindow mHwnd, SW_NORMAL
Case 4: ShowWindow mHwnd, SW_NORMAL
        MoveWindow mHwnd, 0, 0, FenMet(3), FenMet(5), True
End Select

End Sub

Public Property Let FeObn(ByVal Flag As Long)
Attribute FeObn.VB_Description = "Fenster Y Pos"
    mFeObn = Flag
End Property

Public Property Get FeObn() As Long
    FeObn = mFeObn
End Property

Public Property Let FeSta(ByVal Flag As Long)
Attribute FeSta.VB_Description = "Fensterstatus"
    mFeSta = Flag
End Property

Public Property Get FeSta() As Long
    FeSta = mFeSta
End Property

Public Property Let FeTra(Flag As Byte)
Attribute FeTra.VB_Description = "Wert für Fenstertransparenz 0 bis 255"
    mRate = Flag
End Property

Public Property Get FeTra() As Byte
    FeTra = mRate
End Property

Public Property Get hwnd() As Long
    hwnd = mHwnd
End Property

Public Property Let hwnd(ByVal Flag As Long)
    mHwnd = Flag
End Property
