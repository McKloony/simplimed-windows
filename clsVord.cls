VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHwnd As Long
Private mTitl As String

Private mStaZe As Date
Private mEndZe As Date
Private mSpZe1 As Date
Private mSpZe2 As Date
Private mSpZe3 As Date
Private mSpZe4 As Date
Private mSpEn1 As Boolean
Private mSpEn2 As Boolean

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5
Private Const GWL_STYLE = (-16)
Private Const WS_VISIBLE = &H10000000
Private Const WS_BORDER = &H800000
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type WINDOWPLACEMENT
    length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal wIndx As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Property Let FHwnd(ByVal Flag As Long)
Attribute FHwnd.VB_Description = "Fensterhandle der Anwendung"
    mHwnd = Flag
End Property

Public Property Get FHwnd() As Long
    FHwnd = mHwnd
End Property

Public Property Let FTitl(ByVal Flag As String)
Attribute FTitl.VB_Description = "Titel der Anwendung, die in den Vordergrund gesetzt werden soll"
    mTitl = Flag
End Property

Public Property Get FTitl() As String
    FTitl = mTitl
End Property

Private Sub VoAkt(hwnd As Long)
On Error Resume Next

Dim RetWe As Long
Dim WnPlc As WINDOWPLACEMENT
 
With WnPlc
    .length = Len(WnPlc)
    RetWe = GetWindowPlacement(hwnd, WnPlc)
    If RetWe Then
        If .showCmd = SW_SHOWMINIMIZED Then
            .flags = 0
            .showCmd = SW_SHOWNORMAL
            RetWe = SetWindowPlacement(hwnd, WnPlc)
        Else
            RetWe = SetForegroundWindow(hwnd)
            RetWe = BringWindowToTop(hwnd)
        End If
    End If
End With

End Sub

Private Function VoInf(ByVal hwnd As Long) As Boolean
On Error Resume Next
  
Dim RetWe As Long
Dim Style As Long
Dim Titel As String

VoInf = False
Style = GetWindowLong(hwnd, GWL_STYLE)
Style = Style And (WS_VISIBLE Or WS_BORDER)
RetWe = GetWindowTextLength(hwnd) + 1
Titel = Space$(RetWe)
RetWe = GetWindowText(hwnd, Titel, RetWe)
Titel = Left$(Titel, Len(Titel) - 1)

If (Titel <> vbNullString) And Style = (WS_VISIBLE Or WS_BORDER) Then
    If InStr(1, Titel, mTitl, 1) <> 0 Then
        VoInf = True
    End If
End If

End Function

Public Sub VoSet()
Attribute VoSet.VB_Description = "Setzt eine Anwendung in den Vordergrund (benötigt hWnd und Titel)"
On Error Resume Next
'Setzt eine Appllikation in den Vordergrund

Dim RetWe As Boolean
Dim hwnd As Long

hwnd = GetDesktopWindow
RetWe = VoInf(hwnd)
 
hwnd = GetWindow(mHwnd, GW_HWNDFIRST)
  
Do
RetWe = VoInf(hwnd)
If RetWe = True Then
    VoAkt hwnd
    Exit Do
End If
hwnd = GetWindow(hwnd, GW_HWNDNEXT)
Loop Until hwnd = 0

End Sub
