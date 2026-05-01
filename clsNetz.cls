VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHwnd As Long
Private mSend As String
Private mText As String

Public Event NetGefun(ByVal ResNam As String, ByVal ResTyp As String)

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type

Private Type NETRESOURCE_REAL
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    sLocalName As String
    sRemoteName As String
    sComment As String
    sProvider As String
End Type

Private Const RESDTYPE_DIRECTORY& = &H9
Private Const RESDTYPE_DOMAIN& = &H1
Private Const RESDTYPE_FILE& = &H4
Private Const RESDTYPE_GENERIC& = &H0
Private Const RESDTYPE_GROUP& = &H5
Private Const RESDTYPE_NETWORK& = &H6
Private Const RESDTYPE_ROOT& = &H7
Private Const RESDTYPE_SERVER& = &H2
Private Const RESDTYPE_SHARE& = &H3
Private Const RESDTYPE_SHAREADMIN& = &H8

Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&
Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000

Private Const NO_ERROR = 0
Private Const ERROR_MORE_DATA = 234
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF
Private Const MAX_COMPUTERNAME_LENGTH = 15
Private Const NERR_Success As Long = 0&

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function NetMessageBufferSend Lib "netapi32.dll" (lpServerName As Any, lpMsgName As Byte, lpFromName As Any, lpBuf As Byte, ByVal lnBufLen As Long) As Long
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
Private Declare Sub CopyMemByPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, ByVal lLen As Long)
Public Property Let FWnd(Flag As Long)
    mHwnd = Flag
End Property

Public Property Get FWnd() As Long
    FWnd = mHwnd
End Property

Public Property Let NeSnd(Flag As String)
Attribute NeSnd.VB_Description = "Empfänger der Nachricht"
    mSend = Flag
End Property

Public Property Get NeSnd() As String
    NeSnd = mSend
End Property

Public Function NetBen() As String
Attribute NetBen.VB_Description = "Gibt den Benutzernamen zurück"
On Error Resume Next

Dim StBuf As String * 25
Dim RetWe As Long
Dim UsNam As String

RetWe = GetUserName(StBuf, 25)
UsNam = Left(StBuf, InStr(StBuf, Chr(0)) - 1)

NetBen = UsNam

End Function

Public Property Let NeTex(Flag As String)
Attribute NeTex.VB_Description = "Nachrichtentext"
    mText = Flag
End Property

Public Property Get NeTex() As String
    NeTex = mText
End Property

Public Function NetMsg() As Long
On Error Resume Next

Dim WrittenTo() As Byte
Dim MSGBuffer() As Byte

WrittenTo = mSend & vbNullChar
MSGBuffer = mText & vbNullChar

NetMsg = (NetMessageBufferSend(ByVal 0&, WrittenTo(0), ByVal 0&, MSGBuffer(0), UBound(MSGBuffer)) = NERR_Success)

End Function

Public Function NetNam() As String
Attribute NetNam.VB_Description = "Gibt den Netzwerknamen wieder"
On Error Resume Next

Dim RetWe As Long
Dim ReLen As Long
Dim ReNam As String
Dim CName As String
  
ReLen = MAX_COMPUTERNAME_LENGTH + 1
ReNam = Space$(ReLen)
RetWe = GetComputerName(ReNam, ReLen)

If RetWe = 1 Then
    CName = Left$(ReNam, InStr(1, ReNam, Chr$(0)) - 1)
    NetNam = ReNam
Else
    NetNam = vbNullString
End If

End Function

Public Sub NetRec()
On Error Resume Next

Dim RetWe As Long
Dim hEnum As Long
Dim RsWer As Long
Dim MinWe As Long
Dim LenWe As Long
Dim AktZa As Long
Dim BufSi As Long
Dim LetId As Long
Dim ReTyp As String
Dim ReNam As String
Dim ErsRn As Boolean
Dim NetRe() As NETRESOURCE_REAL
Dim NetAp(0 To 256) As NETRESOURCE

ErsRn = True

Do
If ErsRn = True Then
    RetWe = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
        RESOURCEUSAGE_ALL, ByVal 0&, hEnum)
    ErsRn = False
Else
    If NetRe(LetId).dwUsage And RESOURCEUSAGE_CONTAINER Then
        RetWe = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
            RESOURCEUSAGE_ALL, NetRe(LetId), hEnum)
    Else
        RetWe = -1
        hEnum = 0
    End If
    LetId = LetId + 1
End If
If RetWe = NO_ERROR Then
    RsWer = RESOURCE_ENUM_ALL
    Do
    BufSi = UBound(NetAp) * Len(NetAp(0)) / 2
    RetWe = WNetEnumResource(hEnum, RsWer, NetAp(0), BufSi)
    If RsWer > 0 Then
        ReDim Preserve NetRe(0 To MinWe + RsWer - 1) As NETRESOURCE_REAL
        For AktZa = 0 To RsWer - 1
            NetRe(MinWe + AktZa).dwScope = NetAp(AktZa).dwScope
            NetRe(MinWe + AktZa).dwType = NetAp(AktZa).dwType
            NetRe(MinWe + AktZa).dwDisplayType = NetAp(AktZa).dwDisplayType
            NetRe(MinWe + AktZa).dwUsage = NetAp(AktZa).dwUsage
              
            If NetAp(AktZa).pLocalName Then
                LenWe = lstrlen(NetAp(AktZa).pLocalName)
                NetRe(MinWe + AktZa).sLocalName = Space$(LenWe)
                CopyMem ByVal NetRe(MinWe + AktZa).sLocalName, _
                    ByVal NetAp(AktZa).pLocalName, LenWe
            End If
              
            If NetAp(AktZa).pRemoteName Then
                LenWe = lstrlen(NetAp(AktZa).pRemoteName)
                NetRe(MinWe + AktZa).sRemoteName = Space$(LenWe)
                CopyMem ByVal NetRe(MinWe + AktZa).sRemoteName, _
                    ByVal NetAp(AktZa).pRemoteName, LenWe
            End If
                      
            If NetAp(AktZa).pComment Then
                LenWe = lstrlen(NetAp(AktZa).pComment)
                NetRe(MinWe + AktZa).sComment = Space$(LenWe)
                CopyMem ByVal NetRe(MinWe + AktZa).sComment, _
                    ByVal NetAp(AktZa).pComment, LenWe
            End If
              
            If NetAp(AktZa).pProvider Then
                LenWe = lstrlen(NetAp(AktZa).pProvider)
                NetRe(MinWe + AktZa).sProvider = Space$(LenWe)
                CopyMem ByVal NetRe(MinWe + AktZa).sProvider, _
                    ByVal NetAp(AktZa).pProvider, LenWe
            End If
        Next AktZa
    End If
    MinWe = MinWe + RsWer
    Loop While RetWe = ERROR_MORE_DATA
End If
If hEnum Then AktZa = WNetCloseEnum(hEnum)
Loop While LetId < MinWe

If Not (Not NetRe) Then
    For AktZa = 0 To UBound(NetRe)
        Select Case NetRe(AktZa).dwDisplayType
        Case RESDTYPE_DIRECTORY:    ReTyp = "Ordner"
        Case RESDTYPE_DOMAIN:       ReTyp = "Domäne"
        Case RESDTYPE_FILE:         ReTyp = "Datei"
        Case RESDTYPE_GENERIC:      ReTyp = "Generic"
        Case RESDTYPE_GROUP:        ReTyp = "Gruppe"
        Case RESDTYPE_NETWORK:      ReTyp = "Netzwerk"
        Case RESDTYPE_ROOT:         ReTyp = "Root"
        Case RESDTYPE_SERVER:       ReTyp = "Rechner"
        Case RESDTYPE_SHARE:        ReTyp = "Freigabe"
        Case RESDTYPE_SHAREADMIN:   ReTyp = "Adminfreigabe"
        End Select
        ReNam = NetRe(AktZa).sRemoteName & NetRe(AktZa).sComment
        RaiseEvent NetGefun(ReNam, ReTyp)
    Next AktZa
End If

End Sub
Public Sub NetVer(ByVal Flag As Integer)
Attribute NetVer.VB_Description = "Dialog Netzlaufwerk verbinden"
On Error Resume Next

Dim RetWe As Long

Select Case Flag
Case 1: RetWe = WNetConnectionDialog(mHwnd, RESOURCETYPE_DISK)
Case 2: RetWe = WNetDisconnectDialog(mHwnd, RESOURCETYPE_DISK)
End Select

End Sub
