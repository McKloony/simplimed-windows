VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLizenz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CryKey As String 'Verschlüsselungskey
Private HarKey As String 'Hardwarekey
Private HardSe As String 'Seriennummer der Festplatte
Private SerNum As String 'Generierte Seriennummer
Private AnName As String 'Name der Anwendung

Private Const DRIVE_CDROM = 5           'CD-ROM-Laufwerk
Private Const DRIVE_FIXED = 3           'Festplatte
Private Const DRIVE_RAMDISK = 6         'RAM-Disklaufwerk
Private Const DRIVE_REMOTE = 4          'Netzlaufwerk
Private Const DRIVE_REMOVABLE = 2       'Wechselbar, z.B. Diskette oder ZIP-Drive

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Property Get AnwNa() As String
Attribute AnwNa.VB_Description = "Name der Anwendung"
    AnwNa = AnName
End Property

Public Property Let AnwNa(ByVal Flag As String)
    AnName = Flag
End Property

Public Property Get HdSer() As String
    HdSer = HardSe
End Property

Public Property Let HdSer(ByVal Flag As String)
    HardSe = Flag
End Property

Public Property Get KyHar() As String
    KyHar = HarKey
End Property

Public Property Let KyHar(ByVal Flag As String)
    HarKey = Flag
End Property

Public Property Get KySer() As String
    KySer = CryKey
End Property

Public Property Let KySer(ByVal Flag As String)
    CryKey = Flag
End Property

Private Function LiForSe(ByVal StrWe As String) As String
On Error Resume Next
'Formatiert die Seriennummer

StrWe = Right(String(16, "0") & StrWe, 16)

LiForSe = Mid$(StrWe, 1, 4) & Chr$(32) & Mid$(StrWe, 5, 4) & Chr$(32) & Mid$(StrWe, 9, 4) & Chr$(32) & Mid$(StrWe, 13, 4)
  
End Function

Public Sub LiGeSer()
On Error Resume Next
'Generiert eine gültige Seriennummer

Dim HarNr As String
Dim StrWe As String

HarNr = LiSeHar

If (Len(HarNr) = 0) Then
    HarNr = LiLo2St(LiSeZuf)
End If

StrWe = HarNr & HarNr
StrWe = LiVerkn(StrWe, AnName)
StrWe = LiVerkn(StrWe, CryKey)

SerNum = LiSt2He(StrWe)

End Sub

Public Sub LiHdSer()
On Error Resume Next
'Ermittelt die Seriennummer der Hardware

Dim StrWe As String

StrWe = LiSeHar

If (Len(StrWe) = 0) Then
    StrWe = LiLo2St(LiSeZuf)
End If

HardSe = LiSt2He(LiVerkn(StrWe, HarKey))

End Sub

Private Function LiHe2Lo(ByVal HexWe As String) As Long
On Error Resume Next
'Umwandeln Hex in Long

LiHe2Lo = Val("&h" & HexWe)

End Function

Private Function LiHe2St(ByRef HexWe As String) As String
On Error Resume Next
'Umwandeln Hex in String

Dim AktZa As Long
Dim Lange As Long

Lange = Len(HexWe)

For AktZa = 1 To Lange - 1 Step 2
    LiHe2St = LiHe2St & Chr$(Val("&H" & Mid$(HexWe, AktZa, 2)))
Next AktZa

End Function

Private Function LiLo2He(ByVal LoWer As Long) As String
On Error Resume Next
'umwandeln Long in Hex

LiLo2He = UCase(Right(String(7, 48) & Hex(LoWer), 8))

End Function

Private Function LiLo2St(ByVal LoWer As Long) As String
On Error Resume Next
'Umwandeln Long in String

LiLo2St = Chr$(((LoWer And &HFF000000) \ &H1000000) And &HFF)
LiLo2St = LiLo2St & Chr$((LoWer And &HFF0000) \ &H10000)
LiLo2St = LiLo2St & Chr$((LoWer And &HFF00&) \ &H100)
LiLo2St = LiLo2St & Chr$(LoWer And &HFF&)

End Function

Private Function LiSeHar() As String
On Error Resume Next
'Liefert Seriennummer der Hardware
  
Dim AktZa As Long

AktZa = LiVoSer

LiSeHar = IIf(AktZa = 0, "", LiLo2St(AktZa))

End Function

Public Sub LiSeria()
On Error Resume Next
'Generiert eine gültige Seriennummer

Dim HarNr As String
Dim StrWe As String

StrWe = LiLo2St(LiHe2Lo(HardSe))
StrWe = LiVerkn(StrWe, HarKey)

StrWe = StrWe & StrWe
StrWe = LiVerkn(StrWe, AnName)
StrWe = LiVerkn(StrWe, CryKey)

SerNum = LiSt2He(StrWe)

End Sub

Private Function LiSeZuf() As Long
On Error Resume Next
'Generiert eine ZufallsSeriennummer

LiSeZuf = CLng(Rnd() * &H7FFFFFFF)

End Function

Private Function LiSt2He(ByVal StWer As String) As String
On Error Resume Next
'Umwandeln String in Hex

Dim AktZa As Long
Dim Lange As Long

Lange = Len(StWer)

LiSt2He = Space(2 * Lange)
For AktZa = 0 To Lange - 1
    Mid$(LiSt2He, 1 + AktZa * 2) = Right("0" & Hex(Asc(Mid$(StWer, AktZa + 1, 1))), 2)
Next

End Function

Private Function LiVerkn(ByVal StrWe As String, ByVal KeyWe As String) As String
On Error Resume Next
'Verknüpft zwei Strings gleicher Länge

Dim AktZa As Long
Dim Lange As Long

Lange = Len(StrWe)
LiVerkn = Space$(Lange)

For AktZa = 1 To Lange
    Mid$(LiVerkn, AktZa, 1) = Chr$(Asc(Mid$(StrWe, AktZa, 1)) Xor Asc(Mid$(KeyWe, AktZa, 1)))
Next

End Function

Private Function LiVoSer() As Long
On Error Resume Next
'Seriendaten der Festplatte

Dim VolumeNameBuffer As String
Dim VolumeNameSize As Long
Dim VolumeSerialNumber As Long
Dim MaximumComponentLength As Long
Dim FileSystemFlags As Long
Dim FileSystemNameBuffer As String
Dim FileSystemNameSize As Long

Dim AktZa As Long
Dim DrvTy As Long
Dim RetWe As Long
Dim StrSt As Long
Dim AkDrv As String
Dim ReStr As String
Dim DrStr As String * 150

StrSt = 1
RetWe = GetLogicalDriveStrings(150, DrStr)

If RetWe <> 0 Then
    AkDrv = Mid$(DrStr, StrSt, 3)
    Do
    DrvTy = GetDriveType(AkDrv)
    If DrvTy = DRIVE_FIXED Then
        If Not AkDrv = "c:\" Then
            ReStr = AkDrv
        Else
            ReStr = vbNullString
        End If
    Else
        ReStr = vbNullString
    End If
    If ReStr <> vbNullString Then
        Exit Do
    End If
    StrSt = StrSt + 4
    AkDrv = Mid$(DrStr, StrSt, 3)
    Loop While (Mid$(DrStr, StrSt, 1) <> vbNullChar)
Else
    ReStr = "C:\"
End If

GetVolumeInformation ReStr, VolumeNameBuffer, VolumeNameSize, VolumeSerialNumber, MaximumComponentLength, FileSystemFlags, FileSystemNameBuffer, FileSystemNameSize

LiVoSer = VolumeSerialNumber

End Function
Public Property Get SerNu() As String
    SerNu = SerNum
End Property

Public Property Let SerNu(ByVal Flag As String)
    SerNum = Flag
End Property
