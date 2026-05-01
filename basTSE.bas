Attribute VB_Name = "basTSE"
Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, ByRef SIZE) As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Const FILE_DEVICE_FILE_SYSTEM = &H9&
Private Const FILE_ANY_ACCESS = 0
Private Const FILE_READ_ACCESS = &H1
Private Const FILE_WRITE_ACCESS = &H2

Private Const NePIN = "12345"
Private Const NePUK = "123456"
Private Const NeADM = "98765"

Private Const bytespersector = 512

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const INVALID_HANDLE_VALUE = -1&

Private Const FILE_BEGIN = 0

Public RuAry() As Byte

Private ChGlb As Object
Private ChTar As Object
Private ChCry As Object
Private ChBin As Object
'Private ChGlb As New ChilkatGlobal
'Private ChTar As New ChTar

Public TrCou As String
Public SiCou As String
Public TSSig As String

Dim ErgSt As String
Dim ErgOK As String

Public Type Command
    blen(0 To 1) As Byte
    befehl(0 To 509) As Byte
End Type

Public befehl As Command

Private Type pin
    blen(1 To 2) As Byte
    befehl(1 To 2) As Byte
    user_id As Byte
    cpinlen As Byte
    cpin(1 To 5) As Byte
    newpinlen As Byte
    newpin(1 To 5) As Byte
End Type

Private Type puk
    blen(1 To 2) As Byte
    befehl(1 To 2) As Byte
    cpuklen As Byte
    cpuk(1 To 6) As Byte
    newpuklen As Byte
    newpuk(1 To 6) As Byte
End Type

Private Type reg_client
    blen(1 To 2) As Byte
    befehl(1 To 2) As Byte
    len_id As Byte
    id(1 To 251) As Byte
End Type

Private clFil As clsFile
Public Function get_status()
On Error Resume Next

Dim RetWe As Long
Dim TmpSt As String
Dim TmStr As String
Dim StStr As String
Dim SerNr As String
Dim SigNa As String
Dim DaStr As String
Dim AktZa As Integer

RetWe = DirectReadDriveNT(GLTSL & "TSE_INFO.DAT", 0, 0, RuAry(), 512)

If RetWe = 0 Then
    get_status = "Fehler, TSE nicht ansprechbar!"
    Exit Function
End If

SerNr = bytehex(RuAry, 256, 32)

RetWe = RuAry(106)
  
For AktZa = 0 To RetWe - 1
    TmStr = TmStr & Chr(RuAry(107 + AktZa))
Next

SigNa = StrToBase64(TmStr)

TmpSt = bytehex(RuAry, 64, 8)
TmpSt = CLng("&H" & Mid(TmpSt, 1, 16))
DaStr = DateAdd("s", Val(TmpSt), DateSerial(1970, 1, 1))
  
If RuAry(29) = 0 Then
    StStr = StStr & "TSE nicht initialisiert!"
End If

If RuAry(29) = 1 Then
    StStr = StStr & "TSE initialisiert!"
End If
 
If (RuAry(28) And 2) = 2 Then
  TmpSt = bytehex(RuAry, 36, 4)
  TmpSt = CLng("&H" & Mid(TmpSt, 1, 8))
  TmStr = Val(TmpSt / 3600) & ":" & Format(Val(TmpSt Mod 3600) / 60, "0")
  StStr = StStr & vbCrLf & "Selbsttest gültig noch : " & TmStr
Else
  StStr = StStr & vbCrLf & "Selbsttest fehlt!"
End If

TmStr = bytehex(RuAry, 72, 8)
TmpSt = CLng("&H" & Mid(TmStr, 1, 16))

If TmpSt <> "0" Then
    StStr = StStr & vbCrLf & "Speicher belegt: " & TmpSt / 1024 & " KByte " & TmpSt / 512 & " Sectoren"
End If

get_status = Format(DaStr, "dd.mm.yyyy") & ";" & SerNr & ";" & SigNa & ";" & StStr & ";" & TmpSt / 512
    
End Function
Public Function TSE_Login(AdMin As Integer, TSPin As Integer) As String
On Error Resume Next

Dim AktZa As Integer

befehl.blen(0) = 0
befehl.blen(1) = 9
befehl.befehl(0) = &H20
befehl.befehl(1) = &H0

If AdMin = 1 Then
    befehl.befehl(2) = 2
Else
    befehl.befehl(2) = 1
End If

befehl.befehl(3) = 5 'PIN länge immer 5

For AktZa = 0 To 4
    befehl.befehl(AktZa + 4) = Asc(Mid(TSPin, AktZa + 1, 1))
Next

TSE_Send
DoEvents
WindowSleep 1000

TSE_Wart
DoEvents

If Left$(ErgOK, 2) = "00" Then
    TSE_Login = "erfolgreich"
Else
    TSE_Login = "fehlgeschlagen!"
End If

End Function
Private Function TSE_LogNe(pin, AdMin)

Dim AktZa As Integer

befehl.blen(0) = 0
befehl.blen(1) = 9
befehl.befehl(0) = &H20
befehl.befehl(1) = &H0

If AdMin = 1 Then
    befehl.befehl(2) = 2   ' als Admin  2 = timelogin
Else
    befehl.befehl(2) = 1  ' als Admin  2 = timelogin
End If

befehl.befehl(3) = 5     ' Pin länge immer 5

For AktZa = 0 To 4
    befehl.befehl(AktZa + 4) = Asc(Mid(pin, AktZa + 1, 1))
Next

TSE_Send
WindowSleep 1000
TSE_Wart

TSE_LogNe = ErgOK & " Login mit " & pin

End Function
Private Function TSE_Init()
On Error Resume Next

befehl.blen(0) = 0
befehl.blen(1) = 2
befehl.befehl(0) = &H70
befehl.befehl(1) = &H0
TSE_Send
DoEvents
TSE_Init = TSE_Wart
DoEvents

End Function
Private Function TSE_PIN(seed, serial)
On Error GoTo fehler

Dim VarWe As Variant
Dim HexWe As String
Dim AktZa As Integer
Dim puk As String
Dim pin As String
Dim tpin As String
Dim RetWe As Long
Dim AryWe(4) As String

Set ChGlb = CreateObject("Chilkat_9_5_0.Global")
Set ChCry = CreateObject("Chilkat_9_5_0.Crypt2")
Set ChBin = CreateObject("Chilkat_9_5_0.BinData")

RetWe = ChGlb.UnlockBundle("GNTRSC.CB11217_aAaet8Bij08o")

If (RetWe <> 1) Then
    MsgBox ChGlb.LastErrorText, , "DLL fehlt oder ist beschädigt!"
    'End
End If

ChCry.HashAlgorithm = "sha256"
ChCry.Charset = "hex"
ChCry.EncodingMode = "hex"

For AktZa = 0 To Len(seed) - 1
    HexWe = HexWe & Hex(Asc(Mid(seed, AktZa + 1, 1)))
Next

VarWe = HexWe & serial
AktZa = Len(VarWe)
AktZa = Len(HexWe)

AryWe(0) = ChCry.HashStringENC(VarWe)

' die ersten 24 Byte in 8ter Gruppen
AryWe(1) = Mid(AryWe(0), 1, 16)               ' die ersten 8 Byte
AryWe(2) = Mid(AryWe(0), 17, 16)              ' die zweiten 8 Byte
AryWe(3) = Mid(AryWe(0), 33, 16)              ' die dritten 8 Byte

' Kontrollausgabe
'akt_status = VarWe & vbCrLf
'akt_status = akt_status & AryWe(1) & vbCrLf
'akt_status = akt_status & AryWe(2) & vbCrLf
'akt_status = akt_status & AryWe(3) & vbCrLf

ChBin.Clear
RetWe = ChBin.AppendEncoded(AryWe(1), "hex")
VarWe = ChBin.GetEncoded("decimal")
puk = VarWe

ChBin.Clear
RetWe = ChBin.AppendEncoded(AryWe(2), "hex")
VarWe = ChBin.GetEncoded("decimal")
pin = VarWe

ChBin.Clear
RetWe = ChBin.AppendEncoded(AryWe(3), "hex")
tpin = ChBin.GetEncoded("decimal")

puk = Mid(puk, Len(puk) - 5)
pin = Mid(pin, Len(pin) - 4)
tpin = Mid(tpin, Len(tpin) - 4)
TSE_PIN = puk & ";" & pin & ";" & tpin

'akt_status = akt_status & "Pin: " & pin & vbCrLf
'akt_status = akt_status & "Puk: " & puk & vbCrLf
'akt_status = akt_status & "TimePin " & Timepin & vbCrLf
'akt_status = akt_status & HexWe
Exit Function
fehler:
TSE_PIN = "fehler beim pinpuk berechnen!!"

Exit Function

End Function
Private Function TSE_PUK(AlPIN, inewpin)

Dim newpuk As puk
Dim AktZa As Integer
Dim intdatnum As Integer
Dim ErStr As String

If Len(AlPIN) <> Len(inewpin) Then
    TSE_PUK = "Fehler Pin/Puk länge stimmt nicht!"
    MsgBox "hier problem", , AlPIN & Space$(1) & inewpin
    End
    Exit Function
End If

newpuk.befehl(1) = &H23
newpuk.befehl(2) = &H0

newpuk.cpuklen = 6
For AktZa = 1 To 6
    newpuk.cpuk(AktZa) = Asc(Mid(AlPIN, AktZa, 1))
Next

newpuk.newpuklen = 6

For AktZa = 1 To 6
    newpuk.newpuk(AktZa) = Asc(Mid(inewpin, AktZa, 1))
Next

newpuk.blen(1) = 0
newpuk.blen(2) = 16
intdatnum = FreeFile

Open GLTSL & "\TSE_COMM.DAT" For Random Access Write As intdatnum Len = Len(newpuk)
Put #intdatnum, 1, newpuk
Close intdatnum
WindowSleep 1000

ErStr = TSE_Wart

TSE_PUK = "change Puk " & inewpin & " Resultat: " & ErStr

End Function
Public Function TSE_Neue() As String
On Error Resume Next

Dim RetWe As Long
Dim dpuk As String
Dim dpin As String
Dim dtpin As String
Dim ErStr As String
Dim TsAry() As String

Set ChGlb = CreateObject("Chilkat_9_5_0.Global")

RetWe = ChGlb.UnlockBundle("GNTRSC.CB11217_aAaet8Bij08o")

If (RetWe <> 1) Then
    MsgBox ChGlb.LastErrorText, , "DLL fehlt oder ist beschädigt!"
End If

TSEZeig ("TSE Selbsttest : Swissbit")

ErStr = TSE_SeTe("SwissbitSwissbit")

TSEZeig ("Selbsttest läuft...")

ErStr = TSE_PIN("SwissbitSwissbit", GlTSN)
If InStr(1, ErStr, "fehler") Then
    TSE_Neue = ErStr
    MsgBox ErStr, , "hier raus"
    Exit Function
End If

TsAry = Split(ErStr, ";")
dpuk = TsAry(0)
dpin = TsAry(1)
dtpin = TsAry(2)

TSEZeig ("PIN errechnet : " & dpuk & dpin & dtpin)

ErStr = TSE_PUK(dpuk, NePUK)
If InStr(1, ErStr, "fehler") Then
    TSE_Neue = ErStr & " PUK Änderung! Achtung geht nur 3 mal dann TSE hinüber!"
    Exit Function
End If

TSEZeig ("Neuer PUK : " & NePUK)

ErStr = TSE_LogNe(dpin, 0)

If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & "Login PIN"
    Exit Function
End If

TSEZeig ("Login mit TSE PIN : " & dpin)

ErStr = TSE_And(dpin, NePIN, 0)
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " Ändere PIN"
    Exit Function
End If

TSEZeig ("Neuer PIN : " & NeADM)

ErStr = TSE_LogNe(dtpin, 1)                  ' login mit errechneten TimaAdin Pin
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & "Login PIN"
    Exit Function
End If

TSEZeig ("Login mit TSE AdminPIN : " & dtpin)

ErStr = TSE_And(dtpin, NePIN, 1)     ' timeadmin pin auf 98765 ändern
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " Login Time Admin Pin"
    Exit Function
End If

TSEZeig ("neuer AdminPIN : " & NePIN)

ErStr = TSE_LogNe(NePIN, 1)                  ' login mit neuem TimaAdin Pin
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " Login Time Admin Pin"
    Exit Function
End If

TSEZeig ("Login mit AdminPIN : " & NePIN)

ErStr = TSE_Time
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " beim Zeit sezten"
    Exit Function
End If

TSEZeig ("Zeitsynchronisation : " & Now)

ErStr = TSE_Clie(GlTSN)
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " neuer Client eintragen"
    Exit Function
End If

TSEZeig ("Neuer Klient (Kassenname) : " & GlTSN)

ErStr = TSE_SeTe(GlTSN)
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " selbsttest mit neuem Clienten-> " & GlTSN
    Exit Function
End If

TSEZeig ("Selbsttest mit neuem Kassenname : " & GlTSN)

ErStr = TSE_CTSS()
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " enable CTSS "
    Exit Function
End If

TSEZeig ("CTSS aktivieren...")

ErStr = TSE_Init
If InStr(1, ErStr, "fehler", vbTextCompare) Then
    TSE_Neue = ErStr & " Initialisiert in TSE eintragen "
    Exit Function
End If

TSEZeig ("TSE als initialisiert kennzeichnen...")

TSEZeig ("Die TSE wurde erfolgreich eingerichtet")

TSE_Neue = "OK"

End Function
Public Sub TSE_Send()
On Error Resume Next

Dim intdatnum As Integer
Dim dateiname As String
intdatnum = FreeFile

Set clFil = New clsFile

dateiname = GLTSL & "TSE_COMM.DAT"

If clFil.FilVor(dateiname) = True Then
    Open dateiname For Binary As intdatnum
    Put #intdatnum, 1, befehl
    Close intdatnum
End If

Set clFil = Nothing

End Sub

Public Function TSE_SeTe(ByVal CliNa As String) As String
On Error Resume Next
'Selbsttest
  
Dim ErgTe As String
Dim AktZa As Integer

befehl.blen(0) = 0
befehl.blen(1) = Len(CliNa) + 3
befehl.befehl(0) = &H40
befehl.befehl(1) = &H0
befehl.befehl(2) = Len(CliNa)

For AktZa = 0 To Len(CliNa) - 1
    befehl.befehl(AktZa + 3) = Asc(Mid(CliNa, AktZa + 1, 1))
Next

TSE_Send
DoEvents

ErgTe = TSE_Wart()

If Val(ErgTe) > 0 Then
    TSE_SeTe = "TSE Selbsttest fehlgeschlagen!"
Else
    TSE_SeTe = "TSE Selbsttest erfolgreich."
End If

End Function
Private Function TSE_Sign(PayLo As String, TyStr As String, NuSta As Boolean, KaStr As String) As String
On Error Resume Next

Dim TmStr As String
Dim Test() As String

TSE_Tran 0, 0, 0, KaStr, TyStr
TSE_Send
DoEvents

TmStr = TSE_Wart()

If Val(TmStr) > 0 Then
    TSE_Sign = "Fehler Transaktion Start " & TmStr
    Exit Function
End If

TmStr = TSE_Fini()
DoEvents

If NuSta = True Then
    TSE_Sign = TmStr
    Exit Function
End If

Test = Split(TmStr, ";")
DoEvents

If UBound(Test) < 2 Then
    TSE_Sign = "Fehler " & TmStr
    Exit Function
End If

If Val(Test(1)) = 0 Then
    TSE_Sign = "Fehler"
    Exit Function
End If

TSE_Tran 2, Test(0), Len(PayLo), KaStr, TyStr
DoEvents

TSE_Send
DoEvents

TmStr = TSE_Wart()
DoEvents

TSE_Rech PayLo
DoEvents

TmStr = TSE_Fini()
DoEvents

TSE_Sign = TmStr

End Function
Public Function TSE_Strg(ByVal KliNa As String) As Boolean
On Error GoTo MeErr

Dim RetWe As Long
Dim DaDif As Long
Dim TrZah As Long
Dim SgZah As Long
Dim StTSE As String
Dim TmStr As String
Dim ZeiSt As String
Dim ZeiEn As String
Dim ZeSta As String
Dim ZeEnd As String
Dim TSign As String
Dim QrStr As String
Dim StaWe As Integer
Dim AktZa As Integer
Dim ZeSyn As Boolean
Dim TsAry() As String

DaDif = DateDiff("s", CDate("01.01.1970 00:00:00"), Now)
ZeiSt = DateAdd("s", DaDif, DateSerial(1970, 1, 1))
ZeSta = Format(ZeiSt, "yyyy-mm-DDThh:mm:ss.000Z")

RetWe = DirectReadDriveNT(GLTSL & "TSE_INFO.DAT", 0, 0, RuAry(), 512)

If RetWe = 0 Then
    SPopu "TSE nicht ansprechbar!", "Die TSE kann nicht gefunden werden oder es liegt eine Störung vor.", IC48_Forbidden
    frmTSEInit.Show vbModal
Else
    RetWe = DirectReadDriveNT(GLTSL & "TSE_INFO.DAT", 0, 0, RuAry(), 512)
    StaWe = RuAry(28)
    If (StaWe And 2) <> 2 Then
        SPopu "Fehlender Selbsttest!", "Es wurde noch kein TSE Selbsttest durchgeführt.", IC48_Forbidden
        frmTSEInit.Show vbModal
    Else
        RetWe = DirectReadDriveNT(GLTSL & "TSE_INFO.DAT", 0, 0, RuAry(), 512)
        StaWe = RuAry(28)
        If (StaWe And 2) <> 1 Then
            If TSE_Time() = "erfolgreich" Then
                ZeSyn = True
            Else
                SPopu "TSE Zeitsynchronisation", "Die TSE Zeitsynchronisation ist fehlgeschlagen!", IC48_Forbidden
            End If
        End If
        If ZeSyn = True Then
            If GlTSB.BeTyp <> vbNullString Then
                TmStr = GlTSB.BeTyp & "^" & Format$(GlTSB.BeSt0, GlWa1) & "_" & Format$(GlTSB.BeSt1, GlWa1) & "_" & Format$(GlTSB.BeSt2, GlWa1) & "_" & Format$(GlTSB.BeSt4, GlWa1) & "^" & Format$(GlTSB.BeBar, GlWa1) & ":Bar_" & Format$(GlTSB.BeUnb, GlWa1) & ":Unbar"
                TmStr = Replace(TmStr, ",", ".")
                StTSE = TSE_Sign(TmStr, GlTSB.BeTyp, False, KliNa)
                DoEvents
                If StTSE <> vbNullString Then
                    TsAry = Split(StTSE, ";")
                    TSign = CStr(TsAry(2))
                    TrZah = Val(TsAry(0))
                    SgZah = Val(TsAry(1))
                    ZeiEn = CStr(TsAry(3))
                    ZeEnd = Format(ZeiEn, "yyyy-mm-DDThh:mm:ss.000Z")
                    QrStr = "V0" & Chr$(59) & KliNa & Chr$(59) & "Kassenbeleg-V1" & Chr$(59) & TmStr & Chr$(59) & TrZah & Chr$(59) & SgZah & Chr$(59) & ZeSta & Chr$(59) & ZeEnd & Chr$(59) & "ecdsa-plain-SHA384" & Chr$(59) & "unixTime" & Chr$(59) & TSign & Chr$(59) & GlTSK
                    With GlTSB
                        .TraZe = TrZah
                        .SigZe = SgZah
                        .ZeiEn = ZeiEn
                        .ZeiSt = ZeiSt
                        .ZeLog = ZeiSt & " - " & ZeiEn
                        .SigSt = TSign
                        .SigQr = QrStr
                    End With
                    TSE_Strg = True
                End If
            End If
        End If
    End If
End If

Exit Function

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSE_Strg " & Err.Number
Resume Next

End Function
Public Function TSE_TAR(FiNam) As Long
On Error GoTo MeErr

Dim AktZa As Long
Dim GesZa As Long
Dim RetWe As Long
Dim TarNa As String
Dim TaStr As String
Dim DaNam As String
Dim ErStr As String
Dim NeZal As Integer
Dim InDat As Integer

ErStr = TSE_Track(0)

If ErStr = "0" Then
    TSE_TAR = "Fehler TSE gibt keine Daten aus (Selbsttest?)"
    Exit Function
End If

TarNa = Trim(ErStr)
ErStr = TSE_Track(1)
TaStr = ErStr

DaNam = FiNam & "/" & Trim(TarNa)

InDat = FreeFile
Open DaNam For Output As InDat
Print #InDat, TaStr
Close InDat

ErStr = TSE_Track(2)
TarNa = Trim(ErStr)
TaStr = vbNullString
NeZal = 3

For AktZa = 0 To 6
    ErStr = TSE_Track(AktZa + NeZal)
    TaStr = TaStr & ErStr
Next

InDat = FreeFile
DaNam = FiNam & "/" & TarNa
Open DaNam For Output As InDat
Print #InDat, TaStr
Close InDat

NeZal = 10
For AktZa = 0 To 100000
    ErStr = TSE_Track(NeZal + AktZa)
    If ErStr = "0" Then Exit For
        If InStr(1, ErStr, "_Log-") Then
            TarNa = Trim(ErStr)
            TaStr = vbNullString
            ErStr = TSE_Track(AktZa + NeZal + 1)
            TaStr = ErStr
    
            InDat = FreeFile
            DaNam = FiNam & "/" & TarNa
            Open DaNam For Output As InDat
            Print #InDat, TaStr
            Close InDat
            GesZa = GesZa + 1
        End If
Next

Set ChGlb = CreateObject("Chilkat_9_5_0.Global")
Set ChTar = CreateObject("Chilkat_9_5_0.Tar")

RetWe = ChGlb.UnlockBundle("GNTRSC.CB11217_aAaet8Bij08o")

If (RetWe <> 1) Then
    MsgBox ChGlb.LastErrorText, , "DLL fehlt oder ist beschädigt!"
End

End If

ChTar.WriteFormat = "gnu"

RetWe = ChTar.AddDirRoot(FiNam)

If (RetWe <> 1) Then
    MsgBox ChTar.LastErrorText, , "Fehler, Ordner nicht ok"
    Exit Function
End If

RetWe = ChTar.WriteTar(Mid(FiNam, 1, 2) & "\export.tar")

If (RetWe <> 1) Then
    MsgBox ChTar.LastErrorText, , "Fehler beim schreiben der TAR Exportdateien"
    Exit Function
End If

TSE_TAR = GesZa

Exit Function

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaMain " & Err.Number
Resume Next

End Function
Public Function TSE_Time() As String
On Error Resume Next

Dim AktZa As Integer
Dim nsek As Long
Dim ZeStr As String

nsek = DateDiff("s", CDate("01.01.1970 00:00:00"), Now)
ZeStr = "00000000" & Hex$(nsek)
befehl.blen(0) = 0
befehl.blen(1) = 10
befehl.befehl(0) = &H80
befehl.befehl(1) = &H0

For AktZa = 0 To 7
    befehl.befehl(AktZa + 2) = CByte("&H" & Mid(ZeStr, AktZa * 2 + 1, 2)) 'Hex String in Bin Byte
Next

TSE_Send
DoEvents

TSE_Wart
DoEvents

If Left$(ErgOK, 2) = "00" Then
    TSE_Time = "erfolgreich"
Else
    TSE_Time = "fehlgeschlagen!"
End If
  
End Function
Private Function TSE_Track(TrkNr) As String
On Error GoTo fehler

Dim RetWe As Long
Dim ReStr As String
Dim Lange As Integer

RetWe = DirectReadDriveNT(GLTSL & "tse_tar.001", Val(TrkNr), 0, RuAry(), 512)

If RuAry(0) = 0 Or RetWe = 0 Then
    TSE_Track = "0"
    Exit Function
End If

ReStr = StrConv(RuAry, vbUnicode)  '  hört beim ersten 0 auf! Stringende Byte in String

Lange = Len(Chr$(0))
While Len(ReStr) > 0 And Right$(ReStr, Lange) = Chr$(0)
    ReStr = Left$(ReStr, Len(ReStr) - Lange)
Wend

TSE_Track = ReStr

Exit Function

fehler:
TSE_Track = "Fehler beim TSE Trak Nr: " & Format$(TrkNr, "000000")

End Function
Private Sub TSE_Rech(payload)
On Error Resume Next

Dim intdatnum As Integer
Dim processData()  As Byte

processData = StrConv(payload, vbFromUnicode)
intdatnum = FreeFile
Open GLTSL & "TSE_TAR.001" For Binary As #intdatnum
Put #intdatnum, 1, processData
Close intdatnum
  
End Sub
Public Function TSE_Wart() As String
On Error Resume Next

Dim RetWe As Long
Dim AktZa As Integer

ErgSt = vbNullString

nochmals:
  
RetWe = DirectReadDriveNT(GLTSL & "TSE_COMM.DAT", 0, 0, RuAry(), 512)
If RetWe < 10 Then
    'ErgSt = "TSE ausgefallen!"
    TSE_Wart = "Fehler TSE ausgefallen oder nicht vorhanden!"
    Exit Function
End If

If RuAry(4) = 255 Then
    ErgOK = Hex(RuAry(7)) & Hex(RuAry(8))
    For AktZa = 7 To 200
        If RuAry(AktZa) > 15 Then
            ErgSt = ErgSt & Hex(RuAry(AktZa)) & " "
        Else
            ErgSt = ErgSt & "0" & Hex(RuAry(AktZa)) & " "
        End If
    Next
    TSE_Wart = Mid(ErgSt, 1, 5)
    Exit Function
Else
    If RuAry(4) = 253 Then     ' FD so muss RuAry abgeholt werden
        befehl.blen(1) = 2
        befehl.befehl(0) = &H83
        befehl.befehl(1) = &H0
        TSE_Send
   End If
End If

GoTo nochmals

End Function
Public Function DirectReadDriveNT(ByVal sdrive As String, iStartSec As Long, ByVal iOffset As Long, ByRef lpBuffer() As Byte, ByVal cBytes As Long) As Long
On Error Resume Next

Dim hDevice As Long
Dim abBuff() As Byte
Dim nSectors As Integer

nSectors = Int((iOffset + cBytes - 1) / bytespersector) + 1
Rem hDevice = CreateFile("\\.\" & UCase(Left(sDrive, 1)) & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
Rem 4-11-2008 Physical disk read/write modification
If GlRDP = True Then
    hDevice = CreateFile(sdrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, &H20000000, 0&)
Else
    hDevice = CreateFile("\\.\" & sdrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, &H20000000, 0&)
End If
If hDevice = INVALID_HANDLE_VALUE Then Exit Function
Call SetFilePointer(hDevice, iStartSec * bytespersector, 0, FILE_BEGIN)
ReDim lpBuffer(cBytes - 1)
ReDim abBuff(nSectors * bytespersector - 1)
Call ReadFile(hDevice, abBuff(0), UBound(abBuff) + 1, cBytes, 0&)
CloseHandle hDevice
CopyMemory lpBuffer(0), abBuff(iOffset), cBytes
DirectReadDriveNT = cBytes
    
End Function
Public Function DirectWriteDriveNT(ByVal sdrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal sWrite As String) As Boolean
On Error Resume Next
    
Dim filebeginn As Long
Dim Pointer As Long
Dim hDevice As Long
Dim abBuff() As Byte
Dim ab() As Byte
Dim nRead As Long
Dim nSectors As Long
Dim doffset As Long
nSectors = Int((iOffset + Len(sWrite) - 1) / bytespersector) + 1  ' wieviele Sectoren zu schreiben

'hDevice = CreateFile("\\.\" & UCase(Left(sdrive, 1)) & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
If GlRDP = True Then
    hDevice = CreateFile(sdrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
Else
    hDevice = CreateFile("\\.\" & sdrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
End If

If hDevice = INVALID_HANDLE_VALUE Then Exit Function
ReDim abBuff(nSectors * bytespersector - 1)
Pointer = GetFileSizeEx(hDevice, doffset)
'abBuff = ADirectReadDriveNT(sdrive, iStartSec, 0, nSectors * 512)   ' lese diesen Bereicht (Sector)
ab = StrConv(sWrite, vbFromUnicode)
ReDim abuff(nSectors * bytespersector - 1)
ab = StrConv(sWrite, vbFromUnicode)

CopyMemory abBuff(iOffset), ab(0), Len(sWrite)   ' Text in Schreib buffer copieren
doffset = SetFilePointer(hDevice, iStartSec * bytespersector, 0, 0)
'pointer = LockFile(hDevice, LoWord(iStartSec * bytespersector), HiWord(iStartSec * bytespersector), LoWord(nSectors * bytespersector), HiWord(nSectors * bytespersector))
'pointer = LockFile(hDevice, doffset, 0, bytespersector, 0)

DirectWriteDriveNT = WriteFile(hDevice, abBuff(0), UBound(abBuff) + 1, nRead, 0&)
Pointer = FlushFileBuffers(hDevice)
'Call UnlockFile(hDevice, LoWord(iStartSec * bytespersector), HiWord(iStartSec * bytespersector), LoWord(nSectors * bytespersector), HiWord(nSectors * bytespersector))
'pointer = UnlockFile(hDevice, doffset, 0, bytespersector, 0)
CloseHandle hDevice
abBuff = ADirectReadDriveNT(sdrive, iStartSec, 0, nSectors * 512)   ' lese diesen Bereicht (Sector)
    
End Function
Public Function HiWord(ByVal dw As Long) As Integer
    HiWord = (dw And &HFFFF0000) \ 65536
End Function

Public Function LoWord(ByVal dw As Long) As Integer
On Error Resume Next
    
If dw And &H8000& Then
    LoWord = dw Or &HFFFF0000
Else
    LoWord = dw And &HFFFF&
End If

End Function
Public Function ADirectReadDriveNT(ByVal sdrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal cBytes As Long) As Variant
On Error Resume Next

    Dim hDevice As Long
    Dim abBuff() As Byte
    Dim abResult() As Byte
    Dim nSectors As Long
    Dim nRead As Long
    Dim Pointer As Long
    
    nSectors = Int((iOffset + cBytes - 1) / bytespersector) + 1
    'hDevice = CreateFile("\\.\" & UCase(Left(sdrive, 1)) & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    hDevice = CreateFile("\\.\" & sdrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If GlRDP = True Then
        hDevice = CreateFile(sdrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, &H20000000, 0&)
    Else
        hDevice = CreateFile("\\.\" & sdrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, &H20000000, 0&)
    End If

    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    Pointer = SetFilePointer(hDevice, iStartSec * bytespersector, 0, FILE_BEGIN)
    ReDim abResult(cBytes - 1)
    ReDim abBuff(nSectors * bytespersector - 1)
    Pointer = ReadFile(hDevice, abBuff(0), UBound(abBuff) + 1, nRead, 0&)
    CloseHandle hDevice
    CopyMemory abResult(0), abBuff(iOffset), cBytes
    ADirectReadDriveNT = abResult

End Function
'------------------------------------------------------------------------------------------------
' String nach Base64 codieren
'------------------------------------------------------------------------------------------------

Private Function StrToBase64(ByVal sInput As String) As String
On Error Resume Next

  Dim i As Long
  Dim sBase64 As String
  Dim nByte As Long
  Dim nChar As Long
  Dim nOldChar As Long
  Dim sChar As String
  Dim nLen As Long
  Dim sOutput As String
  Dim nPos As Long
  Dim nLenIn    As Long
  Dim nLenAddedBytes As Long
  
  
  '--- eingefügt SD 23.09.19 Padding: Nullbytes ans Ende, damit Eingabestring durch 3 teilbar
  nLenIn = Len(sInput)
  If nLenIn Mod 3 = 1 Then
    sInput = sInput + Chr(0) + Chr(0)
    nLenAddedBytes = 2
  ElseIf nLenIn Mod 3 = 1 Then
    sInput = sInput + Chr(0)
    nLenAddedBytes = 1
  End If
  ' ---
  
  ' Zulässige Zeichen (base64)
  sBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
 
  nLen = Len(sInput)
  If nLen > 0 Then
    ' aus je 3 Bytes werden 4 Base64-codierte Bytes
    sOutput = Space$(nLen / 3 * 4 + 2)
 
    ' String durchlaufen
    nPos = 1
    For i = 1 To nLen
      nByte = nByte + 1
 
      ' Ascii-Zeichen des Originalstrings
      nChar = Asc(Mid$(sInput, i, 1))
      Select Case nByte
        Case 1
          sChar = Mid$(sBase64, 1 + ((Int(nChar / 4) And &H3F)), 1)
          Mid$(sOutput, nPos, 1) = sChar
 
        Case 2
          sChar = Mid$(sBase64, 1 + ((((nOldChar * 16) And &H30) Or _
           ((Int(nChar / 16) And &HF)))), 1)
          Mid$(sOutput, nPos, 1) = sChar
 
        Case 3
          sChar = Mid$(sBase64, 1 + ((((nOldChar * 4) And &H3C) Or _
            ((Int(nChar / 64) And &H3)))), 1) & _
            Mid$(sBase64, 1 + ((nChar And &H3F)), 1)
          Mid$(sOutput, nPos, 2) = sChar
          nPos = nPos + 1
          nByte = 0
 
      End Select
      nPos = nPos + 1
      nOldChar = nChar
    Next i
 
    ' ggf. noch mit = auffüllen
    Select Case nByte
      Case 1
        sChar = Mid$(sBase64, 1 + (((nOldChar * 16) And &H30)), 1)
        Mid$(sOutput, nPos, 3) = sChar & "=="
      Case 2
        sChar = Mid$(sBase64, 1 + (((nOldChar * 4) And &H3C)), 1)
        Mid$(sOutput, nPos, 2) = sChar & "="
    End Select
  End If
 
  '--- eingefügt SD 23.09.19 angefügte Bytes nicht codieren, sondern als = kennzeichnen
  sOutput = RTrim$(sOutput)
  If nLenAddedBytes > 0 Then
    sOutput = Left(sOutput, Len(sOutput) - nLenAddedBytes) + String$(nLenAddedBytes, "=")
  End If
  ' ---
  
  StrToBase64 = RTrim$(sOutput)
  
End Function
Private Sub TSE_Tran(art, trak_nr, dat_len, client, protype) '0 Start 2 Finish
On Error Resume Next

Dim transn As String
Dim prodatalen As String ' prozess Data länge
Dim protypelen As String ' prozess Type länge Beleg?
Dim adddatalen As String ' nicht benutzt
Dim data As String
Dim i As Integer
Dim x As Integer
Dim TEMP() As Byte
  
  befehl.blen(0) = 0
  befehl.blen(1) = 53             ' wird unten dann noch richtig gemacht
  
  For i = 0 To 300
    befehl.befehl(i) = 0
  Next
  befehl.befehl(0) = &H91
  befehl.befehl(1) = &H0
  
  befehl.befehl(2) = art                  ' Transaktion Start = 0
  befehl.befehl(3) = Len(client)          ' Länge des Clients
  TEMP = StrConv(client, vbFromUnicode)
  For i = 0 To Len(client) - 1
    befehl.befehl(i + 4) = TEMP(i) 'Asc(Mid(meineid, i, 1))  'übertrage Client
  Next
  
  
  '********** Transaktionsnummer wenn vorhanden
  transn = "000000000000000" & Hex$(Val(trak_nr))
  transn = Mid(transn, Len(transn) - 15)
  i = i + 4
  For x = 0 To 7
    befehl.befehl(i + x) = CByte("&H" & Mid(transn, x * 2 + 1, 2)) 'Hex String in Bin Byte
  Next
  
  '************** Process Daten länge
  
  i = i + 8
  prodatalen = "000000000000000" & Hex(dat_len) 'prozess Data länge für Datenimport
  prodatalen = Mid(prodatalen, Len(prodatalen) - 15)
  
  For x = 0 To 7
    befehl.befehl(i + x) = CByte("&H" & Mid(prodatalen, x * 2 + 1, 2)) 'Hex String in Bin Byte
  Next
  i = i + 8
  
  '************* Process Type Länge
  
  protypelen = "00000000000000" & Hex(Len(protype)) 'prozess type länge  "Kassenbeleg-V1"
  For x = 0 To 7
    befehl.befehl(i + x) = CByte("&H" & Mid(protypelen, x * 2 + 1, 2)) 'Hex String in Bin Byte
  Next
  
  '*********** Process Type
  i = i + 8
  TEMP = StrConv(protype, vbFromUnicode)
  For x = 0 To Len(protype) - 1
    befehl.befehl(i + x) = TEMP(x) 'Asc(Mid(protype, x, 1))
  Next
  
  '****************** weiter Daten werden nicht benutzt
  i = i + x
  adddatalen = "0000000000000000"  'additeionale Daten nicht benutzt
  For x = 0 To 7
    befehl.befehl(i + x) = CByte("&H" & Mid(adddatalen, x * 2 + 1, 2)) 'Hex String in Bin Byte
  Next
  befehl.blen(1) = Len(protype) + Len(client) + 4 * 8 + 4   ' gesamtlänge des Befehls
  data = Hex(befehl.blen(1)) & vbCrLf
  For i = 0 To 200
    If befehl.befehl(i) > 15 Then
    data = data & Hex(befehl.befehl(i)) & " "
  Else
    data = data & "0" & Hex(befehl.befehl(i)) & " "
  End If
  Next

End Sub
Private Function TSE_And(ioldpin, inewpin, id)
On Error Resume Next

Dim newpin As pin
Dim i As Integer
Dim intdatnum As Integer
Dim ErStr As String

If Len(ioldpin) <> Len(inewpin) Then
  TSE_And = "Fehler Pin/Puk länge stimmt nicht!"
  Exit Function
End If

newpin.befehl(1) = &H24
newpin.befehl(2) = &H0

If id = 1 Then
  newpin.user_id = 2  ' als Admin  2 = timelogin
Else
  newpin.user_id = 1  ' als Admin
End If

newpin.cpinlen = 5
For i = 1 To 5
  newpin.cpin(i) = Asc(Mid(ioldpin, i, 1))
Next
newpin.newpinlen = 5
For i = 1 To 5
  newpin.newpin(i) = Asc(Mid(inewpin, i, 1))
Next
newpin.blen(1) = 0
newpin.blen(2) = 15

intdatnum = FreeFile
Open GLTSL & "\TSE_COMM.DAT" For Random Access Write As intdatnum Len = Len(newpin)
Put #intdatnum, 1, newpin
Close intdatnum
WindowSleep 1000
ErStr = TSE_Wart
TSE_And = "change PIN " & inewpin & " Resultat :" & ErStr

End Function
Private Function TSE_Clie(meineid)

Dim i As Integer
Dim intdatnum As Integer
'Dim client_reg As reg_client
Dim ErStr As String

befehl.blen(0) = 0
befehl.blen(1) = 3 + Len(meineid)
befehl.befehl(0) = &H41
befehl.befehl(1) = &H0
befehl.befehl(2) = Len(meineid)
For i = 1 To Len(meineid)
  befehl.befehl(i + 2) = Asc(Mid(meineid, i, 1))
Next

intdatnum = FreeFile
Open GLTSL & "\TSE_COMM.DAT" For Random Access Write As intdatnum Len = Len(befehl)
Put #intdatnum, 1, befehl
Close intdatnum
WindowSleep 1000
ErStr = TSE_Wart
TSE_Clie = "neuer Client: " & meineid & " Reslutat: " & ErStr
  
End Function
Private Function TSE_CTSS() As String
On Error Resume Next
  
befehl.blen(0) = 0
befehl.blen(1) = 2
befehl.befehl(0) = &H60
befehl.befehl(1) = &H0
TSE_Send
DoEvents
TSE_CTSS = TSE_Wart
DoEvents

End Function
Private Function TSE_Fini()
On Error Resume Next

  Dim Zeile As String
  Dim i As Integer
  Dim fehler As Integer
  Dim longzeile As Long
  Dim zeit As String
  Dim signaturlaenge As Integer
  Dim ergebnis As String
  Dim data As String
  Dim trcnt As Long
  Dim sigcnt As Long
  Dim Signatur As String

nochmals:
  befehl.blen(0) = 0
  befehl.blen(1) = 2
  befehl.befehl(0) = &H95
  befehl.befehl(1) = &H0
  
  TSE_Send
  DoEvents
  
  WindowSleep 1000
  DoEvents
  
  ergebnis = TSE_Wart()
  If Val(ergebnis) > 0 Then     ' Payload wird in der TSE auf die richtige Stelle geschoben das kann dauern!!
  '                               Fehler 1006 kein Payload vorhanden!
    If Val(ergebnis) = 1006 And fehler < 1000 Then
      fehler = fehler + 1
      GoTo nochmals
    End If
    TSE_Fini = "Fehler: Transaktion finish"
    Exit Function
  End If
  Zeile = ""
  
  '*******************************************
  ' Transaktionsnummer -> TransaktionsCounter
  '*******************************************
  For i = 0 To 7
    If RuAry(0 + 7 + i) > 15 Then
      Zeile = Zeile & Hex(RuAry(0 + 7 + i)) 'sig nummer
    Else
      Zeile = Zeile & "0" & Hex(RuAry(0 + 7 + i)) 'sig nummer
    End If
  Next
  If Mid(Zeile, 1, 1) > "0" Then
    MsgBox Mid(Zeile, 1, 6), , "Fehler"
    Exit Function
  End If
  longzeile = CLng("&H" & Zeile)
  trcnt = Val(longzeile)
  '***********************************************
  ' Ende Zeit  -> LogTime
  '***********************************************
  Zeile = ""
  For i = 0 To 7
    If RuAry(40 + 7 + i) > 15 Then
      Zeile = Zeile & Hex(RuAry(40 + 7 + i))
    Else
      Zeile = Zeile & "0" & Hex(RuAry(40 + 7 + i))
    End If
  Next
  longzeile = CLng("&H" & Zeile)
  zeit = DateAdd("s", longzeile, DateSerial(1970, 1, 1))
  'LogTime = zeit
  '*****************************************************
  ' SiCou -> SiCou
  '*****************************************************
  Zeile = ""
  For i = 0 To 7
    If RuAry(48 + 7 + i) > 15 Then
      Zeile = Zeile & Hex(RuAry(48 + 7 + i))
    Else
      Zeile = Zeile & "0" & Hex(RuAry(48 + 7 + i)) ' signaturcounter
    End If
  Next
  longzeile = CLng("&H" & Zeile)
  sigcnt = CLng("&H" & Zeile)
  '**************************************************************
  ' Signaturlänge
  '***************************************************************
  Zeile = ""
  For i = 0 To 7
    If RuAry(56 + 7 + i) > 15 Then
      Zeile = Zeile & Hex(RuAry(56 + 7 + i)) ' signatur länge
    Else
      Zeile = Zeile & "0" & Hex(RuAry(56 + 7 + i)) ' signatur länge
    End If
  Next
  longzeile = CLng("&H" & Zeile)
  signaturlaenge = longzeile
  
  '******************************************************************
  ' Signatur ->Signatur
  '******************************************************************
  Zeile = ""
  
  For i = 0 To signaturlaenge - 1
    If RuAry(64 + 7 + i) > 15 Then
      Zeile = Zeile & Hex(RuAry(64 + 7 + i)) & ""  ' signatur länge
    Else
      Zeile = Zeile & "0" & Hex(RuAry(64 + 7 + i)) & ""  ' signatur länge
    End If
  Next
  
'*******************************
'******************
Zeile = ""
For i = 0 To signaturlaenge - 1
  Zeile = Zeile + Chr(RuAry(64 + 7 + i))
Next


'*******************************
  
  data = data & vbCrLf & Zeile
  Signatur = StrToBase64(Zeile)
  TSE_Fini = trcnt & ";" & sigcnt & ";" & Signatur & ";" & zeit
  Exit Function
'*************************************
' übertrage in Kontrollfenster
data = "Transaktionscounter: " & TrCou & "  "
data = data & "Sig Count:" & SiCou & vbCrLf
data = data & "Sig Länge:" & signaturlaenge & vbCrLf

data = data & Mid(TSSig, 1, 68) & vbCrLf
data = data & Mid(TSSig, 70, 68) & vbCrLf
data = data & Mid(TSSig, 139, 68) & vbCrLf
data = data & Mid(TSSig, 208, 68) & vbCrLf
data = data & Mid(TSSig, 275, 75) & vbCrLf

End Function


Function bytehex(ByRef was, start, wieviele) As String
On Error Resume Next
  
Dim RuAry As String
Dim AktZa As Integer

For AktZa = 0 To wieviele - 1
    If was(start + AktZa) > 15 Then
        RuAry = RuAry & Hex(was(start + AktZa))
    Else
        RuAry = RuAry & "0" & Hex(was(start + AktZa))
    End If
Next

bytehex = RuAry

End Function

Function Str2Hex(c As String) As String
On Error Resume Next

Dim OuStr As String
Dim AktZa As Integer

For AktZa = 1 To Len(c)
    OuStr = OuStr + Right("00" + Hex(Asc(Mid(c, AktZa, 1))), 2)
Next

Str2Hex = OuStr
  
End Function


Public Function Hex2Str(ByVal CeStr As String) As String
On Error Resume Next

Dim AktZa As Integer
Dim OuStr As String
Dim n1 As Integer
Dim n2 As Integer

For AktZa = 1 To Len(CeStr) / 2
    If Val("&h" + Left(CeStr, 2)) = 0 Then
        OuStr = OuStr + "|"
    Else
        OuStr = OuStr + Chr(Val("&h" + Left(CeStr, 2)))
    End If
    CeStr = Mid(CeStr, 3)
Next

AktZa = Len(OuStr)
Hex2Str = OuStr

End Function
