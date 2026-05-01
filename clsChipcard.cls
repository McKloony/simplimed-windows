VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChipcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' clsChipcard - Konsolidiertes Chipkartenleser-Modul
' Ersetzt: cls_ChipRead, clsCardReader und alle Geraetetreiber-Klassen
' ============================================================================

' Windows API Deklarationen
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

' zlib Deklarationen fuer Dekompression
Private Declare Function uncompress Lib "zlibwapi.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress Lib "zlibwapi.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

' ============================================================================
' CT-API Deklarationen fuer alle unterstuetzten Chipkartenleser
' ============================================================================

' ctpcsc32kv.dll - Reiner SCT
Private Declare Function CT_INIT_ctpcsc32kv Lib "ctpcsc32kv.dll" Alias "CT_init" (ByVal ctn As Integer, ByVal pn As Integer) As Byte
Private Declare Function CT_DATA_ctpcsc32kv Lib "ctpcsc32kv.dll" Alias "CT_data" (ByVal ctn As Integer, ByRef dad As Integer, ByRef sad As Integer, ByVal lenc As Integer, ByRef CmAry As Byte, ByRef lenr As Integer, ByRef ReAry As Byte) As Byte
Private Declare Function CT_CLOSE_ctpcsc32kv Lib "ctpcsc32kv.dll" Alias "CT_close" (ByVal ctn As Integer) As Byte

' ctcym.dll - Omnikey
Private Declare Function CT_INIT_ctcym Lib "ctcym.dll" Alias "CT_init" (ByVal ctn As Integer, ByVal pn As Integer) As Byte
Private Declare Function CT_DATA_ctcym Lib "ctcym.dll" Alias "CT_data" (ByVal ctn As Integer, ByRef dad As Integer, ByRef sad As Integer, ByVal lenc As Integer, ByRef CmAry As Byte, ByRef lenr As Integer, ByRef ReAry As Byte) As Byte
Private Declare Function CT_CLOSE_ctcym Lib "ctcym.dll" Alias "CT_close" (ByVal ctn As Integer) As Byte

' Ctvml32 - CT
Private Declare Function CT_INIT_ctvml32 Lib "Ctvml32" Alias "CT_init" (ByVal ctn As Integer, ByVal pn As Integer) As Byte
Private Declare Function CT_DATA_ctvml32 Lib "Ctvml32" Alias "CT_data" (ByVal ctn As Integer, ByRef dad As Integer, ByRef sad As Integer, ByVal lenc As Integer, ByRef CmAry As Byte, ByRef lenr As Integer, ByRef ReAry As Byte) As Byte
Private Declare Function CT_CLOSE_ctvml32 Lib "Ctvml32" Alias "CT_close" (ByVal ctn As Integer) As Byte

' CTORG32.dll - Orga / Cherry (orga930m, orga6000, cherry1530)
Private Declare Function CT_INIT_ctorg32 Lib "CTORG32.dll" Alias "CT_init" (ByVal ctn As Integer, ByVal pn As Integer) As Byte
Private Declare Function CT_DATA_ctorg32 Lib "CTORG32.dll" Alias "CT_data" (ByVal ctn As Integer, ByRef dad As Integer, ByRef sad As Integer, ByVal lenc As Integer, ByRef CmAry As Byte, ByRef lenr As Integer, ByRef ReAry As Byte) As Byte
Private Declare Function CT_CLOSE_ctorg32 Lib "CTORG32.dll" Alias "CT_close" (ByVal ctn As Integer) As Byte

' ctrsct32.dll - Reiner SCT 3.2
Private Declare Function CT_INIT_ctrsct32 Lib "ctrsct32.dll" Alias "CT_init" (ByVal ctn As Integer, ByVal pn As Integer) As Byte
Private Declare Function CT_DATA_ctrsct32 Lib "ctrsct32.dll" Alias "CT_data" (ByVal ctn As Integer, ByRef dad As Integer, ByRef sad As Integer, ByVal lenc As Integer, ByRef CmAry As Byte, ByRef lenr As Integer, ByRef ReAry As Byte) As Byte
Private Declare Function CT_CLOSE_ctrsct32 Lib "ctrsct32.dll" Alias "CT_close" (ByVal ctn As Integer) As Byte

' ctpcsc31.dll - CT-API 3.1
Private Declare Function CT_INIT_ctpcsc31 Lib "ctpcsc31.dll" Alias "CT_init" (ByVal ctn As Integer, ByVal pn As Integer) As Byte
Private Declare Function CT_DATA_ctpcsc31 Lib "ctpcsc31.dll" Alias "CT_data" (ByVal ctn As Integer, ByRef dad As Integer, ByRef sad As Integer, ByVal lenc As Integer, ByRef CmAry As Byte, ByRef lenr As Integer, ByRef ReAry As Byte) As Byte
Private Declare Function CT_CLOSE_ctpcsc31 Lib "ctpcsc31.dll" Alias "CT_close" (ByVal ctn As Integer) As Byte

' ============================================================================
' Private Member-Variablen
' ============================================================================

' Rueckgabewerte
Private RetWe As Long
Private ReStr As String

' Patientendaten
Private CliID As String          ' Versicherten_ID
Private PaVor As String          ' Patient_Vorname
Private PaNam As String          ' Patient_Name
Private PaGeb As String          ' Patient_Geboren
Private PaStr As String          ' Patient_Strasse
Private PaPLZ As String          ' Patient_PLZ
Private PaOrt As String          ' Patient_Ort
Private PaGes As String          ' Patient_Geschlecht
Private HauNu As String          ' Patient_Hausnummer
Private WoSiC As String          ' Wohnsitzlaendercode
Private PaTit As String          ' Patient_Titel
Private PaZun As String          ' Patient_Zusatz (Vorsatzwort)
Private PaLKZ As String          ' Patient_Landeskennzeichen

' Kostentraegerdaten
Private kotrk As String          ' Kostentraegerkennung
Private KoTrN As String          ' Kostentraegername
Private KTrLK As String          ' Kostentraegerlaendercode
Private AKTrN As String          ' Kostentraegername2
Private AKTrK As String          ' Kostentraegerkennung2
Private KoTraKen As String       ' Alternative Kostentraegerkennung
Private KoTraNam As String       ' Alternative Kostentraegername

' eGK-spezifische Daten
Private ElGeK As Boolean         ' IsteGK Flag
Private VeBeg As String          ' Vertragsbegin
Private GKVRK As String          ' Rechtskreis
Private GKVVA As String          ' Versichertenart
Private GKVVS As String          ' Versichertenstatus
Private GKVAM As String          ' Ambulant
Private GKVST As String          ' Stationaer
Private GKVWO As String          ' Wohnortprinzip

' KVK-spezifische Daten
Private KarNr As String          ' Kartennummer
Private KaSta As String          ' Kartenstatus
Private StaEr As String          ' KVKStatus
Private KaGul As String          ' Kartengueltigkeit
Private KaDat As String          ' Kartendatum

' Konfiguration
Private SmRea As String          ' SmartReader Name
Private TePor As Integer         ' TerminalPort

' DLL-Typ Enumeration
Public Enum DLLType
    NoDLL = 0
    StandardDLL = 1
    ActiveXDLL = 2
End Enum

' ============================================================================
' Hilfsfunktion: DLL Verfuegbarkeit pruefen
' ============================================================================
Public Function DLLAvail(sFile As String) As DLLType

Dim hwnd As Long
Dim sPrc As String

hwnd = LoadLibrary(sFile)

If hwnd <> 0 Then
    DLLAvail = StandardDLL

    sPrc = "DllRegisterServer"

    If CBool(GetProcAddress(hwnd, sPrc)) = True Then
        DLLAvail = ActiveXDLL
    End If

    Call FreeLibrary(hwnd)
Else

  DLLAvail = NoDLL

End If

End Function

' ============================================================================
' CT-API Wrapper Funktionen - rufen je nach Reader die richtige DLL auf
' ============================================================================
Private Function CT_INITP(ByVal ctn As Integer, ByVal pn As Integer) As Byte
    Select Case SmRea
    Case "ctpcsc32kv"
        CT_INITP = CT_INIT_ctpcsc32kv(ctn, pn)
    Case "ctcym"
        CT_INITP = CT_INIT_ctcym(ctn, pn)
    Case "ctvml32"
        CT_INITP = CT_INIT_ctvml32(ctn, pn)
    Case "orga930m", "orga6000", "cherry1530"
        CT_INITP = CT_INIT_ctorg32(ctn, pn)
    Case "ctrsct32"
        CT_INITP = CT_INIT_ctrsct32(ctn, pn)
    Case "ctpcsc31"
        CT_INITP = CT_INIT_ctpcsc31(ctn, pn)
    End Select
End Function

Private Function CT_DATAP(ByVal ctn As Integer, ByRef dad As Integer, ByRef sad As Integer, ByVal lenc As Integer, ByRef CmAry As Byte, ByRef lenr As Integer, ByRef ReAry As Byte) As Byte
    Select Case SmRea
    Case "ctpcsc32kv"
        CT_DATAP = CT_DATA_ctpcsc32kv(ctn, dad, sad, lenc, CmAry, lenr, ReAry)
    Case "ctcym"
        CT_DATAP = CT_DATA_ctcym(ctn, dad, sad, lenc, CmAry, lenr, ReAry)
    Case "ctvml32"
        CT_DATAP = CT_DATA_ctvml32(ctn, dad, sad, lenc, CmAry, lenr, ReAry)
    Case "orga930m", "orga6000", "cherry1530"
        CT_DATAP = CT_DATA_ctorg32(ctn, dad, sad, lenc, CmAry, lenr, ReAry)
    Case "ctrsct32"
        CT_DATAP = CT_DATA_ctrsct32(ctn, dad, sad, lenc, CmAry, lenr, ReAry)
    Case "ctpcsc31"
        CT_DATAP = CT_DATA_ctpcsc31(ctn, dad, sad, lenc, CmAry, lenr, ReAry)
    End Select
End Function

Private Function CT_CLOSEP(ByVal ctn As Integer) As Byte
    Select Case SmRea
    Case "ctpcsc32kv"
        CT_CLOSEP = CT_CLOSE_ctpcsc32kv(ctn)
    Case "ctcym"
        CT_CLOSEP = CT_CLOSE_ctcym(ctn)
    Case "ctvml32"
        CT_CLOSEP = CT_CLOSE_ctvml32(ctn)
    Case "orga930m", "orga6000", "cherry1530"
        CT_CLOSEP = CT_CLOSE_ctorg32(ctn)
    Case "ctrsct32"
        CT_CLOSEP = CT_CLOSE_ctrsct32(ctn)
    Case "ctpcsc31"
        CT_CLOSEP = CT_CLOSE_ctpcsc31(ctn)
    End Select
End Function

' ============================================================================
' Hauptfunktion: Karte lesen
' ============================================================================
Public Function CarRead() As Long
On Error GoTo DaErr

Dim lRetWe As Long
Dim sDllName As String

If (DLLAvail("zlibwapi.dll") = NoDLL) Then
    ReStr = "missing zlibwapi.dll"
    CarRead = -1
    Exit Function
End If

' DLL-Name je nach Reader ermitteln
Select Case SmRea
Case "ctpcsc32kv"
    sDllName = "ctpcsc32kv.dll"
Case "ctcym"
    sDllName = "ctcym.dll"
Case "ctvml32"
    sDllName = "Ctvml32.dll"
Case "orga930m", "orga6000", "cherry1530"
    sDllName = "CTORG32.dll"
Case "ctrsct32"
    sDllName = "ctrsct32.dll"
Case "ctpcsc31"
    sDllName = "ctpcsc31.dll"
Case Else
    ReStr = "Unbekannter Kartenleser: " & SmRea
    CarRead = -1
    Exit Function
End Select

' DLL Verfuegbarkeit pruefen
If (DLLAvail(sDllName) = NoDLL) Then
    ReStr = "missing " & sDllName
    CarRead = -1
    Exit Function
End If

' Karte lesen
lRetWe = ReadChipcard()
CarRead = lRetWe

Exit Function

DaErr:
MsgBox Err.Description, 48, "CarRead " & Err.Number
Resume Next

End Function

' ============================================================================
' Interne Funktion: Chipkarte lesen (entspricht clsCardReader.CheCrd)
' ============================================================================
Private Function ReadChipcard() As Long
On Error GoTo DaErr

Dim RetAb As Long
Dim lRetWe As Byte
Dim pn As Integer
Dim KaTyp As Integer
Dim ReAry(0 To 2999) As Byte
Dim CmAry() As Byte

Dim XMLDo As New MSXML2.DOMDocument60

KaTyp = 0
RetAb = 0
pn = TePor

' Reader initialisieren
Call InitReader(ReStr, lRetWe, pn)
DoEvents
Call ErrorLookup(ReStr, lRetWe)
DoEvents
Call CardReaderReset(ReStr, lRetWe, RetAb)
DoEvents
Call RequestIIC(ReStr, lRetWe, RetAb, KaTyp)
DoEvents

' Je nach Kartentyp lesen
If KaTyp = 1 Then
    ' KVK (Krankenversichertenkarte)
    DoEvents
    Call ReadBinaryPatient(lRetWe, RetAb, ReStr)
    ElGeK = False
ElseIf KaTyp = 2 Then
    ' eGK (elektronische Gesundheitskarte)
    DoEvents
    Call ReadEGKCard(lRetWe, RetAb, ReStr)
    ElGeK = True
End If

ReadChipcard = RetAb

DoEvents
Call EraseReader(CmAry, ReAry, lRetWe)
DoEvents
Call Eject(CmAry, ReAry, lRetWe)
DoEvents
Call CloseReader

Exit Function

DaErr:
MsgBox Err.Description, 48, "ReadChipcard " & Err.Number
Resume Next

End Function

' ============================================================================
' Fehlerauswertung
' ============================================================================
Private Function ErrorLookup(ByRef sReStr As String, ByVal code As Integer) As String

If sReStr <> vbNullString Then
    Exit Function
End If
    Select Case code
        Case 129
            sReStr = "Kein Kartenlesegeraet angeschlossen oder falsches gewaehlt"

        Case 255
            sReStr = "Bitte ueberpruefen Sie die Porteinstellungen"

        Case Else
            sReStr = ""
    End Select
End Function

' ============================================================================
' Reader initialisieren
' ============================================================================
Private Function InitReader(ByRef sReStr As String, ByRef lRetWe As Byte, ByRef pn As Integer)
    Dim ctn As Integer
    ctn = 1
    sReStr = "Start von Kartenleser"
    lRetWe = CT_INITP(ctn, pn)
End Function

' ============================================================================
' Reader Reset
' ============================================================================
Private Function CardReaderReset(ByRef sReStr As String, ByRef lRetWe As Byte, ByRef RetAb As Long)
Dim pn As Integer
Dim ctn As Integer
Dim dad As Integer
Dim sad As Integer
Dim lenc As Integer
Dim lenr As Integer
Dim AktZa As Integer

Dim ReAry(0 To 2999) As Byte
Dim CmAry() As Byte
ctn = 1
    If RetAb = 0 Then
        ReDim CmAry(4)
        Erase ReAry

        CmAry(0) = &H20
        CmAry(1) = &H11
        CmAry(2) = &H0
        CmAry(3) = &H0
        CmAry(4) = &H0
        dad = 1
        sad = 2
        lenc = 5
        lenr = 2000

        lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))
        If lRetWe <> 0 Then
            RetAb = lRetWe
            sReStr = "Das Geraet wird zurueckgesetzt"
        Else
            sReStr = "CT_Reset:"
            For AktZa = 0 To lenr - 1
                sReStr = sReStr + CStr(Hex$(ReAry(AktZa)))
            Next
        End If
    End If

End Function

' ============================================================================
' Kartentyp anfordern (KVK oder eGK)
' ============================================================================
Private Function RequestIIC(ByRef sReStr As String, ByRef lRetWe As Byte, ByRef RetAb As Long, ByRef KaTyp As Integer)
    Dim CmAry() As Byte
    Dim ReAry(0 To 2999) As Byte
    Dim AktZa As Integer
    Dim ctn As Integer
    Dim dad As Integer
    Dim sad As Integer
    Dim lenc As Integer
    Dim lenr As Integer

    If RetAb = 0 Then
       ReDim CmAry(5)
       Erase ReAry

       CmAry(0) = &H20
       CmAry(1) = &H12
       CmAry(2) = &H1
       CmAry(3) = &H0
       CmAry(4) = &H1
       CmAry(5) = &H1
       ctn = 1
       dad = 1
       sad = 2
       lenc = 6
       lenr = 2000

        lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))
        If lRetWe <> 0 Then
            RetAb = lRetWe
            sReStr = "Es liegt ein Abfragefehler vor"
        Else
            sReStr = "CT_Request:"
            For AktZa = 0 To lenr - 1
                sReStr = sReStr + CStr(Hex$(ReAry(AktZa)))
            Next

            If sReStr = "CT_Request:620" Then
                sReStr = "Es ist keine Karte eingesteckt"
                KaTyp = 0
            End If

            If sReStr = "CT_Request:901" Then
                sReStr = "eGK erkannt"
                KaTyp = 2
            End If

            If sReStr = "CT_Request:900" Then
                sReStr = "KVK erkannt"
                KaTyp = 1
            End If
        End If
    End If

End Function

' ============================================================================
' KVK (Krankenversichertenkarte) lesen
' ============================================================================
Private Function ReadBinaryPatient(ByRef lRetWe As Byte, ByRef RetAb As Long, ByRef sReStr As String)
On Error Resume Next

Dim CmAry() As Byte
Dim ReAry(0 To 2999) As Byte
Dim AktZa As Integer
Dim ctn As Integer
Dim dad As Integer
Dim sad As Integer
Dim lenc As Integer
Dim lenr As Integer

Dim StVdt As String
Dim FeStr As String

ReDim CmAry(10)
Erase ReAry

Call SelectKVKRoot(lRetWe, sReStr, RetAb)
Call CardReadBinary(lRetWe, ReAry, lenr)

If lRetWe <> 0 Then
    RetAb = lRetWe
    sReStr = "Es ist ein Lesefehler auf der Karte aufgetreten"
Else
    sReStr = vbNullString
    For AktZa = 0 To lenr - 1
        sReStr = sReStr + Chr(Val("&H" & (CStr(Hex$(ReAry(AktZa))))))
        FeStr = FeStr + Chr$(32) + CStr(Hex$(ReAry(AktZa)))
    Next
    StVdt = sReStr
    If InStr(StVdt, Chr(96)) > 0 Then
        If Asc(Mid$(StVdt, InStr(StVdt, Chr(96)) + 1, 1)) <= 128 Then
            StVdt = Mid$(StVdt, InStr(StVdt, Chr(96)) + 2, 127 + Asc(Mid$(StVdt, InStr(StVdt, Chr(96)) + 1, 1)))
        Else
            StVdt = Mid$(StVdt, InStr(StVdt, Chr(96)) + 3, 127 + Asc(Mid$(StVdt, InStr(StVdt, Chr(128)) + 2, 1)))
        End If

        ' Umlaute ersetzen
        StVdt = Replace$(StVdt, Chr(123), Chr(228)) ' ae ersetzen
        StVdt = Replace$(StVdt, Chr(124), Chr(246)) ' oe ersetzen
        StVdt = Replace$(StVdt, Chr(125), Chr(252)) ' ue ersetzen
        StVdt = Replace$(StVdt, Chr(126), Chr(223)) ' ss ersetzen
        StVdt = Replace$(StVdt, Chr(91), Chr(196))  ' Ae ersetzen
        StVdt = Replace$(StVdt, Chr(92), Chr(214))  ' Oe ersetzen
        StVdt = Replace$(StVdt, Chr(93), Chr(220))  ' Ue ersetzen

        ' VDT-Tags auswerten
        If InStr(StVdt, Chr(128)) > 0 Then KoTrN = Mid$(StVdt, InStr(StVdt, Chr(128)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(128)) + 1, 1)))
        If InStr(StVdt, Chr(129)) > 0 Then kotrk = Mid$(StVdt, InStr(StVdt, Chr(129)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(129)) + 1, 1)))
        If InStr(StVdt, Chr(143)) > 0 Then KarNr = Mid$(StVdt, InStr(StVdt, Chr(143)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(143)) + 1, 1)))
        If InStr(StVdt, Chr(130)) > 0 Then CliID = Mid$(StVdt, InStr(StVdt, Chr(130)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(130)) + 1, 1)))
        If InStr(StVdt, Chr(131)) > 0 Then KaSta = Mid$(StVdt, InStr(StVdt, Chr(131)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(131)) + 1, 1)))
        If InStr(StVdt, Chr(144)) > 0 Then StaEr = Mid$(StVdt, InStr(StVdt, Chr(144)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(144)) + 1, 1)))
        If InStr(StVdt, Chr(132)) > 0 Then PaTit = Mid$(StVdt, InStr(StVdt, Chr(132)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(132)) + 1, 1)))
        If InStr(StVdt, Chr(133)) > 0 Then PaVor = Mid$(StVdt, InStr(StVdt, Chr(133)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(133)) + 1, 1)))
        If InStr(StVdt, Chr(134)) > 0 Then PaZun = Mid$(StVdt, InStr(StVdt, Chr(134)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(134)) + 1, 1)))
        If InStr(StVdt, Chr(135)) > 0 Then PaNam = Mid$(StVdt, InStr(StVdt, Chr(135)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(135)) + 1, 1)))
        If InStr(StVdt, Chr(136)) > 0 Then PaGeb = Mid$(StVdt, InStr(StVdt, Chr(136)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(136)) + 1, 1)))
        If InStr(StVdt, Chr(137)) > 0 Then PaStr = Mid$(StVdt, InStr(StVdt, Chr(137)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(137)) + 1, 1)))
        If InStr(StVdt, Chr(138)) > 0 Then PaLKZ = Mid$(StVdt, InStr(StVdt, Chr(138)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(138)) + 1, 1)))
        If InStr(StVdt, Chr(139)) > 0 Then PaPLZ = Mid$(StVdt, InStr(StVdt, Chr(139)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(139)) + 1, 1)))
        If InStr(StVdt, Chr(140)) > 0 Then PaOrt = Mid$(StVdt, InStr(StVdt, Chr(140)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(140)) + 1, 1)))
        If InStr(StVdt, Chr(141)) > 0 Then KaGul = Mid$(StVdt, InStr(StVdt, Chr(141)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(141)) + 1, 1)))
        If InStr(StVdt, Chr(145)) > 0 Then KaDat = Mid$(StVdt, InStr(StVdt, Chr(145)) + 2, Asc(Mid$(StVdt, InStr(StVdt, Chr(145)) + 1, 1)))

    End If
End If

sReStr = "KVK erfolgreich ausgelesen"

End Function

' ============================================================================
' eGK (elektronische Gesundheitskarte) lesen
' ============================================================================
Private Function ReadEGKCard(ByRef lRetWe As Byte, ByRef RetAb As Long, ByRef sReStr As String)
    Dim ReAry(0 To 2999) As Byte
    Dim CmAry() As Byte

    Dim XMLDo As New MSXML2.DOMDocument60

    XMLDo.async = False
    XMLDo.validateOnParse = False
    XMLDo.resolveExternals = False
    XMLDo.preserveWhiteSpace = True

    Dim root As String
    root = SelectEGKRoot(lRetWe)
    Call SelectFile(CmAry, ReAry, lRetWe)

    If RetAb = 0 Then
        Call ReadPerson(sReStr, lRetWe, RetAb, ReAry)
        Dim strString As String
        strString = UncompressZip(ReAry)
        XMLDo.loadXML (strString)
        CliID = GetElementValue("Versicherten_ID", XMLDo)
        PaGeb = GetElementValue("Geburtsdatum", XMLDo)
        PaTit = GetElementValue("Titel", XMLDo)
        PaZun = GetElementValue("Vorsatzwort", XMLDo)
        PaVor = GetElementValue("Vorname", XMLDo)
        PaNam = GetElementValue("Nachname", XMLDo)
        PaGes = GetElementValue("Geschlecht", XMLDo)
        PaPLZ = GetElementValue("Postleitzahl", XMLDo)
        PaOrt = GetElementValue("Ort", XMLDo)
        WoSiC = GetElementValue("Wohnsitzlaendercode", XMLDo)
        PaStr = GetElementValue("Strasse", XMLDo)
        HauNu = GetElementValue("Hausnummer", XMLDo)
        Call ChangeBirthDayFormat

        sReStr = "eGK erfolgreich ausgelesen"
    End If

    If RetAb = 0 Then
        Call ReadInsurance(sReStr, lRetWe, RetAb, ReAry)
        Dim strEFVD As String
        strEFVD = UncompressBytesZip(ReAry)
        XMLDo.loadXML (strEFVD)
        VeBeg = GetElementValue("Beginn", XMLDo)
        kotrk = GetElementValue("Kostentraegerkennung", XMLDo)
        KoTrN = GetElementValue("Name", XMLDo)
        KTrLK = GetElementValue("Kostentraegerlaendercode", XMLDo)
        GKVRK = GetElementValue("Rechtskreis", XMLDo)
        GKVVA = GetElementValue("Versichertenart", XMLDo)
        GKVVS = GetElementValue("Versichertenstatus_RSA", XMLDo)
        GKVAM = GetElementValue("Kostenerstattung_ambulant", XMLDo)
        GKVST = GetElementValue("Kostenerstattung_stationaer", XMLDo)
        GKVWO = GetElementValue("WOP", XMLDo)

        sReStr = "eGK erfolgreich ausgelesen"
    End If

    ' Strasse und Hausnummer kombinieren
    PaStr = PaStr & " " & HauNu

End Function

' ============================================================================
' Hilfsfunktionen fuer Kartenlesung
' ============================================================================
Private Function SelectKVKRoot(ByRef lRetWe As Byte, ByRef sReStr As String, ByRef RetAb As Long)

Dim ReAry(0 To 2999) As Byte
Dim CmAry() As Byte
Dim AktZa As Integer
Dim ctn As Integer
Dim dad As Integer
Dim sad As Integer
Dim lenc As Integer
Dim lenr As Integer

ReDim CmAry(10)
Erase ReAry

CmAry(0) = &H0
CmAry(1) = &HA4
CmAry(2) = &H4
CmAry(3) = &H0
CmAry(4) = &H6
CmAry(5) = &HD2
CmAry(6) = &H76
CmAry(7) = &H0
CmAry(8) = &H0
CmAry(9) = &H1
CmAry(10) = &H1

dad = 0
ctn = 1
sad = 2
lenc = &HB
lenr = 2000

lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))
If lRetWe <> 0 Then
    RetAb = lRetWe
    sReStr = "Bitte zuerst Datei auswaehlen"
Else
    sReStr = "CT_Select_File:"
    For AktZa = 0 To lenr - 1
        sReStr = sReStr + CStr(Hex$(ReAry(AktZa)))
    Next
End If

If sReStr <> "CT_Select_File:900" Then
    sReStr = "Es sind keine Kartendaten vorhanden"
    RetAb = 1
End If

End Function

Private Function SelectEGKRoot(ByRef lRetWe As Byte) As String

Dim ReAry(0 To 2999) As Byte
Dim CmAry() As Byte
Dim retval As String
Dim AktZa As Integer
Dim ctn As Integer
Dim dad As Integer
Dim sad As Integer
Dim lenc As Integer
Dim lenr As Integer

ReDim CmAry(12)
Erase ReAry

CmAry(0) = &H0
CmAry(1) = &HA4
CmAry(2) = &H4
CmAry(3) = &HC
CmAry(4) = &H7
CmAry(5) = &HD2
CmAry(6) = &H76
CmAry(7) = &H0
CmAry(8) = &H1
CmAry(9) = &H44
CmAry(10) = &H80
CmAry(11) = &H0
ctn = 1
dad = 0
sad = 2
lenc = 12
lenr = 2000

lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))

For AktZa = 0 To lenr - 1
    retval = retval + CStr(Hex$(ReAry(AktZa)))
Next
SelectEGKRoot = retval

End Function

Private Function SelectFile(ByRef CmAry() As Byte, ByRef ReAry() As Byte, lRetWe As Byte)

Dim ctn As Integer
Dim dad As Integer
Dim sad As Integer
Dim lenc As Integer
Dim lenr As Integer

ReDim CmAry(11)
Erase ReAry

CmAry(0) = &H0
CmAry(1) = &HA4
CmAry(2) = &H4
CmAry(3) = &HC
CmAry(4) = &H6
CmAry(5) = &HD2
CmAry(6) = &H76
CmAry(7) = &H0
CmAry(8) = &H0
CmAry(9) = &H1
CmAry(10) = &H2
ctn = 1
dad = 0
sad = 2
lenc = 11
lenr = 2000

lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))

End Function

Private Function CardReadBinary(ByRef lRetWe As Byte, ByRef ReAry() As Byte, ByRef lenr As Integer)

Dim CmAry() As Byte
Dim ctn As Integer
Dim dad As Integer
Dim sad As Integer
Dim lenc As Integer

ReDim CmAry(5)
Erase ReAry

CmAry(0) = &H0
CmAry(1) = &HB0
CmAry(2) = &H0
CmAry(3) = &H0
CmAry(4) = &H0
ctn = 1
dad = 0
sad = 2
lenc = 5
lenr = 2000

lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))

End Function

Private Function ReadPerson(ByRef sReStr As String, ByRef lRetWe As Byte, ByRef RetAb As Long, ByRef ReAry() As Byte)
        Dim CmAry() As Byte
        Dim ctn As Integer
        Dim dad As Integer
        Dim sad As Integer
        Dim lenc As Integer
        Dim lenr As Integer

        If RetAb = 0 Then
            ReDim CmAry(7)
            Erase ReAry
            CmAry(0) = &H0
            CmAry(1) = &HB0
            CmAry(2) = &H81
            CmAry(3) = &H0
            CmAry(4) = &H0
            CmAry(5) = &H0
            CmAry(6) = &H0
            ctn = 1
            dad = 0
            sad = 2
            lenc = 7
            lenr = 2000

            lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))
            If lRetWe <> 0 Then
                RetAb = lRetWe
                sReStr = "Es ist ein Lesefehler auf der Karte aufgetreten"
            End If
        End If

End Function

Private Function ReadInsurance(ByRef sReStr As String, ByRef lRetWe As Byte, ByRef RetAb As Long, ByRef ReAry() As Byte)
        Dim CmAry() As Byte
        Dim ctn As Integer
        Dim dad As Integer
        Dim sad As Integer
        Dim lenc As Integer
        Dim lenr As Integer

        If RetAb = 0 Then
            ReDim CmAry(7)
            Erase ReAry
            CmAry(0) = &H0
            CmAry(1) = &HB0
            CmAry(2) = &H82
            CmAry(3) = &H0
            CmAry(4) = &H0
            CmAry(5) = &H0
            CmAry(6) = &H0
            ctn = 1
            dad = 0
            sad = 2
            lenc = 7
            lenr = 2000

            lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))
            If lRetWe <> 0 Then
                RetAb = lRetWe
                sReStr = "Es ist ein Lesefehler auf der Karte aufgetreten"
            End If
        End If

End Function

Private Function EraseReader(ByRef CmAry() As Byte, ByRef ReAry() As Byte, lRetWe As Byte)

Dim ctn As Integer
Dim dad As Integer
Dim sad As Integer
Dim lenc As Integer
Dim lenr As Integer

ReDim CmAry(4)
Erase ReAry
CmAry(0) = &H0
CmAry(1) = &HE
CmAry(2) = &H0
CmAry(3) = &H0
ctn = 1
dad = 0
sad = 2
lenc = 4
lenr = 2000
lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))

End Function

Private Function Eject(ByRef CmAry() As Byte, ByRef ReAry() As Byte, lRetWe As Byte)
    Dim ctn As Integer
    Dim dad As Integer
    Dim sad As Integer
    Dim lenc As Integer
    Dim lenr As Integer

    ReDim CmAry(6)
    Erase ReAry
    CmAry(0) = &H20
    CmAry(1) = &H15
    CmAry(2) = &H1
    CmAry(3) = &H0
    CmAry(4) = &H1
    CmAry(5) = &H0
    dad = 0
    sad = 2
    lenc = 6
    lenr = 2000
    lRetWe = CT_DATAP(ctn, dad, sad, lenc, CmAry(0), lenr, ReAry(0))
End Function

Private Function CloseReader()
    Dim ctn As Integer
    ctn = 1
    CT_CLOSEP (ctn)
End Function

' ============================================================================
' XML-Hilfsfunktionen
' ============================================================================
Private Function GetElementValue(ByVal elementName As String, ByRef XMLDo As MSXML2.DOMDocument60) As String

Dim XMLis As MSXML2.IXMLDOMNodeList
Dim XMLNo As MSXML2.IXMLDOMNode

Set XMLis = XMLDo.documentElement.getElementsByTagName("*")

For Each XMLNo In XMLis
    If XMLNo.baseName = elementName Then
        GetElementValue = XMLNo.nodeTypedValue
    End If
Next

End Function

' ============================================================================
' Dekompression-Funktionen
' ============================================================================
Private Function UncompressZip(ByRef ReAry() As Byte) As String
    Dim startPD As Integer
    Dim length As Long
    Dim compress_array() As Byte
    Dim uncompress_array() As Byte

    length = ReAry(0) * 256
    length = length + ReAry(1)
    length = length + 2 - 8 - 10

    startPD = 2 + 10

    ReDim Preserve compress_array(length - 1)
    ReDim Preserve uncompress_array(length * 5)

    Dim i As Integer
    For i = startPD To startPD + length - 1 - 2
        compress_array(i - startPD + 2) = ReAry(i)
    Next

    compress_array(0) = 120
    compress_array(1) = 156

    uncompress uncompress_array(0), UBound(uncompress_array), compress_array(0), UBound(compress_array)

    Dim strString As String
    strString = StrConv(uncompress_array, vbUnicode)
    UncompressZip = strString
End Function

Private Function UncompressBytesZip(ByRef ReAry() As Byte) As String
    Dim startVD As Integer
    Dim endVD As Integer
    Dim compress_array() As Byte
    Dim uncompress_array() As Byte
    Dim i As Integer

    startVD = ReAry(0) * 256
    startVD = startVD + ReAry(1) + 10

    endVD = ReAry(2) * 256
    endVD = endVD + ReAry(3) - 8

    Erase compress_array
    Erase uncompress_array

    ReDim compress_array(endVD - startVD + 2)
    ReDim uncompress_array((endVD - startVD + 2) * 4)

    compress_array(0) = 120
    compress_array(1) = 156

    For i = startVD To endVD
        compress_array(i - startVD + 2) = ReAry(i)
    Next

    uncompress uncompress_array(0), UBound(uncompress_array), compress_array(0), UBound(compress_array)

    UncompressBytesZip = StrConv(uncompress_array, vbUnicode)

End Function

' ============================================================================
' Geburtsdatum Format konvertieren (YYYYMMDD -> DDMMYYYY)
' ============================================================================
Private Function ChangeBirthDayFormat()
    Dim Year As String
    Year = Mid(PaGeb, 1, 4)
    Dim mouth As String
    mouth = Mid(PaGeb, 5, 2)
    Dim Day As String
    Day = Mid(PaGeb, 7, 2)
    PaGeb = Day & mouth & Year
End Function

' ============================================================================
' Properties - Rueckgabewerte
' ============================================================================
Public Property Get IsteGK() As Boolean
    IsteGK = ElGeK
End Property
Public Property Let IsteGK(Flag As Boolean)
    ElGeK = Flag
End Property

Public Property Let TerminalPort(Flag As Integer)
    TePor = Flag
End Property

Public Property Get Returnstring() As String
    Returnstring = ReStr
End Property
Public Property Let Returnstring(Flag As String)
    ReStr = Flag
End Property

Public Property Get Returnvalue() As Long
    Returnvalue = RetWe
End Property
Public Property Let Returnvalue(Flag As Long)
    RetWe = Flag
End Property

' ============================================================================
' Properties - Patientendaten
' ============================================================================
Public Property Get Versicherten_ID() As String
    Versicherten_ID = CliID
End Property
Public Property Let Versicherten_ID(Flag As String)
    CliID = Flag
End Property

Public Property Get Patient_Vorname() As String
    Patient_Vorname = PaVor
End Property
Public Property Let Patient_Vorname(Flag As String)
    PaVor = Flag
End Property

Public Property Get Patient_Name() As String
    Patient_Name = PaNam
End Property
Public Property Let Patient_Name(Flag As String)
    PaNam = Flag
End Property

Public Property Get Patient_Geboren() As String
    Patient_Geboren = PaGeb
End Property
Public Property Let Patient_Geboren(Flag As String)
    PaGeb = Flag
End Property

Public Property Get Patient_Strasse() As String
    Patient_Strasse = PaStr
End Property
Public Property Let Patient_Strasse(Flag As String)
    PaStr = Flag
End Property

Public Property Get Patient_PLZ() As String
    Patient_PLZ = PaPLZ
End Property
Public Property Let Patient_PLZ(Flag As String)
    PaPLZ = Flag
End Property

Public Property Get Patient_Ort() As String
    Patient_Ort = PaOrt
End Property
Public Property Let Patient_Ort(Flag As String)
    PaOrt = Flag
End Property

Public Property Get Patient_Hausnummer() As String
    Patient_Hausnummer = HauNu
End Property
Public Property Let Patient_Hausnummer(Flag As String)
    HauNu = Flag
End Property

Public Property Get Wohnsitzlaendercode() As String
    Wohnsitzlaendercode = WoSiC
End Property
Public Property Let Wohnsitzlaendercode(Flag As String)
    WoSiC = Flag
End Property

Public Property Get Patient_Geschlecht() As String
    Patient_Geschlecht = PaGes
End Property
Public Property Let Patient_Geschlecht(Flag As String)
    PaGes = Flag
End Property

Public Property Get Patient_Titel() As String
    Patient_Titel = PaTit
End Property
Public Property Let Patient_Titel(Flag As String)
    PaTit = Flag
End Property

Public Property Get Patient_Zusatz() As String
    Patient_Zusatz = PaZun
End Property
Public Property Let Patient_Zusatz(Flag As String)
    PaZun = Flag
End Property

Public Property Get Patient_Landeskennzeichen() As String
    Patient_Landeskennzeichen = PaLKZ
End Property
Public Property Let Patient_Landeskennzeichen(Flag As String)
    PaLKZ = Flag
End Property

' ============================================================================
' Properties - Kostentraegerdaten
' ============================================================================
Public Property Get Kostentraegerkennung() As String
    Kostentraegerkennung = kotrk
End Property
Public Property Let Kostentraegerkennung(Flag As String)
    kotrk = Flag
End Property

Public Property Get Kostentraegername() As String
    Kostentraegername = KoTrN
End Property
Public Property Let Kostentraegername(Flag As String)
    KoTrN = Flag
End Property

Public Property Get Kostentraegerlaendercode() As String
    Kostentraegerlaendercode = KTrLK
End Property
Public Property Let Kostentraegerlaendercode(Flag As String)
    KTrLK = Flag
End Property

Public Property Get Kostentraegername2() As String
    Kostentraegername2 = AKTrN
End Property
Public Property Let Kostentraegername2(Flag As String)
    AKTrN = Flag
End Property

Public Property Get Kostentraegerkennung2() As String
    Kostentraegerkennung2 = AKTrK
End Property
Public Property Let Kostentraegerkennung2(Flag As String)
    AKTrK = Flag
End Property

Public Property Get KostenTragenKennung() As String
    KostenTragenKennung = KoTraKen
End Property
Public Property Let KostenTragenKennung(Flag As String)
    KoTraKen = Flag
End Property

Public Property Get KostenTragenName() As String
    KostenTragenName = KoTraNam
End Property
Public Property Let KostenTragenName(Flag As String)
    KoTraNam = Flag
End Property

' ============================================================================
' Properties - eGK-spezifische Daten
' ============================================================================
Public Property Get Vertragsbegin() As String
    Vertragsbegin = VeBeg
End Property
Public Property Let Vertragsbegin(Flag As String)
    VeBeg = Flag
End Property

Public Property Get Rechtskreis() As String
    Rechtskreis = GKVRK
End Property
Public Property Let Rechtskreis(Flag As String)
    GKVRK = Flag
End Property

Public Property Get Versichertenart() As String
    Versichertenart = GKVVA
End Property
Public Property Let Versichertenart(Flag As String)
    GKVVA = Flag
End Property

Public Property Get Versichertenstatus() As String
    Versichertenstatus = GKVVS
End Property
Public Property Let Versichertenstatus(Flag As String)
    GKVVS = Flag
End Property

Public Property Get Ambulant() As String
    Ambulant = GKVAM
End Property
Public Property Let Ambulant(Flag As String)
    GKVAM = Flag
End Property

Public Property Get Stationaer() As String
    Stationaer = GKVST
End Property
Public Property Let Stationaer(Flag As String)
    GKVST = Flag
End Property

Public Property Get Wohnortprinzip() As String
    Wohnortprinzip = GKVWO
End Property
Public Property Let Wohnortprinzip(Flag As String)
    GKVWO = Flag
End Property

' ============================================================================
' Properties - KVK-spezifische Daten
' ============================================================================
Public Property Get Kartennummer() As String
    Kartennummer = KarNr
End Property
Public Property Let Kartennummer(Flag As String)
    KarNr = Flag
End Property

Public Property Get Kartenstatus() As String
    Kartenstatus = KaSta
End Property
Public Property Let Kartenstatus(Flag As String)
    KaSta = Flag
End Property

Public Property Get KVKStatus() As String
    KVKStatus = StaEr
End Property
Public Property Let KVKStatus(Flag As String)
    StaEr = Flag
End Property

Public Property Get Kartengueltigkeit() As String
    Kartengueltigkeit = KaGul
End Property
Public Property Let Kartengueltigkeit(Flag As String)
    KaGul = Flag
End Property

Public Property Get Kartendatum() As String
    Kartendatum = KaDat
End Property
Public Property Let Kartendatum(Flag As String)
    KaDat = Flag
End Property

' ============================================================================
' Properties - Konfiguration
' ============================================================================
Public Property Get SmartReader() As String
    SmartReader = SmRea
End Property
Public Property Let SmartReader(Flag As String)
    SmRea = Flag
End Property

