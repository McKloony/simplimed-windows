VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------
' clsGenInv - Erzeugt eine Rechnung im Format generalInvoice XML 5.0
'             (Forum Datenaustausch, generalInvoiceRequest_500.xsd,
'              Namespace http://www.forum-datenaustausch.ch/invoice)
'
' Alle Zahlen- und Datumswerte werden bewusst OHNE Format$ und OHNE CStr auf
' Gleitkommawerte erzeugt, damit die Ausgabe unabhaengig von den regionalen
' Windows-Einstellungen (Schweiz, Deutschland, Oesterreich) immer der
' XML-Spezifikation entspricht:
'   - Dezimaltrennzeichen ist immer der Punkt, kein Tausendertrennzeichen
'   - Datum        = JJJJ-MM-TT              (xs:date)
'   - Datum + Zeit = JJJJ-MM-TTThh:mm:ss     (xs:dateTime)
'
' Das XML-Escaping und die UTF-8-Kodierung uebernimmt das DOM (MSXML6);
' es werden keine XML-Fragmente per String-Verkettung gebaut.
'---------------------------------------------------------------------------

'Namespaces
Private Const NS_INVO As String = "http://www.forum-datenaustausch.ch/invoice"
Private Const NS_XSIN As String = "http://www.w3.org/2001/XMLSchema-instance"
Private Const XS_DATE As String = "generalInvoiceRequest_500.xsd"

'Platzhalter-IBAN: formal gueltig (Muster CH + 7 Ziffern + 12 Zeichen), aber
'wegen der Pruefziffern 00 kein zahlungsfaehiges Konto.
Private Const IB_LEER As String = "CH0000000000000000000"

'Indizes der Beteiligten in mPart()
Private Const PA_STEL As Integer = 0    'Rechnungssteller  (biller)
Private Const PA_DEBI As Integer = 1    'Rechnungsempfaenger / Zahlungspflichtiger (debitor)
Private Const PA_ERBR As Integer = 2    'Leistungserbringer (provider)
Private Const PA_VERS As Integer = 3    'Versicherer        (insurance)
Private Const PA_PATI As Integer = 4    'Patient            (patient)
Private Const PA_GARA As Integer = 5    'Garant             (guarantor)
Private Const PA_ZAHL As Integer = 6    'Zahlungsempfaenger (creditor)

'Feldindizes in mPart()
Private Const FE_FIRM As Integer = 0
Private Const FE_ABTL As Integer = 1
Private Const FE_ANRE As Integer = 2
Private Const FE_TITL As Integer = 3
Private Const FE_NAME As Integer = 4
Private Const FE_VORN As Integer = 5
Private Const FE_STRA As Integer = 6
Private Const FE_HAUS As Integer = 7
Private Const FE_PLZZ As Integer = 8
Private Const FE_ORTS As Integer = 9
Private Const FE_KANT As Integer = 10
Private Const FE_LAND As Integer = 11
Private Const FE_LCOD As Integer = 12
Private Const FE_TELE As Integer = 13
Private Const FE_MAIL As Integer = 14

'--- Steuerung des Requests ------------------------------------------------
Public Sprache As String            'de | fr | it
Public Modus As String              'production | test
Public GuidStr As String            '32 Hexzeichen
Public TiersTyp As String           'garant | payant | soldant
Public ErlaubtAend As Boolean       'allowModification bei tiers_payant
Public ErlaubtTS As Boolean         'allowTS bei tiers_soldant
Public Rolle As String              'z.B. naturopathictherapist, physician
Public OrtTyp As String             'practice | hospital | lab | association | company
Public RollenTitel As String        'Praxisname (role_title)
Public AnfrageTyp As String         'invoice | reminder
Public AnfrageSubTyp As String      'normal | copy | refund | storno
Public AnfrageNr As String          'Rechnungsnummer (request_id)
Public AnfrageDatum As Date         'Rechnungsdatum  (request_date)
Public Bemerkung As String          'remark
Public GeneratorName As String
Public GeneratorVers As Long
Public PaketName As String
Public Formatiert As Boolean        'XML eingerueckt ausgeben

'--- Transport -------------------------------------------------------------
Public TransportVon As String       'GLN Absender
Public TransportZu As String        'GLN Empfaenger

'--- Identifikationsnummern ------------------------------------------------
Public StelGln As String
Public StelZsr As String
Public DebiGln As String
Public DebiUid As String
Public ErbrGln As String
Public ErbrGlnOrt As String
Public ErbrZsr As String
Public VersGln As String
Public GaraGln As String

'--- Patient ---------------------------------------------------------------
Public PatGeschlecht As String      'male | female | diverse
Public PatGeburt As Date
Public PatSsn As String             'AHV-Nummer (756...) oder Patientennummer

'--- Zahlung (esrQRRed = QR-Rechnung ohne / mit Creditor Reference) ---------
Public ZahlIban As String
Public ZahlReferenz As String       'RFxx... (optional)
Public ZahlFristTage As Long
Public ZahlGrund As String

'--- Gesetz ----------------------------------------------------------------
Public GesetzTyp As String          'KVG | VVG | UVG | IVG | MVG | ORG
Public GesetzVersNr As String
Public GesetzFallNr As String
Public GesetzFallDatum As Date

'--- Behandlung ------------------------------------------------------------
Public BehVon As Date
Public BehBis As Date
Public BehKanton As String          'ZH, BE, ...
Public BehGrund As String           'disease | accident | maternity | prevention | birthdefect | unknown
Public BehArt As String             'ambulatory | stationary
Public BehPatId As String           'apid
Public BehFallId As String          'acid
Public DiagTyp As String            'ICD | cantonal | DRG | by_contract | freetext | ICPC
Public DiagCode As String
Public DiagText As String

'--- Betraege --------------------------------------------------------------
Public Waehrung As String           'ISO-Code des Rechnungsdatensatzes
Public MwStNummer As String         'CHE...
Public BetragBezahlt As Currency    'bereits bezahlt / Anzahlung
Public BetragMahnung As Currency    'Mahngebuehren
Public RabattProzent As Currency    'Rechnungsrabatt in Prozent
Public RabattText As String
Public RabattCode As String

'--- interne Daten ---------------------------------------------------------
Private mPart(6, 14) As String
Private mXmlDo As Object
Private mFehlr As String
Private mHinwe As String

Private LnTar() As String           'Tarifcode  (tariff_type)
Private LnCod() As String           'Leistungscode
Private LnNam() As String           'Bezeichnung
Private LnBem() As String           'Bemerkung
Private LnAnz() As Currency         'Anzahl
Private LnEin() As Currency         'Einzelpreis
Private LnStu() As Currency         'Steuersatz in Prozent
Private LnDat() As Date             'Leistungsdatum
Private LnCnt As Integer

Private Sub Class_Initialize()

LnCnt = 0
Sprache = "de"
Modus = "production"
TiersTyp = "garant"
Rolle = "naturopathictherapist"
OrtTyp = "practice"
AnfrageTyp = "invoice"
AnfrageSubTyp = "normal"
GeneratorName = "SimpliMed generalInvoice"
GeneratorVers = 500
PaketName = "SimpliMed"
Waehrung = "CHF"
GesetzTyp = "VVG"
BehGrund = "disease"
BehArt = "ambulatory"
DiagTyp = "freetext"
RabattText = "Rabatt"
RabattCode = "RABATT"
Formatiert = True

End Sub

Public Property Get Fehler() As String
Fehler = mFehlr
End Property

Public Property Get Hinweise() As String
Hinweise = mHinwe
End Property

Public Property Get Positionen() As Integer
Positionen = LnCnt
End Property

'---------------------------------------------------------------------------
' Adressen der Beteiligten
'---------------------------------------------------------------------------
Public Sub SetzeSteller(ByVal FirmS As String, ByVal AbtSt As String, ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)
GiSetzeAdr PA_STEL, FirmS, AbtSt, AnreS, TitlS, NamSt, VorSt, StrSt, HauSt, PlzSt, OrtSt, KanSt, LanSt, LcdSt, TelSt, EmlSt
End Sub

Public Sub SetzeDebitor(ByVal FirmS As String, ByVal AbtSt As String, ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)
GiSetzeAdr PA_DEBI, FirmS, AbtSt, AnreS, TitlS, NamSt, VorSt, StrSt, HauSt, PlzSt, OrtSt, KanSt, LanSt, LcdSt, TelSt, EmlSt
End Sub

Public Sub SetzeErbringer(ByVal FirmS As String, ByVal AbtSt As String, ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)
GiSetzeAdr PA_ERBR, FirmS, AbtSt, AnreS, TitlS, NamSt, VorSt, StrSt, HauSt, PlzSt, OrtSt, KanSt, LanSt, LcdSt, TelSt, EmlSt
End Sub

Public Sub SetzeVersicherer(ByVal FirmS As String, ByVal AbtSt As String, ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)
GiSetzeAdr PA_VERS, FirmS, AbtSt, AnreS, TitlS, NamSt, VorSt, StrSt, HauSt, PlzSt, OrtSt, KanSt, LanSt, LcdSt, TelSt, EmlSt
End Sub

Public Sub SetzePatient(ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)
GiSetzeAdr PA_PATI, vbNullString, vbNullString, AnreS, TitlS, NamSt, VorSt, StrSt, HauSt, PlzSt, OrtSt, KanSt, LanSt, LcdSt, TelSt, EmlSt
End Sub

Public Sub SetzeGarant(ByVal FirmS As String, ByVal AbtSt As String, ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)
GiSetzeAdr PA_GARA, FirmS, AbtSt, AnreS, TitlS, NamSt, VorSt, StrSt, HauSt, PlzSt, OrtSt, KanSt, LanSt, LcdSt, TelSt, EmlSt
End Sub

Public Sub SetzeZahlEmpf(ByVal FirmS As String, ByVal AbtSt As String, ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)
GiSetzeAdr PA_ZAHL, FirmS, AbtSt, AnreS, TitlS, NamSt, VorSt, StrSt, HauSt, PlzSt, OrtSt, KanSt, LanSt, LcdSt, TelSt, EmlSt
End Sub

Private Sub GiSetzeAdr(ByVal PaTyp As Integer, ByVal FirmS As String, ByVal AbtSt As String, ByVal AnreS As String, ByVal TitlS As String, ByVal NamSt As String, ByVal VorSt As String, ByVal StrSt As String, ByVal HauSt As String, ByVal PlzSt As String, ByVal OrtSt As String, ByVal KanSt As String, ByVal LanSt As String, ByVal LcdSt As String, ByVal TelSt As String, ByVal EmlSt As String)

mPart(PaTyp, FE_FIRM) = FirmS
mPart(PaTyp, FE_ABTL) = AbtSt
mPart(PaTyp, FE_ANRE) = AnreS
mPart(PaTyp, FE_TITL) = TitlS
mPart(PaTyp, FE_NAME) = NamSt
mPart(PaTyp, FE_VORN) = VorSt
mPart(PaTyp, FE_STRA) = StrSt
mPart(PaTyp, FE_HAUS) = HauSt
mPart(PaTyp, FE_PLZZ) = PlzSt
mPart(PaTyp, FE_ORTS) = OrtSt
mPart(PaTyp, FE_KANT) = KanSt
mPart(PaTyp, FE_LAND) = LanSt
mPart(PaTyp, FE_LCOD) = LcdSt
mPart(PaTyp, FE_TELE) = TelSt
mPart(PaTyp, FE_MAIL) = EmlSt

End Sub

'---------------------------------------------------------------------------
' Rechnungspositionen
'---------------------------------------------------------------------------
Public Function AddInvoiceLine(ByVal TarSt As String, ByVal CodSt As String, ByVal NamSt As String, ByVal Anzah As Currency, ByVal EinPr As Currency, ByVal StuSa As Currency, ByVal DatBe As Date, ByVal BemSt As String, ByVal NeuRe As Boolean) As Integer
'Fuegt eine Leistungsposition hinzu. Die Betraege werden so uebernommen wie
'sie in SimpliMed gespeichert sind, also inklusive Mehrwertsteuer.

If NeuRe = True Then
    LnCnt = 0
End If

ReDim Preserve LnTar(LnCnt)
ReDim Preserve LnCod(LnCnt)
ReDim Preserve LnNam(LnCnt)
ReDim Preserve LnBem(LnCnt)
ReDim Preserve LnAnz(LnCnt)
ReDim Preserve LnEin(LnCnt)
ReDim Preserve LnStu(LnCnt)
ReDim Preserve LnDat(LnCnt)

LnTar(LnCnt) = TarSt
LnCod(LnCnt) = CodSt
LnNam(LnCnt) = NamSt
LnBem(LnCnt) = BemSt
LnAnz(LnCnt) = Anzah
LnEin(LnCnt) = EinPr
LnStu(LnCnt) = StuSa
LnDat(LnCnt) = DatBe
LnCnt = LnCnt + 1

AddInvoiceLine = LnCnt

End Function

Public Sub LeerePositionen()
'Setzt Positionen, Adressen und Meldungen zurueck, damit beim Export
'mehrerer Rechnungen keine Daten der Vorrechnung stehen bleiben.

Dim AktZa As Integer
Dim AktNu As Integer

LnCnt = 0
mFehlr = vbNullString
mHinwe = vbNullString

For AktZa = 0 To 6
    For AktNu = 0 To 14
        mPart(AktZa, AktNu) = vbNullString
    Next AktNu
Next AktZa

VersGln = vbNullString
GaraGln = vbNullString
DebiUid = vbNullString
RabattProzent = 0
BetragBezahlt = 0
BetragMahnung = 0
BehVon = 0
BehBis = 0
DiagCode = vbNullString
DiagText = vbNullString
Bemerkung = vbNullString
GesetzFallDatum = 0

End Sub

'---------------------------------------------------------------------------
' Erzeugt die XML-Datei
'---------------------------------------------------------------------------
Public Function GenerateXMLFile(ByVal FiNam As String) As Boolean
On Error GoTo GiErr

Dim RootN As Object
Dim ProcN As Object
Dim TranN As Object
Dim PaylN As Object
Dim InvoN As Object
Dim BodyN As Object
Dim ProlN As Object
Dim SoftN As Object
Dim TierN As Object
Dim BillN As Object
Dim PartN As Object
Dim BalaN As Object
Dim VatsN As Object
Dim RateN As Object
Dim EsrqN As Object
Dim LawwN As Object
Dim TretN As Object
Dim DiagN As Object
Dim ServN As Object
Dim LineN As Object

Dim AktZa As Integer
Dim AktNu As Integer
Dim RecId As Long
Dim AnzSt As Integer
Dim GesBe As Currency
Dim MwSum As Currency
Dim FaeBe As Currency
Dim PosBe As Currency
Dim RabBe As Currency
Dim StuWe As Currency
Dim SumSt As Currency
Dim MwsBe As Currency
Dim StSat() As Currency
Dim StBet() As Currency
Dim StRab() As Currency
Dim TmpSt As String
Dim TmpDa As Date
Dim ReqDa As Date
Dim GebDa As Date
Dim BegDa As Date
Dim EndDa As Date
Dim IbaSt As String

GenerateXMLFile = False
mFehlr = vbNullString
mHinwe = vbNullString

ReqDa = AnfrageDatum
If ReqDa = 0 Then ReqDa = Date
GebDa = PatGeburt
If GebDa = 0 Then
    GebDa = DateSerial(1900, 1, 1)
    mHinwe = mHinwe & "Fehlendes Geburtsdatum wurde durch 01.01.1900 ersetzt." & vbCrLf
End If

'--- Pflichtangaben absichern -------------------------------------------
'Der Export bricht nicht ab. Angaben, die das Schema zwingend verlangt, im
'Datensatz aber fehlen oder nicht dazu passen, werden durch erkennbare
'Platzhalter ersetzt und als Hinweis gemeldet.
If LnCnt < 1 Then
    AddInvoiceLine vbNullString, "0", "Keine abrechenbare Leistung", 1, 0, 0, ReqDa, vbNullString, True
    mHinwe = mHinwe & "Die Rechnung enthaelt keine Leistungsposition mit Betrag;" & _
             " es wurde eine Nullposition eingesetzt." & vbCrLf
End If

If GiWaehr(Waehrung) = False Then
    mHinwe = mHinwe & "Die Rechnung ist in " & GiKurz(Waehrung, 20) & " ausgestellt." & _
             " generalInvoice 5.0 laesst nur CHF zu; die Betraege wurden unveraendert" & _
             " uebernommen und als CHF ausgewiesen." & vbCrLf
ElseIf GiKurz(Waehrung, 20) = vbNullString Then
    mHinwe = mHinwe & "Im Rechnungsdatensatz ist keine Waehrung hinterlegt;" & _
             " es wurde CHF angenommen." & vbCrLf
End If

IbaSt = GiIbanP(ZahlIban)
If IbaSt = vbNullString Then
    IbaSt = IB_LEER
    mHinwe = mHinwe & "Keine schweizerische oder liechtensteinische IBAN hinterlegt." & _
             " Im Zahlungsteil steht der Platzhalter " & IbaSt & "." & vbCrLf
End If

If GiIstGln(StelGln) = False Then
    mHinwe = mHinwe & "GLN des Mandanten ist nicht gepflegt (13 Ziffern)." & vbCrLf
End If
If GiZsrPr(StelZsr) = vbNullString Then
    mHinwe = mHinwe & "ZSR-Nummer des Mandanten ist nicht gepflegt (Buchstabe und 6 Ziffern)." & vbCrLf
End If
If GiIstMwSt(MwStNummer) = False Then
    mHinwe = mHinwe & "MWST-Nummer des Mandanten ist nicht gepflegt (CHE...)." & vbCrLf
End If
If GiKantOk(BehKanton) = False Then
    mHinwe = mHinwe & "Kanton ist nicht gesetzt (" & GiKurz(BehKanton, 20) & ");" & _
             " es wurde " & GiKanton(BehKanton) & " eingesetzt." & vbCrLf
End If

'--- Betraege und Steuergruppen ermitteln ------------------------------
ReDim StSat(LnCnt)
ReDim StBet(LnCnt)
ReDim StRab(LnCnt)
AnzSt = 0
GesBe = 0

For AktZa = 0 To LnCnt - 1
    PosBe = GiRund(LnAnz(AktZa) * LnEin(AktZa), 2)
    GesBe = GesBe + PosBe
    StuWe = GiRund(LnStu(AktZa), 2)
    If StuWe < 0 Then StuWe = 0
    If StuWe > 100 Then StuWe = 100
    AktNu = GiSteuIdx(StSat, AnzSt, StuWe)
    If AktNu < 0 Then
        StSat(AnzSt) = StuWe
        StBet(AnzSt) = PosBe
        AnzSt = AnzSt + 1
    Else
        StBet(AktNu) = StBet(AktNu) + PosBe
    End If
Next AktZa

'--- Rechnungsrabatt anteilig je Steuersatz ----------------------------
RabBe = 0
If RabattProzent > 0 And RabattProzent < 100 Then
    For AktZa = 0 To AnzSt - 1
        StRab(AktZa) = GiRund(StBet(AktZa) * RabattProzent / 100, 2)
        StBet(AktZa) = StBet(AktZa) - StRab(AktZa)
        RabBe = RabBe + StRab(AktZa)
    Next AktZa
    GesBe = GesBe - RabBe
End If

'--- Mehrwertsteuer aus den Bruttobetraegen herausrechnen --------------
MwSum = 0
For AktZa = 0 To AnzSt - 1
    If StSat(AktZa) > 0 Then
        MwsBe = GiRund(StBet(AktZa) * StSat(AktZa) / (100 + StSat(AktZa)), 2)
    Else
        MwsBe = 0
    End If
    MwSum = MwSum + MwsBe
Next AktZa

If LCase$(TiersTyp) = "garant" Then
    FaeBe = GiRapp(GesBe - BetragBezahlt + BetragMahnung)
Else
    FaeBe = GiRapp(GesBe + BetragMahnung)
End If

'--- Dokument aufbauen -------------------------------------------------
Set mXmlDo = CreateObject("MSXML2.DOMDocument.6.0")
mXmlDo.async = False
mXmlDo.preserveWhiteSpace = True
mXmlDo.appendChild mXmlDo.createProcessingInstruction("xml", "version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & "UTF-8" & Chr$(34))

Set RootN = mXmlDo.createElement("invoice:request")
RootN.setAttribute "xmlns:invoice", NS_INVO
RootN.setAttribute "xmlns:xsi", NS_XSIN
RootN.setAttribute "xsi:schemaLocation", NS_INVO & Chr$(32) & XS_DATE
RootN.setAttribute "language", GiWahl(LCase$(Sprache), "de,fr,it", "de")
RootN.setAttribute "modus", GiWahl(LCase$(Modus), "production,test", "production")
RootN.setAttribute "guid", GiGuid(GuidStr)
mXmlDo.appendChild RootN

'--- processing / transport --------------------------------------------
Set ProcN = GiKind(RootN, "invoice:processing")
Set TranN = GiKind(ProcN, "invoice:transport")
GiSetA TranN, "from", GiGlnPr(TransportVon, "Absender-GLN (transport/from)")
GiSetA TranN, "to", GiGlnPr(TransportZu, "Empfaenger-GLN (transport/to)")

'--- payload ------------------------------------------------------------
Set PaylN = GiKind(RootN, "invoice:payload")
GiSetA PaylN, "request_type", GiWahl(LCase$(AnfrageTyp), "invoice,reminder", "invoice")
GiSetA PaylN, "request_subtype", GiWahl(LCase$(AnfrageSubTyp), "normal,copy,refund,storno", "normal")

Set InvoN = GiKind(PaylN, "invoice:invoice")
GiSetA InvoN, "request_timestamp", GiUnix(ReqDa)
GiSetA InvoN, "request_date", GiDaZe(ReqDa)
TmpSt = GiKurz(AnfrageNr, 35)
If TmpSt = vbNullString Then TmpSt = "1"
GiSetA InvoN, "request_id", TmpSt

'--- body ---------------------------------------------------------------
Set BodyN = GiKind(PaylN, "invoice:body")
GiSetA BodyN, "role_title", GiKurz(RollenTitel, 35)
GiSetA BodyN, "role", GiWahl(LCase$(Rolle), "physician,physiotherapist,chiropractor,ergotherapist,nutritionist,midwife,logotherapist,hospital,rehab_clinic,psychiatric_clinic,pharmacist,dentist,labtechnician,dentaltechnician,othertechnician,psychologist,wholesaler,nursingstaff,transport,druggist,naturopathicdoctor,naturopathictherapist,other", "other")
GiSetA BodyN, "place", GiWahl(LCase$(OrtTyp), "practice,hospital,lab,association,company", "practice")

Set ProlN = GiKind(BodyN, "invoice:prolog")
If GiKurz(PaketName, 50) <> vbNullString Then
    Set SoftN = GiKind(ProlN, "invoice:package")
    GiSetA SoftN, "name", GiKurz(PaketName, 50)
End If
Set SoftN = GiKind(ProlN, "invoice:generator")
TmpSt = GiKurz(GeneratorName, 50)
If TmpSt = vbNullString Then TmpSt = "SimpliMed"
GiSetA SoftN, "name", TmpSt
GiSetA SoftN, "version", GiGanz(GeneratorVers)

If GiKurz(Bemerkung, 350) <> vbNullString Then
    GiSetT GiKind(BodyN, "invoice:remark"), GiKurz(Bemerkung, 350)
End If

'--- tiers_garant / tiers_payant ---------------------------------------
Select Case LCase$(TiersTyp)
Case "payant"
    Set TierN = GiKind(BodyN, "invoice:tiers_payant")
    If ErlaubtAend = True Then
        GiSetA TierN, "allowModification", "1"
    Else
        GiSetA TierN, "allowModification", "0"
    End If
Case "soldant"
    Set TierN = GiKind(BodyN, "invoice:tiers_soldant")
    If ErlaubtTS = True Then
        GiSetA TierN, "allowTS", "1"
    Else
        GiSetA TierN, "allowTS", "0"
    End If
Case Else
    Set TierN = GiKind(BodyN, "invoice:tiers_garant")
End Select

Set BillN = GiKind(TierN, "invoice:billers")
Set PartN = GiKind(BillN, "invoice:biller_gln")
GiSetA PartN, "gln", GiGlnPr(StelGln, "GLN des Rechnungsstellers")
GiBetei PartN, PA_STEL
If GiZsrPr(StelZsr) <> vbNullString Then
    Set PartN = GiKind(BillN, "invoice:biller_zsr")
    GiSetA PartN, "zsr", GiZsrPr(StelZsr)
    GiBetei PartN, PA_STEL
End If

Set PartN = GiKind(TierN, "invoice:debitor")
GiSetA PartN, "gln", GiGlnPr(DebiGln, "GLN des Rechnungsempfaengers")
GiSetA PartN, "uid", GiKurz(DebiUid, 15)
GiBetei PartN, PA_DEBI

Set PartN = GiKind(TierN, "invoice:providers")
Set BillN = GiKind(PartN, "invoice:provider_gln")
GiSetA BillN, "gln", GiGlnPr(ErbrGln, "GLN des Leistungserbringers")
GiSetA BillN, "gln_location", GiGlnPr(ErbrGlnOrt, "GLN des Leistungsorts")
GiBetei BillN, PA_ERBR
If GiZsrPr(ErbrZsr) <> vbNullString Then
    Set BillN = GiKind(PartN, "invoice:provider_zsr")
    GiSetA BillN, "zsr", GiZsrPr(ErbrZsr)
    GiBetei BillN, PA_ERBR
End If

If LCase$(TiersTyp) = "garant" Then
    'insurance ist bei Tiers Garant optional
    If GiIstGln(VersGln) = True Then
        Set PartN = GiKind(TierN, "invoice:insurance")
        GiSetA PartN, "gln", GiGlnPr(VersGln, "GLN des Versicherers")
        GiBetei PartN, PA_VERS
    End If
Else
    'bei Tiers Payant und Tiers Soldant ist insurance zwingend
    Set PartN = GiKind(TierN, "invoice:insurance")
    GiSetA PartN, "gln", GiGlnPr(VersGln, "GLN des Versicherers")
    GiBetei PartN, PA_VERS
End If

Set PartN = GiKind(TierN, "invoice:patient")
GiSetA PartN, "gender", GiGeschl(PatGeschlecht, False)
GiSetA PartN, "sex", GiGeschl(PatGeschlecht, True)
GiSetA PartN, "birthdate", GiDatu(GebDa)
GiSetA PartN, "ssn", GiSsnPr(PatSsn)
GiPerso PartN, PA_PATI

Set PartN = GiKind(TierN, "invoice:guarantor")
If GiIstGln(GaraGln) = True Then
    GiSetA PartN, "gln", GaraGln
End If
GiBetei PartN, PA_GARA

GiKind TierN, "invoice:partners"

Set BalaN = GiKind(TierN, "invoice:balance")
GiSetA BalaN, "currency", "CHF"
GiSetA BalaN, "amount", GiBetr(GesBe)
If BetragMahnung <> 0 Then
    GiSetA BalaN, "amount_reminder", GiBetr(BetragMahnung)
End If
If LCase$(TiersTyp) = "garant" Then
    If BetragBezahlt <> 0 Then
        GiSetA BalaN, "amount_prepaid", GiBetr(BetragBezahlt)
    End If
End If
GiSetA BalaN, "amount_due", GiBetr(FaeBe)

Set VatsN = GiKind(BalaN, "invoice:vat")
GiSetA VatsN, "vat", GiBetr(MwSum)
If GiMwstNr(MwStNummer) <> vbNullString Then
    GiSetA VatsN, "vat_number", GiMwstNr(MwStNummer)
End If
SumSt = 0
For AktZa = 0 To AnzSt - 1
    If StSat(AktZa) > 0 Then
        MwsBe = GiRund(StBet(AktZa) * StSat(AktZa) / (100 + StSat(AktZa)), 2)
    Else
        MwsBe = 0
    End If
    SumSt = SumSt + MwsBe
    Set RateN = GiKind(VatsN, "invoice:vat_rate")
    GiSetA RateN, "vat_rate", GiDezi(StSat(AktZa), 2)
    GiSetA RateN, "amount", GiBetr(StBet(AktZa))
    GiSetA RateN, "vat", GiBetr(MwsBe)
Next AktZa

'--- Zahlungsteil (QR-Rechnung) ----------------------------------------
Set EsrqN = GiKind(BodyN, "invoice:esrQRRed")
GiSetA EsrqN, "subtype", "esrQRRed"
GiSetA EsrqN, "iban", IbaSt
If GiRefPr(ZahlReferenz) <> vbNullString Then
    GiSetA EsrqN, "reference_number", GiRefPr(ZahlReferenz)
End If
If ZahlFristTage > 0 Then
    GiSetA EsrqN, "payment_period", "P" & GiGanz(ZahlFristTage) & "D"
End If
GiSetA EsrqN, "payment_reason", GiKurz(ZahlGrund, 140)
Set PartN = GiKind(EsrqN, "invoice:creditor")
If mPart(PA_ZAHL, FE_FIRM) = vbNullString And mPart(PA_ZAHL, FE_NAME) = vbNullString Then
    GiBetei PartN, PA_STEL
Else
    GiBetei PartN, PA_ZAHL
End If

'--- law ----------------------------------------------------------------
Set LawwN = GiKind(BodyN, "invoice:law")
GiSetA LawwN, "type", GiWahl(UCase$(GesetzTyp), "KVG,VVG,UVG,IVG,MVG,ORG", "VVG")
GiSetA LawwN, "insured_id", GiKurz(GesetzVersNr, 35)
GiSetA LawwN, "case_id", GiKurz(GesetzFallNr, 35)
If GesetzFallDatum > 0 Then
    GiSetA LawwN, "case_date", GiDatu(GesetzFallDatum)
End If

'--- treatment ----------------------------------------------------------
Set TretN = GiKind(BodyN, "invoice:treatment")
BegDa = BehVon
EndDa = BehBis
For AktZa = 0 To LnCnt - 1
    If LnDat(AktZa) > 0 Then
        If BehVon = 0 Then
            If BegDa = 0 Then
                BegDa = LnDat(AktZa)
            ElseIf LnDat(AktZa) < BegDa Then
                BegDa = LnDat(AktZa)
            End If
        End If
        If BehBis = 0 Then
            If LnDat(AktZa) > EndDa Then EndDa = LnDat(AktZa)
        End If
    End If
Next AktZa
If BegDa = 0 Then BegDa = ReqDa
If EndDa = 0 Then EndDa = BegDa
If EndDa < BegDa Then EndDa = BegDa
GiSetA TretN, "date_begin", GiDatu(BegDa)
GiSetA TretN, "date_end", GiDatu(EndDa)
GiSetA TretN, "canton", GiKanton(BehKanton)
GiSetA TretN, "treatment", GiWahl(LCase$(BehArt), "ambulatory,stationary", "ambulatory")
GiSetA TretN, "reason", GiWahl(LCase$(BehGrund), "disease,accident,maternity,prevention,birthdefect,unknown", "disease")
GiSetA TretN, "apid", GiKurz(BehPatId, 35)
GiSetA TretN, "acid", GiKurz(BehFallId, 35)
If GiKurz(DiagCode, 12) <> vbNullString Or GiKurz(DiagText, 350) <> vbNullString Then
    Set DiagN = GiKind(TretN, "invoice:diagnosis")
    GiSetA DiagN, "type", GiWahl(DiagTyp, "ICD,cantonal,DRG,by_contract,freetext,birthdefect,ICPC", "freetext")
    GiSetA DiagN, "code", GiKurz(DiagCode, 12)
    GiSetT DiagN, GiKurz(DiagText, 350)
End If

'--- services -----------------------------------------------------------
Set ServN = GiKind(BodyN, "invoice:services")
RecId = 0
For AktZa = 0 To LnCnt - 1
    RecId = RecId + 1
    PosBe = GiRund(LnAnz(AktZa) * LnEin(AktZa), 2)
    Set LineN = GiKind(ServN, "invoice:service")
    GiSetA LineN, "record_id", GiGanz(RecId)
    GiSetA LineN, "tariff_type", GiTarif(LnTar(AktZa))
    TmpSt = GiKurz(LnCod(AktZa), 30)
    If TmpSt = vbNullString Then TmpSt = GiGanz(RecId)
    GiSetA LineN, "code", TmpSt
    TmpSt = GiKurz(LnNam(AktZa), 350)
    If TmpSt = vbNullString Then TmpSt = "Leistung"
    GiSetA LineN, "name", TmpSt
    'Menge und Einzelpreis mit vier Nachkommastellen, damit
    'quantity * unit den ausgewiesenen amount ergibt
    GiSetA LineN, "quantity", GiDezi(LnAnz(AktZa), 4)
    TmpDa = LnDat(AktZa)
    If TmpDa = 0 Then TmpDa = ReqDa
    GiSetA LineN, "date_begin", GiDaZe(TmpDa)
    GiSetA LineN, "provider_id", GiGlnPr(ErbrGln, vbNullString)
    GiSetA LineN, "responsible_id", GiGlnPr(ErbrGln, vbNullString)
    GiSetA LineN, "unit", GiDezi(LnEin(AktZa), 4)
    GiSetA LineN, "unit_factor", "1"
    GiSetA LineN, "amount", GiBetr(PosBe)
    StuWe = LnStu(AktZa)          'Das Schema laesst nur 0 bis 100 zu
    If StuWe < 0 Then StuWe = 0
    If StuWe > 100 Then StuWe = 100
    If StuWe > 0 Then
        GiSetA LineN, "vat_rate", GiDezi(StuWe, 2)
    End If
    GiSetA LineN, "remark", GiKurz(LnBem(AktZa), 350)
Next AktZa

If RabBe <> 0 Then
    For AktZa = 0 To AnzSt - 1
        MwsBe = StRab(AktZa)
        If MwsBe <> 0 Then
            RecId = RecId + 1
            Set LineN = GiKind(ServN, "invoice:service")
            GiSetA LineN, "record_id", GiGanz(RecId)
            GiSetA LineN, "tariff_type", GiTarif(LnTar(0))
            TmpSt = GiKurz(RabattCode, 30)
            If TmpSt = vbNullString Then TmpSt = "RABATT"
            GiSetA LineN, "code", TmpSt
            TmpSt = GiKurz(RabattText, 350)
            If TmpSt = vbNullString Then TmpSt = "Rabatt"
            GiSetA LineN, "name", TmpSt
            GiSetA LineN, "quantity", "1"
            GiSetA LineN, "date_begin", GiDaZe(ReqDa)
            GiSetA LineN, "provider_id", GiGlnPr(ErbrGln, vbNullString)
            GiSetA LineN, "responsible_id", GiGlnPr(ErbrGln, vbNullString)
            GiSetA LineN, "unit", GiDezi(-MwsBe, 2)
            GiSetA LineN, "unit_factor", "1"
            GiSetA LineN, "amount", GiBetr(-MwsBe)
            If StSat(AktZa) > 0 Then
                GiSetA LineN, "vat_rate", GiDezi(StSat(AktZa), 2)
            End If
        End If
    Next AktZa
End If

If Formatiert = True Then
    GiEinru RootN, 0
End If

mXmlDo.Save FiNam
GenerateXMLFile = True

Set mXmlDo = Nothing

Exit Function

GiErr:
mFehlr = Err.Description & Chr$(32) & "(" & GiGanz(Err.Number) & ")"
Set mXmlDo = Nothing
GenerateXMLFile = False

End Function

'---------------------------------------------------------------------------
' Hilfsroutinen zum Aufbau des DOM
'---------------------------------------------------------------------------
Private Function GiKind(Eltrn As Object, ByVal TagNa As String) As Object

Dim Knote As Object

Set Knote = mXmlDo.createElement(TagNa)
Eltrn.appendChild Knote
Set GiKind = Knote

End Function

Private Sub GiSetA(Knote As Object, ByVal AtNam As String, ByVal AtWer As String)

If AtWer <> vbNullString Then
    Knote.setAttribute AtNam, AtWer
End If

End Sub

Private Sub GiSetT(Knote As Object, ByVal TxWer As String)

Knote.Text = TxWer

End Sub

Private Sub GiBetei(Eltrn As Object, ByVal PaTyp As Integer)
'Erzeugt company oder person, je nachdem was gefuellt ist

If mPart(PaTyp, FE_FIRM) <> vbNullString Then
    GiFirma Eltrn, PaTyp
Else
    GiPerso Eltrn, PaTyp
End If

End Sub

Private Sub GiFirma(Eltrn As Object, ByVal PaTyp As Integer)

Dim FirmN As Object
Dim TmpSt As String

Set FirmN = GiKind(Eltrn, "invoice:company")
TmpSt = GiKurz(mPart(PaTyp, FE_FIRM), 35)
If TmpSt = vbNullString Then TmpSt = "-"
GiSetT GiKind(FirmN, "invoice:companyname"), TmpSt
If GiKurz(mPart(PaTyp, FE_ABTL), 35) <> vbNullString Then
    GiSetT GiKind(FirmN, "invoice:department"), GiKurz(mPart(PaTyp, FE_ABTL), 35)
End If
GiAdres FirmN, PaTyp
GiTelec FirmN, PaTyp
GiOnlin FirmN, PaTyp

End Sub

Private Sub GiPerso(Eltrn As Object, ByVal PaTyp As Integer)

Dim PersN As Object
Dim TmpSt As String

Set PersN = GiKind(Eltrn, "invoice:person")
GiSetA PersN, "salutation", GiKurz(mPart(PaTyp, FE_ANRE), 35)
GiSetA PersN, "title", GiKurz(mPart(PaTyp, FE_TITL), 35)
TmpSt = GiKurz(mPart(PaTyp, FE_NAME), 35)
If TmpSt = vbNullString Then TmpSt = "-"
GiSetT GiKind(PersN, "invoice:familyname"), TmpSt
TmpSt = GiKurz(mPart(PaTyp, FE_VORN), 35)
If TmpSt = vbNullString Then TmpSt = "-"
GiSetT GiKind(PersN, "invoice:givenname"), TmpSt
GiAdres PersN, PaTyp
GiTelec PersN, PaTyp
GiOnlin PersN, PaTyp

End Sub

Private Sub GiAdres(Eltrn As Object, ByVal PaTyp As Integer)

Dim PostN As Object
Dim StraN As Object
Dim ZipsN As Object
Dim LandN As Object
Dim TmpSt As String

Set PostN = GiKind(Eltrn, "invoice:postal")

If GiKurz(mPart(PaTyp, FE_STRA), 35) <> vbNullString Then
    Set StraN = GiKind(PostN, "invoice:street")
    GiSetA StraN, "street_name", GiKurz(mPart(PaTyp, FE_STRA), 35)
    GiSetA StraN, "house_no", GiKurz(mPart(PaTyp, FE_HAUS), 10)
    If GiKurz(mPart(PaTyp, FE_HAUS), 10) <> vbNullString Then
        TmpSt = GiKurz(mPart(PaTyp, FE_STRA) & Chr$(32) & mPart(PaTyp, FE_HAUS), 35)
    Else
        TmpSt = GiKurz(mPart(PaTyp, FE_STRA), 35)
    End If
    GiSetT StraN, TmpSt
End If

Set ZipsN = GiKind(PostN, "invoice:zip")
GiSetA ZipsN, "state_code", GiStaCd(mPart(PaTyp, FE_KANT))
TmpSt = GiKurz(mPart(PaTyp, FE_PLZZ), 9)
If TmpSt = vbNullString Then
    TmpSt = "0000"
    mHinwe = mHinwe & "Fehlende Postleitzahl wurde durch 0000 ersetzt." & vbCrLf
End If
GiSetT ZipsN, TmpSt

TmpSt = GiKurz(mPart(PaTyp, FE_ORTS), 35)
If TmpSt = vbNullString Then
    TmpSt = "-"
    mHinwe = mHinwe & "Fehlender Ort wurde durch - ersetzt." & vbCrLf
End If
GiSetT GiKind(PostN, "invoice:city"), TmpSt

If GiKurz(mPart(PaTyp, FE_LAND), 35) <> vbNullString Then
    Set LandN = GiKind(PostN, "invoice:country")
    GiSetA LandN, "country_code", GiLandC(mPart(PaTyp, FE_LCOD))
    GiSetT LandN, GiKurz(mPart(PaTyp, FE_LAND), 35)
End If

End Sub

Private Sub GiTelec(Eltrn As Object, ByVal PaTyp As Integer)

Dim TeleN As Object

If GiKurz(mPart(PaTyp, FE_TELE), 25) <> vbNullString Then
    Set TeleN = GiKind(Eltrn, "invoice:telecom")
    GiSetT GiKind(TeleN, "invoice:phone"), GiKurz(mPart(PaTyp, FE_TELE), 25)
End If

End Sub

Private Sub GiOnlin(Eltrn As Object, ByVal PaTyp As Integer)

Dim OnliN As Object
Dim TmpSt As String

TmpSt = GiKurz(mPart(PaTyp, FE_MAIL), 80)
If TmpSt <> vbNullString Then
    If InStr(1, TmpSt, "@", 1) > 1 Then
        If InStr(1, TmpSt, "@", 1) < Len(TmpSt) Then
            Set OnliN = GiKind(Eltrn, "invoice:online")
            GiSetT GiKind(OnliN, "invoice:email"), TmpSt
        End If
    End If
End If

End Sub

Private Sub GiEinru(Knote As Object, ByVal Ebene As Integer)
'Rueckt nur Elemente ein, die ausschliesslich Elementkinder besitzen.
'Textinhalte bleiben dadurch unveraendert.

Dim AktZa As Integer
Dim KinAn As Integer
Dim NurEl As Boolean
Dim Kinde As Object

KinAn = Knote.childNodes.length
If KinAn = 0 Then Exit Sub

NurEl = True
For AktZa = 0 To KinAn - 1
    If Knote.childNodes.Item(AktZa).nodeType <> 1 Then
        NurEl = False
        Exit For
    End If
Next AktZa
If NurEl = False Then Exit Sub

For AktZa = KinAn - 1 To 0 Step -1
    Set Kinde = Knote.childNodes.Item(AktZa)
    GiEinru Kinde, Ebene + 1
    Knote.insertBefore mXmlDo.createTextNode(vbCrLf & GiTabs(Ebene + 1)), Kinde
Next AktZa
Knote.appendChild mXmlDo.createTextNode(vbCrLf & GiTabs(Ebene))

End Sub

Private Function GiTabs(ByVal Anzah As Integer) As String

Dim AktZa As Integer
Dim TmpSt As String

For AktZa = 1 To Anzah
    TmpSt = TmpSt & vbTab
Next AktZa
GiTabs = TmpSt

End Function

'---------------------------------------------------------------------------
' Zahlen- und Datumsformatierung (bewusst ohne Format$, damit die Ausgabe
' nicht von den regionalen Windows-Einstellungen abhaengt)
'---------------------------------------------------------------------------
Private Function GiGanz(ByVal Wert As Long) As String
'Ganzzahl ohne Trennzeichen; CStr ist bei Long immer landesunabhaengig

GiGanz = CStr(Wert)

End Function

Private Function GiFuel(ByVal Wert As Long, ByVal Stell As Integer) As String
'Ganzzahl mit fuehrenden Nullen

Dim TmpSt As String

TmpSt = CStr(Abs(Wert))
Do While Len(TmpSt) < Stell
    TmpSt = "0" & TmpSt
Loop
If Wert < 0 Then TmpSt = "-" & TmpSt
GiFuel = TmpSt

End Function

Private Function GiRund(ByVal Wert As Currency, ByVal Stell As Integer) As Currency
'Kaufmaennische Rundung (0.5 wird immer vom Nullpunkt weg gerundet).
'VB6 Round() rundet bankmaessig und ist fuer Geldbetraege ungeeignet.

Dim Fakt As Currency

Select Case Stell
Case 0:    Fakt = 1
Case 1:    Fakt = 10
Case 2:    Fakt = 100
Case 3:    Fakt = 1000
Case Else: Fakt = 10000
End Select

If Wert < 0 Then
    GiRund = -(Int((-Wert) * Fakt + 0.5) / Fakt)
Else
    GiRund = Int(Wert * Fakt + 0.5) / Fakt
End If

End Function

Private Function GiRapp(ByVal Wert As Currency) As Currency
'Schweizer Rappenrundung auf 5 Rappen

If Wert < 0 Then
    GiRapp = -(Int((-Wert) * 20 + 0.5) / 20)
Else
    GiRapp = Int(Wert * 20 + 0.5) / 20
End If

End Function

Private Function GiBetr(ByVal Wert As Currency) As String
'Geldbetrag mit genau zwei Nachkommastellen und Punkt als Trennzeichen

Dim Rest As Currency
Dim Ganz As Currency
Dim Nach As Long
Dim VorZe As String

Rest = GiRund(Wert, 2)
If Rest < 0 Then
    VorZe = "-"
    Rest = -Rest
End If
Ganz = Int(Rest)
Nach = CLng(Int((Rest - Ganz) * 100 + 0.5))
If Nach > 99 Then
    Nach = 0
    Ganz = Ganz + 1
End If
GiBetr = VorZe & GiGanz(CLng(Ganz)) & "." & GiFuel(Nach, 2)

End Function

Private Function GiDezi(ByVal Wert As Currency, ByVal Stell As Integer) As String
'Dezimalzahl mit hoechstens <Stell> Nachkommastellen, ohne ueberfluessige Nullen

Dim Rest As Currency
Dim Ganz As Currency
Dim Nach As Long
Dim Fakt As Long
Dim VorZe As String
Dim NacSt As String

If Stell < 0 Then Stell = 0
If Stell > 4 Then Stell = 4

Select Case Stell
Case 0:    Fakt = 1
Case 1:    Fakt = 10
Case 2:    Fakt = 100
Case 3:    Fakt = 1000
Case Else: Fakt = 10000
End Select

Rest = GiRund(Wert, Stell)
If Rest < 0 Then
    VorZe = "-"
    Rest = -Rest
End If
Ganz = Int(Rest)
Nach = CLng(Int((Rest - Ganz) * Fakt + 0.5))
If Nach >= Fakt Then
    Nach = 0
    Ganz = Ganz + 1
End If

If Stell = 0 Or Nach = 0 Then
    GiDezi = VorZe & GiGanz(CLng(Ganz))
    Exit Function
End If

NacSt = GiFuel(Nach, Stell)
Do While Len(NacSt) > 1
    If Right$(NacSt, 1) <> "0" Then Exit Do
    NacSt = Left$(NacSt, Len(NacSt) - 1)
Loop
GiDezi = VorZe & GiGanz(CLng(Ganz)) & "." & NacSt

End Function

Private Function GiDatu(ByVal Wert As Date) As String
'xs:date, immer JJJJ-MM-TT

If Wert = 0 Then
    GiDatu = vbNullString
    Exit Function
End If
GiDatu = GiFuel(Year(Wert), 4) & "-" & GiFuel(Month(Wert), 2) & "-" & GiFuel(Day(Wert), 2)

End Function

Private Function GiDaZe(ByVal Wert As Date) As String
'xs:dateTime, immer JJJJ-MM-TThh:mm:ss

If Wert = 0 Then
    GiDaZe = vbNullString
    Exit Function
End If
GiDaZe = GiDatu(Wert) & "T" & GiFuel(Hour(Wert), 2) & ":" & GiFuel(Minute(Wert), 2) & ":" & GiFuel(Second(Wert), 2)

End Function

Private Function GiUnix(ByVal Wert As Date) As String
'Sekunden seit 01.01.1970; das Schema laesst 1420066800 bis 2890245600 zu

Dim Sekun As Double
Dim Hochs As Double
Dim Niedr As Long

If Wert = 0 Then
    Sekun = (CDbl(Now) - 25569#) * 86400#
Else
    Sekun = (CDbl(Wert) - 25569#) * 86400#
End If
Sekun = Int(Sekun + 0.5)
If Sekun < 1420066800# Then Sekun = 1420066800#
If Sekun > 2890245600# Then Sekun = 2890245600#

Hochs = Int(Sekun / 1000000#)
Niedr = CLng(Sekun - Hochs * 1000000#)
GiUnix = GiGanz(CLng(Hochs)) & GiFuel(Niedr, 6)

End Function

'---------------------------------------------------------------------------
' Pruef- und Umsetzroutinen
'---------------------------------------------------------------------------
Private Function GiKurz(ByVal TxWer As String, ByVal MaxLa As Integer) As String
'Entfernt Steuerzeichen, kuerzt auf die im Schema erlaubte Laenge

Dim AktZa As Long
Dim CodeZ As Long
Dim Zeich As String
Dim TmpSt As String

If TxWer = vbNullString Then
    GiKurz = vbNullString
    Exit Function
End If

For AktZa = 1 To Len(TxWer)
    Zeich = Mid$(TxWer, AktZa, 1)
    CodeZ = AscW(Zeich)
    If CodeZ < 0 Then
        CodeZ = CodeZ + 65536    'AscW liefert oberhalb U+7FFF negative Werte
    End If
    If CodeZ < 32 Then
        If Len(TmpSt) > 0 Then
            If Right$(TmpSt, 1) <> Chr$(32) Then TmpSt = TmpSt & Chr$(32)
        End If
    ElseIf CodeZ = 127 Then
        'weglassen
    Else
        TmpSt = TmpSt & Zeich
    End If
Next AktZa

TmpSt = Trim$(TmpSt)
If Len(TmpSt) > MaxLa Then
    TmpSt = Trim$(Left$(TmpSt, MaxLa))
End If
GiKurz = TmpSt

End Function

Private Function GiWahl(ByVal Wert As String, ByVal Liste As String, ByVal StaWe As String) As String
'Liefert Wert nur zurueck, wenn er in der Aufzaehlung des Schemas enthalten ist

If Wert <> vbNullString Then
    If InStr(1, "," & Liste & ",", "," & Wert & ",", 1) > 0 Then
        GiWahl = Wert
        Exit Function
    End If
End If
GiWahl = StaWe

End Function

Private Function GiZiffn(ByVal TxWer As String) As String
'Liefert nur die Ziffern eines Strings

Dim AktZa As Long
Dim Zeich As String
Dim TmpSt As String

For AktZa = 1 To Len(TxWer)
    Zeich = Mid$(TxWer, AktZa, 1)
    If Zeich >= "0" And Zeich <= "9" Then
        TmpSt = TmpSt & Zeich
    End If
Next AktZa
GiZiffn = TmpSt

End Function

Public Function GiIstGln(ByVal TxWer As String) As Boolean
'Formatpruefung gemaess Schema: genau 13 Ziffern.
'Die Pruefziffer wird bewusst NICHT erzwungen, da auch offizielle
'Referenzdaten GLN mit abweichender Pruefziffer enthalten.

GiIstGln = (Len(GiZiffn(TxWer)) = 13)

End Function

Public Function GiPruefZi(ByVal TxWer As String) As Boolean
'Prueft die EAN-13 Pruefziffer einer GLN

Dim AktZa As Integer
Dim Ziffe As Integer
Dim Summe As Long
Dim TmpSt As String

GiPruefZi = False
TmpSt = GiZiffn(TxWer)
If Len(TmpSt) <> 13 Then Exit Function

Summe = 0
For AktZa = 1 To 12
    Ziffe = CInt(Mid$(TmpSt, AktZa, 1))
    If AktZa Mod 2 = 0 Then
        Summe = Summe + Ziffe * 3
    Else
        Summe = Summe + Ziffe
    End If
Next AktZa

If ((10 - (Summe Mod 10)) Mod 10) = CInt(Mid$(TmpSt, 13, 1)) Then
    GiPruefZi = True
End If

End Function

Public Function GiBaueGln(ByVal Numme As String) As String
'Bildet aus einer beliebigen Nummer eine formal gueltige 13-stellige GLN.
'Das Praefix 20 kennzeichnet eine nicht registrierte (interne) Nummer.

Dim AktZa As Integer
Dim Ziffe As Integer
Dim Summe As Long
Dim Basis As String
Dim TmpSt As String

TmpSt = GiZiffn(Numme)
If Len(TmpSt) > 10 Then
    TmpSt = Right$(TmpSt, 10)
End If
Do While Len(TmpSt) < 10
    TmpSt = "0" & TmpSt
Loop
Basis = "20" & TmpSt

Summe = 0
For AktZa = 1 To 12
    Ziffe = CInt(Mid$(Basis, AktZa, 1))
    If AktZa Mod 2 = 0 Then
        Summe = Summe + Ziffe * 3
    Else
        Summe = Summe + Ziffe
    End If
Next AktZa

GiBaueGln = Basis & CStr((10 - (Summe Mod 10)) Mod 10)

End Function

Private Function GiGlnPr(ByVal TxWer As String, ByVal FelNa As String) As String
'Liefert eine formal gueltige GLN. Eine hinterlegte 13-stellige Nummer wird
'immer unveraendert uebernommen; nur fehlende oder zu kurze Nummern werden
'durch eine interne Ersatz-GLN ersetzt.

Dim TmpSt As String

TmpSt = GiZiffn(TxWer)

If Len(TmpSt) = 13 Then
    If GiPruefZi(TmpSt) = False And FelNa <> vbNullString Then
        mHinwe = mHinwe & FelNa & ": Pruefziffer der GLN " & TmpSt & " stimmt nicht." & vbCrLf
    End If
    GiGlnPr = TmpSt
    Exit Function
End If

If FelNa <> vbNullString Then
    mHinwe = mHinwe & FelNa & " fehlt oder ist unvollstaendig und wurde durch eine interne GLN ersetzt." & vbCrLf
End If
'Auch bei voellig fehlender Nummer wird eine interne Ersatz-GLN gebildet.
'Die offizielle Kennung 2006666666008 ("kein Empfaenger") wird bewusst nicht
'eingesetzt, damit sie ihre Bedeutung im Transportteil behaelt.
GiGlnPr = GiBaueGln(TmpSt)

End Function

Private Function GiZsrPr(ByVal TxWer As String) As String
'ZSR-Nummer im Format Buchstabe + 6 Ziffern

Dim ErsSt As String
Dim ZifSt As String
Dim TmpSt As String

TmpSt = UCase$(Trim$(TxWer))
If Len(TmpSt) < 2 Then
    GiZsrPr = vbNullString
    Exit Function
End If

ErsSt = Left$(TmpSt, 1)
If ErsSt < "A" Or ErsSt > "Z" Then
    GiZsrPr = vbNullString
    Exit Function
End If

ZifSt = GiZiffn(Mid$(TmpSt, 2))
If Len(ZifSt) <> 6 Then
    GiZsrPr = vbNullString
    Exit Function
End If

GiZsrPr = ErsSt & ZifSt

End Function

Private Function GiSsnPr(ByVal TxWer As String) As String
'Versichertennummer: AHV 756..., 438... oder 4 bis 10 Ziffern

Dim TmpSt As String

TmpSt = GiZiffn(TxWer)
If Len(TmpSt) = 13 Then
    If Left$(TmpSt, 3) = "756" Or Left$(TmpSt, 3) = "438" Then
        GiSsnPr = TmpSt
        Exit Function
    End If
End If

If Len(TmpSt) > 10 Then TmpSt = Right$(TmpSt, 10)
Do While Len(TmpSt) < 4
    TmpSt = "0" & TmpSt
Loop
GiSsnPr = TmpSt

End Function

Private Function GiIbanP(ByVal TxWer As String) As String
'IBAN ohne Leerzeichen; das Schema laesst nur CH und LI zu

Dim AktZa As Long
Dim Zeich As String
Dim TmpSt As String

For AktZa = 1 To Len(TxWer)
    Zeich = UCase$(Mid$(TxWer, AktZa, 1))
    If (Zeich >= "0" And Zeich <= "9") Or (Zeich >= "A" And Zeich <= "Z") Then
        TmpSt = TmpSt & Zeich
    End If
Next AktZa

If Len(TmpSt) = 21 Then
    If Left$(TmpSt, 2) = "CH" Or Left$(TmpSt, 2) = "LI" Then
        If GiZiffn(Mid$(TmpSt, 3, 7)) = Mid$(TmpSt, 3, 7) Then
            GiIbanP = TmpSt
            Exit Function
        End If
    End If
End If


GiIbanP = vbNullString

End Function

Private Function GiRefPr(ByVal TxWer As String) As String
'Creditor Reference RFxx... fuer esrQRRed

Dim AktZa As Long
Dim Zeich As String
Dim TmpSt As String

For AktZa = 1 To Len(TxWer)
    Zeich = UCase$(Mid$(TxWer, AktZa, 1))
    If (Zeich >= "0" And Zeich <= "9") Or (Zeich >= "A" And Zeich <= "Z") Then
        TmpSt = TmpSt & Zeich
    End If
Next AktZa

If Len(TmpSt) < 5 Or Len(TmpSt) > 25 Then
    GiRefPr = vbNullString
    Exit Function
End If
If Left$(TmpSt, 2) <> "RF" Then
    GiRefPr = vbNullString
    Exit Function
End If
If GiZiffn(Mid$(TmpSt, 3, 2)) <> Mid$(TmpSt, 3, 2) Then
    GiRefPr = vbNullString
    Exit Function
End If
GiRefPr = TmpSt

End Function

Public Function GiIstMwSt(ByVal TxWer As String) As Boolean
'Prueft, ob eine gueltige schweizerische MWST-Nummer vorliegt

GiIstMwSt = (GiMwstNr(TxWer) <> vbNullString)

End Function

Private Function GiMwstNr(ByVal TxWer As String) As String
'Schweizer MWST-Nummer im Format CHE123456789

Dim TmpSt As String
Dim ZifSt As String

TmpSt = UCase$(Replace(Trim$(TxWer), Chr$(32), vbNullString))
TmpSt = Replace(TmpSt, ".", vbNullString)
TmpSt = Replace(TmpSt, "-", vbNullString)
If Len(TmpSt) < 3 Then
    GiMwstNr = vbNullString
    Exit Function
End If
If Left$(TmpSt, 3) <> "CHE" Then
    GiMwstNr = vbNullString
    Exit Function
End If
ZifSt = GiZiffn(Mid$(TmpSt, 4))
If Len(ZifSt) < 9 Then
    GiMwstNr = vbNullString
    Exit Function
End If
GiMwstNr = "CHE" & Left$(ZifSt, 9)

End Function

Private Function GiGuid(ByVal TxWer As String) As String
'32 Hexzeichen; erzeugt ersatzweise eine Nummer aus Datum und Zeit

Dim AktZa As Long
Dim Zeich As String
Dim TmpSt As String

For AktZa = 1 To Len(TxWer)
    Zeich = UCase$(Mid$(TxWer, AktZa, 1))
    If (Zeich >= "0" And Zeich <= "9") Or (Zeich >= "A" And Zeich <= "F") Then
        TmpSt = TmpSt & Zeich
    End If
Next AktZa

If Len(TmpSt) >= 32 Then
    GiGuid = Left$(TmpSt, 32)
    Exit Function
End If

TmpSt = TmpSt & GiFuel(Year(Now), 4) & GiFuel(Month(Now), 2) & GiFuel(Day(Now), 2)
TmpSt = TmpSt & GiFuel(Hour(Now), 2) & GiFuel(Minute(Now), 2) & GiFuel(Second(Now), 2)
Do While Len(TmpSt) < 32
    TmpSt = TmpSt & "0"
Loop
GiGuid = Left$(TmpSt, 32)

End Function

Private Function GiGeschl(ByVal TxWer As String, ByVal NurBi As Boolean) As String
'gender laesst male, female und diverse zu, sex nur male und female.
'Der Quelltext bleibt ASCII; Schreibweisen mit Umlaut (maennlich, weiblich)
'werden deshalb ueber den ersten Buchstaben erkannt.

Dim TmpSt As String
Dim KenZe As String

TmpSt = LCase$(Trim$(TxWer))

Select Case TmpSt
Case "male", "maennlich", "mann", "herr", "1"
    KenZe = "m"
Case "female", "weiblich", "frau", "2"
    KenZe = "w"
Case "diverse", "divers", "3"
    KenZe = "d"
Case vbNullString
    KenZe = "?"
Case Else
    KenZe = Left$(TmpSt, 1)
End Select

Select Case KenZe
Case "m", "h"
    GiGeschl = "male"
Case "d", "x"
    If NurBi = True Then
        GiGeschl = "female"
    Else
        GiGeschl = "diverse"
    End If
Case "w", "f"
    GiGeschl = "female"
Case Else
    GiGeschl = "female"
    If NurBi = False Then
        mHinwe = mHinwe & "Geschlecht des Patienten fehlt oder ist unbekannt (" & _
                 GiKurz(TxWer, 20) & "); es wurde female eingesetzt." & vbCrLf
    End If
End Select

End Function

Private Function GiStaCd(ByVal TxWer As String) As String
'Kuerzel fuer zip/@state_code. Bekannte Kantone werden normalisiert, weil
'SimpliMed Basel-Stadt intern als "BuSe" fuehrt. Andere Werte bleiben stehen,
'damit auslaendische Adressen ihre Regionalangabe behalten.

Dim TmpSt As String

TmpSt = UCase$(Trim$(TxWer))
If TmpSt = "BUSE" Then TmpSt = "BS"
If GiKantOk(TmpSt) = True Then
    GiStaCd = TmpSt
Else
    GiStaCd = GiKurz(TxWer, 9)
End If

End Function

Private Function GiKantOk(ByVal TxWer As String) As Boolean
'Prueft, ob ein verwertbares Kantonskuerzel hinterlegt ist

Dim TmpSt As String

TmpSt = UCase$(Trim$(TxWer))
If TmpSt = "BUSE" Then TmpSt = "BS"
GiKantOk = (GiWahl(TmpSt, "AG,AI,AR,BE,BL,BS,FR,GE,GL,GR,JU,LU,NE,NW,OW,SG,SH,SO,SZ,TI,TG,UR,VD,VS,ZG,ZH,LI,A,D,F,I", vbNullString) <> vbNullString)

End Function

Private Function GiKanton(ByVal TxWer As String) As String
'Kantonskuerzel gemaess Aufzaehlung im Schema

Dim TmpSt As String

TmpSt = UCase$(Trim$(TxWer))
If TmpSt = "BUSE" Then TmpSt = "BS"
GiKanton = GiWahl(TmpSt, "AG,AI,AR,BE,BL,BS,FR,GE,GL,GR,JU,LU,NE,NW,OW,SG,SH,SO,SZ,TI,TG,UR,VD,VS,ZG,ZH,LI,A,D,F,I", "ZH")

End Function

Private Function GiLandC(ByVal TxWer As String) As String
'ISO 3166 Alpha 2

Dim AktZa As Long
Dim Zeich As String
Dim TmpSt As String

For AktZa = 1 To Len(TxWer)
    Zeich = UCase$(Mid$(TxWer, AktZa, 1))
    If Zeich >= "A" And Zeich <= "Z" Then
        TmpSt = TmpSt & Zeich
    End If
Next AktZa
If Len(TmpSt) <> 2 Then
    GiLandC = vbNullString
Else
    GiLandC = TmpSt
End If

End Function

Private Function GiTarif(ByVal TxWer As String) As String
'tariff_type besteht aus genau drei Zeichen 0-9 oder A-Z

Dim AktZa As Long
Dim Zeich As String
Dim TmpSt As String

For AktZa = 1 To Len(TxWer)
    Zeich = UCase$(Mid$(TxWer, AktZa, 1))
    If (Zeich >= "0" And Zeich <= "9") Or (Zeich >= "A" And Zeich <= "Z") Then
        TmpSt = TmpSt & Zeich
    End If
Next AktZa
If Len(TmpSt) <> 3 Then
    GiTarif = "590"
Else
    GiTarif = TmpSt
End If

End Function

Private Function GiWaehr(ByVal TxWer As String) As Boolean
'generalInvoice 5.0 schreibt CHF fest vor (balance/@currency fixed="CHF")

Dim TmpSt As String

TmpSt = UCase$(Trim$(TxWer))
TmpSt = Replace(TmpSt, ".", vbNullString)
TmpSt = Replace(TmpSt, Chr$(32), vbNullString)

Select Case TmpSt
Case vbNullString, "CHF", "SFR", "FR", "CH", "SFRS", "FRANKEN", "SCHWEIZERFRANKEN"
    GiWaehr = True
Case Else
    GiWaehr = False
End Select

End Function

Private Function GiSteuIdx(StSat() As Currency, ByVal AnzSt As Integer, ByVal StuWe As Currency) As Integer
'Sucht den Index einer bereits vorhandenen Steuersatzgruppe

Dim AktZa As Integer

GiSteuIdx = -1
For AktZa = 0 To AnzSt - 1
    If StSat(AktZa) = StuWe Then
        GiSteuIdx = AktZa
        Exit Function
    End If
Next AktZa

End Function
