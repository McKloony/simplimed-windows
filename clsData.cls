VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As VB.Form
Private BsB01 As XtremeCommandBars.BackstageButton
Private BsB02 As XtremeCommandBars.BackstageButton
Private BsB03 As XtremeCommandBars.BackstageButton
Private BsB04 As XtremeCommandBars.BackstageButton
Private BsB05 As XtremeCommandBars.BackstageButton
Private BsB06 As XtremeCommandBars.BackstageButton
Private BsB07 As XtremeCommandBars.BackstageButton
Private BsB08 As XtremeCommandBars.BackstageButton
Private BsB09 As XtremeCommandBars.BackstageButton
Private BsB10 As XtremeCommandBars.BackstageButton
Private BsB11 As XtremeCommandBars.BackstageButton
Private BsB12 As XtremeCommandBars.BackstageButton
Private BsB13 As XtremeCommandBars.BackstageButton
Private BsB14 As XtremeCommandBars.BackstageButton
Private BsB15 As XtremeCommandBars.BackstageButton
Private BsB16 As XtremeCommandBars.BackstageButton
Private BsB17 As XtremeCommandBars.BackstageButton
Private BsB18 As XtremeCommandBars.BackstageButton
Private BsB19 As XtremeCommandBars.BackstageButton
Private BsB20 As XtremeCommandBars.BackstageButton
Private CmAcs As XtremeCommandBars.CommandBarActions
Private ShCut As XtremeShortcutBar.ShortcutBar

Private WithEvents clCon As clsConn
Attribute clCon.VB_VarHelpID = -1

Private DaPfa As String 'Path Access Database
Private DaUsr As String 'Username SQL-Server
Private DaSer As String 'Servername SQL-Server
Private DaTab As String 'Databasename SQL-Server
Private DaPas As String 'Datenbankpasswort Access
Private DaUsN As String 'Username Sicheerheitsdatei
Private DaUsP As String 'Passwort Sicheerheitsdatei
Private SerPa As String 'Anwenderpasswort (SQL Server)
Private DaVer As Integer 'SQL-Server Security
Private DaTim As Integer 'Connection Timeout
Private DaAsy As Boolean 'Asynchron Dataconnection
Public Sub FConn(Optional DatSe As String = vbNullString, Optional ByVal DatTa As String = vbNullString, Optional ByVal DatUs As String = vbNullString, Optional ByVal DatPa As String = vbNullString)
On Error GoTo StErr
'Startet den Datenbankzugriff

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim TesNa As String

If GlTyp < 2 Then
    TeTit = "Fehler beim Öffnen der SQL-Datenbank"
    TeMai = "Beim Öffnen der SQL-Datenbank ist ein Fehler aufgetreten. Sollen die Datenbank-Einstellungen jetzt zurückgesetzt werden?"
    TeInh = "Beim Versucht auf den SQL-Server zuzugreifen ist ein Fehler aufgetreten. Entweder sind die Einstellungen im Abschnitt: Datenbankeinstellungen nicht richtig oder der SQL-Server ist nicht erreichbar"
    TeFus = "SimpliMed wird nun beendet. Bitte starten Sie SimpliMed erneut und korrigieren falls erforderlich die Einstellungen im Abschnitt: Datenbankeinstellungen."
Else
    TeTit = "Fehler beim Öffnen de Datenbankdatei"
    TeMai = "Beim Öffnen der Datenbankdatei ist ein Fehler aufgetreten. Sollen die Datenbankdatei-Einstellungen jetzt zurückgesetzt werden?"
    TeInh = "Die Daten konnten nicht korrekt gelesen werden. Entweder ist die Datenbankdatei beschädigt oder Sie versuchen eine Datei zu öffnen, die keine SimpliMed Datenbankdatei ist. Das Programm wird daher beendet"
    TeFus = "Bitte löschen Sie die Konfigurationsdatei im Ordner Praxisdaten und versuchen Sie das Programm danach erneut zu starten"
End If

DaVer = CInt(Right(IniGetVal("System", "DatVer"), 1))
DaAsy = CBool(IniGetVal("System", "DatLad"))
DaTab = IniGetVal("System", "DatTab")  'Databasename SQL-Server
DaTim = IniGetVal("System", "DatAkt")  'Connection Timeout
DaSer = IniGetVal("System", "DatSer")  'Servername SQL-Server
DaUsr = IniGetVal("System", "DatUsr")  'Username SQL-Server
SerPa = IniGetVal("System", "DaUsPa")  'Passwort SQL-Server
DaPas = IniGetVal("System", "DatPas")  'Datenbankpasswort Access
DaUsN = IniGetVal("System", "DaUsNa")  'Username Sicheerheitsdatei
DaUsP = IniGetVal("System", "DaSiPa")  'Passwort Sicheerheitsdatei
DaPfa = IniGetVal("SysPfa", "DatPfa")

If DatSe <> vbNullString Then
    DaSer = DatSe 'Servername SQL-Server
End If

If DatTa <> vbNullString Then
    DaTab = DatTa 'Databasename SQL-Server
End If

If DatUs <> vbNullString Then
    DaUsr = DatUs 'Username SQL-Server
End If

If DatPa <> vbNullString Then
    SerPa = DatPa 'Passwort SQL-Server
End If

If GlVrs = True Then 'Kennwort verschlüsselt
    DaPas = SCrypt(DaPas, False)
    DaUsN = SCrypt(DaUsN, False)
    DaUsP = SCrypt(DaUsP, False)
    SerPa = SCrypt(SerPa, False)
End If
If GlTyp = 3 Then
    DaPas = GlDPa
End If

With clCon
    .DaPrm = adPromptComplete
    .DaSer = DaSer 'Servername SQL-Server
    .DaUsr = DaUsr 'Username SQL-Server
    .DaSQL = SerPa 'Passwort SQL-Server
    .DaPas = DaPas 'Datenbankpasswort Access
    .DaUsN = DaUsN 'Username Sicheerheitsdatei
    .DaUsP = DaUsP 'Passwort Sicheerheitsdatei
    .DaTim = DaTim 'Connection Timeout
    .DaSec = DaVer 'SQL-Server Security
    .DaAsy = DaAsy 'Asynchron Dataconnection
    .DaTyp = GlTyp
    .DaPer = True
    .DaMod = 3
    Select Case GlTyp
    Case 0: 'SQL Native
        If DaTab <> "-" Then
            .DaNam = DaTab
        End If
        .FDBOp
    Case 1: 'SQL OLEDB
        If DaTab <> "-" Then
            .DaNam = DaTab
        End If
        .FDBOp
    Case 2: 'Jet 4.0
        TesNa = SDBTes(DaPfa)
        If TesNa <> vbNullString Then
            .DaNam = TesNa
            .FDBOp
        Else
            FOpn
        End If
    Case 3: 'Jet 12.0
        TesNa = SDBTes(DaPfa)
        If TesNa <> vbNullString Then
            .DaNam = TesNa
            .FDBOp
        Else
            FOpn
        End If
    End Select
End With

Exit Sub

StErr:
If GlDbg = True Then SErLog Err.Description & " FConn " & Err.Number & " - " & Err.Source
SSplS1 'Splashscreen schließen
SSplS2 'Splashscreen schließen
DoEvents
If GlApp = False Then 'Appmode
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, frmMain.hwnd
Else
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, frmApMod.hwnd
End If
If GlMes = 33565 Then
    If GlTyp < 2 Then
        If GlRDP = False Then
            IniSetVal "System", "DBaTyp", "S3"
        End If
    Else
        IniSetVal "SysPfa", "DatPfa", LCase(GlPfa & "Datenbankdatei.dbv")
    End If
End If
If GlApp = False Then 'Appmode
    Unload frmMain
Else
    Unload frmApMod
End If


End Sub
Private Sub FOpn()
On Error GoTo StErr

Dim TrLi1 As XtremeSuiteControls.TreeView
Dim TrLi3 As XtremeSuiteControls.TreeView
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim BaItm As XtremeShortcutBar.ShortcutBarItem

GlDat = False

If GlApp = True Then 'AppMode
    Exit Sub
End If

Set FM = frmMain
Set TrLi1 = FM.trvList1
Set TrLi3 = FM.trvList3
Set BsB01 = FM.bksButt01
Set BsB02 = FM.bksButt02
Set BsB03 = FM.bksButt03
Set BsB04 = FM.bksButt04
Set BsB05 = FM.bksButt05
Set BsB06 = FM.bksButt06
Set BsB07 = FM.bksButt07
Set BsB08 = FM.bksButt08
Set BsB09 = FM.bksButt09
Set BsB10 = FM.bksButt10
Set BsB11 = FM.bksButt11
Set BsB12 = FM.bksButt12
Set BsB13 = FM.bksButt13
Set BsB14 = FM.bksButt14
Set BsB15 = FM.bksButt15
Set BsB16 = FM.bksButt16
Set BsB17 = FM.bksButt17
Set BsB18 = FM.bksButt18
Set BsB19 = FM.bksButt19
Set BsB20 = FM.bksButt20
Set CmBrs = FM.comBar01
Set ShCut = FM.shtCut01
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)

If GlTkB = 1 Then 'Taskbar Einstellung
    ShCut.Item(0).Enabled = False
    ShCut.Item(1).Enabled = False
    ShCut.Item(2).Enabled = False
    ShCut.Item(3).Enabled = False
    ShCut.Item(4).Enabled = False
    ShCut.Item(5).Enabled = False
    ShCut.Item(6).Enabled = False
    ShCut.Item(7).Enabled = False
    ShCut.Item(8).Enabled = False
    DoEvents
Else
    CmAcs(ShoCut_Start).Enabled = False
    CmAcs(ShoCut_Adresse).Enabled = False
    CmAcs(ShoCut_Kranken).Enabled = False
    CmAcs(ShoCut_Abrechn).Enabled = False
    CmAcs(ShoCut_Finanz).Enabled = False
    CmAcs(ShoCut_Termin).Enabled = False
    CmAcs(ShoCut_Labor).Enabled = False
    CmAcs(ShoCut_Texte).Enabled = False
    CmAcs(ShoCut_Katalog).Enabled = False
End If

BsB01.Enabled = False
BsB02.Enabled = False
BsB03.Enabled = False
BsB04.Enabled = False
BsB05.Enabled = False
BsB06.Enabled = False
BsB07.Enabled = False
BsB08.Enabled = False
BsB09.Enabled = False
BsB10.Enabled = False
BsB11.Enabled = False
BsB12.Enabled = False
BsB13.Enabled = False
BsB14.Enabled = False
BsB15.Enabled = False
BsB16.Enabled = False
BsB17.Enabled = False
BsB18.Enabled = False
BsB19.Enabled = False
BsB20.Enabled = True
DoEvents

FCmAr
DoEvents
TrLi1.Nodes("P800").Expanded = True
TrLi3.Nodes("P800").Expanded = True
DoEvents
S_MeCm
DoEvents
SSplS1 'Splashscreen schließen
SSplS2 'Splashscreen schließen
DoEvents
SPrf1
DoEvents
SPrf2
DoEvents

GlAkt = False

If CBool(IniGetVal("Vorgabe", "StMaVo")) = False Then  'Standardmandant vorhanden
    MMain -2
Else
    SVeAn
End If

Exit Sub

StErr:
If GlDbg = True Then SErLog Err.Description & " FOpn " & Err.Number & " - " & Err.Source
Resume Next

End Sub
Private Sub FCmAr()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

Dim AktZa As Integer

ReDim GlTep(1, 4) 'Kalendermarker
GlTep(1, 0) = 1
GlTep(1, 1) = "Termintyp"
GlTep(1, 2) = vbBlack
GlTep(1, 3) = 0

ReDim GlZah(1, 5) 'Zahlungsziele
GlZah(1, 0) = 1
GlZah(1, 1) = "14 Tage"
GlZah(1, 2) = 14
GlZah(1, 3) = -1
GlZah(1, 4) = 14

ReDim GlGKa(30, 2) 'Gebührenkataloge
For AktZa = 1 To 30
    GlGKa(AktZa, 0) = AktZa
    GlGKa(AktZa, 1) = "Katalog " & AktZa
Next AktZa

ReDim GlTar(1, 3) 'Tarife
GlTar(1, 0) = 1
GlTar(1, 1) = 1
GlTar(1, 2) = "Standardtarif"

ReDim GlStu(1, 3) 'Steuersätze
GlStu(1, 0) = 1
GlStu(1, 1) = Format$(0, GlWa1)
GlStu(1, 2) = "ohne Steuer"

GlTyV = True 'Krankenblattypen vorhanden
ReDim GlKrA(27, 5)
GlKrA(1, 0) = 1
GlKrA(1, 1) = "DI"
GlKrA(1, 2) = "Diagnosen"
GlKrA(1, 3) = 8421376
GlKrA(2, 0) = 2
GlKrA(2, 1) = "GE"
GlKrA(2, 2) = "Gebührenziffer"
GlKrA(2, 3) = -2147483640
GlKrA(3, 0) = 3
GlKrA(3, 1) = "LA"
GlKrA(3, 2) = "Laborleistung"
GlKrA(3, 3) = 25800
GlKrA(4, 0) = 4
GlKrA(4, 1) = "ME"
GlKrA(4, 2) = "Medikament"
GlKrA(4, 3) = 16711935
GlKrA(5, 0) = 5
GlKrA(5, 1) = "BE"
GlKrA(5, 2) = "Begründung"
GlKrA(5, 3) = 16711680
GlKrA(6, 0) = 6
GlKrA(6, 1) = "ZA"
GlKrA(6, 2) = "Zahlung"
GlKrA(6, 3) = 52224
GlKrA(7, 0) = 7
GlKrA(7, 1) = "PR"
GlKrA(7, 2) = "Provisionseintrag"
GlKrA(7, 3) = 32768
GlKrA(8, 0) = 8
GlKrA(8, 1) = "IG"
GlKrA(8, 2) = "IGeL Leistungen"
GlKrA(8, 3) = 8421504
GlKrA(9, 0) = 9
GlKrA(9, 1) = "GL"
GlKrA(9, 2) = "Gewerbeleistung"
GlKrA(9, 3) = -2147483640
GlKrA(10, 0) = 21
GlKrA(10, 1) = "BuAn"
GlKrA(10, 2) = "Anamnese"
GlKrA(10, 3) = 255
GlKrA(11, 0) = 22
GlKrA(11, 1) = "NO"
GlKrA(11, 2) = "Notizen"
GlKrA(11, 3) = 10053171
GlKrA(12, 0) = 23
GlKrA(12, 1) = "BF"
GlKrA(12, 2) = "Befunde"
GlKrA(12, 3) = 10040319
GlKrA(13, 0) = 24
GlKrA(13, 1) = "WO"
GlKrA(13, 2) = "Textdokumente"
GlKrA(13, 3) = 6710784
GlKrA(14, 0) = 25
GlKrA(14, 1) = "LB"
GlKrA(14, 2) = "Laborbefunde"
GlKrA(14, 3) = 2162853
GlKrA(15, 0) = 26
GlKrA(15, 1) = "TH"
GlKrA(15, 2) = "Therapiekonzept"
GlKrA(15, 3) = 6736896
GlKrA(16, 0) = 27
GlKrA(16, 1) = "PR"
GlKrA(16, 2) = "Prozedere"
GlKrA(16, 3) = 10053375
GlKrA(17, 0) = 28
GlKrA(17, 1) = "BL"
GlKrA(17, 2) = "Blutdruck"
GlKrA(17, 3) = 39372
GlKrA(18, 0) = 29
GlKrA(18, 1) = "GW"
GlKrA(18, 2) = "Gewicht"
GlKrA(18, 3) = 16750848
GlKrA(19, 0) = 30
GlKrA(19, 1) = "GR"
GlKrA(19, 2) = "Größe"
GlKrA(19, 3) = 16737945
GlKrA(20, 0) = 101
GlKrA(20, 1) = "RP"
GlKrA(20, 2) = "Beleg"
GlKrA(20, 3) = 0
GlKrA(21, 0) = 102
GlKrA(21, 1) = "DA"
GlKrA(21, 2) = "Datei"
GlKrA(21, 3) = 0
GlKrA(22, 0) = 103
GlKrA(22, 1) = "KD"
GlKrA(22, 2) = "Diagnose"
GlKrA(22, 3) = 16512
GlKrA(23, 0) = 104
GlKrA(23, 1) = "PR"
GlKrA(23, 2) = "Protokoll"
GlKrA(23, 3) = 8421504
GlKrA(24, 0) = 105
GlKrA(24, 1) = "BI"
GlKrA(24, 2) = "Bilddatei"
GlKrA(24, 3) = 8388672
GlKrA(25, 0) = 106
GlKrA(25, 1) = "RE"
GlKrA(25, 2) = "Rechnung"
GlKrA(25, 3) = 0
GlKrA(26, 0) = 107
GlKrA(26, 1) = "KM"
GlKrA(26, 2) = "Arzneimittel"
GlKrA(26, 3) = 16711935
GlKrA(27, 0) = 108
GlKrA(27, 1) = "EM"
GlKrA(27, 2) = "Emails"
GlKrA(27, 3) = 8421440

ReDim GlErK(2, 2) 'Einnahmekonten
GlErK(1, 0) = "420000"
GlErK(1, 1) = "420000 Erlöse"
GlErK(2, 0) = "420000"
GlErK(2, 1) = "420000 Eröse"

ReDim GlGeK(2, 4) 'Geldkonten
GlGeK(1, 0) = 1
GlGeK(1, 1) = "Bankkonto"
GlGeK(1, 2) = "170000"
GlGeK(1, 3) = "01 Bankkonto"
GlGeK(1, 4) = 0
GlGeK(2, 0) = 2
GlGeK(2, 1) = "Kasse"
GlGeK(2, 2) = "16000"
GlGeK(2, 3) = "01 Kasse"
GlGeK(2, 4) = -1

ReDim GlBTe(1, 2) 'Buchungstexte
GlBTe(1, 0) = 1
GlBTe(1, 1) = "Bankkonto"

ReDim GlWar(2, 2) 'Wärungen
GlWar(1, 0) = 1
GlWar(1, 1) = "€"
GlWar(1, 2) = "Euro"
GlWar(1, 3) = 1
GlWar(2, 0) = 2
GlWar(2, 1) = "€"
GlWar(2, 2) = "Euro"
GlWar(2, 3) = 1

ReDim GlBtr(1, 4) 'Betreffs
GlBtr(1, 0) = 1
GlBtr(1, 1) = "Testbetreff"
GlBtr(1, 2) = 0
GlBtr(1, 3) = 0
GlBtr(1, 4) = vbWhite

ReDim GlRmu(1, 5) 'Raumzuordnung
GlRmu(1, 0) = 1
GlRmu(1, 1) = "Raum #1"
GlRmu(1, 2) = 0
GlRmu(1, 3) = 0
GlRmu(1, 4) = 0

ReDim GlDia(1, 2) 'Diagnosegruppen
GlDia(1, 0) = 1
GlDia(1, 1) = " Meine Diagnoseauswahl"

ReDim GlMed(1, 2) 'Arzneigruppen
GlMed(1, 0) = 1
GlMed(1, 1) = " Meine Arzneiauswahl"

ReDim GlArt(1, 2) 'Artikelgruppen
GlArt(1, 0) = 1
GlArt(1, 1) = " Meine Artikelauswahl"

ReDim GlAnG(1, 2) 'Fragebogengruppen
GlAnG(1, 0) = 6
GlAnG(1, 1) = " Fragebogengruppe"

ReDim GlLab(1, 2) 'Laborkataloge
GlLab(1, 0) = 1
GlLab(1, 1) = " Laborkatalog"

ReDim GlThe(1, 39)
GlThe(1, 0) = 1
If IniGetVal("Adress", "AVoNa") <> vbNullString Then
    GlThe(1, 1) = IniGetVal("Adress", "AVoNa")
Else
    GlThe(1, 1) = "Vorname"
End If
If IniGetVal("Adress", "AName") <> vbNullString Then
    GlThe(1, 2) = IniGetVal("Adress", "AName")
Else
    GlThe(1, 2) = "Name"
End If
If IniGetVal("Adress", "AStra") <> vbNullString Then
    GlThe(1, 3) = IniGetVal("Adress", "AStra")
Else
    GlThe(1, 3) = "Straße"
End If
If IniGetVal("Adress", "APLZ") <> vbNullString Then
    GlThe(1, 4) = IniGetVal("Adress", "APLZ")
Else
    GlThe(1, 4) = "PLZ"
End If
If IniGetVal("Adress", "AOrt") <> vbNullString Then
    GlThe(1, 5) = IniGetVal("Adress", "AOrt")
Else
    GlThe(1, 5) = "Ort"
End If
If IniGetVal("Adress", "ATele") <> vbNullString Then
    GlThe(1, 6) = IniGetVal("Adress", "ATele")
Else
    GlThe(1, 6) = "Telefon"
End If
If IniGetVal("Adress", "AFax") <> vbNullString Then
    GlThe(1, 7) = IniGetVal("Adress", "AFax")
Else
    GlThe(1, 7) = "Telefax"
End If
If IniGetVal("Adress", "ABank") <> vbNullString Then
    GlThe(1, 8) = IniGetVal("Adress", "ABank")
Else
    GlThe(1, 8) = "Bankname"
End If
If IniGetVal("Adress", "AnBLZ") <> vbNullString Then
    GlThe(1, 9) = IniGetVal("Adress", "AnBLZ")
Else
    GlThe(1, 9) = "BLZ"
End If
If IniGetVal("Adress", "AnKto") <> vbNullString Then
    GlThe(1, 10) = IniGetVal("Adress", "AnKto")
Else
    GlThe(1, 10) = "Konto"
End If
If IniGetVal("System", "PVSNum") <> vbNullString Then
    GlThe(1, 11) = IniGetVal("System", "PVSNum")
Else
    GlThe(1, 11) = "999999"
End If
If IniGetVal("Adress", "ABeru") <> vbNullString Then
    GlThe(1, 12) = IniGetVal("Adress", "ABeru")
Else
    GlThe(1, 12) = "Beruf"
End If
If IniGetVal("Adress", "AName") <> vbNullString Then
    If IniGetVal("Adress", "AVoNa") <> vbNullString Then
        GlThe(1, 13) = IniGetVal("Adress", "AName") & ", " & IniGetVal("Adress", "AVoNa")
    Else
        GlThe(1, 13) = IniGetVal("Adress", "AName")
    End If
Else
    GlThe(1, 13) = "Mandantenname"
End If
If IniGetVal("Adress", "ATite") <> vbNullString Then
    GlThe(1, 14) = IniGetVal("Adress", "ATite")
Else
    GlThe(1, 14) = "Titel"
End If
If IniGetVal("Adress", "AIKNr") <> vbNullString Then
    GlThe(1, 15) = IniGetVal("Adress", "AIKNr")
Else
    GlThe(1, 15) = "KVNummer"
End If
If IniGetVal("Adress", "AEmail") <> vbNullString Then
    GlThe(1, 16) = IniGetVal("Adress", "AEmail")
Else
    GlThe(1, 16) = "Keine@Emailadresse.com"
End If
If IniGetVal("Adress", "AInter") <> vbNullString Then
    GlThe(1, 17) = IniGetVal("Adress", "AInter")
Else
    GlThe(1, 17) = "Internet"
End If
If IniGetVal("Adress", "AIBANr") <> vbNullString Then
    GlThe(1, 18) = IniGetVal("Adress", "AIBANr")
Else
    GlThe(1, 18) = "IBAN"
End If
If IniGetVal("Adress", "APraxis") <> vbNullString Then
    GlThe(1, 19) = IniGetVal("Adress", "APraxis")
Else
    GlThe(1, 19) = "Praxis"
End If
If IniGetVal("Adress", "ABan2") <> vbNullString Then
    GlThe(1, 20) = IniGetVal("Adress", "ABan2")
Else
    GlThe(1, 20) = "Bank2"
End If
If IniGetVal("Adress", "AnBL2") <> vbNullString Then
    GlThe(1, 21) = IniGetVal("Adress", "AnBL2")
Else
    GlThe(1, 21) = "BLZ2"
End If
If IniGetVal("Adress", "AnKt2") <> vbNullString Then
    GlThe(1, 22) = IniGetVal("Adress", "AnKt2")
Else
    GlThe(1, 22) = "Konto2"
End If
If IniGetVal("Adress", "AIBAN2") <> vbNullString Then
    GlThe(1, 23) = IniGetVal("Adress", "AIBAN2")
Else
    GlThe(1, 23) = "IBAN2"
End If
If IniGetVal("Adress", "ASpr1") <> vbNullString Then
    GlThe(1, 24) = IniGetVal("Adress", "ASpr1")
Else
    GlThe(1, 24) = GlSZe 'Sprechzietenstring
End If
If IniGetVal("Adress", "ARas1") <> vbNullString Then
    GlThe(1, 26) = Right$(IniGetVal("Adress", "ARas1"), 1)
Else
    GlThe(1, 26) = GlZeR 'Zeitrasterindex
End If
If IniGetVal("Adress", "ARas2") <> vbNullString Then
    GlThe(1, 27) = Right$(IniGetVal("Adress", "ARas2"), 1)
Else
    GlThe(1, 27) = GlZeR 'Zeitrasterindex
End If
If IniGetVal("Adress", "AOnTe1") <> vbNullString Then
    GlThe(1, 28) = Right$(IniGetVal("Adress", "AOnTe1"), 1)
Else
    GlThe(1, 28) = "999999"
End If
If IniGetVal("Adress", "AOnTe2") <> vbNullString Then
    GlThe(1, 29) = Right$(IniGetVal("Adress", "AOnTe2"), 1)
Else
    GlThe(1, 29) = "123456"
End If
If IniGetVal("Adress", "ATmPl") <> vbNullString Then
    GlThe(1, 30) = Right$(IniGetVal("Adress", "ATmPl"), 1)
Else
    GlThe(1, 30) = False
End If
GlThe(1, 25) = False
If IniGetVal("Adress", "ABICNr") <> vbNullString Then
    GlThe(1, 31) = Right$(IniGetVal("Adress", "ABICNr"), 1)
Else
    GlThe(1, 31) = "BIC"
End If
If IniGetVal("Adress", "ABICN2") <> vbNullString Then
    GlThe(1, 32) = Right$(IniGetVal("Adress", "ABICN2"), 1)
Else
    GlThe(1, 32) = "BIC2"
End If
If IniGetVal("Adress", "AGlIDr") <> vbNullString Then
    GlThe(1, 33) = Right$(IniGetVal("Adress", "AGlIDr"), 1)
Else
    GlThe(1, 33) = "GID"
End If
If IniGetVal("Adress", "AOnFr1") <> vbNullString Then
    GlThe(1, 34) = Right$(IniGetVal("Adress", "AOnFr1"), 1)
Else
    GlThe(1, 34) = "999999"
End If
If IniGetVal("Adress", "AOnFr2") <> vbNullString Then
    GlThe(1, 35) = Right$(IniGetVal("Adress", "AOnFr2"), 1)
Else
    GlThe(1, 35) = "123456"
End If

ReDim GlTmF(20, 3) 'Termingruppen
GlTmF(1, 1) = 16777215 'weiss
GlTmF(2, 1) = 12632319 'hellrot
GlTmF(3, 1) = 12640511 'hellorange
GlTmF(4, 1) = 12648447 'hellgelb
GlTmF(5, 1) = 12648384 'hellgrün
GlTmF(6, 1) = 16777152 'helltürkis
GlTmF(7, 1) = 16769405 'hellblau
GlTmF(8, 1) = 16761087 'hellrosa
GlTmF(9, 1) = 9211135  'rot
GlTmF(10, 1) = 9554175 'orange
GlTmF(11, 1) = 5636095 'gelb
GlTmF(12, 1) = 6619080 'mosgrün
GlTmF(13, 1) = 16776960 'türkis
GlTmF(14, 1) = 14141600 'graublau
GlTmF(15, 1) = 13147135 'mangenta
GlTmF(16, 1) = 16116710 'hellgraublau
GlTmF(17, 1) = 8249855 'gelborange
GlTmF(18, 1) = 3316735 'orange
GlTmF(19, 1) = 65435   'grün
GlTmF(20, 1) = 13816530 'grau
For AktZa = 1 To 20
    GlTmF(AktZa, 0) = "Termingruppe " & AktZa
    GlTmF(AktZa, 2) = AktZa
Next AktZa

ReDim GlReK(1, 2) 'Rechnungskommentare
GlReK(1, 0) = 1
GlReK(1, 1) = vbNullString

ReDim GlPaG(1, 2) 'Patientengruppen
GlPaG(1, 0) = 0
GlPaG(1, 1) = "Adressgruppen"
GlPaG(1, 2) = "001"

ReDim GlEmG(1, 3) 'Emailgruppen
GlPaG(1, 0) = 0
GlPaG(1, 1) = "Emailgruppen"
GlPaG(1, 2) = "001"

ReDim GlLan(4, 3) 'Länder
GlLan(1, 0) = 1
GlLan(1, 1) = vbNullString
GlLan(1, 2) = vbNullString
GlLan(2, 0) = 2
GlLan(2, 1) = "Deutschland"
GlLan(2, 2) = "Deutsch"
GlLan(3, 0) = 3
GlLan(3, 1) = "Österreich"
GlLan(3, 2) = "Deutsch"
GlLan(4, 0) = 4
GlLan(4, 1) = "Schweiz"
GlLan(4, 2) = "Deutsch"

GlStP = Right$(IniGetVal("Vorgabe", "StaEmp"), 1)  'Standardempfängertyp

Select Case GlStD 'Standard-Dezimaltrennzeichen
Case 0: GlFak = CSng(IniGetVal("System", "AbrFak"))
Case 1: GlFak = Replace(IniGetVal("System", "AbrFak"), ",", ".", 1)
End Select

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " FCmAr " & Err.Number & " - " & Err.Source
Resume Next

End Sub
Private Sub Class_Initialize()
    Set clCon = New clsConn
End Sub
Private Sub Class_Terminate()
    Set clCon = Nothing
End Sub
Private Sub clCon_ConCom(adStat As ADODB.EventStatusEnum)
On Error GoTo StErr

Dim VerNr As String
Dim ReNum As Boolean
Dim RetWe As Boolean

If GlDat = True Then
    Exit Sub
End If

Set DB1 = clCon.DaCon

If DB1.State = adStateClosed Then
    DB1.Open
End If

GlSt1 = DB1.State

If GlRDP = False Then
    VerNr = S_Vers()
    ReNum = SVePr(VerNr)
End If

DoEvents
S_Ary0
DoEvents
S_Ary1
DoEvents
S_Ary3
DoEvents
S_Ary2
DoEvents
SStMa

If GlApp = False Then 'AppMode
    GrAd_Lad
    DoEvents
    GrMa_Lad
    DoEvents
    KTree
    DoEvents
    S_MeCm
    DoEvents
    S_WvSpl
    DoEvents
    S_StSt1
    DoEvents
    S_StSt2
    DoEvents
    S_StSt3
    DoEvents
    S_AdSta
    DoEvents
    S_AdSpl
    DoEvents
    S_ReLi
    DoEvents
    S_ReSpl
    DoEvents
    S_AbSpl
    DoEvents
    S_List
    DoEvents
    K_List
    DoEvents
    S_MaSta 'E-Mails
    DoEvents
    S_KrSt
    DoEvents
    S_KrLa
    DoEvents
    S_KrCo True
    DoEvents
    KSuGr
    DoEvents
    K_BuSpl
    DoEvents
    S_KrSpk
    DoEvents

    Select Case GlBut
    Case RibTab_Fragebogen:
                S_AnSpl
                S_AnBoP 'Befüllen des PropertyGrid
    Case RibTab_Krankenbla:
                RetWe = S_KrPk()
                SKrVo
    Case RibTab_Abrechnung:
                S_KrSpl
                RetWe = S_KrPo()
                SAnza
    Case RibTab_Tagesproto:
                S_KrSpl
                S_TaPo
    Case RibTab_Vorbereit:
                S_VoPo
    Case RibTab_Rezeptmodul:
                S_RzSpl
                S_RzLa
    Case RibTab_Belegmodul:
                S_RzSpl
                S_RzLa
    Case RibTab_Bildmodul:
                SBiFi
    Case RibTab_Rechnungen:
                STrNo
    Case RibTab_Mahnwesen:
                STrNo
                S_OPSpl
    Case RibTab_Buchungen:
                STrNo
                S_BuSpl
    Case RibTab_HomeBanki:
                STrNo
                S_BaSpl
    Case RibTab_Statistik:
                SStaSt
    Case RibTab_Ter_Listen:
                STrNo
                S_TeSpl
    Case RibTab_Ter_Akont:
                STrNo
                S_TeSpl
    Case RibTab_Ter_Warte:
                STrNo
                S_WaPo
                S_TeSpl
                S_WaSpl
    Case RibTab_Ter_Raeume:
                STeRa True
    Case RibTab_Ter_Mitarb:
                STeRa True
    Case RibTab_LabBericht:
                S_LbSp1
                S_LaPs
                S_LaBe
    Case RibTab_LabAuftrag:
                S_LaPs
                S_LaBe
    Case RibTab_LabBerichte:
                STrNo
                S_LbSp2
    Case RibTab_LabAuftrage:
                STrNo
    Case RibTab_Tex_Dokumt:
    
    Case RibTab_Tex_Vorlag:
    
    Case RibTab_Tex_Rezept:
    
    Case RibTab_Tex_NewsLe:
    
    End Select
    DoEvents
    
    S_TeLi
    DoEvents
    S_TePi 'Kalndermarker setzen
    DoEvents
End If

If GlApp = False Then 'Appmode
    SDBOpt
End If

If GlDat = False Then
    GlDat = True
End If

SDBSta False
DoEvents

If GlApp = False Then 'ApMode
    SSplS1 'Splashscreen schließen
    DoEvents
    
    SPrf1 Not GlMiV 'Mitarbeiter vorhanden
    DoEvents
    SPrf2
    DoEvents
    
    If CBool(IniGetVal("Vorgabe", "StMaVo")) = False Then  'Standardmandant vorhanden
        MMain -2
    Else
        If GlMaV = False Then 'Mandanten vorhanden
            MMain -2
            GlAdS = True 'Speichern der Adresse erforderlich
        Else
            SVeAn
        End If
    End If
End If

GlAkt = False

Exit Sub

StErr:
If GlDbg = True Then SErLog Err.Description & " clCon_ConCom " & Err.Number & " - " & Err.Source
Resume Next

End Sub
Private Sub clCon_ReSe00(adSta1 As ADODB.EventStatusEnum)
On Error GoTo AlErr

Exit Sub

AlErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ReSe00 " & Err.Number
Resume Next

End Sub

Private Sub clCon_ReSe01(adSta1 As ADODB.EventStatusEnum)
On Error GoTo AlErr

Exit Sub

AlErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ReSe01 " & Err.Number
Resume Next

End Sub


Private Sub clCon_ReSe02(adSta2 As ADODB.EventStatusEnum)
On Error GoTo AlErr

Exit Sub

AlErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ReSe02 " & Err.Number
Resume Next

End Sub

Private Sub clCon_ReSe03(adSta3 As ADODB.EventStatusEnum)
On Error GoTo AlErr

Exit Sub

AlErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ReSe03 " & Err.Number
Resume Next

End Sub
