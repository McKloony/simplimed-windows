# Design: Vereinheitlichung der SQL-Dialekte (GlTyp-Verzweigungen)

Stand: 2026-06-10 · Status: **Umgesetzt** — 495 Dialekt-Konstrukte vereinheitlicht; Restbestand
(150 Blöcke + 34 Zeilen) begründet verzweigt, kategorisiert in
`docs/superpowers/2026-06-10-task9-arbeitsliste.md` und `scripts/gltyp_triage.csv`

## 1. Kontext / Ist-Zustand

SimpliMed (VB6, Ordner `001/`) unterstützt zwei Datenbank-Backends über ADO.
`GlTyp` (gesetzt in `basMain.bas` aus INI `System/DBaTyp`) wählt den Provider in `clsConn.cls`:

| GlTyp | Provider | Engine |
|---|---|---|
| 0 | SQLNCLI.1 (Native Client) | SQL Server |
| 1 | SQLOLEDB.1 | SQL Server |
| 2 | Microsoft.Jet.OLEDB.4.0 (.mdb) | Access/Jet |
| 3 | Microsoft.ACE.OLEDB.12.0 (.accdb) | Access/ACE |

Im Code existieren **677 `GlTyp`-Treffer in 23 Dateien, davon 603 `If GlTyp < 2`-Verzweigungen**,
die überwiegend zwei nahezu identische SQL-Statements bauen (SQL-Server- vs. Access-Variante).

Befund der Stichproben (`basData.bas`, `basDatRe.bas`, `basDatKat.bas`):

1. **Rein kosmetische Paare (geschätzt 85–90 %):** Unterschiede nur in `dbo.`-Präfix,
   eckigen Klammern um Spaltennamen und Schluss-Semikolon.
   Beispiel: `SELECT * FROM dbo.qrySimAdDr3 ORDER BY Matchcode` vs.
   `SELECT * FROM qrySimAdDr3 ORDER BY [Matchcode];`
2. **Echte Ausdrucks-Unterschiede:** Zero-Padding (`RIGHT('00000000' + CONVERT(varchar(10), Mandant), 8)`
   vs. `Format$([Mandant],'00000000')`), Datums-Literale (`CONVERT(DATETIME, '...', 102)` vs. `#...#`).
   Die Datums-**Stringformatierung** ist dabei oft selbst `GlTyp`-verzweigt
   (Muster `basDatKat.bas:1401–1406`: `DatePart`-Konkatenation, SQL Server `yyyy-m-d hh:nn:ss`,
   Access `m/d/yyyy`).
3. **Strukturell verschieden:** Stored Procedures mit `@Parametern` (`DBCmEx6`, z. B. TSE-Signaturen
   `qrySimRzTSE1`) existieren nur im SQL-Server-Zweig.
4. **Versteckte Semantik-Differenzen:** z. B. `basDatKat.bas:1408–1412` — SQL Server filtert per
   Datums-**Bereich** (`>= Tag AND <= Folgetag`), Access per **Gleichheit** (`= #Tag#`).
   Solche Paare sind keine Kosmetik und dürfen nicht automatisch zusammengeführt werden.

## 2. Ziel & Scope

**Ziel:** Ein universeller SQL-Dialekt, den beide Engines über ADO/OLE DB unverändert akzeptieren.
`GlTyp`-Abfragen verbleiben ausschließlich an drei legitimen Stellen:

- (a) Provider-/Verbindungslogik (`clsConn.cls`, `clsData.cls`, `basMain.bas`-Initialisierung)
- (b) innerhalb der Dialekt-Helper in `basFormat.bas`
- (c) strukturell SQL-Server-exklusive Blöcke (Stored Procedures / TSE via `DBCmEx6`)

**Randbedingungen (vom Anwender festgelegt):**

- Beide Backends bleiben dauerhaft unterstützt.
- **Nur VB6-Code wird geändert** — keine Änderungen an Views/gespeicherten Abfragen/Schemata.
- Default-Schema der SQL-Server-Logins ist garantiert `dbo` → `dbo.`-Präfix ist verzichtbar.

**Nicht im Scope:** ADO-`.Filter`-Zeilen (clientseitige Recordset-Filter, eigene Syntax mit `*`),
ADO-Transaktionen (separates offenes Thema), Double→Currency-Umstellung (separates offenes Thema).
Die User-WIP-Dateien `basLayout.bas`, `frmMandant.frm`, `frmOptions.frm` werden erst nach
Commit des WIP angefasst.

## 3. Der universelle Dialekt (Konvention)

| Element | Universelle Form | Begründung |
|---|---|---|
| Objektname | `qrySimAdSu` (ohne `dbo.`) | Default-Schema dbo deckt SQL Server ab; Jet kennt kein Schema |
| Spaltennamen | `[Matchcode]` mit eckigen Klammern | Gültig in T-SQL und Jet; schützt Umlaut-Spalten (`[Zähler]`) und reservierte Wörter (`[Datum]`) |
| Schluss-Semikolon | weglassen | Optional in beiden Dialekten |
| String-Literale | `'...'` + `SqlStr`-Escaping | bereits projektweit umgesetzt |
| LIKE-Wildcards | `%` / `_` | Jet/ACE via ADO läuft im ANSI-92-Modus (Bestandscode nutzt das bereits) |
| Datums-Literale | nur über `SqlDat`-Helper | siehe Abschnitt 4 |
| Dezimalzahlen | nur über `SqlNum`-Helper | siehe Abschnitt 4 |
| Zero-Padding/Sortierausdrücke | nur über `SqlPad`-Helper | siehe Abschnitt 4 |

## 4. Helper in `basFormat.bas` (neben `SqlStr`, Zeile ~10397)

### `SqlDat(Datum As Date, Optional MitZeit As Boolean) As String`

Nimmt einen echten `Date`-Wert (keinen vorformatierten String) und liefert das komplette,
einbettungsfertige Literal:

- **SQL Server (`GlTyp < 2`):** `CONVERT(DATETIME, 'yyyy-mm-dd hh:nn:ss', 120)` —
  Stil 120 passt exakt zum gelieferten Format. (Bestand mischt Stil 102 mit
  Bindestrich-Strings; funktioniert nur dank tolerantem Parser und wird bei der
  Umstellung mit bereinigt.)
- **Access (`GlTyp >= 2`):** `#m/d/yyyy#` bzw. `#m/d/yyyy h:nn:ss#` (US-Reihenfolge).
- Stringaufbau **ausschließlich per `DatePart`-Konkatenation** (bewährtes Bestandsmuster).
  **Niemals `Format$` für SQL-Literale** — dessen Trennzeichen-Platzhalter sind locale-abhängig.

### `SqlNum(Wert As Variant) As String`

Für Dezimaltypen (Currency/Double/Single): `Trim$(Str$(Wert))`. `Str$` garantiert den Punkt
als Dezimaltrenner unabhängig von der Windows-Locale (deutsches `CStr`/`Format$` lieferte `1,5`
und zerstörte das Statement). Integer/Long bleiben bei direkter `&`-Konkatenation.
Grenzfall: `Str$` liefert bei extremen Double-Werten Exponentialschreibweise — für Currency
(Festkomma) ausgeschlossen; der Helper prüft auf `E`/`e` im Ergebnis und loggt via `SErLog`.

### `SqlPad(FeldName As String, Breite As Integer) As String`

Liefert je `GlTyp` den Padding-Ausdruck:
SQL Server `RIGHT('0…0' + CONVERT(varchar(10), Feld), n)`, Access `Format$([Feld],'0…0')`.
Ersetzt das vielfach kopierte `SoStr`-Muster.

Weitere Helper nur bei Bedarf, wenn die Analyse (Phase 1) zusätzliche echte
Differenz-Muster zutage fördert (z. B. `SqlBool`). YAGNI.

## 5. Vorgehen in zwei Phasen

### Phase 1 — Analyse (read-only, ändert keinen Code)

PowerShell-Skript `scripts/analyze_gltyp.ps1`:

1. Parst alle `If GlTyp < 2 … [Else …] End If`-Blöcke in `001/*.bas|*.frm|*.cls` (CP-1252-bytetreu).
2. Erfasst zusätzlich gepaarte, `GlTyp`-verzweigte **Formatierungs-Blöcke**
   (`DaSt* = DatePart(...)`-Muster) vor den SQL-Blöcken.
3. Extrahiert die SQL-String-Paare, normalisiert beide Seiten
   (entfernt `dbo.`, eckige Klammern, Schluss-Semikolon, normalisiert Whitespace).
4. Klassifiziert jeden Block:
   - **MECH** — nach Normalisierung byte-identisch → mechanisch zusammenführbar
   - **DIFF** — echte Differenz → Einzelfall-Liste (Datei, Zeile, beide Strings)
   - **ONLY** — kein Else-Zweig (SQL-Server-only, z. B. TSE) → bleibt verzweigt
   - **DATE/NUM** — Block enthält `#`-Literal, `CONVERT(DATETIME`, Datums-Formatierung
     oder Dezimalzahl-Konkatenation → **unabhängig von der Kosmetik-Klassifikation
     immer manuelles Review**, nie Auto-Patch
5. Ausgabe: CSV-Report + Summenstatistik.

**Gate:** Erst nach Review des Reports (insbesondere DIFF- und DATE/NUM-Liste)
durch den Anwender beginnt Phase 2.

### Phase 2 — Transformation (skriptgestützt, modulweise)

- MECH-Blöcke: Ersetzung des gesamten If/Else/End-If-Blocks durch das eine universelle
  Statement per **CP-1252-Byte-Patch** (Zwei-Phasen-Muster wie `scripts/apply_codereview_fixes.ps1`,
  inkl. Umlaut-Zählung vor/nach). VB6-Dateien werden niemals mit Edit/Write bearbeitet.
- DIFF/DATE/NUM-Fälle: Einzelfallentscheidung — Helper-Einsatz (`SqlDat`/`SqlNum`/`SqlPad`)
  oder begründet verzweigt belassen. Bei Datums-Gleichheit-vs.-Bereich ist das Zielbild
  der Bereichsfilter `>= SqlDat(Tag) AND < SqlDat(Folgetag)` (korrekt auf beiden Engines,
  auch bei Zeitanteilen in den Daten) — aber erst nach Review je Stelle.
- **Ein Commit pro Modul.** Pilot: `basDATEV.bas` (5 Stellen), dann aufsteigend nach
  Trefferzahl; `basDatKat.bas` (93) und `basData.bas` (133) zuletzt.
- Vor jedem Patch-Lauf: Backup nach `backups/gltyp_<datum>/`.

### Modul-Inventar (Treffer `If GlTyp < 2`)

| Datei | Treffer | | Datei | Treffer |
|---|---|---|---|---|
| basData.bas | 133 | | frmAdrFilt.frm | 14 |
| basDatKat.bas | 92 | | frmZeitraum.frm | 13 |
| basDaAdr.bas | 82 | | basLabor.bas | 9 |
| basMain.bas | 66 | | frmIfap.frm | 7 |
| basDatRe.bas | 65 | | frmBuExp.frm | 5 |
| basDaMa.bas | 44 | | basDATEV.bas | 5 |
| clsLisLab.cls | 43 | | clsConn.cls | 5 * |
| | | | übrige 8 Dateien | je ≤ 4 |

\* `clsConn.cls` gehört zur Verbindungslogik (Kategorie (a), Abschnitt 2) und wird nicht umgebaut.

## 6. Verifikation

- Nach jedem Modul: Grep-Nachweis, dass verbleibende `GlTyp`-Treffer nur den legitimen
  Kategorien (a)–(c) aus Abschnitt 2 angehören.
- `SqlDat`/`SqlNum` werden **vor** breitem Einsatz gegen beide Test-Backends verifiziert
  (eine Access- und eine SQL-Server-Testdatenbank), inkl. Randfälle:
  Datum mit/ohne Zeitanteil, Jahreswechsel, Beträge mit Nachkommastellen.
- Nach jedem Commit-Paket: VB6-IDE-Compile + Smoke-Test durch den Anwender auf **beiden**
  Backends. Fokus: Adresssuche, Terminliste, Rechnungsexport (Module mit den meisten Treffern).
- Alle Änderungen ausschließlich per Byte-Patch; `git diff` + manuelles Lesen vor jedem Commit.

## 7. Risiken & Gegenmaßnahmen

| Risiko | Gegenmaßnahme |
|---|---|
| Scheinbar kosmetisches Paar mit versteckter Differenz | Normalisierungs-Analyse: nur byte-identische Paare gelten als MECH; alles andere DIFF |
| Datums-/Zahlenformatierung locale-abhängig (deutsche Locale: Komma, dd.mm.yyyy) | `SqlDat`/`SqlNum` per `DatePart`/`Str$`; `Format$` für SQL-Literale verboten; DATE/NUM-Stellen nie Auto-Patch |
| Semantik-Differenz Gleichheit vs. Bereich bei Datumsfiltern | eigene DIFF-Kategorie, Einzelfall-Review, Zielbild Bereichsfilter |
| Stil-102/120-Inkonsistenzen im Bestand | Umstellung auf `SqlDat` erzeugt konsistente Literale |
| `Str$`-Exponentialschreibweise bei extremen Doubles | Helper prüft auf `E`/`e`, loggt via `SErLog` |
| CP-1252/Umlaute werden durch Tools zerstört | ausschließlich PowerShell-Byte-Patching, Umlaut-Zählung vor/nach |
| Abweichendes Default-Schema bei einzelnen Kunden | laut Anwender garantiert `dbo`; Restrisiko dokumentiert |

## 8. Erfolgskriterien

1. `Grep "If GlTyp < 2"` liefert nach Abschluss nur noch Treffer der Kategorien (a)–(c).
2. Identisches Anwendungsverhalten auf beiden Backends (Smoke-Tests grün).
3. Neue SQL-Statements brauchen keine Verzweigung mehr — eine Schreibweise, dokumentiert
   durch diese Spec als Konvention.
