# Task-9-Arbeitsliste: verbleibende GlTyp-Verzweigungen

Stand: 2026-06-10 nach Triage · Quellen: `scripts/gltyp_report_v32final.csv`, `scripts/gltyp_triage.csv`

## Erledigt (Tasks 8 + 9 bisher)

**389 Blöcke vereinheitlicht:** 320 MECH (Wellen 1–5), 63 konkatenierte Semikolons (Lib v3.2),
6 SqlPad-Helper-Umbauten. Helper `SqlDat`/`SqlNum`/`SqlPad` per Direktfenster auf beiden
Backends verifiziert.

## Verbleibend: 221 Blöcke in drei Gruppen

### Gruppe 1 — MERGE_SAFE: 84 Stellen (57 DIFF + 27 MECH_NOSQL)

Manuell-mechanisch zusammenführbar, je Muster-Gruppe ein Patch-Skript mit Preview + Commit:

| Muster | Beispiele | Hinweis |
|---|---|---|
| `YEAR(`/`MONTH(` (T-SQL, groß) vs. `Year([`/`Month([` (Jet) | frmZeitraum, frmAbschl, frmBuExp, frmReExpo, frmReSam, frmAdrWord, basData, basMain | Vereinheitlicht auf `Year([Feld])` — T-SQL ist case-insensitiv; NUR die reinen Fälle, nicht `DATEPART(wk,…)` vs. `DatePart("ww",…)` (strukturell, bleibt) |
| frmAdrFilt-Like-Gruppe (Blöcke 2674–2753) | `(Feld) Like` vs. `([Feld]) Like` | rein kosmetisch |
| einfache SELECTs mit Spalten-/Whitespace-Kosmetik | basLabor u. a. | je Stelle Preview |
| MECH_NOSQL (27) | `Krit1`/`Krit2`-SQL-Fragmente | Normalisierer-Gleichheit bereits bewiesen, nur SQL-Verb fehlte |

**Pflicht-Bugfix beim Merge:** `basDaAdr.bas:11384` — `"SELECT * FROM qryKat01QORDER BY [ID1];"`
(fehlendes Leerzeichen; Access-Zweig von Katalogknoten K32 seit jeher defekt). Beim
Zusammenführen Leerzeichen ergänzen. **Verifiziert am Code.**

### Gruppe 2 — SqlDat-Kandidaten ✅ ERLEDIGT (Commit 6609d79)

21 Blöcke auf SqlDat umgebaut (Quartals-/Tages-/Zeitraumfilter, 8 Module); Tagesfilter auf
`[Tag, Tag+1)` und Access-HAVING→WHERE je mit User-Freigabe; 2 Bugfixes (Krankenblatt-ORDER-BY
`[CONVERT(CHAR(8),…)]`→`[Druckdatum]` 4×; Geburtstags-Suche CONVERT-Stil-102). 14 geprüfte
Blöcke bleiben strukturell verzweigt: Wochenfilter (`DATEPART(ww)` — Wochennummerierung
engine-abhängig), qryTerWiVor (ORDER BY `Sorter` vs. `ZeiVon`), DBCmRe2-Datums-Parameter,
Krankenblatt-ORDER-BY (Datums-Trunkierung je Engine).

### (ursprüngliche Planung Gruppe 2 — historisch)

| Sub-Muster | Anzahl | Behandlung |
|---|---|---|
| B: `CONVERT(DATETIME,…,102)` vs. `#…#` | 12 | SqlDat-Umbau; Sonderfall basData:11869 (WHERE vs. **HAVING+Between**) einzeln |
| C: einseitige `#`-Literale / deutsche Datums-Strings | 10 | SqlDat-Umbau; basMain:7265ff nutzt server-seitig `'01.01.JJJJ'` → **behebt DATEFORMAT-Risiko** |
| A: DatePart-Formatierungsblöcke (`DaSt1 = …`) | 6 | zusammen mit SQL-Partnerblock auf `SqlDat(…)` inline |
| E: Sonstige | 6 | Einzelsichtung |

Jede Semantik-Angleichung (Gleichheit→Bereich, Between-Grenzen) wird dem Anwender einzeln
vorgelegt (Spec Abschnitt 5).

**Einzelreview zusätzlich:** basMain:13373ff — Access-Zweig enthält
`[CONVERT(CHAR(8), Druckdatum, 24)]` als „Spaltenname" in Klammern (mutmaßlich Alt-Bug,
T-SQL-Ausdruck im Jet-Zweig).

### Gruppe 3 — bleiben begründet verzweigt: ~109 Stellen

| Kategorie | Anzahl | Begründung |
|---|---|---|
| KEEP_STRUCTURAL (Triage) | 81 | u. a.: Boolean `=1` (BIT) vs. `=-1` (Jet) — nie automatisierbar; **unterschiedliche Schlüsselfelder** (basDatKat:12575 `ID1` vs. `[ID0]`, 16926 `ID4` vs. `[ID0]` — mutmaßlich weitere Alt-Bugs, Einzelreview empfohlen!); DBCmEx-/Command-Architektur (adCmdStoredProc vs. Jet Stored Query); RDP-Caption-Logik; `SDBSta`-Signaturen |
| DateNum-D: Recordset-Zeitfelder | 24 | SQL Server speichert Datum+Zeit, Access nur Zeit (`ZeiVon`/`ZeiBis`) — Speichermodell |
| Verbindungslogik (clsConn/clsData) | 7 | Kategorie (a) der Spec |
| ONLY/ONLY1 (TSE/Stored-Procs) | 19 | SQL-Server-exklusiv, Kategorie (c) |

### Task 11 ✅ ERLEDIGT (Commits 9cafcaa, fa2d161, dc510cc)

Parser v4 erkennt `If GlTyp > 1` (Zweige getauscht, Klassen ONLYA/ONLY1A). Von 32 invertierten
Blöcken 17 vereinheitlicht (5 MECH/NOSQL, 12 manuell: SqlDat-Ranges frmAdrFilt, Day()/Month()
statt DATEPART/DatePart('m'), Präfix-Merges); 15 bleiben (Mailing-Boolean, DBCmRe2-Familie,
Datei-Dialog, 3 ONLYA). Von 34 `Select Case GlTyp` 11 vereinheitlicht (Quartals-/Zeitraum-
Nachzügler mit SqlDat, PLZ-/TOP-1-/ID0-Kosmetik); 23 bleiben (16× .dbx/.dbv-Dateinamen,
Boolean, frmOptions=WIP, clsData=Verbindung).

### Endstand (Task 10, Report `scripts/gltyp_report_final.csv`)

**496 Dialekt-Konstrukte vereinheitlicht** (inkl. Opt_PLZ, 2026-06-10 nachgezogen).
Verbleibend 149 Blöcke + 34 Einzelzeilen, alle
begründet: 49 strukturelle (Triage-Begründungen in `gltyp_triage.csv`), 27 Command/StoredProc,
26 Datums-Schema/Wochenfilter, 15 Jet-Boolean vs. BIT, 19 ONLY (TSE/Procs), 3 ONLYA,
7 Verbindungslogik, 3 Helper selbst, 16× .dbx/.dbv u. a.

### Offene User-Entscheide

- ~~PLZ/Land-Wildcard (basDaAdr)~~ ✅ 2026-06-10 entschieden: Ergebnis laut Anwender beidseitig
  identisch → Opt_PLZ vereinheitlicht (Land exakt, universelle Form; Triage-Bucket=MERGED)
- Schlüsselfelder basDatKat:12575/16926 (`ID1`/`ID4` vs. `[ID0]` — Bug oder Absicht?)

## Empfehlung Reihenfolge

1. Gruppe 1 (84 mechanische, inkl. K32-Fix) — Patch-Gruppen mit Preview, 1 Commit je Muster
2. Gruppe 2 (28 SqlDat) — fokussierte Sitzung, Semantik-Freigaben einzeln
3. Schlüsselfeld-Verdachtsfälle aus Gruppe 3 (12575/16926) als Bug-Review an Anwender
4. Task 11, dann Task 10
