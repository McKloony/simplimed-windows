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

### Außerdem offen (eigene Tasks)

- **Task 11:** 32× `If GlTyp > 1` (invertiert, Schwerpunkt basDatRe) per Parser-Erweiterung;
  34× `Select Case GlTyp` manuell sichten (`scripts/gltyp_uncovered_v32final.csv`)
- **Task 10:** Abschluss-Verifikation + Spec-Status

## Empfehlung Reihenfolge

1. Gruppe 1 (84 mechanische, inkl. K32-Fix) — Patch-Gruppen mit Preview, 1 Commit je Muster
2. Gruppe 2 (28 SqlDat) — fokussierte Sitzung, Semantik-Freigaben einzeln
3. Schlüsselfeld-Verdachtsfälle aus Gruppe 3 (12575/16926) als Bug-Review an Anwender
4. Task 11, dann Task 10
