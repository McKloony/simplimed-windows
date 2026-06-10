# SQL-Dialekt-Vereinheitlichung (GlTyp) — Implementierungsplan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Die 603 `If GlTyp < 2`-SQL-Verzweigungen im SimpliMed-VB6-Projekt (`001/`) auf einen universellen SQL-Dialekt zusammenführen; Rest-Verzweigungen nur noch in Verbindungslogik, Dialekt-Helpern und SQL-Server-exklusiven Blöcken.

**Architecture:** Zwei Phasen gemäß Spec `001/docs/superpowers/specs/2026-06-10-sql-dialekt-vereinheitlichung-design.md`. Phase 1: read-only PowerShell-Analyse klassifiziert alle Blöcke (MECH/DIFF/ONLY/ONLY1/COMPLEX + DateNum-Flag) → Report-Gate. Phase 2: drei VB6-Helper (`SqlDat`/`SqlNum`/`SqlPad`) in `basFormat.bas`, dann modulweise skriptgestützte Zusammenführung der MECH-Blöcke per CP-1252-Byte-Patch, ein Commit pro Modul; DIFF/DATE-NUM-Fälle ausschließlich manuell.

**Tech Stack:** Windows PowerShell 5.1 (CP-1252-Byte-Verarbeitung via `[System.IO.File]` + `Encoding.GetEncoding(1252)`), VB6, Git (Repo-Root ist `001/`).

**Eiserne Regeln:**
1. VB6-Dateien (`*.bas`, `*.cls`, `*.frm`) **niemals** mit Edit/Write bearbeiten — nur CP-1252-Byte-Patching per PowerShell.
2. Blöcke mit DateNum-Flag werden **nie** automatisch gepatcht.
3. User-WIP-Dateien `basLayout.bas`, `frmMandant.frm`, `frmOptions.frm` nicht anfassen, solange unkommittiert.
4. `clsConn.cls` wird nicht umgebaut (Verbindungslogik, Kategorie (a) der Spec).
5. Falls das PowerShell-Tool EPERM wirft: Workaround via Bash-Tool → `powershell.exe -NoProfile -ExecutionPolicy Bypass -File <skript>`.
6. Skripte liegen in `C:\Users\schmi\Documents\devcode2\scripts\` und sind **außerhalb** des Git-Repos — Commits betreffen nur Dateien unter `001/`.

---

## Task 1: Parser-Bibliothek `gltyp_lib.ps1`

**Files:**
- Create: `C:\Users\schmi\Documents\devcode2\scripts\gltyp_lib.ps1`

- [ ] **Step 1: Bibliothek schreiben** (Write-Tool, neue Datei, reines ASCII — Umlaute nur als `\u`-Escapes)

```powershell
# gltyp_lib.ps1 - Parser/Normalisierung fuer "If GlTyp < 2"-Bloecke
# Spec: 001/docs/superpowers/specs/2026-06-10-sql-dialekt-vereinheitlichung-design.md
# Dot-sourcen: . (Join-Path $PSScriptRoot 'gltyp_lib.ps1')

$script:Enc1252 = [System.Text.Encoding]::GetEncoding(1252)

function Read-Vb6Text([string]$Path) {
    [System.IO.File]::ReadAllText($Path, $script:Enc1252)
}

function Write-Vb6Text([string]$Path, [string]$Text) {
    [System.IO.File]::WriteAllText($Path, $Text, $script:Enc1252)
}

function Get-UmlautCount([string]$Text) {
    # \u-Escapes statt literaler Umlaute: .ps1 wird ohne BOM gespeichert,
    # PowerShell 5.1 laese literale UTF-8-Umlaute sonst als ANSI-Mojibake
    ([regex]::Matches($Text, '[' + [string][char]0xE4 + [char]0xF6 + [char]0xFC + [char]0xC4 + [char]0xD6 + [char]0xDC + [char]0xDF + ']')).Count
}

function Get-NormalizedSql([string[]]$BranchLines) {
    $t = ($BranchLines -join ' ')
    $t = $t -replace 'dbo\.', ''
    $t = $t -replace '[\[\]]', ''
    $t = $t -replace ';"', '"'
    $t = $t -replace '\s+', ' '
    $t.Trim()
}

function Test-DateNum([string]$BlockText) {
    $BlockText -match '"#|#"|CONVERT\s*\(\s*DATETIME|DatePart\s*\(|Format\$\s*\('
}

function Get-GlTypBlocks([string]$Path) {
    $lines = (Read-Vb6Text $Path) -split "`r`n"
    $blocks = New-Object System.Collections.Generic.List[object]
    $name = Split-Path $Path -Leaf
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $l = $lines[$i]
        if ($l -match '^(\s*)If GlTyp < 2 Then\s*(''.*)?$') {
            # Block-If (nach Then nichts oder nur Kommentar)
            $ifIndent = $Matches[1]
            $srv = @(); $acc = @(); $inElse = $false; $depth = 0
            $cls = $null; $end = -1
            for ($j = $i + 1; $j -lt $lines.Count; $j++) {
                $c = $lines[$j]
                if ($depth -eq 0 -and $c -match '^\s*End If\s*(''.*)?$') { $end = $j; break }
                if ($depth -eq 0 -and $c -match '^\s*ElseIf\b') { $cls = 'COMPLEX' }
                if ($depth -eq 0 -and $c -match '^\s*Else\s*(''.*)?$') { $inElse = $true; continue }
                if ($c -match '^\s*#?End If\b') { $depth-- }
                elseif ($c -notmatch '^\s*ElseIf\b' -and $c -match '^\s*#?If\b.*\bThen\s*(''.*)?$') { $depth++ }
                if ($c -match '\s_\s*$') { $cls = 'COMPLEX' }  # VB6-Zeilenfortsetzung: nicht automatisierbar
                if ($inElse) { $acc += $c } else { $srv += $c }
            }
            if ($end -lt 0) { $cls = 'COMPLEX'; $end = $i }
            if (-not $cls) {
                if (-not $inElse) { $cls = 'ONLY' }
                elseif ((Get-NormalizedSql $srv) -ceq (Get-NormalizedSql $acc)) { $cls = 'MECH' }
                else { $cls = 'DIFF' }
            }
            $blockText = ($lines[$i..([Math]::Max($end, $i))] -join "`r`n")
            $blocks.Add([pscustomobject]@{
                File = $name; StartLine = $i + 1; EndLine = $end + 1
                Class = $cls; DateNum = (Test-DateNum $blockText)
                IfIndent = $ifIndent
                SrvLines = $srv; AccLines = $acc
                SrvText = (($srv | ForEach-Object { $_.Trim() }) -join ' / ')
                AccText = (($acc | ForEach-Object { $_.Trim() }) -join ' / ')
            })
            if ($end -gt $i) { $i = $end }
        }
        elseif ($l -match '^(\s*)If GlTyp < 2 Then\s+\S') {
            # Einzeiler: If GlTyp < 2 Then <code>
            $blocks.Add([pscustomobject]@{
                File = $name; StartLine = $i + 1; EndLine = $i + 1
                Class = 'ONLY1'; DateNum = (Test-DateNum $l)
                IfIndent = $Matches[1]
                SrvLines = @($l); AccLines = @()
                SrvText = $l.Trim(); AccText = ''
            })
        }
    }
    $blocks
}
```

- [ ] **Step 2: Syntax-Check**

Run: `powershell.exe -NoProfile -Command ". 'C:\Users\schmi\Documents\devcode2\scripts\gltyp_lib.ps1'; 'LIB OK'"`
Expected: `LIB OK`

---

## Task 2: Test-Fixture (CP-1252) mit bekannten Blocktypen

**Files:**
- Create: `C:\Users\schmi\Documents\devcode2\scripts\tests\fixture_gltyp.bas` (CP-1252!)

- [ ] **Step 1: Fixture als UTF-8-Zwischendatei schreiben** (Write-Tool → `scripts\tests\fixture_gltyp.utf8.tmp`)

```vb
Attribute VB_Name = "fixture_gltyp"
'Testfixture fuer analyze_gltyp.ps1 / transform_gltyp.ps1 - kein Teil des VB6-Projekts
Public Sub TestBloecke()
Dim SQL1 As String
Dim DaSt1 As String

'Block 1: MECH - nur dbo./Klammern/Semikolon (mit Umlaut-Spalte)
If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTest1 ORDER BY Zähler"
Else
    SQL1 = "SELECT * FROM qryTest1 ORDER BY [Zähler];"
End If

'Block 2: DIFF - ORDER BY unterscheidet sich wirklich
If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTest2 ORDER BY Feld1"
Else
    SQL1 = "SELECT * FROM qryTest2 ORDER BY [Feld2];"
End If

'Block 3: ONLY - kein Else-Zweig (TSE-Muster)
If GlTyp < 2 Then
    DBCmEx6 "qryTestTSE", "@A", "@B", "@C", "@D", "@E", "@F", vbNullString, vbNullString, vbNullString, 0, 0, 1
End If

'Block 4: ONLY1 - Einzeiler
If GlTyp < 2 Then CmCon.Enabled = False

'Block 5: DIFF + DateNum - Datums-Literale
If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTest3 WHERE (Datum >= CONVERT(DATETIME, '" & DaSt1 & "', 102))"
Else
    SQL1 = "SELECT * FROM qryTest3 WHERE ([Datum] = #" & DaSt1 & "#);"
End If

'Block 6: DIFF + DateNum - DatePart-Formatierungsblock
If GlTyp < 2 Then
    DaSt1 = DatePart("yyyy", Date) & "-" & DatePart("m", Date) & "-" & DatePart("d", Date) & " 00:00:00"
Else
    DaSt1 = DatePart("m", Date) & "/" & DatePart("d", Date) & "/" & DatePart("yyyy", Date)
End If

'Block 7: COMPLEX - ElseIf
If GlTyp < 2 Then
    SQL1 = "A"
ElseIf GlTyp = 2 Then
    SQL1 = "B"
Else
    SQL1 = "C"
End If

'Block 8: MECH mit verschachteltem Select Case (mehrzeilige Zweige)
If GlTyp < 2 Then
    Select Case GlPaK
    Case "P801": SQL1 = "SELECT * FROM dbo.qryTest4 ORDER BY ID0"
    End Select
Else
    Select Case GlPaK
    Case "P801": SQL1 = "SELECT * FROM qryTest4 ORDER BY [ID0];"
    End Select
End If
End Sub
```

- [ ] **Step 2: Nach CP-1252 konvertieren und Zwischendatei löschen**

```powershell
$d = 'C:\Users\schmi\Documents\devcode2\scripts\tests'
[System.IO.File]::WriteAllText("$d\fixture_gltyp.bas",
    [System.IO.File]::ReadAllText("$d\fixture_gltyp.utf8.tmp", [System.Text.Encoding]::UTF8),
    [System.Text.Encoding]::GetEncoding(1252))
Remove-Item "$d\fixture_gltyp.utf8.tmp"
(([System.IO.File]::ReadAllBytes("$d\fixture_gltyp.bas") | Where-Object { $_ -eq 0xE4 }).Count)
```

Expected: letzte Zeile gibt `2` aus (zwei `ä` in „Zähler"-Zeilen als CP-1252-Byte 0xE4) — beweist korrekte Kodierung.

---

## Task 3: Analyse-Skript `analyze_gltyp.ps1` + Fixture-Test

**Files:**
- Create: `C:\Users\schmi\Documents\devcode2\scripts\analyze_gltyp.ps1`

- [ ] **Step 1: Skript schreiben** (Write-Tool)

```powershell
# analyze_gltyp.ps1 - Phase 1 (read-only): klassifiziert alle "If GlTyp < 2"-Bloecke
# Aufruf: powershell -File analyze_gltyp.ps1 [-SourceDir <dir>] [-OutCsv <pfad>]
param(
    [string]$SourceDir = 'C:\Users\schmi\Documents\devcode2\001',
    [string]$OutCsv = 'C:\Users\schmi\Documents\devcode2\scripts\gltyp_report.csv'
)
. (Join-Path $PSScriptRoot 'gltyp_lib.ps1')

$all = New-Object System.Collections.Generic.List[object]
Get-ChildItem $SourceDir -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm' } | ForEach-Object {
    foreach ($b in (Get-GlTypBlocks $_.FullName)) { $all.Add($b) }
}
$all | Select-Object File, StartLine, EndLine, Class, DateNum, SrvText, AccText |
    Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8

'Bloecke gesamt: ' + $all.Count
$all | Group-Object Class | Sort-Object Name | Format-Table Name, Count -AutoSize
'DateNum-markiert: ' + (@($all | Where-Object { $_.DateNum })).Count
'MECH auto-patchbar (ohne DateNum): ' + (@($all | Where-Object { $_.Class -eq 'MECH' -and -not $_.DateNum })).Count
'Report: ' + $OutCsv
```

- [ ] **Step 2: Gegen Fixture laufen lassen (der Test)**

Run:
```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\schmi\Documents\devcode2\scripts\analyze_gltyp.ps1 -SourceDir C:\Users\schmi\Documents\devcode2\scripts\tests -OutCsv C:\Users\schmi\Documents\devcode2\scripts\tests\fixture_report.csv
```

Expected (exakt diese Zählung, sonst Parser fixen bis sie stimmt):
```
Bloecke gesamt: 8
COMPLEX  1
DIFF     3
MECH     2
ONLY     1
ONLY1    1
DateNum-markiert: 2
MECH auto-patchbar (ohne DateNum): 2
```

- [ ] **Step 3: CSV-Umlaut-Probe**

Run: `powershell.exe -NoProfile -Command "(Import-Csv 'C:\Users\schmi\Documents\devcode2\scripts\tests\fixture_report.csv')[0].AccText"`
Expected: enthält `[Zähler];` mit intaktem Umlaut.

---

## Task 4: Analyse über das Projekt + Report-Gate  **[GATE: User]**

- [ ] **Step 1: Analyse über `001/` laufen lassen**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\schmi\Documents\devcode2\scripts\analyze_gltyp.ps1`
Expected: `Bloecke gesamt:` plausibel zu den 603 Grep-Treffern (Block-Ifs + Einzeiler; geringe Abweichung durch COMPLEX/Sonderformen ist ok und wird erklärt).

- [ ] **Step 2: Plausibilisierung gegen bekannte Stellen**

Stichproben im CSV prüfen:
- `basDatKat.bas` Zeile ~1401 (DatePart-Block) → Class=DIFF, DateNum=True
- `basDatKat.bas` Zeile ~1408 (CONVERT/#-SQL) → Class=DIFF, DateNum=True
- `basDatRe.bas` Zeile ~2649 → Class=MECH, DateNum=False
- `clsConn.cls`-Treffer → ONLY/ONLY1 (werden ohnehin nicht angefasst)

- [ ] **Step 3: Zusammenfassung an User, Report-Review abwarten**

Dem User melden: Gesamtzahl, Verteilung je Klasse, MECH-auto-patchbar-Zahl, Pfad zum CSV. **STOPP bis User die DIFF- und DateNum-Listen freigibt.** (Spec Abschnitt 5, Phase-1-Gate.)

---

## Task 5: Helper `SqlDat`/`SqlNum`/`SqlPad` in `basFormat.bas`  **[GATE: User-Verifikation]**

**Files:**
- Create: `C:\Users\schmi\Documents\devcode2\scripts\insert_sql_helpers.ps1`
- Modify (per Byte-Patch): `C:\Users\schmi\Documents\devcode2\001\basFormat.bas` (Anhang ans Dateiende, nach `SqlStr` ~Z. 10401)

- [ ] **Step 1: Einfüge-Skript schreiben** (Write-Tool; Helper-Code ist bewusst umlautfrei)

```powershell
# insert_sql_helpers.ps1 - haengt SqlDat/SqlNum/SqlPad an basFormat.bas an (CP-1252)
param([string]$Target = 'C:\Users\schmi\Documents\devcode2\001\basFormat.bas')
. (Join-Path $PSScriptRoot 'gltyp_lib.ps1')

$text = Read-Vb6Text $Target
if ($text -match 'Public Function SqlDat') { 'SqlDat bereits vorhanden - Abbruch'; exit 1 }
$umlBefore = Get-UmlautCount $text

$helpers = @'
Public Function SqlDat(ByVal DatVal As Date, Optional ByVal MitZeit As Boolean = False) As String
'Liefert einbettungsfertiges Datums-Literal je nach GlTyp (Spec 2026-06-10)
'SQL Server: CONVERT(DATETIME, 'yyyy-mm-dd hh:nn:ss', 120) - Access/Jet: #m/d/yyyy[ h:nn:ss]#
'Aufbau per DatePart - niemals Format$ fuer SQL-Literale (Locale-Risiko)
Dim TmDat As String
Dim TmZei As String
If GlTyp < 2 Then
    TmDat = DatePart("yyyy", DatVal) & "-" & Right$("0" & DatePart("m", DatVal), 2) & "-" & Right$("0" & DatePart("d", DatVal), 2)
    If MitZeit = True Then
        TmZei = Right$("0" & DatePart("h", DatVal), 2) & ":" & Right$("0" & DatePart("n", DatVal), 2) & ":" & Right$("0" & DatePart("s", DatVal), 2)
    Else
        TmZei = "00:00:00"
    End If
    SqlDat = "CONVERT(DATETIME, '" & TmDat & " " & TmZei & "', 120)"
Else
    TmDat = DatePart("m", DatVal) & "/" & DatePart("d", DatVal) & "/" & DatePart("yyyy", DatVal)
    If MitZeit = True Then
        SqlDat = "#" & TmDat & " " & DatePart("h", DatVal) & ":" & Right$("0" & DatePart("n", DatVal), 2) & ":" & Right$("0" & DatePart("s", DatVal), 2) & "#"
    Else
        SqlDat = "#" & TmDat & "#"
    End If
End If
End Function

Public Function SqlNum(ByVal NumVal As Variant) As String
'Dezimalzahl Locale-sicher fuer SQL-Einbettung (Punkt als Dezimaltrenner); Null ergibt Leerstring
'CStr/Format$ liefern in deutscher Locale Komma und sind fuer SQL-Literale verboten
If IsNull(NumVal) Then Exit Function
SqlNum = Trim$(Str$(NumVal))
If InStr(1, SqlNum, "E", vbTextCompare) > 0 Then
    If GlDbg = True Then SErLog "SqlNum Exponentialschreibweise: " & SqlNum & " SqlNum 0"
End If
End Function

Public Function SqlPad(ByVal FldNam As String, ByVal PadLen As Integer) As String
'Liefert Zero-Padding-Ausdruck (ORDER BY/SELECT) je nach GlTyp; ersetzt das SoStr-Muster
If GlTyp < 2 Then
    SqlPad = "RIGHT ('" & String$(PadLen, "0") & "' + CONVERT (varchar(10), " & FldNam & "), " & PadLen & ")"
Else
    SqlPad = "Format$([" & FldNam & "],'" & String$(PadLen, "0") & "')"
End If
End Function
'@

$helpers = $helpers -replace "`r?`n", "`r`n"
$newText = $text.TrimEnd("`r`n") + "`r`n`r`n" + $helpers + "`r`n"
Write-Vb6Text $Target $newText

$check = Read-Vb6Text $Target
"Umlaute vorher/nachher: $umlBefore / " + (Get-UmlautCount $check)
if ((Get-UmlautCount $check) -ne $umlBefore) { throw 'Umlaut-Differenz!' }
if ($check -notmatch 'Public Function SqlPad') { throw 'Einfuegen fehlgeschlagen' }
'OK: Helper angehaengt'
```

- [ ] **Step 2: Backup + Skript ausführen**

```powershell
New-Item -ItemType Directory -Force 'C:\Users\schmi\Documents\devcode2\backups\gltyp_20260610' | Out-Null
Copy-Item 'C:\Users\schmi\Documents\devcode2\001\basFormat.bas' 'C:\Users\schmi\Documents\devcode2\backups\gltyp_20260610\basFormat.pre_helpers.bas'
```
Dann: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\schmi\Documents\devcode2\scripts\insert_sql_helpers.ps1`
Expected: `Umlaute vorher/nachher: <n> / <n>` (gleich) und `OK: Helper angehaengt`

- [ ] **Step 3: Diff lesen**

Run: `git -C C:\Users\schmi\Documents\devcode2\001 diff -- basFormat.bas`
Expected: ausschließlich angehängte Helper am Dateiende, keinerlei andere Änderung.

- [ ] **Step 4: USER-GATE — Compile + Direktfenster-Verifikation auf beiden Backends**

Der User prüft im VB6-Direktfenster (GlTyp temporär umsetzen, danach auf Originalwert zurück!):

```
GlTyp = 1
? SqlDat(DateSerial(2026, 6, 5))
' Erwartet: CONVERT(DATETIME, '2026-06-05 00:00:00', 120)
? SqlDat(DateSerial(2026, 12, 31) + TimeSerial(14, 5, 9), True)
' Erwartet: CONVERT(DATETIME, '2026-12-31 14:05:09', 120)
? SqlNum(1234.56)
' Erwartet: 1234.56   (Punkt! Bei deutschem CStr waere es 1234,56)
? SqlPad("Mandant", 8)
' Erwartet: RIGHT ('00000000' + CONVERT (varchar(10), Mandant), 8)
GlTyp = 3
? SqlDat(DateSerial(2026, 6, 5))
' Erwartet: #6/5/2026#
? SqlDat(DateSerial(2026, 6, 5) + TimeSerial(8, 30, 0), True)
' Erwartet: #6/5/2026 8:30:00#
? SqlPad("Mandant", 8)
' Erwartet: Format$([Mandant],'00000000')
```

Zusätzlich je eine echte Probe-Query gegen SQL-Server- und Access-Testdatenbank (z. B. `SELECT`-Filter auf eine Datumsspalte mit `SqlDat`). **STOPP bis grün.** (Spec Abschnitt 6.)

- [ ] **Step 5: Commit**

```
git -C C:\Users\schmi\Documents\devcode2\001 add basFormat.bas
git -C C:\Users\schmi\Documents\devcode2\001 commit -m "basFormat: SQL-Dialekt-Helper SqlDat/SqlNum/SqlPad (Spec 2026-06-10)"
```

---

## Task 6: Transformations-Skript `transform_gltyp.ps1` + Fixture-Test

**Files:**
- Create: `C:\Users\schmi\Documents\devcode2\scripts\transform_gltyp.ps1`

- [ ] **Step 1: Skript schreiben** (Write-Tool)

```powershell
# transform_gltyp.ps1 - Phase 2: fuehrt MECH-Bloecke (ohne DateNum) auf den universellen Dialekt zusammen
# Universal-Form = Access-Zweig (Klammern bleiben) ohne dbo. (hat er nie) und ohne Schluss-Semikolon
# Aufruf: -FileName basDATEV.bas           -> nur Preview-Datei erzeugen
#         -FileName basDATEV.bas -Apply    -> Backup + Patch + Verifikation
param(
    [Parameter(Mandatory = $true)][string]$FileName,
    [switch]$Apply,
    [string]$SourceDir = 'C:\Users\schmi\Documents\devcode2\001',
    [string]$BackupDir = 'C:\Users\schmi\Documents\devcode2\backups\gltyp_20260610'
)
. (Join-Path $PSScriptRoot 'gltyp_lib.ps1')

$path = Join-Path $SourceDir $FileName
$blocks = @(Get-GlTypBlocks $path | Where-Object { $_.Class -eq 'MECH' -and -not $_.DateNum })
'MECH-Bloecke (ohne DateNum) in {0}: {1}' -f $FileName, $blocks.Count
if (-not $blocks.Count) { exit 0 }

$origText = Read-Vb6Text $path
$lines = $origText -split "`r`n"
$preview = New-Object System.Collections.Generic.List[string]

# rueckwaerts ersetzen, damit Zeilenindizes der frueheren Bloecke gueltig bleiben
foreach ($b in ($blocks | Sort-Object StartLine -Descending)) {
    $accIndent = ''
    if ($b.AccLines[0] -match '^(\s*)') { $accIndent = $Matches[1] }
    $newLines = @(foreach ($al in $b.AccLines) {
        $t = $al
        if ($t.StartsWith($accIndent)) { $t = $b.IfIndent + $t.Substring($accIndent.Length) }
        $t -replace ';"', '"'
    })
    $preview.Add(('--- {0}:{1}-{2} ---' -f $b.File, $b.StartLine, $b.EndLine))
    foreach ($x in $lines[($b.StartLine - 1)..($b.EndLine - 1)]) { $preview.Add('ALT | ' + $x) }
    foreach ($x in $newLines) { $preview.Add('NEU | ' + $x) }
    $preview.Add('')
    $before = @(); if ($b.StartLine -gt 1) { $before = $lines[0..($b.StartLine - 2)] }
    $after = @(); if ($b.EndLine -lt $lines.Count) { $after = $lines[$b.EndLine..($lines.Count - 1)] }
    $lines = @($before) + $newLines + @($after)
}

$previewPath = Join-Path $PSScriptRoot ('gltyp_preview_' + $FileName + '.txt')
[System.IO.File]::WriteAllLines($previewPath, $preview)
'Preview: ' + $previewPath

if ($Apply) {
    if (-not (Test-Path $BackupDir)) { New-Item -ItemType Directory -Force $BackupDir | Out-Null }
    Copy-Item $path (Join-Path $BackupDir $FileName) -Force
    $newText = $lines -join "`r`n"
    Write-Vb6Text $path $newText
    if ((Read-Vb6Text $path) -cne $newText) { throw 'VERIFIKATION FEHLGESCHLAGEN: Datei weicht vom Soll ab' }
    'OK: {0} Bloecke gepatcht, Backup unter {1}' -f $blocks.Count, $BackupDir
}
```

- [ ] **Step 2: Test gegen Arbeitskopie der Fixture**

```powershell
$t = 'C:\Users\schmi\Documents\devcode2\scripts\tests'
Copy-Item "$t\fixture_gltyp.bas" "$t\fixture_work.bas" -Force
```
Run:
```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\schmi\Documents\devcode2\scripts\transform_gltyp.ps1 -FileName fixture_work.bas -SourceDir C:\Users\schmi\Documents\devcode2\scripts\tests -BackupDir C:\Users\schmi\Documents\devcode2\scripts\tests\backup -Apply
```
Expected: `MECH-Bloecke (ohne DateNum) in fixture_work.bas: 2` und `OK: 2 Bloecke gepatcht...`

- [ ] **Step 3: Ergebnis inhaltlich prüfen**

`fixture_work.bas` lesen (PowerShell `Read-Vb6Text`, nicht Read-Tool-Annahme über Encoding) und verifizieren:
- Block 1 ist ersetzt durch genau eine Zeile auf Spalte 0 (If stand auf Spalte 0): `SQL1 = "SELECT * FROM qryTest1 ORDER BY [Zähler]"` (Umlaut intakt, kein Semikolon, Klammern erhalten)
- Block 8 ist ersetzt durch das dedentete `Select Case` mit `SQL1 = "SELECT * FROM qryTest4 ORDER BY [ID0]"`
- Blöcke 2–7 unverändert
- erneuter Analyse-Lauf über `fixture_work.bas` meldet `MECH ... 0` auto-patchbar

- [ ] **Step 4: Aufräumen**

```powershell
Remove-Item 'C:\Users\schmi\Documents\devcode2\scripts\tests\fixture_work.bas', 'C:\Users\schmi\Documents\devcode2\scripts\tests\backup' -Recurse -Force
```

---

## Task 7: Pilot-Modul `basDATEV.bas`  **[GATE: User]**

- [ ] **Step 1: Preview erzeugen** — `transform_gltyp.ps1 -FileName basDATEV.bas`
- [ ] **Step 2: Preview vollständig lesen** (`scripts\gltyp_preview_basDATEV.bas.txt`) — jede ALT/NEU-Paarung manuell prüfen: NEU muss exakt dem Access-Zweig ohne Schluss-Semikolon entsprechen
- [ ] **Step 3: Anwenden** — `transform_gltyp.ps1 -FileName basDATEV.bas -Apply`
- [ ] **Step 4: Git-Diff lesen** — `git -C ...\001 diff -- basDATEV.bas`; zusätzlich Umlaut-Probe: `git diff` darf keine Umlaut-Zeilen außerhalb der ersetzten Blöcke zeigen
- [ ] **Step 5: USER-GATE — VB6-Compile + DATEV-Export-Smoke-Test** (beide Backends). **STOPP bis grün.**
- [ ] **Step 6: Commit** — `git -C ...\001 add basDATEV.bas` + `git commit -m "basDATEV: GlTyp-SQL-Verzweigungen vereinheitlicht (MECH-Bloecke)"`

---

## Task 8: Wellen 2–5 — alle übrigen Module

Je Modul dieser vollständige Ablauf (`<Modul>` = Dateiname):

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\schmi\Documents\devcode2\scripts\transform_gltyp.ps1 -FileName <Modul>
REM Preview scripts\gltyp_preview_<Modul>.txt vollstaendig lesen: jede ALT/NEU-Paarung pruefen
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\schmi\Documents\devcode2\scripts\transform_gltyp.ps1 -FileName <Modul> -Apply
git -C C:\Users\schmi\Documents\devcode2\001 diff -- <Modul>
git -C C:\Users\schmi\Documents\devcode2\001 add <Modul>
git -C C:\Users\schmi\Documents\devcode2\001 commit -m "<Modul>: GlTyp-SQL-Verzweigungen vereinheitlicht (MECH-Bloecke)"
```

**Nach jeder Welle:** User-Compile + Smoke-Test, erst dann nächste Welle.

- [ ] **Welle 2 (klein):** `clsData.cls`, `frmAdrWord.frm`, `frmReSam.frm`, `frmAbschl.frm`, `frmReExpo.frm`, `basFormat.bas`, `frmBuExp.frm`, `frmIfap.frm`, `basLabor.bas`
  - Übersprungen solange WIP unkommittiert: `frmOptions.frm`, `basLayout.bas`, `frmMandant.frm`
- [ ] **Welle 2 Smoke-Test:** Rechnungs-/Buchungsexport-Dialoge öffnen, Labor-Liste
- [ ] **Welle 3 (mittel):** `frmZeitraum.frm`, `frmAdrFilt.frm`, `clsLisLab.cls`, `basDaMa.bas`
- [ ] **Welle 3 Smoke-Test:** Adressfilter, Zeitraum-Auswahl, Labor-Listen, Mandanten-Daten
- [ ] **Welle 4 (groß):** `basDatRe.bas`, `basMain.bas`, `basDaAdr.bas`
- [ ] **Welle 4 Smoke-Test:** Rechnungserstellung/-export, Adresssuche („O'Brien"-Apostroph-Fall erneut), Programmstart/-ende
- [ ] **Welle 5 (sehr groß):** `basDatKat.bas`, `basData.bas`
- [ ] **Welle 5 Smoke-Test:** Adresssuche alle Suchindizes, Terminkalender/Wiedervorlage, Krankenblatt, PAD-Summen gegen Altexport

Nicht transformiert (begründet): `clsConn.cls` (Verbindungslogik).

---

## Task 9: DIFF- und DateNum-Fälle einzeln abarbeiten

**Vorgehen je Fall (aus `gltyp_report.csv`, Klassen DIFF sowie MECH+DateNum):**

1. Stelle im Code lesen (Read-Tool nur zum Lesen), beide Zweige verstehen.
2. Entscheidung dokumentieren: (a) Helper-Umbau (`SqlDat`/`SqlNum`/`SqlPad`), (b) begründet verzweigt belassen (z. B. Stored-Proc/TSE), (c) echte Altlast → Liste für User.
3. Umbau ausschließlich per Einzel-Byte-Patch-Skript nach dem Muster der bestehenden `scripts/patch_*.ps1` (Zwei-Phasen: erst Treffer-Verifikation, dann Ersetzung; Umlaut-Count vor/nach).
4. Bei Datums-Gleichheit-vs.-Bereich-Differenzen: Zielbild ist der Bereichsfilter — **jede** solche semantische Angleichung einzeln dem User vorlegen (Spec Abschnitt 5).

**Durchgerechnetes Beispiel** (`basDatKat.bas:1401–1412`): Die beiden If-Blöcke (DatePart-Formatierung + SQL) werden zusammen ersetzt durch:

```vb
SQL1 = "SELECT * FROM qrySimAbSav WHERE (([ID0] = " & GlAdr & ") AND ([ID1] = 107) AND ([Datum] >= " & SqlDat(GlTag(1)) & ") AND ([Datum] < " & SqlDat(GlTag(1) + 1) & "))"
```

(Semantik-Angleichung: SQL Server filterte `>= Tag AND <= Folgetag-Mitternacht`, Access `= Tag`; vereinheitlicht auf `>= Tag AND < Folgetag` — vor Anwendung dem User vorlegen.)

- [ ] **Step 1: DIFF-Liste aus CSV gruppieren** (gleiche Muster bündeln: SoStr/SqlPad-Gruppe, Datums-Gruppe, Rest)
- [ ] **Step 2: SoStr/SqlPad-Gruppe umbauen** (mechanischstes Muster zuerst), Commit je Datei-Gruppe
- [ ] **Step 3: Datums-Gruppe mit User-Freigabe je Semantik-Änderung umbauen**, Commit je Datei-Gruppe
- [ ] **Step 4: Rest-Fälle einzeln entscheiden/dokumentieren**, verbleibende Verzweigungen mit Begründungskommentar versehen ist NICHT nötig — die Spec-Kategorien (a)–(c) genügen
- [ ] **Step 5: User-Compile + Smoke-Test beider Backends**

---

## Task 10: Abschluss-Verifikation

- [ ] **Step 1: Erfolgskriterium 1 der Spec prüfen**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\schmi\Documents\devcode2\scripts\analyze_gltyp.ps1 -OutCsv C:\Users\schmi\Documents\devcode2\scripts\gltyp_report_final.csv`
Expected: `MECH auto-patchbar (ohne DateNum): 0`; alle verbleibenden Blöcke gehören zu Kategorie (a) Verbindungslogik, (b) Helper, (c) Stored-Proc/TSE oder sind dokumentierte User-Entscheide.

- [ ] **Step 2: Spec-Status aktualisieren** (Status-Zeile auf „Umgesetzt <Datum>", Write/Edit erlaubt — Markdown, kein VB6)
- [ ] **Step 3: Abschlussbericht an User** — Zahlen vorher/nachher, Liste der bewusst verbliebenen Verzweigungen, Hinweis: `backups/gltyp_20260610/` nach grünem Gesamttest löschbar

---

## Addendum A (2026-06-10, nach Quality-Review Task 1): Lib v2

Der Code-Review der ersten `gltyp_lib.ps1`-Fassung fand drei False-MECH-Risiken im Normalisierer
(Klammer-/Whitespace-/Semikolon-Entfernung wirkte auch innerhalb von SQL-Datenliteralen) sowie
Lücken im DateNum-Netz und Parser-Randfälle. **Die Code-Blöcke in Task 1–3 oben gelten als
überholt, wo dieses Addendum sie ersetzt.**

### A.1 `Get-NormalizedSql` v2 — quote-bewusst

Vertrag: Zwei Branches gelten genau dann als gleich (→ MECH), wenn ihr Code-/SQL-Skelett nach
Kosmetik-Normalisierung gleich ist **und** alle SQL-Datenliterale **byte-exakt** gleich sind.

Algorithmus, pro Branch (`SrvLines`/`AccLines`), Zustände werden je Zeile zurückgesetzt:

1. Zeichenweiser Scan mit zwei Zuständen: `inVb` (innerhalb VB6-`"…"`), `inSql` (innerhalb
   SQL-`'…'`, nur toggelbar wenn `inVb`).
   - `"` toggelt `inVb`; Escape `""` (zwei `"` in Folge bei `inVb=$true`) zählt als Literalzeichen,
     kein Toggle.
   - `'` bei `inVb=$true` toggelt `inSql`; `'` bei `inVb=$false` beginnt VB6-Kommentar →
     Rest der Zeile verwerfen (Kommentare beeinflussen Verhalten nicht; reduziert False-DIFF).
   - Zeichen bei `inSql=$true` → Daten-Puffer (aktuelles Literal). Beim Wechsel `inSql`
     false→true wird im Skelett ein Platzhalter `[char]1` eingefügt und ein neues
     Daten-Segment begonnen.
   - Alle übrigen Zeichen → Skelett-Puffer.
2. Endet eine Zeile mit `inSql=$true` (SQL-String per Konkatenation über Zeilen offen):
   Branch nicht normalisierbar → Rückgabe `([char]2 + 'RAW' + [char]2 + (Zeilen roh gejoint))`
   → faktisch DIFF, außer beide Branches sind byte-identisch.
3. Skelett-Transformationen (NUR Skelett, nie Daten-Segmente), Reihenfolge:
   `[\[\]]` → '' · `dbo\.` → '' · `\s+` → ' ' · Trim · am Ende des Gesamt-Skeletts `;\s*$` → ''
   (entfernt exakt das Statement-Endsemikolon; ein `;` mitten im Skelett — z. B. Datenlisten
   `"Mo;Di;" & x` — bleibt und erzeugt DIFF).
4. Rückgabe: Skelett + `[char]1` + (Daten-Segmente exakt, mit `[char]1` gejoint).

Damit gelöst: LIKE-Escapes `[_]`/`[%]` in Datenliteralen bleiben erhalten (C1), Whitespace in
Datenliteralen bleibt exakt (C2), nur das Statement-Endsemikolon wird entfernt (C3),
Kommentar-Differenzen erzeugen kein False-DIFF mehr.

### A.2 Parser-Fixes

- Depth-Increment-Regex härten: `'^\s*#?If\b[^'']*\bThen\s*(''.*)?$'` (kein Match, wenn vor
  `Then` ein Apostroph liegt — Kommentare, die auf „then" enden, korrumpieren die Tiefe nicht mehr).
- Guard: wird `$depth -lt 0`, Block als COMPLEX klassifizieren (Parser out of sync).
- `^\s*(Else|End If)\s*:` auf Tiefe 0 → COMPLEX (Doppelpunkt-Statement-Separator).
- Einzeiler mit Zeilenfortsetzung (`If GlTyp < 2 Then _`) → COMPLEX statt ONLY1.
- Einzeiler mit Inline-Else (`If GlTyp < 2 Then x Else y`) → neue Klasse **ONLY1E**
  (ehrliches Inventar: es existiert ein Access-Zweig).
- Skelett-Transform-Reihenfolge Klammern-vor-dbo deckt `[dbo].`-Schreibweise gratis ab.

### A.3 DateNum-Regex erweitern

`'"#|#"|CONVERT\s*\(\s*(SMALL)?DATETIME|CONVERT\s*\(\s*N?(VAR)?CHAR|DatePart\s*\(|Format\$?\s*\(|FormatNumber\s*\(|FormatDateTime\s*\(|CDate\s*\('`

(deckt `Format(` ohne `$`, `CONVERT(VARCHAR, …, 104)`-Richtung, `CDate`, `FormatNumber`,
`FormatDateTime` ab.)

### A.4 Fixture v2 (Task 2) — fünf zusätzliche Blöcke

- Block 9 (DIFF): LIKE-Escape — Srv `… LIKE 'A[_]B'` vs. Acc `… LIKE 'A_B'`
- Block 10 (DIFF): Mid-Skelett-Semikolon — Srv `SQL1 = "Mo;Di" & SuStr` vs. Acc `SQL1 = "Mo;Di;" & SuStr`
- Block 11 (ONLY1E): `If GlTyp < 2 Then CmCon.Enabled = False Else CmCon.Enabled = True`
- Block 12 (DIFF): Datenliteral-Whitespace — Srv `"REPLACE(N, '  ', ' ')"` vs. Acc `"REPLACE(N, ' ', ' ')"`
- Block 13 (MECH): identisch bis auf Kosmetik **und** ein Trailing-Kommentar nur im Srv-Zweig
  (beweist Kommentar-Stripping)

**Neue Erwartungswerte Task 3 Fixture-Lauf:** Blöcke gesamt 13 — MECH 3, DIFF 6, ONLY 1,
ONLY1 1, ONLY1E 1, COMPLEX 1; DateNum-markiert 2; MECH auto-patchbar (ohne DateNum) 3.

### A.5 Task 3 zusätzlich: GlTyp-Coverage-Report

`analyze_gltyp.ps1` ergänzt einen Reconciliation-Report `scripts/gltyp_uncovered.csv`:
alle Zeilen mit `\bGlTyp\b`, die weder Header noch innerhalb [StartLine..EndLine] eines
erkannten Blocks liegen (Datei, Zeile, Inhalt). Bekannter Treffer, der dort erscheinen MUSS:
`basMain.bas:37502` (`ElseIf GlTyp < 2 Then` — vom Parser bewusst nicht als Block erfasst).
Initialisierungs-/Zuweisungszeilen (`GlTyp = …`, `clsConn`-Logik) erscheinen dort erwartungsgemäß.

### A.6 Task 6 zusätzlich: Roundtrip-Guard + Semikolon-Regel

- `transform_gltyp.ps1` prüft vor jeder Transformation:
  `($lines -join "`r`n") -ceq $origText`, sonst Abbruch (schützt Dateien mit vereinzelten
  LF-only-Zeilenenden vor stiller Normalisierung beim Zurückschreiben).
- Semikolon-Strip beim Bauen der Ersatzzeilen NUR am Zeilenende:
  `-replace ';"(\s*)$', '"$1'` statt global `-replace ';"', '"'` (ein `;"` mitten in der
  Zeile — Datenliteral-Ende vor Konkatenation — bleibt unangetastet; bei MECH-Blöcken hatte
  der Srv-Zweig dasselbe `;` im Skelett, Beibehalten ist also korrekt).

## Offene Abhängigkeiten

- Tasks 4, 5, 7 und jede Welle in Task 8 enthalten **harte User-Gates** (Report-Review, Compile, Smoke-Tests) — keine Fortsetzung ohne Freigabe.
- `frmOptions.frm`, `basLayout.bas`, `frmMandant.frm` erst nach User-WIP-Commit transformieren (dann wie Task 7).
