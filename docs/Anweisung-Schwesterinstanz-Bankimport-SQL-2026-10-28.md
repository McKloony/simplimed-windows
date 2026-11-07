# SQL-Server-Anpassung fuer den Bankimport

## Auftrag

Der Benutzer hat die Anpassung der Access-Datenbanken in
`Z:\Data\Database\Abrechnung` beauftragt. Die SQL-Server-Anpassung soll durch
die Schwesterinstanz erfolgen. Dieses Dokument ist die technische Uebergabe;
von dieser Instanz wurde keine SQL-Server-Datenbank geaendert.

Access-Stand nach Freigabe des Vollzugriffs: Alle neun gefundenen Access-
Datenbanken im Zielordner einschliesslich `_Datenbank`, `_Beispieldaten`, `_CrossMed` und `_Abfragen`
wurden gesichert und um diesen Index ergaenzt. Indexdefinition und erhaltene
Umsatzzahlen sind nach erneutem Oeffnen bestaetigt. Siehe
`Pruefbericht-Access-Bankindex-2026-10-28.txt`.
Die beiden DAX-Dateien in `_Abfragen` wurden am 28.10.2026 nachtraeglich erfasst
und angepasst; sie waren im ersten Lauf mit sieben Dateien nicht enthalten.

Ergaenze in der vorgesehenen SimpliMed/CrossMed-Datenbank einen passenden
Index fuer den Bestandsabgleich des Bankimports. Keine Buchungen loeschen,
keine Dublettenbereinigung, keine Datentyp- oder View-Aenderung.

## Grundlage und Befund

- Aktuellerer lokaler Export: `895b.sql`, Dateistand 16.10.2026, 08:21:27.
- SHA-256: `56D8F1B68452F40EA9C5D55DEB401295297AE2404E42A2849714D1315C6DD43A`.
- `Tabelle_Banking`: Zeile 1043; bestehende Indizes: Zeilen 53224 bis 53303.
- Dort vorhanden: einzelne Indizes auf Datum, GuiID, ID0, ID1, IDM und IDR
  sowie der gruppierte Primaerschluessel auf ID2.
- Ein gemeinsamer Index auf IDB, IDM, Datum fehlt im Export. Den Live-Stand
  vor der Ausfuehrung pruefen; der Export belegt nicht den aktuellen Serverstand.
- `sql/SQL_Server.sql` ist aelter (17.04.2026) und enthaelt ebenfalls keinen
  passenden gemeinsamen Index.
- Die neue Routine `BankLoad` in `001/basBankimp.bas` liest einmal je Import:

```sql
SELECT [ID2], [Datum], [Betrag], [IDKurz], [Suche]
FROM [Tabelle_Banking]
WHERE [IDB] = ? AND [IDM] = ? AND [Datum] >= ? AND [Datum] < ?
```

Die Fragezeichen werden durch ADO gebunden: Konto, Mandant, erster Tag,
exklusive obere Grenze am Tag nach dem letzten Datum.

## Vorgehen

1. Die richtige Anwendungsdatenbank und Tabelle identifizieren. Den vorhandenen
   Sicherungs-/Wiederherstellungsstand kontrollieren. Indexanlage in einem
   geeigneten Wartungsfenster ausfuehren; sie kann Schreibzugriffe blockieren.
2. Bestehende Indizes samt Schluesselreihenfolge lesen:

```sql
SELECT DB_NAME() AS [Datenbank],
       OBJECT_SCHEMA_NAME(i.[object_id]) AS [SchemaName],
       i.[name] AS [IndexName], i.[type_desc], i.[is_unique],
       i.[is_disabled], i.[is_hypothetical], i.[has_filter], i.[filter_definition],
       ic.[key_ordinal], ic.[is_included_column], c.[name] AS [Spalte]
FROM sys.indexes AS i
JOIN sys.index_columns AS ic
  ON ic.[object_id] = i.[object_id] AND ic.[index_id] = i.[index_id]
JOIN sys.columns AS c
  ON c.[object_id] = ic.[object_id] AND c.[column_id] = ic.[column_id]
WHERE i.[object_id] = OBJECT_ID(N'Tabelle_Banking', N'U')
ORDER BY i.[index_id], ic.[key_ordinal], ic.[index_column_id]
```

3. Ein aktiver, ungefilterter, nicht hypothetischer Rowstore-Index mit den
   ersten drei **Schluesselspalten** `(IDB, IDM, Datum)` oder
   `(IDM, IDB, Datum)` reicht aus. Weitere nachfolgende Spalten sind erlaubt.
   INCLUDE-Spalten zaehlen nicht als Schluesselspalten. Bei einem passenden
   vorhandenen Index nichts zusaetzlich anlegen und dessen Namen protokollieren.
4. Fehlt ein solcher Index, genau diesen nicht eindeutigen Index anlegen:

```sql
CREATE INDEX [IX_Banking_IDB_IDM_Datum]
ON [Tabelle_Banking] ([IDB], [IDM], [Datum])
```

   Das SQL liegt auch in `sql/Bankimport_Index.sql`. Ist dieser Name bereits
   mit anderer Definition belegt, den Konflikt untersuchen; keinen bestehenden
   Index automatisch loeschen oder ueberschreiben. Keine UNIQUE-Einschraenkung:
   mehrere Umsaetze desselben Kontos/Mandanten am selben Tag sind legitim.
5. Die Metadatenabfrage erneut ausfuehren und Reihenfolge, Aktivitaet und
   Nicht-Eindeutigkeit des neu angelegten Index bestaetigen.
6. Die reale BankLoad-Abfrage mit repraesentativem Konto, Mandant und Zeitraum
   vor/nach der Anpassung vergleichen: Ausfuehrungsplan, logische Reads und
   Laufzeit. Bei kleinen Tabellen kann ein Scan weiterhin guenstiger sein.
   Keine Umsatztexte oder Zugangsdaten in das Ergebnisprotokoll aufnehmen.
7. Import und Wiederholungsimport mit dem neuen Programmstand auf einer
   Testkopie pruefen. Den neuen Index und das Ergebnis in der ueblichen
   Schema-/Migrationsverwaltung dokumentieren, damit Folgeinstallationen
   denselben Stand erhalten. Den Export nicht ungeprueft komplett ausfuehren.

## Grenzen und Ruecknahme

Der Index unterstuetzt die gefilterte Suche. Die einzelnen `RS.Update`-
Schreibzugriffe des Imports bleiben bestehen. Ein schnellerer Gesamtimport
ist deshalb erst durch Messung belegt. Zusaetzliche Indizes brauchen Speicher
und muessen beim Schreiben gepflegt werden. Bestehende Einzelindizes bleiben
unveraendert, da andere Programmteile sie verwenden koennen.

Falls die Messung keinen Nutzen zeigt oder Probleme entstehen, darf nur der
in diesem Auftrag neu angelegte Index wieder entfernt werden:

```sql
DROP INDEX [IX_Banking_IDB_IDM_Datum] ON [Tabelle_Banking]
```

Ruecknahme nur ausfuehren, wenn das Aenderungsprotokoll dessen Neuanlage belegt.
Keine Datenbankwiederherstellung ueber spaeter hinzugekommene Buchungen.

## Rueckmeldung der Schwesterinstanz

Melde Zielserver/Datenbank, Zeitpunkt, vorhandener oder neu angelegter Index,
Pruefergebnis und Messwerte. Keine Passwoerter, Verbindungszeichenfolgen oder
Patienten-/Umsatzdaten in die Rueckmeldung schreiben.

SQL Server wurde hier nicht ausgefuehrt. Index-DDL und BankLoad-Abfrage wurden
bereits auf einer separaten temporaeren ACE-Datenbank geprueft.

Microsoft-Referenzen: [Indexauswahl](https://learn.microsoft.com/en-us/sql/relational-databases/sql-server-index-design-guide),
[Indexspalten und Schluesselreihenfolge](https://learn.microsoft.com/en-us/sql/relational-databases/system-catalog-views/sys-index-columns-transact-sql).
