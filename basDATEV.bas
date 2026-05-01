Attribute VB_Name = "basDATEV"
Option Explicit
'================================================================================
' basDATEV.bas - DATEV Export Module for SimpliMed
'================================================================================
' Purpose:     Export accounting postings to DATEV CSV format (Buchungsstapel)
'              Export DATEV XML for document linking (Belege)
'
' Replaces:    frmBuExp.FWeit(), basDaMa.S_Expor(), basData.S_BuEx(),
'              basDatRe.S_DaExB(), basDatRe.S_DaExP(), basDatRe.S_DaExF(),
'              basDatRe.S_DaExX()
'
' DATEV Specs: EXTF Format Version 700, Format Category 21 (Buchungsstapel)
'              XML Namespace http://xml.datev.de/bedi/tps/document/v06.0
'
' Author:      SimpliMed Development
' Created:     2025
'================================================================================

'--------------------------------------------------------------------------------
' Module-Level Constants
'--------------------------------------------------------------------------------
Private Const DATEV_VERSION As Integer = 700
Private Const DATEV_FORMAT_CATEGORY As Integer = 21
Private Const DATEV_FORMAT_NAME As String = "Buchungsstapel"
Private Const DATEV_SEPARATOR As String = ";"
Private Const DATEV_TEXT_QUALIFIER As String = """"
Private Const DATEV_CURRENCY As String = "EUR"
Private Const DATEV_ENCODING As String = "UTF-8"
Private Const DATEV_GENERATING_SYSTEM As String = "SimpliMed"

Private Const XML_NAMESPACE As String = "http://xml.datev.de/bedi/tps/document/v06.0"
Private Const XML_SCHEMA_LOCATION As String = "http://xml.datev.de/bedi/tps/document/v06.0 document_v060.xsd"
Private Const XML_VERSION As String = "6.0"

' Ledger XML namespace and schema (for Option B: strukturierte Belegsatzdaten)
Private Const LEDGER_XML_NAMESPACE As String = "http://xml.datev.de/bedi/tps/ledger/v060"
Private Const LEDGER_XML_SCHEMA_LOCATION As String = "http://xml.datev.de/bedi/tps/ledger/v060 Belegverwaltung_online_ledger_import_v060.xsd"
Private Const LEDGER_XML_VERSION As String = "6.0"
Private Const LEDGER_XML_DATA_TEXT As String = "Kopie nur zur Verbuchung berechtigt nicht zum Vorsteuerabzug"

Private Const BELEGLINK_PREFIX As String = "BEDI"
Private Const BELEGINFO_PATIENT_ART As String = "PATIENTENNR"
Private Const BELEGINFO_DEBITORNR_ART As String = "DEBITORENNR"

' Ledger XML extension types for structured document linking
Private Const XML_EXT_ACCOUNTS_RECEIVABLE As String = "accountsReceivableLedger"  ' Ausgangsrechnungen (Einnahmen)
Private Const XML_EXT_ACCOUNTS_PAYABLE As String = "accountsPayableLedger"        ' Eingangsrechnungen (Ausgaben)
Private Const XML_EXT_CASH_LEDGER As String = "cashLedger"                        ' Kassenbuch (Einnahmen + Ausgaben)
Private Const XML_EXT_FILE As String = "File"

' Ledger XML folder types (property key="3")
Private Const XML_FOLDER_OUTGOING As String = "Ausgangsrechnungen"
Private Const XML_FOLDER_INCOMING As String = "Eingangsrechnungen"
Private Const XML_FOLDER_CASH As String = "Kasse"                                  ' Kassenbuch folder name

Private Const MAX_BOOKING_TEXT_LENGTH As Integer = 60
Private Const MAX_BELEGFELD1_LENGTH As Integer = 12
Private Const MAX_FILENAME_LENGTH As Integer = 46

' Performance tuning constants
Private Const PROGRESS_UPDATE_INTERVAL As Integer = 25      ' Update UI every N records
Private Const INITIAL_BUFFER_SIZE As Long = 65536           ' 64KB initial string buffer estimate
Private Const LINE_BUFFER_SIZE As Integer = 2048            ' Estimated bytes per CSV line
Private Const DOEVENTS_INTERVAL As Integer = 50             ' DoEvents every N records

' Amount constraints
Private Const MIN_VALID_AMOUNT As Currency = 0.01           ' Minimum valid amount (1 cent)
Private Const MAX_VALID_AMOUNT As Currency = 999999999.99   ' Maximum DATEV amount

'--------------------------------------------------------------------------------
' Module-Level Type Definitions
'--------------------------------------------------------------------------------

' Export configuration passed from callers
Public Type DATEV_ExportConfig
    ExportPath As String            ' Target export directory
    ExportFileName As String        ' User-selected filename (without extension)
    Beraternummer As Long           ' DATEV Beraternummer (from GlDvB)
    Mandantennummer As Long         ' DATEV Mandantennummer (from GlDvM)
    DateFrom As Date                ' Export date range start
    DateTo As Date                  ' Export date range end
    WJBeginn As Date                ' Wirtschaftsjahrbeginn
    FourDigitAccounts As Boolean    ' GldKt: Use 4-digit accounts (vs 6-digit)
    SwapDebitCredit As Boolean      ' GlTSH: Swap SOLL/HABEN
    IncludePatientNumber As Boolean ' GlDeN: Include patient number
    ReplaceAccountWithDebtor As Boolean ' GlDeE: Replace account (Konto) with debtor number
    IncludePatientName As Boolean   ' GlDaP: Include patient name
    ExportDocuments As Boolean      ' GlBlE: Export linked documents (Belege)
    MandantNr As Long               ' Current mandant number for filtering
    EmailAfterExport As Integer     ' 0=No, 1=Yes - send via email
    CompressOutput As Boolean       ' Create ZIP archive
    EncryptOutput As Boolean        ' Encrypt ZIP archive
    EncryptPassword As String       ' Encryption password
    ExportCSV As Boolean            ' Generate DATEV CSV file
    ExportXML As Boolean            ' Generate DATEV Beleglink XML file
    UseLedgerXML As Boolean         ' Option B: Generate ledger.xml for structured data import
End Type

' Single booking record for CSV export
Public Type DATEV_BookingRecord
    Umsatz As Currency              ' Amount (always positive)
    SollHaben As String             ' "S" or "H"
    Konto As Long                   ' Debit account
    Gegenkonto As Long              ' Credit account
    BUSchluessel As String          ' Tax key (Steuerschluessel)
    Belegdatum As Date              ' Document date
    Belegfeld1 As String            ' Document field 1 (invoice number)
    Belegfeld2 As String            ' Document field 2
    Buchungstext As String          ' Booking description
    Beleglink As String             ' BEDI + GUID for document linking
    Steuersatz As Single            ' Tax rate
    guid As String                  ' Unique identifier
    Kostenstelle1 As String         ' Cost center 1
    Kostenstelle2 As String         ' Cost center 2
    Festschreibung As Boolean       ' Lock flag
    PatientNummer As String         ' Patient number (optional)
    PatientName As String           ' Patient name (optional)
End Type

' Document record for XML export
Public Type DATEV_DocumentRecord
    guid As String                  ' Document GUID (same as Beleglink without BEDI prefix)
    FileName As String              ' PDF filename
    FilePath As String              ' Full path to PDF
    Description As String           ' Document description
    DocumentDate As Date            ' Document date
    Keywords As String              ' Search keywords
    InvoiceNumber As String         ' Related invoice number
End Type

' Export result for status reporting
Public Type DATEV_ExportResult
    success As Boolean
    CSVFilePath As String
    XMLFilePath As String
    ZipFilePath As String
    RecordCount As Long
    DocumentCount As Long
    ErrorMessage As String
    ErrorCode As Long
    ' Konsistenzpruefung
    PDFFileCount As Long            ' Anzahl physisch vorhandener PDF-Dateien
    XMLDocumentCount As Long        ' Anzahl Dokument-Referenzen in document.xml
    ConsistencyOK As Boolean        ' True wenn PDF-Anzahl = XML-Referenzen
    ConsistencyMessage As String    ' Meldung bei Inkonsistenz
    ExportPath As String            ' Pfad fuer Cleanup-Funktion
    ' Belegzaehlung nach Typ
    DebitDocCount As Long           ' Anzahl Debitorenbelege (Einnahmen, BuTyp=2)
    KreditDocCount As Long          ' Anzahl Kreditorenbelege (Ausgaben, BuTyp=1)
End Type

'--------------------------------------------------------------------------------
' Module-Level Variables
'--------------------------------------------------------------------------------
Private m_clFil As clsFile                  ' File operations class
Private m_clLis As clsLisLab                ' PDF/Beleg generation class
Private m_Cancelled As Boolean              ' User cancelled export
Private m_Config As DATEV_ExportConfig      ' Current export configuration
Private m_ZipFiles() As String              ' Files to include in ZIP
Private m_ZipFileCount As Integer           ' Number of files for ZIP
Private m_DocumentGUIDs As Collection       ' Track unique document GUIDs
Private m_ExportedFileNames As Collection   ' Track exported filenames (for iteration-based lookup)
Private m_InvoiceCount As Long              ' Counter for invoice IDs (GloDr)
Private m_InvalidDocuments As Collection    ' Track expense documents not found (GuiID -> True)
Private m_InvoiceGUIDs As Collection        ' Store GuiIDs for PDF renaming
Private m_InvoiceRechNrs As Collection      ' Store RechNrs parallel to GUIDs (for iteration-based lookup)
Private m_InvMod As Boolean                 ' True wenn Debitoren-Export (RibTab_Abrechnung/Rechnungen)
Private m_OrgPfa As String                  ' Original export path (before subfolder)
Private m_SubNam As String                  ' Subfolder name for export files

' Performance optimization - cached values
Private m_ColumnHeaders As String           ' Cached column headers (built once)
Private m_Q As String                       ' Cached quote character
Private m_Sep As String                     ' Cached separator
Private m_EmptyQuoted As String             ' Cached empty quoted field ("")
Private m_LineBuffer() As String            ' Array-based line buffer for fast concatenation
Private m_LineBufferPos As Integer          ' Current position in line buffer
Private m_LastProgressUpdate As Long        ' Last record number when progress was updated

' Dual progress bar tracking (prbStat1=Detail, prbStat2=Overall)
Private m_TotalPhases As Integer            ' Total number of phases in export
Private m_CurrentPhase As Integer           ' Current phase (1-based)
Private m_PhaseNames() As String            ' Names of each phase

'--------------------------------------------------------------------------------
' PUBLIC API - Entry Points (Replace Legacy Functions)
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' SendDatevEmail - Helper to send email after export
'--------------------------------------------------------------------------------
Private Sub SendDatevEmail(ByVal ExTyp As String, ByVal FilePath As String)
    Dim EmBet As String
    Dim EmTex As String
    Dim DaNam As String
    Dim FiNam As String
    Dim p As Long
    
    If ExTyp = "A" Then
        EmBet = "DATEV Belegarchivierung"
        EmTex = "DATEV 6.0 Belegarchivierung"
    ElseIf ExTyp = "B" Then
        EmBet = "DATEV Belegsatzdaten"
        EmTex = "DATEV 6.0 Belegsatzdaten"
    End If
    
    If Len(EmBet) > 0 Then
        FiNam = FilePath
        p = InStrRev(FiNam, "\")
        If p > 0 Then
            DaNam = Mid$(FiNam, p + 1)
        Else
            DaNam = FiNam
        End If
        
        If GlLog = True Then SLogi "DATEV: SendDatevEmail calling SMaNe for " & DaNam
        SMaNe 0, , , EmTex, EmBet & " - " & GlMan(GlSMa, 1), FiNam
    End If
End Sub

'================================================================================
' DATEV_Expor
'--------------------------------------------------------------------------------
' Purpose:     Export selected bookings to DATEV format (drop-in for S_Expor)
'              Call this from frmBuExp.FWeit() instead of S_Expor
'
' Parameters:  ExTyp       - Export type: "A" = Dokumentenarchivierung (Document-XML)
'                                         "B" = Ledger-Integration (Document-XML + Ledger-XML)
'              EmVer       - Email after export (0=No, 1=Yes)
'              BelEx       - Export documents (Belege) and compress to ZIP
'
' Option A:    PDF Dateien + document.xml (einfaches Archivformat) + CSV
' Option B:    PDF Dateien + document.xml + ledger.xml (strukturierte Belegsatzdaten) + CSV
'
' Usage:       DATEV_Expor "A", EmlVe, BelEx  ' Dokumentenarchivierung
'              DATEV_Expor "B", EmlVe, BelEx  ' Ledger-Integration
'================================================================================
Public Sub DATEV_Expor(ByVal ExTyp As String, Optional ByVal EmVer As Integer = 0, Optional ByVal BelEx As Boolean = False)
On Error GoTo ErrHandler
' Exportiert die markierten Buchungen aus dem ReportControl (wie S_Expor)
' Verwendet Column Captions als Feldnamen im Recordset

    Dim Config As DATEV_ExportConfig
    Dim Result As DATEV_ExportResult
    Dim RST As ADODB.Recordset
    Dim RpCo1 As XtremeReportControl.ReportControl
    Dim RpCo3 As XtremeReportControl.ReportControl
    Dim RpCo4 As XtremeReportControl.ReportControl
    Dim RpCls As XtremeReportControl.ReportColumns
    Dim RpSel As XtremeReportControl.ReportSelectedRows
    Dim RpRow As XtremeReportControl.ReportRow
    Dim RpCol As XtremeReportControl.ReportColumn
    Dim FM As frmMain
    Dim GesPo As Integer
    Dim AktPo As Integer
    Dim IdxNr As Long
    Dim CapSt As String
    Dim TmStr As String
    Dim AkZa1 As Integer
    Dim AkZa2 As Integer
    Dim clFil As clsFile
    Dim ExOrd As String
    Dim DaExt As String
    Dim DaNam As String
    Dim FiNam As String
    Dim DaPfa As String
    Dim DaNaO As String

    Set FM = frmMain
    Set clFil = New clsFile

    ' Exportordner ermitteln
    If GlRDP = True Then
        If clFil.FilDir(GlIPf) = False Then
            ExOrd = GlDpf & "Import\"
        Else
            ExOrd = GlIPf
        End If
    Else
        If clFil.FilDir(GlExO) = False Then
            ExOrd = GlDpf & "Export\"
        Else
            ExOrd = GlExO
        End If
    End If
    If Right$(ExOrd, 1) <> "\" Then
        ExOrd = ExOrd & "\"
    End If

    ' Get ReportControl based on current tab (wie S_Expor)
    Set RpCo1 = FM.repCont1
    Set RpCo3 = FM.repCont3
    Set RpCo4 = FM.repCont4
    Select Case GlBut
    Case RibTab_Abrechnung:
        Set RpCls = RpCo3.Columns
        Set RpSel = RpCo3.SelectedRows
    Case RibTab_Rechnungen:
        Set RpCls = RpCo4.Columns
        Set RpSel = RpCo4.SelectedRows
    Case RibTab_Buchungen:
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
    Case Else
        SPopu "DATEV-Export", "DATEV-Export nur von Buchungen moeglich", IC48_Warning
        Exit Sub
    End Select

    ' Debitoren-Modus setzen (Abrechnung/Rechnungen = Rechnungsexport)
    m_InvMod = (GlBut = RibTab_Abrechnung Or GlBut = RibTab_Rechnungen)
    If GlLog = True Then SLogi "DATEV_Expor: m_InvMod = " & m_InvMod & " (GlBut = " & GlBut & ")"

    GesPo = RpSel.Count
    If GesPo = 0 Then
        SPopu "DATEV-Export", "Keine Buchungen ausgewaehlt", IC48_Warning
        Exit Sub
    End If

    ' Speichern-Dialog ZUERST anzeigen (vor Fortschrittsdialog)
    DaExt = "csv"
    DaNam = "EXTF_DATEV_" & Format$(Now, "YYYYMMDD_HHMM") & ".csv"

    With clFil
        .hwnd = FM.hwnd
        .StaVe = ExOrd
        .DaExt = DaExt
        .DaNam = ExOrd & DaNam
        .DaTit = "Bitte Name und Ordner der Exportdatei angeben"
        .DaStr = "DATEV 4.0 Dateien (*.csv)" & Chr(0) & "*.csv" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        FiNam = .FilSav
    End With

    If FiNam = vbNullString Then
        Set RpCo1 = Nothing
        Set RpCls = Nothing
        Set RpSel = Nothing
        Set clFil = Nothing
        Set FM = Nothing
        Exit Sub
    End If

    If Right$(FiNam, 4) <> "." & DaExt Then
        FiNam = FiNam & "." & DaExt
    End If

    With clFil
        .FilPfa FiNam
        DaPfa = .DaPfa & "\"
        DaNaO = .DaNaO
        If .FilVor(FiNam) = True Then
            .DaLoe = FiNam & vbNullChar
            .FilLoe
        End If
    End With

    If GlBlE = False Then 'DATEV Belegexport
        If LCase(DaPfa) <> LCase(GlExO) Then
            IniSetVal "System", "ExpOrd", LCase(DaPfa)
            GlExO = DaPfa
        End If
    End If

    ' Determine number of phases based on export options
    ' Phase 1: Daten laden, Phase 2: CSV, Phase 3: XML (optional), Phase 4: PDF (optional), Phase 5: ZIP (optional)
    Dim NumPhases As Integer
    NumPhases = 2  ' Minimum: Daten laden + CSV
    If BelEx Then NumPhases = NumPhases + 3  ' XML + PDF + ZIP

    ' Initialize dual progress dialog
    InitProgressWithPhases "DATEV Export", NumPhases
    StartPhase 1, "Lade Daten", GesPo

    ' Recordset mit Column Captions als Feldnamen erstellen (wie S_Expor)
    Set RST = New ADODB.Recordset
    For Each RpCol In RpCls
        CapSt = RpCol.Caption
        If CapSt <> vbNullString Then
            RST.Fields.Append CapSt, adVariant
        Else
            RST.Fields.Append "S" & RpCol.Index, adVariant
        End If
    Next RpCol
    If RST.State = adStateClosed Then
        RST.Open
    End If

    ' Daten aus ReportControl in Recordset kopieren (wie S_Expor)
    AktPo = 0
    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then

            ' Berichtdatum setzen (wie S_Expor)
            ' Debitoren: ID1 (Rechnung-PK) statt Buh_ID0 (Buchungs-PK)
            If m_InvMod Then
                Set RpCol = RpCls.Find(Rec_ID1)
                IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                ' Kein qrySimBuBer fuer Rechnungen - Berichtdatum wird nicht gesetzt
            Else
                Set RpCol = RpCls.Find(Buh_ID0)
                IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                DBCmEx2 "qrySimBuBer", "@IdDat", "@IdxNr", Date, IdxNr
            End If

            ' Neue Zeile anlegen
            RST.AddNew
            For Each RpCol In RpCls
                CapSt = RpCol.Caption
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    If RpRow.Record(RpCol.ItemIndex).Caption = "GuiID" Then
                        TmStr = CreateID("B")
                    Else
                        TmStr = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    ' Sonderzeichen entfernen (wie S_Expor)
                    For AkZa2 = 1 To 10
                        For AkZa1 = 30 To 180
                            If (AkZa1 < 34 Or AkZa1 > 122) Then
                                If AkZa1 <> 32 Then
                                    TmStr = Replace(TmStr, Chr$(AkZa1), vbNullString, 1)
                                End If
                            End If
                        Next AkZa1
                    Next AkZa2
                    If RST.Fields(RpCol.ItemIndex).Type = adBoolean Then
                        RST.Fields(RpCol.ItemIndex).Value = CBool(TmStr)
                    Else
                        If TmStr <> vbNullString Then
                            RST.Fields(RpCol.ItemIndex).Value = TmStr
                        Else
                            RST.Fields(RpCol.ItemIndex).Value = 0
                        End If
                    End If
                End If
            Next RpCol
            RST.Update

            AktPo = AktPo + 1

            ' Update progress every 10 records
            If AktPo Mod 10 = 0 Then
                UpdateProgress AktPo, GesPo, "Lade Datensatz " & AktPo & " von " & GesPo & "..."
                DoEvents
            End If
        End If
    Next RpRow

    ' Complete phase 1 (Daten laden)
    CompletePhase

    If RST.RecordCount = 0 Then
        HideProgressDialog
        SPopu "DATEV-Export", "Keine Buchungsdaten gefunden", IC48_Warning
        RST.Close
        Set RST = Nothing
        Exit Sub
    End If

    RST.MoveFirst

    ' Configure export
    Config = DATEV_GetDefaultConfig()
    ' Derive MandantNr from first record (ReportControl caption "Mandant")
    If HasField(RST, "Mandant") Then
        Config.MandantNr = SafeLongField(RST, "Mandant")
    Else
        Config.MandantNr = 0
    End If
    Config.EmailAfterExport = EmVer
    Config.ExportDocuments = BelEx
    Config.CompressOutput = BelEx

    ' Set export type based on ExTyp parameter
    ' Option A: Dokumentenarchivierung - Document-XML (einfaches Archivformat)
    ' Option B: Ledger-Integration - Document-XML + Ledger-XML (strukturierte Belegsatzdaten)
    Select Case UCase$(ExTyp)
    Case "A"
        ' Option A: Dokumentenarchivierung
        Config.ExportCSV = True
        Config.ExportXML = Config.ExportDocuments
        Config.UseLedgerXML = False
    Case "B"
        ' Option B: Ledger-Integration mit strukturierten Belegsatzdaten
        Config.ExportCSV = True
        Config.ExportXML = Config.ExportDocuments
        Config.UseLedgerXML = True
    Case Else
        ' Default: Option A (Dokumentenarchivierung)
        Config.ExportCSV = True
        Config.ExportXML = Config.ExportDocuments
        Config.UseLedgerXML = False
    End Select

    ' Exportpfad und Dateiname in Config setzen (Dialog wurde bereits angezeigt)
    Config.ExportPath = DaPfa
    Config.ExportFileName = DaNaO

    ' Execute export mit Column Captions
    Result = DATEV_ExportFromReportControl(Config, RST)

    ' Konsistenzpruefung und Belegzaehlung bei Belegexport
    If Config.ExportDocuments And Result.success Then
        If GlLog = True Then SLogi "=== Calling ValidateExportConsistency ==="
        DoEvents
        ValidateExportConsistency Result, Config.ExportPath
        If GlLog = True Then SLogi "=== ValidateExportConsistency returned ==="
        DoEvents
        ' Belegzaehlung nach Typ (Debitoren/Kreditoren)
        If GlLog = True Then SLogi "=== Calling CountDocumentsByType ==="
        DoEvents
        CountDocumentsByType RST, Result, Config.ExportPath
        If GlLog = True Then SLogi "=== CountDocumentsByType returned ==="
        DoEvents
    End If

    ' Show result
    Dim Tit1 As String, Mld1 As String
    Dim DeleteResult As Long
    Tit1 = "DATEV Export"
    If Result.success Then
        ' Zusammenfassung formatieren (mindestens 3 Zeichen)
        If Config.ExportDocuments Then
            ' Mit Belegexport: Buchungen + Debitorenbelege + Kreditorenbelege
            Mld1 = Format$(Result.RecordCount, "000") & " Buchungen exportiert" & vbCrLf & _
                   Format$(Result.DebitDocCount, "000") & " Debitorenbelege exportiert" & vbCrLf & _
                   Format$(Result.KreditDocCount, "000") & " Kreditorenbelege exportiert"
            ' Konsistenz-Warnung anhaengen wenn noetig
            If Not Result.ConsistencyOK Then
                Mld1 = Mld1 & vbCrLf & vbCrLf & Result.ConsistencyMessage
            End If
        Else
            ' Ohne Belegexport: nur Buchungen
            Mld1 = Format$(Result.RecordCount, "000") & " Buchungen exportiert"
        End If
        If GlLog = True Then SLogi "=== SPopu Ergebnis anzeigen ==="
        DoEvents
        SPopu Tit1, Mld1, IC48_Information
        If GlLog = True Then SLogi "=== SPopu beendet ==="
        DoEvents

        ' Email senden wenn angefordert
        If EmVer = 1 Then
            Dim FinalPath As String
            If Len(Result.ZipFilePath) > 0 Then
                FinalPath = Result.ZipFilePath
            Else
                FinalPath = Result.CSVFilePath
            End If
            If Len(FinalPath) > 0 Then
                SendDatevEmail ExTyp, FinalPath
            End If
        End If
    Else
        Mld1 = "DATEV-Export fehlgeschlagen:" & vbCrLf & Result.ErrorMessage
        SPopu Tit1, Mld1, IC48_Information
    End If

    ' Cleanup
    If RST.State = adStateOpen Then RST.Close
    Set RST = Nothing
    Set RpCo1 = Nothing
    Set RpCo3 = Nothing
    Set RpCo4 = Nothing
    Set RpCls = Nothing
    Set RpSel = Nothing
    Set clFil = Nothing
    Set FM = Nothing

    Exit Sub

ErrHandler:
    HideProgressDialog
    If GlLog = True Then SLogi "=== DATEV_Expor ERROR ==="
    DoEvents
    If GlLog = True Then SLogi "  Err.Number: " & Err.Number
    If GlLog = True Then SLogi "  Err.Description: " & Err.Description
    If GlLog = True Then SLogi "  Err.Source: " & Err.Source
    DoEvents
    SPopu "DATEV_Expor " & Err.Number, Err.Description, IC48_Warning
    If Not RST Is Nothing Then
        If RST.State = adStateOpen Then RST.Close
        Set RST = Nothing
    End If
    Set clFil = Nothing
End Sub

'================================================================================
' DATEV_BuEx
'--------------------------------------------------------------------------------
' Purpose:     Export bookings from database by SQL criteria to DATEV format
'              Replaces: basData.S_BuEx() for DATEV exports
'              This function queries the database directly (for period-based export)
'
' Parameters:  ExTyp       - Export type: "A" = Dokumentenarchivierung (Document-XML)
'                                         "B" = Ledger-Integration (Document-XML + Ledger-XML)
'              EmVer       - Email after export (0=No, 1=Yes)
'              Krite       - SQL WHERE criteria (e.g., "(Datum >= #01.01.2024#)")
'                            Already contains mandant filter (IDT) via caller
'              BelEx       - Export documents (Belege) and compress to ZIP
'
' Option A:    PDF Dateien + document.xml (einfaches Archivformat) + CSV
' Option B:    PDF Dateien + document.xml + ledger.xml (strukturierte Belegsatzdaten) + CSV
'
' Usage:       DATEV_BuEx "A", EmlVe, Krite, BelEx  ' Dokumentenarchivierung
'              DATEV_BuEx "B", EmlVe, Krite, BelEx  ' Ledger-Integration
'================================================================================
Public Sub DATEV_BuEx(ByVal ExTyp As String, ByVal EmVer As Integer, ByVal Krite As String, Optional ByVal BelEx As Boolean = False)
On Error GoTo ErrHandler

    Dim Config As DATEV_ExportConfig
    Dim Result As DATEV_ExportResult
    Dim RST As ADODB.Recordset
    Dim SQL1 As String
    Dim GesPo As Long
    Dim clFil As clsFile
    Dim FM As frmMain
    Dim ExOrd As String
    Dim DaExt As String
    Dim DaNam As String
    Dim FiNam As String
    Dim DaPfa As String
    Dim DaNaO As String

    Set FM = frmMain
    Set clFil = New clsFile

    ' Exportordner ermitteln
    If GlRDP = True Then
        If clFil.FilDir(GlIPf) = False Then
            ExOrd = GlDpf & "Import\"
        Else
            ExOrd = GlIPf
        End If
    Else
        If clFil.FilDir(GlExO) = False Then
            ExOrd = GlDpf & "Export\"
        Else
            ExOrd = GlExO
        End If
    End If
    If Right$(ExOrd, 1) <> "\" Then
        ExOrd = ExOrd & "\"
    End If

    ' Build SQL query based on current tab
    Select Case GlBut
    Case RibTab_Abrechnung:
        If GlTyp < 2 Then
            SQL1 = "SELECT * FROM dbo.qrySimReSu WHERE " & Krite & " ORDER BY ID1, Datum"
        Else
            SQL1 = "SELECT * FROM qrySimReSu WHERE " & Krite & " ORDER BY [ID1], [Datum];"
        End If
    Case RibTab_Rechnungen:
        If GlTyp < 2 Then
            SQL1 = "SELECT * FROM dbo.qrySimReSu WHERE " & Krite & " ORDER BY ID1, Datum"
        Else
            SQL1 = "SELECT * FROM qrySimReSu WHERE " & Krite & " ORDER BY [ID1], [Datum];"
        End If
    Case RibTab_Mahnwesen:
        If GlTyp < 2 Then
            SQL1 = "SELECT * FROM dbo.qrySimOPSu WHERE " & Krite & " ORDER BY IDR, Datum"
        Else
            SQL1 = "SELECT * FROM qrySimOPSu WHERE " & Krite & " ORDER BY [IDR], [Datum];"
        End If
    Case RibTab_Buchungen:
        If GlTyp < 2 Then
            SQL1 = "SELECT * FROM dbo.qrySimBuSu WHERE " & Krite & " ORDER BY Datum"
        Else
            SQL1 = "SELECT * FROM qrySimBuSu WHERE " & Krite & " ORDER BY [Datum];"
        End If
    Case Else
        SPopu "DATEV-Export", "DATEV-Export nur von Buchungen moeglich", IC48_Warning
        Exit Sub
    End Select

    ' Debitoren-Modus setzen (Abrechnung/Rechnungen = Rechnungsexport)
    m_InvMod = (GlBut = RibTab_Abrechnung Or GlBut = RibTab_Rechnungen)
    If GlLog = True Then SLogi "DATEV_BuEx: m_InvMod = " & m_InvMod & " (GlBut = " & GlBut & ")"

    ' Speichern-Dialog ZUERST anzeigen (vor Datenbankabfrage)
    DaExt = "csv"
    DaNam = "EXTF_DATEV_" & Format$(Now, "YYYYMMDD_HHMM") & ".csv"

    With clFil
        .hwnd = FM.hwnd
        .StaVe = ExOrd
        .DaExt = DaExt
        .DaNam = ExOrd & DaNam
        .DaTit = "Bitte Name und Ordner der Exportdatei angeben"
        .DaStr = "DATEV 4.0 Dateien (*.csv)" & Chr(0) & "*.csv" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        FiNam = .FilSav
    End With

    If FiNam = vbNullString Then
        Set clFil = Nothing
        Set FM = Nothing
        Exit Sub
    End If

    If Right$(FiNam, 4) <> "." & DaExt Then
        FiNam = FiNam & "." & DaExt
    End If

    With clFil
        .FilPfa FiNam
        DaPfa = .DaPfa & "\"
        DaNaO = .DaNaO
        If .FilVor(FiNam) = True Then
            .DaLoe = FiNam & vbNullChar
            .FilLoe
        End If
    End With

    If GlBlE = False Then 'DATEV Belegexport
        If LCase(DaPfa) <> LCase(GlExO) Then
            IniSetVal "System", "ExpOrd", LCase(DaPfa)
            GlExO = DaPfa
        End If
    End If

    ' Execute database query
    Set RST = New ADODB.Recordset
    With RST
        .CursorLocation = adUseClient
        .Source = SQL1
        .ActiveConnection = DB1
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open Options:=adCmdText
    End With
    GesPo = RST.RecordCount

    If GesPo = 0 Then
        SPopu "DATEV-Export", "Keine Buchungen im angegebenen Zeitraum gefunden", IC48_Warning
        RST.Close
        Set RST = Nothing
        Exit Sub
    End If

    RST.MoveFirst

    ' Configure export
    Config = DATEV_GetDefaultConfig()
    ' Derive MandantNr from first record (already filtered via Krite)
    If HasField(RST, "IDT") Then
        Config.MandantNr = SafeLong(RST.Fields("IDT").Value)
    Else
        Config.MandantNr = 0
    End If
    Config.EmailAfterExport = EmVer
    Config.ExportDocuments = BelEx
    Config.CompressOutput = BelEx

    ' Set export type based on ExTyp parameter
    ' Option A: Dokumentenarchivierung - Document-XML (einfaches Archivformat)
    ' Option B: Ledger-Integration - Document-XML + Ledger-XML (strukturierte Belegsatzdaten)
    Select Case UCase$(ExTyp)
    Case "A"
        ' Option A: Dokumentenarchivierung
        Config.ExportCSV = True
        Config.ExportXML = Config.ExportDocuments  ' XML nur wenn ExportDocuments = True
        Config.UseLedgerXML = False
    Case "B"
        ' Option B: Ledger-Integration mit strukturierten Belegsatzdaten
        Config.ExportCSV = True
        Config.ExportXML = Config.ExportDocuments  ' XML nur wenn ExportDocuments = True
        Config.UseLedgerXML = True
    Case Else
        ' Default: Option A (Dokumentenarchivierung)
        Config.ExportCSV = True
        Config.ExportXML = Config.ExportDocuments
        Config.UseLedgerXML = False
    End Select

    ' Exportpfad und Dateiname in Config setzen (Dialog wurde bereits angezeigt)
    Config.ExportPath = DaPfa
    Config.ExportFileName = DaNaO

    ' Execute export
    Result = DATEV_ExportSelected(Config, RST)

    ' Konsistenzpruefung und Belegzaehlung bei Belegexport
    If Config.ExportDocuments And Result.success Then
        If GlLog = True Then SLogi "=== Calling ValidateExportConsistency ==="
        DoEvents
        ValidateExportConsistency Result, Config.ExportPath
        If GlLog = True Then SLogi "=== ValidateExportConsistency returned ==="
        DoEvents
        ' Belegzaehlung nach Typ (Debitoren/Kreditoren)
        If GlLog = True Then SLogi "=== Calling CountDocumentsByType ==="
        DoEvents
        CountDocumentsByType RST, Result, Config.ExportPath
        If GlLog = True Then SLogi "=== CountDocumentsByType returned ==="
        DoEvents
    End If

    ' Show result
    Dim Tit1 As String, Mld1 As String
    Dim DeleteResult As Long
    Tit1 = "DATEV Export"
    If Result.success Then
        ' Zusammenfassung formatieren (mindestens 3 Zeichen)
        If Config.ExportDocuments Then
            ' Mit Belegexport: Buchungen + Debitorenbelege + Kreditorenbelege
            Mld1 = Format$(Result.RecordCount, "000") & " Buchungen exportiert" & vbCrLf & _
                   Format$(Result.DebitDocCount, "000") & " Debitorenbelege exportiert" & vbCrLf & _
                   Format$(Result.KreditDocCount, "000") & " Kreditorenbelege exportiert"
            ' Konsistenz-Warnung anhaengen wenn noetig
            If Not Result.ConsistencyOK Then
                Mld1 = Mld1 & vbCrLf & vbCrLf & Result.ConsistencyMessage
            End If
        Else
            ' Ohne Belegexport: nur Buchungen
            Mld1 = Format$(Result.RecordCount, "000") & " Buchungen exportiert"
        End If
        If GlLog = True Then SLogi "=== SPopu Ergebnis anzeigen ==="
        DoEvents
        SPopu Tit1, Mld1, IC48_Information
        If GlLog = True Then SLogi "=== SPopu beendet ==="
        DoEvents

        ' Email senden wenn angefordert
        If EmVer = 1 Then
            Dim FinalPath As String
            If Len(Result.ZipFilePath) > 0 Then
                FinalPath = Result.ZipFilePath
            Else
                FinalPath = Result.CSVFilePath
            End If
            If Len(FinalPath) > 0 Then
                SendDatevEmail ExTyp, FinalPath
            End If
        End If
    Else
        Mld1 = "DATEV-Export fehlgeschlagen:" & vbCrLf & Result.ErrorMessage
        SPopu Tit1, Mld1, IC48_Information
    End If

    ' Cleanup
    If RST.State = adStateOpen Then RST.Close
    Set RST = Nothing
    Set clFil = Nothing
    Set FM = Nothing

    Exit Sub

ErrHandler:
    If GlLog = True Then SLogi "=== DATEV_BuEx ERROR ==="
    DoEvents
    If GlLog = True Then SLogi "  Err.Number: " & Err.Number
    If GlLog = True Then SLogi "  Err.Description: " & Err.Description
    If GlLog = True Then SLogi "  Err.Source: " & Err.Source
    DoEvents
    SPopu "DATEV_BuEx " & Err.Number, Err.Description, IC48_Warning
    If Not RST Is Nothing Then
        If RST.State = adStateOpen Then RST.Close
        Set RST = Nothing
    End If
    Set clFil = Nothing
    Set FM = Nothing
End Sub

'================================================================================
' DATEV_ExportFromReportControl
'--------------------------------------------------------------------------------
' Purpose:     Export bookings from ReportControl to DATEV format
'              Uses Column Captions as field names (like S_Expor in basDaMa.bas)
'
' Parameters:  Config      - Export configuration
'              RST         - Recordset with Column Captions as field names
'                           (Mandant, Mitarbeiter, Sachkonto, etc.)
'
' Returns:     DATEV_ExportResult with status and file paths
'
' Note:        This function mirrors DATEV_ExportSelected but uses Column Captions
'              Field mapping (Column Caption -> DB Field):
'              - "Mandant" -> "IDT"
'              - "Mitarbeiter" -> "IDM"
'              - "Sachkonto" -> "IDK"
'              - "Gegenkonto" -> "IDG"
'              - "Sachkontenbezeichnung" -> "IDKurz"
'              - "Belegzeichen" -> "RechNr"
'              - "Nummer" -> "Beleg"
'              - "Buchungstext" -> "Buchtext"
'================================================================================
Public Function DATEV_ExportFromReportControl(ByRef Config As DATEV_ExportConfig, _
                                              ByRef RST As ADODB.Recordset) As DATEV_ExportResult
On Error GoTo ErrHandler

    Dim Result As DATEV_ExportResult
    Dim CSVContent As String
    Dim XMLContent As String
    Dim DataLine As String
    Dim RecordCount As Long
    Dim ExportPath As String
    Dim CSVFileName As String
    Dim XMLFileName As String
    Dim DaSt1 As String
    Dim DaSt2 As String
    Dim DaSt3 As String
    Dim DaSt4 As String
    Dim TmpDat As String
    Dim BerSt As String
    Dim MaNam As String
    Dim AnzKo As Integer

    ' Initialize result
    Result.success = False
    Result.RecordCount = 0
    Result.DocumentCount = 0
    Result.ErrorMessage = vbNullString
    Result.ErrorCode = 0

    ' Initialize module state
    m_Cancelled = False
    m_Config = Config

    ' Validate configuration
    If Not ValidateConfig(Config) Then
        Result.ErrorMessage = "Ungueltige Exportkonfiguration"
        Result.ErrorCode = 1001
        DATEV_ExportFromReportControl = Result
        Exit Function
    End If

    ' Validate recordset
    If RST Is Nothing Then
        Result.ErrorMessage = "Keine Daten zum Exportieren vorhanden"
        Result.ErrorCode = 1002
        DATEV_ExportFromReportControl = Result
        Exit Function
    End If

    If RST.EOF And RST.BOF Then
        Result.ErrorMessage = "Keine Datensaetze im Ergebnis"
        Result.ErrorCode = 1003
        DATEV_ExportFromReportControl = Result
        Exit Function
    End If

    ' Phase tracking for dual progress (phases continue from DATEV_Expor)
    ' Phase 1 was "Daten laden" in DATEV_Expor
    Dim PhaseNum As Integer
    PhaseNum = 1  ' Start after phase 1

    ' Initialize file operations
    Set m_clFil = New clsFile

    ' Initialize collections and arrays
    ReDim GloDr(0)
    m_InvoiceCount = 0
    Set m_DocumentGUIDs = New Collection
    Set m_ExportedFileNames = New Collection
    Set m_InvoiceGUIDs = New Collection
    Set m_InvoiceRechNrs = New Collection
    Set m_InvalidDocuments = New Collection  ' Initialize early to avoid Nothing errors
    ReDim m_ZipFiles(0)
    m_ZipFileCount = 0

    ' Ensure export directory exists
    If Not EnsureExportDirectory(Config.ExportPath) Then
        Result.ErrorMessage = "Exportverzeichnis kann nicht erstellt werden: " & Config.ExportPath
        Result.ErrorCode = 1004
        DATEV_ExportFromReportControl = Result
        Exit Function
    End If

    ' Setup subfolder for export if compression is requested
    If Config.CompressOutput Then
        If Not SetSubFo(Config) Then
            Result.ErrorMessage = "Export-Unterordner konnte nicht erstellt werden"
            Result.ErrorCode = 1005
            DATEV_ExportFromReportControl = Result
            Exit Function
        End If
    Else
        ' No compression - save original path for compatibility
        m_OrgPfa = Config.ExportPath
        m_SubNam = vbNullString
    End If

    ' Get date range from recordset
    RST.MoveFirst
    If HasField(RST, "Datum") Then
        If IsDate(RST.Fields("Datum").Value) Then
            DaSt1 = Format$(RST.Fields("Datum").Value, "YYYYMMDD")
            DaSt3 = DatePart("yyyy", CDate(RST.Fields("Datum").Value), vbMonday) & "0101"
        Else
            DaSt1 = Format$(Date, "YYYYMMDD")
            DaSt3 = DatePart("yyyy", Date, vbMonday) & "0101"
        End If
    Else
        DaSt1 = Format$(Date, "YYYYMMDD")
        DaSt3 = DatePart("yyyy", Date, vbMonday) & "0101"
    End If
    RST.MoveLast
    If HasField(RST, "Datum") And IsDate(RST.Fields("Datum").Value) Then
        DaSt2 = Format$(RST.Fields("Datum").Value, "YYYYMMDD")
    Else
        DaSt2 = Format$(Date, "YYYYMMDD")
    End If
    RST.MoveFirst

    ' Ensure DatumVon <= DatumBis (swap if recordset is sorted descending)
    If DaSt1 > DaSt2 Then
        TmpDat = DaSt1
        DaSt1 = DaSt2
        DaSt2 = TmpDat
    End If

    ' DATEV header fields - ErzeugtAm ohne Leerzeichen (YYYYMMDDHHMMSS)
    DaSt4 = Format$(Now, "YYYYMMDD") & Format$(Now, "HHMMSS")
    BerSt = Left$(Format$(GlDvB, "00000"), 5)
    MaNam = Left$(Format$(GlDvM, "00000"), 5)
    If GldKt = True Then
        AnzKo = 4
    Else
        AnzKo = 6
    End If

    ' Build CSV content
    If Config.ExportCSV Then
        ' Phase 2: CSV Buchungsstapel
        PhaseNum = PhaseNum + 1
        StartPhase PhaseNum, "CSV Buchungsstapel", RST.RecordCount

        ' Build DATEV header line
        CSVContent = BuildDATEVHeaderLine(DaSt1, DaSt2, DaSt3, DaSt4, BerSt, MaNam, AnzKo)
        CSVContent = CSVContent & vbCrLf & BuildColumnHeaderLine() & vbCrLf

        ' Process each record with Column Captions
        RecordCount = 0
        Do While Not RST.EOF
            DataLine = BuildCSVDataLineFromReportControl(RST, Config)
            If Len(DataLine) > 0 Then
                CSVContent = CSVContent & DataLine & vbCrLf
                RecordCount = RecordCount + 1
            End If

            ' Collect invoice IDs for PDF generation (Einnahmen)
            If Config.ExportDocuments Then
                CollectInvoiceIDFromReportControl RST
            End If

            ' Update progress (dual mode)
            UpdateDualProgress RecordCount, RST.RecordCount, "Buchung " & RecordCount & " von " & RST.RecordCount

            RST.MoveNext
            DoEvents
        Loop

        ' Complete phase 2
        CompletePhase

        ' Write CSV file in ANSI/Windows-1252 encoding (DATEV EXTF requirement)
        ' Benutzer-gewaehlten Dateinamen verwenden falls vorhanden
        If Len(Config.ExportFileName) > 0 Then
            CSVFileName = Config.ExportPath & Config.ExportFileName & ".csv"
        Else
            CSVFileName = Config.ExportPath & "EXTF_Buchungsstapel_" & Format$(Date, "YYYYMMDD") & "_" & Format$(Time, "HHMMSS") & ".csv"
        End If
        If WriteCSVFileAnsi(CSVFileName, CSVContent) Then
            Result.CSVFilePath = CSVFileName
            ' Add to ZIP list
            AddToZipListRC CSVFileName
        Else
            Result.ErrorMessage = "CSV-Datei konnte nicht geschrieben werden"
            Result.ErrorCode = 2001
            GoTo Cleanup
        End If
    End If

    ' Generate XML if requested
    If Config.ExportXML And Config.ExportDocuments Then
        ' Phase 3: XML Belegverknuepfung
        PhaseNum = PhaseNum + 1
        StartPhase PhaseNum, "XML Belegverknuepfung", RST.RecordCount

        RST.MoveFirst
        XMLFileName = Config.ExportPath & "document.xml"
        GenerateXMLFromReportControl RST, XMLFileName, Config
        If m_clFil.FilVor(XMLFileName) Then
            Result.XMLFilePath = XMLFileName
            AddToZipListRC XMLFileName
        End If

        CompletePhase

        ' Option B: Generate ledger.xml for structured booking data import
        If Config.UseLedgerXML Then
            Dim LedgerXMLPath As String
            If GlLog = True Then SLogi "DATEV: Option B aktiv, starte ledger.xml Generierung"
            DoEvents

            ' Reset recordset position for ledger.xml generation
            On Error Resume Next
            RST.MoveFirst
            If Err.Number <> 0 Then
                If GlLog = True Then SLogi "DATEV: RST.MoveFirst Fehler: " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler

            LedgerXMLPath = GenerateLedgerXMLFromRC(RST, Config)
            If Len(LedgerXMLPath) = 0 Then
                If GlLog = True Then SLogi "DATEV: ledger.xml nicht erstellt!"
            Else
                If GlLog = True Then SLogi "DATEV: ledger.xml erstellt: " & LedgerXMLPath
                ' Note: GenerateLedgerXMLFromRC adds to ZIP list internally via AddToZipList
            End If
        Else
            If GlLog = True Then SLogi "DATEV: UseLedgerXML = False, keine ledger.xml"
        End If
    End If

    ' Generate PDF documents if requested
    If Config.ExportDocuments And m_InvoiceCount > 0 Then
        ' Phase 4: PDF Belege
        PhaseNum = PhaseNum + 1
        StartPhase PhaseNum, "PDF Belege", m_InvoiceCount

        GenerateInvoicePDFs Config

        CompletePhase
    End If

    ' Create ZIP archive if requested
    If Config.CompressOutput And m_ZipFileCount > 0 Then
        ' Phase 5: ZIP Archiv
        PhaseNum = PhaseNum + 1
        StartPhase PhaseNum, "ZIP Archiv", m_ZipFileCount

        frmStatus.Hide
        Result.ZipFilePath = CreateZIPArchiveFromReportControl(Config)

        CompletePhase
    End If

    ' Success
    Result.success = True
    Result.RecordCount = RecordCount

Cleanup:
    ' Hide progress dialog
    HideProgressDialog

    ' Cleanup
    Set m_clFil = Nothing
    Set m_DocumentGUIDs = Nothing
    Set m_ExportedFileNames = Nothing
    Set m_InvoiceGUIDs = Nothing
    Set m_InvoiceRechNrs = Nothing

    DATEV_ExportFromReportControl = Result
    Exit Function

ErrHandler:
    Result.success = False
    Result.ErrorMessage = "Exportfehler: " & Err.Description
    Result.ErrorCode = Err.Number
    LogError "DATEV_ExportFromReportControl", Err.Number, Err.Description
    Resume Cleanup
End Function

'================================================================================
' BuildCSVDataLineFromReportControl
'--------------------------------------------------------------------------------
' Purpose:     Build a single CSV data line using Column Captions
'              (Mandant, Mitarbeiter, Sachkonto, etc.)
'
' Parameters:  RST         - Recordset with Column Captions as field names
'              Config      - Export configuration
'
' Returns:     Formatted CSV line string or empty string if invalid
'================================================================================
Private Function BuildCSVDataLineFromReportControl(ByRef RST As ADODB.Recordset, _
                                                   ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler

    Dim ExSep As String
    Dim BuDat As Date
    Dim IdxNr As Long
    Dim BuTyp As Integer
    Dim GeKto As Integer
    Dim ManNr As Long
    Dim MitNr As Long
    Dim PatNr As Long
    Dim KtoSo As Long
    Dim KtoHa As Long
    Dim Steue As Single
    Dim Storn As Boolean
    Dim BLock As Boolean
    Dim GesBe As String
    Dim Kennz As String
    Dim StSch As String
    Dim ReStr As String
    Dim BelFe As String
    Dim BuStr As String
    Dim PaNum As String
    Dim Koste As String
    Dim BuGui As String
    Dim BeGui As String
    Dim DaNam As String
    Dim Beleg As String
    Dim Komme As String
    Dim KSoSt As String
    Dim KHaSt As String
    Dim TmpSt As String
    Dim BeVor As Boolean
    Dim AktZa As Integer
    Dim PidNr As Long
    Dim DebNr As Long

    ExSep = Chr$(59)

    ' Extract date (required)
    If Not HasField(RST, "Datum") Then
        BuildCSVDataLineFromReportControl = vbNullString
        Exit Function
    End If
    If IsNull(RST.Fields("Datum").Value) Or Not IsDate(RST.Fields("Datum").Value) Then
        BuildCSVDataLineFromReportControl = vbNullString
        Exit Function
    End If
    BuDat = CDate(RST.Fields("Datum").Value)

    ' Extract basic fields using Column Captions
    ' Debitoren-Modus: andere Caption-Namen und Feldzuordnung
    If m_InvMod Then
        ' Caption "ID1" = Rechnung-PK, "Mandant" = IDP (Patient-Ref)
        IdxNr = SafeLongField(RST, "ID1")
        BuTyp = 2  ' Immer Einnahme
        GeKto = 0
        ManNr = SafeLongField(RST, "Mandant")
        MitNr = 0  ' Kein Mitarbeiter-Caption im Rechnungs-RC
    Else
        IdxNr = SafeLongField(RST, "ID0")
        BuTyp = SafeIntField(RST, "IDA")
        GeKto = SafeIntField(RST, "IDB")
        ' Column Captions: "Mandant" und "Mitarbeiter" statt IDT/IDM
        ManNr = SafeLongField(RST, "Mandant")
        MitNr = SafeLongField(RST, "Mitarbeiter")
    End If

    Steue = SafeSingleField(RST, "Steuer")
    Storn = SafeBoolField(RST, "Storniert")
    BLock = SafeBoolField(RST, "Lock")
    If m_InvMod Then
        ' Rechnungen (manuelle Auswahl): Rec_ID0 enthaelt die Patientennummer
        PatNr = SafeLongField(RST, "ID0")
        If PatNr = 0 Then PatNr = SafeLongField(RST, "Mandant")
    Else
        ' Buchungen (manuelle Auswahl): Buh_IDP enthaelt die Patientennummer
        PatNr = SafeLongField(RST, "IDP")
        If PatNr = 0 Then PatNr = SafeLongField(RST, "Mandant")
        If PatNr = 0 Then PatNr = SafeLongField(RST, "ID0")
    End If

    ' Determine amount (Einnahme/Ausgabe)
    ' Debitoren: "Betrag" Caption statt Einnahme/Ausgabe
    If m_InvMod Then
        Dim TmpBet As Variant
        TmpBet = RST.Fields("Betrag").Value
        If Not IsNull(TmpBet) And IsNumeric(TmpBet) And CDbl(TmpBet) > 0 Then
            GesBe = FormatAmountGermanOptimized(Abs(CCur(TmpBet)))
        Else
            GesBe = vbNullString
        End If
    Else
        GesBe = DetermineAmountFromReportControl(RST, BuTyp)
    End If
    If GesBe = "0" Or GesBe = vbNullString Then
        BuildCSVDataLineFromReportControl = vbNullString
        Exit Function
    End If

    ' Determine Soll/Haben
    ' Debitoren: immer Einnahme (S ohne Tausch)
    If m_InvMod Then
        Kennz = IIf(Config.SwapDebitCredit, "S", "H")
    Else
        Kennz = DetermineSollHabenFromReportControl(RST, BuTyp, Config.SwapDebitCredit)
    End If

    ' Get accounts
    ' Debitoren: Konto = Geldkonto (Bank), Gegenkonto = Erloskonto
    If m_InvMod Then
        KtoSo = S_DaKoK(GlGkB)
        KtoHa = GlSE2
        If GlLog = True Then SLogi "DATEV_Expor: InvMod KtoSo=" & KtoSo & " (GlGkB=" & GlGkB & ") KtoHa=" & KtoHa
    ElseIf GlBuc = True Then
        ' Column Captions: "Sachkonto" und "Gegenkonto"
        KtoSo = S_DaKoZ(SafeLongField(RST, "Sachkonto"))
        KtoHa = S_DaKoK(GeKto)
    Else
        If HasField(RST, "Sollkonto") Then
            KtoSo = S_DaKoZ(SafeLongField(RST, "Sollkonto"))
            KtoHa = S_DaKoZ(SafeLongField(RST, "Habenkonto"))
        Else
            KtoSo = S_DaKoZ(SafeLongField(RST, "Sachkonto"))
            KtoHa = S_DaKoK(GeKto)
        End If
    End If

    ' Format account numbers
    KSoSt = S_DaExF(KtoSo)
    KHaSt = S_DaExF(KtoHa)

    ' Determine tax key
    StSch = GetTaxKey(Steue, Kennz)

    ' Get document info
    ' Debitoren: kein Datei-Caption, generiere Dateiname aus "Rechnung"
    Dim InvStr As String
    If m_InvMod Then
        DaNam = vbNullString
        InvStr = SafeStringField(RST, "Rechnung")
        If Len(InvStr) > 0 Then
            DaNam = "Rechnung_Beleg_" & InvStr & ".pdf"
        End If
    Else
        DaNam = SafeStringField(RST, "Datei")
        If Len(DaNam) > 46 Then
            DaNam = Left$(DaNam, 42) & Right$(DaNam, 4)
        End If

        ' Default filename for revenue without document
        If Len(DaNam) = 0 And BuTyp = 2 Then
            ' Column Caption: "Belegzeichen" statt "RechNr"
            InvStr = SafeStringField(RST, "Belegzeichen")
            If Len(InvStr) > 0 Then
                DaNam = "Rechnung_Beleg_" & InvStr & ".pdf"
            End If
        End If
    End If

    ' Get GUID for Beleglink
    BuGui = SafeStringField(RST, "GuiID")
    BeVor = False
    If Len(BuGui) > 0 And Len(DaNam) > 0 Then
        BeGui = DATEV_FormatGUIDForXML(BuGui)
        If Not IsDocumentAlreadyExported(DaNam) Then
            TrackExportedDocument DaNam, BuGui
        Else
            BeVor = True
        End If
    End If

    ' Get text fields
    ' Debitoren: "Rechnung" Caption statt "Belegzeichen"
    If m_InvMod Then
        ReStr = SafeStringField(RST, "Rechnung")
        If Len(ReStr) > 12 Then ReStr = Left$(ReStr, 12)
        ' Beleginfo Art 2: RechNr als Kennung
        Beleg = SafeStringField(RST, "Rechnung")
        If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)
    Else
        ' Column Caption: "Belegzeichen" statt "RechNr"
        ReStr = SafeStringField(RST, "Belegzeichen")
        If Len(ReStr) > 12 Then ReStr = Left$(ReStr, 12)

        ' Column Caption: "Nummer" statt "Beleg"
        ' Beleginfo - Art 2: max 20 Zeichen laut DATEV-Spezifikation
        Beleg = SafeStringField(RST, "Nummer")
        If Len(Beleg) > 0 And IsNumeric(Beleg) Then
            Beleg = Format$(CLng(Beleg), "00000000")
        End If
        If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)
    End If

    ' DATEV-Regel: Beleginfo - Art 2 und Inhalt 2 muessen beide gefuellt oder beide leer sein
    ' DaNam = Inhalt 2 (BEDI-Dateiname), Beleg = Art 2
    If Len(DaNam) = 0 Then
        Beleg = vbNullString  ' Wenn kein Dokument, auch Art 2 leer
    ElseIf Len(Beleg) = 0 Then
        Beleg = Format$(IdxNr, "00000000")  ' Fallback: ID als Art 2 wenn Dokument vorhanden
        If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)
    End If

    Komme = SafeStringField(RST, "Kommentar")
    If Len(Komme) > 0 And IsNumeric(Komme) Then
        Komme = Format$(CLng(Komme), "00000000")
    End If

    ' Stornierung Prefix
    If Storn = True Then
        BelFe = "[STORNIERT] "
    Else
        BelFe = Space$(12)
    End If

    ' Buchungstext
    ' Debitoren: Patient + Rechnung statt Buchungstext-Caption
    If m_InvMod Then
        Dim PaTxt As String
        PaTxt = SafeStringField(RST, "Patient")
        If Len(PaTxt) > 0 And Len(InvStr) > 0 Then
            BuStr = PaTxt & " Rech." & InvStr
        ElseIf Len(PaTxt) > 0 Then
            BuStr = PaTxt
        ElseIf Len(InvStr) > 0 Then
            BuStr = "Rechnung " & InvStr
        Else
            BuStr = "Rechnungsexport"
        End If
    Else
        ' Column Caption: "Buchungstext" statt "Buchtext"
        BuStr = SafeStringField(RST, "Buchungstext")
    End If
    If Len(BuStr) > 60 Then BuStr = Left$(BuStr, 60)
    BuStr = Replace(BuStr, ExSep, vbNullString)

    ' Patient number
    PaNum = vbNullString
    If Config.IncludePatientNumber Then
        PidNr = PatNr
        If PidNr <= 0 Then
            Dim TmSt2 As String
            On Error Resume Next
            TmSt2 = S_AdIdx(IdxNr, "IDP")
            On Error GoTo ErrHandler
            If Len(TmSt2) > 0 And IsNumeric(TmSt2) Then
                PidNr = CLng(TmSt2)
            End If
        End If

        If PidNr > 0 Then
            If Config.FourDigitAccounts Then
                PaNum = Format$(PidNr, "00000")
            Else
                PaNum = Format$(PidNr, "0000000")
            End If
        End If
    Else
        If Config.FourDigitAccounts Then
            PaNum = "6" & Format(BuDat, "yy") & Format(BuDat, "mm")
        Else
            PaNum = "6" & Format(BuDat, "yyyy") & Format(BuDat, "mm")
        End If
    End If

    ' DATEV-konforme Debitorennummer (Patientennr + Basis)
    DebNr = 0
    If PidNr > 0 Then
        If Config.FourDigitAccounts Then
            DebNr = 10000 + PidNr
        Else
            DebNr = 1000000 + PidNr
        End If
    End If

    ' GlDeE: Replace account (Konto) with debtor number (invoices only)
    If Config.ReplaceAccountWithDebtor And m_InvMod Then
        If DebNr > 0 Then KSoSt = CStr(DebNr)
    End If

    ' Kostenstelle
    Koste = vbNullString
    For AktZa = 1 To UBound(GlThe)
        If ManNr = GlThe(AktZa, 0) Then
            If GlThe(AktZa, 47) <> vbNullString Then
                If Len(GlThe(AktZa, 47)) > 8 Then
                    Koste = Left$(GlThe(AktZa, 47), 8)
                Else
                    Koste = GlThe(AktZa, 47)
                End If
            End If
            Exit For
        End If
    Next AktZa

    ' Build CSV line (DATEV format)
    TmpSt = vbNullString
    TmpSt = TmpSt & GesBe & ExSep  'Umsatz
    TmpSt = TmpSt & Chr$(34) & Kennz & Chr$(34) & ExSep   'Soll/Haben
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep  'WKZ Umsatz
    TmpSt = TmpSt & vbNullString & ExSep 'Kurs
    TmpSt = TmpSt & vbNullString & ExSep 'Basisumsatz
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'WKZ Basis
    TmpSt = TmpSt & KSoSt & ExSep 'Sachkonto
    TmpSt = TmpSt & KHaSt & ExSep 'Gegenkonto
    TmpSt = TmpSt & Chr$(34) & StSch & Chr$(34) & ExSep 'BU-Schluessel
    TmpSt = TmpSt & Format(BuDat, "ddmm") & ExSep 'Belegdatum
    TmpSt = TmpSt & Chr$(34) & ReStr & Chr$(34) & ExSep 'Belegfeld1
    TmpSt = TmpSt & Chr$(34) & BelFe & Chr$(34) & ExSep 'Belegfeld2
    TmpSt = TmpSt & vbNullString & ExSep 'Skonto
    TmpSt = TmpSt & Chr$(34) & BuStr & Chr$(34) & ExSep 'Buchungstext
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Postensperre (Text)
    TmpSt = TmpSt & Chr$(34) & PaNum & Chr$(34) & ExSep 'Adressnummer
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Geschaeftspartnerbank (Text)
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Sachverhalt (Text)
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Zinssperre (Text)
    ' Beleglink - use DATEV_CreateBeleglink for correct format
    ' BEDI nur setzen wenn Beleg gueltig (bei Ausgaben: Datei existiert in GlBPf)
    If Len(DaNam) > 0 And Len(BeGui) > 0 And BeVor = False And IsExpenseDocumentValid(BuGui) Then
        TmpSt = TmpSt & Chr$(34) & DATEV_CreateBeleglink(BeGui) & Chr$(34) & ExSep
    Else
        TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep
    End If
    If DebNr > 0 Then
        TmpSt = TmpSt & Chr$(34) & BELEGINFO_DEBITORNR_ART & Chr$(34) & ExSep 'Beleginfo Art 1: Debitorennr
        TmpSt = TmpSt & Chr$(34) & CStr(DebNr) & Chr$(34) & ExSep             'Beleginfo Inhalt 1: Debitorennr
    Else
        TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Beleginfo Art 1 (leer)
        TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Beleginfo Inhalt 1 (leer)
    End If
    TmpSt = TmpSt & Chr$(34) & Beleg & Chr$(34) & ExSep  'Beleginfo2a
    TmpSt = TmpSt & Chr$(34) & DaNam & Chr$(34) & ExSep 'Beleginfo2b
    ' Beleginfo 3: Patientennummer als Debitoren-Referenz
    If Len(PaNum) > 0 Then
        TmpSt = TmpSt & Chr$(34) & BELEGINFO_PATIENT_ART & Chr$(34) & ExSep
        TmpSt = TmpSt & Chr$(34) & PaNum & Chr$(34) & ExSep
    Else
        TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    End If
    ' Empty fields 4a-8b
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    TmpSt = TmpSt & Chr$(34) & Koste & Chr$(34) & ExSep 'KOST1
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'KOST2
    TmpSt = TmpSt & vbNullString & ExSep 'Kostmenge
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'UStID
    TmpSt = TmpSt & vbNullString & ExSep 'EU-Steuersatz
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Versteuerungsart
    ' L+L fields
    TmpSt = TmpSt & vbNullString & ExSep & vbNullString & ExSep
    ' BU49 fields
    TmpSt = TmpSt & vbNullString & ExSep & vbNullString & ExSep & vbNullString & ExSep
    ' Info fields 1-20 (mostly empty)
    If Len(Komme) > 0 Then
        TmpSt = TmpSt & Chr$(34) & "KOMMENTAR" & Chr$(34) & ExSep & Chr$(34) & Komme & Chr$(34) & ExSep
    Else
        TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    End If
    Dim ii As Integer
    For ii = 2 To 20
        TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep & Chr$(34) & Chr$(34) & ExSep
    Next ii
    ' Remaining fields
    TmpSt = TmpSt & vbNullString & ExSep 'Stueck (Numeric)
    TmpSt = TmpSt & vbNullString & ExSep 'Gewicht (Numeric)
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Zahlweise (Text)
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Forderungsart (Text)
    TmpSt = TmpSt & Format(BuDat, "yyyy") & ExSep 'Veranlagungsjahr
    TmpSt = TmpSt & vbNullString & ExSep 'Faelligkeit
    TmpSt = TmpSt & vbNullString & ExSep 'Skontotyp
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Auftragsnummer
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Buchungstyp
    TmpSt = TmpSt & vbNullString & ExSep 'Ust-Schluessel
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'EU-Land
    TmpSt = TmpSt & vbNullString & ExSep 'L+L3
    TmpSt = TmpSt & vbNullString & ExSep 'L+L Steuersatz
    TmpSt = TmpSt & vbNullString & ExSep 'Erloeskonto
    TmpSt = TmpSt & Chr$(34) & "WK" & Chr$(34) & ExSep 'Herkunft
    TmpSt = TmpSt & Chr$(34) & BuGui & Chr$(34) & ExSep 'GuiID
    TmpSt = TmpSt & vbNullString & ExSep 'KOST-Datum
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Mandatsreferenz
    TmpSt = TmpSt & Chr$(34) & "0" & Chr$(34) & ExSep 'Skontosperre
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Gesellschaftername
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Beteiligtennummer
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Identifikationsnummer
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Zeichnernummer
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'Postensperre
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'SoBil-Sachverhalt
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep 'SoBil-Buchung
    TmpSt = TmpSt & Chr$(34) & "0" & Chr$(34) & ExSep 'Festschreibung (114)

    ' Fields 115-116: Leistungsdatum, Datum Zuord. Steuerperiode (empty, unquoted)
    TmpSt = TmpSt & vbNullString & ExSep   ' 115: Leistungsdatum
    TmpSt = TmpSt & vbNullString & ExSep   ' 116: Datum Zuord. Steuerperiode

    ' Fields 117-125: DATEV v700 Format v13 additional fields
    ' Textfelder mit "" quotieren, numerische Felder leer
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep   ' 117: Generalumkehr (GU) - Text
    TmpSt = TmpSt & vbNullString & ExSep          ' 118: Steuersatz - Numeric
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep   ' 119: Land - Text
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep   ' 120: Abrechnungsreferenz - Text
    TmpSt = TmpSt & vbNullString & ExSep          ' 121: BVV-Position - Numeric
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep   ' 122: EU-Land u. UStID (Ursprungsland) - Text
    TmpSt = TmpSt & Chr$(34) & Chr$(34) & ExSep   ' 123: EU-USt-IdNr (Ursprung) - Text
    TmpSt = TmpSt & vbNullString & ExSep          ' 124: Sachverhalt Warenbewegung - Numeric
    TmpSt = TmpSt & vbNullString                  ' 125: Steuerschluessel Devisen - Numeric (last field)

    BuildCSVDataLineFromReportControl = TmpSt

    ' GoBD Festschreibung
    If GlBuG = True And BLock = False And IdxNr > 0 Then
        DBCmEx2 "qrySimBuL7", "@IdLock", "@IdxNr", -1, IdxNr
    End If

    Exit Function

ErrHandler:
    LogError "BuildCSVDataLineFromReportControl", Err.Number, Err.Description
    BuildCSVDataLineFromReportControl = vbNullString
End Function

'================================================================================
' Helper functions for ReportControl export
'================================================================================
Private Function HasField(ByRef RST As ADODB.Recordset, ByVal FieldName As String) As Boolean
    On Error Resume Next
    Dim Fld As ADODB.Field
    HasField = False
    For Each Fld In RST.Fields
        If UCase$(Fld.Name) = UCase$(FieldName) Then
            HasField = True
            Exit For
        End If
    Next Fld
End Function

Private Function SafeLongField(ByRef RST As ADODB.Recordset, ByVal FieldName As String) As Long
    On Error Resume Next
    SafeLongField = 0
    If HasField(RST, FieldName) Then
        If Not IsNull(RST.Fields(FieldName).Value) Then
            If IsNumeric(RST.Fields(FieldName).Value) Then
                SafeLongField = CLng(RST.Fields(FieldName).Value)
            End If
        End If
    End If
End Function

Private Function SafeIntField(ByRef RST As ADODB.Recordset, ByVal FieldName As String) As Integer
    On Error Resume Next
    SafeIntField = 0
    If HasField(RST, FieldName) Then
        If Not IsNull(RST.Fields(FieldName).Value) Then
            If IsNumeric(RST.Fields(FieldName).Value) Then
                SafeIntField = CInt(RST.Fields(FieldName).Value)
            End If
        End If
    End If
End Function

Private Function SafeSingleField(ByRef RST As ADODB.Recordset, ByVal FieldName As String) As Single
    On Error Resume Next
    SafeSingleField = 0
    If HasField(RST, FieldName) Then
        If Not IsNull(RST.Fields(FieldName).Value) Then
            If IsNumeric(RST.Fields(FieldName).Value) Then
                SafeSingleField = CSng(RST.Fields(FieldName).Value)
            End If
        End If
    End If
End Function

Private Function SafeBoolField(ByRef RST As ADODB.Recordset, ByVal FieldName As String) As Boolean
    On Error Resume Next
    SafeBoolField = False
    If HasField(RST, FieldName) Then
        If Not IsNull(RST.Fields(FieldName).Value) Then
            SafeBoolField = CBool(RST.Fields(FieldName).Value)
        End If
    End If
End Function

Private Function SafeStringField(ByRef RST As ADODB.Recordset, ByVal FieldName As String) As String
    On Error Resume Next
    SafeStringField = vbNullString
    If HasField(RST, FieldName) Then
        If Not IsNull(RST.Fields(FieldName).Value) Then
            SafeStringField = CStr(RST.Fields(FieldName).Value)
        End If
    End If
End Function

Private Function S_DaKoZ(TmVal As Variant) As Long
'Konvertiert Variant zu Long, behandelt Null-Werte
On Error Resume Next

    If IsNull(TmVal) = True Then
        S_DaKoZ = 0
    Else
        S_DaKoZ = Val(TmVal)
    End If

End Function

Private Function S_DaExF(ByVal KntNr As Long) As String
'Sachkontenformatierung
On Error GoTo OpErr

    Dim KntSt As String
    Dim KoStr As String
    Dim Lange As Integer

    KoStr = CStr(KntNr)

    If KntNr = 0 Then
        If GldKt = True Then
            KoStr = "0000"
        Else
            KoStr = "000000"
        End If
    Else
        If GldKt = True Then
            Lange = Len(KoStr)
            If Lange = 2 Then
                KntSt = KoStr & "00"
            ElseIf Lange = 3 Then
                KntSt = KoStr & "0"
            ElseIf Lange = 4 Then
                KntSt = KoStr
            ElseIf Lange > 4 Then
                KntSt = Left$(KoStr, 4)
            Else
                KntSt = Format$(KntNr, "0000")
            End If
        Else
            Lange = Len(KoStr)
            If Lange = 2 Then
                KntSt = KoStr & "0000"
            ElseIf Lange = 3 Then
                KntSt = KoStr & "000"
            ElseIf Lange = 4 Then
                KntSt = KoStr & "00"
            ElseIf Lange = 5 Then
                KntSt = KoStr & "0"
            ElseIf Lange = 6 Then
                KntSt = KoStr
            ElseIf Lange > 6 Then
                KntSt = Left$(KoStr, 6)
            Else
                KntSt = Format$(KntNr, "000000")
            End If
        End If
    End If

    S_DaExF = KntSt

Exit Function

OpErr:
If GlDbg = True Then SPopu "S_DaExF " & Err.Number, Err.Description, IC48_Warning
Resume Next

End Function

'--------------------------------------------------------------------------------
' S_DaKoK - Gegenkonto ermitteln aus Geldkonten-Array
'--------------------------------------------------------------------------------
' Purpose:     Ermittelt das DATEV-Sachkonto fuer ein Geldkonto (Bank/Kasse)
'              Urspruenglich aus basDatRe.bas - jetzt lokal integriert
'
' Parameters:  GeKto   - Geldkonto-ID (IDB aus Buchung)
'
' Returns:     DATEV-Sachkontonummer (4- oder 6-stellig je nach GldKt)
'--------------------------------------------------------------------------------
Private Function S_DaKoK(ByVal GeKto As Integer) As Long
On Error Resume Next

    Dim SaKto As Long
    Dim AktZa As Integer

    For AktZa = 1 To UBound(GlGeK)  ' Geldkonten-Array
        If GlGeK(AktZa, 0) = GeKto Then
            If GlGeK(AktZa, 2) <> vbNullString Then
                SaKto = S_DaExF(GlGeK(AktZa, 2))
            Else
                ' Fallback: Standard-Bankkonto
                If GldKt = True Then
                    SaKto = 1200
                Else
                    SaKto = 120000
                End If
            End If
            Exit For
        End If
    Next AktZa

    S_DaKoK = SaKto

End Function

'--------------------------------------------------------------------------------
' S_AdIdx - Patientendetails aus Adressindex ermitteln
'--------------------------------------------------------------------------------
' Purpose:     Ermittelt Feldwert aus qryAdrIdx fuer eine Patienten-ID
'              Urspruenglich aus basData.bas - jetzt lokal integriert
'
' Parameters:  PatNr   - Patienten-ID (IdxNr)
'              FelNa   - Feldname (z.B. "IDP" fuer Patientennummer)
'
' Returns:     Feldwert als String, oder Leerstring/0 bei Fehler
'--------------------------------------------------------------------------------
Private Function S_AdIdx(ByVal PatNr As Long, ByVal FelNa As String) As String
On Error GoTo LoErr

    Dim RST As ADODB.Recordset

    Set RST = New ADODB.Recordset
    RST.CursorLocation = adUseClient
    Set RST = DBCmRe1("qryAdrIdx", "@IdxNr", PatNr)

    If RST.RecordCount > 0 Then
        If Not IsNull(RST.Fields(FelNa).Value) And RST.Fields(FelNa).Value <> vbNullString Then
            S_AdIdx = RST.Fields(FelNa).Value
        Else
            ' Leerer Wert - Typ-abhaengiger Default
            Select Case RST.Fields(FelNa).Type
            Case adBigInt, adBoolean, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt
                S_AdIdx = "0"
            Case Else
                S_AdIdx = vbNullString
            End Select
        End If
    Else
        ' Kein Datensatz gefunden
        S_AdIdx = vbNullString
    End If

    RST.Close
    Set RST = Nothing
    Exit Function

LoErr:
    If GlDbg = True Then SPopu "S_AdIdx " & Err.Number, Err.Description, IC48_Warning
    S_AdIdx = vbNullString
    Resume Next

End Function

Private Function DetermineAmountFromReportControl(ByRef RST As ADODB.Recordset, ByVal BuTyp As Integer) As String
    Dim Amount As Currency
    Amount = 0

    On Error Resume Next
    If GlBuc = True Then
        Select Case BuTyp
        Case 1: ' Ausgaben
            If HasField(RST, "Ausgabe") Then Amount = CCur(RST.Fields("Ausgabe").Value)
            If Amount = 0 And HasField(RST, "Einnahme") Then Amount = CCur(RST.Fields("Einnahme").Value)
        Case 2: ' Einnahmen
            If HasField(RST, "Einnahme") Then Amount = CCur(RST.Fields("Einnahme").Value)
            If Amount = 0 And HasField(RST, "Ausgabe") Then Amount = CCur(RST.Fields("Ausgabe").Value)
        Case Else:
            If HasField(RST, "Einnahme") And CCur(RST.Fields("Einnahme").Value) > 0 Then
                Amount = CCur(RST.Fields("Einnahme").Value)
            ElseIf HasField(RST, "Ausgabe") And CCur(RST.Fields("Ausgabe").Value) > 0 Then
                Amount = CCur(RST.Fields("Ausgabe").Value)
            End If
        End Select
    Else
        If HasField(RST, "Einnahme") Then Amount = CCur(RST.Fields("Einnahme").Value)
    End If
    On Error GoTo 0

    If Amount < 0.01 Then
        DetermineAmountFromReportControl = vbNullString
    Else
        ' Use FormatAmountGermanOptimized for consistent DATEV format
        ' (no thousands separator, comma as decimal separator, 2 decimal places)
        DetermineAmountFromReportControl = FormatAmountGermanOptimized(Amount)
    End If
End Function

Private Function DetermineSollHabenFromReportControl(ByRef RST As ADODB.Recordset, _
                                                     ByVal BuTyp As Integer, _
                                                     ByVal SwapDebitCredit As Boolean) As String
    Dim Kennz As String

    If GlBuc = True Then
        Select Case BuTyp
        Case 1: ' Ausgaben
            If SwapDebitCredit Then Kennz = "S" Else Kennz = "H"
        Case 2: ' Einnahmen
            If SwapDebitCredit Then Kennz = "H" Else Kennz = "S"
        Case Else:
            On Error Resume Next
            If HasField(RST, "Einnahme") And CCur(RST.Fields("Einnahme").Value) > 0 Then
                If SwapDebitCredit Then Kennz = "H" Else Kennz = "S"
            ElseIf HasField(RST, "Ausgabe") And CCur(RST.Fields("Ausgabe").Value) > 0 Then
                If SwapDebitCredit Then Kennz = "S" Else Kennz = "H"
            Else
                Kennz = "S"
            End If
            On Error GoTo 0
        End Select
    Else
        If SwapDebitCredit Then Kennz = "H" Else Kennz = "S"
    End If

    DetermineSollHabenFromReportControl = Kennz
End Function

Private Sub CollectInvoiceIDFromReportControl(ByRef RST As ADODB.Recordset)
    On Error Resume Next
    Dim InvID As Long

    ' Debitoren-Modus: ID1 (Rechnung-PK) statt IDR (Buchungs-Referenz)
    If m_InvMod Then
        If HasField(RST, "ID1") Then
            If Not IsNull(RST.Fields("ID1").Value) Then
                InvID = CLng(RST.Fields("ID1").Value)
                If InvID > 0 Then
                    m_InvoiceCount = m_InvoiceCount + 1
                    ReDim Preserve GloDr(m_InvoiceCount)
                    GloDr(m_InvoiceCount) = InvID
                End If
            End If
        End If
    Else
        ' Column Caption: "IDR" ist auch im ReportControl vorhanden
        If HasField(RST, "IDR") Then
            If Not IsNull(RST.Fields("IDR").Value) Then
                InvID = CLng(RST.Fields("IDR").Value)
                If InvID > 0 Then
                    m_InvoiceCount = m_InvoiceCount + 1
                    ReDim Preserve GloDr(m_InvoiceCount)
                    GloDr(m_InvoiceCount) = InvID
                End If
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub GenerateXMLFromReportControl(ByRef RST As ADODB.Recordset, _
                                         ByVal XMLFileName As String, _
                                         ByRef Config As DATEV_ExportConfig)
'--------------------------------------------------------------------------------
' Purpose:     Generate document.xml using Column Captions from ReportControl
'              Replaces legacy S_DaExX call with new BuildDocumentXMLElement
'
' Parameters:  RST         - Recordset with Column Captions as field names
'              XMLFileName - Full path for output document.xml
'              Config      - Export configuration
'--------------------------------------------------------------------------------
On Error GoTo ErrHandler

    Dim XMLContent As String
    Dim DocumentsDir As String
    Dim DocumentCount As Long
    Dim RecordCount As Long
    Dim CurrentRecord As Long
    Dim ProcessedDocs As Collection

    ' Initialize
    DocumentCount = 0
    Set ProcessedDocs = New Collection
    DocumentsDir = Config.ExportPath

    ' Build XML header
    XMLContent = BuildXMLHeader(Config)

    ' Process records for documents
    RecordCount = RST.RecordCount
    RST.MoveFirst
    CurrentRecord = 0

    Do While Not RST.EOF
        CurrentRecord = CurrentRecord + 1

        ' Update progress
        If (CurrentRecord Mod PROGRESS_UPDATE_INTERVAL) = 0 Then
            UpdateDualProgress CurrentRecord, RecordCount, "XML Beleg " & CurrentRecord & " von " & RecordCount
            DoEvents
        End If

        ' Process document for this record using Column Captions
        Dim DocXML As String
        DocXML = ProcessDocumentForXMLFromRC(RST, Config, DocumentsDir, ProcessedDocs)
        If Len(DocXML) > 0 Then
            XMLContent = XMLContent & DocXML
            DocumentCount = DocumentCount + 1
        End If

        RST.MoveNext
    Loop

    ' Build XML footer
    XMLContent = XMLContent & BuildXMLFooter(Config)

    ' Write XML file
    If WriteXMLFile(XMLFileName, XMLContent) Then
        If GlLog = True Then SLogi "DATEV: document.xml erstellt mit " & DocumentCount & " Dokumenten"
    Else
        If GlLog = True Then SLogi "DATEV: document.xml konnte nicht geschrieben werden"
    End If

    ' Cleanup
    Set ProcessedDocs = Nothing
    Exit Sub

ErrHandler:
    Set ProcessedDocs = Nothing
    LogError "GenerateXMLFromReportControl", Err.Number, Err.Description
End Sub

'--------------------------------------------------------------------------------
' ProcessDocumentForXMLFromRC - Process document using Column Captions
'--------------------------------------------------------------------------------
' Purpose:     Creates XML document element for ReportControl recordset
'              Uses Column Captions: "IDA", "GuiID", "Datei", "Belegzeichen", etc.
'
' Parameters:  RST           - Recordset with Column Captions
'              Config        - Export configuration
'              DocumentsDir  - Directory for document copies
'              ProcessedDocs - Collection to track processed GUIDs
'
' Returns:     XML string for one document element, or empty if no document
'--------------------------------------------------------------------------------
Private Function ProcessDocumentForXMLFromRC(ByRef RST As ADODB.Recordset, _
                                             ByRef Config As DATEV_ExportConfig, _
                                             ByVal DocumentsDir As String, _
                                             ByRef ProcessedDocs As Collection) As String
On Error GoTo ErrHandler

    Dim DocXML As String
    Dim BuTyp As Integer
    Dim BuGui As String
    Dim DaNam As String
    Dim DaPfa As String
    Dim ReStr As String
    Dim Komme As String
    Dim BuDat As Date
    Dim BuDatVal As Variant
    Dim IsRevenue As Boolean
    Dim FormattedGUID As String
    Dim TargetFileName As String
    Dim TargetPath As String
    Dim DocExists As Boolean
    Dim CleanGUID As String
    Dim LedgerFileName As String

    DocXML = vbNullString

    ' Get booking type (invoice mode: always revenue)
    If m_InvMod Then
        BuTyp = 2
    Else
        BuTyp = SafeIntField(RST, "IDA")
    End If
    IsRevenue = (BuTyp = 2)

    ' Skip records with zero/invalid amount (same validation as CSV export)
    ' This ensures XML only contains documents for records that are in the CSV
    Dim Amount As Currency
    Amount = DetermineAmountValueFromRC(RST, BuTyp)
    If Amount < MIN_VALID_AMOUNT Then
        ProcessDocumentForXMLFromRC = vbNullString
        Exit Function
    End If

    ' Get GUID - Column Caption: "GuiID"
    BuGui = SafeStringField(RST, "GuiID")
    If Len(BuGui) = 0 Then
        ProcessDocumentForXMLFromRC = vbNullString
        Exit Function
    End If

    ' Skip payment bookings (K prefix) - they don't have document files
    ' Only process R (invoices), B (expenses), G (other business docs)
    If Len(BuGui) > 0 Then
        Dim GuidPrefix As String
        GuidPrefix = UCase$(Left$(BuGui, 1))
        If GuidPrefix = "K" Then
            If GlLog = True Then SLogi "  >>> ProcessDocumentForXMLFromRC: Skipped (payment booking, no document)"
            ProcessDocumentForXMLFromRC = vbNullString
            Exit Function
        End If
    End If

    ' Check if expense document was validated as invalid (skip in invoice mode)
    If Not m_InvMod Then
        If Not IsExpenseDocumentValid(BuGui) Then
            ProcessDocumentForXMLFromRC = vbNullString
            Exit Function
        End If
    End If

    ' Format GUID for XML (8-4-4-4-12 lowercase)
    FormattedGUID = DATEV_FormatGUIDForXML(BuGui)

    ' Check if already processed (split posting handling)
    Dim AlreadyProcessed As Boolean
    Dim j As Long
    AlreadyProcessed = False
    If ProcessedDocs.Count > 0 Then
        For j = 1 To ProcessedDocs.Count
            If StrComp(ProcessedDocs.Item(j), FormattedGUID, vbTextCompare) = 0 Then
                AlreadyProcessed = True
                Exit For
            End If
        Next j
    End If
    If AlreadyProcessed Then
        ProcessDocumentForXMLFromRC = vbNullString
        Exit Function
    End If

    ' Get document info - Column Captions
    If m_InvMod Then
        DaNam = vbNullString
        ReStr = SanitizeTextField(SafeStringField(RST, "Rechnung"), MAX_BELEGFELD1_LENGTH)
    Else
        DaNam = SafeStringField(RST, "Datei")
        ReStr = SanitizeTextField(SafeStringField(RST, "Belegzeichen"), MAX_BELEGFELD1_LENGTH)
    End If
    DaPfa = vbNullString
    Komme = SanitizeTextField(SafeStringField(RST, "Kommentar"), 60)

    ' Get document date - Column Caption: "Datum"
    BuDatVal = RST.Fields("Datum").Value
    If Not IsNull(BuDatVal) And IsDate(BuDatVal) Then
        BuDat = CDate(BuDatVal)
    Else
        BuDat = Date
    End If

    ' Determine if we have/need a document
    If Len(DaNam) = 0 Then
        If IsRevenue Then
            ' Revenue without document - generate default filename
            If Len(ReStr) > 0 Then
                DaNam = "Rechnung_Beleg_" & SanitizeFileName(ReStr) & ".pdf"
            Else
                DaNam = "Rechnung_Beleg_" & SanitizeFileName(FormattedGUID) & ".pdf"
            End If
        Else
            ' Expense without document - skip
            ProcessDocumentForXMLFromRC = vbNullString
            Exit Function
        End If
    End If

    ' Find document source path
    Dim SourceFound As Boolean
    SourceFound = False
    If Len(DaPfa) = 0 Then
        Dim FoundPath As String
        FoundPath = FindDocumentPath(DaNam, Config.MandantNr)
        If Len(FoundPath) > 0 Then
            DaPfa = FoundPath
            SourceFound = True
        End If
    End If
    DocExists = SourceFound

    ' Skip if document does not exist - ensures XML matches CSV Beleglinks
    ' Exception: Revenue/invoice documents are generated during export
    If Not DocExists Then
        ' For expenses, require physical document
        ' For revenues, allow (PDF will be generated later in export)
        If Not IsRevenue Then
            If GlLog = True Then SLogi "  >>> ProcessDocumentForXMLFromRC: Skipped (expense document not found)"
            ProcessDocumentForXMLFromRC = vbNullString
            Exit Function
        Else
            If GlLog = True Then SLogi "  >>> ProcessDocumentForXMLFromRC: Revenue document will be generated"
        End If
    End If

    ' Prepare target filename
    TargetFileName = SanitizeFileName(DaNam)
    If Len(TargetFileName) = 0 Then
        TargetFileName = "Rechnung_Beleg_" & SanitizeFileName(FormattedGUID) & ".pdf"
    End If

    ' Ensure .pdf extension for revenues
    If IsRevenue Then
        If LCase$(Right$(TargetFileName, 4)) <> ".pdf" Then
            TargetFileName = TargetFileName & ".pdf"
        End If
    End If

    ' Create BEDI filename for document.xml
    CleanGUID = Replace(FormattedGUID, "-", vbNullString)
    CleanGUID = UCase$(CleanGUID)
    LedgerFileName = BELEGLINK_PREFIX & CleanGUID & GetFileExtension(TargetFileName)
    TargetPath = DocumentsDir & LedgerFileName

    ' Copy document if source exists and ExportDocuments is enabled
    If DocExists And Config.ExportDocuments Then
        If CopyDocumentToExport(DaPfa, TargetPath) Then
            AddToZipListRC LedgerFileName
        End If
    End If

    ' Mark document as processed
    ProcessedDocs.Add FormattedGUID, FormattedGUID

    ' Build XML document element with property keys
    ' Option A format: File extension with property keys for Buchungsperiode and Rechnungsnummer
    If Config.UseLedgerXML Then
        ' Option B: Ledger format
        DocXML = BuildLedgerDocumentXMLElement(FormattedGUID, CleanGUID, LedgerFileName, BuDat, IsRevenue)
    Else
        ' Option A: Simple document format with property keys
        DocXML = BuildDocumentXMLElement(FormattedGUID, LedgerFileName, BuDat, ReStr, Komme)
    End If

    ProcessDocumentForXMLFromRC = DocXML
    Exit Function

ErrHandler:
    LogError "ProcessDocumentForXMLFromRC", Err.Number, Err.Description
    ProcessDocumentForXMLFromRC = vbNullString
End Function

Private Function CreateZIPArchiveFromReportControl(ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler
' Erstellt ZIP-Archiv aus dem Export-Unterordner

Dim ZipName As String
Dim ZipPath As String
Dim SrcPath As String

' Benutzer-gewaehlten Dateinamen verwenden falls vorhanden
If Len(Config.ExportFileName) > 0 Then
    ZipName = Config.ExportFileName & ".zip"
Else
    ZipName = "EXTF_DATEV_" & Format$(Now, "YYYYMMDD_HHMM") & ".zip"
End If

' ZIP file goes to original export path (not in subfolder)
ZipPath = m_OrgPfa & ZipName

' Source path is the subfolder (Config.ExportPath now points to subfolder)
' SimpliZip expects a real folder/file path, not a wildcard mask
SrcPath = Config.ExportPath

If GlLog = True Then SLogi "=== CreateZIPArchiveFromReportControl ==="
If GlLog = True Then SLogi "  ZIP File: " & ZipPath
If GlLog = True Then SLogi "  Source Folder: " & SrcPath
DoEvents

' Compress folder and delete source via SZipp (basWindow)
' SZipp waits synchronously for SimpliZip.exe to finish
If SZipp(ZipPath, SrcPath, True) = True Then
    If GlLog = True Then SLogi "SZipp erfolgreich beendet"
    If m_clFil.FilVor(ZipPath) Then
        If GlLog = True Then SLogi "ZIP-Datei erstellt: " & ZipPath
        CreateZIPArchiveFromReportControl = ZipPath
    Else
        If GlLog = True Then SLogi "SZipp beendet, aber ZIP-Datei nicht gefunden: " & ZipPath
        CreateZIPArchiveFromReportControl = vbNullString
    End If
Else
    If GlLog = True Then SLogi "SZipp fehlgeschlagen - Dateien bleiben unkomprimiert"
    CreateZIPArchiveFromReportControl = vbNullString
End If

Exit Function

ErrHandler:
If GlLog = True Then SLogi "=== CreateZIPArchiveFromReportControl ERROR ==="
If GlLog = True Then SLogi "  Err.Number: " & Err.Number
If GlLog = True Then SLogi "  Err.Description: " & Err.Description
DoEvents
CreateZIPArchiveFromReportControl = vbNullString

End Function

Private Sub AddToZipListRC(ByVal FilePath As String)
    ' Version fuer ReportControl Export - aktualisiert m_ZipFiles Array

    m_ZipFileCount = m_ZipFileCount + 1
    ReDim Preserve m_ZipFiles(m_ZipFileCount - 1)
    m_ZipFiles(m_ZipFileCount - 1) = FilePath
End Sub

Private Function BuildDATEVHeaderLine(ByVal DaSt1 As String, ByVal DaSt2 As String, _
                                       ByVal DaSt3 As String, ByVal DaSt4 As String, _
                                       ByVal BerSt As String, ByVal MaNam As String, _
                                       ByVal AnzKo As Integer) As String
    Dim ExSep As String
    Dim CSVep As String
    Dim ManNa As String
    Dim AktZa As Integer

    ExSep = Chr$(59)
    CSVep = Chr$(34) & Chr$(59) & Chr$(34)

    ' Mandantenname ermitteln
    ManNa = vbNullString
    If m_Config.MandantNr > 0 Then
        For AktZa = 1 To UBound(GlMan)
            If m_Config.MandantNr = GlMan(AktZa, 2) Then
                ManNa = GlMan(AktZa, 3)
                Exit For
            End If
        Next AktZa
    Else
        ManNa = GlMan(GlSMa, 3)
    End If
    If Len(ManNa) > 2 Then
        ManNa = SUmw(ManNa, False, True, True)
        If Len(ManNa) > 25 Then ManNa = Left$(ManNa, 25)
    Else
        ManNa = "Admin"
    End If
    ManNa = Replace(ManNa, Chr$(32), vbNullString)
    ManNa = Replace(ManNa, Chr$(58), vbNullString)

    ' Build EXTF header line (31 fields per DATEV v700 spec)
    ' 1=EXTF, 2=Version, 3=Category, 4=Name, 5=FormatVer, 6=Generated, 7=Imported,
    ' 8=Origin, 9=ExportedBy, 10=ImportedBy, 11=Berater, 12=Mandant, 13=WJBeginn,
    ' 14=Kontenlaenge, 15=DatumVon, 16=DatumBis, 17=Bezeichnung, 18=Diktat,
    ' 19=Buchungstyp, 20=Rechnungslegung, 21=Festschreibung, 22=WKZ, 23-26=reserved,
    ' 27=SKR, 28=Branchenloesung, 29-30=reserved, 31=Anwendungsinfo
    BuildDATEVHeaderLine = Chr$(34) & "EXTF" & Chr$(34) & ExSep & _
        "700" & ExSep & "21" & ExSep & _
        Chr$(34) & "Buchungsstapel" & Chr$(34) & ExSep & "13" & ExSep & _
        DaSt4 & "000" & ExSep & ExSep & _
        Chr$(34) & "SV" & Chr$(34) & ExSep & _
        Chr$(34) & ManNa & Chr$(34) & ExSep & _
        Chr$(34) & Chr$(34) & ExSep & _
        BerSt & ExSep & MaNam & ExSep & DaSt3 & ExSep & AnzKo & ExSep & _
        DaSt1 & ExSep & DaSt2 & ExSep & _
        Chr$(34) & "Buchungen" & Chr$(34) & ExSep & _
        Chr$(34) & Chr$(34) & ExSep & _
        "1" & ExSep & "0" & ExSep & "0" & ExSep & _
        Chr$(34) & "EUR" & Chr$(34) & ExSep & _
        ExSep & ExSep & ExSep & ExSep & _
        Chr$(34) & Chr$(34) & ExSep & _
        ExSep & ExSep & ExSep & _
        Chr$(34) & Chr$(34)
End Function

Private Function BuildColumnHeaderLine() As String
    BuildColumnHeaderLine = "Umsatz;Soll/Haben;WKZ_Umsatz;Kurs;Basisumsatz;WKZ_Basisumsatz;Konto;Gegenkonto;BU-Schluessel;Belegdatum;Belegfeld1;Belegfeld2;Skonto;Buchungstext;" & _
        "Postensperre;Adressnummer;Geschaeftspartnerbank;Sachverhalt;Zinssperre;Beleglink;Beleginfo1a;Beleginfo1b;Beleginfo2a;Beleginfo2b;Beleginfo3a;Beleginfo3b;" & _
        "Beleginfo4a;Beleginfo4b;Beleginfo5a;Beleginfo5b;Beleginfo6a;Beleginfo6b;Beleginfo7a;Beleginfo7b;Beleginfo8a;Beleginfo8b;KOST1;KOST2;Menge;UStID;" & _
        "Steuersatz;Versteuerungsart;L+L1;L+L2;BU49a;BU49b;BU49c;Info1a;Infor1b;Info2a;Info2b;Zusatzin3a;Info3b;Info4a;Info4b;Info5a;Info5b;Info6a;Info6b;" & _
        "Info7a;Info7b;Info8a;Info8b;Info9a;Info9b;Info10a;Info10b;Info11a;Info11b;Info12a;Info12b;Info13a;Info13b;Info14a;Info14b;Info15a;Info15b;Info16a;" & _
        "Info16b;Info17a;Info17b;Info18a;Info18b;Info19a;Info19b;Info20a;Info20b;Stueck;Gewicht;Zahlweise;Forderungsart;Veranlagungsjahr;Faelligkeit;Skontotyp;" & _
        "Auftragsnummer;Buchungstyp;Ust-Schluessel;EU-Land;L+L3;Steuersatz;Erloeskonto;Herkunft;GUID;KOST-Datum;SEPA-Mandatsreferenz;Skontosperre;Gesellschaftername;" & _
        "Beteiligtennummer;Identifikationsnummer;Zeichnernummer;Postensperre;SoBil-Sachverhalt;SoBil-Buchung;Festschreibung"
End Function

'================================================================================
' DATEV_ExportSelected
'--------------------------------------------------------------------------------
' Purpose:     Export selected bookings to DATEV format
'              Replaces: basDaMa.S_Expor() for DATEV exports
'
' Parameters:  Config      - Export configuration
'              RST         - Recordset with selected bookings
'
' Returns:     DATEV_ExportResult with status and file paths
'================================================================================
Public Function DATEV_ExportSelected(ByRef Config As DATEV_ExportConfig, _
                                     ByRef RST As ADODB.Recordset) As DATEV_ExportResult
On Error GoTo ErrHandler

    Dim Result As DATEV_ExportResult
    Dim ValidCount As Long

    ' Initialize result
    Result.success = False
    Result.RecordCount = 0
    Result.DocumentCount = 0
    Result.ErrorMessage = vbNullString
    Result.ErrorCode = 0

    ' Initialize module state
    m_Cancelled = False
    m_Config = Config
    m_ColumnHeaders = vbNullString  ' Reset cached column headers

    ' Validate configuration
    If Not ValidateConfig(Config) Then
        Result.ErrorMessage = "Ungueltige Exportkonfiguration: Beraternummer=" & Config.Beraternummer & _
                             ", Mandantennummer=" & Config.Mandantennummer & _
                             ", Pfad=" & Config.ExportPath
        Result.ErrorCode = 1001
        DATEV_ExportSelected = Result
        Exit Function
    End If

    ' Validate recordset
    If RST Is Nothing Then
        Result.ErrorMessage = "Keine Daten zum Exportieren vorhanden"
        Result.ErrorCode = 1002
        DATEV_ExportSelected = Result
        Exit Function
    End If

    If RST.EOF And RST.BOF Then
        Result.ErrorMessage = "Keine Datensaetze im Ergebnis"
        Result.ErrorCode = 1003
        DATEV_ExportSelected = Result
        Exit Function
    End If

    ' Initialize file operations
    Set m_clFil = New clsFile

    ' Initialize invoice ID collection for PDF generation
    ReDim GloDr(0)
    m_InvoiceCount = 0

    If GlLog = True Then SLogi "=== DATEV_ExportSelected START ==="
    If GlLog = True Then SLogi "Config.ExportDocuments = " & Config.ExportDocuments
    If GlLog = True Then SLogi "Config.ExportCSV = " & Config.ExportCSV
    If GlLog = True Then SLogi "Config.ExportXML = " & Config.ExportXML
    If GlLog = True Then SLogi "Config.ExportPath = " & Config.ExportPath
    DoEvents

    ' Ensure export directory exists
    If Not EnsureExportDirectory(Config.ExportPath) Then
        Result.ErrorMessage = "Exportverzeichnis kann nicht erstellt werden: " & Config.ExportPath
        Result.ErrorCode = 1004
        GoTo Cleanup
    End If

    ' Setup subfolder for export if compression is requested
    If Config.CompressOutput And Config.ExportDocuments Then
        If Not SetSubFo(Config) Then
            Result.ErrorMessage = "Export-Unterordner konnte nicht erstellt werden"
            Result.ErrorCode = 1005
            GoTo Cleanup
        End If
    Else
        ' No compression - save original path for compatibility
        m_OrgPfa = Config.ExportPath
        m_SubNam = vbNullString
    End If

    ' Initialize ZIP file list
    ReDim m_ZipFiles(0)
    m_ZipFileCount = 0

    ' Initialize document tracking collection
    Set m_DocumentGUIDs = New Collection
    Set m_ExportedFileNames = New Collection
    Set m_InvoiceGUIDs = New Collection
    Set m_InvoiceRechNrs = New Collection
    Set m_InvalidDocuments = New Collection  ' Initialize early to avoid Nothing errors

    ' Show progress dialog early with Marquee mode
    ' This ensures the user sees activity immediately during data loading
    ShowProgressDialogInit "DATEV Export"

    ' Validate expense documents BEFORE export (only if document export is enabled)
    ' This checks if all expense documents (Datei field) exist in GlBPf
    ' Results are stored in m_InvalidDocuments for fast lookup during export
    ' Debitoren-Modus: keine Validierung noetig (kein Datei-Feld, nur generierte PDFs)
    If Config.ExportDocuments And Not m_InvMod Then
        ValidateExpenseDocuments RST, Config
    End If

    ' Collect invoice IDs for PDF generation BEFORE any export
    ' This must happen regardless of CSV/XML export type
    If Config.ExportDocuments Then
        If GlLog = True Then SLogi "=== Collecting Invoice IDs ==="
        RST.MoveFirst
        Do While Not RST.EOF
            CollectInvoiceID RST
            RST.MoveNext
        Loop
        RST.MoveFirst
        If GlLog = True Then SLogi "Total m_InvoiceCount = " & m_InvoiceCount
        DoEvents
    End If

    ' Generate CSV export if requested
    If Config.ExportCSV Then
        ShowProgressDialog "DATEV Export - CSV Buchungsstapel", RST.RecordCount

        Result.CSVFilePath = GenerateCSVExport(RST, Config, ValidCount)
        If Result.CSVFilePath = vbNullString Then
            If m_Cancelled Then
                Result.ErrorMessage = "Export durch Benutzer abgebrochen"
                Result.ErrorCode = 0
            Else
                Result.ErrorMessage = "CSV-Datei konnte nicht erstellt werden"
                Result.ErrorCode = 2001
            End If
            GoTo Cleanup
        End If
    End If

    ' Generate XML document linking if requested
    If Config.ExportXML Then
        If Config.ExportCSV Then
            frmStatus.Caption = "DATEV Export - XML Belegverknuepfung"
            DoEvents
        Else
            ShowProgressDialog "DATEV Export - XML Belegverknuepfung", RST.RecordCount
        End If
        RST.MoveFirst
        Result.XMLFilePath = GenerateXMLExport(RST, Config, Result.CSVFilePath)
        If Result.XMLFilePath = vbNullString And Not m_Cancelled Then
            If Config.ExportCSV Then
                ' XML generation failed but CSV succeeded - continue with warning
                Result.ErrorMessage = "CSV exportiert, aber XML-Generierung fehlgeschlagen"
            Else
                Result.ErrorMessage = "XML-Datei konnte nicht erstellt werden"
                Result.ErrorCode = 2002
                GoTo Cleanup
            End If
        End If
    End If

    ' Check if any export was done
    If Not Config.ExportCSV And Not Config.ExportXML Then
        Result.ErrorMessage = "Kein Export-Typ ausgew?hlt"
        Result.ErrorCode = 2003
        GoTo Cleanup
    End If

    ' Generate PDF documents BEFORE ZIP if ExportDocuments is enabled and invoices were found
    If GlLog = True Then SLogi "=== PDF Generation Check ==="
    If GlLog = True Then SLogi "Config.ExportDocuments = " & Config.ExportDocuments
    If GlLog = True Then SLogi "m_InvoiceCount = " & m_InvoiceCount
    DoEvents
    If Config.ExportDocuments And m_InvoiceCount > 0 Then
        ' Update status for PDF generation phase
        frmStatus.lblLab01.Caption = "Generiere " & m_InvoiceCount & " PDF-Beleg(e)..."
        DoEvents
        If GlLog = True Then SLogi ">>> Calling GenerateInvoicePDFs..."
        GenerateInvoicePDFs Config
        If GlLog = True Then SLogi ">>> GenerateInvoicePDFs completed"
        DoEvents
    Else
        If GlLog = True Then SLogi ">>> PDF generation SKIPPED (ExportDocuments=" & Config.ExportDocuments & ", InvoiceCount=" & m_InvoiceCount & ")"
        DoEvents
    End If

    ' Create ZIP archive if requested (AFTER PDF generation so PDFs are included)
    ' Skip ZIP creation if only CSV is exported (no documents) - ZIP makes no sense for single file
    If Config.CompressOutput And Config.ExportDocuments Then
        ' Update status for ZIP creation phase
        frmStatus.lblLab01.Caption = "Erstelle ZIP-Archiv..."
        DoEvents
        If GlLog = True Then SLogi "=== Creating ZIP Archive ==="
        If GlLog = True Then SLogi "m_ZipFileCount = " & m_ZipFileCount
        Dim zz As Integer
        If GlLog = True Then
            For zz = 0 To m_ZipFileCount - 1
                SLogi "  m_ZipFiles(" & zz & ") = " & m_ZipFiles(zz)
            Next zz
        End If
        DoEvents
        Result.ZipFilePath = CreateZIPArchive(Config, Result.CSVFilePath, Result.XMLFilePath)
        If GlLog = True Then SLogi "ZIP created: " & Result.ZipFilePath
        DoEvents

    End If

    ' Get final counts
    If Config.ExportCSV Then
        Result.RecordCount = ValidCount
    Else
        Result.RecordCount = RST.RecordCount
    End If
    If Not m_DocumentGUIDs Is Nothing Then
        Result.DocumentCount = m_DocumentGUIDs.Count
    End If
    Result.success = True

    ' Reset invoice collection
    ReDim GloDr(0)
    m_InvoiceCount = 0
    If GlLog = True Then SLogi "=== DATEV_ExportSelected END ==="
    DoEvents

Cleanup:
    ' Always cleanup resources
    HideProgressDialog
    CleanupModuleState

    ' Send via email if requested and export succeeded
    If Result.success And Config.EmailAfterExport > 0 Then
        SendExportViaEmail Config, Result
    End If

    DATEV_ExportSelected = Result
    Exit Function

ErrHandler:
    Result.success = False
    Result.ErrorMessage = "Exportfehler: " & Err.Description
    Result.ErrorCode = Err.Number
    LogError "DATEV_ExportSelected", Err.Number, Err.Description
    Resume Cleanup
End Function

'================================================================================
' CleanupModuleState - Release resources after export
'================================================================================
Private Sub CleanupModuleState()
    On Error Resume Next

    ' Release file operations object
    Set m_clFil = Nothing

    ' Release PDF generator object
    Set m_clLis = Nothing

    ' Clear document tracking collection
    Set m_DocumentGUIDs = Nothing
    Set m_ExportedFileNames = Nothing
    Set m_InvoiceGUIDs = Nothing
    Set m_InvoiceRechNrs = Nothing

    ' Clear cached values
    m_ColumnHeaders = vbNullString
    m_Q = vbNullString
    m_Sep = vbNullString
    m_EmptyQuoted = vbNullString

    ' Clear line buffer
    Erase m_LineBuffer
    m_LineBufferPos = 0

    ' Clear ZIP file list
    Erase m_ZipFiles
    m_ZipFileCount = 0

    ' Reset invoice mode flag
    m_InvMod = False

    On Error GoTo 0
End Sub

'================================================================================
' CollectInvoiceID - Collect invoice IDs from recordset for PDF generation
' Exakt wie in S_BuEx (basData.bas Zeile 13840-13843): CLng() Konvertierung!
' HINWEIS: ZIP-Liste wird NACH LLExDv in AddGeneratedPDFsToZipList befuellt,
'          um sicherzustellen, dass die tatsaechlichen Dateinamen verwendet werden
'================================================================================
Private Sub CollectInvoiceID(ByRef RST As ADODB.Recordset)
On Error Resume Next

    Dim InvID As Long

    ' Debitoren-Modus: ID1 (Rechnung-PK) statt IDR (Buchungs-Referenz)
    If m_InvMod Then
        InvID = CLng(RST.Fields("ID1").Value)
    Else
        ' Exakt wie S_BuEx: CLng() fuer explizite Long-Konvertierung
        InvID = CLng(RST.Fields("IDR").Value)
    End If

    If InvID > 0 Then
        m_InvoiceCount = m_InvoiceCount + 1
        ReDim Preserve GloDr(m_InvoiceCount)
        GloDr(m_InvoiceCount) = InvID
        ' PDF-Dateinamen werden NACH LLExDv via AddGeneratedPDFsToZipList hinzugefuegt
    End If

    Err.Clear
End Sub

'================================================================================
' GenerateInvoicePDFs - Generate PDF documents using clsLisLab (like s_buex/s_expor)
'================================================================================
Private Sub GenerateInvoicePDFs(ByRef Config As DATEV_ExportConfig)
On Error Resume Next

    Dim DaPfa As String
    Dim DaNaO As String
    Dim FiNam As String
    Dim Formu As Boolean
    Dim i As Long

    If GlLog = True Then SLogi "=== GenerateInvoicePDFs START ==="
    If GlLog = True Then SLogi "m_InvoiceCount = " & m_InvoiceCount
    DoEvents

    ' List all collected invoice IDs
    If GlLog = True Then
        SLogi "GloDr contents:"
        For i = 1 To m_InvoiceCount
            If i <= 10 Then
                SLogi "  GloDr(" & i & ") = " & GloDr(i)
            End If
        Next i
        If m_InvoiceCount > 10 Then
            SLogi "  ... and " & (m_InvoiceCount - 10) & " more"
        End If
        DoEvents
    End If

    ' Get export path
    DaPfa = Config.ExportPath
    If Right$(DaPfa, 1) <> "\" Then DaPfa = DaPfa & "\"

    ' Generate base filename - Benutzer-gewaehlten Dateinamen verwenden falls vorhanden
    If Len(Config.ExportFileName) > 0 Then
        DaNaO = Config.ExportFileName
    Else
        DaNaO = "EXTF_DATEV_" & Format$(Now, "YYYYMMDD_HHMM")
    End If

    ' Formular-Pruefung: Zugeordnetes Formular vorhanden?
    FiNam = GlFrO & S_FoCh("Rechnu")
    If m_clFil.FilVor(FiNam) = True Then
        Formu = True
    Else
        Formu = False
    End If

    ' Fallback auf Standardformular wenn zugeordnetes Formular nicht existiert
    If Formu = False Then
        If GlFrn <> vbNullString Then
            FiNam = GlFrn & "standardrechnung.blg"
        Else
            FiNam = GlFrO & "standardrechnung.blg"
        End If
        If GlLog = True Then SLogi "Formular-Fallback: " & FiNam
    End If

    If GlLog = True Then SLogi "DaPfa = " & DaPfa
    If GlLog = True Then SLogi "DaNaO = " & DaNaO
    If GlLog = True Then SLogi "FiNam (Formular) = " & FiNam
    If GlLog = True Then SLogi "GlTmp = " & GlTmp
    DoEvents

    ' Create and configure PDF generator
    If GlLog = True Then SLogi "Creating clsLisLab..."
    Set m_clLis = New clsLisLab
    If m_clLis Is Nothing Then
        If GlLog = True Then SLogi "ERROR: clsLisLab creation failed!"
        Exit Sub
    End If
    If GlLog = True Then SLogi "clsLisLab created successfully"
    DoEvents

    With m_clLis
        .ForNam = "Rechnu"
        .FilNam = FiNam
        .PfaTmp = GlTmp
        .ExpFmt = "PDF"
        .StaVer = DaPfa
        .DatNam = DaPfa & DaNaO & ".pdf"
        .DruDia = False
        .DruVor = False
        .MitaVo = GlMiV
        .ArztVo = GlArV
        .MandVo = GlMaV

        If GlLog = True Then SLogi "Calling .LLExDv..."
        If GlLog = True Then SLogi "  .ForNam = " & .ForNam
        If GlLog = True Then SLogi "  .FilNam = " & .FilNam
        If GlLog = True Then SLogi "  .ExpFmt = " & .ExpFmt
        If GlLog = True Then SLogi "  .StaVer = " & .StaVer
        If GlLog = True Then SLogi "  .DatNam = " & .DatNam
        DoEvents

        .LLExDv

        ' frmStatus wieder anzeigen nach LLExDv (LLExDv hat eigene Statusanzeige)
        On Error Resume Next
        frmStatus.Show vbModeless
        frmStatus.Caption = "DATEV Export - PDF Belege"
        DoEvents
        On Error GoTo 0

        If GlLog = True Then SLogi ".LLExDv completed"
        DoEvents
        If Err.Number <> 0 Then
            If GlLog = True Then SLogi "ERROR after LLExDv: " & Err.Number & " - " & Err.Description
            DoEvents
        End If
    End With
    Set m_clLis = Nothing

    ' Nach LLExDv: Tatsaechlich erstellte PDFs zur ZIP-Liste hinzufuegen
    ' LLExDv erstellt Dateien mit Namen aus qryPrSimRe1.Rechnungsnummer
    ' Diese koennen von RechNr (aus qrySimBuSu) abweichen
    AddGeneratedPDFsToZipList DaPfa

    If GlLog = True Then SLogi "=== GenerateInvoicePDFs END ==="
    DoEvents

End Sub

'================================================================================
' AddGeneratedPDFsToZipList - Sucht erstellte PDF-Rechnungen, benennt sie in BEDI-Format um
'                              und fuegt sie zur ZIP-Liste hinzu
'================================================================================
Private Sub AddGeneratedPDFsToZipList(ByVal ExportPath As String)
On Error Resume Next

    Dim FileName As String
    Dim RechNr As String
    Dim GuiID As String
    Dim CleanGUID As String
    Dim NewFileName As String
    Dim SourcePath As String
    Dim TargetPath As String
    Dim FileList As Collection
    Dim i As Long

    If GlLog = True Then SLogi "=== AddGeneratedPDFsToZipList START ==="
    If GlLog = True Then SLogi "Suche in: " & ExportPath
    DoEvents

    ' WICHTIG: Erst alle Dateinamen sammeln, dann verarbeiten
    ' Dir$() hat einen internen Zustand der durch Kill/Name/Dir$ zerstoert wird
    Set FileList = New Collection
    FileName = Dir$(ExportPath & "Rechnung_Beleg_*.pdf")
    Do While FileName <> vbNullString
        FileList.Add FileName
        FileName = Dir$()
    Loop

    If GlLog = True Then SLogi "Gefunden: " & FileList.Count & " Dateien"
    DoEvents

    ' Status aktualisieren vor Verarbeitung
    On Error Resume Next
    If FileList.Count > 0 Then
        frmStatus.lblLab01.Caption = "Benenne " & FileList.Count & " Debitorenbelege um..."
        frmStatus.prbStat1.Scrolling = xtpProgressBarStandard
        frmStatus.prbStat1.Max = FileList.Count
        frmStatus.prbStat1.Value = 0
        DoEvents
    End If
    On Error GoTo 0

    ' Jetzt die gesammelten Dateien verarbeiten
    For i = 1 To FileList.Count
        FileName = FileList.Item(i)

        ' Fortschritt aktualisieren
        On Error Resume Next
        frmStatus.lblLab01.Caption = "Debitorenbeleg " & i & " von " & FileList.Count & "..."
        frmStatus.prbStat1.Value = i
        DoEvents
        On Error GoTo 0

        If GlLog = True Then SLogi "  Verarbeite: " & FileName

        ' Extrahiere RechNr aus Dateinamen: Rechnung_Beleg_<RechNr>.pdf
        ' "Rechnung_Beleg_" = 15 Zeichen, ".pdf" = 4 Zeichen
        RechNr = Mid$(FileName, 16, Len(FileName) - 19)
        If GlLog = True Then SLogi "    RechNr extrahiert: " & RechNr

        ' Suche GuiID in der Mapping-Collection (uses iteration-based lookup)
        GuiID = GetInvoiceGUID(RechNr)

        If Len(GuiID) > 0 Then
            ' GUID gefunden - Datei in BEDI-Format umbenennen
            ' DATEV_FormatGUIDForXML entfernt B/R/G Praefix und formatiert als Standard-GUID
            Dim FormattedGUID As String
            FormattedGUID = DATEV_FormatGUIDForXML(GuiID)
            CleanGUID = UCase$(Replace(FormattedGUID, "-", vbNullString))
            NewFileName = BELEGLINK_PREFIX & CleanGUID & ".pdf"

            SourcePath = ExportPath & FileName
            TargetPath = ExportPath & NewFileName

            If GlLog = True Then SLogi "    Umbenennung: " & FileName & " -> " & NewFileName

            ' Pruefen ob Zieldatei bereits existiert (mit clFil.FilVor, nicht Dir$!)
            If m_clFil.FilVor(TargetPath) Then
                ' Zieldatei existiert bereits - Quelldatei loeschen, Ziel zur Liste
                If GlLog = True Then SLogi "    Zieldatei existiert bereits, Quelldatei wird geloescht"
                On Error Resume Next
                Kill SourcePath
                On Error GoTo 0
                AddToZipList NewFileName
            Else
                ' Datei umbenennen (VB6 Name-Statement)
                On Error Resume Next
                Name SourcePath As TargetPath
                If Err.Number = 0 Then
                    If GlLog = True Then SLogi "    Erfolgreich umbenannt!"
                    AddToZipList NewFileName
                Else
                    If GlLog = True Then SLogi "    Fehler beim Umbenennen: " & Err.Description
                    ' Bei Fehler: Originaldatei zur ZIP-Liste hinzufuegen
                    AddToZipList FileName
                End If
                On Error GoTo 0
            End If
        Else
            ' Kein GUID gefunden - keine Buchung vorhanden, Datei loeschen
            If GlLog = True Then SLogi "    Kein GUID gefunden, keine Buchung - Datei wird geloescht"
            On Error Resume Next
            Kill ExportPath & FileName
            On Error GoTo 0
            ' Datei wird NICHT zur ZIP-Liste hinzugefuegt
        End If
    Next i

    If GlLog = True Then SLogi "=== AddGeneratedPDFsToZipList END ==="
    DoEvents

    Set FileList = Nothing
    Err.Clear
End Sub

'================================================================================
' ValidateExportConsistency - Prueft Konsistenz zwischen PDF-Dateien und XML-Referenzen
'--------------------------------------------------------------------------------
' Purpose:     Zaehlt physisch vorhandene PDF-Dateien und vergleicht mit
'              Dokument-Referenzen in document.xml
'
' Parameters:  Result      - Export-Ergebnis (wird mit Konsistenz-Info ergaenzt)
'              ExportPath  - Pfad zum Export-Verzeichnis
'================================================================================
Public Sub ValidateExportConsistency(ByRef Result As DATEV_ExportResult, _
                                     ByVal ExportPath As String)
On Error Resume Next

    Dim PDFCount As Long
    Dim XMLRefCount As Long
    Dim FileName As String
    Dim XMLPath As String
    Dim FileNum As Integer
    Dim XMLContent As String
    Dim SearchPos As Long
    Dim XMLExists As Boolean

    If GlLog = True Then SLogi "=== ValidateExportConsistency START ==="
    DoEvents

    ' Speichere ExportPath fuer Cleanup
    Result.ExportPath = ExportPath

    ' Zaehle physische PDF-Dateien (BEDI*.pdf)
    PDFCount = 0
    FileName = Dir$(ExportPath & "BEDI*.pdf")
    Do While FileName <> vbNullString
        PDFCount = PDFCount + 1
        FileName = Dir$()
    Loop
    Result.PDFFileCount = PDFCount

    If GlLog = True Then SLogi "  PDF-Dateien gefunden: " & PDFCount
    DoEvents

    ' Zaehle Dokument-Referenzen in document.xml
    If GlLog = True Then SLogi "  ExportPath: " & ExportPath
    DoEvents
    XMLPath = ExportPath & "document.xml"
    If GlLog = True Then SLogi "  XMLPath: " & XMLPath
    DoEvents
    XMLRefCount = 0

    ' Pruefe ob XML-Datei existiert (mit Dir$ statt m_clFil, da m_clFil evtl. Nothing ist)
    If GlLog = True Then SLogi "  Pruefe XML-Existenz mit Dir$..."
    DoEvents
    XMLExists = (Dir$(XMLPath) <> vbNullString)
    If GlLog = True Then SLogi "  XMLExists: " & XMLExists
    DoEvents

    If XMLExists Then
        ' Datei lesen und xsi:type="File" zaehlen
        FileNum = FreeFile
        Open XMLPath For Binary Access Read As #FileNum
        XMLContent = Space$(LOF(FileNum))
        Get #FileNum, , XMLContent
        Close #FileNum

        ' Zaehle Vorkommen von 'xsi:type="File"'
        SearchPos = 1
        Do
            SearchPos = InStr(SearchPos, XMLContent, "xsi:type=""File""", vbTextCompare)
            If SearchPos > 0 Then
                XMLRefCount = XMLRefCount + 1
                SearchPos = SearchPos + 1
            End If
        Loop While SearchPos > 0
    End If
    Result.XMLDocumentCount = XMLRefCount

    If GlLog = True Then SLogi "  XML-Referenzen gefunden: " & XMLRefCount
    DoEvents

    ' Konsistenz pruefen
    If PDFCount = XMLRefCount Then
        Result.ConsistencyOK = True
        Result.ConsistencyMessage = vbNullString
    Else
        Result.ConsistencyOK = False
        If PDFCount < XMLRefCount Then
            Result.ConsistencyMessage = "WARNUNG: " & (XMLRefCount - PDFCount) & _
                " PDF-Dateien fehlen (" & PDFCount & " vorhanden, " & XMLRefCount & " referenziert)"
        Else
            Result.ConsistencyMessage = "WARNUNG: " & (PDFCount - XMLRefCount) & _
                " PDF-Dateien ohne XML-Referenz (" & PDFCount & " vorhanden, " & XMLRefCount & " referenziert)"
        End If
        If GlLog = True Then SLogi "  " & Result.ConsistencyMessage
    End If

    If GlLog = True Then SLogi "=== ValidateExportConsistency END ==="
    DoEvents

End Sub

'================================================================================
' CountDocumentsByType - Zaehlt exportierte Belege nach Typ (Debitoren/Kreditoren)
'--------------------------------------------------------------------------------
' Purpose:     Zaehlt die Anzahl der exportierten Belege getrennt nach
'              Debitorenbelegen (Einnahmen) und Kreditorenbelegen (Ausgaben).
'              Verwendet das Recordset um den BuTyp (IDA-Feld) zu bestimmen.
'
' Parameters:  RST         - Recordset mit Buchungen
'              Result      - Export-Ergebnis (wird mit Belegzaehlung ergaenzt)
'              ExportPath  - Pfad zum Export-Verzeichnis (fuer Datei-Pruefung)
'
' Logik:       - Debitor (BuTyp=2, Einnahme): IDR > 0 = PDF wird generiert
'              - Kreditor (BuTyp=1, Ausgabe): Datei-Feld nicht leer UND Datei existiert
'================================================================================
Public Sub CountDocumentsByType(ByRef RST As ADODB.Recordset, _
                                ByRef Result As DATEV_ExportResult, _
                                ByVal ExportPath As String)
On Error Resume Next

    Dim BuTyp As Integer
    Dim IDR As Long
    Dim DaNam As String
    Dim FullPath As String
    Dim clFil As clsFile
    Dim GuiID As String

    If GlLog = True Then SLogi "CountDocumentsByType: START"
    DoEvents

    Result.DebitDocCount = 0
    Result.KreditDocCount = 0

    If GlLog = True Then SLogi "CountDocumentsByType: Counters initialized"
    DoEvents

    ' Pruefe ob Recordset gueltig und offen ist
    If RST Is Nothing Then
        If GlLog = True Then SLogi "CountDocumentsByType: RST Is Nothing - Exit"
        DoEvents
        Exit Sub
    End If

    If GlLog = True Then SLogi "CountDocumentsByType: RST not Nothing, checking State..."
    DoEvents

    ' Pruefe ob Recordset offen ist (State = 1 = adStateOpen)
    If RST.State <> 1 Then
        If GlLog = True Then SLogi "CountDocumentsByType: RST not open (State=" & RST.State & ") - Exit"
        DoEvents
        Exit Sub
    End If

    If GlLog = True Then SLogi "CountDocumentsByType: RST State OK, checking EOF/BOF..."
    DoEvents

    ' Pruefe ob Recordset leer ist
    If RST.EOF And RST.BOF Then
        If GlLog = True Then SLogi "CountDocumentsByType: RST is empty - Exit"
        DoEvents
        Exit Sub
    End If

    If GlLog = True Then SLogi "CountDocumentsByType: RST has records, creating clFil..."
    DoEvents

    Set clFil = New clsFile

    If GlLog = True Then SLogi "CountDocumentsByType: clFil created, calling MoveFirst..."
    DoEvents

    RST.MoveFirst

    If GlLog = True Then SLogi "CountDocumentsByType: MoveFirst done, starting loop..."
    DoEvents

    Do While Not RST.EOF
        ' BuTyp ermitteln (IDA-Feld)
        BuTyp = 0
        If HasField(RST, "IDA") Then
            If Not IsNull(RST.Fields("IDA").Value) Then
                BuTyp = CInt(RST.Fields("IDA").Value)
            End If
        End If

        If BuTyp = 2 Then
            ' Debitor (Einnahme): Pruefe ob IDR > 0 (= Rechnung vorhanden)
            IDR = 0
            If HasField(RST, "IDR") Then
                If Not IsNull(RST.Fields("IDR").Value) Then
                    IDR = CLng(RST.Fields("IDR").Value)
                End If
            End If
            If IDR > 0 Then
                Result.DebitDocCount = Result.DebitDocCount + 1
            End If
        ElseIf BuTyp = 1 Then
            ' Kreditor (Ausgabe): Pruefe ob Datei-Feld nicht leer UND Datei existiert
            DaNam = vbNullString
            If HasField(RST, "Datei") Then
                If Not IsNull(RST.Fields("Datei").Value) Then
                    DaNam = Trim$(RST.Fields("Datei").Value & vbNullString)
                End If
            End If
            If Len(DaNam) > 0 Then
                ' Pruefe ob BEDI-Datei im Exportverzeichnis existiert
                ' Die Ausgabebelege werden als BEDI<GUID>.pdf kopiert
                GuiID = vbNullString
                If HasField(RST, "GuiID") Then
                    If Not IsNull(RST.Fields("GuiID").Value) Then
                        GuiID = Trim$(RST.Fields("GuiID").Value & vbNullString)
                    End If
                End If
                If Len(GuiID) > 0 Then
                    FullPath = ExportPath
                    If Right$(FullPath, 1) <> "\" Then FullPath = FullPath & "\"
                    FullPath = FullPath & "BEDI" & GuiID & ".pdf"
                    If clFil.FilVor(FullPath) = True Then
                        Result.KreditDocCount = Result.KreditDocCount + 1
                    End If
                End If
            End If
        End If

        RST.MoveNext
    Loop

    Set clFil = Nothing

    If GlLog = True Then
        SLogi "CountDocumentsByType: DebitDocCount=" & Result.DebitDocCount & _
              ", KreditDocCount=" & Result.KreditDocCount
    End If

End Sub

'================================================================================
' ValidateExpenseDocuments - Vorpruefung der Ausgabe-Belege
'--------------------------------------------------------------------------------
' Purpose:     Prueft VOR dem Export ob alle Ausgabe-Belege (Datei-Feld) existieren.
'              Ergebnisse werden in m_InvalidDocuments gespeichert fuer schnellen
'              Zugriff waehrend des Exports.
'
' Logik:       - Nur Ausgaben haben einen Dateinamen im Feld "Datei"
'              - Einnahmen haben IDR > 0 aber kein Datei-Feld (PDFs werden generiert)
'              - Wenn Datei nicht in GlBPf gefunden: GuiID in m_InvalidDocuments
'              - Im Export: Keine BEDI wenn GuiID in m_InvalidDocuments
'
' Parameters:  RST         - Recordset mit Buchungen
'              Config      - Export-Konfiguration
'================================================================================
Private Sub ValidateExpenseDocuments(ByRef RST As ADODB.Recordset, _
                                     ByRef Config As DATEV_ExportConfig)
On Error Resume Next

    Dim DaNam As String
    Dim BuGui As String
    Dim FullPath As String
    Dim RecordCount As Long
    Dim CurrentRecord As Long
    Dim ExpenseCount As Long
    Dim RevenueCount As Long
    Dim InvalidCount As Long
    Dim BuTyp As Integer
    Dim BelZei As String
    Dim FoundPath As String
    Dim MandNr As Long

    If GlLog = True Then SLogi "=== ValidateExpenseDocuments START ==="
    DoEvents

    ' Initialize collection for invalid documents
    Set m_InvalidDocuments = New Collection

    ' Count records for progress
    RST.MoveFirst
    RecordCount = RST.RecordCount
    If RecordCount = 0 Then
        If GlLog = True Then SLogi "Keine Datensaetze vorhanden"
        DoEvents
        Exit Sub
    End If

    ' Show progress for validation phase
    ShowProgressDialog "DATEV Export - Belegpruefung", RecordCount

    CurrentRecord = 0
    ExpenseCount = 0
    RevenueCount = 0
    InvalidCount = 0

    ' Loop through all records
    Do While Not RST.EOF
        CurrentRecord = CurrentRecord + 1

        ' Update progress every 10 records
        If CurrentRecord Mod 10 = 0 Or CurrentRecord = RecordCount Then
            UpdateProgress CurrentRecord, RecordCount, "Pruefe Beleg " & CurrentRecord & " von " & RecordCount
        End If

        ' Get booking type (IDA column = BuTyp)
        BuTyp = 0
        If HasField(RST, "IDA") Then
            If Not IsNull(RST.Fields("IDA").Value) Then
                BuTyp = CInt(RST.Fields("IDA").Value)
            End If
        End If

        ' Get GUID for this record
        BuGui = vbNullString
        If Not IsNull(RST.Fields("GuiID").Value) Then
            BuGui = Trim$(RST.Fields("GuiID").Value & vbNullString)
        End If

        ' Get document filename from Datei field
        DaNam = vbNullString
        If Not IsNull(RST.Fields("Datei").Value) Then
            DaNam = Trim$(RST.Fields("Datei").Value & vbNullString)
        End If

        ' Check expenses: records with Datei field
        If Len(DaNam) > 0 Then
            ExpenseCount = ExpenseCount + 1

            ' Check if file exists in GlBPf
            If Len(BuGui) > 0 Then
                FullPath = GlBPf
                If Right$(FullPath, 1) <> "\" Then FullPath = FullPath & "\"
                FullPath = FullPath & DaNam

                If m_clFil.FilVor(FullPath) = False Then
                    ' File not found - mark as invalid
                    InvalidCount = InvalidCount + 1
                    m_InvalidDocuments.Add BuGui  ' Store GuiID as value (no key) for safe iteration
                    If GlLog = True Then SLogi "  NICHT GEFUNDEN: " & DaNam & " (GuiID: " & BuGui & ")"
                Else
                    If GlLog = True And CurrentRecord <= 5 Then SLogi "  OK: " & DaNam
                End If
            End If
        ' Check revenues: BuTyp=2 with Belegzeichen but no Datei
        ElseIf BuTyp = 2 And Len(BuGui) > 0 Then
            ' Get Belegzeichen (invoice number)
            BelZei = vbNullString
            If HasField(RST, "Belegzeichen") Then
                If Not IsNull(RST.Fields("Belegzeichen").Value) Then
                    BelZei = Trim$(RST.Fields("Belegzeichen").Value & vbNullString)
                End If
            End If

            ' Revenue with invoice number - check if document can be found
            If Len(BelZei) > 0 Then
                RevenueCount = RevenueCount + 1

                ' Generate expected filename for revenue document
                DaNam = "Rechnung_Beleg_" & BelZei & ".pdf"

                ' Try to find document using FindDocumentPath (same as XML generation)
                ' Parameters: FileName, MandantNr (get from RST or default to 0)
                MandNr = 0
                If HasField(RST, "Mandant") Then
                    If Not IsNull(RST.Fields("Mandant").Value) Then
                        MandNr = CLng(RST.Fields("Mandant").Value)
                    End If
                End If
                FoundPath = FindDocumentPath(DaNam, MandNr)

                If Len(FoundPath) = 0 Then
                    ' Document not found - mark as invalid
                    InvalidCount = InvalidCount + 1
                    m_InvalidDocuments.Add BuGui
                    If GlLog = True And InvalidCount <= 10 Then SLogi "  EINNAHME OHNE BELEG: " & DaNam & " (GuiID: " & BuGui & ")"
                End If
            End If
        End If

        RST.MoveNext
    Loop

    ' Reset recordset position
    RST.MoveFirst

    If GlLog = True Then
        SLogi "Ausgaben mit Datei: " & ExpenseCount
        SLogi "Einnahmen mit Belegzeichen: " & RevenueCount
        SLogi "Davon nicht gefunden: " & InvalidCount
        SLogi "=== ValidateExpenseDocuments END ==="
        DoEvents
    End If

    Err.Clear
End Sub

'================================================================================
' IsExpenseDocumentValid - Prueft ob Beleg gueltig ist (Ausgaben UND Einnahmen)
'--------------------------------------------------------------------------------
' Purpose:     Schnelle Pruefung waehrend des Exports ob ein Beleg gueltig ist.
'              Verwendet die in ValidateExpenseDocuments erstellte Collection.
'              Prueft sowohl Ausgaben (Datei-Feld) als auch Einnahmen (Belegzeichen).
'
' Parameters:  GuiID       - GUID der Buchung
'
' Returns:     True wenn Beleg gueltig (gefunden) oder keine Pruefung noetig
'              False wenn Beleg ungueltig (nicht gefunden)
'================================================================================
Private Function IsExpenseDocumentValid(ByVal GuiID As String) As Boolean
    Dim i As Long
    Dim StoredGuiID As String

    ' If no GuiID, assume valid (nothing to check)
    If Len(GuiID) = 0 Then
        IsExpenseDocumentValid = True
        Exit Function
    End If

    ' If collection not initialized, create empty collection and assume valid
    If m_InvalidDocuments Is Nothing Then
        Set m_InvalidDocuments = New Collection
        IsExpenseDocumentValid = True
        Exit Function
    End If

    ' If collection is empty, all documents are valid
    If m_InvalidDocuments.Count = 0 Then
        IsExpenseDocumentValid = True
        Exit Function
    End If

    ' Iterate through collection to find GuiID (stored as values, not keys)
    ' This avoids errors when checking for non-existent keys
    For i = 1 To m_InvalidDocuments.Count
        StoredGuiID = m_InvalidDocuments.Item(i)
        If StrComp(StoredGuiID, GuiID, vbTextCompare) = 0 Then
            ' Found in invalid collection -> document is invalid
            IsExpenseDocumentValid = False
            Exit Function
        End If
    Next i

    ' Not found in invalid collection -> document is valid
    IsExpenseDocumentValid = True
End Function

'================================================================================
' DATEV_ExportByDateRange
'--------------------------------------------------------------------------------
' Purpose:     Export bookings by date criteria to DATEV format
'              Replaces: basData.S_BuEx() for DATEV exports
'
' Parameters:  Config      - Export configuration
'              SQLCriteria - SQL WHERE clause for date filtering
'
' Returns:     DATEV_ExportResult with status and file paths
'================================================================================
Public Function DATEV_ExportByDateRange(ByRef Config As DATEV_ExportConfig, _
                                        ByVal SQLCriteria As String) As DATEV_ExportResult
On Error GoTo ErrHandler

    Dim Result As DATEV_ExportResult
    Dim RST As ADODB.Recordset
    Dim SQL1 As String

    ' Initialize
    Result.success = False
    m_Config = Config

    ' Show progress dialog early with Marquee mode during data loading
    ShowProgressDialogInit "DATEV Export"

    ' Build SQL query based on database type
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qrySimBuSu WHERE " & SQLCriteria & " ORDER BY Datum"
    Else
        SQL1 = "SELECT * FROM qrySimBuSu WHERE " & SQLCriteria & " ORDER BY [Datum];"
    End If

    ' Execute query
    Set RST = New ADODB.Recordset
    With RST
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open SQL1, DB1, , , adCmdText
    End With

    If RST.EOF And RST.BOF Then
        Result.ErrorMessage = "No bookings found for specified criteria"
        Result.ErrorCode = 1003
        RST.Close
        Set RST = Nothing
        HideProgressDialog  ' Close dialog on early exit
        DATEV_ExportByDateRange = Result
        Exit Function
    End If

    ' Delegate to main export function
    Result = DATEV_ExportSelected(Config, RST)

    RST.Close
    Set RST = Nothing

    DATEV_ExportByDateRange = Result
    Exit Function

ErrHandler:
    Result.success = False
    Result.ErrorMessage = "Export error: " & Err.Description
    Result.ErrorCode = Err.Number
    If Not RST Is Nothing Then
        If RST.State = adStateOpen Then RST.Close
        Set RST = Nothing
    End If
    HideProgressDialog  ' Close dialog on error
    If GlLog = True Then SLogi "=== DATEV_ExportByDateRange ERROR ==="
    If GlLog = True Then SLogi "  Err.Number: " & Err.Number
    If GlLog = True Then SLogi "  Err.Description: " & Err.Description
    If GlLog = True Then SLogi "  Err.Source: " & Err.Source
    DoEvents
    SPopu "DATEV_ExportByDateRange " & Err.Number, Err.Description, IC48_Warning
    DATEV_ExportByDateRange = Result
End Function

'================================================================================
' DATEV_GetDefaultConfig
'--------------------------------------------------------------------------------
' Purpose:     Create default export configuration from global settings
'              Use this to initialize a Config structure before export
'
' Returns:     DATEV_ExportConfig populated with current global settings
'================================================================================
Public Function DATEV_GetDefaultConfig() As DATEV_ExportConfig
    Dim Config As DATEV_ExportConfig

    ' DATEV settings from globals
    Config.Beraternummer = GlDvB
    Config.Mandantennummer = GlDvM
    Config.FourDigitAccounts = GldKt
    Config.SwapDebitCredit = GlTSH
    Config.IncludePatientNumber = GlDeN
    Config.ReplaceAccountWithDebtor = GlDeE
    Config.IncludePatientName = GlDaP
    Config.ExportDocuments = GlBlE

    ' Default date range (current fiscal year)
    Config.DateFrom = DateSerial(Year(Date), 1, 1)
    Config.DateTo = Date
    Config.WJBeginn = DateSerial(Year(Date), 1, 1)

    ' Default export path
    If GlRDP = True Then
        If Len(GlIPf) > 0 Then
            Config.ExportPath = GlIPf
        Else
            Config.ExportPath = GlDpf & "Import\"
        End If
    Else
        If Len(GlExO) > 0 Then
            Config.ExportPath = GlExO
        Else
            Config.ExportPath = GlDpf & "Export\"
        End If
    End If

    ' Ensure trailing backslash
    If Right$(Config.ExportPath, 1) <> "\" Then
        Config.ExportPath = Config.ExportPath & "\"
    End If

    ' Default options
    Config.EmailAfterExport = 0
    Config.CompressOutput = False
    Config.EncryptOutput = False
    Config.EncryptPassword = vbNullString

    ' Default: export both CSV and XML
    Config.ExportCSV = True
    Config.ExportXML = True

    DATEV_GetDefaultConfig = Config
End Function

'================================================================================
' DATEV_FormatAccountNumber
'--------------------------------------------------------------------------------
' Purpose:     Format account number according to DATEV requirements
'              Replaces: basDatRe.S_DaExF()
'
' Parameters:  AccountNo   - Raw account number
'              FourDigit   - True for 4-digit, False for 6-digit
'
' Returns:     Formatted account number string
'================================================================================
Public Function DATEV_FormatAccountNumber(ByVal AccountNo As Long, _
                                          ByVal FourDigit As Boolean) As String
    If FourDigit Then
        DATEV_FormatAccountNumber = Format$(AccountNo, "0000")
    Else
        DATEV_FormatAccountNumber = Format$(AccountNo, "000000")
    End If
End Function

'================================================================================
' DATEV_CreateBeleglink
'--------------------------------------------------------------------------------
' Purpose:     Create DATEV Beleglink from GUID
'              Format: BEDI "GUID-with-dashes" (8-4-4-4-12 format)
'              Per DATEV specification, the GUID must be in quotes with dashes
'
' Parameters:  GUID        - Source GUID (with or without "G/R/B/K" prefix)
'
' Returns:     DATEV Beleglink string, e.g. BEDI "9e5dcc50-fdd9-46c3-82ac-a2c5dfdce141"
'================================================================================
Public Function DATEV_CreateBeleglink(ByVal guid As String) As String
    Dim CleanGUID As String
    Dim FirstChar As String
    Dim FormattedGUID As String

    If Len(guid) = 0 Then
        DATEV_CreateBeleglink = vbNullString
        Exit Function
    End If

    ' Remove G/R/B/K prefix if present (from CreateID function)
    ' G = Gutschrift, R = Rechnung, B = Beleg, K = Kreditor
    FirstChar = UCase$(Left$(guid, 1))
    If FirstChar = "G" Or FirstChar = "R" Or FirstChar = "B" Or FirstChar = "K" Then
        CleanGUID = Mid$(guid, 2)
    Else
        CleanGUID = guid
    End If

    ' Remove any dashes or braces
    CleanGUID = Replace(CleanGUID, "-", vbNullString)
    CleanGUID = Replace(CleanGUID, "{", vbNullString)
    CleanGUID = Replace(CleanGUID, "}", vbNullString)

    ' Ensure exactly 32 hex characters (pad with leading zeros if needed)
    If Len(CleanGUID) < 32 Then
        CleanGUID = String$(32 - Len(CleanGUID), "0") & CleanGUID
    ElseIf Len(CleanGUID) > 32 Then
        ' Take only first 32 characters
        CleanGUID = Left$(CleanGUID, 32)
    End If

    ' Format as 8-4-4-4-12 with dashes (DATEV XSD specification)
    FormattedGUID = Mid$(CleanGUID, 1, 8) & "-" & _
                    Mid$(CleanGUID, 9, 4) & "-" & _
                    Mid$(CleanGUID, 13, 4) & "-" & _
                    Mid$(CleanGUID, 17, 4) & "-" & _
                    Mid$(CleanGUID, 21, 12)

    ' DATEV Beleglink format: BEDI "guid-with-dashes"
    ' Use doubled quotes for CSV escaping (inner quotes must be escaped)
    DATEV_CreateBeleglink = BELEGLINK_PREFIX & " """"" & FormattedGUID & """"""
End Function

'================================================================================
' DATEV_FormatGUIDForXML
'--------------------------------------------------------------------------------
' Purpose:     Format GUID for XML document.xml (lowercase with dashes)
'              Per DATEV XSD specification: 8-4-4-4-12 format, max 36 chars
'
' Parameters:  GUID        - Source GUID (with or without "G/R/B/K" prefix)
'
' Returns:     XML-formatted GUID string, e.g. 9e5dcc50-fdd9-46c3-82ac-a2c5dfdce141
'================================================================================
Public Function DATEV_FormatGUIDForXML(ByVal guid As String) As String
    Dim CleanGUID As String
    Dim FirstChar As String

    If Len(guid) = 0 Then
        DATEV_FormatGUIDForXML = vbNullString
        Exit Function
    End If

    ' Remove G/R/B/K prefix if present (from CreateID function)
    ' G = Gutschrift, R = Rechnung, B = Beleg, K = Kreditor
    FirstChar = UCase$(Left$(guid, 1))
    If FirstChar = "G" Or FirstChar = "R" Or FirstChar = "B" Or FirstChar = "K" Then
        CleanGUID = Mid$(guid, 2)
    Else
        CleanGUID = guid
    End If

    ' Remove existing dashes/braces
    CleanGUID = Replace(CleanGUID, "-", vbNullString)
    CleanGUID = Replace(CleanGUID, "{", vbNullString)
    CleanGUID = Replace(CleanGUID, "}", vbNullString)

    ' Ensure exactly 32 hex characters
    If Len(CleanGUID) < 32 Then
        CleanGUID = String$(32 - Len(CleanGUID), "0") & CleanGUID
    ElseIf Len(CleanGUID) > 32 Then
        ' Take only first 32 characters
        CleanGUID = Left$(CleanGUID, 32)
    End If

    ' Format as 8-4-4-4-12 with dashes, lowercase (DATEV XSD specification)
    DATEV_FormatGUIDForXML = LCase$(Mid$(CleanGUID, 1, 8) & "-" & _
                                    Mid$(CleanGUID, 9, 4) & "-" & _
                                    Mid$(CleanGUID, 13, 4) & "-" & _
                                    Mid$(CleanGUID, 17, 4) & "-" & _
                                    Mid$(CleanGUID, 21, 12))
End Function

'--------------------------------------------------------------------------------
' PRIVATE - CSV Generation (Optimized Implementation)
'--------------------------------------------------------------------------------

Private Function GenerateCSVExport(ByRef RST As ADODB.Recordset, _
                                   ByRef Config As DATEV_ExportConfig, _
                                   Optional ByRef OutValidCount As Long = 0) As String
On Error GoTo ErrHandler

    Dim CSVLines() As String        ' Array-based buffer for all lines
    Dim LineCount As Long           ' Current line count
    Dim MaxLines As Long            ' Allocated array size
    Dim CSVFilePath As String
    Dim DataLine As String
    Dim RecordCount As Long
    Dim CurrentRecord As Long
    Dim ValidRecordCount As Long
    Dim MinDate As Date
    Dim MaxDate As Date
    Dim RecDateVal As Variant
    Dim FinalContent As String

    ' Initialize cached values for performance
    InitializeCachedValues

    ' Initialize counters
    ValidRecordCount = 0
    MinDate = #12/31/9999#
    MaxDate = #1/1/100#
    m_LastProgressUpdate = 0

    ' Get record count and pre-allocate array
    RecordCount = RST.RecordCount
    If RecordCount <= 0 Then
        GenerateCSVExport = vbNullString
        Exit Function
    End If

    ' Pre-allocate lines array (header + column headers + data lines)
    MaxLines = RecordCount + 10
    ReDim CSVLines(0 To MaxLines)
    LineCount = 0

    ' Single pass: determine date range AND collect data
    ' This avoids iterating twice through the recordset
    RST.MoveFirst
    CurrentRecord = 0

    Do While Not RST.EOF
        CurrentRecord = CurrentRecord + 1

        ' Update progress at intervals (not every record)
        If (CurrentRecord - m_LastProgressUpdate) >= PROGRESS_UPDATE_INTERVAL Then
            UpdateProgress CurrentRecord, RecordCount, "Verarbeite Buchung " & CurrentRecord & " von " & RecordCount
            m_LastProgressUpdate = CurrentRecord
        End If

        ' Check for cancellation at intervals
        If (CurrentRecord Mod DOEVENTS_INTERVAL) = 0 Then
            DoEvents
            If CheckCancelled() Then
                m_Cancelled = True
                Erase CSVLines
                GenerateCSVExport = vbNullString
                Exit Function
            End If
        End If

        ' Track date range from actual data
        RecDateVal = RST.Fields("Datum").Value
        If Not IsNull(RecDateVal) Then
            If IsDate(RecDateVal) Then
                Dim RecDate As Date
                RecDate = CDate(RecDateVal)
                ' Validate date is reasonable (1900-2100)
                If RecDate >= #1/1/1900# And RecDate <= #12/31/2100# Then
                    If RecDate < MinDate Then MinDate = RecDate
                    If RecDate > MaxDate Then MaxDate = RecDate
                End If
            End If
        End If

        ' Build data line
        DataLine = BuildCSVDataLineOptimized(RST, Config)
        If Len(DataLine) > 0 Then
            ' Store in array (offset by 2 for header lines)
            CSVLines(LineCount + 2) = DataLine
            ValidRecordCount = ValidRecordCount + 1
            LineCount = LineCount + 1

            ' Note: Invoice IDs are collected before export in DATEV_ExportSelected
        End If

        RST.MoveNext
    Loop

    If GlLog = True Then SLogi "=== CSV Export Loop completed ==="
    If GlLog = True Then SLogi "ValidRecordCount = " & ValidRecordCount
    DoEvents

    ' Final progress update
    UpdateProgress RecordCount, RecordCount, "Finalisiere Export..."

    ' Check if we have any valid records
    If ValidRecordCount = 0 Then
        Erase CSVLines
        GenerateCSVExport = vbNullString
        Exit Function
    End If

    ' Update config with actual date range
    If MinDate <= MaxDate Then
        Config.DateFrom = MinDate
        Config.DateTo = MaxDate
        ' Fix: Update WJBeginn to match fiscal year of actual data (not current system date)
        ' This prevents year boundary errors when exporting historical data
        Config.WJBeginn = DateSerial(Year(MinDate), 1, 1)
        If GlLog = True Then SLogi "=== WJBeginn updated to match data year: " & Format$(Config.WJBeginn, "yyyy-mm-dd") & " (data year: " & Year(MinDate) & ")"
    End If

    ' Build header lines (now that we know the date range)
    CSVLines(0) = BuildEXTFHeader(Config)
    CSVLines(1) = GetCachedColumnHeaders()

    ' Resize array to actual size
    ReDim Preserve CSVLines(0 To LineCount + 1)

    ' Join all lines with CRLF (single allocation)
    FinalContent = Join(CSVLines, vbCrLf) & vbCrLf

    ' Free array memory before file write
    Erase CSVLines

    ' Generate filename
    CSVFilePath = GenerateCSVFilename(Config)

    ' Write file using clFil
    If WriteCSVFile(CSVFilePath, FinalContent) Then
        ' Add to ZIP file list (relative path/filename only)
        AddToZipList Mid$(CSVFilePath, InStrRev(CSVFilePath, "\") + 1)
        GenerateCSVExport = CSVFilePath
        OutValidCount = ValidRecordCount
    Else
        GenerateCSVExport = vbNullString
    End If

    Exit Function

ErrHandler:
    Erase CSVLines
    LogError "GenerateCSVExport", Err.Number, Err.Description
    GenerateCSVExport = vbNullString
End Function

'--------------------------------------------------------------------------------
' PRIVATE - Initialize Cached Values (called once per export)
'--------------------------------------------------------------------------------
Private Sub InitializeCachedValues()
    m_Q = Chr$(34)
    m_Sep = DATEV_SEPARATOR
    m_EmptyQuoted = m_Q & m_Q

    ' Pre-allocate line buffer for field assembly
    ReDim m_LineBuffer(0 To 120)  ' 116 fields + margin
    m_LineBufferPos = 0
End Sub

'--------------------------------------------------------------------------------
' PRIVATE - Get Cached Column Headers
'--------------------------------------------------------------------------------
Private Function GetCachedColumnHeaders() As String
    ' Build column headers only once and cache
    If Len(m_ColumnHeaders) = 0 Then
        m_ColumnHeaders = BuildColumnHeaders()
    End If
    GetCachedColumnHeaders = m_ColumnHeaders
End Function

Private Function BuildEXTFHeader(ByRef Config As DATEV_ExportConfig) As String
    ' EXTF header format per DATEV specification version 700
    ' 21 fields in header line

    Dim Header As String
    Dim GeneratedOn As String
    Dim WJBeginn As String
    Dim DateFrom As String
    Dim DateTo As String
    Dim AccountLength As Integer
    Dim Q As String

    Q = Chr$(34) ' Quote character

    GeneratedOn = Format$(Now, "yyyymmddhhnnss") & "000"
    WJBeginn = Format$(Config.WJBeginn, "yyyymmdd")
    DateFrom = Format$(Config.DateFrom, "yyyymmdd")
    DateTo = Format$(Config.DateTo, "yyyymmdd")

    ' Diagnostic logging for header date fields
    If GlLog = True Then SLogi "=== EXTF Header: WJBeginn=" & WJBeginn & ", DateFrom=" & DateFrom & ", DateTo=" & DateTo

    If Config.FourDigitAccounts Then
        AccountLength = 4
    Else
        AccountLength = 6
    End If

    ' Build EXTF header line (31 fields per DATEV v700 spec)
    Header = Q & "EXTF" & Q & DATEV_SEPARATOR                               ' 1: Identifier
    Header = Header & CStr(DATEV_VERSION) & DATEV_SEPARATOR                 ' 2: Version (700)
    Header = Header & CStr(DATEV_FORMAT_CATEGORY) & DATEV_SEPARATOR         ' 3: Format category (21)
    Header = Header & Q & DATEV_FORMAT_NAME & Q & DATEV_SEPARATOR           ' 4: Format name
    Header = Header & "13" & DATEV_SEPARATOR                                ' 5: Format version (13 for v700)
    Header = Header & GeneratedOn & DATEV_SEPARATOR                         ' 6: Generated timestamp
    Header = Header & DATEV_SEPARATOR                                       ' 7: Imported (empty)
    Header = Header & Q & "RE" & Q & DATEV_SEPARATOR                        ' 8: Origin (RE=Rechnungswesen)
    Header = Header & Q & Q & DATEV_SEPARATOR                               ' 9: Exported by (empty)
    Header = Header & Q & Q & DATEV_SEPARATOR                               ' 10: Imported by (empty)
    Header = Header & Format$(Config.Beraternummer, "0") & DATEV_SEPARATOR  ' 11: Beraternummer
    Header = Header & Format$(Config.Mandantennummer, "0") & DATEV_SEPARATOR ' 12: Mandantennummer
    Header = Header & WJBeginn & DATEV_SEPARATOR                            ' 13: WJ-Beginn
    Header = Header & CStr(AccountLength) & DATEV_SEPARATOR                 ' 14: Sachkontenlaenge
    Header = Header & DateFrom & DATEV_SEPARATOR                            ' 15: Datum von
    Header = Header & DateTo & DATEV_SEPARATOR                              ' 16: Datum bis
    Header = Header & Q & "SimpliMed Export" & Q & DATEV_SEPARATOR          ' 17: Bezeichnung
    Header = Header & Q & Q & DATEV_SEPARATOR                               ' 18: Diktatkuerzel
    Header = Header & "1" & DATEV_SEPARATOR                                 ' 19: Buchungstyp (1=Fibu)
    Header = Header & "0" & DATEV_SEPARATOR                                 ' 20: Rechnungslegungszweck
    Header = Header & "0" & DATEV_SEPARATOR                                 ' 21: Festschreibung (0=keine)
    Header = Header & Q & DATEV_CURRENCY & Q & DATEV_SEPARATOR              ' 22: WKZ
    Header = Header & DATEV_SEPARATOR                                       ' 23: reserviert
    Header = Header & DATEV_SEPARATOR                                       ' 24: Derivatskennzeichen
    Header = Header & DATEV_SEPARATOR                                       ' 25: reserviert
    Header = Header & DATEV_SEPARATOR                                       ' 26: reserviert
    Header = Header & Q & Q & DATEV_SEPARATOR                               ' 27: SKR
    Header = Header & DATEV_SEPARATOR                                       ' 28: Branchenloesung-Id
    Header = Header & DATEV_SEPARATOR                                       ' 29: reserviert
    Header = Header & DATEV_SEPARATOR                                       ' 30: reserviert
    Header = Header & Q & Q                                                 ' 31: Anwendungsinformation

    BuildEXTFHeader = Header
End Function

Private Function BuildColumnHeaders() As String
    ' DATEV column headers - 125 fields for Buchungsstapel format v13
    ' Must match exactly the DATEV specification
    ' Split into multiple assignments to avoid VB6 line continuation limit (max ~25)

    Dim H As String

    ' Fields 1-20
    H = "Umsatz (ohne Soll/Haben-Kz);Soll/Haben-Kennzeichen;WKZ Umsatz;"
    H = H & "Kurs;Basis-Umsatz;WKZ Basis-Umsatz;Konto;Gegenkonto (ohne BU-Schluessel);"
    H = H & "BU-Schluessel;Belegdatum;Belegfeld 1;Belegfeld 2;Skonto;Buchungstext;"
    H = H & "Postensperre;Diverse Adressnummer;Geschaeftspartnerbank;Sachverhalt;"
    H = H & "Zinssperre;Beleglink;"

    ' Fields 21-40 (Beleginfo + KOST)
    H = H & "Beleginfo - Art 1;Beleginfo - Inhalt 1;"
    H = H & "Beleginfo - Art 2;Beleginfo - Inhalt 2;Beleginfo - Art 3;Beleginfo - Inhalt 3;"
    H = H & "Beleginfo - Art 4;Beleginfo - Inhalt 4;Beleginfo - Art 5;Beleginfo - Inhalt 5;"
    H = H & "Beleginfo - Art 6;Beleginfo - Inhalt 6;Beleginfo - Art 7;Beleginfo - Inhalt 7;"
    H = H & "Beleginfo - Art 8;Beleginfo - Inhalt 8;KOST1 - Kostenstelle;"
    H = H & "KOST2 - Kostenstelle;Kost-Menge;EU-Land u. UStID;"

    ' Fields 41-50
    H = H & "EU-Steuersatz;Abw. Versteuerungsart;Sachverhalt L+L;Funktionsergaenzung L+L;"
    H = H & "BU 49 Hauptfunktionstyp;BU 49 Hauptfunktionsnummer;BU 49 Funktionsergaenzung;"
    H = H & "Zusatzinformation - Art 1;Zusatzinformation- Inhalt 1;Zusatzinformation - Art 2;"

    ' Fields 51-70 (Zusatzinformation 2-11)
    H = H & "Zusatzinformation- Inhalt 2;Zusatzinformation - Art 3;Zusatzinformation- Inhalt 3;"
    H = H & "Zusatzinformation - Art 4;Zusatzinformation- Inhalt 4;Zusatzinformation - Art 5;"
    H = H & "Zusatzinformation- Inhalt 5;Zusatzinformation - Art 6;Zusatzinformation- Inhalt 6;"
    H = H & "Zusatzinformation - Art 7;Zusatzinformation- Inhalt 7;Zusatzinformation - Art 8;"
    H = H & "Zusatzinformation- Inhalt 8;Zusatzinformation - Art 9;Zusatzinformation- Inhalt 9;"
    H = H & "Zusatzinformation - Art 10;Zusatzinformation- Inhalt 10;Zusatzinformation - Art 11;"
    H = H & "Zusatzinformation- Inhalt 11;"

    ' Fields 71-88 (Zusatzinformation 12-20 + misc)
    H = H & "Zusatzinformation - Art 12;Zusatzinformation- Inhalt 12;"
    H = H & "Zusatzinformation - Art 13;Zusatzinformation- Inhalt 13;"
    H = H & "Zusatzinformation - Art 14;Zusatzinformation- Inhalt 14;"
    H = H & "Zusatzinformation - Art 15;Zusatzinformation- Inhalt 15;"
    H = H & "Zusatzinformation - Art 16;Zusatzinformation- Inhalt 16;"
    H = H & "Zusatzinformation - Art 17;Zusatzinformation- Inhalt 17;"
    H = H & "Zusatzinformation - Art 18;Zusatzinformation- Inhalt 18;"
    H = H & "Zusatzinformation - Art 19;Zusatzinformation- Inhalt 19;"
    H = H & "Zusatzinformation - Art 20;Zusatzinformation- Inhalt 20;"

    ' Fields 89-100
    H = H & "Stueck;Gewicht;Zahlweise;Forderungsart;Veranlagungsjahr;Zugeordnete Faelligkeit;"
    H = H & "Skontotyp;Auftragsnummer;Buchungstyp;USt-Schluessel (Anzahlungen);"
    H = H & "EU-Land (Anzahlungen);Sachverhalt L+L (Anzahlungen);"

    ' Fields 101-116
    H = H & "EU-Steuersatz (Anzahlungen);Erloeskonto (Anzahlungen);"
    H = H & "Herkunft-Kz;Buchungs GUID;KOST-Datum;SEPA-Mandatsreferenz;Skontosperre;"
    H = H & "Gesellschaftername;Beteiligtennummer;Identifikationsnummer;Zeichnernummer;"
    H = H & "Postensperre bis;Bezeichnung SoBil-Sachverhalt;Kennzeichen SoBil-Buchung;"
    H = H & "Festschreibung;Leistungsdatum;Datum Zuord. Steuerperiode;"

    ' Fields 117-125 (DATEV v700 Format v13)
    H = H & "Generalumkehr (Storno);Steuersatz;Land;"
    H = H & "Abrechnungsreferenz;BVV-Position;EU-Land u. UStID (Ursprungsland);"
    H = H & "EU-USt-IdNr (Ursprung);Sachverhalt Warenbewegung;Steuerschloessel Devisen"

    BuildColumnHeaders = H
End Function

'--------------------------------------------------------------------------------
' PRIVATE - Optimized CSV Data Line Builder (Array-based, proper edge cases)
'--------------------------------------------------------------------------------
Private Function BuildCSVDataLineOptimized(ByRef RST As ADODB.Recordset, _
                                           ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler

    Dim Fields(0 To 124) As String  ' 125 fields (0-124) per DATEV v700 format v13
    Dim FieldIdx As Integer
    Dim i As Integer

    ' Field values
    Dim BuDat As Date
    Dim BuDatVal As Variant
    Dim IdxNr As Long
    Dim BuTyp As Integer
    Dim GeKto As Integer
    Dim ManNr As Long
    Dim MitNr As Long
    Dim PatNr As Long
    Dim KtoSo As Long
    Dim KtoHa As Long
    Dim Steue As Single
    Dim Storn As Boolean
    Dim BLock As Boolean
    Dim Amount As Currency

    Dim GesBe As String
    Dim Kennz As String
    Dim StSch As String
    Dim ReStr As String
    Dim BelFe As String
    Dim BuStr As String
    Dim PaNum As String
    Dim Koste As String
    Dim BuGui As String
    Dim BeGui As String
    Dim DaNam As String
    Dim BeDaNam As String  ' BEDI-Dateiname fuer Beleginfo-Feld
    Dim Beleg As String
    Dim Komme As String
    Dim KSoSt As String
    Dim KHaSt As String
    Dim TmSt2 As String
    Dim BeVor As Boolean
    Dim PidNr As Long
    Dim DebNr As Long
    Dim BePatArt As String

    ' ========================================================================
    ' Step 1: Validate and extract date (required field)
    ' ========================================================================
    BuDatVal = RST.Fields("Datum").Value
    If IsNull(BuDatVal) Then
        BuildCSVDataLineOptimized = vbNullString
        Exit Function
    End If

    If Not IsDate(BuDatVal) Then
        BuildCSVDataLineOptimized = vbNullString
        Exit Function
    End If

    BuDat = CDate(BuDatVal)

    ' Validate date is reasonable (1900-2100)
    If BuDat < #1/1/1900# Or BuDat > #12/31/2100# Then
        BuildCSVDataLineOptimized = vbNullString
        Exit Function
    End If

    ' ========================================================================
    ' Step 2: Extract all field values with safe conversions
    ' ========================================================================
    If m_InvMod Then
        ' Debitoren: ID1 als Kennung, BuTyp=2 (immer Einnahme)
        IdxNr = SafeLong(RST.Fields("ID1").Value)
        BuTyp = 2
        GeKto = 0
        ManNr = SafeLong(RST.Fields("IDT").Value)
        MitNr = SafeLong(RST.Fields("IDM").Value)
    Else
        IdxNr = SafeLong(RST.Fields("ID0").Value)
        BuTyp = SafeInt(RST.Fields("IDA").Value)
        GeKto = SafeInt(RST.Fields("IDB").Value)
        ManNr = SafeLong(RST.Fields("IDT").Value)
        MitNr = SafeLong(RST.Fields("IDM").Value)
    End If
    Steue = SafeSingle(RST.Fields("Steuer").Value)
    Storn = SafeBool(RST.Fields("Storniert").Value)
    BLock = SafeBool(RST.Fields("Lock").Value)
    If m_InvMod Then
        ' qrySimReSu: Patientennummer liegt in ID0
        PatNr = SafeLong(GetFieldValue(RST, "ID0", "IDP"))
    Else
        ' qrySimBuSu: Patientennummer liegt in IDP
        PatNr = SafeLong(GetFieldValue(RST, "IDP", "ID0"))
    End If

    ' ========================================================================
    ' Step 3: Determine amount (with edge case handling)
    ' ========================================================================
    Amount = DetermineAmountValue(RST, BuTyp)

    ' Skip zero or invalid amounts
    If Amount < MIN_VALID_AMOUNT Then
        BuildCSVDataLineOptimized = vbNullString
        Exit Function
    End If

    ' Cap at maximum valid amount
    If Amount > MAX_VALID_AMOUNT Then
        Amount = MAX_VALID_AMOUNT
    End If

    ' Format amount with German decimal
    GesBe = FormatAmountGermanOptimized(Amount)

    ' Determine Soll/Haben
    Kennz = DetermineSollHabenValue(RST, BuTyp, Config.SwapDebitCredit)

    ' ========================================================================
    ' Step 4: Get accounts
    ' ========================================================================
    If m_InvMod Then
        ' Debitoren: Konto = Geldkonto (Bank), Gegenkonto = Erloskonto
        KtoSo = GetCashAccountNumber(GlGkB)
        KtoHa = GlSE2
        If GlLog = True Then SLogi "DATEV_BuEx: InvMod KtoSo=" & KtoSo & " (GlGkB=" & GlGkB & ") KtoHa=" & KtoHa
    ElseIf GlBuc = True Then
        KtoSo = SafeLong(RST.Fields("IDK").Value)
        KtoHa = GetCashAccountNumber(GeKto)
    Else
        KtoSo = SafeLong(RST.Fields("IDK").Value)
        KtoHa = SafeLong(RST.Fields("IDG").Value)
    End If

    ' Format account numbers
    KSoSt = FormatAccountNumber(KtoSo, Config.FourDigitAccounts)
    KHaSt = FormatAccountNumber(KtoHa, Config.FourDigitAccounts)

    ' Determine tax key
    StSch = GetTaxKey(Steue, Kennz)

    ' ========================================================================
    ' Step 5: Get document info and handle Beleglink
    ' ========================================================================
    Dim InvStr As String
    If m_InvMod Then
        ' Debitoren: kein Datei-Feld, PDF wird generiert aus RechNr
        DaNam = vbNullString
        InvStr = SafeString(RST.Fields("RechNr").Value)
        If Len(InvStr) > 0 Then
            DaNam = "Rechnung_Beleg_" & SanitizeFileName(InvStr) & ".pdf"
        End If
    Else
        DaNam = SanitizeTextField(SafeString(RST.Fields("Datei").Value), MAX_FILENAME_LENGTH)

        ' Generate default filename from invoice number if no document (NUR Einnahmen!)
        ' Wie in S_BuEx: BlgNa = "Rechnung_Beleg_" & RS125.Fields("RechNr").Value & ".pdf"
        ' Ausgaben ohne zugeordnete Datei: KEIN Standard-Dateiname, KEIN BEDI-Link
        If Len(DaNam) = 0 And BuTyp = 2 Then
            InvStr = SafeString(RST.Fields("RechNr").Value)
            If Len(InvStr) > 0 Then
                DaNam = "Rechnung_Beleg_" & SanitizeFileName(InvStr) & ".pdf"
            End If
        End If
    End If

    ' Get GUID for Beleglink
    BuGui = SafeString(RST.Fields("GuiID").Value)

    ' Store RechNr -> GUID mapping for revenue PDF renaming
    ' Used by AddGeneratedPDFsToZipList to rename Rechnung_Beleg_*.pdf to BEDI*.pdf
    If BuTyp = 2 And Len(BuGui) > 0 Then
        Dim RechNrKey As String
        RechNrKey = SanitizeFileName(SafeString(RST.Fields("RechNr").Value))
        If Len(RechNrKey) > 0 Then
            RegisterInvoiceGUID RechNrKey, BuGui
        End If
    End If

    BeVor = False
    If Len(BuGui) > 0 And Len(DaNam) > 0 Then
        BeGui = DATEV_FormatGUIDForXML(BuGui)
        ' Track document for deduplication (split postings)
        If Not IsDocumentAlreadyExported(DaNam) Then
            TrackExportedDocument DaNam, BuGui
        Else
            BeVor = True
        End If
    End If

    ' BEDI-Dateiname fuer Beleginfo-Feld berechnen
    ' Format: BEDI<GUID ohne Bindestriche>.<Dateiendung>
    ' DATEV_FormatGUIDForXML entfernt B/R/G Praefix fuer konsistente Benennung
    If Len(BuGui) > 0 And Len(DaNam) > 0 Then
        Dim CleanGUID As String
        CleanGUID = DATEV_FormatGUIDForXML(BuGui)
        BeDaNam = BELEGLINK_PREFIX & UCase$(Replace(CleanGUID, "-", vbNullString)) & GetFileExtension(DaNam)
    Else
        BeDaNam = DaNam
    End If

    ' ========================================================================
    ' Step 6: Get text fields with proper sanitization
    ' ========================================================================
    If m_InvMod Then
        ' Debitoren: Belegfeld1 = RechNr (keine Beleg-Spalte in qrySimReSu)
        ReStr = SanitizeTextField(SafeString(RST.Fields("RechNr").Value), MAX_BELEGFELD1_LENGTH)
        ' Beleginfo Art 2: RechNr als Kennung (max 20 Zeichen)
        Beleg = SafeString(RST.Fields("RechNr").Value)
        If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)
    Else
        ReStr = SanitizeTextField(SafeString(RST.Fields("RechNr").Value), MAX_BELEGFELD1_LENGTH)

        ' Beleginfo - Art 2: max 20 Zeichen laut DATEV-Spezifikation
        Beleg = SafeString(RST.Fields("IDKurz").Value)
        If Len(Beleg) > 0 And IsNumeric(Beleg) Then
            On Error Resume Next
            Beleg = Format$(CLng(Beleg), "00000000")
            On Error GoTo ErrHandler
        End If
        If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)
    End If

    ' DATEV-Regel: Beleginfo - Art 2 und Inhalt 2 muessen beide gefuellt oder beide leer sein
    ' BeDaNam = Inhalt 2 (BEDI-Dateiname), Beleg = Art 2
    If Len(BeDaNam) = 0 Then
        Beleg = vbNullString  ' Wenn kein Dokument, auch Art 2 leer
    ElseIf Len(Beleg) = 0 Then
        ' Fallback: IDX als Art 2 wenn Dokument vorhanden aber kein IDKurz
        Beleg = Format$(IdxNr, "00000000")
        If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)
    End If

    Komme = SanitizeTextField(SafeString(RST.Fields("Kommentar").Value), 60)

    ' Get patient number if configured
    PaNum = vbNullString
    If Config.IncludePatientNumber Then
        PidNr = PatNr
        If PidNr <= 0 Then
            On Error Resume Next
            TmSt2 = S_AdIdx(IdxNr, "IDP")
            On Error GoTo ErrHandler
            If Len(TmSt2) > 0 And IsNumeric(TmSt2) Then
                PidNr = CLng(TmSt2)
            End If
        End If

        If PidNr > 0 Then
            If Config.FourDigitAccounts Then
                PaNum = Format$(PidNr, "00000")
            Else
                PaNum = Format$(PidNr, "0000000")
            End If
        End If
    End If

    ' DATEV-konforme Debitorennummer (Patientennr + Basis)
    DebNr = 0
    If PidNr > 0 Then
        If Config.FourDigitAccounts Then
            DebNr = 10000 + PidNr
        Else
            DebNr = 1000000 + PidNr
        End If
    End If

    ' GlDeE: Replace account (Konto) with debtor number (invoices only)
    If Config.ReplaceAccountWithDebtor And m_InvMod Then
        If DebNr > 0 Then KSoSt = CStr(DebNr)
    End If

    ' Get cost center
    Koste = GetCostCenter(ManNr)

    ' Build booking text with sanitization
    BuStr = BuildBookingTextOptimized(RST, Config.IncludePatientName, Storn)

    ' Belegfeld 2 (storno marker)
    If Storn Then
        BelFe = "[STORNIERT]"
    Else
        BelFe = vbNullString
    End If

    ' ========================================================================
    ' Step 7: Build all 116 fields into array (fast)
    ' ========================================================================

    ' Field 1: Umsatz
    Fields(0) = GesBe

    ' Field 2: Soll/Haben-Kennzeichen
    Fields(1) = QuoteField(Kennz)

    ' Field 3: WKZ Umsatz (empty)
    Fields(2) = m_EmptyQuoted

    ' Field 4: Kurs (empty)
    Fields(3) = vbNullString

    ' Field 5: Basis-Umsatz (empty)
    Fields(4) = vbNullString

    ' Field 6: WKZ Basis-Umsatz (empty)
    Fields(5) = m_EmptyQuoted

    ' Field 7: Konto
    Fields(6) = KSoSt

    ' Field 8: Gegenkonto
    Fields(7) = KHaSt

    ' Field 9: BU-Schluessel
    Fields(8) = QuoteField(StSch)

    ' Field 10: Belegdatum (DDMM)
    Fields(9) = Format$(BuDat, "ddmm")

    ' Field 11: Belegfeld 1
    Fields(10) = QuoteField(ReStr)

    ' Field 12: Belegfeld 2
    Fields(11) = QuoteField(BelFe)

    ' Field 13: Skonto (empty)
    Fields(12) = vbNullString

    ' Field 14: Buchungstext
    Fields(13) = QuoteField(BuStr)

    ' Field 15: Postensperre (empty, Text)
    Fields(14) = m_EmptyQuoted

    ' Field 16: Diverse Adressnummer
    Fields(15) = QuoteField(PaNum)

    ' Field 17: Geschaeftspartnerbank (empty, Text)
    Fields(16) = m_EmptyQuoted

    ' Field 18: Sachverhalt (empty, Text)
    Fields(17) = m_EmptyQuoted

    ' Field 19: Zinssperre (empty, Text)
    Fields(18) = m_EmptyQuoted

    ' Field 20: Beleglink
    ' BEDI nur setzen wenn: Dateiname vorhanden, nicht bereits exportiert, GUID vorhanden,
    ' UND Beleg gueltig (bei Ausgaben: Datei existiert in GlBPf)
    ' Format: "BEDI" + GUID (32 hex, uppercase, no dashes) via DATEV_CreateBeleglink
    If Len(DaNam) > 0 And Not BeVor And Len(BuGui) > 0 And IsExpenseDocumentValid(BuGui) Then
        Fields(19) = QuoteField(DATEV_CreateBeleglink(BuGui))
    Else
        Fields(19) = m_EmptyQuoted
    End If

    ' Fields 21-36: Beleginfo 1-8 (Art + Inhalt pairs)
    If DebNr > 0 Then
        Fields(20) = QuoteField(BELEGINFO_DEBITORNR_ART) ' Art 1: Debitorennr
        Fields(21) = QuoteField(CStr(DebNr))             ' Inhalt 1: Debitorennr
    Else
        Fields(20) = m_EmptyQuoted                       ' Art 1 (leer)
        Fields(21) = m_EmptyQuoted                       ' Inhalt 1 (leer)
    End If
    Fields(22) = QuoteField(Beleg)                     ' Art 2
    Fields(23) = QuoteField(BeDaNam)                   ' Inhalt 2 (BEDI-Dateiname)
    BePatArt = vbNullString
    If Len(PaNum) > 0 Then
        BePatArt = BELEGINFO_PATIENT_ART
    End If
    Fields(24) = QuoteField(BePatArt)                  ' Art 3 (Patientennummer)
    Fields(25) = QuoteField(PaNum)                     ' Inhalt 3 (Patientennummer)
    For i = 26 To 35
        Fields(i) = m_EmptyQuoted
    Next i

    ' Field 37: KOST1
    Fields(36) = QuoteField(Koste)

    ' Field 38: KOST2 (empty)
    Fields(37) = m_EmptyQuoted

    ' Field 39: Kost-Menge (empty)
    Fields(38) = vbNullString

    ' Field 40: EU-Land u. UStID (empty)
    Fields(39) = m_EmptyQuoted

    ' Field 41: EU-Steuersatz (empty)
    Fields(40) = vbNullString

    ' Field 42: Abw. Versteuerungsart (empty)
    Fields(41) = m_EmptyQuoted

    ' Field 43-44: Sachverhalt/Funktionsergaenzung L+L (empty)
    Fields(42) = vbNullString
    Fields(43) = vbNullString

    ' Fields 45-47: BU 49 fields (empty)
    Fields(44) = vbNullString
    Fields(45) = vbNullString
    Fields(46) = vbNullString

    ' Fields 48-87: Zusatzinformation 1-20 (40 fields)
    If Len(Komme) > 0 Then
        Fields(47) = QuoteField("KOMMENTAR")  ' Art 1
        Fields(48) = QuoteField(Komme)        ' Inhalt 1 (Kommentar)
    Else
        Fields(47) = m_EmptyQuoted            ' Art 1 (leer)
        Fields(48) = m_EmptyQuoted            ' Inhalt 1 (leer)
    End If
    For i = 49 To 86
        Fields(i) = m_EmptyQuoted
    Next i

    ' Fields 88-90: Stueck, Gewicht, Zahlweise
    Fields(87) = vbNullString  ' Stueck (Numeric)
    Fields(88) = vbNullString  ' Gewicht (Numeric)
    Fields(89) = m_EmptyQuoted ' Zahlweise (Text)

    ' Field 91: Forderungsart (empty)
    Fields(90) = m_EmptyQuoted

    ' Field 92: Veranlagungsjahr
    Fields(91) = Format$(BuDat, "yyyy")

    ' Fields 93-94: Zugeordnete Faelligkeit, Skontotyp (empty)
    Fields(92) = vbNullString
    Fields(93) = vbNullString

    ' Fields 95-96: Auftragsnummer, Buchungstyp (empty)
    Fields(94) = m_EmptyQuoted
    Fields(95) = m_EmptyQuoted

    ' Field 97: USt-Schluessel Anzahlungen (empty)
    Fields(96) = vbNullString

    ' Field 98: EU-Land Anzahlungen (empty)
    Fields(97) = m_EmptyQuoted

    ' Fields 99-101: Sachverhalt, EU-Steuersatz, Erloeskonto Anzahlungen (empty)
    Fields(98) = vbNullString
    Fields(99) = vbNullString
    Fields(100) = vbNullString

    ' Field 102: Herkunft-Kz
    Fields(101) = QuoteField("SM")

    ' Field 103: Buchungs GUID
    Fields(102) = QuoteField(BuGui)

    ' Field 104: KOST-Datum (empty)
    Fields(103) = vbNullString

    ' Field 105: SEPA-Mandatsreferenz (empty)
    Fields(104) = m_EmptyQuoted

    ' Field 106: Skontosperre
    Fields(105) = QuoteField("0")

    ' Fields 107-111: Empty quoted fields
    For i = 106 To 110
        Fields(i) = m_EmptyQuoted
    Next i

    ' Field 112-113: SoBil fields (empty)
    Fields(111) = m_EmptyQuoted
    Fields(112) = m_EmptyQuoted

    ' Field 114: Festschreibung
    If BLock Then
        Fields(113) = QuoteField("1")
    Else
        Fields(113) = QuoteField("0")
    End If

    ' Fields 115-116: Leistungsdatum, Datum Zuord. Steuerperiode (empty)
    Fields(114) = vbNullString
    Fields(115) = vbNullString

    ' Fields 117-125: DATEV v700 Format v13 additional fields (all empty)
    ' Textfelder mit "" quotieren, numerische Felder leer lassen
    ' HINWEIS: DATEV Prueftool gibt Warnungen fuer leere Textfelder - das ist normal
    Fields(116) = m_EmptyQuoted ' 117: Generalumkehr (Storno) - Text
    Fields(117) = vbNullString  ' 118: Steuersatz - Numeric
    Fields(118) = m_EmptyQuoted ' 119: Land - Text
    Fields(119) = m_EmptyQuoted ' 120: Abrechnungsreferenz - Text
    Fields(120) = vbNullString  ' 121: BVV-Position - Numeric
    Fields(121) = m_EmptyQuoted ' 122: EU-Land u. UStID (Ursprungsland) - Text
    Fields(122) = m_EmptyQuoted ' 123: EU-USt-IdNr (Ursprung) - Text
    Fields(123) = vbNullString  ' 124: Sachverhalt Warenbewegung - Numeric
    Fields(124) = vbNullString  ' 125: Steuerschloessel Devisen - Numeric

    ' ========================================================================
    ' Step 8: Join all fields with separator (single operation)
    ' ========================================================================
    BuildCSVDataLineOptimized = Join(Fields, m_Sep)
    Exit Function

ErrHandler:
    LogError "BuildCSVDataLineOptimized", Err.Number, Err.Description
    BuildCSVDataLineOptimized = vbNullString
End Function

'--------------------------------------------------------------------------------
' PRIVATE - Text Sanitization for DATEV CSV
'--------------------------------------------------------------------------------
Private Function SanitizeTextField(ByVal Text As String, ByVal MaxLen As Integer) As String
    ' Sanitize text field for DATEV CSV:
    ' - Remove/replace problematic characters
    ' - Truncate to max length
    ' - Handle encoding issues
    ' Optimized: uses array-based character filtering

    Dim Result As String
    Dim TextLen As Long
    Dim CharArray() As Byte
    Dim OutArray() As Byte
    Dim i As Long
    Dim j As Long
    Dim b As Byte

    If Len(Text) = 0 Then
        SanitizeTextField = vbNullString
        Exit Function
    End If

    Result = Text

    ' Remove line breaks (CR, LF, CRLF) - Replace is fast for these
    Result = Replace(Result, vbCrLf, " ")
    Result = Replace(Result, vbCr, " ")
    Result = Replace(Result, vbLf, " ")

    ' Remove tabs
    Result = Replace(Result, vbTab, " ")

    ' Escape double quotes by doubling them (CSV standard)
    If Len(m_Q) > 0 Then
        Result = Replace(Result, m_Q, m_Q & m_Q)
    Else
        Result = Replace(Result, Chr$(34), Chr$(34) & Chr$(34))
    End If

    ' Remove semicolons (DATEV separator) - replace with comma
    Result = Replace(Result, ";", ",")

    ' Remove control characters (ASCII 0-31) using byte array for speed
    TextLen = Len(Result)
    If TextLen > 0 Then
        ' Convert to byte array (Unicode = 2 bytes per char)
        CharArray = Result
        ReDim OutArray(0 To UBound(CharArray))
        j = 0

        ' Process in pairs (Unicode)
        For i = 0 To UBound(CharArray) - 1 Step 2
            b = CharArray(i)  ' Low byte (ASCII value for Latin chars)
            ' Keep if printable (>= 32) or high byte is non-zero (extended Unicode)
            If b >= 32 Or CharArray(i + 1) <> 0 Then
                OutArray(j) = CharArray(i)
                OutArray(j + 1) = CharArray(i + 1)
                j = j + 2
            End If
        Next i

        ' Convert back to string
        If j > 0 Then
            ReDim Preserve OutArray(0 To j - 1)
            Result = OutArray
        Else
            Result = vbNullString
        End If
    End If

    ' Collapse multiple spaces (limited iterations to prevent infinite loop)
    Dim LoopCount As Integer
    LoopCount = 0
    Do While InStr(Result, "  ") > 0 And LoopCount < 100
        Result = Replace(Result, "  ", " ")
        LoopCount = LoopCount + 1
    Loop

    ' Trim
    Result = Trim$(Result)

    ' Truncate if needed
    If Len(Result) > MaxLen Then
        Result = Left$(Result, MaxLen)
    End If

    SanitizeTextField = Result
End Function

Private Function SanitizeFileName(ByVal FileName As String) As String
    ' DATEV Unternehmen Online konforme Dateinamen
    ' - Erlaubt: A-Z, a-z, 0-9, _ (Unterstrich), - (Bindestrich), . (Punkt)
    ' - Umlaute werden ersetzt: oe->ae, oe->oe, oe->ue, oe->ss
    ' - Sonderzeichen werden durch Unterstrich ersetzt
    ' - Max 46 Zeichen (DATEV-Limit)

    Dim Result As String

    If Len(FileName) = 0 Then
        SanitizeFileName = vbNullString
        Exit Function
    End If

    Result = FileName

    ' Replace German umlauts (DATEV requirement)
    Result = Replace(Result, "oe", "ae")
    Result = Replace(Result, "oe", "oe")
    Result = Replace(Result, "oe", "ue")
    Result = Replace(Result, "oe", "Ae")
    Result = Replace(Result, "oe", "Oe")
    Result = Replace(Result, "oe", "Ue")
    Result = Replace(Result, "oe", "ss")

    ' Replace spaces with underscore
    Result = Replace(Result, " ", "_")

    ' Replace invalid filename characters with underscore
    Result = Replace(Result, "\", "_")
    Result = Replace(Result, "/", "_")
    Result = Replace(Result, ":", "_")
    Result = Replace(Result, "*", "_")
    Result = Replace(Result, "?", "_")
    Result = Replace(Result, Chr$(34), "_")  ' Double quote
    Result = Replace(Result, "<", "_")
    Result = Replace(Result, ">", "_")
    Result = Replace(Result, "|", "_")

    ' Replace additional DATEV-problematic characters
    Result = Replace(Result, "#", "_")
    Result = Replace(Result, "%", "_")
    Result = Replace(Result, "&", "_")
    Result = Replace(Result, "{", "_")
    Result = Replace(Result, "}", "_")
    Result = Replace(Result, "$", "_")
    Result = Replace(Result, "!", "_")
    Result = Replace(Result, "'", "_")
    Result = Replace(Result, "@", "_")
    Result = Replace(Result, "+", "_")
    Result = Replace(Result, "=", "_")
    Result = Replace(Result, ";", "_")
    Result = Replace(Result, ",", "_")

    ' Remove control characters using byte array
    Dim CharArray() As Byte
    Dim OutArray() As Byte
    Dim i As Long
    Dim j As Long
    Dim b As Byte

    If Len(Result) > 0 Then
        CharArray = Result
        ReDim OutArray(0 To UBound(CharArray))
        j = 0

        For i = 0 To UBound(CharArray) - 1 Step 2
            b = CharArray(i)
            If b >= 32 Or CharArray(i + 1) <> 0 Then
                OutArray(j) = CharArray(i)
                OutArray(j + 1) = CharArray(i + 1)
                j = j + 2
            End If
        Next i

        If j > 0 Then
            ReDim Preserve OutArray(0 To j - 1)
            Result = OutArray
        Else
            Result = vbNullString
        End If
    End If

    ' Remove multiple consecutive underscores
    Do While InStr(Result, "__") > 0
        Result = Replace(Result, "__", "_")
    Loop

    ' Remove leading/trailing underscores
    Do While Left$(Result, 1) = "_"
        Result = Mid$(Result, 2)
    Loop
    Do While Right$(Result, 1) = "_"
        Result = Left$(Result, Len(Result) - 1)
    Loop

    ' Truncate to safe length (MAX_FILENAME_LENGTH or 46 for DATEV)
    ' IMPORTANT: Preserve file extension when truncating!
    If Len(Result) > MAX_FILENAME_LENGTH Then
        Dim ExtPos As Long
        Dim FileExt As String
        Dim BaseName As String

        ' Find last dot for file extension
        ExtPos = InStrRev(Result, ".")
        If ExtPos > 0 And (Len(Result) - ExtPos) <= 5 Then  ' Extension max 5 chars (e.g., .jpeg, .tiff)
            FileExt = Mid$(Result, ExtPos)  ' Includes the dot
            BaseName = Left$(Result, ExtPos - 1)
            ' Truncate base name, keeping extension
            If Len(BaseName) + Len(FileExt) > MAX_FILENAME_LENGTH Then
                BaseName = Left$(BaseName, MAX_FILENAME_LENGTH - Len(FileExt))
            End If
            Result = BaseName & FileExt
        Else
            ' No valid extension found, just truncate
            Result = Left$(Result, MAX_FILENAME_LENGTH)
        End If
    End If

    SanitizeFileName = Result
End Function

Private Function QuoteField(ByVal Value As String) As String
    ' Wrap field value in quotes
    QuoteField = m_Q & Value & m_Q
End Function

'--------------------------------------------------------------------------------
' PRIVATE - Optimized Amount Handling
'--------------------------------------------------------------------------------
Private Function DetermineAmountValue(ByRef RST As ADODB.Recordset, _
                                      ByVal BuTyp As Integer) As Currency
    ' Returns amount as Currency (always positive)
    ' Handles edge cases: null, negative, overflow

    Dim Amount As Currency
    Dim TempVal As Variant

    Amount = 0

    On Error Resume Next

    ' Debitoren-Modus: Betrag-Feld aus qrySimReSu (nur Einnahmen)
    If m_InvMod Then
        TempVal = RST.Fields("Betrag").Value
        If Not IsNull(TempVal) And IsNumeric(TempVal) Then
            Amount = Abs(CCur(TempVal))
        End If
        On Error GoTo 0
        DetermineAmountValue = Amount
        Exit Function
    End If

    If GlBuc = True Then
        ' Simple bookkeeping
        Select Case BuTyp
        Case 1 ' Ausgabe
            TempVal = RST.Fields("Ausgabe").Value
            If Not IsNull(TempVal) And IsNumeric(TempVal) Then
                Amount = Abs(CCur(TempVal))
            Else
                TempVal = RST.Fields("Einnahme").Value
                If Not IsNull(TempVal) And IsNumeric(TempVal) Then
                    Amount = Abs(CCur(TempVal))
                End If
            End If
        Case 2 ' Einnahme
            TempVal = RST.Fields("Einnahme").Value
            If Not IsNull(TempVal) And IsNumeric(TempVal) Then
                Amount = Abs(CCur(TempVal))
            Else
                TempVal = RST.Fields("Ausgabe").Value
                If Not IsNull(TempVal) And IsNumeric(TempVal) Then
                    Amount = Abs(CCur(TempVal))
                End If
            End If
        Case Else
            TempVal = RST.Fields("Einnahme").Value
            If Not IsNull(TempVal) And IsNumeric(TempVal) And CDbl(TempVal) > 0 Then
                Amount = Abs(CCur(TempVal))
            Else
                TempVal = RST.Fields("Ausgabe").Value
                If Not IsNull(TempVal) And IsNumeric(TempVal) And CDbl(TempVal) > 0 Then
                    Amount = Abs(CCur(TempVal))
                End If
            End If
        End Select
    Else
        ' Double-entry bookkeeping - Feld "Betrag" existiert nicht, verwende Einnahme/Ausgabe
        TempVal = RST.Fields("Einnahme").Value
        If Not IsNull(TempVal) And IsNumeric(TempVal) And CDbl(TempVal) > 0 Then
            Amount = Abs(CCur(TempVal))
        Else
            TempVal = RST.Fields("Ausgabe").Value
            If Not IsNull(TempVal) And IsNumeric(TempVal) And CDbl(TempVal) > 0 Then
                Amount = Abs(CCur(TempVal))
            End If
        End If
    End If

    On Error GoTo 0

    DetermineAmountValue = Amount
End Function

Private Function DetermineSollHabenValue(ByRef RST As ADODB.Recordset, _
                                         ByVal BuTyp As Integer, _
                                         ByVal SwapDebitCredit As Boolean) As String
    ' Returns "S" or "H" based on booking type and swap setting

    Dim Result As String
    Dim TempVal As Variant

    On Error Resume Next

    ' Debitoren-Modus: immer Einnahme (Revenue = Soll ohne Tausch)
    If m_InvMod Then
        Result = IIf(SwapDebitCredit, "H", "S")
        On Error GoTo 0
        DetermineSollHabenValue = Result
        Exit Function
    End If

    If GlBuc = True Then
        ' Simple bookkeeping
        Select Case BuTyp
        Case 1 ' Ausgabe
            Result = IIf(SwapDebitCredit, "S", "H")
        Case 2 ' Einnahme
            Result = IIf(SwapDebitCredit, "H", "S")
        Case Else
            TempVal = RST.Fields("Einnahme").Value
            If Not IsNull(TempVal) And IsNumeric(TempVal) And CDbl(TempVal) > 0 Then
                Result = IIf(SwapDebitCredit, "H", "S")
            Else
                Result = IIf(SwapDebitCredit, "S", "H")
            End If
        End Select
    Else
        ' Double-entry bookkeeping
        Result = IIf(SwapDebitCredit, "S", "H")
    End If

    On Error GoTo 0

    DetermineSollHabenValue = Result
End Function

Private Function FormatAmountGermanOptimized(ByVal Amount As Currency) As String
    ' Format amount for DATEV: German decimal (comma), 2 decimal places
    ' No thousands separator, comma as decimal separator

    Dim Formatted As String

    ' Use simple Format with fixed 2 decimal places (no thousands separator)
    Formatted = Format$(Amount, "0.00")

    ' Replace dot with comma for German decimal format
    Formatted = Replace(Formatted, ".", ",")

    FormatAmountGermanOptimized = Formatted
End Function

Private Function BuildBookingTextOptimized(ByRef RST As ADODB.Recordset, _
                                           ByVal IncludePatientName As Boolean, _
                                           ByVal Storn As Boolean) As String
    ' Build booking text with proper sanitization

    Dim BuStr As String
    Dim PaStr As String
    Dim BuTex As String
    Dim RcNr As String

    On Error Resume Next

    ' Debitoren-Modus: IDKurz (Patient) + RechNr als Buchungstext
    If m_InvMod Then
        PaStr = SafeString(RST.Fields("IDKurz").Value)
        RcNr = SafeString(RST.Fields("RechNr").Value)
        If Len(PaStr) > 0 And Len(RcNr) > 0 Then
            BuStr = PaStr & " Rech." & RcNr
        ElseIf Len(PaStr) > 0 Then
            BuStr = PaStr
        ElseIf Len(RcNr) > 0 Then
            BuStr = "Rechnung " & RcNr
        Else
            BuStr = "Rechnungsexport"
        End If
        If Storn Then BuStr = "[STORNIERT] " & BuStr
        BuStr = SanitizeTextField(BuStr, MAX_BOOKING_TEXT_LENGTH)
        On Error GoTo 0
        BuildBookingTextOptimized = BuStr
        Exit Function
    End If

    ' Get patient name if available
    PaStr = vbNullString
    If IncludePatientName Then
        PaStr = SafeString(RST.Fields("Patient").Value)
    End If

    ' Get booking text
    BuTex = SafeString(RST.Fields("Buchtext").Value)

    On Error GoTo 0

    ' Build combined text
    If Len(PaStr) > 0 Then
        If Len(BuTex) > 0 Then
            BuStr = PaStr & " " & BuTex
        Else
            BuStr = PaStr
        End If
    Else
        BuStr = BuTex
    End If

    ' Add storno prefix if applicable
    If Storn Then
        BuStr = "[STORNIERT] " & BuStr
    End If

    ' Sanitize and truncate
    BuStr = SanitizeTextField(BuStr, MAX_BOOKING_TEXT_LENGTH)

    BuildBookingTextOptimized = BuStr
End Function

'--------------------------------------------------------------------------------
' PRIVATE - Legacy BuildCSVDataLine (kept for compatibility, delegates to optimized)
'--------------------------------------------------------------------------------
Private Function BuildCSVDataLine(ByRef RST As ADODB.Recordset, _
                                  ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler

    Dim Line As String
    Dim Q As String
    Dim Sep As String

    ' Field values
    Dim BuDat As Date
    Dim IdxNr As Long
    Dim BuTyp As Integer
    Dim GeKto As Integer
    Dim ManNr As Long
    Dim MitNr As Long
    Dim PatNr As Long
    Dim KtoSo As Long
    Dim KtoHa As Long
    Dim Steue As Single
    Dim Storn As Boolean
    Dim BLock As Boolean

    Dim GesBe As String
    Dim Kennz As String
    Dim StSch As String
    Dim ReStr As String
    Dim BelFe As String
    Dim BuStr As String
    Dim PaNum As String
    Dim Koste As String
    Dim BuGui As String
    Dim BeGui As String
    Dim DaNam As String
    Dim BeDaNam As String  ' BEDI-Dateiname fuer Beleginfo-Feld
    Dim Beleg As String
    Dim Komme As String
    Dim KSoSt As String
    Dim KHaSt As String
    Dim PaStr As String
    Dim BuTex As String
    Dim TmSt1 As String
    Dim TmSt2 As String
    Dim BePatArt As String

    Dim AktZa As Integer
    Dim BeVor As Boolean
    Dim PidNr As Long
    Dim DebNr As Long

    Q = Chr$(34)
    Sep = DATEV_SEPARATOR

    ' Validate record has date
    If IsNull(RST.Fields("Datum").Value) Then
        BuildCSVDataLine = vbNullString
        Exit Function
    End If

    BuDat = CDate(RST.Fields("Datum").Value)

    ' Extract basic field values
    IdxNr = SafeLong(RST.Fields("ID0").Value)
    BuTyp = SafeInt(RST.Fields("IDA").Value)
    GeKto = SafeInt(RST.Fields("IDB").Value)
    ManNr = SafeLong(RST.Fields("IDT").Value)
    MitNr = SafeLong(RST.Fields("IDM").Value)
    Steue = SafeSingle(RST.Fields("Steuer").Value)
    Storn = SafeBool(RST.Fields("Storniert").Value)
    BLock = SafeBool(RST.Fields("Lock").Value)
    If m_InvMod Then
        ' qrySimReSu: Patientennummer liegt in ID0
        PatNr = SafeLong(GetFieldValue(RST, "ID0", "IDP"))
    Else
        ' qrySimBuSu: Patientennummer liegt in IDP
        PatNr = SafeLong(GetFieldValue(RST, "IDP", "ID0"))
    End If

    ' Get document filename if present
    DaNam = SafeString(RST.Fields("Datei").Value)
    If Len(DaNam) > MAX_FILENAME_LENGTH Then
        Dim DaExt As String
        DaExt = Right$(DaNam, 4)
        DaNam = Left$(DaNam, MAX_FILENAME_LENGTH - 4) & DaExt
    End If

    ' Generate default filename from invoice number if no document (NUR Einnahmen!)
    ' Wie in S_BuEx: BlgNa = "Rechnung_Beleg_" & RS125.Fields("RechNr").Value & ".pdf"
    ' Ausgaben ohne zugeordnete Datei: KEIN Standard-Dateiname, KEIN BEDI-Link
    If DaNam = vbNullString And BuTyp = 2 Then
        Dim InvStr As String
        InvStr = SafeString(RST.Fields("RechNr").Value)
        If Len(InvStr) > 0 Then
            DaNam = "Rechnung_Beleg_" & SanitizeFileName(InvStr) & ".pdf"
        End If
    End If

    ' Get GUID for Beleglink
    BuGui = SafeString(RST.Fields("GuiID").Value)

    ' Store RechNr -> GUID mapping for revenue PDF renaming
    ' Used by AddGeneratedPDFsToZipList to rename Rechnung_Beleg_*.pdf to BEDI*.pdf
    If BuTyp = 2 And Len(BuGui) > 0 Then
        Dim RechNrKey As String
        RechNrKey = SanitizeFileName(SafeString(RST.Fields("RechNr").Value))
        If Len(RechNrKey) > 0 Then
            RegisterInvoiceGUID RechNrKey, BuGui
        End If
    End If

    If Len(BuGui) > 0 And Len(DaNam) > 0 Then
        BeGui = DATEV_FormatGUIDForXML(BuGui)
        ' Track document for deduplication
        If Not IsDocumentAlreadyExported(DaNam) Then
            TrackExportedDocument DaNam, BuGui
            BeVor = False
        Else
            BeVor = True ' Document already exported (split posting)
        End If
    End If

    ' BEDI-Dateiname fuer Beleginfo-Feld berechnen
    ' Format: BEDI<GUID ohne Bindestriche>.<Dateiendung>
    ' DATEV_FormatGUIDForXML entfernt B/R/G Praefix fuer konsistente Benennung
    If Len(BuGui) > 0 And Len(DaNam) > 0 Then
        Dim CleanGUID As String
        CleanGUID = DATEV_FormatGUIDForXML(BuGui)
        BeDaNam = BELEGLINK_PREFIX & UCase$(Replace(CleanGUID, "-", vbNullString)) & GetFileExtension(DaNam)
    Else
        BeDaNam = DaNam
    End If

    ' Get accounts - handle simple vs double-entry bookkeeping
    If GlBuc = True Then
        KtoSo = SafeLong(RST.Fields("IDK").Value)
        KtoHa = GetCashAccountNumber(GeKto)
    Else
        KtoSo = SafeLong(RST.Fields("IDK").Value)
        KtoHa = SafeLong(RST.Fields("IDG").Value)
    End If

    ' Determine amount and Soll/Haben based on booking type
    GesBe = DetermineAmount(RST, BuTyp, Config)
    Kennz = DetermineSollHaben(RST, BuTyp, Config)

    ' Skip if amount is zero
    If GesBe = "0,00" Or GesBe = vbNullString Then
        BuildCSVDataLine = vbNullString
        Exit Function
    End If

    ' Determine tax key (BU-Schluessel)
    StSch = GetTaxKey(Steue, Kennz)

    ' Format account numbers
    KSoSt = FormatAccountNumber(KtoSo, Config.FourDigitAccounts)
    KHaSt = FormatAccountNumber(KtoHa, Config.FourDigitAccounts)

    ' Get invoice number (Belegfeld 1) - RechNr from qrySimBuSu
    ReStr = SafeString(RST.Fields("RechNr").Value)
    If Len(ReStr) > MAX_BELEGFELD1_LENGTH Then
        ReStr = Left$(ReStr, MAX_BELEGFELD1_LENGTH)
    End If

    ' Get booking number for Beleginfo - Art 2: max 20 Zeichen laut DATEV-Spezifikation
    Beleg = SafeString(RST.Fields("IDKurz").Value)
    If Len(Beleg) > 0 And IsNumeric(Beleg) Then
        Beleg = Format$(CLng(Beleg), "00000000")
    End If
    If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)

    ' DATEV-Regel: Beleginfo - Art 2 und Inhalt 2 muessen beide gefuellt oder beide leer sein
    ' BeDaNam = Inhalt 2 (BEDI-Dateiname), Beleg = Art 2
    If Len(BeDaNam) = 0 Then
        Beleg = vbNullString  ' Wenn kein Dokument, auch Art 2 leer
    ElseIf Len(Beleg) = 0 Then
        ' Fallback: IDX als Art 2 wenn Dokument vorhanden aber kein IDKurz
        Beleg = Format$(IdxNr, "00000000")
        If Len(Beleg) > 20 Then Beleg = Left$(Beleg, 20)
    End If

    ' Get comment
    Komme = SafeString(RST.Fields("Kommentar").Value)
    If Len(Komme) > 0 And IsNumeric(Komme) Then
        Komme = Format$(CLng(Komme), "00000000")
    End If

    ' Get patient number if configured
    PaNum = vbNullString
    If Config.IncludePatientNumber Then
        PidNr = PatNr
        If PidNr <= 0 Then
            TmSt2 = S_AdIdx(IdxNr, "IDP")
            If Len(TmSt2) > 0 And IsNumeric(TmSt2) Then
                PidNr = CLng(TmSt2)
            End If
        End If

        If PidNr > 0 Then
            If Config.FourDigitAccounts Then
                PaNum = Format$(PidNr, "00000")
            Else
                PaNum = Format$(PidNr, "0000000")
            End If
        End If
    End If

    ' DATEV-konforme Debitorennummer (Patientennr + Basis)
    DebNr = 0
    If PidNr > 0 Then
        If Config.FourDigitAccounts Then
            DebNr = 10000 + PidNr
        Else
            DebNr = 1000000 + PidNr
        End If
    End If

    ' GlDeE: Replace account (Konto) with debtor number (invoices only)
    If Config.ReplaceAccountWithDebtor And m_InvMod Then
        If DebNr > 0 Then KSoSt = CStr(DebNr)
    End If

    ' Get cost center
    Koste = GetCostCenter(ManNr)

    ' Build booking text
    BuStr = BuildBookingText(RST, Config, Storn)

    ' Build Belegfeld 2 (storniert marker or space)
    If Storn Then
        BelFe = "[STORNIERT]"
    Else
        BelFe = vbNullString
    End If

    ' ========================================================================
    ' Build CSV line - 116 fields
    ' ========================================================================
    Line = vbNullString

    ' Field 1: Umsatz (Amount without sign, German decimal format)
    Line = Line & GesBe & Sep

    ' Field 2: Soll/Haben-Kennzeichen
    Line = Line & Q & Kennz & Q & Sep

    ' Field 3: WKZ Umsatz (currency)
    Line = Line & Q & Q & Sep

    ' Field 4: Kurs (exchange rate)
    Line = Line & Sep

    ' Field 5: Basis-Umsatz
    Line = Line & Sep

    ' Field 6: WKZ Basis-Umsatz
    Line = Line & Q & Q & Sep

    ' Field 7: Konto (debit account)
    Line = Line & KSoSt & Sep

    ' Field 8: Gegenkonto (credit account)
    Line = Line & KHaSt & Sep

    ' Field 9: BU-Schluessel (tax key)
    Line = Line & Q & StSch & Q & Sep

    ' Field 10: Belegdatum (DDMM format)
    Line = Line & Format$(BuDat, "ddmm") & Sep

    ' Field 11: Belegfeld 1 (invoice number)
    Line = Line & Q & ReStr & Q & Sep

    ' Field 12: Belegfeld 2
    Line = Line & Q & BelFe & Q & Sep

    ' Field 13: Skonto
    Line = Line & Sep

    ' Field 14: Buchungstext (max 60 chars)
    Line = Line & Q & BuStr & Q & Sep

    ' Field 15: Postensperre (Text)
    Line = Line & Q & Q & Sep

    ' Field 16: Diverse Adressnummer (patient number)
    Line = Line & Q & PaNum & Q & Sep

    ' Field 17: Geschaeftspartnerbank (Text)
    Line = Line & Q & Q & Sep

    ' Field 18: Sachverhalt (Text)
    Line = Line & Q & Q & Sep

    ' Field 19: Zinssperre (Text)
    Line = Line & Q & Q & Sep

    ' Field 20: Beleglink (BEDI + GUID)
    ' BEDI nur setzen wenn Beleg gueltig (bei Ausgaben: Datei existiert in GlBPf)
    ' Format: "BEDI" + GUID (32 hex, uppercase, no dashes) via DATEV_CreateBeleglink
    If Len(DaNam) > 0 And Not BeVor And Len(BuGui) > 0 And IsExpenseDocumentValid(BuGui) Then
        Line = Line & Q & DATEV_CreateBeleglink(BuGui) & Q & Sep
    Else
        Line = Line & Q & Q & Sep
    End If

    ' Fields 21-36: Beleginfo 1-8 (Art + Inhalt pairs)
    If DebNr > 0 Then
        Line = Line & Q & BELEGINFO_DEBITORNR_ART & Q & Sep  ' Beleginfo-Art 1: Debitorennr
        Line = Line & Q & CStr(DebNr) & Q & Sep              ' Beleginfo-Inhalt 1: Debitorennr
    Else
        Line = Line & Q & Q & Sep                            ' Beleginfo-Art 1 (leer)
        Line = Line & Q & Q & Sep                            ' Beleginfo-Inhalt 1 (leer)
    End If
    Line = Line & Q & Beleg & Q & Sep                     ' Beleginfo-Art 2 (Belegnummer)
    Line = Line & Q & BeDaNam & Q & Sep                   ' Beleginfo-Inhalt 2 (BEDI-Dateiname)
    BePatArt = vbNullString
    If Len(PaNum) > 0 Then
        BePatArt = BELEGINFO_PATIENT_ART
    End If
    Line = Line & Q & BePatArt & Q & Sep                  ' Beleginfo-Art 3 (Patientennummer)
    Line = Line & Q & PaNum & Q & Sep                     ' Beleginfo-Inhalt 3 (Patientennummer)
    Line = Line & Q & Q & Sep                             ' Beleginfo-Art 4
    Line = Line & Q & Q & Sep                             ' Beleginfo-Inhalt 4
    Line = Line & Q & Q & Sep                             ' Beleginfo-Art 5
    Line = Line & Q & Q & Sep                             ' Beleginfo-Inhalt 5
    Line = Line & Q & Q & Sep                             ' Beleginfo-Art 6
    Line = Line & Q & Q & Sep                             ' Beleginfo-Inhalt 6
    Line = Line & Q & Q & Sep                             ' Beleginfo-Art 7
    Line = Line & Q & Q & Sep                             ' Beleginfo-Inhalt 7
    Line = Line & Q & Q & Sep                             ' Beleginfo-Art 8
    Line = Line & Q & Q & Sep                             ' Beleginfo-Inhalt 8

    ' Field 37: KOST1 - Kostenstelle
    Line = Line & Q & Koste & Q & Sep

    ' Field 38: KOST2 - Kostenstelle
    Line = Line & Q & Q & Sep

    ' Field 39: Kost-Menge
    Line = Line & Sep

    ' Field 40: EU-Land u. UStID
    Line = Line & Q & Q & Sep

    ' Field 41: EU-Steuersatz
    Line = Line & Sep

    ' Field 42: Abw. Versteuerungsart
    Line = Line & Q & Q & Sep

    ' Field 43: Sachverhalt L+L
    Line = Line & Sep

    ' Field 44: Funktionsergaenzung L+L
    Line = Line & Sep

    ' Fields 45-47: BU 49 fields
    Line = Line & Sep  ' BU 49 Hauptfunktionstyp
    Line = Line & Sep  ' BU 49 Hauptfunktionsnummer
    Line = Line & Sep  ' BU 49 Funktionsergaenzung

    ' Fields 48-87: Zusatzinformation 1-20 (Art + Inhalt pairs)
    If Len(Komme) > 0 Then
        Line = Line & Q & "KOMMENTAR" & Q & Sep  ' Art 1
        Line = Line & Q & Komme & Q & Sep        ' Inhalt 1 (Kommentar)
    Else
        Line = Line & Q & Q & Sep                ' Art 1 (leer)
        Line = Line & Q & Q & Sep                ' Inhalt 1 (leer)
    End If
    Dim i As Integer
    For i = 2 To 20
        Line = Line & Q & Q & Sep  ' Art
        Line = Line & Q & Q & Sep  ' Inhalt
    Next i

    ' Field 88: Stueck (Numeric)
    Line = Line & Sep

    ' Field 89: Gewicht (Numeric)
    Line = Line & Sep

    ' Field 90: Zahlweise (Text)
    Line = Line & Q & Q & Sep

    ' Field 91: Forderungsart (Text)
    Line = Line & Q & Q & Sep

    ' Field 92: Veranlagungsjahr
    Line = Line & Format$(BuDat, "yyyy") & Sep

    ' Field 93: Zugeordnete Faelligkeit
    Line = Line & Sep

    ' Field 94: Skontotyp
    Line = Line & Sep

    ' Field 95: Auftragsnummer
    Line = Line & Q & Q & Sep

    ' Field 96: Buchungstyp
    Line = Line & Q & Q & Sep

    ' Field 97: USt-Schluessel (Anzahlungen)
    Line = Line & Sep

    ' Field 98: EU-Land (Anzahlungen)
    Line = Line & Q & Q & Sep

    ' Field 99: Sachverhalt L+L (Anzahlungen)
    Line = Line & Sep

    ' Field 100: EU-Steuersatz (Anzahlungen)
    Line = Line & Sep

    ' Field 101: Erloeskonto (Anzahlungen)
    Line = Line & Sep

    ' Field 102: Herkunft-Kz
    Line = Line & Q & "SM" & Q & Sep

    ' Field 103: Buchungs GUID
    Line = Line & Q & BuGui & Q & Sep

    ' Field 104: KOST-Datum
    Line = Line & Sep

    ' Field 105: SEPA-Mandatsreferenz
    Line = Line & Q & Q & Sep

    ' Field 106: Skontosperre
    Line = Line & Q & "0" & Q & Sep

    ' Field 107: Gesellschaftername
    Line = Line & Q & Q & Sep

    ' Field 108: Beteiligtennummer
    Line = Line & Q & Q & Sep

    ' Field 109: Identifikationsnummer
    Line = Line & Q & Q & Sep

    ' Field 110: Zeichnernummer
    Line = Line & Q & Q & Sep

    ' Field 111: Postensperre bis
    Line = Line & Q & Q & Sep

    ' Field 112: Bezeichnung SoBil-Sachverhalt
    Line = Line & Q & Q & Sep

    ' Field 113: Kennzeichen SoBil-Buchung
    Line = Line & Q & Q & Sep

    ' Field 114: Festschreibung
    If BLock Then
        Line = Line & Q & "1" & Q & Sep
    Else
        Line = Line & Q & "0" & Q & Sep
    End If

    ' Field 115: Leistungsdatum
    Line = Line & Sep

    ' Field 116: Datum Zuord. Steuerperiode
    Line = Line & Sep

    ' Fields 117-125: DATEV v700 Format v13 additional fields (all empty)
    ' Textfelder mit "" quotieren, numerische Felder leer
    Line = Line & Q & Q & Sep  ' 117: Generalumkehr (Storno) - Text
    Line = Line & Sep          ' 118: Steuersatz - Numeric
    Line = Line & Q & Q & Sep  ' 119: Land - Text
    Line = Line & Q & Q & Sep  ' 120: Abrechnungsreferenz - Text
    Line = Line & Sep          ' 121: BVV-Position - Numeric
    Line = Line & Q & Q & Sep  ' 122: EU-Land u. UStID (Ursprungsland) - Text
    Line = Line & Q & Q & Sep  ' 123: EU-USt-IdNr (Ursprung) - Text
    Line = Line & Sep          ' 124: Sachverhalt Warenbewegung - Numeric
    Line = Line & vbNullString ' 125: Steuerschloessel Devisen - Numeric (last field)

    BuildCSVDataLine = Line
    Exit Function

ErrHandler:
    If GlLog = True Then SLogi "=== BuildCSVDataLine ERROR ==="
    If GlLog = True Then SLogi "  Err.Number: " & Err.Number
    If GlLog = True Then SLogi "  Err.Description: " & Err.Description
    If GlLog = True Then SLogi "  Err.Source: " & Err.Source
    DoEvents
    SPopu "BuildCSVDataLine " & Err.Number, Err.Description, IC48_Warning
    BuildCSVDataLine = vbNullString
End Function

'--------------------------------------------------------------------------------
' PRIVATE - CSV Helper Functions
'--------------------------------------------------------------------------------

Private Function DetermineAmount(ByRef RST As ADODB.Recordset, _
                                 ByVal BuTyp As Integer, _
                                 ByRef Config As DATEV_ExportConfig) As String
    ' Returns amount in German decimal format (comma as decimal separator)
    ' Amount is always positive in DATEV format

    Dim Amount As Currency
    Amount = 0

    On Error Resume Next

    If GlBuc = True Then
        ' Simple bookkeeping
        Select Case BuTyp
        Case 1: ' Ausgabe (expense)
            If Not IsNull(RST.Fields("Ausgabe").Value) Then
                Amount = Abs(CCur(RST.Fields("Ausgabe").Value))
            ElseIf Not IsNull(RST.Fields("Einnahme").Value) Then
                Amount = Abs(CCur(RST.Fields("Einnahme").Value))
            End If
        Case 2: ' Einnahme (income)
            If Not IsNull(RST.Fields("Einnahme").Value) Then
                Amount = Abs(CCur(RST.Fields("Einnahme").Value))
            ElseIf Not IsNull(RST.Fields("Ausgabe").Value) Then
                Amount = Abs(CCur(RST.Fields("Ausgabe").Value))
            End If
        Case Else:
            If Not IsNull(RST.Fields("Einnahme").Value) And RST.Fields("Einnahme").Value > 0 Then
                Amount = Abs(CCur(RST.Fields("Einnahme").Value))
            ElseIf Not IsNull(RST.Fields("Ausgabe").Value) And RST.Fields("Ausgabe").Value > 0 Then
                Amount = Abs(CCur(RST.Fields("Ausgabe").Value))
            End If
        End Select
    Else
        ' Double-entry bookkeeping - Feld "Betrag" existiert nicht, verwende Einnahme/Ausgabe
        If Not IsNull(RST.Fields("Einnahme").Value) And RST.Fields("Einnahme").Value > 0 Then
            Amount = Abs(CCur(RST.Fields("Einnahme").Value))
        ElseIf Not IsNull(RST.Fields("Ausgabe").Value) And RST.Fields("Ausgabe").Value > 0 Then
            Amount = Abs(CCur(RST.Fields("Ausgabe").Value))
        End If
    End If

    On Error GoTo 0

    ' Format with German decimal (comma)
    DetermineAmount = FormatAmountGerman(Amount)
End Function

Private Function DetermineSollHaben(ByRef RST As ADODB.Recordset, _
                                    ByVal BuTyp As Integer, _
                                    ByRef Config As DATEV_ExportConfig) As String
    ' Returns "S" (Soll/Debit) or "H" (Haben/Credit)
    ' Respects GlTSH swap setting

    Dim Result As String

    On Error Resume Next

    If GlBuc = True Then
        ' Simple bookkeeping
        Select Case BuTyp
        Case 1: ' Ausgabe
            Result = IIf(Config.SwapDebitCredit, "S", "H")
        Case 2: ' Einnahme
            Result = IIf(Config.SwapDebitCredit, "H", "S")
        Case Else:
            If Not IsNull(RST.Fields("Einnahme").Value) And RST.Fields("Einnahme").Value > 0 Then
                Result = IIf(Config.SwapDebitCredit, "H", "S")
            Else
                Result = IIf(Config.SwapDebitCredit, "S", "H")
            End If
        End Select
    Else
        ' Double-entry bookkeeping - default to S, swap if configured
        Result = IIf(Config.SwapDebitCredit, "H", "S")
    End If

    On Error GoTo 0

    DetermineSollHaben = Result
End Function

Private Function GetTaxKey(ByVal TaxRate As Single, ByVal SollHaben As String) As String
    ' Returns DATEV BU-Schluessel (tax key) based on tax rate and S/H
    ' H = Umsatzsteuer (sales tax), S = Vorsteuer (input tax)
    ' Handles various formats: 0.07, 0.19, 7, 7.0, 19, 19.0, etc.

    Dim NormalizedRate As Single

    ' Normalize tax rate to percentage (0-100 scale)
    ' Handles both decimal (0.07, 0.19) and percentage (7, 19) formats
    If TaxRate > 0 And TaxRate < 1 Then
        ' Decimal format (0.07, 0.19) - convert to percentage
        NormalizedRate = TaxRate * 100
    Else
        ' Already in percentage format (7, 19)
        NormalizedRate = TaxRate
    End If

    ' Round to handle floating point imprecision
    NormalizedRate = CSng(Int(NormalizedRate + 0.5))

    If SollHaben = "H" Then
        ' Umsatzsteuer (sales tax)
        ' DATEV XSD erfordert BU-Schluessel OHNE fuehrende Nullen
        Select Case NormalizedRate
        Case 7: GetTaxKey = "2"       ' 7% USt
        Case 9: GetTaxKey = vbNullString  ' 9% (special rate, no automatic key)
        Case 16: GetTaxKey = "5"      ' 16% USt (Corona reduced rate 2020)
        Case 19: GetTaxKey = "3"      ' 19% USt (standard rate)
        Case 5: GetTaxKey = "4"       ' 5% USt (Corona reduced rate 2020)
        Case 0: GetTaxKey = vbNullString  ' Tax-free
        Case Else: GetTaxKey = vbNullString
        End Select
    Else
        ' Vorsteuer (input tax)
        ' DATEV XSD erfordert BU-Schluessel OHNE fuehrende Nullen
        Select Case NormalizedRate
        Case 7: GetTaxKey = "8"       ' 7% VSt
        Case 9: GetTaxKey = vbNullString  ' 9% (special rate, no automatic key)
        Case 16: GetTaxKey = "7"      ' 16% VSt (Corona reduced rate 2020)
        Case 19: GetTaxKey = "9"      ' 19% VSt (standard rate)
        Case 5: GetTaxKey = "6"       ' 5% VSt (Corona reduced rate 2020)
        Case 0: GetTaxKey = vbNullString  ' Tax-free
        Case Else: GetTaxKey = vbNullString
        End Select
    End If
End Function

Private Function FormatAccountNumber(ByVal AccountNo As Long, _
                                     ByVal FourDigit As Boolean) As String
    ' Format account number with correct length
    Dim KoStr As String
    Dim Lange As Integer

    If AccountNo = 0 Then
        If FourDigit Then
            FormatAccountNumber = "0000"
        Else
            FormatAccountNumber = "000000"
        End If
        Exit Function
    End If

    KoStr = CStr(AccountNo)
    Lange = Len(KoStr)

    If FourDigit Then
        Select Case Lange
        Case 1: FormatAccountNumber = KoStr & "000"
        Case 2: FormatAccountNumber = KoStr & "00"
        Case 3: FormatAccountNumber = KoStr & "0"
        Case 4: FormatAccountNumber = KoStr
        Case Else: FormatAccountNumber = Left$(KoStr, 4)
        End Select
    Else
        Select Case Lange
        Case 1: FormatAccountNumber = KoStr & "00000"
        Case 2: FormatAccountNumber = KoStr & "0000"
        Case 3: FormatAccountNumber = KoStr & "000"
        Case 4: FormatAccountNumber = KoStr & "00"
        Case 5: FormatAccountNumber = KoStr & "0"
        Case 6: FormatAccountNumber = KoStr
        Case Else: FormatAccountNumber = Left$(KoStr, 6)
        End Select
    End If
End Function

Private Function FormatAmountGerman(ByVal Amount As Currency) As String
    ' Format amount for DATEV: German decimal (comma), no thousands separator
    Dim Formatted As String

    Formatted = Format$(Amount, "0.00")
    ' Replace dot with comma for German format
    Formatted = Replace(Formatted, ".", ",")

    FormatAmountGerman = Formatted
End Function

Private Function GetCashAccountNumber(ByVal GeKto As Integer) As Long
    ' Get the DATEV account number for a cash/bank account
    ' Uses GlGeK global array like S_DaKoK

    Dim SaKto As Long
    Dim AktZa As Integer

    On Error Resume Next

    For AktZa = 1 To UBound(GlGeK)
        If GlGeK(AktZa, 0) = GeKto Then
            If GlGeK(AktZa, 2) <> vbNullString Then
                SaKto = CLng(GlGeK(AktZa, 2))
            Else
                If GldKt = True Then
                    SaKto = 1200
                Else
                    SaKto = 120000
                End If
            End If
            Exit For
        End If
    Next AktZa

    On Error GoTo 0

    GetCashAccountNumber = SaKto
End Function

Private Function GetCostCenter(ByVal ManNr As Long) As String
    ' Get cost center for mandant from GlThe array
    Dim AktZa As Integer
    Dim Koste As String

    On Error Resume Next

    For AktZa = 1 To UBound(GlThe)
        If GlThe(AktZa, 0) = ManNr Then
            If GlThe(AktZa, 47) <> vbNullString Then
                Koste = Left$(CStr(GlThe(AktZa, 47)), 8)
            End If
            Exit For
        End If
    Next AktZa

    On Error GoTo 0

    GetCostCenter = Koste
End Function

Private Function BuildBookingText(ByRef RST As ADODB.Recordset, _
                                  ByRef Config As DATEV_ExportConfig, _
                                  ByVal Storn As Boolean) As String
    ' Build booking text, optionally including patient name
    ' Max 60 characters

    Dim BuStr As String
    Dim PaStr As String
    Dim BuTex As String
    Dim Prefix As String

    On Error Resume Next

    ' Get patient name if available
    PaStr = SafeString(RST.Fields("Patient").Value)

    ' Get booking text
    BuTex = SafeString(RST.Fields("Buchtext").Value)

    ' Build combined text
    If Config.IncludePatientName And Len(PaStr) > 0 Then
        If Len(BuTex) > 0 Then
            BuStr = PaStr & " " & BuTex
        Else
            BuStr = PaStr
        End If
    Else
        BuStr = BuTex
    End If

    ' Add storno prefix if applicable
    If Storn Then
        Prefix = "[STORNIERT] "
        BuStr = Prefix & BuStr
    End If

    ' Truncate to max length
    If Len(BuStr) > MAX_BOOKING_TEXT_LENGTH Then
        BuStr = Left$(BuStr, MAX_BOOKING_TEXT_LENGTH)
    End If

    ' Remove semicolons (DATEV separator)
    BuStr = Replace(BuStr, DATEV_SEPARATOR, vbNullString)

    ' Remove quotes
    BuStr = Replace(BuStr, Chr$(34), vbNullString)

    On Error GoTo 0

    BuildBookingText = Trim$(BuStr)
End Function

Private Function SanitizeCollectionKey(ByVal KeyValue As String) As String
    ' Sanitize a string for use as VB6 Collection key
    ' VB6 Collections throw Error 5 on problematic characters
    ' Use very strict filtering: only alphanumeric, dot, underscore, hyphen
    On Error Resume Next

    Dim i As Long
    Dim c As String
    Dim a As Integer
    Dim Result As String
    Dim KeyLen As Long

    Result = vbNullString
    KeyLen = Len(KeyValue)

    If KeyLen = 0 Then
        SanitizeCollectionKey = vbNullString
        Exit Function
    End If

    For i = 1 To KeyLen
        c = Mid$(KeyValue, i, 1)
        a = Asc(c)
        ' Only allow: A-Z (65-90), a-z (97-122), 0-9 (48-57), dot (46), underscore (95), hyphen (45)
        Select Case a
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95
                Result = Result & c
        End Select
    Next i

    SanitizeCollectionKey = Result
    On Error GoTo 0
End Function

Private Function IsDocumentAlreadyExported(ByVal FileName As String) As Boolean
    ' Check if document was already exported (for split postings)
    ' Uses iteration instead of key-based lookup to avoid Error 5

    ' Check for empty/null filename
    If Len(FileName) = 0 Then
        IsDocumentAlreadyExported = False
        Exit Function
    End If

    ' Check if collection exists and has items
    If m_ExportedFileNames Is Nothing Then
        IsDocumentAlreadyExported = False
        Exit Function
    End If

    If m_ExportedFileNames.Count = 0 Then
        IsDocumentAlreadyExported = False
        Exit Function
    End If

    ' Sanitize filename for comparison
    Dim CleanKey As String
    CleanKey = SanitizeCollectionKey(FileName)

    If Len(CleanKey) = 0 Then
        IsDocumentAlreadyExported = False
        Exit Function
    End If

    ' Search through collection by iterating (avoids Error 5 on key access)
    Dim i As Long
    Dim StoredKey As String
    On Error Resume Next
    For i = 1 To m_ExportedFileNames.Count
        StoredKey = m_ExportedFileNames(i)
        If StrComp(StoredKey, CleanKey, vbTextCompare) = 0 Then
            IsDocumentAlreadyExported = True
            Exit Function
        End If
    Next i
    On Error GoTo 0

    IsDocumentAlreadyExported = False
End Function

Private Sub TrackExportedDocument(ByVal FileName As String, ByVal guid As String)
    ' Track exported document for deduplication
    On Error Resume Next

    ' Check for empty parameters
    If Len(FileName) = 0 Then Exit Sub
    If Len(guid) = 0 Then Exit Sub

    ' Check if collections exist
    If m_DocumentGUIDs Is Nothing Then Exit Sub
    If m_ExportedFileNames Is Nothing Then Exit Sub

    ' Sanitize filename for lookup
    Dim CleanKey As String
    CleanKey = SanitizeCollectionKey(FileName)

    If Len(CleanKey) = 0 Then Exit Sub

    ' Add GUID to collection (no key - avoids Error 5)
    m_DocumentGUIDs.Add guid

    ' Add filename to parallel collection for iteration-based lookup
    m_ExportedFileNames.Add CleanKey

    On Error GoTo 0
End Sub

Private Function GetInvoiceGUID(ByVal RechNr As String) As String
    ' Get stored GUID for a RechNr, returns empty string if not found
    ' Uses iteration instead of key-based lookup to avoid Error 5

    GetInvoiceGUID = vbNullString

    If m_InvoiceRechNrs Is Nothing Then Exit Function
    If m_InvoiceGUIDs Is Nothing Then Exit Function
    If Len(RechNr) = 0 Then Exit Function
    If m_InvoiceRechNrs.Count = 0 Then Exit Function

    Dim CleanKey As String
    CleanKey = SanitizeCollectionKey(RechNr)
    If Len(CleanKey) = 0 Then Exit Function

    ' Search through collection by iterating
    Dim i As Long
    Dim StoredRechNr As String
    On Error Resume Next
    For i = 1 To m_InvoiceRechNrs.Count
        StoredRechNr = m_InvoiceRechNrs(i)
        If StrComp(StoredRechNr, CleanKey, vbTextCompare) = 0 Then
            GetInvoiceGUID = m_InvoiceGUIDs(i)
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Private Sub RegisterInvoiceGUID(ByVal RechNr As String, ByVal guid As String)
    ' Register RechNr -> GUID mapping for PDF renaming
    ' Handles split postings (same RechNr, same GUID) and detects data inconsistencies
    ' Uses parallel collections to avoid Error 5 on key-based access

    If m_InvoiceRechNrs Is Nothing Then Exit Sub
    If m_InvoiceGUIDs Is Nothing Then Exit Sub
    If Len(RechNr) = 0 Then Exit Sub
    If Len(guid) = 0 Then Exit Sub

    Dim CleanKey As String
    CleanKey = SanitizeCollectionKey(RechNr)
    If Len(CleanKey) = 0 Then Exit Sub

    ' Check if RechNr already registered (search by iteration)
    Dim ExistingGUID As String
    ExistingGUID = GetInvoiceGUID(CleanKey)

    If Len(ExistingGUID) = 0 Then
        ' Not found - add new mapping to both collections
        On Error Resume Next
        m_InvoiceRechNrs.Add CleanKey
        m_InvoiceGUIDs.Add guid
        On Error GoTo 0
    ElseIf ExistingGUID <> guid Then
        ' Different GUID for same RechNr - this is a data inconsistency!
        ' Log warning but continue (use first GUID found)
        If GlLog = True Then
            SLogi "WARNUNG: RechNr '" & RechNr & "' hat unterschiedliche GUIDs:"
            SLogi "  Gespeichert: " & ExistingGUID
            SLogi "  Neu gefunden: " & guid
            SLogi "  -> Verwende erste GUID"
        End If
    End If
    ' Else: Same GUID = split posting, already registered, no action needed
End Sub

'--------------------------------------------------------------------------------
' PRIVATE - Safe Type Conversion Helpers
'--------------------------------------------------------------------------------

Private Function SafeString(ByVal Value As Variant) As String
    If IsNull(Value) Then
        SafeString = vbNullString
    Else
        SafeString = Trim$(CStr(Value))
    End If
End Function

Private Function SafeLong(ByVal Value As Variant) As Long
    If IsNull(Value) Then
        SafeLong = 0
    ElseIf IsNumeric(Value) Then
        SafeLong = CLng(Value)
    Else
        SafeLong = 0
    End If
End Function

Private Function SafeInt(ByVal Value As Variant) As Integer
    If IsNull(Value) Then
        SafeInt = 0
    ElseIf IsNumeric(Value) Then
        SafeInt = CInt(Value)
    Else
        SafeInt = 0
    End If
End Function

Private Function SafeSingle(ByVal Value As Variant) As Single
    If IsNull(Value) Then
        SafeSingle = 0
    ElseIf IsNumeric(Value) Then
        SafeSingle = CSng(Value)
    Else
        SafeSingle = 0
    End If
End Function

Private Function SafeBool(ByVal Value As Variant) As Boolean
    If IsNull(Value) Then
        SafeBool = False
    Else
        SafeBool = CBool(Value)
    End If
End Function

Private Function SafeCurrency(ByVal Value As Variant) As Currency
    If IsNull(Value) Then
        SafeCurrency = 0
    ElseIf IsNumeric(Value) Then
        SafeCurrency = CCur(Value)
    Else
        SafeCurrency = 0
    End If
End Function

'--------------------------------------------------------------------------------
' GetFieldValue - Get field value with fallback field names
'--------------------------------------------------------------------------------
Private Function GetFieldValue(ByRef RST As ADODB.Recordset, _
                               ParamArray FieldNames() As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim FldName As Variant
    Dim Fld As ADODB.Field
    Dim FieldFound As Boolean

    GetFieldValue = Null

    ' Check if RST is valid
    If RST Is Nothing Then Exit Function
    If RST.State <> adStateOpen Then Exit Function

    On Error Resume Next
    For Each FldName In FieldNames
        If Len(CStr(FldName)) > 0 Then
            ' Check if field exists by iterating through Fields collection
            FieldFound = False
            For j = 0 To RST.Fields.Count - 1
                If UCase$(RST.Fields(j).Name) = UCase$(CStr(FldName)) Then
                    FieldFound = True
                    Exit For
                End If
            Next j

            If FieldFound Then
                Err.Clear
                GetFieldValue = RST.Fields(CStr(FldName)).Value
                If Err.Number = 0 Then
                    Exit Function
                End If
            End If
        End If
    Next FldName
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' FieldExists - Check if a field exists in the recordset
'--------------------------------------------------------------------------------
Private Function FieldExists(ByRef RST As ADODB.Recordset, ByVal FieldName As String) As Boolean
    Dim Fld As ADODB.Field
    On Error Resume Next
    Set Fld = RST.Fields(FieldName)
    FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GenerateCSVFilename(ByRef Config As DATEV_ExportConfig) As String
    Dim FileName As String

    ' Benutzer-gewaehlten Dateinamen verwenden falls vorhanden
    If Len(Config.ExportFileName) > 0 Then
        FileName = Config.ExportFileName & ".csv"
    Else
        ' Fallback: EXTF_Buchungsstapel_YYYYMMDD_HHMMSS.csv
        FileName = "EXTF_Buchungsstapel_" & _
                   Format$(Date, "yyyymmdd") & "_" & _
                   Format$(Time, "hhnnss") & ".csv"
    End If

    GenerateCSVFilename = Config.ExportPath & FileName
End Function

Private Function WriteCSVFile(ByVal FilePath As String, ByVal Content As String) As Boolean
On Error GoTo ErrHandler

    ' Write as ANSI/Windows-1252 (DATEV EXTF requirement)
    ' StrConv with vbFromUnicode converts to system ANSI codepage (Windows-1252)
    Dim FileNum As Integer
    Dim Buffer() As Byte

    Buffer = StrConv(Content, vbFromUnicode)
    FileNum = FreeFile
    Open FilePath For Binary Access Write As #FileNum
    Put #FileNum, , Buffer
    Close #FileNum

    ' Verify file was created
    WriteCSVFile = m_clFil.FilVor(FilePath)
    Exit Function

ErrHandler:
    LogError "WriteCSVFile", Err.Number, Err.Description
    WriteCSVFile = False
End Function

Private Function WriteCSVFileAnsi(ByVal FilePath As String, ByVal Content As String) As Boolean
On Error GoTo ErrHandler

    ' Write as ANSI/Windows-1252 (DATEV EXTF requirement)
    ' StrConv with vbFromUnicode converts to system ANSI codepage (Windows-1252)
    Dim FileNum As Integer
    Dim Buffer() As Byte

    Buffer = StrConv(Content, vbFromUnicode)
    FileNum = FreeFile
    Open FilePath For Binary Access Write As #FileNum
    Put #FileNum, , Buffer
    Close #FileNum

    ' Verify file was created
    WriteCSVFileAnsi = m_clFil.FilVor(FilePath)
    Exit Function

ErrHandler:
    LogError "WriteCSVFileAnsi", Err.Number, Err.Description
    WriteCSVFileAnsi = False
End Function

'--------------------------------------------------------------------------------
' PRIVATE - XML Generation (Full Implementation)
'--------------------------------------------------------------------------------

Private Function GenerateXMLExport(ByRef RST As ADODB.Recordset, _
                                   ByRef Config As DATEV_ExportConfig, _
                                   ByVal CSVFilePath As String) As String
On Error GoTo ErrHandler

    Dim XMLContent As String
    Dim XMLFilePath As String
    Dim DocumentsDir As String
    Dim DocumentCount As Long
    Dim RecordCount As Long
    Dim CurrentRecord As Long
    Dim ProcessedDocs As Collection

    ' Initialize
    DocumentCount = 0
    Set ProcessedDocs = New Collection

    ' Documents go into the export subfolder directly
    DocumentsDir = Config.ExportPath
    ' If Config.ExportDocuments Then
    '     If Not EnsureExportDirectory(DocumentsDir) Then
    '         LogError "GenerateXMLExport", 3001, "Belegverzeichnis konnte nicht erstellt werden: " & DocumentsDir
    '         GenerateXMLExport = vbNullString
    '         Exit Function
    '     End If
    ' End If

    ' Reset progress for XML/PDF phase with clear phase label
    ResetProgressForPhase "DATEV Export - XML Belegverknoepfung", RST.RecordCount

    ' Build XML header
    XMLContent = BuildXMLHeader(Config)

    ' Process records for documents
    RecordCount = RST.RecordCount
    RST.MoveFirst
    CurrentRecord = 0
    m_LastProgressUpdate = 0

    Do While Not RST.EOF
        CurrentRecord = CurrentRecord + 1

        ' Update progress at intervals (consistent with GenerateCSVExport)
        If (CurrentRecord - m_LastProgressUpdate) >= PROGRESS_UPDATE_INTERVAL Then
            UpdateProgress CurrentRecord, RecordCount, "Verarbeite Beleg " & CurrentRecord & " von " & RecordCount
            m_LastProgressUpdate = CurrentRecord
        End If

        ' Check for cancellation
        If (CurrentRecord Mod DOEVENTS_INTERVAL) = 0 Then
            DoEvents
            If CheckCancelled() Then
                m_Cancelled = True
                Set ProcessedDocs = Nothing
                GenerateXMLExport = vbNullString
                Exit Function
            End If
        End If

        ' Process document for this record
        Dim DocXML As String
        DocXML = ProcessDocumentForXML(RST, Config, DocumentsDir, ProcessedDocs)
        If Len(DocXML) > 0 Then
            XMLContent = XMLContent & DocXML
            DocumentCount = DocumentCount + 1
        End If

        RST.MoveNext
    Loop

    ' Final progress update
    UpdateProgress RecordCount, RecordCount, "Finalisiere XML..."

    ' Build XML footer
    XMLContent = XMLContent & BuildXMLFooter(Config)

    ' Log document count summary
    If GlLog = True Then
        SLogi "DATEV: XML generation complete - " & DocumentCount & " document elements generated from " & RecordCount & " records"
        If DocumentCount = 0 Then
            SLogi "DATEV: WARNING - No document elements generated! Check if records have GUIDs and meet validation criteria."
        End If
    End If

    ' Generate XML filename
    XMLFilePath = GenerateXMLFilename(Config)

    ' Write XML file
    If WriteXMLFile(XMLFilePath, XMLContent) Then
        ' Add to ZIP file list (relative path)
        AddToZipList "document.xml"
        GenerateXMLExport = XMLFilePath

        ' Option B: Generate ledger.xml for structured booking data import
        If Config.UseLedgerXML Then
            Dim LedgerXMLPath As String
            If GlLog = True Then SLogi "DATEV: Option B aktiv, starte ledger.xml Generierung"
            DoEvents
            On Error Resume Next
            RST.MoveFirst  ' Reset recordset for second pass
            If Err.Number <> 0 Then
                If GlLog = True Then SLogi "DATEV: RST.MoveFirst Fehler: " & Err.Description
                LogError "GenerateXMLExport", 3011, "RST.MoveFirst fehlgeschlagen: " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
            LedgerXMLPath = GenerateLedgerXML(RST, Config)
            If Len(LedgerXMLPath) = 0 Then
                If GlLog = True Then SLogi "DATEV: ledger.xml nicht erstellt!"
                LogError "GenerateXMLExport", 3010, "ledger.xml konnte nicht erstellt werden"
                ' Continue anyway - document.xml was successful
            Else
                If GlLog = True Then SLogi "DATEV: ledger.xml erstellt: " & LedgerXMLPath
            End If
        Else
            If GlLog = True Then SLogi "DATEV: UseLedgerXML = False, keine ledger.xml"
        End If
    Else
        GenerateXMLExport = vbNullString
    End If

    ' Cleanup
    Set ProcessedDocs = Nothing

    Exit Function

ErrHandler:
    Set ProcessedDocs = Nothing
    LogError "GenerateXMLExport", Err.Number, Err.Description
    GenerateXMLExport = vbNullString
End Function

'--------------------------------------------------------------------------------
' PRIVATE - XML Building Functions
'--------------------------------------------------------------------------------

Private Function BuildXMLHeader(ByRef Config As DATEV_ExportConfig) As String
    Dim Header As String
    Dim GeneratedDate As String

    GeneratedDate = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")

    ' document.xml always uses archive format with document namespace
    ' Both Option A and Option B use the same document.xml structure
    ' Option B adds cashLedger/accountsPayable/accountsReceivable extensions that reference ledger.xml
    If GlLog = True Then
        If Config.UseLedgerXML Then
            SLogi "DATEV: Building document.xml header (Option B: with ledger.xml references)"
        Else
            SLogi "DATEV: Building document.xml header (Option A: standalone documents)"
        End If
    End If

    Header = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    Header = Header & "<archive xmlns=""" & XML_NAMESPACE & """" & vbCrLf
    Header = Header & "         xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & vbCrLf
    Header = Header & "         xsi:schemaLocation=""" & XML_SCHEMA_LOCATION & """" & vbCrLf
    Header = Header & "         version=""" & XML_VERSION & """" & vbCrLf
    Header = Header & "         generatingSystem=""" & DATEV_GENERATING_SYSTEM & """>" & vbCrLf
    Header = Header & "  <header>" & vbCrLf
    Header = Header & "    <date>" & GeneratedDate & "</date>" & vbCrLf
    Header = Header & "    <description>SimpliMed DATEV Export</description>" & vbCrLf
    Header = Header & "    <consultantNumber>" & Format$(Config.Beraternummer, "0") & "</consultantNumber>" & vbCrLf
    Header = Header & "    <clientNumber>" & Format$(Config.Mandantennummer, "0") & "</clientNumber>" & vbCrLf
    Header = Header & "  </header>" & vbCrLf
    Header = Header & "  <content>" & vbCrLf

    BuildXMLHeader = Header
End Function

Private Function BuildXMLFooter(ByRef Config As DATEV_ExportConfig) As String
    ' document.xml always uses archive format
    BuildXMLFooter = "  </content>" & vbCrLf & "</archive>" & vbCrLf
End Function

Private Function ProcessDocumentForXML(ByRef RST As ADODB.Recordset, _
                                       ByRef Config As DATEV_ExportConfig, _
                                       ByVal DocumentsDir As String, _
                                       ByRef ProcessedDocs As Collection) As String
On Error GoTo ErrHandler

    Dim DocXML As String
    Dim BuTyp As Integer
    Dim BuGui As String
    Dim DaNam As String
    Dim DaPfa As String
    Dim ReStr As String
    Dim Komme As String
    Dim BuDat As Date
    Dim BuDatVal As Variant
    Dim IsRevenue As Boolean
    Dim FormattedGUID As String
    Dim TargetFileName As String
    Dim TargetPath As String
    Dim DocExists As Boolean

    DocXML = vbNullString

    ' Get booking type to determine if revenue or expense
    ' Debitoren-Modus: immer Einnahme (BuTyp=2)
    If m_InvMod Then
        BuTyp = 2
    Else
        BuTyp = SafeInt(RST.Fields("IDA").Value)
    End If
    IsRevenue = (BuTyp = 2) ' Type 2 = Einnahme (Revenue)

    ' Skip records with zero/invalid amount (same validation as CSV export)
    ' This ensures XML only contains documents for records that are in the CSV
    Dim Amount As Currency
    Amount = DetermineAmountValue(RST, BuTyp)
    If Amount < MIN_VALID_AMOUNT Then
        ProcessDocumentForXML = vbNullString
        Exit Function
    End If

    ' Get GUID
    BuGui = SafeString(RST.Fields("GuiID").Value)
    If Len(BuGui) = 0 Then
        ' No GUID - cannot link document
        ProcessDocumentForXML = vbNullString
        Exit Function
    End If

    ' Skip payment bookings (K prefix) - they don't have document files
    ' Only process R (invoices), B (expenses), G (other business docs)
    If Len(BuGui) > 0 Then
        Dim GuidPrefix As String
        GuidPrefix = UCase$(Left$(BuGui, 1))
        If GuidPrefix = "K" Then
            If GlLog = True Then SLogi "  >>> ProcessDocumentForXML: Skipped (payment booking, no document)"
            ProcessDocumentForXML = vbNullString
            Exit Function
        End If
    End If

    ' Check if expense document was validated as invalid (file not found in GlBPf)
    ' Debitoren: keine Validierung noetig (nur generierte PDFs, keine Ausgabebelege)
    If Not m_InvMod Then
        If Not IsExpenseDocumentValid(BuGui) Then
            If GlLog = True Then SLogi "ProcessDocumentForXML: Skipped - document invalid (GuiID: " & BuGui & ")"
            ProcessDocumentForXML = vbNullString
            Exit Function
        End If
    End If

    ' Format GUID for XML (8-4-4-4-12 lowercase)
    FormattedGUID = DATEV_FormatGUIDForXML(BuGui)

    ' Check if this document was already processed (split posting handling)
    ' Use iteration to avoid Error 5 when key doesn't exist
    Dim AlreadyProcessed As Boolean
    Dim j As Long
    Dim StoredGUID As String
    AlreadyProcessed = False
    If ProcessedDocs.Count > 0 Then
        For j = 1 To ProcessedDocs.Count
            StoredGUID = ProcessedDocs.Item(j)
            If StrComp(StoredGUID, FormattedGUID, vbTextCompare) = 0 Then
                AlreadyProcessed = True
                Exit For
            End If
        Next j
    End If
    If AlreadyProcessed Then
        ' Document already processed - this is a split posting
        ProcessDocumentForXML = vbNullString
        Exit Function
    End If

    ' Get document info
    ' Debitoren: kein Datei-Feld, Dateiname wird aus RechNr generiert
    If m_InvMod Then
        DaNam = vbNullString  ' Wird weiter unten aus RechNr generiert
    Else
        DaNam = SafeString(RST.Fields("Datei").Value)
    End If
    DaPfa = vbNullString  ' Feld "Pfad" existiert nicht in qrySimBuSu
    ' RechNr aus Recordset - gilt fuer Debitoren (Einnahmen) UND Kreditoren (Ausgaben)
    ' Validierung: max 12 Zeichen (MAX_BELEGFELD1_LENGTH), Sonderzeichen werden entfernt
    ReStr = SanitizeTextField(SafeString(RST.Fields("RechNr").Value), MAX_BELEGFELD1_LENGTH)
    Komme = SanitizeTextField(SafeString(RST.Fields("Kommentar").Value), 60)

    ' Debug: Show document info for expenses
    If GlLog = True And Not IsRevenue And Len(DaNam) > 0 Then
        SLogi "ProcessDocumentForXML: Expense with document"
        SLogi "  DaNam (Datei) = " & DaNam
        SLogi "  DaPfa (Pfad) = " & DaPfa
    End If

    ' Get document date
    BuDatVal = RST.Fields("Datum").Value
    If Not IsNull(BuDatVal) And IsDate(BuDatVal) Then
        BuDat = CDate(BuDatVal)
    Else
        BuDat = Date
    End If

    ' Determine if we have/need a document
    If Len(DaNam) = 0 Then
        ' No document filename
        If IsRevenue Then
            ' Revenue without document - generate default filename (like basData.bas)
            If Len(ReStr) > 0 Then
                DaNam = "Rechnung_Beleg_" & SanitizeFileName(ReStr) & ".pdf"
            Else
                DaNam = "Rechnung_Beleg_" & SanitizeFileName(FormattedGUID) & ".pdf"
            End If
        Else
            ' Expense without document filename
            If Not Config.UseLedgerXML Then
                ' Option A: Skip XML entry but posting is still in CSV
                ProcessDocumentForXML = vbNullString
                Exit Function
            Else
                ' Option B: Generate default filename for ledger.xml linking
                DaNam = "Beleg_" & SanitizeFileName(FormattedGUID) & ".pdf"
                If GlLog = True Then SLogi "  >>> Option B: Generated default filename for expense: " & DaNam
            End If
        End If
    End If

    ' Determine source path
    Dim SourceFound As Boolean
    SourceFound = False
    
    If Len(DaPfa) > 0 Then
        ' Full path provided in DB - verify it exists
        Dim TryPath As String
        If Right$(DaPfa, 1) <> "\" Then
            TryPath = DaPfa & "\" & DaNam
        Else
            TryPath = DaPfa & DaNam
        End If
        
        If m_clFil.FilVor(TryPath) Then
            DaPfa = TryPath
            SourceFound = True
        End If
    End If
    
    If Not SourceFound Then
        ' Not found in DB path or no DB path - check common document locations
        Dim FoundPath As String
        FoundPath = FindDocumentPath(DaNam, Config.MandantNr)
        If Len(FoundPath) > 0 Then
            DaPfa = FoundPath
            SourceFound = True
        End If
    End If

    ' Check if source document exists
    DocExists = SourceFound

    ' Debug: Show search result for expenses
    If GlLog = True And Not IsRevenue Then
        SLogi "  FindDocumentPath result = " & DaPfa
        SLogi "  DocExists = " & DocExists
    End If

    ' Skip if document does not exist - ensures XML matches CSV Beleglinks
    ' Option A: Expenses need existing documents, revenues are generated during export
    ' Option B: Allow XML entry even without physical files (data is in ledger.xml)
    If Not DocExists Then
        If Not Config.UseLedgerXML Then
            ' Option A: Require physical documents for EXPENSES only
            ' Revenues/invoices are generated during export, so allow them to proceed
            If Not IsRevenue Then
                If GlLog = True Then SLogi "  >>> Skipped (expense document not found, no XML entry)"
                ProcessDocumentForXML = vbNullString
                Exit Function
            Else
                If GlLog = True Then SLogi "  >>> Option A: Revenue document will be generated (not yet existing)"
            End If
        Else
            ' Option B: Continue without physical file (ledger.xml contains the data)
            If GlLog = True Then SLogi "  >>> Option B: Generating XML entry without physical file (data in ledger.xml)"
        End If
    End If

    ' Document exists - prepare for copy
    TargetFileName = SanitizeFileName(DaNam)

    ' Ensure target filename is valid
    If Len(TargetFileName) = 0 Then
        TargetFileName = "Rechnung_Beleg_" & SanitizeFileName(FormattedGUID) & ".pdf"
    End If

    ' For revenues (generated PDFs), ensure .pdf extension
    ' For expenses, keep original file extension (can be JPG, PNG, TIFF, DOC, etc.)
    If IsRevenue Then
        If LCase$(Right$(TargetFileName, 4)) <> ".pdf" Then
            TargetFileName = TargetFileName & ".pdf"
        End If
        ' Re-sanitize if we added extension to ensure length limit
        If Len(TargetFileName) > MAX_FILENAME_LENGTH Then
            TargetFileName = SanitizeFileName(TargetFileName)
        End If
    End If

    ' Determine if we use the Ledger format based on export option
    ' Option A (UseLedgerXML=False): Simple document.xml with File extensions + BEDI filename for CSV linking
    ' Option B (UseLedgerXML=True): document.xml with accountsReceivableLedger extension + ledger.xml
    Dim UseLedgerFormat As Boolean
    Dim CleanGUID As String
    Dim LedgerFileName As String

    UseLedgerFormat = Config.UseLedgerXML  ' Option B uses extended Ledger format in document.xml

    ' Create clean GUID (no dashes, uppercase) for BEDI filename
    CleanGUID = Replace(FormattedGUID, "-", vbNullString)
    CleanGUID = UCase$(CleanGUID)

    ' Filename fuer document.xml:
    ' Alle Belege (Einnahmen und Ausgaben) werden zu BEDI<GUID>.<ext> umbenannt
    ' Dies ist DATEV-konform und ermoeglicht konsistente Belegverknuepfung
    LedgerFileName = BELEGLINK_PREFIX & CleanGUID & GetFileExtension(TargetFileName)

    ' Build target path
    ' All documents go to DocumentsDir
    TargetPath = DocumentsDir & LedgerFileName

    ' Copy document if source exists and ExportDocuments is enabled
    If DocExists And Config.ExportDocuments Then
        ' Debug: Show copy attempt for expenses
        If GlLog = True And Not IsRevenue Then
            SLogi "  Copying expense document:"
            SLogi "    Source: " & DaPfa
            SLogi "    Target: " & TargetPath
        End If
        If Not CopyDocumentToExport(DaPfa, TargetPath) Then
            ' Copy failed - log but continue
            LogError "ProcessDocumentForXML", 3002, "Beleg konnte nicht kopiert werden: " & DaPfa
            If GlLog = True Then SLogi "  >>> Copy FAILED"
            ' For revenues, still create XML entry (document should exist)
            ' For expenses, skip
            If Not IsRevenue Then
                ProcessDocumentForXML = vbNullString
                Exit Function
            End If
        Else
            ' Successfully copied - register to ZIP with relative path
            AddToZipList LedgerFileName
            If GlLog = True And Not IsRevenue Then
                SLogi "  >>> Expense document copied and added to ZIP: " & LedgerFileName
            End If
        End If
    End If

    ' Mark document as processed
    ProcessedDocs.Add FormattedGUID, FormattedGUID

    ' Build XML document element - choose format based on export option
    If UseLedgerFormat Then
        ' Option B: Ledger format with accountsReceivableLedger/accountsPayableLedger extension
        ' This format includes datafile reference for structured data import
        DocXML = BuildLedgerDocumentXMLElement(FormattedGUID, CleanGUID, LedgerFileName, BuDat, IsRevenue)
        If GlLog = True Then
            SLogi "  >>> Generated Ledger format XML element (Option B)"
            SLogi "      GUID=" & FormattedGUID & ", File=" & LedgerFileName
        End If
    Else
        ' Option A: Simple document format with File extension and GUID attribute
        ' BEDI filename enables CSV Beleglink matching
        DocXML = BuildDocumentXMLElement(FormattedGUID, LedgerFileName, BuDat, ReStr, Komme)
        If GlLog = True Then
            SLogi "  >>> Generated simple format XML element (Option A)"
            SLogi "      GUID=" & FormattedGUID & ", File=" & LedgerFileName
        End If
    End If

    ProcessDocumentForXML = DocXML
    Exit Function

ErrHandler:
    LogError "ProcessDocumentForXML", Err.Number, Err.Description
    ProcessDocumentForXML = vbNullString
End Function

Private Function BuildDocumentXMLElement(ByVal guid As String, _
                                         ByVal FileName As String, _
                                         ByVal DocDate As Date, _
                                         ByVal InvoiceNumber As String, _
                                         ByVal Description As String) As String
    Dim XML As String
    Dim EscFileName As String
    Dim DateProperty As String
    Dim HasProperties As Boolean

    ' Escape XML special characters for filename
    EscFileName = EscapeXML(FileName)

    ' Format date as YYYY-MM for property key="1" (Buchungsperiode)
    DateProperty = Format$(DocDate, "yyyy-mm")

    ' Determine if we have properties to add
    ' Property key="1" (Buchungsperiode) is always useful
    ' Property key="2" (Rechnungsnummer) only if provided and valid
    HasProperties = True

    ' Build document element per DATEV v06.0 schema (Option A: Dokumentenarchivierung)
    ' Uses guid attribute for document identification (links to CSV Beleglink via BEDI prefix in filename)
    ' Uses xsi:type="File" with name attribute for PDF files
    ' Property keys: 1=Buchungsperiode (YYYY-MM), 2=Rechnungsnummer/Belegnummer
    ' Path is relative to document.xml (flat structure for compatibility)
    XML = "    <document guid=""" & guid & """>" & vbCrLf

    If HasProperties Then
        ' Extension with child property elements (not self-closing)
        XML = XML & "      <extension xsi:type=""File"" name=""" & EscFileName & """>" & vbCrLf

        ' Property key="1": Buchungsperiode (always set)
        XML = XML & "        <property value=""" & DateProperty & """ key=""1""/>" & vbCrLf

        ' Property key="2": Rechnungsnummer (only if valid - max 12 chars, sanitized)
        ' Applies to both Debitoren (Einnahmen) and Kreditoren (Ausgaben)
        If Len(InvoiceNumber) > 0 Then
            XML = XML & "        <property value=""" & EscapeXML(InvoiceNumber) & """ key=""2""/>" & vbCrLf
        End If

        XML = XML & "      </extension>" & vbCrLf
    Else
        ' Self-closing extension (no properties)
        XML = XML & "      <extension xsi:type=""File"" name=""" & EscFileName & """/>" & vbCrLf
    End If

    XML = XML & "    </document>" & vbCrLf

    BuildDocumentXMLElement = XML
End Function

'--------------------------------------------------------------------------------
' BuildLedgerDocumentXMLElement - Build XML element for structured document linking
'--------------------------------------------------------------------------------
' Purpose:     Creates XML document element with cashLedger extension for Option B
'              This format includes datafile reference to ledger.xml for structured data import
'              Uses cashLedger for mixed Einnahmen/Ausgaben (Kassenbuch format)
'
' Parameters:  FormattedGUID  - Document GUID with dashes (for guid attribute)
'              CleanGUID      - Document GUID without dashes (for BEDI filename)
'              BelegFileName  - Full BEDI filename (e.g., BEDI123...456.pdf)
'              DocDate        - Document date (for property key="1")
'              IsRevenue      - (unused, kept for compatibility)
'
' Returns:     XML string for one document element with cashLedger extension and datafile
'--------------------------------------------------------------------------------
Private Function BuildLedgerDocumentXMLElement(ByVal FormattedGUID As String, _
                                               ByVal CleanGUID As String, _
                                               ByVal BelegFileName As String, _
                                               ByVal DocDate As Date, _
                                               ByVal IsRevenue As Boolean) As String
    Dim XML As String
    Dim EscFileName As String
    Dim DateProperty As String

    ' Format date as YYYY-MM for property key="1" (Buchungsperiode)
    DateProperty = Format$(DocDate, "yyyy-mm")

    ' Escape XML special characters for filename
    EscFileName = EscapeXML(BelegFileName)

    ' Build document element with cashLedger extension (Option B format)
    ' Per DATEV XSD: extension with datafile attribute references the ledger.xml
    ' guid attribute enables document identification
    ' Using cashLedger for mixed Einnahmen/Ausgaben (xsd:choice only allows ONE type)
    XML = "    <document guid=""" & FormattedGUID & """>" & vbCrLf

    ' cashLedger extension with datafile - links to ledger.xml for structured booking data
    ' property key="1" = Buchungsperiode (YYYY-MM)
    ' property key="3" = Kassenbezeichnung (e.g. "Kasse")
    XML = XML & "      <extension xsi:type=""" & XML_EXT_CASH_LEDGER & """ datafile=""ledger.xml"">" & vbCrLf
    XML = XML & "        <property value=""" & DateProperty & """ key=""1""/>" & vbCrLf
    XML = XML & "        <property value=""" & XML_FOLDER_CASH & """ key=""3""/>" & vbCrLf
    XML = XML & "      </extension>" & vbCrLf

    ' File extension - references the actual document file (PDF)
    XML = XML & "      <extension xsi:type=""" & XML_EXT_FILE & """ name=""" & EscFileName & """/>" & vbCrLf

    XML = XML & "    </document>" & vbCrLf

    BuildLedgerDocumentXMLElement = XML
End Function

'--------------------------------------------------------------------------------
' GetFileExtension - Extract file extension including the dot
'--------------------------------------------------------------------------------
Private Function GetFileExtension(ByVal FileName As String) As String
    Dim DotPos As Long

    DotPos = InStrRev(FileName, ".")
    If DotPos > 0 Then
        GetFileExtension = Mid$(FileName, DotPos)
    Else
        GetFileExtension = ".pdf"  ' Default extension
    End If
End Function

'================================================================================
' GenerateLedgerXML - Generate ledger.xml for Option B (strukturierte Belegsatzdaten)
'--------------------------------------------------------------------------------
' Purpose:     Creates a separate ledger.xml file with structured booking data
'              in the DATEV LedgerImport format. This file is referenced by
'              the document.xml via the datafile attribute.
'
' Parameters:  RST         - Recordset with booking data
'              Config      - Export configuration
'
' Returns:     Full path to generated ledger.xml, or empty string on failure
'
' Note:        This function is ONLY called for Option B exports.
'              Option A uses only document.xml without ledger.xml.
'================================================================================
Private Function GenerateLedgerXML(ByRef RST As ADODB.Recordset, _
                                   ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler

    Dim XMLContent As String
    Dim LedgerFilePath As String
    Dim TotalAmount As Currency
    Dim TmpAmount As Currency
    Dim TmpBuTyp As Integer
    Dim RecordCount As Long
    Dim CurrentRecord As Long
    Dim ConsolidatedDate As Date
    Dim BookingXML As String

    If GlLog = True Then SLogi "DATEV: GenerateLedgerXML gestartet"
    DoEvents

    ' Validate recordset
    If RST Is Nothing Then
        If GlLog = True Then SLogi "DATEV: RST Is Nothing!"
        LogError "GenerateLedgerXML", 3001, "Recordset ist Nothing"
        GenerateLedgerXML = vbNullString
        Exit Function
    End If

    If RST.State <> 1 Then  ' adStateOpen = 1
        LogError "GenerateLedgerXML", 3002, "Recordset ist nicht geoeffnet"
        GenerateLedgerXML = vbNullString
        Exit Function
    End If

    If RST.EOF And RST.BOF Then
        LogError "GenerateLedgerXML", 3003, "Recordset ist leer"
        GenerateLedgerXML = vbNullString
        Exit Function
    End If

    ' Initialize
    TotalAmount = 0
    RecordCount = RST.RecordCount
    ConsolidatedDate = Config.DateTo  ' Use end date as consolidated date
    If ConsolidatedDate = 0 Then ConsolidatedDate = Date

    ' Build XML header for LedgerImport
    XMLContent = BuildLedgerXMLHeader(Config, ConsolidatedDate)

    ' First pass: calculate total amount for consolidate element
    ' For cashLedger: positive = Einnahme (BuTyp=2), negative = Ausgabe (BuTyp=1)
    RST.MoveFirst
    Do While Not RST.EOF
        If m_InvMod Then
            TmpBuTyp = 2
        Else
            TmpBuTyp = SafeInt(RST.Fields("IDA").Value)
        End If
        TmpAmount = DetermineAmountValue(RST, TmpBuTyp)
        If TmpAmount >= 0.01 Then  ' Only count valid amounts
            If TmpBuTyp = 2 Then
                ' Einnahme: positive
                TotalAmount = TotalAmount + TmpAmount
            Else
                ' Ausgabe: negative
                TotalAmount = TotalAmount - TmpAmount
            End If
        End If
        RST.MoveNext
    Loop

    ' Add consolidate opening tag with totals
    XMLContent = XMLContent & "  <consolidate consolidatedAmount=""" & FormatLedgerAmount(TotalAmount) & """" & vbCrLf
    XMLContent = XMLContent & "               consolidatedDate=""" & Format$(ConsolidatedDate, "yyyy-mm-dd") & """" & vbCrLf
    XMLContent = XMLContent & "               consolidatedCurrencyCode=""" & DATEV_CURRENCY & """>" & vbCrLf

    ' Second pass: generate booking records
    RST.MoveFirst
    CurrentRecord = 0

    Do While Not RST.EOF
        CurrentRecord = CurrentRecord + 1

        ' Build individual booking record
        BookingXML = BuildLedgerBookingElement(RST, Config)
        If Len(BookingXML) > 0 Then
            XMLContent = XMLContent & BookingXML
        End If

        RST.MoveNext
    Loop

    ' Close consolidate and LedgerImport
    XMLContent = XMLContent & "  </consolidate>" & vbCrLf
    XMLContent = XMLContent & "</LedgerImport>" & vbCrLf

    ' Generate filename
    LedgerFilePath = Config.ExportPath & "ledger.xml"
    If GlLog = True Then SLogi "DATEV: Schreibe ledger.xml nach: " & LedgerFilePath
    If GlLog = True Then SLogi "DATEV: XML Loenge: " & Len(XMLContent) & " Zeichen"
    DoEvents

    ' Write file
    If WriteLedgerXMLFile(LedgerFilePath, XMLContent) Then
        If GlLog = True Then SLogi "DATEV: ledger.xml erfolgreich geschrieben!"
        DoEvents
        ' Add to ZIP file list
        AddToZipList "ledger.xml"
        GenerateLedgerXML = LedgerFilePath
    Else
        If GlLog = True Then SLogi "DATEV: WriteLedgerXMLFile fehlgeschlagen!"
        DoEvents
        GenerateLedgerXML = vbNullString
    End If

    Exit Function

ErrHandler:
    If GlLog = True Then SLogi "DATEV: GenerateLedgerXML Fehler: " & Err.Number & " - " & Err.Description
    DoEvents
    LogError "GenerateLedgerXML", Err.Number, Err.Description
    GenerateLedgerXML = vbNullString
End Function

'--------------------------------------------------------------------------------
' GenerateLedgerXMLFromRC - Generate ledger.xml for ReportControl Recordset
'--------------------------------------------------------------------------------
' Purpose:     Creates ledger.xml using Column Caption field names
'              (Mandant, Sachkonto, Belegzeichen, etc. statt IDT, IDK, RechNr)
'--------------------------------------------------------------------------------
Private Function GenerateLedgerXMLFromRC(ByRef RST As ADODB.Recordset, _
                                         ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler

    Dim XMLContent As String
    Dim BookingXML As String
    Dim LedgerFilePath As String
    Dim ConsolidatedDate As Date
    Dim CurrentRecord As Long
    Dim BuDatVal As Variant

    If GlLog = True Then SLogi "DATEV: GenerateLedgerXMLFromRC gestartet"
    DoEvents

    ' Validate recordset
    If RST Is Nothing Then
        LogError "GenerateLedgerXMLFromRC", 3001, "Recordset ist Nothing"
        GenerateLedgerXMLFromRC = vbNullString
        Exit Function
    End If
    If RST.State <> adStateOpen Then
        LogError "GenerateLedgerXMLFromRC", 3002, "Recordset ist nicht geoeffnet"
        GenerateLedgerXMLFromRC = vbNullString
        Exit Function
    End If
    If RST.EOF And RST.BOF Then
        LogError "GenerateLedgerXMLFromRC", 3003, "Recordset ist leer"
        GenerateLedgerXMLFromRC = vbNullString
        Exit Function
    End If

    ' Get consolidated date from first record
    RST.MoveFirst
    BuDatVal = RST.Fields("Datum").Value
    If Not IsNull(BuDatVal) And IsDate(BuDatVal) Then
        ConsolidatedDate = CDate(BuDatVal)
    Else
        ConsolidatedDate = Date
    End If

    ' Build XML header
    XMLContent = BuildLedgerXMLHeader(Config, ConsolidatedDate)

    ' Generate booking records
    RST.MoveFirst
    CurrentRecord = 0

    Do While Not RST.EOF
        CurrentRecord = CurrentRecord + 1

        ' Build individual booking record using RC field names
        BookingXML = BuildLedgerBookingElementFromRC(RST, Config)
        If Len(BookingXML) > 0 Then
            XMLContent = XMLContent & BookingXML
        End If

        RST.MoveNext
    Loop

    ' Close consolidate and LedgerImport
    XMLContent = XMLContent & "  </consolidate>" & vbCrLf
    XMLContent = XMLContent & "</LedgerImport>" & vbCrLf

    ' Generate filename
    LedgerFilePath = Config.ExportPath & "ledger.xml"
    If GlLog = True Then SLogi "DATEV: Schreibe ledger.xml nach: " & LedgerFilePath
    DoEvents

    ' Write file
    If WriteLedgerXMLFile(LedgerFilePath, XMLContent) Then
        If GlLog = True Then SLogi "DATEV: ledger.xml erfolgreich geschrieben!"
        DoEvents
        ' Add to ZIP file list
        AddToZipList "ledger.xml"
        GenerateLedgerXMLFromRC = LedgerFilePath
    Else
        If GlLog = True Then SLogi "DATEV: WriteLedgerXMLFile fehlgeschlagen!"
        DoEvents
        GenerateLedgerXMLFromRC = vbNullString
    End If

    Exit Function

ErrHandler:
    If GlLog = True Then SLogi "DATEV: GenerateLedgerXMLFromRC Fehler: " & Err.Number & " - " & Err.Description
    DoEvents
    LogError "GenerateLedgerXMLFromRC", Err.Number, Err.Description
    GenerateLedgerXMLFromRC = vbNullString
End Function

'--------------------------------------------------------------------------------
' BuildLedgerBookingElementFromRC - Build booking element using RC field names
'--------------------------------------------------------------------------------
Private Function BuildLedgerBookingElementFromRC(ByRef RST As ADODB.Recordset, _
                                                 ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler

    Dim XML As String
    Dim BuTyp As Integer
    Dim IsRevenue As Boolean
    Dim BuDat As Date
    Dim Betrag As Currency
    Dim Konto As Long
    Dim BuSchl As String
    Dim RechNr As String
    Dim BuText As String
    Dim guid As String
    Dim BuDatVal As Variant
    Dim Steue As Single
    Dim Kennz As String

    ' Get booking type (invoice mode: always revenue)
    If m_InvMod Then
        BuTyp = 2
    Else
        BuTyp = SafeIntField(RST, "IDA")
    End If
    IsRevenue = (BuTyp = 2)  ' Type 2 = Einnahme (Revenue)

    ' Get booking date - Datum field exists in both formats
    BuDatVal = RST.Fields("Datum").Value
    If Not IsNull(BuDatVal) And IsDate(BuDatVal) Then
        BuDat = CDate(BuDatVal)
    Else
        BuDat = Date
    End If

    ' Get amount using RC helper function
    Betrag = DetermineAmountValueFromRC(RST, BuTyp)
    If Betrag < 0.01 Then
        ' Skip zero amounts
        BuildLedgerBookingElementFromRC = vbNullString
        Exit Function
    End If

    ' Get account number (invoice mode: use default revenue account)
    If m_InvMod Then
        Konto = GlSE2
    Else
        Konto = SafeLongField(RST, "Sachkonto")
    End If

    ' Get tax key - Steuer field exists in both formats
    Steue = SafeSingleField(RST, "Steuer")
    If IsRevenue Then
        Kennz = "H"
    Else
        Kennz = "S"
    End If
    BuSchl = GetTaxKey(Steue, Kennz)

    ' Get invoice number (invoice mode: use "Rechnung" caption)
    If m_InvMod Then
        RechNr = SanitizeTextField(SafeStringField(RST, "Rechnung"), MAX_BELEGFELD1_LENGTH)
    Else
        RechNr = SanitizeTextField(SafeStringField(RST, "Belegzeichen"), MAX_BELEGFELD1_LENGTH)
    End If

    ' Get booking text (invoice mode: build from Patient + Rechnung)
    If m_InvMod Then
        BuText = SafeStringField(RST, "Patient")
        If Len(BuText) > 0 And Len(RechNr) > 0 Then
            BuText = BuText & " Rech." & RechNr
        ElseIf Len(RechNr) > 0 Then
            BuText = "Rech." & RechNr
        End If
    Else
        BuText = SafeStringField(RST, "Buchungstext")
    End If
    If Len(BuText) = 0 Then
        ' Fallback to Kommentar if exists
        If HasField(RST, "Kommentar") Then
            BuText = SafeStringField(RST, "Kommentar")
        End If
    End If
    BuText = SanitizeTextField(BuText, MAX_BOOKING_TEXT_LENGTH)
    ' Fallback: bookingText is REQUIRED
    If Len(BuText) = 0 Then
        If IsRevenue Then
            BuText = "Einnahme " & Format$(BuDat, "dd.mm.yyyy")
        Else
            BuText = "Ausgabe " & Format$(BuDat, "dd.mm.yyyy")
        End If
    End If

    ' Get GUID - GuiID field exists in both formats
    guid = SafeStringField(RST, "GuiID")

    ' Build XML element - cashLedger format
    XML = "    <cashLedger>" & vbCrLf

    ' Required: date
    XML = XML & "      <date>" & Format$(BuDat, "yyyy-mm-dd") & "</date>" & vbCrLf

    ' Required: amount (always positive, 2 decimals)
    XML = XML & "      <amount>" & FormatLedgerAmount(Betrag) & "</amount>" & vbCrLf

    ' Required: accountNo
    XML = XML & "      <accountNo>" & Format$(Konto, "0") & "</accountNo>" & vbCrLf

    ' Optional: buCode (tax key)
    If Len(BuSchl) > 0 Then
        XML = XML & "      <buCode>" & BuSchl & "</buCode>" & vbCrLf
    End If

    ' Required: currencyCode
    XML = XML & "      <currencyCode>EUR</currencyCode>" & vbCrLf

    ' Optional: invoiceId (Belegfeld1)
    If Len(RechNr) > 0 Then
        XML = XML & "      <invoiceId>" & EscapeXML(RechNr) & "</invoiceId>" & vbCrLf
    End If

    ' Required: bookingText
    XML = XML & "      <bookingText>" & EscapeXML(BuText) & "</bookingText>" & vbCrLf

    ' Optional: belegLink (GUID reference)
    If Len(guid) > 0 Then
        XML = XML & "      <belegLink>" & DATEV_FormatGUIDForXML(guid) & "</belegLink>" & vbCrLf
    End If

    XML = XML & "    </cashLedger>" & vbCrLf

    BuildLedgerBookingElementFromRC = XML
    Exit Function

ErrHandler:
    LogError "BuildLedgerBookingElementFromRC", Err.Number, Err.Description
    BuildLedgerBookingElementFromRC = vbNullString
End Function

'--------------------------------------------------------------------------------
' DetermineAmountValueFromRC - Get amount value using RC field names
'--------------------------------------------------------------------------------
Private Function DetermineAmountValueFromRC(ByRef RST As ADODB.Recordset, _
                                            ByVal BuTyp As Integer) As Currency
On Error Resume Next
    Dim Amount As Currency
    Amount = 0

    ' Invoice mode: read Betrag field directly
    If m_InvMod Then
        If HasField(RST, "Betrag") Then
            Amount = Abs(CCur(RST.Fields("Betrag").Value))
        End If
        DetermineAmountValueFromRC = Amount
        Exit Function
    End If

    ' Try Einnahme/Ausgabe fields first
    If BuTyp = 2 Then  ' Einnahme
        If HasField(RST, "Einnahme") Then
            Amount = CCur(RST.Fields("Einnahme").Value)
        End If
        If Amount = 0 And HasField(RST, "Ausgabe") Then
            Amount = CCur(RST.Fields("Ausgabe").Value)
        End If
    Else  ' Ausgabe
        If HasField(RST, "Ausgabe") Then
            Amount = CCur(RST.Fields("Ausgabe").Value)
        End If
        If Amount = 0 And HasField(RST, "Einnahme") Then
            Amount = CCur(RST.Fields("Einnahme").Value)
        End If
    End If

    ' Fallback to Saldo if available
    If Amount = 0 And HasField(RST, "Saldo") Then
        Amount = Abs(CCur(RST.Fields("Saldo").Value))
    End If

    DetermineAmountValueFromRC = Abs(Amount)
End Function

'--------------------------------------------------------------------------------
' BuildLedgerXMLHeader - Build XML header for LedgerImport format
'--------------------------------------------------------------------------------
Private Function BuildLedgerXMLHeader(ByRef Config As DATEV_ExportConfig, _
                                      ByVal ConsolidatedDate As Date) As String
    Dim Header As String

    Header = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    Header = Header & "<LedgerImport xmlns=""" & LEDGER_XML_NAMESPACE & """" & vbCrLf
    Header = Header & "              xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & vbCrLf
    Header = Header & "              xsi:schemaLocation=""" & LEDGER_XML_SCHEMA_LOCATION & """" & vbCrLf
    Header = Header & "              version=""" & LEDGER_XML_VERSION & """" & vbCrLf
    Header = Header & "              generator_info=""" & DATEV_GENERATING_SYSTEM & """" & vbCrLf
    Header = Header & "              xml_data=""" & LEDGER_XML_DATA_TEXT & """>" & vbCrLf

    BuildLedgerXMLHeader = Header
End Function

'--------------------------------------------------------------------------------
' BuildLedgerBookingElement - Build single booking element for ledger.xml
'--------------------------------------------------------------------------------
' Uses cashLedger format for all bookings (Kassenbuch).
' This allows mixed Einnahmen and Ausgaben in one consolidate element.
' Element order per XSD: date, amount, accountNo, buCode, currencyCode, invoiceId, bookingText
' Note: cashLedger has no bpAccountNo (that's only in accountsReceivable/Payable)
'--------------------------------------------------------------------------------
Private Function BuildLedgerBookingElement(ByRef RST As ADODB.Recordset, _
                                           ByRef Config As DATEV_ExportConfig) As String
On Error GoTo ErrHandler

    Dim XML As String
    Dim BuTyp As Integer
    Dim IsRevenue As Boolean
    Dim BuDat As Date
    Dim Betrag As Currency
    Dim Konto As Long
    Dim BuSchl As String
    Dim RechNr As String
    Dim BuText As String
    Dim guid As String
    Dim BuDatVal As Variant
    Dim Steue As Single
    Dim Kennz As String

    ' Get booking type (for determining amount sign)
    ' Debitoren-Modus: immer Einnahme (BuTyp=2)
    If m_InvMod Then
        BuTyp = 2
    Else
        BuTyp = SafeInt(RST.Fields("IDA").Value)
    End If
    IsRevenue = (BuTyp = 2)  ' Type 2 = Einnahme (Revenue)

    ' Get booking date
    BuDatVal = RST.Fields("Datum").Value
    If Not IsNull(BuDatVal) And IsDate(BuDatVal) Then
        BuDat = CDate(BuDatVal)
    Else
        BuDat = Date
    End If

    ' Get amount (always positive in ledger.xml per XSD)
    ' Use DetermineAmountValue which handles both simple (Einnahme/Ausgabe)
    ' and double-entry (Betrag) bookkeeping field structures
    Betrag = DetermineAmountValue(RST, BuTyp)
    If Betrag < 0.01 Then
        ' Skip zero amounts (XSD: amount 0.00 is not allowed)
        BuildLedgerBookingElement = vbNullString
        Exit Function
    End If

    ' Get account number (Gegenkonto/Sachkonto)
    ' Debitoren: Erloskonto als Sachkonto
    If m_InvMod Then
        Konto = GlSE2
    Else
        Konto = SafeLong(RST.Fields("IDK").Value)
    End If

    ' Get tax key from Steuer field
    Steue = SafeSingle(RST.Fields("Steuer").Value)
    If IsRevenue Then
        Kennz = "H"  ' Haben fuer Einnahmen
    Else
        Kennz = "S"  ' Soll fuer Ausgaben
    End If
    BuSchl = GetTaxKey(Steue, Kennz)

    ' Get invoice number (Belegfeld1) - optional for cashLedger
    RechNr = SanitizeTextField(SafeString(RST.Fields("RechNr").Value), MAX_BELEGFELD1_LENGTH)

    ' Get booking text - REQUIRED for cashLedger!
    BuText = SanitizeTextField(SafeString(RST.Fields("Kommentar").Value), MAX_BOOKING_TEXT_LENGTH)
    If Len(BuText) = 0 Then
        ' Debitoren: IDKurz + RechNr statt Buchtext
        If m_InvMod Then
            BuText = SafeString(RST.Fields("IDKurz").Value)
            If Len(BuText) > 0 And Len(RechNr) > 0 Then
                BuText = BuText & " Rech." & RechNr
            End If
        Else
            BuText = SafeString(RST.Fields("Buchtext").Value)
        End If
        BuText = SanitizeTextField(BuText, MAX_BOOKING_TEXT_LENGTH)
    End If
    ' Fallback: bookingText is REQUIRED, generate if still empty
    If Len(BuText) = 0 Then
        If IsRevenue Then
            BuText = "Einnahme " & Format$(BuDat, "dd.mm.yyyy")
        Else
            BuText = "Ausgabe " & Format$(BuDat, "dd.mm.yyyy")
        End If
    End If

    ' Get GUID for belegLink
    guid = SafeString(RST.Fields("GuiID").Value)

    ' Build XML element - cashLedger format
    ' Element order per XSD: date, amount, accountNo, buCode, currencyCode, invoiceId, bookingText
    XML = "    <cashLedger>" & vbCrLf

    ' Required: date
    XML = XML & "      <date>" & Format$(BuDat, "yyyy-mm-dd") & "</date>" & vbCrLf

    ' Required: amount (positive = Einnahme, negative = Ausgabe in Kassenbuch)
    ' Per XSD: positive = Einnahme, negative = Ausgabe
    If IsRevenue Then
        XML = XML & "      <amount>" & FormatLedgerAmount(Betrag) & "</amount>" & vbCrLf
    Else
        XML = XML & "      <amount>" & FormatLedgerAmount(-Betrag) & "</amount>" & vbCrLf
    End If

    ' Optional: accountNo (Gegenkonto) - only if > 0
    If Konto > 0 Then
        XML = XML & "      <accountNo>" & Format$(Konto, "0") & "</accountNo>" & vbCrLf
    End If

    ' Optional: buCode (BU-Schluessel) - only if valid
    If Len(BuSchl) > 0 And BuSchl <> "0" Then
        XML = XML & "      <buCode>" & EscapeXML(BuSchl) & "</buCode>" & vbCrLf
    End If

    ' Required: currencyCode (fixed EUR for cashLedger)
    XML = XML & "      <currencyCode>EUR</currencyCode>" & vbCrLf

    ' Optional: invoiceId (Rechnungsnummer) - optional for cashLedger
    If Len(RechNr) > 0 Then
        XML = XML & "      <invoiceId>" & EscapeXML(RechNr) & "</invoiceId>" & vbCrLf
    End If

    ' Required: bookingText (Belegtext)
    XML = XML & "      <bookingText>" & EscapeXML(BuText) & "</bookingText>" & vbCrLf

    ' Close element
    XML = XML & "    </cashLedger>" & vbCrLf

    BuildLedgerBookingElement = XML
    Exit Function

ErrHandler:
    LogError "BuildLedgerBookingElement", Err.Number, Err.Description
    BuildLedgerBookingElement = vbNullString
End Function

'--------------------------------------------------------------------------------
' FormatLedgerAmount - Format currency amount for ledger.xml (always with 2 decimals)
'--------------------------------------------------------------------------------
Private Function FormatLedgerAmount(ByVal Amount As Currency) As String
    ' Format as decimal with exactly 2 decimal places, using dot as separator
    FormatLedgerAmount = Format$(Amount, "0.00")
    ' Ensure dot is used as decimal separator (not comma)
    FormatLedgerAmount = Replace(FormatLedgerAmount, ",", ".")
End Function

'--------------------------------------------------------------------------------
' WriteLedgerXMLFile - Write ledger.xml content to file with UTF-8 encoding
'--------------------------------------------------------------------------------
Private Function WriteLedgerXMLFile(ByVal FilePath As String, ByVal Content As String) As Boolean
On Error GoTo ErrHandler

    Dim fso As Object
    Dim TextStream As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create file with UTF-8 encoding (using ADODB.Stream for proper UTF-8)
    Dim ADOStream As Object
    Set ADOStream = CreateObject("ADODB.Stream")

    With ADOStream
        .Type = 2  ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText Content
        .SaveToFile FilePath, 2  ' adSaveCreateOverWrite
        .Close
    End With

    Set ADOStream = Nothing
    Set fso = Nothing

    WriteLedgerXMLFile = True
    Exit Function

ErrHandler:
    LogError "WriteLedgerXMLFile", Err.Number, Err.Description
    WriteLedgerXMLFile = False
End Function

Private Function EscapeXML(ByVal Text As String) As String
    ' Escape XML special characters
    If Len(Text) = 0 Then
        EscapeXML = vbNullString
        Exit Function
    End If

    Dim Result As String
    Result = Text

    ' Must escape & first (before other entities that contain &)
    Result = Replace(Result, "&", "&amp;")
    Result = Replace(Result, "<", "&lt;")
    Result = Replace(Result, ">", "&gt;")
    Result = Replace(Result, """", "&quot;")
    Result = Replace(Result, "'", "&apos;")

    ' Remove control characters
    Dim i As Integer
    For i = 0 To 31
        If i <> 9 And i <> 10 And i <> 13 Then ' Keep tab, LF, CR
            Result = Replace(Result, Chr$(i), vbNullString)
        End If
    Next i

    EscapeXML = Result
End Function

Private Function FindDocumentPath(ByVal FileName As String, ByVal MandantNr As Long) As String
    ' Search for document in common locations
    Dim SearchPaths(0 To 5) As String
    Dim i As Integer
    Dim FullPath As String

    ' Build search paths
    SearchPaths(0) = GlBPf                          ' Bilder folder (GlBPf = Bilderpfad from INI)
    SearchPaths(1) = GlDpf & "Belege\"              ' Main document folder
    SearchPaths(2) = GlDpf & "Dokumente\"           ' Documents folder
    SearchPaths(3) = GlDpf & "Bilder\"              ' Images folder (alternative)
    SearchPaths(4) = GlDpf & "PDF\"                 ' PDF folder
    SearchPaths(5) = GlDpf & "Rechnungen\"          ' Invoices folder

    ' Search each location
    For i = 0 To UBound(SearchPaths)
        Dim BasePath As String
        BasePath = SearchPaths(i)
        
        If Len(BasePath) > 0 Then
            ' Ensure trailing backslash
            If Right$(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
            
            FullPath = BasePath & FileName
            If m_clFil.FilVor(FullPath) Then
                FindDocumentPath = FullPath
                Exit Function
            End If
        End If
    Next i

    ' Not found
    FindDocumentPath = vbNullString
End Function

Private Function CopyDocumentToExport(ByVal SourcePath As String, _
                                      ByVal TargetPath As String) As Boolean
On Error GoTo ErrHandler

    ' Check if source exists
    If Not m_clFil.FilVor(SourcePath) Then
        CopyDocumentToExport = False
        Exit Function
    End If

    ' Check if target already exists (from split posting or previous run)
    If m_clFil.FilVor(TargetPath) Then
        CopyDocumentToExport = True
        Exit Function
    End If

    ' Use VB6 FileCopy for synchronous copy with rename
    ' This ensures the file is named correctly (TargetFileName) at destination
    FileCopy SourcePath, TargetPath

    ' Verify copy success
    CopyDocumentToExport = m_clFil.FilVor(TargetPath)
    Exit Function

ErrHandler:
    ' If VB6 FileCopy fails (e.g. open file, permissions), try FSO as fallback
    Err.Clear
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        fso.CopyFile SourcePath, TargetPath, True
        CopyDocumentToExport = m_clFil.FilVor(TargetPath)
        Set fso = Nothing
    Else
        CopyDocumentToExport = False
    End If
    
    If Not CopyDocumentToExport Then
        LogError "CopyDocumentToExport", Err.Number, Err.Description
    End If
End Function

Private Function GenerateXMLFilename(ByRef Config As DATEV_ExportConfig) As String
    Dim FileName As String

    ' Format: document.xml (standard DATEV name)
    FileName = "document.xml"

    GenerateXMLFilename = Config.ExportPath & FileName
End Function

Private Function WriteXMLFile(ByVal FilePath As String, ByVal Content As String) As Boolean
On Error GoTo ErrHandler

    ' Delete existing file first to ensure clean overwrite
    If m_clFil.FilVor(FilePath) Then
        On Error Resume Next
        Kill FilePath
        On Error GoTo ErrHandler
    End If

    ' Use clFil for UTF-8 file write
    m_clFil.FilCnWr FilePath, Content

    ' Verify file was created
    WriteXMLFile = m_clFil.FilVor(FilePath)
    Exit Function

ErrHandler:
    LogError "WriteXMLFile", Err.Number, Err.Description
    WriteXMLFile = False
End Function

'--------------------------------------------------------------------------------
' PRIVATE - ZIP Archive Creation
'--------------------------------------------------------------------------------

Private Function CreateZIPArchive(ByRef Config As DATEV_ExportConfig, _
                                  ByVal CSVFilePath As String, _
                                  ByVal XMLFilePath As String) As String
On Error GoTo ErrHandler

    Dim ZipFilePath As String
    Dim ZipFileName As String
    Dim SrcPath As String

    ' Benutzer-gewaehlten Dateinamen verwenden falls vorhanden
    If Len(Config.ExportFileName) > 0 Then
        ZipFileName = Config.ExportFileName & ".zip"
    Else
        ' Fallback: Generate ZIP filename with Option A/B distinction
        ' Option A (UseLedgerXML=False): DATEV_Archiv_... (Dokumentenarchivierung)
        ' Option B (UseLedgerXML=True):  DATEV_Ledger_... (Ledger-Integration)
        If Config.UseLedgerXML Then
            ZipFileName = "DATEV_Ledger_" & _
                          Format$(Date, "yyyymmdd") & "_" & _
                          Format$(Time, "hhnnss") & ".zip"
        Else
            ZipFileName = "DATEV_Archiv_" & _
                          Format$(Date, "yyyymmdd") & "_" & _
                          Format$(Time, "hhnnss") & ".zip"
        End If
    End If

    ' ZIP file goes to original export path (not in subfolder)
    ZipFilePath = m_OrgPfa & ZipFileName

    ' Source path is the subfolder (Config.ExportPath now points to subfolder)
    ' SimpliZip expects a real folder/file path, not a wildcard mask
    SrcPath = Config.ExportPath

    If GlLog = True Then SLogi "=== CreateZIPArchive ==="
    If GlLog = True Then SLogi "  ZIP File: " & ZipFilePath
    If GlLog = True Then SLogi "  Source Folder: " & SrcPath
    DoEvents

    ' Compress folder and delete source via SZipp (basWindow)
    ' SZipp waits synchronously for SimpliZip.exe to finish
    If SZipp(ZipFilePath, SrcPath, True) Then
        If GlLog = True Then SLogi "SZipp erfolgreich beendet"
        If m_clFil.FilVor(ZipFilePath) Then
            If GlLog = True Then SLogi "ZIP-Datei erstellt: " & ZipFilePath
            CreateZIPArchive = ZipFilePath
        Else
            If GlLog = True Then SLogi "SZipp beendet, aber ZIP-Datei nicht gefunden: " & ZipFilePath
            CreateZIPArchive = vbNullString
        End If
    Else
        If GlLog = True Then SLogi "SZipp fehlgeschlagen - Dateien bleiben unkomprimiert"
        CreateZIPArchive = vbNullString
    End If

    Exit Function

ErrHandler:
    If GlLog = True Then SLogi "=== CreateZIPArchive ERROR ==="
    If GlLog = True Then SLogi "  Err.Number: " & Err.Number
    If GlLog = True Then SLogi "  Err.Description: " & Err.Description
    If GlLog = True Then SLogi "  Err.Source: " & Err.Source
    DoEvents
    SPopu "CreateZIPArchive " & Err.Number, Err.Description, IC48_Warning
    CreateZIPArchive = vbNullString
End Function

Private Sub AddToZipList(ByVal FilePath As String)
    m_ZipFileCount = m_ZipFileCount + 1
    ReDim Preserve m_ZipFiles(m_ZipFileCount - 1)
    m_ZipFiles(m_ZipFileCount - 1) = FilePath
End Sub

'--------------------------------------------------------------------------------
' PRIVATE - Dual Progress Dialog Integration
' prbStat1 = Detail progress (current phase)
' prbStat2 = Overall progress (all phases)
'--------------------------------------------------------------------------------

Private Sub InitProgressWithPhases(ByVal Title As String, ByVal TotalPhases As Integer)
On Error Resume Next
    ' Initialize progress dialog with phase tracking
    ' prbStat1: Detail (0-100 for current phase)
    ' prbStat2: Overall (0-100 across all phases)

    m_TotalPhases = TotalPhases
    m_CurrentPhase = 0
    ReDim m_PhaseNames(1 To TotalPhases)

    frmStatus.Show vbModeless
    frmStatus.Caption = Title
    frmStatus.lblLab01.Caption = "Initialisiere..."

    ' prbStat1: Detail progress (current phase) - Standard mode
    With frmStatus.prbStat1
        .Scrolling = xtpProgressBarStandard
        .Min = 0
        .Max = 100
        .Value = 0
    End With

    ' prbStat2: Overall progress (all phases) - Standard mode
    With frmStatus.prbStat2
        .Scrolling = xtpProgressBarStandard
        .Min = 0
        .Max = 100
        .Value = 0
    End With

    ' Reset cancel flag
    frmStatus.txtDummy.Text = "A"

    DoEvents
End Sub

Private Sub StartPhase(ByVal PhaseNumber As Integer, ByVal PhaseName As String, ByVal MaxItems As Long)
On Error Resume Next
    ' Start a new phase - resets prbStat1 and updates prbStat2

    If PhaseNumber < 1 Then PhaseNumber = 1
    If PhaseNumber > m_TotalPhases Then PhaseNumber = m_TotalPhases

    m_CurrentPhase = PhaseNumber
    If PhaseNumber <= UBound(m_PhaseNames) Then
        m_PhaseNames(PhaseNumber) = PhaseName
    End If

    ' Update title and label
    frmStatus.Caption = "DATEV Export - " & PhaseName
    frmStatus.lblLab01.Caption = PhaseName & "..."

    ' Reset detail progress (prbStat1) for new phase
    With frmStatus.prbStat1
        .Min = 0
        .Max = IIf(MaxItems > 0, MaxItems, 100)
        .Value = 0
    End With

    ' Update overall progress (prbStat2) to start of this phase
    Dim OverallPct As Long
    OverallPct = ((m_CurrentPhase - 1) * 100) \ m_TotalPhases
    frmStatus.prbStat2.Value = OverallPct

    DoEvents
End Sub

Private Sub UpdateDualProgress(ByVal CurrentItem As Long, ByVal MaxItems As Long, ByVal StatusText As String)
On Error Resume Next
    ' Update both progress bars
    ' prbStat1: Detail progress within current phase
    ' prbStat2: Overall progress including current phase partial

    Dim DetailPct As Long
    Dim OverallPct As Long
    Dim PhaseContribution As Long

    ' Clamp values
    If CurrentItem < 0 Then CurrentItem = 0
    If MaxItems < 1 Then MaxItems = 1
    If CurrentItem > MaxItems Then CurrentItem = MaxItems

    ' Calculate detail percentage (0-100 within phase)
    DetailPct = (CurrentItem * 100) \ MaxItems

    ' Calculate overall percentage
    ' Each phase contributes (100 / m_TotalPhases) percent
    ' Current phase adds partial progress
    PhaseContribution = 100 \ m_TotalPhases
    OverallPct = ((m_CurrentPhase - 1) * PhaseContribution) + _
                 ((DetailPct * PhaseContribution) \ 100)

    ' Update prbStat1 (detail) - use actual item count for smoother progress
    If frmStatus.prbStat1.Max <> MaxItems Then
        frmStatus.prbStat1.Max = MaxItems
    End If
    frmStatus.prbStat1.Value = CurrentItem

    ' Update prbStat2 (overall) - percentage based
    If OverallPct > 100 Then OverallPct = 100
    frmStatus.prbStat2.Value = OverallPct

    ' Update status text
    frmStatus.lblLab01.Caption = StatusText
End Sub

Private Sub CompletePhase()
On Error Resume Next
    ' Mark current phase as complete (100%)

    ' Set detail to 100%
    frmStatus.prbStat1.Value = frmStatus.prbStat1.Max

    ' Update overall to end of current phase
    Dim OverallPct As Long
    OverallPct = (m_CurrentPhase * 100) \ m_TotalPhases
    If OverallPct > 100 Then OverallPct = 100
    frmStatus.prbStat2.Value = OverallPct

    DoEvents
End Sub

Private Sub ShowProgressDialogInit(ByVal Title As String)
On Error Resume Next
    ' Legacy function - show dialog with Marquee mode (indeterminate)
    ' Used when phase count is unknown

    m_TotalPhases = 1
    m_CurrentPhase = 1

    frmStatus.Show vbModeless
    frmStatus.Caption = Title
    frmStatus.lblLab01.Caption = "Lade Daten..."

    With frmStatus.prbStat1
        .Min = 0
        .Max = 100
        .Value = 0
        .Scrolling = xtpProgressBarMarquee
    End With

    With frmStatus.prbStat2
        .Min = 0
        .Max = 100
        .Value = 0
        .Scrolling = xtpProgressBarMarquee
    End With

    frmStatus.txtDummy.Text = "A"
    DoEvents
End Sub

Private Sub ShowProgressDialog(ByVal Title As String, ByVal MaxValue As Long)
On Error Resume Next
    ' Legacy function - show dialog with known count (single phase)

    If MaxValue < 1 Then MaxValue = 1

    m_TotalPhases = 1
    m_CurrentPhase = 1

    If Not frmStatus.Visible Then
        frmStatus.Show vbModeless
    End If

    frmStatus.Caption = Title
    frmStatus.lblLab01.Caption = "Initialisiere..."

    With frmStatus.prbStat1
        .Scrolling = xtpProgressBarStandard
        .Min = 0
        .Max = MaxValue
        .Value = 0
    End With

    With frmStatus.prbStat2
        .Scrolling = xtpProgressBarStandard
        .Min = 0
        .Max = 100
        .Value = 0
    End With

    frmStatus.txtDummy.Text = "A"
    DoEvents
End Sub

Private Sub ResetProgressForPhase(ByVal PhaseTitle As String, ByVal MaxValue As Long)
On Error Resume Next
    ' Legacy function - reset progress bars for a new phase (used by DATEV_BuEx)

    If MaxValue < 1 Then MaxValue = 1

    frmStatus.Caption = PhaseTitle
    frmStatus.lblLab01.Caption = "Initialisiere Phase..."

    With frmStatus.prbStat1
        .Max = MaxValue
        .Value = 0
    End With

    With frmStatus.prbStat2
        .Max = MaxValue
        .Value = 0
    End With

    DoEvents
End Sub

Private Sub UpdateProgress(ByVal CurrentValue As Long, ByVal MaxValue As Long, ByVal StatusText As String)
On Error Resume Next
    ' Legacy function - update progress (works with single phase or dual)

    If m_TotalPhases > 1 Then
        ' Use dual progress mode
        UpdateDualProgress CurrentValue, MaxValue, StatusText
    Else
        ' Single phase mode - both bars show same progress
        If CurrentValue < 0 Then CurrentValue = 0
        If MaxValue < 1 Then MaxValue = 1
        If CurrentValue > MaxValue Then CurrentValue = MaxValue

        If frmStatus.prbStat1.Max <> MaxValue Then
            frmStatus.prbStat1.Max = MaxValue
        End If

        frmStatus.prbStat1.Value = CurrentValue
        frmStatus.prbStat2.Value = (CurrentValue * 100) \ MaxValue
        frmStatus.lblLab01.Caption = StatusText
    End If
End Sub

Private Sub HideProgressDialog()
On Error Resume Next

    frmStatus.Hide
    Unload frmStatus
    Set frmStatus = Nothing

    m_TotalPhases = 0
    m_CurrentPhase = 0

    DoEvents
End Sub

Private Function CheckCancelled() As Boolean
On Error Resume Next

    If frmStatus.txtDummy.Text = "B" Then
        CheckCancelled = True
    Else
        CheckCancelled = False
    End If
End Function

'--------------------------------------------------------------------------------
' PRIVATE - Validation Helpers
'--------------------------------------------------------------------------------

Private Function ValidateConfig(ByRef Config As DATEV_ExportConfig) As Boolean
    ' Validate required configuration values

    If Config.Beraternummer <= 0 Then
        ValidateConfig = False
        Exit Function
    End If

    If Config.Mandantennummer <= 0 Then
        ValidateConfig = False
        Exit Function
    End If

    If Len(Config.ExportPath) = 0 Then
        ValidateConfig = False
        Exit Function
    End If

    ValidateConfig = True
End Function

Private Function EnsureExportDirectory(ByVal Path As String) As Boolean
On Error GoTo ErrHandler

    ' Use clFil to ensure directory exists
    EnsureExportDirectory = m_clFil.FilDir(Path)
    Exit Function

ErrHandler:
    EnsureExportDirectory = False
End Function

'--------------------------------------------------------------------------------
' PRIVATE - Email Sending (Stub)
'--------------------------------------------------------------------------------

Private Sub SendExportViaEmail(ByRef Config As DATEV_ExportConfig, _
                               ByRef Result As DATEV_ExportResult)
On Error Resume Next

    ' STUB: Email functionality to be implemented based on existing patterns
    ' Will use existing email infrastructure from the codebase
End Sub

'--------------------------------------------------------------------------------
' PRIVATE - ZIP Compression Helpers
'--------------------------------------------------------------------------------

Private Function SetSubFo(ByRef Config As DATEV_ExportConfig) As Boolean
On Error GoTo ErrHdl
' Setup subfolder structure for export
' Modifies Config.ExportPath to point to subfolder
' Stores original path in m_OrgPfa
' Returns: True if subfolder was created successfully

Dim SubNam As String

SetSubFo = False

' Save original export path
m_OrgPfa = Config.ExportPath

' Ensure original path ends with backslash
If Right$(m_OrgPfa, 1) <> "\" Then
    m_OrgPfa = m_OrgPfa & "\"
End If

' Create subfolder name based on export filename or timestamp
If Len(Config.ExportFileName) > 0 Then
    SubNam = Config.ExportFileName
Else
    SubNam = "DATEV_" & Format$(Now, "YYYYMMDD_HHNNSS")
End If

' Store subfolder name
m_SubNam = SubNam

' Update Config.ExportPath to point to subfolder
Config.ExportPath = m_OrgPfa & SubNam & "\"

' Create subfolder directory
If Not EnsureExportDirectory(Config.ExportPath) Then
    If GlLog = True Then SLogi "Konnte Unterordner nicht erstellen: " & Config.ExportPath
    Config.ExportPath = m_OrgPfa  ' Restore original path
    Exit Function
End If

If GlLog = True Then SLogi "Export-Unterordner erstellt: " & Config.ExportPath

SetSubFo = True
Exit Function

ErrHdl:
If GlLog = True Then SLogi "SetSubFo Fehler: " & Err.Number & " - " & Err.Description
Config.ExportPath = m_OrgPfa  ' Restore original path on error
SetSubFo = False

End Function


'--------------------------------------------------------------------------------
' PRIVATE - Error Handling Helpers
'--------------------------------------------------------------------------------

Private Sub LogError(ByVal ProcName As String, ByVal ErrNum As Long, ByVal ErrDesc As String)
    ' Log error and show popup
    If GlLog = True Then SLogi "=== LogError: basDATEV." & ProcName & " ==="
    If GlLog = True Then SLogi "  Err.Number: " & ErrNum
    If GlLog = True Then SLogi "  Err.Description: " & ErrDesc
    DoEvents
    SPopu "basDATEV." & ProcName & " " & ErrNum, ErrDesc, IC48_Warning

    ' Also log to error log if available
    On Error Resume Next
    SErLog ErrDesc & " basDATEV." & ProcName & " " & ErrNum
End Sub

'--------------------------------------------------------------------------------
' Migration Notes for Legacy Callers
'--------------------------------------------------------------------------------
' To migrate from legacy functions to basDATEV:
'
' 1. Replace basDaMa.S_Expor() "csv"/"datev" calls:
'    Old: S_Expor "csv", EmlVe, ManNr, ExKom, ExVer
'    New: DATEV_Expor "A", EmlVe, BelEx   (Dokumentenarchivierung)
'         DATEV_Expor "B", EmlVe, BelEx   (Ledger-Integration)
'         (ManNr removed - derived from recordset field "Mandant")
'         (BelEx = True setzt automatisch ExportDocuments und CompressOutput)
'
' 2. Replace basData.S_BuEx() calls:
'    Old: S_BuEx "csv", Krite, ManNr, EmlVe, ExKom, ExVer
'    New: DATEV_BuEx "A", EmlVe, Krite, BelEx   (Dokumentenarchivierung)
'         DATEV_BuEx "B", EmlVe, Krite, BelEx   (Ledger-Integration)
'         (ManNr removed - already in Krite, derived from recordset)
'         (BelEx = True setzt automatisch ExportDocuments und CompressOutput)
'
' 3. Replace basDatRe.S_DaExF() calls:
'    Old: FormattedAccount = S_DaExF(AccountNo)
'    New: FormattedAccount = DATEV_FormatAccountNumber(AccountNo, GldKt)
'
' 4. Replace Beleglink generation:
'    Old: BeGui = "BEDI" & CleanGUID(GuiID)
'    New: BeGui = DATEV_CreateBeleglink(GuiID)
'
' 5. Replace basDatRe.S_DaExX() XML generation:
'    Old: S_DaExX RST, FiNam, ExTyp, ExNam  (uses XML v4.0)
'    New: XML is generated automatically by DATEV_ExportSelected when
'         Config.ExportDocuments = True (uses XML v6.0)
'
' 6. Replace basDatRe.S_DaExB() and S_DaExP() CSV line building:
'    Old: TmpSt = S_DaExB(RST, "csv", ZeiEx)
'    New: Handled internally by GenerateCSVExport via BuildCSVDataLineOptimized
'
'--------------------------------------------------------------------------------
' LEGACY FUNCTION MAPPING (Features/Outcomes)
'--------------------------------------------------------------------------------
' Legacy Function              -> New basDATEV Function(s)
' ---------------------------     ------------------------------------------
' frmBuExp.FWeit()             -> Caller uses DATEV_ExportSelected/ByDateRange
' basDaMa.S_Expor("csv/datev") -> DATEV_Expor("A"/"B", ...) with DATEV_ExportSelected()
' basData.S_BuEx("csv/datev")  -> DATEV_BuEx("A"/"B", ...) with DATEV_ExportByDateRange()
' basDatRe.S_DaExB()           -> BuildCSVDataLineOptimized() [internal]
' basDatRe.S_DaExP()           -> BuildCSVDataLineOptimized() [internal]
' basDatRe.S_DaExF()           -> DATEV_FormatAccountNumber() [public]
' basDatRe.S_DaExX()           -> GenerateXMLExport() + GenerateLedgerXML() [internal]
'
'--------------------------------------------------------------------------------
' KEY CHANGES FROM LEGACY
'--------------------------------------------------------------------------------
' - XML version upgraded: v4.0 -> v6.0 (current DATEV Unternehmen Online)
' - CSV format: EXTF v700 with 116 fields (compliant with DATEV Proeftool)
' - All file I/O via clsFile (FilCnWr for UTF-8, FilCop for documents)
' - Two-phase progress: Phase 1 = CSV, Phase 2 = XML + PDF documents
' - Split posting deduplication via m_DocumentGUIDs collection
' - Revenue vs Expense handling: revenues require documents, expenses optional
' - Performance: array-based line building, interval-based UI updates
'
'--------------------------------------------------------------------------------
' EXPORT OPTIONS (ExTyp Parameter)
'--------------------------------------------------------------------------------
' Two export options are available via the ExTyp parameter:
'
' OPTION A: Dokumentenarchivierung (Document-XML)
' ------------------------------------------------
'    - ExTyp = "A"
'    - Purpose: Simple document archiving in DATEV Unternehmen Online
'    - Generated files:
'      * EXTF_*.csv (DATEV Buchungsstapel)
'      * document.xml (archive format, namespace: document/v06.0)
'      * BEDI{GUID}.pdf files (document files)
'    - document.xml structure:
'      <archive xmlns="http://xml.datev.de/bedi/tps/document/v06.0">
'        <document guid="{GUID}">
'          <extension xsi:type="File" name="BEDI{GUID}.pdf"/>
'        </document>
'      </archive>
'
' OPTION B: Ledger-Integration (Document-XML + Ledger-XML)
' ---------------------------------------------------------
'    - ExTyp = "B"
'    - Purpose: Automatic document-to-booking linking in DATEV Buchhaltung
'    - Generated files:
'      * EXTF_*.csv (DATEV Buchungsstapel)
'      * document.xml (archive format with datafile reference)
'      * ledger.xml (structured booking data, namespace: ledger/v060)
'      * BEDI{GUID}.pdf files (document files)
'    - document.xml structure (references ledger.xml):
'      <archive xmlns="http://xml.datev.de/bedi/tps/document/v06.0">
'        <document guid="{GUID}">
'          <extension xsi:type="accountsReceivableLedger" datafile="ledger.xml">
'            <property value="YYYY-MM" key="1"/>
'            <property value="Ausgangsrechnungen" key="3"/>
'          </extension>
'          <extension xsi:type="File" name="BEDI{GUID}.pdf"/>
'        </document>
'      </archive>
'    - ledger.xml structure (structured booking data):
'      <LedgerImport xmlns="http://xml.datev.de/bedi/tps/ledger/v060">
'        <consolidate><accountsReceivableLedger>
'          <date>YYYY-MM-DD</date>
'          <amount>123.45</amount>
'          <accountNo>10000</accountNo>
'          <belegLink>BEDI{GUID}</belegLink>
'          ...
'        </accountsReceivableLedger></consolidate>
'      </LedgerImport>
'
' Both options use BEDI{GUID} filename format for CSV "Beleglink" compatibility.
' Option B enables DATEV to automatically link documents to their bookings.
'==================================================================
