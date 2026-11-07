Attribute VB_Name = "basBankimp"
Option Explicit

' ------------------------------------------------------------------------------
' Module:       basBankimp
' Description:  Universal CSV & MT940 Import for Bank Transactions
' ------------------------------------------------------------------------------

' Required Object References
Private clFil As clsFile
Private FM As Form
Private CoDia As XtremeSuiteControls.CommonDialog

Private Declare Function BnkAcquire Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef Prov As Long, ByVal Container As String, ByVal Provider As String, ByVal ProvType As Long, ByVal Flags As Long) As Long
Private Declare Function BnkCreate Lib "advapi32.dll" Alias "CryptCreateHash" (ByVal Prov As Long, ByVal AlgId As Long, ByVal Key As Long, ByVal Flags As Long, ByRef Hash As Long) As Long
Private Declare Function BnkHash Lib "advapi32.dll" Alias "CryptHashData" (ByVal Hash As Long, ByVal DataPtr As Long, ByVal DataLen As Long, ByVal Flags As Long) As Long
Private Declare Function BnkGet Lib "advapi32.dll" Alias "CryptGetHashParam" (ByVal Hash As Long, ByVal Param As Long, ByRef Buffer As Any, ByRef Size As Long, ByVal Flags As Long) As Long
Private Declare Function BnkDestroy Lib "advapi32.dll" Alias "CryptDestroyHash" (ByVal Hash As Long) As Long
Private Declare Function BnkRelease Lib "advapi32.dll" Alias "CryptReleaseContext" (ByVal Prov As Long, ByVal Flags As Long) As Long

' Public Entry Point
Public Sub Imp01(ByVal IdBnk As Long, ByVal ManNr As Long, ByVal MitNr As Long, Optional ByVal PthFl As String = "")
    On Error GoTo LiErr

    Static ImportBusy As Boolean
    If ImportBusy Then Exit Sub
    ImportBusy = True
    Dim NewBank As Boolean, DidImport As Boolean, ChkDup As Boolean
    Dim BankMap() As Long
    Dim FiNam As String
    Dim ImOrd As String
    Dim TmpSt As String
    Dim RS121 As ADODB.Recordset
    
    ' Progress Bar References
    Dim PrBr1 As Object
    Dim PrBr2 As Object
    Dim TxDum As Object
    
    ' Initialize Objects
    Set FM = frmMain
    Set CoDia = FM.comDialo
    Set clFil = New clsFile
    clFil.hwnd = FM.hwnd
    
    ' --------------------------------------------------------------------------
    ' 1. File Selection Logic
    ' --------------------------------------------------------------------------
    If Len(PthFl) > 0 Then
        FiNam = PthFl
    Else
        ' Determine Initial Directory
        If GlRDP = True Then
            If clFil.FilDir(GlIPf) = False Then
                ImOrd = GlDpf & "Import\"
            Else
                ImOrd = GlIPf
            End If
        Else
            If clFil.FilDir(GlImO) = False Then
                ImOrd = GlIPf
            Else
                ImOrd = GlImO
            End If
        End If
        If Right$(ImOrd, 1) <> "\" Then ImOrd = ImOrd & "\"
        
        ' Open Dialog via clsFile
        With clFil
            .hwnd = FM.hwnd
            .StaVe = ImOrd
            .DaTit = "Bitte waehlen Sie die gewuenschten Dateien aus"
            .DaStr = "Unterstuetzte Formate (*.csv;*.txt;*.sta)" & Chr(0) & "*.csv;*.txt;*.sta" & Chr(0) & _
                     "Bankdaten (*.csv;*.sta)" & Chr(0) & "*.csv;*.sta" & Chr(0) & _
                     "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0)
            FiNam = .FilOpn
        End With
        
        If Len(FiNam) = 0 Then
            Set CoDia = Nothing
            Set clFil = Nothing
            GoTo Cleanup
        End If
    End If

    ' Check File Existence
    With clFil
        .FilPfa FiNam
    End With
    
    If clFil.FilVor(FiNam) = False Then
        SMeFr "Bankimport", "Datei nicht gefunden", FiNam, vbNullString, False, 2, True, FM.hwnd
        GoTo Cleanup
    End If

    If GlLog = True Then SLogi "Imp01: Starting import of file: " & FiNam

    ' --------------------------------------------------------------------------
    ' 2. Database Setup
    ' --------------------------------------------------------------------------
    If GlLog = True Then SLogi "Imp01: Opening recordset..."
    Set RS121 = New ADODB.Recordset
    With RS121
        .CursorLocation = adUseClient
        .Source = "qrySimBaNe"
        .ActiveConnection = DB1
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
        .Open Options:=adCmdTableDirect
    End With
    If GlLog = True Then SLogi "Imp01: Recordset opened successfully"

    If Not RS121.Supports(adAddNew) Then
        SMeFr "Bankimport", "Die Datenbank erlaubt kein Hinzufuegen von Umsaetzen.", vbNullString, vbNullString, False, 2, True, FM.hwnd
        GoTo Cleanup
    End If
    
    ' --------------------------------------------------------------------------
    ' 3. Read & Analyze
    ' --------------------------------------------------------------------------
    
    ' Show Status Form
    If GlLog = True Then SLogi "Imp01: Showing status form..."
    frmStatus.Show
    DoEvents
    Set PrBr1 = frmStatus.prbStat1
    Set PrBr2 = frmStatus.prbStat2
    Set TxDum = frmStatus.txtDummy
    frmStatus.Caption = "Universal Import... Analysiere"
    If GlLog = True Then SLogi "Imp01: Status form shown"

    ' Read File
    If GlLog = True Then SLogi "Imp01: Reading file content..."
    With clFil
        .FilPfa FiNam
    End With

    ' --- Robust multi-encoding read (UTF-16 LE/BE w/ or w/o BOM, UTF-8, ANSI) ---
    ' Legacy FilReSt truncates UTF-16 files at the first NUL byte; read & decode here.
    Dim UniDun As Boolean
    UniDun = False
    TmpSt = ReadBankFile(FiNam, UniDun)
    If Len(TmpSt) = 0 And UniDun = False Then TmpSt = clFil.FilReSt
    If GlLog = True Then SLogi "Imp01: File read, length=" & Len(TmpSt)

    If TmpSt = vbNullString Then
        If GlLog = True Then SLogi "Imp01: File is empty, exiting"
        GoTo Cleanup
    End If
    
    ' UTF-8 to ANSI Conversion
    If GlLog = True Then SLogi "Imp01: Checking encoding..."
    If UniDun Then
        If GlLog = True Then SLogi "Imp01: UTF-16 already decoded, skipping UTF-8 check"
    ElseIf IsUTF8(TmpSt) Then
        If GlLog = True Then SLogi "Imp01: UTF-8 detected, converting..."
        TmpSt = ConvUTF8(TmpSt)
    Else
        If GlLog = True Then SLogi "Imp01: ANSI encoding, no conversion needed"
    End If
    If GlLog = True Then SLogi "Imp01: Encoding done, length=" & Len(TmpSt)

    ' Check for MT940 format (only check first 500 chars to avoid false positives)
    Dim HeadTx As String
    HeadTx = Left$(TmpSt, 500)
    If InStr(HeadTx, ":20:") > 0 Or InStr(HeadTx, ":60F:") > 0 Or InStr(HeadTx, ":60M:") > 0 Then
        If GlLog = True Then SLogi "Imp01: MT940 format detected"
        ImpMT940 TmpSt, RS121, IdBnk, ManNr, MitNr
        DidImport = True
        ChkDup = True 'MT940 speichert ohne Bestandsabgleich: danach S_BaDop
        GoTo Cleanup
    End If
    If GlLog = True Then SLogi "Imp01: CSV format detected"
    
    ' --------------------------------------------------------------------------
    ' 4. CSV Parsing
    ' --------------------------------------------------------------------------
    Dim Delim As String
    Dim RowsC As Collection

    Delim = FindD(TmpSt)
    If GlLog = True Then SLogi "Imp01: Detected delimiter=[" & Delim & "]"

    ' Fix concatenated records (Volksbank format where line breaks are missing)
    If GlLog = True Then SLogi "Imp01: Calling FixConc..."
    TmpSt = FixConc(TmpSt, Delim)
    If GlLog = True Then SLogi "Imp01: FixConc done, length=" & Len(TmpSt)

    If GlLog = True Then SLogi "Imp01: Calling ParsC..."
    Set RowsC = ParsC(TmpSt, Delim)

    If GlLog = True Then SLogi "Imp01: RowsC.Count=" & RowsC.Count

    ' Fix Deutsche Bank comma-delimited files where decimal comma splits amounts
    If Delim = "," And RowsC.Count > 1 Then
        If GlLog = True Then SLogi "Imp01: Checking for Deutsche Bank comma-delimiter fix..."
        Set RowsC = FixDBK(RowsC)
    End If

    If RowsC.Count < 1 Then
        If GlLog = True Then SLogi "Imp01: No rows parsed, exiting"
        GoTo Cleanup
    End If

    ' NEW: Header Detection
    Dim HdrRow As Long
    HdrRow = FindHeader(RowsC)
    NewBank = BankCols(RowsC, BankMap, HdrRow)
    If NewBank Then
        PrBr1.Min = 0
        PrBr1.Max = RowsC.Count
        frmStatus.Caption = "Bankumsaetze abgleichen..."
        DidImport = True
        BankImport RowsC, BankMap, HdrRow, RS121, IdBnk, ManNr, MitNr
        GoTo Cleanup
    End If
    If GlLog = True Then SLogi "Imp01: HdrRow=" & HdrRow

    ' Column Detection
    Dim IdxDa As Integer, IdxAm As Integer, IdxAm2 As Integer
    IdxAm2 = -1
    Dim ColIg As Collection
    Set ColIg = New Collection

    ' --- NEW LOGIC: ALWAYS run ScanD first, then override for special formats ---
    ' 1. Perform general analysis to find date/amount columns and all constant columns to be ignored.
    ScanD RowsC, IdxDa, IdxAm, ColIg, HdrRow
    
    ' 2. Detect special formats (e.g., StarMoney)
    Dim IsStarMoneyFmt As Boolean
    Dim IsStarMoneyFmt2 As Boolean
    Dim IsStarMoneyFmt3 As Boolean ' NEW: For StarMoney Delux format
    IsStarMoneyFmt = False
    IsStarMoneyFmt2 = False
    IsStarMoneyFmt3 = False

    On Error Resume Next
    If HdrRow > 0 And HdrRow <= RowsC.Count Then
        Dim HdrV As Variant
        HdrV = RowsC(HdrRow)
        
        ' --- Check for StarMoney Delux format (66 columns, 0 to 65) ---
        If UBound(HdrV) >= 65 Then
            Dim FirstHeader As String
            FirstHeader = LCase$(ClnTx(CStr(HdrV(0))))
            Dim BetragHeader As String
            If UBound(HdrV) >= 7 Then BetragHeader = LCase$(ClnTx(CStr(HdrV(7))))
            Dim BuchDatumHeader As String
            If UBound(HdrV) >= 26 Then BuchDatumHeader = LCase$(ClnTx(CStr(HdrV(26))))
            Dim Vwz1Header As String
            If UBound(HdrV) >= 12 Then Vwz1Header = LCase$(ClnTx(CStr(HdrV(12))))
            
            If GlLog = True Then
                SLogi "Imp01: Delux Detection: FirstHeader=" & FirstHeader & ", BetragHeader=" & BetragHeader & ", BuchDatumHeader=" & BuchDatumHeader & ", Vwz1Header=" & Vwz1Header
            End If

            If FirstHeader = "saldo" And BetragHeader = "betrag" And BuchDatumHeader = "buchdatum" And Vwz1Header = "vwz1" Then
                IsStarMoneyFmt3 = True
                If GlLog = True Then SLogi "Imp01: StarMoney Delux format detected."
            End If
        End If
        ' --- END Check for StarMoney Delux format ---

        ' --- Check for existing StarMoney CSV format ---
        If Not IsStarMoneyFmt3 Then ' Only check if not already detected as Delux
            Dim h As Long
            Dim hasVwz14 As Boolean
            hasVwz14 = False
            For h = 0 To UBound(HdrV)
                If InStr(1, CStr(HdrV(h)), "Verwendungszweckzeile 14", vbTextCompare) > 0 Then
                    hasVwz14 = True
                    Exit For
                End If
            Next h

            If hasVwz14 Then
                Dim CurrentFirstHeader As String
                CurrentFirstHeader = LCase$(ClnTx(CStr(HdrV(0))))
                
                If CurrentFirstHeader = "kontonummer" Then
                    IsStarMoneyFmt = True
                ElseIf CurrentFirstHeader = "saldo" Then
                    IsStarMoneyFmt2 = True
                End If
            End If
        End If
    End If
    On Error GoTo LiErr

    ' 3. If a special format is detected, OVERRIDE the indices and add extra ignored columns.
    If IsStarMoneyFmt3 Then ' NEW: Handle StarMoney Delux format
        If GlLog = True Then SLogi "Imp01: StarMoney Delux format detected. Overriding indices."
        IdxDa = 26 ' BuchDatum
        IdxAm = 7  ' Betrag

        ' Add redundant columns for this format
        If Not InCol(ColIg, 1) Then ColIg.Add 1, "K1"   ' SdoWaehr
        If Not InCol(ColIg, 5) Then ColIg.Add 5, "K5"   ' Storno
        If Not InCol(ColIg, 6) Then ColIg.Add 6, "K6"   ' OrigBtg
        If Not InCol(ColIg, 8) Then ColIg.Add 8, "K8"   ' BtgWaehr
        If Not InCol(ColIg, 9) Then ColIg.Add 9, "K9"   ' OCMTBetr
        If Not InCol(ColIg, 10) Then ColIg.Add 10, "K10" ' OCMTWaehr
        ' WertDatum (Column 27) is also a date column, likely redundant if BuchDatum (26) is used
        If Not InCol(ColIg, 27) Then ColIg.Add 27, "K27" ' WertDatum
        
    ElseIf IsStarMoneyFmt Then
        If GlLog = True Then SLogi "Imp01: StarMoney CSV format detected. Overriding indices."
        IdxDa = 7  ' Buchungstag
        IdxAm = 4  ' Betrag

        ' Add redundant columns that ScanD might have missed for this format
        If Not InCol(ColIg, 46) Then ColIg.Add 46, "K46" ' Wertstellungstag
        If Not InCol(ColIg, 75) Then ColIg.Add 75, "K75" ' Soll
        If Not InCol(ColIg, 76) Then ColIg.Add 76, "K76" ' Haben
        
    ElseIf IsStarMoneyFmt2 Then
        If GlLog = True Then SLogi "Imp01: StarMoney TXT format detected. Overriding indices."
        IdxDa = 26 ' BuchDatum
        IdxAm = 7  ' Betrag

        ' Add redundant columns for this format
        If Not InCol(ColIg, 27) Then ColIg.Add 27, "K27" ' WertDatum
    End If
    ' --- END OF NEW LOGIC ---

    ' FALLBACK: If detection failed and we tried a shifted header, try standard row 1
    If (IdxDa < 0 Or IdxAm < 0) And HdrRow > 1 Then
        If GlLog = True Then SLogi "Imp01: Fallback - trying HdrRow=1"
        HdrRow = 1
        Set ColIg = New Collection
        ScanD RowsC, IdxDa, IdxAm, ColIg, HdrRow
    End If

    If IdxDa < 0 Or IdxAm < 0 Then
        If GlLog = True Then SLogi "Imp01: Column detection failed! IdxDa=" & IdxDa & " IdxAm=" & IdxAm
        SMeFr "Bankimport", "Konnte Datum- oder Betragsspalte nicht automatisch erkennen.", vbNullString, vbNullString, False, 2, True, FM.hwnd
        GoTo Cleanup
    End If

    ' Detect Soll/Haben sister column (e.g. Targobank: separate debit/credit columns)
    Call FindSister(RowsC, IdxDa, IdxAm, IdxAm2, HdrRow)
    If IdxAm2 >= 0 Then
        If Not InCol(ColIg, CInt(IdxAm2)) Then ColIg.Add CInt(IdxAm2), "K" & CStr(IdxAm2)
        If GlLog = True Then SLogi "Imp01: Soll/Haben sister column at " & IdxAm2
    End If
    
    ' DEBUG LOG via SLogi
    If GlLog = True Then
        SLogi "DEBUG: Start Import Analysis"
        SLogi "DEBUG: HdrRow: " & HdrRow
        SLogi "DEBUG: IdxDa: " & IdxDa & " (Expected Buchungstag)"
        SLogi "DEBUG: IdxAm: " & IdxAm & " (Expected Betrag)"
    End If
    
    ' --------------------------------------------------------------------------
    ' 5. CSV Import Loop
    ' --------------------------------------------------------------------------
    PrBr1.Min = 0
    PrBr1.Max = RowsC.Count
    frmStatus.Caption = "Importiere Daten..."
    
    Dim i As Long
    Dim RowV As Variant
    Dim sDat As String, sAmt As String, sTxt As String, sAmtAlt As String
    Dim dAmt As Double, dtDat As Date
    Dim j As Integer
    Dim sTmp As String
    Dim sSch As String
    Dim NeuBu As Boolean
    Dim RowUB As Long
    Dim PlanC As Collection
    Dim MinDay As Date, MaxDay As Date
    
    Dim StartLoop As Long
    StartLoop = HdrRow + 1

    If GlLog = True Then SLogi "Import: Starting loop from row " & StartLoop & " to " & RowsC.Count
    If GlLog = True Then SLogi "Import: IdxDa=" & IdxDa & " IdxAm=" & IdxAm

    DidImport = True 'ab hier werden Umsaetze gespeichert: S_List auch nach Abbruch
    Set PlanC = New Collection
    For i = StartLoop To RowsC.Count
        RowV = RowsC(i)
        NeuBu = False

        ' Safely check array bounds
        On Error Resume Next
        RowUB = -1
        RowUB = UBound(RowV)
        On Error GoTo LiErr
        
        ' DEBUG Log for first 10 rows
        If GlLog = True And i < StartLoop + 10 Then
            SLogi "DEBUG: Row " & i & ": Cols=" & RowUB
        End If

        If RowUB >= IdxDa And RowUB >= IdxAm And IdxDa >= 0 And IdxAm >= 0 Then
            sDat = ClnTx(CStr(RowV(IdxDa)))
            sAmt = ClnTx(CStr(RowV(IdxAm)))
            sAmt = ClnAm(sAmt)

            ' Soll/Haben sister fallback: if primary amount is empty/zero, take sister value
            If IdxAm2 >= 0 Then
                If RowUB >= IdxAm2 Then
                    If Not IsNumeric(sAmt) Or Val(sAmt) = 0 Then
                        sAmtAlt = ClnAm(ClnTx(CStr(RowV(IdxAm2))))
                        If IsNumeric(sAmtAlt) Then
                            If Val(sAmtAlt) <> 0 Then sAmt = sAmtAlt
                        End If
                    End If
                End If
            End If

            dtDat = ParseDate(sDat)
            
            ' DEBUG Log details
            If GlLog = True And i < StartLoop + 10 Then
                SLogi "DEBUG:   RawDate='" & sDat & "' -> Parsed=" & dtDat
                SLogi "DEBUG:   RawAmt='" & CStr(RowV(IdxAm)) & "' -> ClnAmt='" & sAmt & "' -> IsNum=" & IsNumeric(sAmt)
            End If

            If IsNumeric(sAmt) And dtDat > 0 Then
                dAmt = Val(sAmt)

                If Year(dtDat) > 1900 Then
                     ' Text Construction
                    sTxt = ""
                    For j = 0 To RowUB
                        If j <> IdxDa And j <> IdxAm And Not InCol(ColIg, CInt(j)) Then
                            sTmp = ClnTx(CStr(RowV(j)))
                            If Len(sTmp) > 0 Then
                                If Len(sTxt) > 0 Then sTxt = sTxt & " "
                                sTxt = sTxt & sTmp
                            End If
                        End If
                    Next j

                    NeuBu = True
                    If GlLog = True Then SLogi "Import Row " & i & ": Date=" & sDat & " Amt=" & dAmt
                End If
            Else
                If GlLog = True Then SLogi "Import Row " & i & " Skipped: sAmt=[" & sAmt & "] dtDat=" & dtDat
            End If
        Else
            If GlLog = True Then SLogi "Import Row " & i & " Skipped: RowUB=" & RowUB
        End If

        If NeuBu Then
            'Gespeichert wird erst nach dem Bestandsabgleich in BankStore
            PlanC.Add Array(dtDat, dAmt, sTxt, BankParts(RowV, IdxDa, IdxAm), ParseDate(sDat, True))
            BankSpan dtDat, MinDay, MaxDay
             If GlLog = True And i < StartLoop + 10 Then SLogi "DEBUG:   -> PLANNED"
        Else
             If GlLog = True And i < StartLoop + 10 Then SLogi "DEBUG:   -> SKIPPED"
        End If

        PrBr1.Value = i
        DoEvents
    Next i
    If PlanC.Count > 0 Then
        frmStatus.Caption = "Bankumsaetze abgleichen..."
        BankStore PlanC, RS121, IdBnk, ManNr, MitNr, MinDay, MaxDay
    End If

Cleanup:
    ' --------------------------------------------------------------------------
    ' 6. Cleanup & Post-Processing
    ' --------------------------------------------------------------------------
    On Error Resume Next
    Unload frmStatus
    Set frmStatus = Nothing
    DoEvents
    
    If Not RS121 Is Nothing Then
        If RS121.State = adStateOpen Then RS121.Close
        Set RS121 = Nothing
    End If
    
    Set CoDia = Nothing
    Set clFil = Nothing
    
    DoEvents
    'CSV-Importe gleichen vor dem Speichern mit dem Bestand ab (BankStore, BankImport); S_BaDop wuerde dort
    'auch echte gleiche Buchungen zum Loeschen anbieten
    If DidImport And ChkDup Then S_BaDop
    
    DoEvents
    If DidImport Then S_List
    ImportBusy = False
    
    Exit Sub

LiErr:
    If GlLog = True Then SLogi "Imp01 ERROR: " & Err.Number & " - " & Err.Description
    If Err.Number = vbObjectError + 2210 Or Err.Number = vbObjectError + 2211 Then
        SMeFr "Bankimport", "Bankimport abgebrochen", Err.Description, vbNullString, False, 2, True, FM.hwnd
    Else
        SMeFr "Bankimport", "Bankimport abgebrochen", "Fehler " & CStr(Err.Number) & ". Bitte den Import pruefen.", "Bereits gespeicherte Umsaetze bleiben erhalten.", False, 2, True, FM.hwnd
    End If
    Resume Cleanup
End Sub

' ------------------------------------------------------------------------------
' MT940 Parser
' ------------------------------------------------------------------------------
Private Sub ImpMT940(ByVal s As String, ByRef RS As ADODB.Recordset, ByVal IdBnk As Long, ByVal ManNr As Long, ByVal MitNr As Long)
    Dim Lines() As String
    Dim i As Long, Ln As String
    Dim CurDat As Date, CurAmt As Double
    Dim CurTxt As String, CurTag As String
    Dim HasTrans As Boolean
    
    ' Normalize line endings
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    Lines = Split(s, vbLf)
    
    HasTrans = False
    
    For i = 0 To UBound(Lines)
        Ln = Trim$(Lines(i))
        If Len(Ln) > 0 Then
            If Left$(Ln, 1) = ":" Then
                ' Process previous if new 61 starts
                If InStr(Ln, ":61:") = 1 Then
                    If HasTrans Then
                        AddRec RS, IdBnk, ManNr, MitNr, CurDat, CurAmt, CurTxt
                    End If
                    
                    ' Parse New 61
                    ' :61:YYMMDD(MMDD)DC(Mark)Amount...
                    Dim Content As String, OffSet As Long
                    Dim sD As String, sA As String, sM As String
                    
                    Content = Mid$(Ln, 5)
                    sD = Left$(Content, 6)
                    CurDat = ParseDate(sD)
                    
                    OffSet = 7
                    ' Skip Entry Date if present (4 digits)
                    If IsNumeric(Mid$(Content, OffSet, 4)) Then OffSet = OffSet + 4
                    
                    ' DC Mark (starts at OffSet)
                    sM = Mid$(Content, OffSet, 2) ' CR, D, C, DR, RC, RD
                    Dim Sig As Double
                    Sig = 1
                    ' Standard: C/CR/RC=Credit(+), D/DR/RD=Debit(-)
                    ' Note: RC/RD are Reversals. Logic may vary. Assume C=Pos, D=Neg
                    If InStr(sM, "D") > 0 Then Sig = -1
                    
                    ' Move past Mark to Amount
                    If IsNumeric(Right$(sM, 1)) Or Right$(sM, 1) = "," Then
                        ' Mark was 1 char
                        OffSet = OffSet + 1
                    Else
                        ' Mark was 2 chars
                        OffSet = OffSet + 2
                    End If
                    
                    ' Extract Amount (until N or non-numeric/comma)
                    ' MT940 amount uses comma.
                    Dim j As Long
                    Dim RemS As String
                    RemS = Mid$(Content, OffSet)
                    For j = 1 To Len(RemS)
                        Dim c As String * 1
                        c = Mid$(RemS, j, 1)
                        If Not (IsNumeric(c) Or c = "," Or c = ".") Then Exit For
                    Next j
                    sA = Left$(RemS, j - 1)
                    CurAmt = Val(ClnAm(sA)) * Sig

                    CurTxt = ""
                    CurTag = "61"
                    HasTrans = True
                    
                ElseIf InStr(Ln, ":86:") = 1 Then
                    CurTag = "86"
                    CurTxt = CurTxt & Mid$(Ln, 5)
                Else
                    CurTag = Mid$(Ln, 2, InStr(2, Ln, ":") - 2)
                End If
            Else
                ' Continuation
                If CurTag = "86" Then CurTxt = CurTxt & " " & Ln
            End If
        End If
    Next i
    
    ' Add last
    If HasTrans Then
         AddRec RS, IdBnk, ManNr, MitNr, CurDat, CurAmt, CurTxt
    End If
End Sub

Private Sub AddRec(RS As ADODB.Recordset, IdBnk As Long, ManNr As Long, MitNr As Long, dtDat As Date, dAmt As Double, sTxt As String, Optional ByVal StableKey As String)
    Dim BaGui As String
    Dim sSch As String
    
    BaGui = CreateID("B")
    
    ' Clean Text (remove MT940 tags like ?00, ?10 if present)
    sTxt = Replace(sTxt, "?", " ")
    ' Simple cleanup, can be enhanced
    
    RS.AddNew
    RS.Fields("Betrag").Value = Round(dAmt, 2)
    RS.Fields("Datum").Value = dtDat
    RS.Fields("IDB").Value = IdBnk
    RS.Fields("IDM").Value = ManNr
    RS.Fields("IDT").Value = MitNr
    RS.Fields("GuiID").Value = BaGui
    
    If Len(sTxt) > 255 Then
        RS.Fields("IDKurz").Value = Left$(sTxt, 255)
    Else
        RS.Fields("IDKurz").Value = IIf(Len(sTxt) > 0, sTxt, "kein Buchungstext")
    End If
    
    sSch = CStr(dtDat) & " " & Left$(sTxt, 200) & " " & CStr(dAmt)
    If Len(sSch) > 255 Then sSch = Left$(sSch, 255)
    If Len(StableKey) > 0 Then sSch = StableKey
    RS.Fields("Suche").Value = sSch
    
    RS.Update
End Sub

' ------------------------------------------------------------------------------
' Helpers (Universal Parsing)
' ------------------------------------------------------------------------------

Private Function ParseDate(ByVal s As String, Optional ByVal OldRule As Boolean) As Date
    On Error GoTo ErrHnd
    ParseDate = 0
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    Dim lD As Long, lM As Long, lY As Long
    Dim Parts() As String
    Dim PartCount As Long
    Dim Fict As Date

    ' Fiktive Bankvaluta wie 29.02.25, 31-04-2025 oder 2025-02-30 (Monatsende): IsDate liest sie in anderer
    ' Reihenfolge (29.02.25 -> 25.02.2029), daher vorab der letzte Tag des Monats (BankFict). OldRule: bisherige
    ' Lesart fuer stabile Kennungen und frueher gespeicherte Umsaetze
    If Not OldRule Then
        Fict = BankFict(s)
        If Fict > 0 Then
            ParseDate = Fict
            Exit Function
        End If
    End If

    ' Try native VB date parsing first
    If IsDate(s) Then
        ParseDate = CDate(s)
        If Err.Number = 0 Then Exit Function
        Err.Clear
    End If

    ' Try DD.MM.YYYY (German format)
    If InStr(s, ".") > 0 Then
        Parts = Split(s, ".")
        PartCount = -1
        PartCount = UBound(Parts)
        If Err.Number <> 0 Then Err.Clear: GoTo TryDash
        If PartCount = 2 Then
            If IsNumeric(Parts(0)) And IsNumeric(Parts(1)) And IsNumeric(Parts(2)) Then
                lD = CLng(Parts(0))
                lM = CLng(Parts(1))
                lY = CLng(Parts(2))
                If lY < 100 Then lY = 2000 + lY
                If lD >= 1 And lD <= 31 And lM >= 1 And lM <= 12 And lY > 1900 And lY < 2100 Then
                    ParseDate = BankDate(lY, lM, lD, OldRule)
                    If Err.Number = 0 Then Exit Function
                    Err.Clear
                End If
            End If
        End If
    End If

TryDash:
    ' Try DD-MM-YYYY
    If InStr(s, "-") > 0 Then
        Parts = Split(s, "-")
        PartCount = -1
        PartCount = UBound(Parts)
        If Err.Number <> 0 Then Err.Clear: GoTo TrySlash
        If PartCount = 2 Then
            If IsNumeric(Parts(0)) And IsNumeric(Parts(1)) And IsNumeric(Parts(2)) Then
                lD = CLng(Parts(0))
                lM = CLng(Parts(1))
                lY = CLng(Parts(2))
                If lY < 100 Then lY = 2000 + lY
                If lD >= 1 And lD <= 31 And lM >= 1 And lM <= 12 And lY > 1900 And lY < 2100 Then
                    ParseDate = BankDate(lY, lM, lD, OldRule)
                    If Err.Number = 0 Then Exit Function
                    Err.Clear
                End If
            End If
        End If
    End If

TrySlash:
    ' Try DD/MM/YYYY
    If InStr(s, "/") > 0 Then
        Parts = Split(s, "/")
        PartCount = -1
        PartCount = UBound(Parts)
        If Err.Number <> 0 Then Err.Clear: GoTo TryNumeric
        If PartCount = 2 Then
            If IsNumeric(Parts(0)) And IsNumeric(Parts(1)) And IsNumeric(Parts(2)) Then
                lD = CLng(Parts(0))
                lM = CLng(Parts(1))
                lY = CLng(Parts(2))
                If lY < 100 Then lY = 2000 + lY
                If lD >= 1 And lD <= 31 And lM >= 1 And lM <= 12 And lY > 1900 And lY < 2100 Then
                    ParseDate = BankDate(lY, lM, lD, OldRule)
                    If Err.Number = 0 Then Exit Function
                    Err.Clear
                End If
            End If
        End If
    End If

TryNumeric:
    ' Try YYYYMMDD
    If Len(s) = 8 Then
        If Not IsNumeric(s) Then GoTo SkipYYYYMMDD
        ' Strict check: must not contain separators
        If InStr(s, ".") > 0 Or InStr(s, ",") > 0 Or InStr(s, "-") > 0 Or InStr(s, "+") > 0 Then GoTo SkipYYYYMMDD
        
        On Error Resume Next
        lY = CLng(Left$(s, 4))
        lM = CLng(Mid$(s, 5, 2))
        lD = CLng(Right$(s, 2))
        On Error GoTo ErrHnd
        
        If lD >= 1 And lD <= 31 And lM >= 1 And lM <= 12 And lY > 1900 And lY < 2100 Then
            ParseDate = BankDate(lY, lM, lD, OldRule)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        End If
    End If
SkipYYYYMMDD:

    ' Try YYMMDD
    If Len(s) = 6 Then
        If Not IsNumeric(s) Then GoTo SkipYYMMDD
        ' Strict check: must not contain separators
        If InStr(s, ".") > 0 Or InStr(s, ",") > 0 Or InStr(s, "-") > 0 Or InStr(s, "+") > 0 Then GoTo SkipYYMMDD

        On Error Resume Next
        lY = 2000 + CLng(Left$(s, 2))
        lM = CLng(Mid$(s, 3, 2))
        lD = CLng(Right$(s, 2))
        On Error GoTo ErrHnd
        
        If lD >= 1 And lD <= 31 And lM >= 1 And lM <= 12 Then
            ParseDate = BankDate(lY, lM, lD, OldRule)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        End If
    End If
SkipYYMMDD:
    
    Exit Function

ErrHnd:
    ParseDate = 0
    Err.Clear
End Function

Private Function BankDate(ByVal lY As Long, ByVal lM As Long, ByVal lD As Long, ByVal OldRule As Boolean) As Date
    'Datum aus Jahr, Monat (1-12) und Tag (1-31); ein Tag ueber dem Monatsende (fiktive Bankvaluta
    'wie 30.02.) ergibt den letzten Tag des Monats statt eines Tages im Folgemonat (OldRule: Folgemonat)
    Dim LastD As Long
    LastD = Day(DateSerial(CInt(lY), CInt(lM) + 1, 0))
    If lD > LastD And Not OldRule Then lD = LastD
    BankDate = DateSerial(CInt(lY), CInt(lM), CInt(lD))
End Function

Private Function BankFict(ByVal s As String) As Date
    'Fiktive Bankvaluta ueber dem Monatsende (29.02.25, 31-04-2025, 31/04/2025, 2025-02-30) als letzter Tag des
    'Monats. Nur Ziffern und ein Trennzeichen (. - /); Tag und Monat ein- oder zweistellig, Jahr zwei- oder
    'vierstellig, vierstelliges Jahr vorn = Jahr-Monat-Tag. Sonst 0: bisherige Erkennung in ParseDate
    Dim Parts() As String
    Dim Sep As String
    Dim lD As Long, lM As Long, lY As Long, n As Long, Code As Long
    If InStr(s, ".") > 0 Then
        Sep = "."
    ElseIf InStr(s, "-") > 0 Then
        Sep = "-"
    ElseIf InStr(s, "/") > 0 Then
        Sep = "/"
    Else
        Exit Function
    End If
    For n = 1 To Len(s)
        Code = AscW(Mid$(s, n, 1))
        If (Code < 48 Or Code > 57) And Mid$(s, n, 1) <> Sep Then Exit Function
    Next n
    Parts = Split(s, Sep)
    If UBound(Parts) <> 2 Then Exit Function
    If Len(Parts(0)) = 0 Or Len(Parts(1)) = 0 Or Len(Parts(1)) > 2 Then Exit Function
    If Len(Parts(0)) = 4 And Len(Parts(2)) >= 1 And Len(Parts(2)) <= 2 Then
        lY = CLng(Parts(0))
        lM = CLng(Parts(1))
        lD = CLng(Parts(2))
    ElseIf Len(Parts(0)) <= 2 And (Len(Parts(2)) = 2 Or Len(Parts(2)) = 4) Then
        lD = CLng(Parts(0))
        lM = CLng(Parts(1))
        lY = CLng(Parts(2))
        If lY < 100 Then lY = 2000 + lY
    Else
        Exit Function
    End If
    If lD < 29 Or lD > 31 Or lM < 1 Or lM > 12 Or lY <= 1900 Or lY >= 2100 Then Exit Function
    If lD > Day(DateSerial(CInt(lY), CInt(lM) + 1, 0)) Then BankFict = BankDate(lY, lM, lD, False)
End Function

Private Function FindD(ByVal s As String) As String
    Dim CntS As Long, CntC As Long, CntT As Long
    Dim SubS As String
    SubS = Left$(s, 2000)
    CntS = Len(SubS) - Len(Replace(SubS, ";", ""))
    CntC = Len(SubS) - Len(Replace(SubS, ",", ""))
    CntT = Len(SubS) - Len(Replace(SubS, vbTab, ""))
    
    If CntT > CntS And CntT > CntC Then
        FindD = vbTab
    ElseIf CntS > CntC Then
        FindD = ";"
    Else
        FindD = ","
    End If
End Function

Private Function ParsC(ByVal s As String, ByVal D As String) As Collection
    Dim Res As Collection
    Set Res = New Collection
    Dim Fld As String
    Dim InQ As Boolean
    Dim c As String
    Dim i As Long, L As Long
    Dim Fields As Collection
    Set Fields = New Collection
    Dim Q As String
    Q = Chr$(34)

    L = Len(s)
    InQ = False

    If GlLog = True Then SLogi "ParsC: Input length=" & L & " Delim=[" & D & "]"

    For i = 1 To L
        c = Mid$(s, i, 1)
        If InQ Then
            If c = Q Then
                If i < L Then
                    If Mid$(s, i + 1, 1) = Q Then
                        Fld = Fld & Q
                        i = i + 1
                    Else
                        InQ = False
                    End If
                Else
                    InQ = False
                End If
            Else
                Fld = Fld & c
            End If
        Else
            If c = Q Then
                InQ = True
            ElseIf c = D Then
                Fields.Add Fld
                Fld = ""
            ElseIf c = vbCr Or c = vbLf Then
                If c = vbCr And i < L Then
                    If Mid$(s, i + 1, 1) = vbLf Then i = i + 1
                End If
                Fields.Add Fld
                Fld = ""
                If Not BankEmpty(Fields) Then
                    Res.Add ColToArr(Fields)
                End If
                Set Fields = New Collection
            Else
                Fld = Fld & c
            End If
        End If
    Next i
    If Fields.Count > 0 Or Len(Fld) > 0 Then
        Fields.Add Fld
        If Not BankEmpty(Fields) Then Res.Add ColToArr(Fields)
    End If

    If GlLog = True Then SLogi "ParsC: Parsed " & Res.Count & " rows"

    Set ParsC = Res
End Function

Private Function FixDBK(ByVal Rows As Collection) As Collection
    ' -------------------------------------------------------------------------
    ' Fix Deutsche Bank comma-delimited CSV where decimal comma in amounts
    ' causes field splitting. Example: amount "99,64" becomes fields "99","64"
    ' expanding 18 columns to 19-21. We detect the DB header pattern and
    ' rejoin split amount fields at known positions (Betrag=11, Soll=15, Haben=16).
    ' -------------------------------------------------------------------------
    Dim Result As Collection
    Set Result = New Collection

    Dim HdIdx As Long
    Dim HdrV As Variant
    Dim ExpCo As Long  ' expected column count from header
    ExpCo = 0
    HdIdx = 0

    ' Step 1: Find the Deutsche Bank header row
    ' Header signature: first col contains "buchungstag", col 11 = "betrag", col 17 = "waehrung"/"w" (truncated umlaut)
    Dim i As Long
    Dim Limit As Long
    Limit = Rows.Count
    If Limit > 30 Then Limit = 30

    For i = 1 To Limit
        HdrV = Rows(i)
        If IsArray(HdrV) Then
            If UBound(HdrV) >= 17 Then
                Dim Col00 As String
                Dim Col11 As String
                Col00 = LCase$(Trim$(CStr(HdrV(0))))
                Col11 = LCase$(Trim$(CStr(HdrV(11))))
                If InStr(Col00, "buchungstag") > 0 And InStr(Col11, "betrag") > 0 Then
                    HdIdx = i
                    ExpCo = UBound(HdrV) + 1
                    If GlLog = True Then SLogi "FixDBK: DB header at row " & i & " ExpCo=" & ExpCo
                    Exit For
                End If
            End If
        End If
    Next i

    ' If no Deutsche Bank header found, return unchanged
    If HdIdx = 0 Then
        If GlLog = True Then SLogi "FixDBK: No DB header found, skipping"
        Set FixDBK = Rows
        Exit Function
    End If

    ' Step 2: Process each row
    Dim RowV As Variant
    Dim RowUB As Long
    Dim Extra As Long
    Dim Fixed As Long
    Fixed = 0

    For i = 1 To Rows.Count
        RowV = Rows(i)
        If Not IsArray(RowV) Then
            Result.Add RowV
            GoTo NxtRw
        End If

        RowUB = UBound(RowV) + 1  ' actual column count

        ' Only fix data rows after header that have too many columns
        If i > HdIdx And RowUB > ExpCo Then
            Extra = RowUB - ExpCo  ' how many extra fields from split decimals

            ' Rebuild row: rejoin decimal fragments at amount positions
            ' Deutsche Bank has 3 amount fields that can split:
            '   Betrag (col 11), Soll (col 15), Haben (col 16)
            ' Each split adds 1 extra column. We process right-to-left
            ' to keep index positions stable.
            Dim NewRw() As String
            ReDim NewRw(ExpCo - 1)

            Dim SrcIx As Long  ' source index in original (expanded) row
            Dim DstIx As Long  ' destination index in fixed row
            SrcIx = 0
            DstIx = 0

            Do While DstIx < ExpCo And SrcIx <= UBound(RowV)
                NewRw(DstIx) = CStr(RowV(SrcIx))

                ' Check if this is an amount column and the row has excess fields
                If Extra > 0 Then
                    If DstIx = 11 Or DstIx = 15 Or DstIx = 16 Then
                        ' Only rejoin if the integer part is non-empty (a real amount).
                        ' Empty amount fields (e.g. Soll on credit rows) must stay empty;
                        ' otherwise the next field (Haben integer) gets wrongly consumed.
                        If Len(NewRw(DstIx)) > 0 Then
                            If SrcIx + 1 <= UBound(RowV) Then
                                Dim NxtFl As String
                                NxtFl = Trim$(CStr(RowV(SrcIx + 1)))
                                If Len(NxtFl) >= 1 And Len(NxtFl) <= 2 Then
                                    If IsNum(NxtFl) Then
                                        ' Rejoin: "99" + "," + "64" -> "99,64"
                                        NewRw(DstIx) = NewRw(DstIx) & "," & NxtFl
                                        SrcIx = SrcIx + 1
                                        Extra = Extra - 1
                                        If GlLog = True And Fixed < 10 Then
                                            SLogi "FixDBK: Row " & i & " Col " & DstIx & " rejoined -> [" & NewRw(DstIx) & "]"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                SrcIx = SrcIx + 1
                DstIx = DstIx + 1
            Loop

            Result.Add NewRw
            Fixed = Fixed + 1
        Else
            Result.Add RowV
        End If
NxtRw:
    Next i

    If GlLog = True Then SLogi "FixDBK: Fixed " & Fixed & " rows, total=" & Result.Count
    Set FixDBK = Result
End Function

Private Function IsNum(ByVal s As String) As Boolean
    ' Check if string contains only digit characters (0-9)
    Dim k As Long
    IsNum = False
    If Len(s) = 0 Then Exit Function
    For k = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, k, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next k
    IsNum = True
End Function

Private Function ColToArr(ByVal c As Collection) As String()
    Dim a() As String
    Dim i As Long
    If c.Count > 0 Then
        ReDim a(c.Count - 1)
        For i = 1 To c.Count
            a(i - 1) = c(i)
        Next i
    Else
        ReDim a(0)
    End If
    ColToArr = a
End Function

Private Function FindHeader(ByVal Rows As Collection) As Long
    Dim i As Long, j As Long
    Dim RowV As Variant
    Dim s As String
    Dim Score As Long
    Dim BestScore As Long, BestRow As Long

    If GlLog = True Then SLogi "FindHeader: Rows.Count=" & Rows.Count

    BestRow = 1
    BestScore = -1

    Dim Limit As Long
    Limit = 50
    If Limit > Rows.Count Then Limit = Rows.Count

    For i = 1 To Limit
        RowV = Rows(i)
        Score = 0

        ' Check if row has data
        If IsArray(RowV) Then
            ' Bonus for number of columns (Header usually has many columns)
            If UBound(RowV) > 2 Then Score = Score + 2

            For j = LBound(RowV) To UBound(RowV)
                s = LCase(ClnTx(CStr(RowV(j))))
                If Len(s) > 0 Then
                    ' Date Keywords
                    If InStr(s, "datum") > 0 Or InStr(s, "date") > 0 Or InStr(s, "valuta") > 0 Or InStr(s, "buchung") > 0 Or InStr(s, "zeitraum") > 0 Then
                        Score = Score + 10
                    End If
                    ' Amount Keywords
                    If InStr(s, "betrag") > 0 Or InStr(s, "amount") > 0 Or InStr(s, "saldo") > 0 Or InStr(s, "wert") > 0 Or InStr(s, "soll") > 0 Or InStr(s, "haben") > 0 Or InStr(s, "umsatz") > 0 Then
                        Score = Score + 10
                    End If
                    ' Structure Keywords
                    If InStr(s, "iban") > 0 Or InStr(s, "bic") > 0 Or InStr(s, "text") > 0 Or InStr(s, "verwendung") > 0 Or InStr(s, "empfnger") > 0 Or InStr(s, "auftraggeber") > 0 Or InStr(s, "whrung") > 0 Then
                        Score = Score + 5
                    End If
                End If
            Next j
        End If

        If Score > BestScore Then
            BestScore = Score
            BestRow = i
            If GlLog = True Then SLogi "FindHeader: Row " & i & " Score=" & Score & " (new best)"
        End If
    Next i

    If GlLog = True Then SLogi "FindHeader: Result BestRow=" & BestRow & " BestScore=" & BestScore

    FindHeader = BestRow
End Function

Private Sub ScanD(ByVal Rows As Collection, ByRef IdxDa As Integer, ByRef IdxAm As Integer, ByRef IgCol As Collection, Optional ByVal StartRow As Long = 1)
    IdxDa = -1
    IdxAm = -1
    If Rows.Count = 0 Then
        If GlLog = True Then SLogi "ScanD: Rows.Count=0, exiting"
        Exit Sub
    End If
    If StartRow < 1 Then StartRow = 1
    If StartRow > Rows.Count Then
        If GlLog = True Then SLogi "ScanD: StartRow > Rows.Count, exiting"
        Exit Sub
    End If

    Dim R As Long, c As Long, MaxC As Long
    Dim RowV As Variant
    Dim TmpV As String
    MaxC = UBound(Rows(StartRow))

    If GlLog = True Then SLogi "ScanD: StartRow=" & StartRow & " MaxC=" & MaxC & " Rows.Count=" & Rows.Count

    Dim ScDat() As Long, ScNum() As Long
    Dim ScDec() As Long ' Score for Decimal Separators
    Dim IsDatCol As Boolean, DatPrio As Integer ' For date column priority detection
    ReDim ScDat(MaxC), ScNum(MaxC), ScDec(MaxC)

    Dim Limit As Long
    Limit = StartRow + 50
    If Limit > Rows.Count Then Limit = Rows.Count

    If GlLog = True Then SLogi "ScanD: Scanning rows " & (StartRow + 1) & " to " & Limit

    ' Pass 1: Scoring
    For R = StartRow + 1 To Limit
        RowV = Rows(R)
        If UBound(RowV) >= MaxC Then
            For c = 0 To MaxC
                Dim v As String
                v = ClnTx(RowV(c))
                ' Check Date using ParseDate logic equivalent
                If ParseDate(v) > 0 Then ScDat(c) = ScDat(c) + 1

                If IsNumeric(ClnAm(v)) And Len(v) > 0 Then
                    ScNum(c) = ScNum(c) + 1
                    ' Check for Decimal Separator (Comma or Dot) indicating currency
                    If InStr(v, ",") > 0 Or InStr(v, ".") > 0 Then
                        ScDec(c) = ScDec(c) + 1
                    End If
                End If
            Next c
        End If
    Next R

    ' Log column scores
    If GlLog = True Then
        For c = 0 To MaxC
            If ScDat(c) > 0 Or ScNum(c) > 0 Then
                SLogi "ScanD: Col " & c & " ScDat=" & ScDat(c) & " ScNum=" & ScNum(c) & " ScDec=" & ScDec(c)
            End If
        Next c
    End If

    ' Best Date
    Dim BestS As Long
    For c = 0 To MaxC
        If ScDat(c) > BestS Then
            BestS = ScDat(c)
            IdxDa = c
        End If
    Next c

    If GlLog = True Then SLogi "ScanD: Best Date column=" & IdxDa & " Score=" & BestS

    ' Best Amount (vs Balance & Variance & Decimals & Headers)
    Dim NegC() As Long
    Dim VarC() As Boolean
    ReDim NegC(MaxC), VarC(MaxC)
    Dim NumCols As Collection
    Set NumCols = New Collection

    For c = 0 To MaxC
        ' Adjusted Limit check for percentage calculation
        If ScNum(c) > ((Limit - StartRow) / 2) And c <> IdxDa Then
            NumCols.Add c
            If GlLog = True Then SLogi "ScanD: NumCol candidate: " & c
            
            Dim BaseV As String
            BaseV = "|INIT|"
            
            For R = StartRow + 1 To Limit
                If UBound(Rows(R)) >= c Then
                    Dim ValS As String
                    ValS = ClnAm(Rows(R)(c))
                    
                    ' Check Negatives
                    If IsNumeric(ValS) Then
                        If Val(ValS) < 0 Then NegC(c) = NegC(c) + 1
                    End If
                    
                    ' Check Variance
                    If BaseV = "|INIT|" Then
                        BaseV = ValS
                    Else
                        If ValS <> BaseV Then VarC(c) = True
                    End If
                End If
            Next R
        End If
    Next c
    
    ' Header Priority Search
    Dim HdrIdx_Betrag As Integer
    Dim HdrIdx_Umsatz As Integer
    Dim HdrIdx_Saldo As Integer
    Dim HdrIdx_Datum As Integer
    HdrIdx_Betrag = -1
    HdrIdx_Umsatz = -1
    HdrIdx_Saldo = -1
    HdrIdx_Datum = -1

    If StartRow <= Rows.Count Then
        Dim HdrRowV As Variant
        HdrRowV = Rows(StartRow)
        For c = 0 To UBound(HdrRowV)
            Dim HdrTxt As String
            Dim HdrRaw As String
            HdrRaw = CStr(HdrRowV(c))
            HdrTxt = LCase(ClnTx(HdrRaw))
            HdrTxt = Replace(HdrTxt, Chr(34), "") ' Remove quotes
            HdrTxt = Replace(HdrTxt, Chr(160), " ") ' Non-breaking space to regular space
            HdrTxt = Trim$(HdrTxt)

            If GlLog = True And c < 15 Then SLogi "ScanD: Header[" & c & "] raw=[" & HdrRaw & "] clean=[" & HdrTxt & "]"

            ' Amount column detection - prioritize exact match, then substring
            If HdrTxt = "betrag" Or HdrTxt = "amount" Then
                HdrIdx_Betrag = c
                If GlLog = True Then SLogi "ScanD: Exact match 'Betrag' at column " & c
            ElseIf InStr(HdrTxt, "betrag") > 0 And InStr(HdrTxt, "ursprung") = 0 And InStr(HdrTxt, "aquivalenz") = 0 And HdrIdx_Betrag = -1 Then
                HdrIdx_Betrag = c
                If GlLog = True Then SLogi "ScanD: Substring match 'betrag' at column " & c
            End If

            If HdrTxt = "umsatz" And InStr(HdrTxt, "gekennzeichnet") = 0 And HdrIdx_Umsatz = -1 Then HdrIdx_Umsatz = c
            If InStr(HdrTxt, "saldo") > 0 And HdrIdx_Saldo = -1 Then HdrIdx_Saldo = c

            ' Date Header Detection with Priority System
            ' HIGH PRIORITY: datum, buchung, date (actual date columns)
            ' LOW PRIORITY: valuta (only if combined with datum, to avoid currency columns)
            IsDatCol = False
            DatPrio = 0

            ' Strict exclusion of text/description columns AND saldo columns
            If InStr(HdrTxt, "text") = 0 And InStr(HdrTxt, "zweck") = 0 And InStr(HdrTxt, "verwendung") = 0 And InStr(HdrTxt, "saldo") = 0 Then
                ' HIGH PRIORITY: Exact or substring match for datum/date
                ' For buchung: only match if not combined with id or anzahl (excludes buchungs-id, anzahl splittbuchungen)
                If InStr(HdrTxt, "datum") > 0 Or HdrTxt = "date" Then
                    IsDatCol = True
                    DatPrio = 100
                    If GlLog = True Then SLogi "ScanD: High-priority date column at " & c & " [" & HdrTxt & "] Prio=" & DatPrio
                ElseIf InStr(HdrTxt, "buchung") > 0 And InStr(HdrTxt, "id") = 0 And InStr(HdrTxt, "anzahl") = 0 Then
                    IsDatCol = True
                    DatPrio = 100
                    If GlLog = True Then SLogi "ScanD: High-priority date column at " & c & " [" & HdrTxt & "] Prio=" & DatPrio
                ' LOW PRIORITY: valuta only if combined with datum (e.g. "valuta datum")
                ElseIf InStr(HdrTxt, "valuta") > 0 Then
                    ' Accept valuta only if it also contains datum (e.g., "valuta datum", "valutadatum")
                    If InStr(HdrTxt, "datum") > 0 Then
                        IsDatCol = True
                        DatPrio = 50
                        If GlLog = True Then SLogi "ScanD: Low-priority valuta+datum column at " & c & " [" & HdrTxt & "] Prio=" & DatPrio
                    Else
                        ' Plain "valuta" without "datum" is likely a currency column - skip it
                        If GlLog = True Then SLogi "ScanD: Skipping plain valuta column (likely currency) at " & c & " [" & HdrTxt & "]"
                    End If
                End If

                ' Update HdrIdx_Datum if this is a date column and has higher or equal priority
                If IsDatCol Then
                    If HdrIdx_Datum = -1 Then
                        ' No date column set yet - use this one
                        HdrIdx_Datum = c
                        If GlLog = True Then SLogi "ScanD: Date column set to " & c & " [" & HdrTxt & "]"
                    ElseIf DatPrio > 50 Then
                        ' Higher priority column found - override previous
                        If GlLog = True Then SLogi "ScanD: Overriding date column from " & HdrIdx_Datum & " to " & c & " (higher priority)"
                        HdrIdx_Datum = c
                    End If
                End If
            End If
        Next c
    End If
    If GlLog = True Then SLogi "ScanD: Header Search -> Betrag=" & HdrIdx_Betrag & " Umsatz=" & HdrIdx_Umsatz & " Saldo=" & HdrIdx_Saldo & " Datum=" & HdrIdx_Datum

    If NumCols.Count = 1 Then
        IdxAm = NumCols(1)
    ElseIf NumCols.Count >= 2 Then
        Dim BestScore As Long, BestC As Integer
        BestScore = -2147483647 ' Min Long
        BestC = -1
        
        Dim VVar As Variant
        For Each VVar In NumCols
            c = VVar
            
            Dim Score As Long
            Score = NegC(c) * 10
            
            If VarC(c) Then Score = Score + 50
            
            Score = Score + (ScDec(c) * 5)
            
            ' Header Analysis (Row StartRow)
            Dim Hdr As String
            If StartRow <= Rows.Count Then
                If UBound(Rows(StartRow)) >= c Then
                    Hdr = LCase(ClnTx(CStr(Rows(StartRow)(c))))
                End If
            End If
            
            ' Explicit Keywords (German/English/Latin-roots)
            If InStr(Hdr, "betrag") > 0 And InStr(Hdr, "ursprung") = 0 Then
                Score = Score + 1000
            ElseIf InStr(Hdr, "amount") > 0 Or InStr(Hdr, "umsatz") > 0 Then
                Score = Score + 1000
            End If
            
            ' Penalize likely non-transaction columns
            If InStr(Hdr, "saldo") > 0 Then Score = Score - 500 ' Balance
            If InStr(Hdr, "valuta") > 0 Then Score = Score - 500 ' Date related
            
            If GlLog = True Then SLogi "ScanD: Score for Col " & c & " [" & Hdr & "] = " & Score & " (Neg=" & NegC(c) & " Var=" & VarC(c) & " Dec=" & ScDec(c) & ")"
            
            If Score > BestScore Then
                BestScore = Score
                BestC = c
            End If
        Next VVar
        IdxAm = BestC
    End If
    
    ' FORCE: Header Match Override (Trust Headers over Statistics)
    ' Prioritize "Umsatz" over generic "Betrag" (more specific for bank transactions)
    If HdrIdx_Umsatz > -1 Then
        IdxAm = HdrIdx_Umsatz
        If GlLog = True Then SLogi "ScanD: Forced to Header 'Umsatz' -> " & IdxAm
    ElseIf HdrIdx_Betrag > -1 Then
        IdxAm = HdrIdx_Betrag
        If GlLog = True Then SLogi "ScanD: Forced to Header 'Betrag' -> " & IdxAm
    End If
    
    If HdrIdx_Datum > -1 Then
        IdxDa = HdrIdx_Datum
        If GlLog = True Then SLogi "ScanD: Forced to Header 'Datum/Buchung' -> " & IdxDa
    End If
    
    ' FALLBACK 3: Max ScNum (Last Resort)
    If IdxAm = -1 And NumCols.Count > 0 Then
        Dim MaxN As Long
        MaxN = -1
        For Each VVar In NumCols
            c = VVar
            If ScNum(c) > MaxN Then
                MaxN = ScNum(c)
                IdxAm = c
            End If
        Next VVar
        If GlLog = True Then SLogi "ScanD: Fallback IdxAm=" & IdxAm & " with ScNum=" & MaxN
    End If
    
    ' Constant Columns (Robust version)
    ' This logic identifies columns where the value is the same across all sample rows.
    ' It performs a case-insensitive and whitespace-agnostic comparison.
    Dim IsCon As Boolean
    For c = 0 To MaxC
        IsCon = True
        Dim BaseV2 As String
        BaseV2 = ""
        
        ' Find the first non-empty, cleaned value to use as the base for comparison
        For R = StartRow + 1 To Limit
            If UBound(Rows(R)) >= c Then
                TmpV = ClnTx(Rows(R)(c))
                If Len(TmpV) > 0 Then
                    BaseV2 = LCase$(TmpV)
                    Exit For
                End If
            End If
        Next R

        ' If a base value was found, compare all other rows against it.
        ' If no base value was found, the column is empty in the sample, thus constant.
        If Len(BaseV2) > 0 Then
            If GlLog = True Then SLogi "ScanD-Const: BaseV2 for Col " & c & " is [" & BaseV2 & "]"
            For R = StartRow + 1 To Limit
                If UBound(Rows(R)) >= c Then
                    TmpV = ClnTx(Rows(R)(c))
                    
                    ' If a non-empty value differs from the base, the column is not constant.
                    If Len(TmpV) > 0 And LCase$(TmpV) <> BaseV2 Then
                        If GlLog = True Then
                            SLogi "ScanD-Const: Col " & c & " is NOT constant. Row " & R & " value [" & LCase$(TmpV) & "] <> BaseV2 [" & BaseV2 & "]"
                            SLogi "ScanD-Const: Raw value: [" & Rows(R)(c) & "]"
                        End If
                        IsCon = False
                        Exit For
                    End If
                End If
            Next R
        Else
            If GlLog = True Then SLogi "ScanD-Const: Col " & c & " has no base value, considered constant."
        End If

        ' If the column was determined to be constant, add it to the ignore collection.
        If IsCon Then
            If Not InCol(IgCol, c) Then
                On Error Resume Next
                IgCol.Add c, "K" & CStr(c)
                If GlLog = True Then SLogi "ScanD: Ignored constant column " & c
                On Error GoTo 0
            End If
        End If
    Next c

    ' NEW: Ignore redundant Date columns to prevent leaks in Description
    ' If a column scores high as a Date but is not the primary Date column, ignore it
    For c = 0 To MaxC
        If c <> IdxDa Then
            ' If > 50% of rows parsed as valid dates
            If ScDat(c) > ((Limit - StartRow) / 2) Then
                If Not InCol(IgCol, c) Then
                    If GlLog = True Then SLogi "ScanD: Ignored redundant Date Column " & c & " Score=" & ScDat(c)
                    On Error Resume Next
                    IgCol.Add c, "K" & CStr(c)
                    On Error GoTo 0
                End If
            End If
        End If
    Next c

    If GlLog = True Then SLogi "ScanD: Final IdxDa=" & IdxDa & " IdxAm=" & IdxAm & " NumCols.Count=" & NumCols.Count
End Sub

Private Function InCol(ByVal Col As Collection, ByVal Key As Variant) As Boolean
    InCol = False

    ' Guard against Null, Empty, or invalid Key values
    If IsNull(Key) Then Exit Function
    If IsEmpty(Key) Then Exit Function
    If Col Is Nothing Then Exit Function
    If Col.Count = 0 Then Exit Function

    ' Iterate through collection to find the value (safest VB6 method)
    ' The collection stores column indices as values with "K" & index as key
    Dim i As Long
    Dim StoredVal As Long
    For i = 1 To Col.Count
        StoredVal = Col(i)
        If StoredVal = CLng(Key) Then
            InCol = True
            Exit Function
        End If
    Next i
End Function

Private Function FixConc(ByVal RawTxt As String, ByVal Delim As String) As String
    ' Fix files with concatenated records (missing line breaks)
    ' Detects when the first column value repeats mid-line and inserts breaks
    'Schneller: jeder Kandidat wird nur einmal gezaehlt, die Zaehlung endet beim Schwellwert, das Ergebnis
    'entsteht per Join (vorher wurde fuer jede Zeile die ganze Datei durchsucht; Ausgabe unveraendert)
    Dim Lines() As String
    Dim i As Long
    Dim Ln As String
    Dim FstCol As String
    Dim SrcPat As String
    Dim PatPos As Long
    Dim SecPos As Long
    Dim PatCnt As Long
    Dim SplitN As Long
    Dim FirstLn As Long
    Dim Tried As Object

    If GlLog = True Then SLogi "FixConc: Starting analysis, Delim=[" & Delim & "]"

    ' Normalize line endings first
    RawTxt = Replace(RawTxt, vbCrLf, vbLf)
    RawTxt = Replace(RawTxt, vbCr, vbLf)
    Lines = Split(RawTxt, vbLf)

    If GlLog = True Then SLogi "FixConc: Total lines before fix: " & (UBound(Lines) + 1)

    ' Find first non-empty line with data to detect the repeating column
    Set Tried = Nothing
    FstCol = ""
    For i = 0 To UBound(Lines)
        Ln = Trim$(Lines(i))
        If Len(Ln) > 10 Then
            ' Check if line has the delimiter
            PatPos = InStr(Ln, Delim)
            If PatPos > 1 Then
                FstCol = Left$(Ln, PatPos - 1)
                ' Remove BOM if present (UTF-8 BOM: EF BB BF or Unicode BOM)
                If Len(FstCol) > 0 Then
                    If Asc(Left$(FstCol, 1)) = 239 Or Asc(Left$(FstCol, 1)) = 254 Or Asc(Left$(FstCol, 1)) = 255 Then
                        FstCol = Mid$(FstCol, 2)
                    End If
                    ' Check for UTF-8 BOM sequence (3 bytes at start)
                    If Len(FstCol) >= 3 Then
                        If Asc(Mid$(FstCol, 1, 1)) = 239 And Asc(Mid$(FstCol, 2, 1)) = 187 And Asc(Mid$(FstCol, 3, 1)) = 191 Then
                            FstCol = Mid$(FstCol, 4)
                        End If
                    End If
                End If
                ' Remove quotes if present
                If Left$(FstCol, 1) = Chr$(34) Then
                    FstCol = Mid$(FstCol, 2)
                End If
                If Right$(FstCol, 1) = Chr$(34) Then
                    FstCol = Left$(FstCol, Len(FstCol) - 1)
                End If

                If GlLog = True Then SLogi "FixConc: Line " & i & " FstCol=[" & FstCol & "] HasSpace=" & (InStr(FstCol, " ") > 0)

                ' Skip if it looks like a header (has spaces) or too short
                If Len(FstCol) > 5 And InStr(FstCol, " ") = 0 Then
                    'Ein bereits verworfener Kandidat wird nicht erneut gezaehlt (gleiche Zeichenfolge, gleiches Ergebnis)
                    If Tried Is Nothing Then Set Tried = CreateObject("Scripting.Dictionary")
                    If Not Tried.Exists(FstCol) Then
                        Tried.Add FstCol, True
                        ' Found a potential pattern - verify it appears multiple times
                        PatCnt = CountStr(RawTxt, FstCol & Delim, UBound(Lines) + 6)
                        If GlLog = True Then SLogi "FixConc: Pattern [" & FstCol & "] count=" & PatCnt & " threshold=" & (UBound(Lines) + 5)
                        If PatCnt > UBound(Lines) + 5 Then
                            If GlLog = True Then SLogi "FixConc: Pattern accepted!"
                            Exit For
                        End If
                    End If
                End If
                FstCol = ""
            End If
        End If
    Next i

    ' If no pattern found, return original
    If Len(FstCol) < 3 Then
        If GlLog = True Then SLogi "FixConc: No pattern found, returning original"
        FixConc = RawTxt
        Exit Function
    End If

    If GlLog = True Then SLogi "FixConc: Using pattern: [" & FstCol & "]"

    ' Build search pattern: the first column followed by delimiter
    SrcPat = FstCol & Delim
    SplitN = 0
    FirstLn = -1

    ' Process each line and split where pattern repeats
    For i = 0 To UBound(Lines)
        Ln = Lines(i)
        If Len(Ln) > 0 Then
            ' Check if pattern appears more than once in this line
            PatPos = InStr(Ln, SrcPat)
            If PatPos > 0 Then
                SecPos = InStr(PatPos + Len(SrcPat), Ln, SrcPat)
                ' Distance Check: multiple records on this line - split them
                If SecPos > 0 Then
                    If SecPos - PatPos > 20 Then
                        Lines(i) = SplitRec(Ln, SrcPat)
                        SplitN = SplitN + 1
                        If GlLog = True Then SLogi "FixConc: Split line " & i & " (PatPos=" & PatPos & " SecPos=" & SecPos & ")"
                    End If
                End If
            End If
            If FirstLn < 0 Then FirstLn = i
        End If
    Next i

    If GlLog = True Then SLogi "FixConc: Split " & SplitN & " lines"

    'Wie bisher: alle Zeilen ab der ersten nicht leeren, mit vbLf verbunden
    If FirstLn < 0 Then
        FixConc = vbNullString
    Else
        FixConc = Mid$(Join(Lines, vbLf), FirstLn + 1)
    End If
    If GlLog = True Then SLogi "FixConc: Total lines after fix: " & (UBound(Split(FixConc, vbLf)) + 1)
End Function

Private Function SplitRec(ByVal Ln As String, ByVal SrcPat As String) As String
    ' Split a line that contains multiple concatenated records
    ' MODIFIED: Ignores patterns found within 20 chars of the previous one
    Dim Result As String
    Dim PatPos As Long
    Dim NextP As Long
    Dim SegLen As Long

    Result = ""
    PatPos = 1

    Do
        ' Find next occurrence of pattern after current position
        NextP = InStr(PatPos, Ln, SrcPat)
        
        ' Filter: If next occurrence is too close (< 20 chars), it's likely a second column.
        Do While NextP > 0 And (NextP - PatPos < 20) And NextP <> PatPos
             NextP = InStr(NextP + 1, Ln, SrcPat)
        Loop

        If NextP = 0 Then
            ' No more patterns - add remainder
            If PatPos <= Len(Ln) Then
                If Len(Result) > 0 Then Result = Result & vbLf
                Result = Result & Mid$(Ln, PatPos)
            End If
            Exit Do
        End If

        If NextP = PatPos Then
            ' Pattern at start - find next occurrence
            NextP = InStr(PatPos + Len(SrcPat), Ln, SrcPat)
            
            ' Filter loop for the subsequent occurrence too
            Do While NextP > 0 And (NextP - PatPos < 20)
                 NextP = InStr(NextP + 1, Ln, SrcPat)
            Loop
            
            If NextP = 0 Then
                ' Only one record - add full line
                If Len(Result) > 0 Then Result = Result & vbLf
                Result = Result & Mid$(Ln, PatPos)
                Exit Do
            End If
            ' Add segment up to next pattern
            If Len(Result) > 0 Then Result = Result & vbLf
            Result = Result & Mid$(Ln, PatPos, NextP - PatPos)
            PatPos = NextP
        Else
            ' Pattern found mid-segment - add up to pattern
            If Len(Result) > 0 Then Result = Result & vbLf
            Result = Result & Mid$(Ln, PatPos, NextP - PatPos)
            PatPos = NextP
        End If
    Loop

    SplitRec = Result
End Function

Private Function CountStr(ByVal Source As String, ByVal Search As String, ByVal Limit As Long) As Long
    ' Count occurrences of Search in Source
    'Limit > 0: die Zaehlung endet nach Limit Treffern (0 = alle zaehlen)
    Dim Cnt As Long
    Dim Pos As Long

    Cnt = 0
    Pos = 1
    Do
        Pos = InStr(Pos, Source, Search)
        If Pos = 0 Then Exit Do
        Cnt = Cnt + 1
        If Cnt = Limit Then Exit Do
        Pos = Pos + 1
    Loop
    CountStr = Cnt
End Function

Private Function ClnTx(ByVal s As String) As String
    Dim Res As String
    Res = s
    ' Remove UTF-8 BOM if present
    If Len(Res) > 0 Then
        If Asc(Left$(Res, 1)) = 239 Or Asc(Left$(Res, 1)) = 254 Or Asc(Left$(Res, 1)) = 255 Then
            Res = Mid$(Res, 2)
        End If
    End If
    If Len(Res) >= 3 Then
        If Asc(Mid$(Res, 1, 1)) = 239 And Asc(Mid$(Res, 2, 1)) = 187 And Asc(Mid$(Res, 3, 1)) = 191 Then
            Res = Mid$(Res, 4)
        End If
    End If
    
    ' Remove problematic characters
    Res = Replace(Res, """", "") ' Quotes
    Res = Replace(Res, Chr(0), "") ' Null characters

    ' Normalize various whitespace characters to a standard space
    Res = Replace(Res, vbTab, " ")
    Res = Replace(Res, Chr(160), " ") ' Non-breaking space
    
    ' Replace HTML line break tags with spaces
    Res = Replace(Res, "<BR>", " ", , , vbTextCompare)
    Res = Replace(Res, "<SC>", " ", , , vbTextCompare)
    ' Clean up multiple spaces
    Do While InStr(Res, "  ") > 0
        Res = Replace(Res, "  ", " ")
    Loop
    ClnTx = Trim$(Res)
End Function

Private Function ClnAm(ByVal s As String) As String
    ' Clean and normalize amount strings for Val conversion (locale-independent)
    ' Handles German (1.234,56) and English (1,234.56) formats
    Dim Res As String
    Dim P1 As Long
    Dim P2 As Long

    Res = s
    Res = Replace(Res, "EUR", "")
    Res = Replace(Res, " ", "")
    Res = Replace(Res, Chr(34), "")
    
    P1 = InStr(Res, ".")
    P2 = InStr(Res, ",")

    If GlLog = True Then SLogi "ClnAm Input: [" & s & "] -> [" & Res & "] P1=" & P1 & " P2=" & P2

    If P1 > 0 And P2 > 0 Then
        ' Both separators present: determine which is thousands vs decimal
        If P1 < P2 Then
            ' Format: 1.234,56 (German) - dot is thousands, comma is decimal
            Res = Replace(Res, ".", "")
            Res = Replace(Res, ",", ".")
        Else
            ' Format: 1,234.56 (English) - comma is thousands, dot is decimal
            Res = Replace(Res, ",", "")
        End If
    ElseIf P1 > 0 And P2 = 0 Then
        ' Only dot present: check if thousands separator (3 digits after)
        If IsThouSep(Res, P1) Then
            ' Thousands separator: -3.300 -> -3300
            Res = Replace(Res, ".", "")
        End If
        ' If decimal separator, keep dot as-is (universal format)
    ElseIf P2 > 0 And P1 = 0 Then
        ' Only comma present: check if thousands separator (3 digits after)
        If IsThouSep(Res, P2) Then
            ' Thousands separator: -3,300 -> -3300
            Res = Replace(Res, ",", "")
        Else
            ' Decimal separator: -3,30 -> -3.30 (convert to universal dot format)
            Res = Replace(Res, ",", ".")
        End If
    End If

    If GlLog = True Then SLogi "ClnAm Output: [" & Res & "]"
    ClnAm = Res
End Function

Private Function IsThouSep(ByVal NumStr As String, ByVal SepPos As Long) As Boolean
    ' Determine if separator at SepPos is a thousands separator
    ' Returns True if exactly 3 digits follow the separator (no more separators after)
    Dim AfterS As String
    Dim DigCnt As Long
    Dim i As Long
    Dim c As String

    IsThouSep = False

    If SepPos < 1 Or SepPos >= Len(NumStr) Then Exit Function

    AfterS = Mid$(NumStr, SepPos + 1)
    DigCnt = 0

    ' Count consecutive digits after separator
    For i = 1 To Len(AfterS)
        c = Mid$(AfterS, i, 1)
        If c >= "0" And c <= "9" Then
            DigCnt = DigCnt + 1
        Else
            ' Non-digit found - stop counting
            Exit For
        End If
    Next i

    ' Thousands separator has exactly 3 digits after, no more separators
    ' Also check that we're at end of string or next char is not a digit
    If DigCnt = 3 Then
        ' Check there's no decimal part hidden after
        If Len(AfterS) = 3 Then
            IsThouSep = True
        ElseIf Len(AfterS) > 3 Then
            ' Check what follows the 3 digits
            c = Mid$(AfterS, 4, 1)
            If c <> "." And c <> "," Then
                ' Could be end of number or non-numeric char
                IsThouSep = True
            End If
        End If
    End If
End Function

Private Function IsUTF8(ByVal RawTxt As String) As Boolean
    ' Detect if string contains UTF-8 encoded data
    ' Checks for BOM or typical UTF-8 multi-byte sequences
    Dim i As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim UTF8Cnt As Long

    IsUTF8 = False
    If Len(RawTxt) < 2 Then Exit Function

    ' Check for UTF-8 BOM (EF BB BF)
    If Len(RawTxt) >= 3 Then
        b1 = Asc(Mid$(RawTxt, 1, 1))
        b2 = Asc(Mid$(RawTxt, 2, 1))
        b3 = Asc(Mid$(RawTxt, 3, 1))
        If b1 = 239 And b2 = 187 And b3 = 191 Then
            IsUTF8 = True
            Exit Function
        End If
    End If

    ' Scan for UTF-8 multi-byte sequences (German umlauts, etc.)
    ' UTF-8 2-byte: 110xxxxx 10xxxxxx (C0-DF followed by 80-BF)
    UTF8Cnt = 0
    For i = 1 To Len(RawTxt) - 1
        b1 = Asc(Mid$(RawTxt, i, 1))
        b2 = Asc(Mid$(RawTxt, i + 1, 1))

        ' Check for 2-byte UTF-8 sequence (covers ä ö ü ß etc.)
        If b1 >= 194 And b1 <= 223 Then
            If b2 >= 128 And b2 <= 191 Then
                UTF8Cnt = UTF8Cnt + 1
                If UTF8Cnt >= 3 Then
                    IsUTF8 = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function ConvUTF8(ByVal RawTxt As String) As String
    ' Convert UTF-8 encoded string to ANSI (Windows-1252)
    ' Uses ADODB.Stream for reliable conversion
    Dim oStm As Object
    Dim Bytes() As Byte

    On Error GoTo ErrHnd

    If Len(RawTxt) = 0 Then
        ConvUTF8 = ""
        Exit Function
    End If

    ' Convert VB string back to raw bytes
    Bytes = StrConv(RawTxt, vbFromUnicode)

    ' Use ADODB.Stream to decode UTF-8
    Set oStm = CreateObject("ADODB.Stream")

    ' Write bytes as binary
    oStm.Type = 1  ' adTypeBinary
    oStm.Open
    oStm.Write Bytes

    ' Read back as UTF-8 text
    oStm.Position = 0
    oStm.Type = 2  ' adTypeText
    oStm.Charset = "UTF-8"

    ConvUTF8 = oStm.ReadText

    oStm.Close
    Set oStm = Nothing

    If GlLog = True Then SLogi "ConvUTF8: Converted " & Len(RawTxt) & " -> " & Len(ConvUTF8) & " chars"
    Exit Function

ErrHnd:
    If GlLog = True Then SLogi "ConvUTF8 Error: " & Err.Number & " - " & Err.Description
    ConvUTF8 = RawTxt
End Function

Private Function ReadBankFile(ByVal Path As String, ByRef DecodedUnicode As Boolean) As String
    ' Reads a bank-statement file as raw bytes and decodes it with full encoding
    ' detection: UTF-16 LE/BE (with or without BOM). For ANSI and UTF-8 files it
    ' returns the legacy "bytes as CP-1252 chars" string so the existing
    ' IsUTF8/ConvUTF8 pipeline keeps working byte-for-byte unchanged.
    ' DecodedUnicode is set True ONLY for UTF-16; the caller must then skip ConvUTF8.
    On Error GoTo ErrHnd
    DecodedUnicode = False
    ReadBankFile = ""

    Dim n As Long
    n = FileLen(Path)
    If n <= 0 Then Exit Function

    Dim b() As Byte
    ReDim b(n - 1)
    Dim fNum As Integer
    fNum = FreeFile
    Open Path For Binary Access Read As #fNum
    Get #fNum, 1, b
    Close #fNum

    ' BOM-based detection
    If n >= 2 Then
        If b(0) = 255 And b(1) = 254 Then
            ReadBankFile = DecodeUTF16(b, 2, n, False)
            DecodedUnicode = True
            Exit Function
        ElseIf b(0) = 254 And b(1) = 255 Then
            ReadBankFile = DecodeUTF16(b, 2, n, True)
            DecodedUnicode = True
            Exit Function
        End If
    End If

    ' BOM-less UTF-16 heuristic: distribution of NUL bytes across even/odd offsets
    Dim lim As Long, i As Long
    Dim zEven As Long, zOdd As Long
    lim = n - 1
    If lim > 1999 Then lim = 1999
    For i = 0 To lim
        If b(i) = 0 Then
            If (i Mod 2) = 0 Then
                zEven = zEven + 1
            Else
                zOdd = zOdd + 1
            End If
        End If
    Next i

    Dim pairs As Long
    pairs = (lim + 1) \ 2
    If pairs > 0 Then
        If zOdd > zEven And zOdd >= (pairs \ 2) Then
            ReadBankFile = DecodeUTF16(b, 0, n, False)
            DecodedUnicode = True
            Exit Function
        ElseIf zEven > zOdd And zEven >= (pairs \ 2) Then
            ReadBankFile = DecodeUTF16(b, 0, n, True)
            DecodedUnicode = True
            Exit Function
        End If
    End If

    ' ANSI / UTF-8: legacy representation (no NUL-truncation)
    ReadBankFile = StrConv(b, vbUnicode)
    Exit Function

ErrHnd:
    If GlLog = True Then SLogi "ReadBankFile Error: " & Err.Number & " - " & Err.Description
    DecodedUnicode = False
    ReadBankFile = ""
End Function

Private Function DecodeUTF16(ByRef b() As Byte, ByVal StartIx As Long, ByVal nBytes As Long, ByVal BigEndian As Boolean) As String
    ' Decode UTF-16 bytes b(StartIx .. nBytes-1) into a VB6 Unicode string.
    On Error GoTo ErrHnd
    DecodeUTF16 = ""

    Dim cnt As Long
    cnt = nBytes - StartIx
    If cnt < 2 Then Exit Function
    If (cnt Mod 2) <> 0 Then cnt = cnt - 1

    Dim tmp() As Byte
    ReDim tmp(cnt - 1)
    Dim i As Long
    If BigEndian Then
        For i = 0 To cnt - 2 Step 2
            tmp(i) = b(StartIx + i + 1)
            tmp(i + 1) = b(StartIx + i)
        Next i
    Else
        For i = 0 To cnt - 1
            tmp(i) = b(StartIx + i)
        Next i
    End If

    Dim s As String
    s = tmp
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then s = Mid$(s, 2)
    End If
    DecodeUTF16 = s
    Exit Function

ErrHnd:
    If GlLog = True Then SLogi "DecodeUTF16 Error: " & Err.Number & " - " & Err.Description
    DecodeUTF16 = ""
End Function


Private Sub FindSister(ByVal Rows As Collection, ByVal IdxDa As Integer, ByVal IdxAm As Integer, ByRef IdxAm2 As Integer, Optional ByVal StartRow As Long = 1)
    ' Detect a Soll/Haben sister column adjacent to IdxAm.
    ' Triggers ONLY when ALL of the following hold:
    '   - sister is direct neighbour of IdxAm (and not the date column)
    '   - both columns each carry >= 3 non-zero values in the sample window
    '   - no row has both columns simultaneously non-zero (mutually exclusive)
    '   - IdxAm is overwhelmingly negative, sister is overwhelmingly positive
    ' Conservative by design: single-column formats (DKB, Sparkasse, Comdirect,
    ' StarMoney etc.) cannot satisfy condition 4 because they mix +/- in one column.
    IdxAm2 = -1
    If IdxAm < 0 Then Exit Sub
    If Rows Is Nothing Then Exit Sub
    If Rows.Count = 0 Then Exit Sub
    If StartRow < 1 Then StartRow = 1
    If StartRow > Rows.Count Then Exit Sub

    Dim MaxC As Long
    On Error Resume Next
    MaxC = UBound(Rows(StartRow))
    On Error GoTo 0
    If MaxC < 1 Then Exit Sub

    Dim Lim As Long
    Lim = StartRow + 50
    If Lim > Rows.Count Then Lim = Rows.Count

    Dim Cands(1) As Integer
    Cands(0) = IdxAm + 1
    Cands(1) = IdxAm - 1

    Dim k As Integer
    For k = 0 To 1
        Dim cs As Integer
        cs = Cands(k)
        If cs >= 0 And cs <= MaxC And cs <> IdxDa And cs <> IdxAm Then
            Dim CntA As Long, CntB As Long, CntBoth As Long
            Dim NegA As Long, PosB As Long
            CntA = 0: CntB = 0: CntBoth = 0: NegA = 0: PosB = 0

            Dim r As Long
            For r = StartRow + 1 To Lim
                Dim RowV As Variant
                RowV = Rows(r)
                Dim Ub As Long
                Ub = -1
                On Error Resume Next
                Ub = UBound(RowV)
                On Error GoTo 0

                If Ub >= IdxAm And Ub >= cs Then
                    Dim sA As String, sB As String
                    sA = ClnAm(ClnTx(CStr(RowV(IdxAm))))
                    sB = ClnAm(ClnTx(CStr(RowV(cs))))

                    Dim hA As Boolean, hB As Boolean
                    hA = False: hB = False
                    If Len(sA) > 0 And IsNumeric(sA) Then
                        If Val(sA) <> 0 Then hA = True
                    End If
                    If Len(sB) > 0 And IsNumeric(sB) Then
                        If Val(sB) <> 0 Then hB = True
                    End If

                    If hA And hB Then
                        CntBoth = CntBoth + 1
                    ElseIf hA Then
                        CntA = CntA + 1
                        If Val(sA) < 0 Then NegA = NegA + 1
                    ElseIf hB Then
                        CntB = CntB + 1
                        If Val(sB) > 0 Then PosB = PosB + 1
                    End If
                End If
            Next r

            If GlLog = True Then SLogi "FindSister: Cand=" & cs & " CntA=" & CntA & " CntB=" & CntB & " Both=" & CntBoth & " NegA=" & NegA & " PosB=" & PosB

            ' Conservative match: 3+ rows in each, none-overlap, sign separation
            If CntA >= 3 And CntB >= 3 And CntBoth = 0 Then
                If NegA >= CntA - 1 And PosB >= CntB - 1 Then
                    IdxAm2 = cs
                    If GlLog = True Then SLogi "FindSister: Sister column detected -> " & IdxAm2
                    Exit Sub
                End If
            End If
        End If
    Next k
End Sub

Private Function BankEmpty(ByVal Fields As Collection) As Boolean
    Dim n As Long
    BankEmpty = True
    For n = 1 To Fields.Count
        If Len(Trim$(Replace(CStr(Fields(n)), vbTab, vbNullString))) > 0 Then
            BankEmpty = False
            Exit Function
        End If
    Next n
End Function

Private Function BankCols(ByVal Rows As Collection, ByRef ColMap() As Long, ByVal HdrRow As Long) As Boolean
    'Erkennt das vollstaendige CSV-Schema der Volks- und Raiffeisenbanken (z.B. i300)
    'Andere Exporte behalten die allgemeine Spaltenerkennung in Imp01
    Dim Names As Variant, Header As Variant
    Dim Duplicate As Boolean
    Dim c As Long, n As Long
    Dim HdrTxt As String
    If HdrRow < 1 Or HdrRow > Rows.Count Then Exit Function
    Header = Rows(HdrRow)
    Names = Array("iban auftragskonto", "buchungstag", "valutadatum", "name zahlungsbeteiligter", _
                  "iban zahlungsbeteiligter", "bic (swift-code) zahlungsbeteiligter", "buchungstext", _
                  "verwendungszweck", "betrag", "waehrung", "saldo nach buchung", "glaeubiger id", "mandatsreferenz")
    ReDim ColMap(UBound(Names))
    For n = 0 To UBound(Names)
        ColMap(n) = -1
    Next n
    For c = 0 To UBound(Header)
        HdrTxt = LCase$(ClnTx(CStr(Header(c))))
        HdrTxt = Replace(HdrTxt, ChrW(228), "ae")
        For n = 0 To UBound(Names)
            If HdrTxt = Names(n) Then
                If ColMap(n) <> -1 Then Duplicate = True
                ColMap(n) = c
            End If
        Next n
    Next c
    For n = 0 To UBound(Names)
        If ColMap(n) < 0 Then Exit Function
    Next n
    If Duplicate Then Err.Raise vbObjectError + 2210, "BankCols", "Eine Bankspalte ist mehrfach vorhanden."
    BankCols = True
End Function

Private Function BankField(ByVal Value As String) As String
    Value = Replace(Value, vbCr, " ")
    Value = Replace(Value, vbLf, " ")
    Value = Replace(Value, vbTab, " ")
    Value = Replace(Value, ChrW(160), " ")
    Do While InStr(Value, "  ") > 0
        Value = Replace(Value, "  ", " ")
    Loop
    BankField = Trim$(Value)
End Function

Private Function BankAmt(ByVal Value As String, ByRef Amount As Currency) As Boolean
    'Bankbetrag locale-unabhaengig lesen; False bei leerem oder ungueltigem Wert
    Dim Clean As String
    Clean = ClnAm(ClnTx(Value))
    If Len(Clean) = 0 Then Exit Function
    If Not IsNumeric(Clean) Then Exit Function
    Amount = CCur(Round(Val(Clean), 2))
    BankAmt = True
End Function

Private Function BankKey(ByVal RowV As Variant, ByRef ColMap() As Long, ByVal IdBnk As Long, ByVal ManNr As Long, ByVal Prov As Long) As String
    'Stabile Kennung eines Umsatzes; Format BKI1 unveraendert, bereits gespeicherte Kennungen bleiben gueltig
    Dim n As Long
    Dim Value As String, KeyText As String
    Dim Amt As Currency
    KeyText = CStr(IdBnk) & ":" & CStr(ManNr) & ":"
    For n = 0 To UBound(ColMap)
        Value = BankField(CStr(RowV(ColMap(n))))
        Select Case n
        Case 0, 4, 5, 9, 11
            Value = UCase$(Replace(Value, " ", vbNullString))
        Case 1, 2
            'Datum in bisheriger Lesart (fiktive Valuta), damit gespeicherte Kennungen gueltig bleiben
            Value = CStr(CLng(DateValue(ParseDate(Value, True))))
        Case 8, 10
            If BankAmt(Value, Amt) Then Value = Trim$(Str$(Amt))
        End Select
        KeyText = KeyText & CStr(Len(Value)) & ":" & Value
    Next n
    BankKey = "BKI1:" & BankHash(Prov, KeyText)
End Function

Private Function BankHash(ByVal Prov As Long, ByVal KeyText As String) As String
    'SHA-256 (CryptoAPI) ueber den Kennungstext; der Kontext wird einmal je Import geoeffnet
    Dim Hash As Long, Size As Long, n As Long
    Dim HashOk As Boolean
    Dim Bytes(31) As Byte
    If BnkCreate(Prov, &H800C&, 0, 0, Hash) <> 0 Then
        If BnkHash(Hash, StrPtr(KeyText), LenB(KeyText), 0) <> 0 Then
            Size = 32
            If BnkGet(Hash, 2, Bytes(0), Size, 0) <> 0 Then HashOk = (Size = 32)
        End If
        BnkDestroy Hash
    End If
    If Not HashOk Then Err.Raise vbObjectError + 2211, "BankHash", "Die Bankkennung konnte nicht erstellt werden. Es wurden keine Umsaetze importiert."
    For n = 0 To 31
        BankHash = BankHash & Right$("0" & Hex$(Bytes(n)), 2)
    Next n
End Function

Private Function BankText(ByVal RowV As Variant, ByRef ColMap() As Long) As String
    Dim n As Variant, Value As String
    For Each n In Array(3, 4, 5, 6, 7, 10, 11, 12)
        Value = BankField(CStr(RowV(ColMap(n))))
        If Len(Value) > 0 Then
            If Len(BankText) > 0 Then BankText = BankText & " "
            BankText = BankText & Value
        End If
    Next n
End Function

Private Function BankBucket(ByVal BookDay As Date, ByVal Amount As Double) As String
    ' Match the legacy REAL/Single amount representation, scoped by BankLoad.
    BankBucket = CStr(CLng(DateValue(BookDay))) & ":" & CStr(CSng(Amount))
End Function

Private Sub BankSpan(ByVal BkDay As Date, ByRef MinDay As Date, ByRef MaxDay As Date)
    'Erweitert den Zeitraum des Bestandsabgleichs um ein gueltiges Datum
    If Year(BkDay) <= 1900 Then Exit Sub
    If Year(MinDay) <= 1900 Or BkDay < MinDay Then MinDay = BkDay
    If BkDay > MaxDay Then MaxDay = BkDay
End Sub

Private Function BankLoad(ByVal IdBnk As Long, ByVal ManNr As Long, ByVal MinDay As Date, ByVal MaxDay As Date) As ADODB.Recordset
    'Liest den Bestand des Kontos im Zeitraum einmal je Import
    'Literale per SqlNum/SqlDat: adDBTimeStamp-Parameter vertauschen unter Jet/ACE Tag und Monat (Tag <= 12)
    On Error GoTo LoadErr
    Dim OldRows As ADODB.Recordset
    Dim SqlText As String
    Dim ErrNum As Long
    SqlText = "SELECT [ID2], [Datum], [Betrag], [IDKurz], [Suche] FROM [Tabelle_Banking] WHERE [IDB] = " & SqlNum(IdBnk) & _
              " AND [IDM] = " & SqlNum(ManNr) & " AND [Datum] >= " & SqlDat(DateValue(MinDay)) & _
              " AND [Datum] < " & SqlDat(DateAdd("d", 1, DateValue(MaxDay)))
    Set OldRows = New ADODB.Recordset
    OldRows.CursorLocation = adUseClient
    OldRows.Open SqlText, DB1, adOpenStatic, adLockReadOnly, adCmdText
    Set OldRows.ActiveConnection = Nothing
    Set BankLoad = OldRows
    Exit Function
LoadErr:
    ErrNum = Err.Number
    On Error Resume Next
    If Not OldRows Is Nothing Then
        If OldRows.State <> adStateClosed Then OldRows.Close
    End If
    Set OldRows = Nothing
    On Error GoTo 0
    Err.Raise ErrNum, "BankLoad", "Vorhandene Bankumsaetze konnten nicht gelesen werden. Der Import wurde abgebrochen."
End Function

Private Function BankMatch(ByVal Cells As Variant, ByVal Pos As Long, ByVal Prefix As String, ByVal Target As String) As Boolean
    Dim Added As String
    If Left$(Prefix, 255) = Target Then
        BankMatch = True
        Exit Function
    End If
    If Pos > UBound(Cells) Or Len(Prefix) > Len(Target) Then Exit Function
    If Left$(Target, Len(Prefix)) <> Prefix Then Exit Function
    If Len(Cells(Pos)) > 0 Then
        Added = Prefix
        If Len(Added) > 0 Then Added = Added & " "
        Added = Added & Cells(Pos)
        If BankMatch(Cells, Pos + 1, Added, Target) Then
            BankMatch = True
            Exit Function
        End If
    End If
    BankMatch = BankMatch(Cells, Pos + 1, Prefix, Target)
End Function

Private Function BankCells(ByVal RowV As Variant, ByRef ColMap() As Long) As Variant
    'Textbausteine wie im bisherigen Import (ClnTx, "?" als Leerzeichen) fuer den Altbestandsabgleich
    'Folge: Buchungstag, Valuta, Name, IBAN, BIC, Buchungstext, Zweck, Waehrung, Saldo, Glaeubiger, Mandat
    Dim Cells(10) As String
    Dim Indexes As Variant
    Dim n As Long
    Indexes = Array(1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12)
    For n = 0 To 10
        Cells(n) = Replace(ClnTx(CStr(RowV(ColMap(Indexes(n))))), "?", " ")
    Next n
    BankCells = Cells
End Function

Private Function BankLegacy(ByVal Cells As Variant, ByVal OldText As String) As Long
    'Vergleich mit Altbestand ohne stabile Kennung (Datum und Betrag stimmen bereits ueberein):
    '0 = anderer Umsatz, 1 = sicher derselbe, 2 = derselbe bei gekuerztem oder leerem Alttext
    If OldText = "kein Buchungstext" Then
        BankLegacy = 2
        Exit Function
    End If
    If Not BankMatch(Cells, 0, vbNullString, OldText) Then Exit Function
    BankLegacy = 2
    If Len(OldText) >= 255 Then Exit Function
    'Saldo nach Buchung unterscheidet sonst gleiche Umsaetze eines Tages
    If Len(Cells(8)) > 0 Then
        If InStr(1, OldText, Cells(8), vbBinaryCompare) > 0 Then BankLegacy = 1
    ElseIf Len(Cells(6)) > 0 Then
        If InStr(1, OldText, Cells(6), vbBinaryCompare) > 0 Then BankLegacy = 1
    End If
End Function

Private Sub BankImport(ByVal Rows As Collection, ByRef ColMap() As Long, ByVal HdrRow As Long, ByRef Target As ADODB.Recordset, ByVal IdBnk As Long, ByVal ManNr As Long, ByVal MitNr As Long)
    On Error GoTo BankErr
    'Import im VR-Schema: stabile Kennung je Umsatz in [Suche], Abgleich mit dem Bestand vor dem
    'ersten Schreibzugriff; Zeilen ohne gueltiges Datum oder Betrag werden wie bisher uebersprungen
    Dim Plans As Collection, Bucket As Collection
    Dim Known As Object, Seen As Object, Legacy As Object, Used As Object, Extra As Object
    Dim RowV As Variant, Plan As Variant, OldRec As Variant, Cells As Variant
    Dim State() As Byte, Sure() As String, Weak() As String, Texts() As String
    Dim BookDay As Date, ValueDay As Date, StoreDay As Date, MinDay As Date, MaxDay As Date, OldDay As Date
    Dim Amount As Currency
    Dim KeyText As String, BaseKey As String, DayKey As String, Done As String
    Dim RawBook As String, RawValue As String
    Dim n As Long, k As Long, b As Long, c As Long, HdrUB As Long, Prov As Long
    Dim Match As Long, Skipped As Long, Added As Long, BadRows As Long, SureCnt As Long, WeakCnt As Long, Ambig As Long
    Dim SureAt As Long, WeakAt As Long
    Dim StateN() As Byte, SureN() As String, WeakN() As String
    Dim UsedN As Object
    Dim Fit As Boolean
    Dim ErrNum As Long, ErrText As String

    Set Plans = New Collection
    Set Known = CreateObject("Scripting.Dictionary")
    Set Seen = CreateObject("Scripting.Dictionary")
    Set Legacy = CreateObject("Scripting.Dictionary")
    Set Used = CreateObject("Scripting.Dictionary")
    Set Extra = CreateObject("Scripting.Dictionary")
    If BnkAcquire(Prov, vbNullString, vbNullString, 24, &HF0000000) = 0 Then
        Prov = 0
        Err.Raise vbObjectError + 2211, "BankImport", "Die Bankkennung konnte nicht erstellt werden. Es wurden keine Umsaetze importiert."
    End If

    '1. Datei pruefen und planen: ein Durchlauf, noch kein Schreibzugriff
    HdrUB = UBound(Rows(HdrRow))
    For Each RowV In Rows
        n = n + 1
        If n > HdrRow Then
            'Zusaetzliche Felder hinter der Kopfzeile (z.B. Trennzeichen am Zeilenende) nur, wenn alle leer sind
            Fit = (UBound(RowV) = HdrUB)
            If UBound(RowV) > HdrUB Then
                Fit = True
                For c = HdrUB + 1 To UBound(RowV)
                    If Len(BankField(ClnTx(CStr(RowV(c))))) > 0 Then Fit = False
                Next c
            End If
            If Not Fit Then
                BadRows = BadRows + 1
            Else
                RawBook = ClnTx(CStr(RowV(ColMap(1))))
                RawValue = ClnTx(CStr(RowV(ColMap(2))))
                BookDay = ParseDate(RawBook)
                ValueDay = ParseDate(RawValue)
                'Datum wie im bisherigen CSV-Import und bei MT940: Valuta, ersatzweise Buchungstag
                If Year(ValueDay) > 1900 Then StoreDay = ValueDay Else StoreDay = BookDay
                If Year(StoreDay) <= 1900 Then
                    BadRows = BadRows + 1
                ElseIf Not BankAmt(CStr(RowV(ColMap(8))), Amount) Then
                    BadRows = BadRows + 1
                Else
                    'Gleiche Zeilen einer Datei werden gezaehlt, damit Wiederholungen nichts verdoppeln
                    BaseKey = BankKey(RowV, ColMap, IdBnk, ManNr, Prov)
                    If Seen.Exists(BaseKey) Then Seen(BaseKey) = Seen(BaseKey) + 1 Else Seen.Add BaseKey, 1&
                    KeyText = BaseKey & ":" & CStr(Seen(BaseKey))
                    '(7), (8): Buchungstag und Valuta in bisheriger Lesart, unter der ein frueherer Import
                    'fiktive Valuta gespeichert hat (31.04.2025 -> 01.05.2025, 29.02.25 -> 25.02.2029)
                    Plans.Add Array(RowV, StoreDay, BookDay, ValueDay, Amount, KeyText, BankText(RowV, ColMap), ParseDate(RawBook, True), ParseDate(RawValue, True))
                    BankSpan BookDay, MinDay, MaxDay
                    BankSpan ValueDay, MinDay, MaxDay
                End If
            End If
        End If
    Next RowV
    If Plans.Count > 0 Then
        '2. Bestand des Kontos im Zeitraum einmal lesen, dazu die Tage bisheriger Lesart ausserhalb davon
        For Each Plan In Plans
            For b = 7 To 8
                OldDay = Plan(b)
                If Year(OldDay) > 1900 Then
                    OldDay = DateValue(OldDay)
                    If OldDay < DateValue(MinDay) Or OldDay > DateValue(MaxDay) Then Extra(CStr(CLng(OldDay))) = OldDay
                End If
            Next b
        Next Plan
        BankRead IdBnk, ManNr, MinDay, MaxDay, Extra, Legacy, Known

        '3. Abgleich: stabile Kennung, dann sichere Treffer im Altbestand unter Valuta- und Buchungstag, auch in
        '   bisheriger Lesart (dort nur mit Buchungstag, Valuta oder Saldo der Zeile im Alttext; je BankAssign), danach
        '   Treffer mit gekuerztem oder leerem Alttext (BankAssign); jeder Altumsatz deckt hoechstens eine Zeile ab
        ReDim State(1 To Plans.Count)
        ReDim Sure(1 To Plans.Count)
        ReDim Weak(1 To Plans.Count)
        ReDim Texts(1 To Plans.Count)
        ReDim SureN(1 To Plans.Count)
        ReDim WeakN(1 To Plans.Count)
        k = 0
        For Each Plan In Plans
            k = k + 1
            'Speichertag und Text wie in AddRec gespeichert (BankAssign: gleiche Zeilen sind austauschbar)
            Texts(k) = Replace(CStr(Plan(6)), "?", " ")
            If Len(Texts(k)) = 0 Then Texts(k) = "kein Buchungstext" Else Texts(k) = Left$(Texts(k), 255)
            Texts(k) = CStr(CLng(DateValue(Plan(1)))) & "|" & Texts(k)
            If Known.Exists(Plan(5)) Then
                State(k) = 1
            ElseIf Legacy.Count > 0 Then
                Cells = Empty
                Done = "|"
                SureAt = 0
                WeakAt = 0
                For b = 1 To 4
                    Select Case b
                    Case 1: DayKey = BankBucket(Plan(3), Plan(4))
                    Case 2: DayKey = BankBucket(Plan(2), Plan(4))
                    Case 3: DayKey = BankBucket(Plan(8), Plan(4))
                    Case Else: DayKey = BankBucket(Plan(7), Plan(4))
                    End Select
                    If InStr(Done, "|" & DayKey & "|") = 0 Then
                        Done = Done & DayKey & "|"
                        If Legacy.Exists(DayKey) Then
                            If IsEmpty(Cells) Then Cells = BankCells(Plan(0), ColMap)
                            Set Bucket = Legacy(DayKey)
                            For Each OldRec In Bucket
                                If Not Used.Exists(OldRec(0)) Then
                                    Match = BankLegacy(Cells, CStr(OldRec(1)))
                                    'Tag bisheriger Lesart: ohne Datum oder Saldo der Zeile im Alttext koennte eine
                                    'echte Buchung dieses Tages gemeint sein (auch ein leerer Alttext belegt nichts)
                                    If b > 2 Then
                                        If Not BankProof(Cells, CStr(OldRec(1))) Then Match = 0
                                    End If
                                    'Zwei Sichten (siehe unten): alle Tage (Sure/Weak) und je Stufe nur der erste
                                    'Tag mit Treffern (SureN/WeakN)
                                    If Match = 1 Then
                                        Sure(k) = Sure(k) & " " & OldRec(0)
                                        If SureAt = 0 Or SureAt = b Then
                                            SureN(k) = SureN(k) & " " & OldRec(0)
                                            SureAt = b
                                        End If
                                    ElseIf Match = 2 Then
                                        Weak(k) = Weak(k) & " " & OldRec(0)
                                        If WeakAt = 0 Or WeakAt = b Then
                                            WeakN(k) = WeakN(k) & " " & OldRec(0)
                                            WeakAt = b
                                        End If
                                    End If
                                End If
                            Next OldRec
                        End If
                    End If
                Next b
            End If
        Next Plan
        'Sichere Treffer unter echtem Tag und Tag bisheriger Lesart in einer Stufe: konkurrieren verschiedene Zeilen
        'um einen Altumsatz (z.B. gleicher Saldo bei April- und Maibuchung), ist nicht entscheidbar, welche er ist
        'Zwei Sichten: mit allen Tagen kann ein gleichlautender Altumsatz eines anderen Tages Ausweichpartner einer
        'Zeile werden (eine neue Zeile verbraucht dann deren Altumsatz); mit nur dem ersten Tag je Stufe faellt der
        'Altumsatz einer Zeile weg, deren Valuta auf den Altumsatz einer anderen faellt. Vorhanden ist eine Zeile nur,
        'wenn beide Sichten sie zuordnen, sonst wird sie gespeichert (State 2, Dublette statt Verlust)
        ReDim StateN(1 To Plans.Count)
        For k = 1 To Plans.Count
            StateN(k) = State(k)
        Next k
        Set UsedN = CreateObject("Scripting.Dictionary")
        SureCnt = BankAssign(Sure, State, Used, Texts)
        WeakCnt = BankAssign(Weak, State, Used, Texts)
        BankAssign SureN, StateN, UsedN, Texts
        BankAssign WeakN, StateN, UsedN, Texts
        For k = 1 To Plans.Count
            If State(k) = 1 And StateN(k) <> 1 Then State(k) = 2
        Next k

        '4. Nur neue Umsaetze speichern
        k = 0
        For Each Plan In Plans
            k = k + 1
            If State(k) = 1 Then
                Skipped = Skipped + 1
            Else
                If State(k) = 2 Then Ambig = Ambig + 1
                AddRec Target, IdBnk, ManNr, MitNr, CDate(Plan(1)), CDbl(Plan(4)), CStr(Plan(6)), CStr(Plan(5))
                Added = Added + 1
                frmStatus.prbStat1.Value = Added
                DoEvents
            End If
        Next Plan
    End If

    BnkRelease Prov, 0
    Prov = 0
    If GlLog = True Then SLogi "BankImport: Added=" & CStr(Added) & "; Skipped=" & CStr(Skipped) & "; LegacySure=" & CStr(SureCnt) & "; LegacyWeak=" & CStr(WeakCnt) & "; Unklar=" & CStr(Ambig) & "; Invalid=" & CStr(BadRows) & "; ExtraDays=" & CStr(Extra.Count)
    Exit Sub

BankErr:
    ErrNum = Err.Number
    ErrText = Err.Description
    On Error Resume Next
    If Prov <> 0 Then BnkRelease Prov, 0
    On Error GoTo 0
    Err.Raise ErrNum, "BankImport", ErrText
End Sub

Private Function BankParts(ByVal RowV As Variant, ByVal IdxDa As Integer, ByVal IdxAm As Integer) As Variant
    'Textbausteine einer Zeile wie im allgemeinen CSV-Import (ClnTx, "?" als Leerzeichen), ohne Datum und Betrag
    Dim Cells() As String
    Dim j As Long
    ReDim Cells(UBound(RowV))
    For j = 0 To UBound(RowV)
        If j <> IdxDa And j <> IdxAm Then Cells(j) = BankNorm(Replace(ClnTx(CStr(RowV(j))), "?", " "))
    Next j
    BankParts = Cells
End Function

Private Function BankNorm(ByVal Value As String) As String
    'Vergleichstext: vorgemerkte Umsaetze (Sparkasse: Info "Umsatz vorgemerkt") gelten als gebucht; Sparkasse
    'CSV-CAMT V2 kuerzt Buchungstexte, die V8 ausschreibt. Der gespeicherte Text bleibt wie geliefert
    Value = Replace(Value, "Umsatz vorgemerkt", "Umsatz gebucht", 1, -1, vbTextCompare)
    Value = Replace(Value, "UEBERTRAG (GUTSCHR. UEBERW)", "UEBERTRAG (GUTSCHRIFT UEBERWEISUNG)", 1, -1, vbTextCompare)
    BankNorm = Replace(Value, "GUTSCHR. UEBERWEISUNG", "GUTSCHRIFT UEBERWEISUNG", 1, -1, vbTextCompare)
End Function

Private Function BankFits(ByVal Cells As Variant, ByVal Target As String, ByVal Cut As Boolean) As Boolean
    'True, wenn Target aus einer Auswahl der Zellen in Spaltenfolge entsteht; Leerzeichen zaehlen nicht
    '(Exportformate verbinden Textzeilen verschieden). Cut: Target ist ein auf 255 Zeichen gekuerzter Text.
    'Merkt erreichbare Textlaengen statt Rekursion
    Dim Reach() As Boolean
    Dim n As Long, L As Long, TLen As Long
    Dim Part As String
    Target = Replace(Target, " ", vbNullString)
    TLen = Len(Target)
    If TLen = 0 Then Exit Function
    ReDim Reach(TLen)
    Reach(0) = True
    For n = 0 To UBound(Cells)
        Part = Replace(Cells(n), " ", vbNullString)
        If Len(Part) > 0 Then
            For L = TLen - 1 To 0 Step -1
                If Reach(L) Then
                    If L + Len(Part) <= TLen Then
                        If Mid$(Target, L + 1, Len(Part)) = Part Then Reach(L + Len(Part)) = True
                    ElseIf Cut Then
                        If Mid$(Target, L + 1) = Left$(Part, TLen - L) Then Reach(TLen) = True
                    End If
                End If
            Next L
            If Reach(TLen) Then Exit For
        End If
    Next n
    BankFits = Reach(TLen)
End Function

Private Function BankMod97(ByVal Text As String) As Long
    'Rest modulo 97 nach ISO 7064 (Buchstaben A=10 bis Z=35), fuer IBAN und Glaeubiger-ID
    Dim n As Long, Code As Long, Rest As Long
    For n = 1 To Len(Text)
        Code = AscW(Mid$(Text, n, 1))
        If Code <= 57 Then
            Rest = (Rest * 10 + Code - 48) Mod 97
        Else
            Rest = (Rest * 100 + Code - 55) Mod 97
        End If
    Next n
    BankMod97 = Rest
End Function

Private Function BankIbOk(ByVal Token As String) As Boolean
    'IBAN nach ISO 13616: Laenderkennung, Pruefziffer mod 97; deutsche IBAN genau 20 Ziffern nach "DE"
    Dim n As Long, Code As Long
    If Len(Token) < 15 Or Len(Token) > 34 Then Exit Function
    For n = 1 To Len(Token)
        Code = AscW(Mid$(Token, n, 1))
        If n <= 2 Then
            If Code < 65 Or Code > 90 Then Exit Function
        ElseIf n <= 4 Or Left$(Token, 2) = "DE" Then
            If Code < 48 Or Code > 57 Then Exit Function
        ElseIf Not ((Code >= 48 And Code <= 57) Or (Code >= 65 And Code <= 90)) Then
            Exit Function
        End If
    Next n
    If Left$(Token, 2) = "DE" And Len(Token) <> 22 Then Exit Function
    BankIbOk = (BankMod97(Mid$(Token, 5) & Left$(Token, 4)) = 1)
End Function

Private Function BankCiOk(ByVal Token As String) As Boolean
    'SEPA-Glaeubiger-ID: Laenderkennung, Pruefziffer, 3 Zeichen Geschaeftsbereich, nationale Kennung;
    'Pruefziffer mod 97 ohne Geschaeftsbereich; deutsche Glaeubiger-ID 18 Zeichen, Kennung 11 Ziffern
    Dim n As Long, Code As Long
    If Len(Token) < 9 Or Len(Token) > 35 Then Exit Function
    For n = 1 To Len(Token)
        Code = AscW(Mid$(Token, n, 1))
        If n <= 2 Then
            If Code < 65 Or Code > 90 Then Exit Function
        ElseIf n <= 4 Or (n >= 8 And Left$(Token, 2) = "DE") Then
            If Code < 48 Or Code > 57 Then Exit Function
        ElseIf Not ((Code >= 48 And Code <= 57) Or (Code >= 65 And Code <= 90)) Then
            Exit Function
        End If
    Next n
    If Left$(Token, 2) = "DE" And Len(Token) <> 18 Then Exit Function
    BankCiOk = (BankMod97(Mid$(Token, 8) & Left$(Token, 4)) = 1)
End Function

Private Function BankIds(ByVal Text As String, ByVal Kind As Long) As String
    'Kennungen eines Textes als Liste " ID1 ID2 ": Kind 1 = IBAN, 2 = Glaeubiger-ID.
    'Woerter aus Buchstaben und Ziffern; ein MT940-Kennzeichen wie "SVWZ+" am Wortende wird abgetrennt
    Dim n As Long, Code As Long, Start As Long
    Dim Token As String
    Dim IsId As Boolean
    BankIds = " "
    Text = UCase$(Text) & " "
    For n = 1 To Len(Text)
        Code = AscW(Mid$(Text, n, 1))
        If (Code >= 48 And Code <= 57) Or (Code >= 65 And Code <= 90) Then
            If Start = 0 Then Start = n
        ElseIf Start > 0 Then
            If n - Start >= 9 And n - Start <= 40 Then
                Token = Mid$(Text, Start, n - Start)
                If Code = 43 And Len(Token) > 12 Then
                    If InStr("|EREF|KREF|MREF|CRED|DEBT|SVWZ|ABWA|ABWE|", "|" & Right$(Token, 4) & "|") > 0 Then Token = Left$(Token, Len(Token) - 4)
                End If
                If Kind = 1 Then IsId = BankIbOk(Token) Else IsId = BankCiOk(Token)
                If IsId Then
                    If InStr(BankIds, " " & Token & " ") = 0 Then BankIds = BankIds & Token & " "
                End If
            End If
            Start = 0
        End If
    Next n
End Function

Private Function BankHasId(ByVal List As String, ByVal OldText As String, ByVal Cut As Boolean) As Boolean
    'True, wenn der Alttext eine Kennung der Liste (" ID1 ID2 ") enthaelt; bei gekuerztem Alttext
    '(Cut) genuegt ein Kennungsanfang ab 4 Zeichen (Land und Pruefziffer) als letztes Wort
    Dim Ident As Variant
    Dim Tail As String
    Dim n As Long, Code As Long
    If Len(List) < 2 Or Len(OldText) = 0 Then Exit Function
    OldText = UCase$(OldText)
    If Cut Then
        For n = Len(OldText) To 1 Step -1
            Code = AscW(Mid$(OldText, n, 1))
            If Not ((Code >= 48 And Code <= 57) Or (Code >= 65 And Code <= 90)) Then Exit For
        Next n
        Tail = Mid$(OldText, n + 1)
    End If
    For Each Ident In Split(Trim$(List), " ")
        If InStr(1, OldText, Ident, vbBinaryCompare) > 0 Then
            BankHasId = True
            Exit Function
        ElseIf Len(Tail) >= 4 Then
            If Left$(Ident, Len(Tail)) = Tail Then
                BankHasId = True
                Exit Function
            End If
        End If
    Next Ident
End Function

Private Function BankBack(ByVal NewText As String, ByVal OldText As String, ByVal OldCut As Boolean) As Boolean
    'Gegenrichtung zu BankFits: neuer Text aus einer Auswahl der Woerter des Alttexts, Leerzeichen zaehlen
    'nicht (Alttext mit einem Feld mehr, z.B. Kontonummer im CAMT-Format). Der neue Text muss mindestens
    '80 Prozent des Alttexts abdecken; gekuerzte Alttexte prueft nur BankFits
    Dim NewLen As Long, OldLen As Long
    If OldCut Then Exit Function
    NewLen = Len(Replace(NewText, " ", vbNullString))
    OldLen = Len(Replace(OldText, " ", vbNullString))
    If NewLen < 20 Or NewLen * 5 < OldLen * 4 Then Exit Function
    BankBack = BankFits(Split(OldText, " "), NewText, Len(NewText) >= 255)
End Function

Private Sub BankStore(ByVal Plans As Collection, ByRef Target As ADODB.Recordset, ByVal IdBnk As Long, ByVal ManNr As Long, ByVal MitNr As Long, ByVal MinDay As Date, ByVal MaxDay As Date)
    On Error GoTo BankErr
    'Allgemeiner CSV-Import: Abgleich mit dem Bestand des Kontos vor dem Speichern.
    'Der Buchungstext haengt von den konstanten Spalten der ersten 50 Zeilen ab (ScanD) und vom Exportformat.
    'Ein Altumsatz mit gleichem Tag und Betrag gilt als vorhanden bei 1. gleichem Text, 2. Text aus einer
    'Auswahl der Zellen der Zeile oder umgekehrt (Leerzeichen zaehlen nicht), 3. gleicher IBAN, 4. gleicher Glaeubiger-ID
    'des Zahlungsbeteiligten (anderes Exportformat). Jeder Altumsatz deckt hoechstens eine Zeile ab; ausser bei
    'gleichem vollstaendigem Text entscheidet BankAssign. Ist offen, welche Zeile neu ist, wird gespeichert
    Dim Legacy As Object, Used As Object, IdCnt As Object, Extra As Object
    Dim Bucket As Collection
    Dim Plan As Variant, OldRec As Variant, Ident As Variant
    Dim State() As Byte, Weak() As String, ByIban() As String, ByCi() As String, Texts() As String
    Dim PlanIb() As String, PlanCi() As String
    Dim BucketId As String, OwnText As String, OldText As String, NormOld As String, IdText As String, Cands As String
    Dim k As Long, Kind As Long, Skipped As Long, Added As Long, Ambig As Long
    Dim WeakCnt As Long, IbanHit As Long, CiHit As Long
    Dim Cut As Boolean, SameTx As Boolean
    Dim OldDay As Date
    Dim b As Long
    Dim ErrNum As Long, ErrText As String

    Set Legacy = CreateObject("Scripting.Dictionary")
    Set Used = CreateObject("Scripting.Dictionary")
    '1. Bestand des Kontos im Zeitraum einmal lesen, dazu einzelne Tage, unter denen ein frueherer Import
    '   fiktive Valuta gespeichert hat (ParseDate bisherige Lesart, z.B. 29.02.25 -> 25.02.2029)
    Set Extra = CreateObject("Scripting.Dictionary")
    For Each Plan In Plans
        OldDay = DateValue(Plan(4))
        If Year(OldDay) > 1900 And OldDay <> DateValue(Plan(0)) Then
            If OldDay < DateValue(MinDay) Or OldDay > DateValue(MaxDay) Then Extra(CStr(CLng(OldDay))) = OldDay
        End If
    Next Plan
    BankRead IdBnk, ManNr, MinDay, MaxDay, Extra, Legacy, Nothing

    ReDim State(1 To Plans.Count)
    ReDim Weak(1 To Plans.Count)
    ReDim ByIban(1 To Plans.Count)
    ReDim ByCi(1 To Plans.Count)
    ReDim PlanIb(1 To Plans.Count)
    ReDim PlanCi(1 To Plans.Count)
    ReDim Texts(1 To Plans.Count)
    If Legacy.Count > 0 Then
        '2. IBANs und Glaeubiger-IDs je Zeile; eine Kennung in jeder Zeile gehoert zum eigenen Konto
        Set IdCnt = CreateObject("Scripting.Dictionary")
        k = 0
        For Each Plan In Plans
            k = k + 1
            IdText = Join(Plan(3), " ")
            PlanIb(k) = BankIds(IdText, 1)
            PlanCi(k) = BankIds(IdText, 2)
            IdText = Trim$(PlanIb(k) & PlanCi(k))
            If Len(IdText) > 0 Then
                For Each Ident In Split(IdText, " ")
                    If Len(Ident) > 0 Then IdCnt(Ident) = IdCnt(Ident) + 1
                Next Ident
            End If
        Next Plan
        For k = 1 To Plans.Count
            For Kind = 1 To 2
                If Kind = 1 Then Cands = PlanIb(k) Else Cands = PlanCi(k)
                IdText = " "
                If Plans.Count > 1 And Len(Cands) > 1 Then
                    For Each Ident In Split(Trim$(Cands), " ")
                        If IdCnt(Ident) < Plans.Count Then IdText = IdText & Ident & " "
                    Next Ident
                End If
                If Kind = 1 Then PlanIb(k) = IdText Else PlanCi(k) = IdText
            Next Kind
        Next k

        '3. Abgleich: gleicher Text sofort, Kandidaten fuer Teiltext, IBAN und Glaeubiger-ID merken;
        '   bei fiktiver Valuta auch unter dem frueher gespeicherten Tag (Plan(4))
        k = 0
        For Each Plan In Plans
            k = k + 1
            'Text wie in AddRec gespeichert
            OwnText = Replace(CStr(Plan(2)), "?", " ")
            If Len(OwnText) = 0 Then OwnText = "kein Buchungstext" Else OwnText = Left$(OwnText, 255)
            'Speichertag und Text (BankAssign: gleiche Zeilen sind austauschbar)
            Texts(k) = CStr(CLng(DateValue(Plan(0)))) & "|" & OwnText
            OwnText = BankNorm(OwnText)
            For b = 1 To 2
                If b = 1 Then
                    BucketId = BankBucket(Plan(0), Round(CDbl(Plan(1)), 2))
                ElseIf DateValue(Plan(4)) <> DateValue(Plan(0)) Then
                    BucketId = BankBucket(Plan(4), Round(CDbl(Plan(1)), 2))
                Else
                    Exit For
                End If
                If Legacy.Exists(BucketId) Then
                    Set Bucket = Legacy(BucketId)
                    For Each OldRec In Bucket
                        If Not Used.Exists(OldRec(0)) Then
                            OldText = CStr(OldRec(1))
                            NormOld = BankNorm(OldText)
                            Cut = (Len(OldText) >= 255)
                            SameTx = (NormOld = OwnText)
                            If b = 2 Then
                                'Unter dem frueher gespeicherten Tag (fiktive Valuta) nur gleicher, nicht leerer Text;
                                'sonst koennte eine echte Buchung dieses Tages gemeint sein
                                If SameTx And OldText <> "kein Buchungstext" Then Weak(k) = Weak(k) & " " & OldRec(0)
                            ElseIf SameTx And Not Cut Then
                                Used(OldRec(0)) = True
                                State(k) = 1
                                Exit For
                            ElseIf SameTx Or OldText = "kein Buchungstext" Then
                                'gleicher gekuerzter Text oder leerer Alttext: Zuordnung in BankAssign
                                Weak(k) = Weak(k) & " " & OldRec(0)
                            ElseIf BankFits(Plan(3), NormOld, Cut) Then
                                Weak(k) = Weak(k) & " " & OldRec(0)
                            ElseIf BankBack(OwnText, NormOld, Cut) Then
                                Weak(k) = Weak(k) & " " & OldRec(0)
                            ElseIf BankHasId(PlanIb(k), OldText, Cut) Then
                                ByIban(k) = ByIban(k) & " " & OldRec(0)
                            ElseIf BankHasId(PlanCi(k), OldText, Cut) Then
                                ByCi(k) = ByCi(k) & " " & OldRec(0)
                            End If
                        End If
                    Next OldRec
                End If
                If State(k) = 1 Then Exit For
            Next b
        Next Plan
        'Teiltext vor IBAN vor Glaeubiger-ID, je Stufe groesste eins-zu-eins-Zuordnung; Zeilen mit unklarer
        'Zuordnung und verschiedenem Text werden gespeichert (BankAssign)
        WeakCnt = BankAssign(Weak, State, Used, Texts)
        IbanHit = BankAssign(ByIban, State, Used, Texts)
        CiHit = BankAssign(ByCi, State, Used, Texts)
    End If

    '4. Nur neue Umsaetze speichern
    k = 0
    For Each Plan In Plans
        k = k + 1
        If State(k) = 1 Then
            Skipped = Skipped + 1
        Else
            If State(k) = 2 Then Ambig = Ambig + 1
            AddRec Target, IdBnk, ManNr, MitNr, CDate(Plan(0)), CDbl(Plan(1)), CStr(Plan(2)), vbNullString
            Added = Added + 1
            DoEvents
        End If
    Next Plan
    If GlLog = True Then SLogi "BankStore: Added=" & CStr(Added) & "; Skipped=" & CStr(Skipped) & "; Weak=" & CStr(WeakCnt) & "; Iban=" & CStr(IbanHit) & "; Ci=" & CStr(CiHit) & "; Unklar=" & CStr(Ambig) & "; ExtraDays=" & CStr(Extra.Count)
    Exit Sub

BankErr:
    ErrNum = Err.Number
    ErrText = Err.Description
    If GlLog = True Then SLogi "BankStore Err: " & ErrText & " #" & ErrNum
    On Error GoTo 0
    Err.Raise ErrNum, "BankStore", ErrText
End Sub

Private Sub BankRead(ByVal IdBnk As Long, ByVal ManNr As Long, ByVal MinDay As Date, ByVal MaxDay As Date, ByVal Extra As Object, ByVal Legacy As Object, ByVal Known As Object)
    'Bestand des Kontos einmal lesen (BankImport, BankStore): Zeitraum MinDay bis MaxDay und einzelne weitere
    'Tage (Extra: frueher gespeicherte Tage fiktiver Valuta). Mit Known: stabile Kennungen (BKI1) dorthin;
    'alle anderen Umsaetze nach Tag und Betrag in Legacy (je Bucket eine Collection aus Array(ID2, IDKurz))
    On Error GoTo BankErr
    Dim OldRows As ADODB.Recordset
    Dim Bucket As Collection
    Dim Days As Variant
    Dim OldKey As String, BucketId As String
    Dim n As Long
    Dim ErrNum As Long, ErrText As String
    Days = Extra.Items
    For n = 0 To Extra.Count
        If n = 0 Then
            Set OldRows = BankLoad(IdBnk, ManNr, MinDay, MaxDay)
        Else
            Set OldRows = BankLoad(IdBnk, ManNr, Days(n - 1), Days(n - 1))
        End If
        Do While Not OldRows.EOF
            OldKey = vbNullString & OldRows.Fields("Suche").Value
            If Left$(OldKey, 5) = "BKI1:" And Not Known Is Nothing Then
                Known(OldKey) = True
            ElseIf Not IsNull(OldRows.Fields("Datum").Value) And Not IsNull(OldRows.Fields("Betrag").Value) Then
                BucketId = BankBucket(OldRows.Fields("Datum").Value, OldRows.Fields("Betrag").Value)
                If Not Legacy.Exists(BucketId) Then
                    Set Bucket = New Collection
                    Legacy.Add BucketId, Bucket
                End If
                Set Bucket = Legacy(BucketId)
                Bucket.Add Array(CStr(OldRows.Fields("ID2").Value), vbNullString & OldRows.Fields("IDKurz").Value)
            End If
            OldRows.MoveNext
        Loop
        OldRows.Close
        Set OldRows = Nothing
    Next n
    Exit Sub

BankErr:
    ErrNum = Err.Number
    ErrText = Err.Description
    On Error Resume Next
    If Not OldRows Is Nothing Then
        If OldRows.State <> adStateClosed Then OldRows.Close
    End If
    Set OldRows = Nothing
    On Error GoTo 0
    Err.Raise ErrNum, "BankRead", ErrText
End Sub

Private Function BankAssign(ByRef Cands() As String, ByRef State() As Byte, ByVal Used As Object, ByRef Texts() As String) As Long
    'Eins-zu-eins-Zuordnung offener Zeilen (State 0) zu den freien Altumsaetzen ihrer Kandidatenliste (" ID ID"):
    'groesste Zuordnung ueber erweiternde Wege (BankAug) statt Dateireihenfolge; Zeilen mit Partner gelten als
    'vorhanden (State 1). Bleibt eine Zeile ohne Partner, ist offen, welche der ueber wechselnde Wege von ihr
    'erreichbaren Zeilen neu ist. Unterscheiden sie sich in Speichertag oder -text (Texts), werden sie alle gespeichert
    '(State 2, Dublette statt Verlust) und ihre Altumsaetze gesperrt; bei gleichem Text ist die Wahl ohne Folgen,
    'die Zeile ohne Partner bleibt offen fuer die naechste Stufe. Rueckgabe: Zahl der als vorhanden erkannten Zeilen
    Dim Lst() As String, Got() As String
    Dim Mate As Object, Seen As Object, Loose As Object, Block As Object
    Dim Queue As Collection
    Dim OldId As Variant
    Dim k As Long, n As Long, r As Long
    Dim Mixed As Boolean
    ReDim Lst(UBound(Cands))
    ReDim Got(UBound(Cands))
    Set Mate = CreateObject("Scripting.Dictionary")
    Set Seen = CreateObject("Scripting.Dictionary")
    Set Loose = CreateObject("Scripting.Dictionary")
    Set Block = CreateObject("Scripting.Dictionary")
    For k = LBound(Cands) To UBound(Cands)
        If State(k) = 0 And Len(Cands(k)) > 0 Then
            For Each OldId In Split(Trim$(Cands(k)), " ")
                If Not Used.Exists(OldId) Then Lst(k) = Lst(k) & " " & OldId
            Next OldId
            If Len(Lst(k)) > 0 Then
                Seen.RemoveAll
                BankAug k, Lst, Mate, Seen
            End If
        End If
    Next k
    For Each OldId In Mate.Keys
        Got(Mate(OldId)) = OldId
    Next OldId
    For k = LBound(Cands) To UBound(Cands)
        If Len(Lst(k)) > 0 And Len(Got(k)) = 0 Then
            'Zeile ohne Partner: alle ueber wechselnde Wege erreichbaren Zeilen und Altumsaetze sammeln
            Seen.RemoveAll
            Set Queue = New Collection
            Queue.Add k
            Mixed = False
            n = 1
            Do While n <= Queue.Count
                r = Queue(n)
                If Texts(r) <> Texts(k) Then Mixed = True
                For Each OldId In Split(Trim$(Lst(r)), " ")
                    If Not Seen.Exists(OldId) Then
                        Seen.Add OldId, True
                        If Mate.Exists(OldId) Then Queue.Add CLng(Mate(OldId))
                    End If
                Next OldId
                n = n + 1
            Loop
            If Mixed Then
                For n = 1 To Queue.Count
                    Loose(CStr(Queue(n))) = True
                Next n
                For Each OldId In Seen.Keys
                    Block(OldId) = True
                Next OldId
            End If
        End If
    Next k
    For k = LBound(Cands) To UBound(Cands)
        If Len(Lst(k)) > 0 Then
            If Loose.Exists(CStr(k)) Then
                State(k) = 2
            ElseIf Len(Got(k)) > 0 Then
                State(k) = 1
                Used(Got(k)) = True
                BankAssign = BankAssign + 1
            End If
        End If
    Next k
    For Each OldId In Block.Keys
        Used(OldId) = True
    Next OldId
End Function

Private Function BankAug(ByVal k As Long, ByRef Lst() As String, ByVal Mate As Object, ByVal Seen As Object) As Boolean
    'Erweiternder Weg fuer Zeile k (BankAssign): ein freier Altumsatz oder einer, dessen Zeile ausweichen kann
    Dim OldId As Variant
    For Each OldId In Split(Trim$(Lst(k)), " ")
        If Not Seen.Exists(OldId) Then
            Seen.Add OldId, True
            If Not Mate.Exists(OldId) Then
                Mate.Add OldId, k
                BankAug = True
                Exit Function
            ElseIf BankAug(CLng(Mate(OldId)), Lst, Mate, Seen) Then
                Mate(OldId) = k
                BankAug = True
                Exit Function
            End If
        End If
    Next OldId
End Function

Private Function BankProof(ByVal Cells As Variant, ByVal OldText As String) As Boolean
    'True, wenn der Alttext Buchungstag, Valuta oder Saldo nach Buchung der Zeile als ganzes Wort enthaelt (Zellen
    '0, 1 und 8 aus BankCells): Beleg fuer einen Altumsatz, den ein frueherer Import unter dem Tag fiktiver Valuta
    'gespeichert hat (BankImport). Der fruehere allgemeine Import speichert den Saldo, aber kein Datum im Text
    Dim n As Variant
    OldText = " " & OldText & " "
    For Each n In Array(0, 1, 8)
        If Len(Cells(n)) > 0 Then
            If InStr(1, OldText, " " & Cells(n) & " ", vbBinaryCompare) > 0 Then BankProof = True
        End If
    Next n
End Function
