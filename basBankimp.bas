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

' Public Entry Point
Public Sub Imp01(ByVal IdBnk As Long, ByVal ManNr As Long, ByVal MitNr As Long, Optional ByVal PthFl As String = "")
    On Error GoTo LiErr

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
            Exit Sub
        End If
    End If

    ' Check File Existence
    With clFil
        .FilPfa FiNam
    End With
    
    If clFil.FilVor(FiNam) = False Then
        MsgBox "Datei nicht gefunden: " & FiNam, vbExclamation
        Exit Sub
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
        MsgBox "Datenbank untersttzt kein Hinzufgen.", vbCritical
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
        TmpSt = .FilReSt
    End With
    If GlLog = True Then SLogi "Imp01: File read, length=" & Len(TmpSt)

    If TmpSt = vbNullString Then
        If GlLog = True Then SLogi "Imp01: File is empty, exiting"
        GoTo Cleanup
    End If
    
    ' UTF-8 to ANSI Conversion
    If GlLog = True Then SLogi "Imp01: Checking encoding..."
    If IsUTF8(TmpSt) Then
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
    If GlLog = True Then SLogi "Imp01: HdrRow=" & HdrRow

    ' Column Detection
    Dim IdxDa As Integer, IdxAm As Integer
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
        MsgBox "Konnte Datum- oder Betragsspalte nicht automatisch erkennen.", vbExclamation
        GoTo Cleanup
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
    Dim sDat As String, sAmt As String, sTxt As String
    Dim dAmt As Double, dtDat As Date
    Dim j As Integer
    Dim sTmp As String
    Dim sSch As String
    Dim NeuBu As Boolean
    Dim RowUB As Long
    
    Dim StartLoop As Long
    StartLoop = HdrRow + 1

    If GlLog = True Then SLogi "Import: Starting loop from row " & StartLoop & " to " & RowsC.Count
    If GlLog = True Then SLogi "Import: IdxDa=" & IdxDa & " IdxAm=" & IdxAm

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
            AddRec RS121, IdBnk, ManNr, MitNr, dtDat, dAmt, sTxt
             If GlLog = True And i < StartLoop + 10 Then SLogi "DEBUG:   -> ADDED"
        Else
             If GlLog = True And i < StartLoop + 10 Then SLogi "DEBUG:   -> SKIPPED"
        End If

        PrBr1.Value = i
        DoEvents
    Next i

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
    S_BaDop ' Remove Duplicates
    
    DoEvents
    S_List  ' Refresh List?
    
    Exit Sub

LiErr:
    If GlLog = True Then SLogi "Imp01 ERROR: " & Err.Number & " - " & Err.Description
    If GlDbg = True Then MsgBox Err.Description, 48, "Imp01 " & Err.Number
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

Private Sub AddRec(RS As ADODB.Recordset, IdBnk As Long, ManNr As Long, MitNr As Long, dtDat As Date, dAmt As Double, sTxt As String)
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
    RS.Fields("Suche").Value = sSch
    
    RS.Update
End Sub

' ------------------------------------------------------------------------------
' Helpers (Universal Parsing)
' ------------------------------------------------------------------------------

Private Function ParseDate(ByVal s As String) As Date
    On Error GoTo ErrHnd
    ParseDate = 0
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    Dim lD As Long, lM As Long, lY As Long
    Dim Parts() As String
    Dim PartCount As Long

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
                    ParseDate = DateSerial(CInt(lY), CInt(lM), CInt(lD))
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
                    ParseDate = DateSerial(CInt(lY), CInt(lM), CInt(lD))
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
                    ParseDate = DateSerial(CInt(lY), CInt(lM), CInt(lD))
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
            ParseDate = DateSerial(CInt(lY), CInt(lM), CInt(lD))
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
            ParseDate = DateSerial(CInt(lY), CInt(lM), CInt(lD))
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
                If Fields.Count > 0 Then
                    Res.Add ColToArr(Fields)
                    Set Fields = New Collection
                End If
            Else
                Fld = Fld & c
            End If
        End If
    Next i
    If Fields.Count > 0 Or Len(Fld) > 0 Then
        Fields.Add Fld
        Res.Add ColToArr(Fields)
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
    Dim Lines() As String
    Dim i As Long
    Dim Ln As String
    Dim FstCol As String
    Dim OutTxt As String
    Dim SrcPat As String
    Dim PatPos As Long
    Dim SecPos As Long
    Dim PatCnt As Long
    Dim SplitN As Long

    If GlLog = True Then SLogi "FixConc: Starting analysis, Delim=[" & Delim & "]"

    ' Normalize line endings first
    RawTxt = Replace(RawTxt, vbCrLf, vbLf)
    RawTxt = Replace(RawTxt, vbCr, vbLf)
    Lines = Split(RawTxt, vbLf)

    If GlLog = True Then SLogi "FixConc: Total lines before fix: " & (UBound(Lines) + 1)

    ' Find first non-empty line with data to detect the repeating column
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
                    ' Found a potential pattern - verify it appears multiple times
                    PatCnt = CountStr(RawTxt, FstCol & Delim)
                    If GlLog = True Then SLogi "FixConc: Pattern [" & FstCol & "] count=" & PatCnt & " threshold=" & (UBound(Lines) + 5)
                    If PatCnt > UBound(Lines) + 5 Then
                        If GlLog = True Then SLogi "FixConc: Pattern accepted!"
                        Exit For
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

    ' Process each line and split where pattern repeats
    OutTxt = ""
    For i = 0 To UBound(Lines)
        Ln = Lines(i)
        If Len(Ln) > 0 Then
            ' Check if pattern appears more than once in this line
            PatPos = InStr(Ln, SrcPat)
            If PatPos > 0 Then
                SecPos = InStr(PatPos + Len(SrcPat), Ln, SrcPat)
                
                ' NEW: Distance Check
                Dim MinDist As Boolean
                MinDist = False
                If SecPos > 0 Then
                    If SecPos - PatPos > 20 Then MinDist = True
                End If

                If SecPos > 0 And MinDist Then
                    ' Multiple records on this line - split them
                    Ln = SplitRec(Ln, SrcPat)
                    SplitN = SplitN + 1
                    If GlLog = True Then SLogi "FixConc: Split line " & i & " (PatPos=" & PatPos & " SecPos=" & SecPos & ")"
                End If
            End If
        End If
        If Len(OutTxt) > 0 Then
            OutTxt = OutTxt & vbLf
        End If
        OutTxt = OutTxt & Ln
    Next i

    If GlLog = True Then SLogi "FixConc: Split " & SplitN & " lines"

    ' Count lines after fix
    Dim LnAftr() As String
    LnAftr = Split(OutTxt, vbLf)
    If GlLog = True Then SLogi "FixConc: Total lines after fix: " & (UBound(LnAftr) + 1)

    FixConc = OutTxt
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

Private Function CountStr(ByVal Source As String, ByVal Search As String) As Long
    ' Count occurrences of Search in Source
    Dim Cnt As Long
    Dim Pos As Long

    Cnt = 0
    Pos = 1
    Do
        Pos = InStr(Pos, Source, Search)
        If Pos = 0 Then Exit Do
        Cnt = Cnt + 1
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
