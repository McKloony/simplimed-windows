Attribute VB_Name = "basPAD"
Option Explicit

' ============================================================================
' Module: basPAD
' Purpose: PAD and PADNext invoice export functionality
' Migration: S_ReExP migrated from basDatRe
' New: S_ReExN for PADNext XML format export
' ============================================================================

' --- CommonDialog Flags ---
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_PATHMUSTEXIST = &H800

' --- PADNext Namespace ---
Private Const PADX_NS As String = "http://padinfo.de/ns/pad"

' --- Module-level Form and Control References ---
Private FM As Form
Private TxDum As VB.TextBox
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private CoDia As XtremeSuiteControls.CommonDialog

' --- Module-level Recordsets ---
Private RS120 As ADODB.Recordset
Private RS125 As ADODB.Recordset
Private RS126 As ADODB.Recordset
Private RS128 As ADODB.Recordset

' --- Module-level Report Controls ---
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn

' --- Module-level Class Instances ---
Private clFil As clsFile
Private clNet As clsNetz

' ============================================================================
' Function: S_ReExP
' Purpose: Export invoices in legacy PAD2 format for billing offices
' Origin: Migrated from basDatRe
' Returns: Boolean - True if export successful
' ============================================================================
Public Function S_ReExP(Optional ByVal ExpOr As Boolean = False) As Boolean
On Error GoTo LiErr

Dim TmpDa As Date
Dim IdxNr As Long
Dim RowNr As Long
Dim ManNr As Long
Dim AktRe As Long
Dim AnzSe As Long
Dim ZeiGe As Long
Dim ReZei As Long
Dim AnzRe As Long
Dim AnzPo As Long
Dim Lerze As Long
Dim ExOrd As String
Dim RecSt As String
Dim MeStr As String
Dim DaNam As String
Dim FilNa As String
Dim TmpTx As String
Dim TmpKo As String
Dim TmpDi As String
Dim TmPfa As String
Dim TmpGe As String
Dim TmpEi As String
Dim TmpKu As String
Dim HoSum As Double
Dim KoSum As Double
Dim AnSum As Double
Dim ReBet As Double
Dim BeBet As Double
Dim ReStu As Single
Dim LeStu As Single
Dim ReWar As String * 3
Dim StuRe As String * 4
Dim StuLe As String * 4
Dim PatNr As String * 4
Dim ReEmp As String * 50
Dim ReAdr As String * 30
Dim ReStr As String * 30
Dim ReAnr As String * 1
Dim DiaKa As String * 50
Dim DiaMa As String * 50
Dim Thera As String * 50
Dim Behan As String * 40
Dim LeTex As String * 40
Dim KoTex As String * 40
Dim VolNa As String * 11
Dim DatNa As String * 12
Dim ComNa As String * 10
Dim KarNr As String * 12
Dim Gebor As String * 8
Dim ReNum As String * 6
Dim KunNr As String * 6
Dim Gesch As String * 1
Dim VeArt As String * 2
Dim Ziffe As String * 7
Dim Aznah As String * 4
Dim PosKz As String * 1
Dim VoZei As String * 1
Dim Fakto As String * 8
Dim Einze As String * 7
Dim Gesam As String * 7
Dim SumHo As String * 8
Dim SumGe As String * 8
Dim SumKo As String * 8
Dim SuMan As String * 8
Dim PosDa As String * 6
Dim VerNr As String * 3
Dim KatNr As String * 3
Dim VerSt As String * 94
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim ZeTyp As Integer
Dim ZwZei As Integer
Dim AnzAn As Integer
Dim StaWe As Integer
Dim AktWe As Integer
Dim AktPo As Integer
Dim LeZif As Integer
Dim ReSto As Boolean
Dim ZeiUm As Boolean
Dim Umlau As Boolean
Dim Anlog As Boolean
Dim RetWe As Boolean
Dim Frage As Integer
Dim Mld1, Mld2, Tit1 As String
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMain
Set CoDia = FM.comDialo
Set RpCo4 = FM.repCont4
Set RpCo3 = FM.repCont3

Set clFil = New clsFile
Set clNet = New clsNetz
clFil.hwnd = FM.hwnd

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select

Set RpRow = RpSel(0)
If RpRow.GroupRow = False Then
    Set RpCol = RpCls.Find(Rec_IDP)
    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
        ManNr = RpRow.Record(RpCol.ItemIndex).Value
    End If
End If

ZeiGe = 0
AnzAn = 0
ZeiUm = CBool(IniGetVal("System", "PADZei"))
Umlau = CBool(IniGetVal("System", "PADUml"))
Mld1 = "Die Datei existiert bereits, soll diese uberschrieben werden?"
Mld2 = "Im Mandanteneingabedialog wurde keine gultige PAD Kundennummer erfasst."
Tit1 = "PAD Export"
AnzSe = RpSel.Count

TmpKu = S_AdIdx(ManNr, "Postfach")

If TmpKu <> vbNullString Then
    If TmpKu <> "999999" Then
        If TmpKu <> "0" Then
            If IsNumeric(TmpKu) = True Then
                KunNr = Format$(TmpKu, "000000")
            Else
                KunNr = Left$(TmpKu, 6)
            End If
            DaNam = "PV" & KunNr & ".dat"
        Else
            WindowMess Mld2, Dial3, Tit1, FM.hwnd
            Exit Function
        End If
    Else
        WindowMess Mld2, Dial3, Tit1, FM.hwnd
        Exit Function
    End If
Else
    WindowMess Mld2, Dial3, Tit1, FM.hwnd
    Exit Function
End If

If AnzSe > 0 Then
    ReDim GloRe(AnzSe)
Else
    Exit Function
End If

If GlRDP = True Then
    If ExpOr = True Then
        ExOrd = GlDpf & "Export\"
    Else
        If clFil.FilDir(GlIPf) = False Then
            ExOrd = GlDpf & "Import\"
        Else
            ExOrd = GlIPf
        End If
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

ComNa = clNet.NetNam

With CoDia
    .CancelError = True
    .DialogStyle = 1
    .DefaultExt = "*.dat"
    .Filter = "PAD/PVS Dateien (*.dat)|*.dat|Alle Dateien (*.*)|*.*"
    .FilterIndex = 0
    .DialogTitle = "Bitte Name und Ordner der Abrechnungsdatei angeben"
    .FileName = ExOrd & DaNam
    .InitDir = ExOrd
    .ShowSave
    FilNa = .FileName
    If .FileTitle = vbNullString Then
        Set CoDia = Nothing
        Set RpSel = Nothing
        Set RpCls = Nothing
        Set RpCo3 = Nothing
        Set RpCo4 = Nothing
        Set clFil = Nothing
        Exit Function
    End If
End With

If LCase(Right$(FilNa, 4)) <> ".dat" Then
    FilNa = FilNa & ".dat"
End If

With clFil
    If .FilVor(FilNa) = True Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            .DaLoe = FilNa & vbNullChar
            .FilLoe
        Else
            Set CoDia = Nothing
            Set RpSel = Nothing
            Set RpCls = Nothing
            Set RpCo3 = Nothing
            Set RpCo4 = Nothing
            Set clFil = Nothing
            Exit Function
        End If
    End If
End With

VolNa = "PV" & KunNr
DatNa = DaNam

If Not IsNull(FilNa) And Not FilNa = vbNullString Then
    If GlLog = True Then SLogi "basPAD.S_ReExP: Starting PAD export for " & AnzSe & " invoices"

    DBCmEx0 "qrySimPADZu"
    DBCmEx0 "qryAdrMark"
    DoEvents

    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            GloRe(AktPo) = RowNr

            Set RpCol = RpCls.Find(Rec_ID1)
            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Rec_Betrag)
            ReBet = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Rec_Bezahlt)
            BeBet = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Rec_Storniert)
            ReSto = CBool(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Rec_Steuer)
            ReStu = CSng(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Rec_IDP)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                ManNr = 0
            End If
            StuRe = Format$(ReStu, "00") & "00"

            If ReBet > 0 Then
                If ReSto = False Then
                    DBCmEx2 "qrySimReMa3", "@Druck", "@IdxNr", -1, IdxNr
                    DBCmEx2 "qrySimReOP", "@IdOff", "@IdxNr", -1, IdxNr
                    DBCmEx2 "qrySimReTyp", "@ReTyp", "@IdxNr", "A", IdxNr
                    DBCmEx2 "qrySimReSm5", "@IdBez", "@IdxNr", ReBet, IdxNr
                End If
            End If
        End If
    Next RpRow

    Set RS120 = New ADODB.Recordset
    RS120.CursorLocation = adUseClient
    Set RS120 = DBCmRe1("qryAdrIdx", "@IdxNr", ManNr)
    If RS120.RecordCount > 0 Then
        If RS120.Fields("Postfach").Value <> vbNullString Then
            TmpKu = RS120.Fields("Postfach").Value
            If IsNumeric(TmpKu) = True Then
                KunNr = Format$(TmpKu, "000000")
            Else
                KunNr = Left$(TmpKu, 6)
            End If
        Else
            KunNr = "000000"
        End If
    End If
    RS120.Close
    Set RS120 = Nothing

    MeStr = KunNr & "000000" & "000" & VolNa & DatNa & Format$(Now, "ddmmyy") & Format$(Now, "hhmm") & "201605" & ComNa

    Set RS125 = New ADODB.Recordset
    With RS125
        .CursorLocation = adUseClient
        .Source = "qrySimPAD1"
        .ActiveConnection = DB1
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open Options:=adCmdTableDirect
    End With

    AktWe = 1
    AnzRe = RS125.RecordCount
    If AnzRe > 0 Then
        frmStatus.Show
        DoEvents
        frmStatus.Caption = "PAD/PVS Export"
        Set PrBr1 = frmStatus.prbStat1
        Set PrBr2 = frmStatus.prbStat2
        Set TxDum = frmStatus.txtDummy
        PrBr2.Min = 0
        PrBr2.Max = AnzRe

        ' PAD spec (SA000 Feld 10): Waehrungskennung = Konstante EUR
        ReWar = "EUR"
        MeStr = MeStr & ReWar & Space$(61) & vbCrLf
        ZeiGe = ZeiGe + 1

        Do
        ReZei = 0
        DiaMa = vbNullString
        DiaKa = vbNullString
        AktRe = RS125.Fields("IDR").Value
        RecSt = RS125.Fields("RechNr").Value
        ReNum = Format$(Right$(RecSt, 4), "000000")

        MeStr = MeStr & KunNr & ReNum & "100" & Space$(50) & Space$(50) & Space$(13) & vbCrLf
        ZeiGe = ZeiGe + 1

        If Not IsNull(RS125.Fields("R_Anrede").Value) Then
            If LCase(RS125.Fields("R_Anrede").Value) = "firma" Then
                ReAnr = "I"
            Else
                ReAnr = UCase(Left$(RS125.Fields("R_Anrede").Value, 1))
                Select Case ReAnr
                Case "H":
                Case "F":
                Case "I":
                Case "P":
                Case Else: ReAnr = "O"
                End Select
            End If
        Else
            ReAnr = "O"
        End If

        If Not IsNull(RS125.Fields("Geboren").Value) Then
            Gebor = Format$(RS125.Fields("Geboren").Value, "ddmmyyyy")
        Else
            Gebor = "00000000"
        End If

        If ReAnr = "I" Then
            If RS125.Fields("R_Firma1").Value <> vbNullString Then
                If RS125.Fields("R_Name").Value <> vbNullString Then
                    If RS125.Fields("R_Vorname").Value <> vbNullString Then
                        If RS125.Fields("R_Titel").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value, 50), False, Umlau)
                        End If
                    Else
                        If RS125.Fields("R_Titel").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value & Chr$(44), 50), False, Umlau)
                        End If
                    End If
                Else
                    If RS125.Fields("R_Vorname").Value <> vbNullString Then
                        If RS125.Fields("R_Titel").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & Chr$(44) & RS125.Fields("R_Vorname").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & Chr$(44) & RS125.Fields("R_Vorname").Value, 50), False, Umlau)
                        End If
                    Else
                        If RS125.Fields("R_Titel").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44), 50), False, Umlau)
                        End If
                    End If
                End If
            Else
                If RS125.Fields("R_Name").Value <> vbNullString Then
                    If RS125.Fields("R_Vorname").Value <> vbNullString Then
                        If RS125.Fields("R_Titel").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value, 50), False, Umlau)
                        End If
                    Else
                        If RS125.Fields("R_Titel").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value, 50), False, Umlau)
                        End If
                    End If
                End If
            End If
            If RS125.Fields("R_Strasse").Value <> vbNullString Then
                ReStr = SUmw(Left$(RS125.Fields("R_Strasse").Value, 40), False, Umlau)
            Else
                ReStr = vbNullString
            End If
        Else
            If RS125.Fields("R_Name").Value <> vbNullString Then
                If RS125.Fields("R_Vorname").Value <> vbNullString Then
                    If RS125.Fields("R_Titel").Value <> vbNullString Then
                        If RS125.Fields("R_Firma1").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        End If
                    Else
                        If RS125.Fields("R_Firma1").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Vorname").Value, 50), False, Umlau)
                        End If
                    End If
                Else
                    If RS125.Fields("R_Titel").Value <> vbNullString Then
                        If RS125.Fields("R_Firma1").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value & Chr$(44) & RS125.Fields("R_Titel").Value, 50), False, Umlau)
                        End If
                    Else
                        If RS125.Fields("R_Firma1").Value <> vbNullString Then
                            ReEmp = SUmw(Left$(RS125.Fields("R_Firma1").Value & Chr$(44) & RS125.Fields("R_Name").Value, 50), False, Umlau)
                        Else
                            ReEmp = SUmw(Left$(RS125.Fields("R_Name").Value, 50), False, Umlau)
                        End If
                    End If
                End If
                If RS125.Fields("R_Strasse").Value <> vbNullString Then
                    ReStr = SUmw(Left$(RS125.Fields("R_Strasse").Value, 40), False, Umlau)
                Else
                    ReStr = vbNullString
                End If
            ElseIf RS125.Fields("Name").Value <> vbNullString Then
                If RS125.Fields("Vorname").Value <> vbNullString Then
                    If RS125.Fields("Titel").Value <> vbNullString Then
                        ReEmp = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Vorname").Value & Chr$(44) & RS125.Fields("Titel").Value, 50), False, Umlau)
                    Else
                        ReEmp = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Vorname").Value, 50), False, Umlau)
                    End If
                Else
                    If RS125.Fields("Titel").Value <> vbNullString Then
                        ReEmp = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Titel").Value, 50), False, Umlau)
                    Else
                        ReEmp = SUmw(Left$(RS125.Fields("Name").Value, 40), False, Umlau)
                    End If
                End If
                If RS125.Fields("R_Strasse").Value <> vbNullString Then
                    ReStr = SUmw(Left$(RS125.Fields("R_Strasse").Value, 40), False, Umlau)
                Else
                    If RS125.Fields("Strasse").Value <> vbNullString Then
                        ReStr = SUmw(Left$(RS125.Fields("Strasse").Value, 40), False, Umlau)
                    Else
                        ReStr = vbNullString
                    End If
                End If
            End If
        End If

        If RS125.Fields("R_Land").Value <> vbNullString Then
            If RS125.Fields("R_PLZ").Value <> vbNullString Then
                If RS125.Fields("R_Ort").Value <> vbNullString Then
                    ReAdr = RS125.Fields("R_Land").Value & " " & RS125.Fields("R_PLZ").Value & " " & RS125.Fields("R_Ort").Value
                Else
                    ReAdr = RS125.Fields("R_Land").Value & " " & RS125.Fields("R_PLZ").Value
                End If
            Else
                If RS125.Fields("R_Ort").Value <> vbNullString Then
                    ReAdr = RS125.Fields("R_Land").Value & " " & RS125.Fields("R_Ort").Value
                Else
                    ReAdr = RS125.Fields("R_Land").Value
                End If
            End If
        Else
            If RS125.Fields("R_PLZ").Value <> vbNullString Then
                If RS125.Fields("R_Ort").Value <> vbNullString Then
                    ReAdr = RS125.Fields("R_PLZ").Value & " " & RS125.Fields("R_Ort").Value
                Else
                    ReAdr = RS125.Fields("R_PLZ").Value
                End If
            Else
                If RS125.Fields("R_Ort").Value <> vbNullString Then
                    ReAdr = RS125.Fields("R_Ort").Value
                Else
                    ReAdr = vbNullString
                End If
            End If
        End If

        If RS125.Fields("Name").Value <> vbNullString Then
            If RS125.Fields("Vorname").Value <> vbNullString Then
                If RS125.Fields("Titel").Value <> vbNullString Then
                    Behan = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Vorname").Value & Chr$(44) & RS125.Fields("Titel").Value, 50), False, Umlau)
                Else
                    Behan = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Vorname").Value, 50), False, Umlau)
                End If
            Else
                If RS125.Fields("Titel").Value <> vbNullString Then
                    Behan = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Titel").Value, 50), False, Umlau)
                Else
                    Behan = SUmw(Left$(RS125.Fields("Name").Value, 40), False, Umlau)
                End If
            End If
        ElseIf RS125.Fields("Name").Value <> vbNullString Then
            If RS125.Fields("Vorname").Value <> vbNullString Then
                If RS125.Fields("Titel").Value <> vbNullString Then
                    Behan = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Vorname").Value & Chr$(44) & RS125.Fields("Titel").Value, 50), False, Umlau)
                Else
                    Behan = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Vorname").Value, 50), False, Umlau)
                End If
            Else
                If RS125.Fields("Titel").Value <> vbNullString Then
                    Behan = SUmw(Left$(RS125.Fields("Name").Value & Chr$(44) & RS125.Fields("Titel").Value, 50), False, Umlau)
                Else
                    Behan = SUmw(Left$(RS125.Fields("Name").Value, 40), False, Umlau)
                End If
            End If

        End If

        If RS125.Fields("Geschlecht").Value = "M" Then
            Gesch = "m"
        Else
            Gesch = "w"
        End If
        If IsNull(RS125.Fields("Abteilung").Value) Then
            VeArt = "01"
        Else
            If RS125.Fields("Abteilung").Value <> vbNullString Then
                If IsNumeric(RS125.Fields("Abteilung").Value) Then
                    VeArt = Format$(RS125.Fields("Abteilung").Value, "00")
                Else
                    VeArt = "01"
                End If
            Else
                VeArt = "01"
            End If
        End If
        If RS125.Fields("Kartennummer").Value <> vbNullString Then
            KarNr = RS125.Fields("Kartennummer").Value
        Else
            KarNr = Space$(12)
        End If
        If RS125.Fields("PatNr").Value <> vbNullString Then
            PatNr = Format$(RS125.Fields("PatNr").Value, "0000")
        Else
            PatNr = "0000"
        End If

        Set RS126 = New ADODB.Recordset
        RS126.CursorLocation = adUseClient
        Set RS126 = DBCmRe1("qrySimPAD2", "@IdxNr", AktRe)

        AktPo = 1
        AnzPo = RS126.RecordCount

        If AnzPo > 0 Then
            PrBr1.Min = 0
            PrBr1.Max = AnzPo
            PrBr1.Value = 0

            MeStr = MeStr & KunNr & ReNum & "200" & Space$(4) & VeArt & "N" & ReAnr & ReEmp & ReStr & "0000000" & KarNr & PatNr & Space$(2) & vbCrLf
            ZeiGe = ZeiGe + 1
            ReZei = ReZei + 1

            If LCase(ReEmp) <> LCase(Behan) Then
                MeStr = MeStr & KunNr & ReNum & "300" & ReAdr & Behan & Gebor & "00000000000000N000 000" & Gesch & ReWar & StuRe & Space$(5) & vbCrLf
            Else
                MeStr = MeStr & KunNr & ReNum & "300" & ReAdr & Space$(40) & Gebor & "00000000000000N000 000" & Gesch & ReWar & StuRe & Space$(5) & vbCrLf
            End If
            ZeiGe = ZeiGe + 1
            ReZei = ReZei + 1

            If RS125.Fields("ID3").Value <> vbNullString Then
                KatNr = Format$(RS125.Fields("ID3").Value, "000")
            Else
                KatNr = "000"
            End If
            If RS125.Fields("Tarif").Value <> vbNullString Then
                VerNr = Format$(RS125.Fields("Tarif").Value, "000")
            Else
                VerNr = "000"
            End If
            If RS125.Fields("Versicherung").Value <> vbNullString Then
                If VerNr = "001" Then
                    If RS125.Fields("Versicherer").Value <> vbNullString Then
                        VerSt = " Liquidiert wurde nach " & RS125.Fields("Versicherer").Value
                    Else
                        VerSt = vbNullString
                    End If
                Else
                    VerSt = " Liquidiert wurde nach " & RS125.Fields("Versicherung").Value
                End If
            Else
                If RS125.Fields("Versicherer").Value <> vbNullString Then
                    VerSt = " Liquidiert wurde nach " & RS125.Fields("Versicherer").Value
                Else
                    VerSt = vbNullString
                End If
            End If

            If GlGbK = True Then
                VerSt = SUmw(VerSt, False, Umlau)
                MeStr = MeStr & KunNr & ReNum & "500" & KatNr & VerNr & VerSt & Space$(13) & vbCrLf
            Else
                MeStr = MeStr & KunNr & ReNum & "500" & Space$(113) & vbCrLf
            End If
            ZeiGe = ZeiGe + 1
            ReZei = ReZei + 1

            If RS125.Fields("Diagnose").Value <> vbNullString Then
                If Len(RS125.Fields("Diagnose").Value) > 0 Then
                    TmpDi = Trim$(RS125.Fields("Diagnose").Value)
                    TmpDi = SNaFi(TmpDi, False, Umlau, True, False)
                    StaWe = 1
                    Do
                    DiaMa = Mid$(TmpDi, StaWe, 50)
                    Lerze = InStrRev(DiaMa, Chr$(32), -1, 1)
                    If Lerze = 1 Then
                        DiaMa = Trim$(Mid$(TmpDi, StaWe, 50))
                        StaWe = StaWe + 50
                    ElseIf Lerze > 1 Then
                        DiaMa = Trim$(Mid$(TmpDi, StaWe, Lerze - 1))
                        StaWe = StaWe + (Lerze - 1)
                    Else
                        StaWe = StaWe + 50
                    End If
                    MeStr = MeStr & KunNr & ReNum & "600" & DiaMa & Space$(63) & vbCrLf
                    ZeiGe = ZeiGe + 1
                    ReZei = ReZei + 1
                    Loop Until StaWe >= Len(TmpDi)
                End If
            End If

            Set RS128 = New ADODB.Recordset
            RS128.CursorLocation = adUseClient
            If GlDSo = True Then
                Set RS128 = DBCmRe1("qrySimAbDiag", "@IdxNr", AktRe)
            Else
                Set RS128 = DBCmRe1("qrySimAbDia5", "@IdxNr", AktRe)
            End If
            If RS128.RecordCount > 0 Then
                Do
                If RS128.Fields("IDKurz").Value <> vbNullString Then
                    If RS128.Fields("Datum").Value <> vbNullString Then
                        If GldId = True Then
                            If GlICD = True Then
                                TmpDi = RS128.Fields("Datum").Value & ": " & SUmw(RS128.Fields("IDKurz").Value, False, Umlau) & " (" & Trim$(RS128.Fields("GOID").Value) & ")"
                            Else
                                TmpDi = RS128.Fields("Datum").Value & ": " & SUmw(RS128.Fields("IDKurz").Value, False, Umlau)
                            End If
                        Else
                            If GlICD = True Then
                                TmpDi = SUmw(RS128.Fields("IDKurz").Value, False, Umlau) & " (" & RS128.Fields("GOID").Value & ")"
                            Else
                                TmpDi = SUmw(RS128.Fields("IDKurz").Value, False, Umlau)
                            End If
                        End If
                    Else
                        If GlICD = True Then
                            TmpDi = SUmw(RS128.Fields("IDKurz").Value, False, Umlau) & " (" & RS128.Fields("GOID").Value & ")"
                        Else
                            TmpDi = SUmw(RS128.Fields("IDKurz").Value, False, Umlau)
                        End If
                    End If

                    StaWe = 1
                    Do
                    DiaKa = Mid$(TmpDi, StaWe, 50)
                    Lerze = InStrRev(DiaKa, Chr$(32), -1, 1)
                    If Lerze = 1 Then
                        DiaKa = Trim$(Mid$(TmpDi, StaWe, 50))
                        StaWe = StaWe + 50
                    ElseIf Lerze > 1 Then
                        DiaKa = Trim$(Mid$(TmpDi, StaWe, Lerze - 1))
                        StaWe = StaWe + (Lerze - 1)
                    Else
                        StaWe = StaWe + 50
                    End If
                    MeStr = MeStr & KunNr & ReNum & "600" & DiaKa & Space$(63) & vbCrLf
                    ZeiGe = ZeiGe + 1
                    ReZei = ReZei + 1
                    Loop Until StaWe >= Len(TmpDi)

                End If
                RS128.MoveNext
                Loop Until RS128.EOF
            End If
            RS128.Close
            Set RS128 = Nothing

            HoSum = 0
            KoSum = 0
            AnSum = 0

            Do
            If RS126.Fields("Text").Value <> vbNullString Then
                If IsDate(RS126.Fields("Datum").Value) = True Then
                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then
                        If Left$(Round(RS126.Fields("GesBetrag").Value, 2), 1) = "-" Then
                            VoZei = "-"
                        Else
                            VoZei = Chr$(32)
                        End If
                    Else
                        VoZei = Chr$(32)
                    End If

                    ZeTyp = RS126.Fields("ID1").Value
                    LeStu = RS126.Fields("Steuer").Value
                    TmpGe = Replace(Format$(Round(RS126.Fields("GesBetrag").Value, 2), "#####,##00000.00"), ".", vbNullString, 1, , 1)
                    TmpEi = Replace(Format$(RS126.Fields("Betrag").Value, "#####,##00000.00"), ".", vbNullString, 1, , 1)

                    Select Case ZeTyp
                    Case 1: PosKz = "T"
                    Case 2: PosKz = Chr$(32)
                            Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                    Case 3: PosKz = Chr$(32)
                            Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                    Case 4: Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                            If GlPvM = True Then
                                PosKz = Chr$(32)
                            Else
                                PosKz = "M"
                            End If
                    Case 5:
                            If GlPvM = True Then
                                PosKz = Chr$(32)
                            Else
                                PosKz = "B"
                            End If
                    Case 6: PosKz = Chr$(32)
                            Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                            If RS126.Fields("AnzBetrag").Value <> vbNullString Then
                                AnSum = AnSum + Round(RS126.Fields("AnzBetrag").Value, 2)
                            End If
                    Case 7: PosKz = Chr$(32)
                            Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                    Case 8: PosKz = "H"
                            Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                    Case 9: PosKz = "H"
                            Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                    Case Else:
                            PosKz = Chr$(32)
                            Einze = Left$(TmpEi, 5) & Right$(TmpEi, 2)
                            Gesam = Left$(TmpGe, 5) & Right$(TmpGe, 2)
                    End Select

                    If RS126.Fields("GONr").Value <> vbNullString Then
                        LeZif = Len(RS126.Fields("GONr").Value)
                        If GlAnl = True Then
                            If Len(RS126.Fields("Analog").Value) > 0 Then
                                Anlog = CBool(RS126.Fields("Analog").Value)
                            Else
                                Anlog = 0
                            End If
                        Else
                            Anlog = 0
                        End If
                        If Anlog = True Then
                            If LeZif >= 7 Then
                                Ziffe = Left$(RS126.Fields("GONr").Value, 7)
                            ElseIf LeZif <= 4 Then
                                Ziffe = Space$(4 - LeZif) & RS126.Fields("GONr").Value & "(a)"
                            ElseIf LeZif <= 6 Then
                                Ziffe = Space$(6 - LeZif) & RS126.Fields("GONr").Value & "a"
                            Else
                                Ziffe = RS126.Fields("GONr").Value
                            End If
                        Else
                            If LeZif >= 7 Then
                                Ziffe = Left$(RS126.Fields("GONr").Value, 7)
                            Else
                                Ziffe = Space$(7 - LeZif) & RS126.Fields("GONr").Value
                            End If
                        End If
                    Else
                        Ziffe = Space$(7)
                    End If

                    TmPfa = Format$(Round(RS126.Fields("Multi").Value, 2), "##,######00.000000")
                    Fakto = Left$(TmPfa, 2) & Right$(TmPfa, 6)

                    StaWe = 1
                    ZwZei = 1
                    Do
                    If Len(RS126.Fields("Text").Value) > 0 Then
                        TmpTx = Trim$(RS126.Fields("Text").Value)
                        TmpTx = Replace(TmpTx, vbCrLf, Chr$(32), 1, , 1)
                        TmpTx = SUmw(TmpTx, False, Umlau)
                        If Len(RS126.Fields("Kommentar").Value) > 0 Then
                            TmpKo = Trim$(RS126.Fields("Kommentar").Value)
                            TmpKo = Replace(TmpKo, vbCrLf, Chr$(32), 1, , 1)
                            TmpKo = SUmw(TmpKo, False, Umlau)
                        Else
                            TmpKo = vbNullString
                        End If

                        LeTex = Mid$(TmpTx & Chr$(32) & TmpKo, StaWe, 40)
                        Lerze = InStrRev(LeTex, Chr$(32), -1, 1)
                        If Lerze = 1 Then
                            LeTex = Trim$(Mid$(TmpTx & Chr$(32) & TmpKo, StaWe, 40))
                            StaWe = StaWe + 40
                        ElseIf Lerze > 1 Then
                            LeTex = Trim$(Mid$(TmpTx & Chr$(32) & TmpKo, StaWe, Lerze - 1))
                            StaWe = StaWe + (Lerze - 1)
                        Else
                            StaWe = StaWe + 40
                        End If

                        TmpDa = RS126.Fields("Datum").Value
                        PosDa = Format$(DatePart("d", TmpDa, vbMonday), "00") & Format$(DatePart("m", TmpDa, vbMonday), "00") & Right$(DatePart("yyyy", TmpDa, vbMonday), 2)

                        If IsNull(RS126.Fields("Anzahl").Value) Then
                            Aznah = "0000"
                        Else
                            Aznah = Format$(RS126.Fields("Anzahl").Value, "0000")
                        End If

                        Select Case ZeTyp
                        Case 1:
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                MeStr = MeStr & Space$(7) & "0000" & "00000000" & VoZei & "0000000" & VoZei & "0000000" & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                        Case 2:
                            If ZwZei = 1 Then
                                ReZei = ReZei + 1
                                If PosKz = Chr$(32) Then
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Ziffe & Aznah & Fakto & VoZei & Gesam & VoZei & Einze & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then HoSum = HoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                Else
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Space$(7) & Aznah & Fakto & " 0000000 0000000 0000000" & VoZei & Gesam & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then KoSum = KoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                End If
                            Else
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                            End If
                        Case 3:
                            If ZwZei = 1 Then
                                ReZei = ReZei + 1
                                If PosKz = Chr$(32) Then
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Ziffe & Aznah & Fakto & VoZei & Gesam & VoZei & Einze & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then HoSum = HoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                Else
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Space$(7) & Aznah & Fakto & " 0000000 0000000 0000000" & VoZei & Gesam & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then KoSum = KoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                End If
                            Else
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                            End If
                        Case 4:
                                If ZwZei = 1 Then
                                    ReZei = ReZei + 1
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    If GlPvM = True Then
                                        MeStr = MeStr & Ziffe & Aznah & Fakto & VoZei & Gesam & VoZei & Einze & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    Else
                                        MeStr = MeStr & Space$(7) & Aznah & Fakto & VoZei & "0000000" & VoZei & "0000000 0000000 " & Gesam & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    End If
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then HoSum = HoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                Else
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                End If
                        Case 5:
                            If ZwZei = 1 Then
                                ReZei = ReZei + 1
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                MeStr = MeStr & Space$(7) & "0000" & "00000000" & VoZei & "0000000" & VoZei & "0000000" & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                            Else
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                            End If
                        Case 6:
                            AnzAn = AnzAn + 1
                        Case 7:
                            If ZwZei = 1 Then
                                ReZei = ReZei + 1
                                If PosKz = Chr$(32) Then
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Ziffe & Aznah & Fakto & VoZei & Gesam & VoZei & Einze & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then HoSum = HoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                Else
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Space$(7) & Aznah & Fakto & " 0000000 0000000 0000000" & VoZei & Gesam & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then KoSum = KoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                End If
                            Else
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                            End If
                        Case 8:
                            If ZwZei = 1 Then
                                ReZei = ReZei + 1
                                If PosKz = Chr$(32) Then
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Space$(7) & Aznah & Fakto & VoZei & Gesam & VoZei & Einze & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then HoSum = HoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                Else
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Space$(7) & Aznah & Fakto & " 0000000 0000000 0000000" & VoZei & Gesam & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then KoSum = KoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                End If
                            Else
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                            End If
                        Case 9:
                            StuLe = Format$(LeStu, "00") & "00"
                            If ZwZei = 1 Then
                                ReZei = ReZei + 1
                                If PosKz = Chr$(32) Then
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Ziffe & Aznah & Fakto & VoZei & Gesam & VoZei & Einze & " 0000000 0000000" & LeTex & Space(9) & StuLe & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then HoSum = HoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                Else
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Space$(7) & Aznah & Fakto & " 0000000 0000000 0000000" & VoZei & Gesam & LeTex & Space(9) & StuLe & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then KoSum = KoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                End If
                            Else
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & StuLe & Space(2) & vbCrLf
                            End If
                        Case Else:
                            If ZwZei = 1 Then
                                ReZei = ReZei + 1
                                If PosKz = Chr$(32) Then
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Ziffe & Aznah & Fakto & VoZei & Gesam & VoZei & Einze & " 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then HoSum = HoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                Else
                                    MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz
                                    MeStr = MeStr & Space$(7) & Aznah & Fakto & " 0000000 0000000 0000000" & VoZei & Gesam & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                                    If Round(RS126.Fields("GesBetrag").Value, 2) <> vbNullString Then KoSum = KoSum + Round(RS126.Fields("GesBetrag").Value, 2)
                                End If
                            Else
                                MeStr = MeStr & KunNr & ReNum & "700" & PosDa & PosKz & Space$(7) & "000000000000" & " 0000000 0000000 0000000 0000000" & LeTex & Space(9) & "0000" & Space(2) & vbCrLf
                            End If
                        End Select
                        ZeiGe = ZeiGe + 1
                        ZwZei = ZwZei + 1
                    Else
                        TmpTx = vbNullString
                        Exit Do
                    End If

                    Loop Until StaWe >= Len(TmpTx & TmpKo)
                End If
            End If

            DoEvents
            If AktPo < AnzPo Then PrBr1.Value = AktPo
            AktPo = AktPo + 1
            RS126.MoveNext
            Loop Until RS126.EOF

            If HoSum < 0 Then
                SumHo = Mid$(Format$(HoSum, "000000.00"), 2, 6) & Right$(Format$(HoSum, "000000.00"), 2)
                SumGe = Mid$(Format$(HoSum - AnSum, "000000.00"), 2, 6) & Right$(Format$(HoSum - AnSum, "000000.00"), 2)
                VoZei = "-"
            Else
                SumHo = Left$(Format$(HoSum, "000000.00"), 6) & Right$(Format$(HoSum, "000000.00"), 2)
                VoZei = Chr$(32)
            End If

            If KoSum < 0 Then
                VoZei = "-"
                SumKo = Mid$(Format$(KoSum, "000000.00"), 2, 6) & Right$(Format$(KoSum, "000000.00"), 2)
            Else
                SumKo = Left$(Format$(KoSum, "000000.00"), 6) & Right$(Format$(KoSum, "000000.00"), 2)
                VoZei = Chr$(32)
            End If

            If AnSum < 0 Then
                VoZei = "-"
                SuMan = Mid$(Format$(AnSum, "000000.00"), 2, 6) & Right$(Format$(AnSum, "000000.00"), 2)
            Else
                SuMan = Left$(Format$(AnSum, "000000.00"), 6) & Right$(Format$(AnSum, "000000.00"), 2)
                VoZei = Chr$(32)
            End If

            ReZei = ReZei + 1
            MeStr = MeStr & KunNr & ReNum & "900" & VoZei & SumHo & " 00000000" & VoZei & SumKo & "-" & SuMan & " 00000000" & Format$(ReZei, "00000") & Space$(63) & vbCrLf
            ZeiGe = ZeiGe + 1
        End If
        RS126.Close
        Set RS126 = Nothing

        DoEvents
        If AktWe < AnzRe Then PrBr2.Value = AktWe
        AktWe = AktWe + 1
        RS125.MoveNext
        Loop Until RS125.EOF

        ZeiGe = ZeiGe + 1
        MeStr = MeStr & KunNr & "000000" & "990" & Format$((ZeiGe) - AnzAn, "000000") & Format$(AnzRe, "000000") & Space$(101) & vbCrLf
        Unload frmStatus
        Set frmStatus = Nothing
        DoEvents
    End If
    RS125.Close
    Set RS125 = Nothing

    DBCmEx0 "qrySimPADZu"
    DoEvents

    MeStr = Replace(MeStr, Chr$(0), Chr$(32), 1, , 1)
    If ZeiUm = False Then
        MeStr = Replace(MeStr, vbCrLf, vbNullString, 1, , 1)
    End If

    If AnzRe > 0 Then
        If Umlau = True Then
            With clFil
                .FilPfa FilNa
                .StrDa = MeStr
                RetWe = .FilWrSt
                .StrDa = vbNullString
            End With
        Else
            SCaKo MeStr, "ibm850", FilNa
        End If
    End If
End If

DoEvents
GlNeK = GlKoX
With GlNeK
    .PatNr = ManNr
    .IdxNr = 0
    .EiDat = Format$(Date, "dd.mm.yyyy")
    .EiZei = TimeValue(Now)
    .EiTyp = 104
    .TeStr = "Exportiert - PAD Datei - (" & AnzSe & " Positionen)"
    .ZiStr = Format$(Now, "hh:mm") & " Uhr"
    .NeuEi = True
    .KeiAk = True
    .Mitar = GlMiA(GlSmI, 2)
End With
S_Prot

Set CoDia = Nothing
Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Set clFil = Nothing
Set clNet = Nothing

TeTit = "Rechnungsexport"
TeMai = AnzRe & " Rechnungen wurden exportiert"
TeInh = "Die markieren Rechnungen wurden unter der Datei: " & DaNam & " im folgenden Ordner abgelegt."
TeFus = "Die Abrechnungsdatei steht im folgenden Ordner zum Upload bereit:" & vbCrLf & FilNa

If AnzRe > 0 Then
    S_ReExP = True
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    If GlLog = True Then SLogi "basPAD.S_ReExP: PAD export completed: " & AnzRe & " invoices"
End If

Exit Function

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_ReExP " & Err.Number
If GlLog = True Then SLogi "basPAD.S_ReExP: Error " & Err.Number & " - " & Err.Description
Exit Function

End Function

' ============================================================================
' Function: S_ReExN
' Purpose: Export invoices in PADNext XML format (version 2.12)
' Returns: Boolean - True if export successful
' ============================================================================
Public Function S_ReExN(Optional ByVal ExpOr As Boolean = False) As Boolean
On Error GoTo LiErr

Dim xmlDoc As Object
Dim xmlRoot As Object
Dim xmlRech As Object
Dim xmlFall As Object
Dim xmlPos As Object
Dim xmlElem As Object

Dim FilNa As String
Dim DaNam As String
Dim AnzRe As Long
Dim AnzSe As Long
Dim AktRe As Long
Dim AktWe As Long
Dim ReNum As String
Dim RecSt As String
Dim IdxNr As Long
Dim RowNr As Long
Dim AktPo As Long
Dim ManNr As Long
Dim ReBet As Currency
Dim BeBet As Currency
Dim ReSto As Boolean
Dim ReStu As Single
Dim StuRe As String
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim Gebor As String
Dim ReAnr As String
Dim ReEmp As String
Dim DiaTe As String
Dim DiaKa As String
Dim Anzah As String
Dim LeDat As String
Dim LeZif As String
Dim PosNr As Integer
Dim Mld1, Mld2, Tit1 As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

If GlLog = True Then SLogi "basPAD.S_ReExN: START - PADNext XML export"

S_ReExN = False
Set FM = frmMain
Set CoDia = FM.comDialo
Set RpCo4 = FM.repCont4
Set RpCo3 = FM.repCont3

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select

AnzSe = RpSel.Count
AktPo = 0

If AnzSe <= 0 Then
    If GlLog = True Then SLogi "basPAD.S_ReExN: ERROR - No rows selected"
    MsgBox "Keine Rechnungen ausgewaehlt.", vbExclamation, "PADNext Export"
    S_ReExN = False
    Exit Function
End If

' Get ManNr from first selected row for validation
Set RpRow = RpSel(0)
If RpRow.GroupRow = False Then
    Set RpCol = RpCls.Find(Rec_IDP)
    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
        ManNr = RpRow.Record(RpCol.ItemIndex).Value
    End If
End If

' Validate Kundennummer (required for PADNext rechnungsersteller)
Dim TmpKuV As String
TmpKuV = S_AdIdx(ManNr, "Postfach")

If TmpKuV = vbNullString Then
    If GlLog = True Then SLogi "basPAD.S_ReExN: ERROR - No Kundennummer (Postfach) found"
    MsgBox "Im Mandanteneingabedialog wurde keine gueltige PAD Kundennummer erfasst.", vbExclamation, "PADNext Export"
    S_ReExN = False
    Exit Function
End If

If TmpKuV = "0" Or TmpKuV = "999999" Then
    If GlLog = True Then SLogi "basPAD.S_ReExN: ERROR - Invalid Kundennummer: " & TmpKuV
    MsgBox "Im Mandanteneingabedialog wurde keine gueltige PAD Kundennummer erfasst.", vbExclamation, "PADNext Export"
    S_ReExN = False
    Exit Function
End If

If GlLog = True Then SLogi "basPAD.S_ReExN: Kundennummer validated: " & TmpKuV

ReDim GloRe(AnzSe)

DaNam = "PN" & Format$(Now, "YYYYMMDD_HHMMSS") & ".xml"

Dim ExOrd As String
ExOrd = IniGetVal("Standardordner", "PAD-Export")

If Len(ExOrd) = 0 Or InStr(ExOrd, ":") = 0 Then
    ExOrd = GlExO
    If Len(ExOrd) = 0 Or InStr(ExOrd, ":") = 0 Then
        ExOrd = GlDpf & "Export\"
    End If
End If

If Right$(ExOrd, 1) <> "\" Then
    ExOrd = ExOrd & "\"
End If

With CoDia
    .CancelError = True
    .DefaultExt = "*.xml"
    .DialogTitle = "Bitte Name und Ordner der Abrechnungsdatei angeben"
    .Filter = "XML-Dateien (*.xml)|*.xml|Alle Dateien (*.*)|*.*"
    .FilterIndex = 0
    .FileName = ExOrd & DaNam
    .InitDir = ExOrd
    .ShowSave
    FilNa = .FileName
    If .FileTitle = vbNullString Then
        Set CoDia = Nothing
        Set RpSel = Nothing
        Set RpCls = Nothing
        Set RpCo3 = Nothing
        Set RpCo4 = Nothing
        Exit Function
    End If
End With

If LCase(Right$(FilNa, 4)) <> ".xml" Then
    FilNa = FilNa & ".xml"
End If
IniSetVal "Standardordner", "PAD-Export", Left$(FilNa, InStrRev(FilNa, "\"))

If InStr(FilNa, ":") = 0 Then
    If GlLog = True Then SLogi "basPAD.S_ReExN: ERROR - Invalid file path: " & FilNa
    MsgBox "Fehler: Ungultiger Dateipfad. Bitte einen vollstandigen Pfad angeben.", vbExclamation, "PADNext Export"
    S_ReExN = False
    Exit Function
End If

If GlLog = True Then SLogi "basPAD.S_ReExN: Export file = " & FilNa

Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.validateOnParse = False
xmlDoc.resolveExternals = False
xmlDoc.appendChild xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""iso-8859-15""")

Set xmlRoot = PadXmlEl(xmlDoc, "rechnungen")
xmlRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
xmlRoot.setAttribute "xsi:schemaLocation", "http://padinfo.de/ns/pad http://padinfo.de/ns/pad/padx_adl_v2.12.xsd"
xmlDoc.appendChild xmlRoot

Set xmlElem = PadXmlEl(xmlDoc, "nachrichtentyp")
xmlElem.setAttribute "version", "02.12"
xmlElem.Text = "ADL"
xmlRoot.appendChild xmlElem

' Add rechnungsersteller element (required by PADNext)
' Note: ManNr was already retrieved and validated at function start
Set xmlElem = S_ReExNAddRErs(xmlDoc, ManNr)
If Not xmlElem Is Nothing Then
    xmlRoot.appendChild xmlElem
End If

' Add leistungserbringer element (required by PADNext)
Set xmlElem = S_ReExNAddLErb(xmlDoc, ManNr)
If Not xmlElem Is Nothing Then
    xmlRoot.appendChild xmlElem
End If

DBCmEx0 "qrySimPADZu"
DBCmEx0 "qryAdrMark"
DoEvents

For Each RpRow In RpSel
    If RpRow.GroupRow = False Then
        RowNr = RpRow.Index
        GloRe(AktPo) = RowNr

        Set RpCol = RpCls.Find(Rec_ID1)
        IdxNr = RpRow.Record(RpCol.ItemIndex).Value

        Set RpCol = RpCls.Find(Rec_Betrag)
        ReBet = RpRow.Record(RpCol.ItemIndex).Value

        Set RpCol = RpCls.Find(Rec_Bezahlt)
        BeBet = RpRow.Record(RpCol.ItemIndex).Value

        Set RpCol = RpCls.Find(Rec_Storniert)
        ReSto = CBool(RpRow.Record(RpCol.ItemIndex).Value)

        Set RpCol = RpCls.Find(Rec_Steuer)
        ReStu = CSng(RpRow.Record(RpCol.ItemIndex).Value)

        Set RpCol = RpCls.Find(Rec_IDP)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            ManNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            ManNr = 0
        End If

        If ReBet > 0 Then
            If ReSto = False Then
                DBCmEx2 "qrySimReMa3", "@Druck", "@IdxNr", -1, IdxNr
                DBCmEx2 "qrySimReOP", "@IdOff", "@IdxNr", -1, IdxNr
                DBCmEx2 "qrySimReTyp", "@ReTyp", "@IdxNr", "A", IdxNr
                DBCmEx2 "qrySimReSm5", "@IdBez", "@IdxNr", ReBet, IdxNr
            End If
        End If
    End If
Next RpRow

Set RS125 = New ADODB.Recordset
With RS125
    .CursorLocation = adUseClient
    .Source = "qrySimPAD1"
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdTableDirect
End With

AnzRe = RS125.RecordCount
If GlLog = True Then SLogi "basPAD.S_ReExN: Exporting " & AnzRe & " invoices"
xmlRoot.setAttribute "anzahl", CStr(AnzRe)

If AnzRe > 0 Then
    frmStatus.Show vbModeless
    frmStatus.Caption = "PADNext XML Export"
    DoEvents
    Set PrBr1 = frmStatus.prbStat1
    Set PrBr2 = frmStatus.prbStat2

    PrBr1.Min = 0
    PrBr1.Max = AnzRe
    PrBr1.Value = 0
    PrBr1.Visible = True

    PrBr2.Min = 0
    PrBr2.Max = 100
    PrBr2.Value = 0
    PrBr2.Visible = True

    frmStatus.lblLab01.Caption = "Bereite Export vor..."

    AktWe = 1
    PosNr = 0

    Do While Not RS125.EOF
        PrBr1.Value = AktWe
        frmStatus.lblLab01.Caption = "Rechnung " & AktWe & " von " & AnzRe
        DoEvents

        PrBr2.Value = 0

        AktRe = RS125.Fields("IDR").Value
        RecSt = RS125.Fields("RechNr").Value
        ReNum = Format$(Right$(RecSt, 4), "000000")

        Set xmlRech = PadXmlEl(xmlDoc, "rechnung")
        xmlRech.setAttribute "id", RecSt
        xmlRech.setAttribute "abrechnungsform", "22222222"
        xmlRech.setAttribute "druckkennzeichen", "1"
        xmlRech.setAttribute "eabgabe", "0"
        xmlRoot.appendChild xmlRech

        Set xmlElem = S_ReExNAddEmpf(xmlDoc, RS125)
        xmlRech.appendChild xmlElem

        Set xmlFall = S_ReExNAddFall(xmlDoc, RS125, AktRe)
        xmlRech.appendChild xmlFall

        RS125.MoveNext
        AktWe = AktWe + 1
    Loop

    RS125.Close
    Set RS125 = Nothing

    frmStatus.Hide
    DoEvents
End If

xmlDoc.Save FilNa

If Dir$(FilNa) = vbNullString Then
    If GlLog = True Then SLogi "basPAD.S_ReExN: ERROR - File not saved: " & FilNa
End If

If GlLog = True Then SLogi "basPAD.S_ReExN: Completed - " & AnzRe & " invoices to " & FilNa

DoEvents
GlNeK = GlKoX
With GlNeK
    .PatNr = ManNr
    .IdxNr = 0
    .EiDat = Format$(Date, "dd.mm.yyyy")
    .EiZei = TimeValue(Now)
    .EiTyp = 104
    .TeStr = "Exportiert - PADNext Datei - (" & AnzRe & " Positionen)"
    .ZiStr = Format$(Now, "hh:mm") & " Uhr"
    .NeuEi = True
    .KeiAk = True
    .Mitar = GlMiA(GlSmI, 2)
End With
S_Prot

TeTit = "PADNext XML Export"
TeMai = AnzRe & " Rechnungen wurden exportiert"
TeInh = "Die markierten Rechnungen wurden im PADNext-Format unter: " & DaNam & " abgelegt."
TeFus = "Die PADNext-Exportdatei steht im folgenden Ordner zum Upload bereit:" & vbCrLf & FilNa

If AnzRe > 0 Then
    S_ReExN = True
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
End If

Exit Function

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_ReExN " & Err.Number
If GlLog = True Then SLogi "basPAD.S_ReExN: ERROR " & Err.Number & " - " & Err.Description
S_ReExN = False
Exit Function

End Function

' ============================================================================
' Function: S_ReExNAddEmpf
' Purpose: Add invoice recipient element to PADNext XML
' ============================================================================
Private Function S_ReExNAddEmpf(xmlDoc As Object, RS As ADODB.Recordset) As Object
On Error GoTo LiErr

Dim xmlEmpf As Object
Dim xmlPers As Object
Dim xmlElem As Object
Dim xmlAddr As Object
Dim xmlHaus As Object

Set xmlEmpf = PadXmlEl(xmlDoc, "rechnungsempfaenger")
Set xmlPers = PadXmlEl(xmlDoc, "person")

' Anrede
If Not IsNull(RS.Fields("R_Anrede").Value) And RS.Fields("R_Anrede").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "anrede")
    xmlElem.Text = RS.Fields("R_Anrede").Value
    xmlPers.appendChild xmlElem
End If

' Titel
If Not IsNull(RS.Fields("R_Titel").Value) And RS.Fields("R_Titel").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "titel")
    xmlElem.Text = RS.Fields("R_Titel").Value
    xmlPers.appendChild xmlElem
End If

' Vorname
If Not IsNull(RS.Fields("R_Vorname").Value) And RS.Fields("R_Vorname").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "vorname")
    xmlElem.Text = RS.Fields("R_Vorname").Value
    xmlPers.appendChild xmlElem
End If

' Name
If Not IsNull(RS.Fields("R_Name").Value) And RS.Fields("R_Name").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "name")
    xmlElem.Text = RS.Fields("R_Name").Value
    xmlPers.appendChild xmlElem
End If

' Geburtsdatum
If Not IsNull(RS.Fields("Geboren").Value) Then
    If IsDate(RS.Fields("Geboren").Value) Then
        Set xmlElem = PadXmlEl(xmlDoc, "gebdatum")
        xmlElem.Text = Format$(CDate(RS.Fields("Geboren").Value), "yyyy-MM-dd")
        xmlPers.appendChild xmlElem
    End If
End If

' Anschrift
Set xmlAddr = PadXmlEl(xmlDoc, "anschrift")
Set xmlHaus = PadXmlEl(xmlDoc, "hausadresse")

' Land - convert to ISO code
Dim TmpLnd As String
TmpLnd = vbNullString
If Not IsNull(RS.Fields("R_Land").Value) Then
    TmpLnd = RS.Fields("R_Land").Value
End If
Set xmlElem = PadXmlEl(xmlDoc, "land")
xmlElem.Text = PadLandCd(TmpLnd)
xmlHaus.appendChild xmlElem

' PLZ
If Not IsNull(RS.Fields("R_PLZ").Value) And RS.Fields("R_PLZ").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "plz")
    xmlElem.Text = RS.Fields("R_PLZ").Value
    xmlHaus.appendChild xmlElem
End If

' Ort
If Not IsNull(RS.Fields("R_Ort").Value) And RS.Fields("R_Ort").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "ort")
    xmlElem.Text = RS.Fields("R_Ort").Value
    xmlHaus.appendChild xmlElem
End If

' Strasse
If Not IsNull(RS.Fields("R_Strasse").Value) And RS.Fields("R_Strasse").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "strasse")
    xmlElem.Text = RS.Fields("R_Strasse").Value
    xmlHaus.appendChild xmlElem
End If

xmlAddr.appendChild xmlHaus
xmlPers.appendChild xmlAddr
xmlEmpf.appendChild xmlPers

Set S_ReExNAddEmpf = xmlEmpf
Exit Function

LiErr:
If GlLog = True Then SLogi "basPAD.S_ReExNAddEmpf: ERROR " & Err.Number & " - " & Err.Description
Set S_ReExNAddEmpf = Nothing
End Function

' ============================================================================
' Function: S_ReExNAddFall
' Purpose: Add billing case element to PADNext XML
' ============================================================================
Private Function S_ReExNAddFall(xmlDoc As Object, RS As ADODB.Recordset, AktRe As Long) As Object
On Error GoTo LiErr

Dim xmlFall As Object
Dim xmlHuman As Object
Dim xmlBeh As Object
Dim xmlVers As Object
Dim xmlDiag As Object
Dim xmlPos As Object
Dim xmlPosEl As Object
Dim xmlElem As Object
Dim DiaTe As String
Dim DiaKa As String
Dim PosNr As Integer
Dim GeschPad As String


Set xmlFall = PadXmlEl(xmlDoc, "abrechnungsfall")
Set xmlHuman = PadXmlEl(xmlDoc, "humanmedizin")

' Leistungserbringer ID
Set xmlElem = PadXmlEl(xmlDoc, "leistungserbringerid")
xmlElem.Text = "B1"
xmlHuman.appendChild xmlElem

' Behandelter
Set xmlBeh = PadXmlEl(xmlDoc, "behandelter")

' Vorname (patient first name)
If Not IsNull(RS.Fields("Vorname").Value) And RS.Fields("Vorname").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "vorname")
    xmlElem.Text = RS.Fields("Vorname").Value
    xmlBeh.appendChild xmlElem
End If

' Name (patient last name)
If Not IsNull(RS.Fields("Name").Value) And RS.Fields("Name").Value <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "name")
    xmlElem.Text = RS.Fields("Name").Value
    xmlBeh.appendChild xmlElem
End If

If Not IsNull(RS.Fields("Geboren").Value) Then
    If IsDate(RS.Fields("Geboren").Value) Then
        Set xmlElem = PadXmlEl(xmlDoc, "gebdatum")
        xmlElem.Text = Format$(CDate(RS.Fields("Geboren").Value), "yyyy-MM-dd")
        xmlBeh.appendChild xmlElem
    End If
End If

' Geschlecht (patient gender) - PADNext v2.12: m, w, u
If Not IsNull(RS.Fields("Geschlecht").Value) And RS.Fields("Geschlecht").Value <> vbNullString Then
    Select Case UCase$(Left$(RS.Fields("Geschlecht").Value, 1))
    Case "M": GeschPad = "m"
    Case "W": GeschPad = "w"
    Case Else: GeschPad = "u"
    End Select
    Set xmlElem = PadXmlEl(xmlDoc, "geschlecht")
    xmlElem.Text = GeschPad
    xmlBeh.appendChild xmlElem
End If

xmlHuman.appendChild xmlBeh

' Behandlungsart
Set xmlElem = PadXmlEl(xmlDoc, "behandlungsart")
xmlElem.Text = "0"
xmlHuman.appendChild xmlElem

' Vertragsart
Set xmlElem = PadXmlEl(xmlDoc, "vertragsart")
xmlElem.Text = "01"
xmlHuman.appendChild xmlElem

' Diagnosen
Set RS128 = New ADODB.Recordset
RS128.CursorLocation = adUseClient
Set RS128 = DBCmRe1("qrySimAbDiag", "@IdxNr", AktRe)

If RS128.RecordCount > 0 Then
    Do While Not RS128.EOF
        Set xmlDiag = PadXmlEl(xmlDoc, "diagnose")

        DiaTe = RS128.Fields("IDKurz").Value
        DiaKa = RS128.Fields("GOID").Value

        If DiaTe <> vbNullString Then
            Set xmlElem = PadXmlEl(xmlDoc, "text")
            xmlElem.Text = DiaTe
            xmlDiag.appendChild xmlElem
        End If

        If DiaKa <> vbNullString Then
            Set xmlElem = PadXmlEl(xmlDoc, "code")
            xmlElem.setAttribute "system", "ICD-10"
            xmlElem.Text = DiaKa
            xmlDiag.appendChild xmlElem
        End If

        xmlHuman.appendChild xmlDiag
        RS128.MoveNext
    Loop
End If
RS128.Close
Set RS128 = Nothing

' Positionen
Set xmlPos = PadXmlEl(xmlDoc, "positionen")

Set RS126 = New ADODB.Recordset
RS126.CursorLocation = adUseClient
Set RS126 = DBCmRe1("qrySimPAD2", "@IdxNr", AktRe)

PosNr = 0

' Configure PrBr2 for position progress (inner loop)
If RS126.RecordCount > 0 Then
    PrBr2.Min = 0
    PrBr2.Max = RS126.RecordCount
    PrBr2.Value = 0
End If

If RS126.RecordCount > 0 Then
    Do While Not RS126.EOF
        PosNr = PosNr + 1

        ' Update position progress bar (inner loop)
        If PosNr Mod 5 = 0 Or PosNr = 1 Then
            PrBr2.Value = PosNr
            DoEvents
        End If


            Set xmlPosEl = PadXmlEl(xmlDoc, "goziffer")
            xmlPosEl.setAttribute "positionsnr", CStr(PosNr)
            xmlPosEl.setAttribute "go", "GOAE"
            xmlPosEl.setAttribute "goversion", "02.01.2002"

            ' Ziffer (GONr)
            If Not IsNull(RS126.Fields("GONr").Value) And RS126.Fields("GONr").Value <> vbNullString Then
                xmlPosEl.setAttribute "ziffer", RS126.Fields("GONr").Value
            End If

            ' Datum
            If Not IsNull(RS126.Fields("Datum").Value) Then
                If IsDate(RS126.Fields("Datum").Value) Then
                    Set xmlElem = PadXmlEl(xmlDoc, "datum")
                    xmlElem.Text = Format$(CDate(RS126.Fields("Datum").Value), "yyyy-MM-dd")
                    xmlPosEl.appendChild xmlElem
                End If
            End If

            ' Anzahl
            If Not IsNull(RS126.Fields("Anzahl").Value) Then
                Set xmlElem = PadXmlEl(xmlDoc, "anzahl")
                xmlElem.Text = CStr(RS126.Fields("Anzahl").Value)
                xmlPosEl.appendChild xmlElem
            End If

            ' Text (Leistungsbeschreibung)
            If Not IsNull(RS126.Fields("Text").Value) And RS126.Fields("Text").Value <> vbNullString Then
                Set xmlElem = PadXmlEl(xmlDoc, "text")
                xmlElem.Text = RS126.Fields("Text").Value
                xmlPosEl.appendChild xmlElem
            End If

            ' Faktor (Multi)
            If Not IsNull(RS126.Fields("Multi").Value) Then
                If IsNumeric(RS126.Fields("Multi").Value) Then
                    Set xmlElem = PadXmlEl(xmlDoc, "faktor")
                    xmlElem.Text = Format$(CDbl(RS126.Fields("Multi").Value), "0.000000")
                    xmlPosEl.appendChild xmlElem
                Else
                End If
            End If

            ' Gesamtbetrag
            If Not IsNull(RS126.Fields("GesBetrag").Value) Then
                If IsNumeric(RS126.Fields("GesBetrag").Value) Then
                    Set xmlElem = PadXmlEl(xmlDoc, "gesamtbetrag")
                    xmlElem.Text = Format$(CCur(RS126.Fields("GesBetrag").Value), "0.00")
                    xmlPosEl.appendChild xmlElem
                Else
                End If
            End If

            xmlPos.appendChild xmlPosEl
        RS126.MoveNext
    Loop

    ' Set position progress bar to maximum when done (inner loop)
    PrBr2.Value = PrBr2.Max
    DoEvents
End If

xmlPos.setAttribute "posanzahl", CStr(PosNr)
xmlHuman.appendChild xmlPos

RS126.Close
Set RS126 = Nothing

' Log completion of position processing

xmlFall.appendChild xmlHuman
Set S_ReExNAddFall = xmlFall

Exit Function

LiErr:
If GlLog = True Then SLogi "basPAD.S_ReExNAddFall: ERROR " & Err.Number & " - " & Err.Description
Set S_ReExNAddFall = Nothing
End Function

' ============================================================================
' Function: S_ReExNAddRErs
' Purpose: Add rechnungsersteller element to PADNext XML
' Note: kundennr is REQUIRED in rechnungsersteller per XSD specification
' ============================================================================
Private Function S_ReExNAddRErs(xmlDoc As Object, ByVal ManNr As Long) As Object
On Error GoTo LiErr

Dim xmlRErs As Object
Dim xmlElem As Object
Dim xmlAnsc As Object
Dim xmlHaus As Object
Dim TmpNam As String
Dim TmpStr As String
Dim KunNrT As String
Dim TmpTit As String
Dim TmpVor As String
Dim TmpFir As String
Dim xmlBank As Object
Dim TmpIBA As String
Dim TmpBIC As String
Dim TmpBank As String

If GlLog = True Then SLogi "basPAD.S_ReExNAddRErs: Creating rechnungsersteller for ManNr=" & ManNr

If ManNr = 0 Then
    If GlLog = True Then SLogi "basPAD.S_ReExNAddRErs: Invalid ManNr=0"
    Set S_ReExNAddRErs = Nothing
    Exit Function
End If

Set xmlRErs = PadXmlEl(xmlDoc, "rechnungsersteller")

' Name (required) - combine Titel + Vorname + Name
TmpNam = vbNullString
TmpTit = S_AdIdx(ManNr, "Titel")
TmpVor = S_AdIdx(ManNr, "Vorname")
TmpStr = S_AdIdx(ManNr, "Name")
TmpFir = S_AdIdx(ManNr, "Firma1")

If TmpTit <> vbNullString Then
    TmpNam = Trim$(TmpTit)
End If
If TmpVor <> vbNullString Then
    If TmpNam <> vbNullString Then TmpNam = TmpNam & " "
    TmpNam = TmpNam & Trim$(TmpVor)
End If
If TmpStr <> vbNullString Then
    If TmpNam <> vbNullString Then TmpNam = TmpNam & " "
    TmpNam = TmpNam & Trim$(TmpStr)
End If

If TmpNam = vbNullString Then
    ' Fallback to Firma1 if no name
    If TmpFir <> vbNullString Then
        TmpNam = Trim$(TmpFir)
    End If
End If

If TmpNam <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "name")
    xmlElem.Text = Left$(TmpNam, 40)
    xmlRErs.appendChild xmlElem
Else
    If GlLog = True Then SLogi "basPAD.S_ReExNAddRErs: WARNING - No name found for mandant"
End If

' Kundennr (required) - from Postfach field
KunNrT = S_AdIdx(ManNr, "Postfach")
If KunNrT <> vbNullString Then
    KunNrT = Trim$(KunNrT)
    ' Remove leading zeros for XML positive integer
    If IsNumeric(KunNrT) Then
        KunNrT = CStr(CLng(KunNrT))
    End If
End If

If KunNrT <> vbNullString And KunNrT <> "0" And KunNrT <> "999999" Then
    Set xmlElem = PadXmlEl(xmlDoc, "kundennr")
    xmlElem.Text = KunNrT
    xmlRErs.appendChild xmlElem
Else
    If GlLog = True Then SLogi "basPAD.S_ReExNAddRErs: WARNING - No valid kundennr (Postfach) found"
End If

' Anschrift (required)
Set xmlAnsc = PadXmlEl(xmlDoc, "anschrift")
Set xmlHaus = PadXmlEl(xmlDoc, "hausadresse")

' Land - convert to ISO code
TmpStr = S_AdIdx(ManNr, "Land")
Set xmlElem = PadXmlEl(xmlDoc, "land")
xmlElem.Text = PadLandCd(TmpStr)
xmlHaus.appendChild xmlElem

' PLZ
TmpStr = S_AdIdx(ManNr, "PLZ")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "plz")
    xmlElem.Text = Trim$(TmpStr)
    xmlHaus.appendChild xmlElem
End If

' Ort
TmpStr = S_AdIdx(ManNr, "Ort")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "ort")
    xmlElem.Text = Trim$(TmpStr)
    xmlHaus.appendChild xmlElem
End If

' Strasse - construct field name with special character
TmpStr = S_AdIdx(ManNr, "Stra" & Chr$(223) & "e")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "strasse")
    xmlElem.Text = Trim$(TmpStr)
    xmlHaus.appendChild xmlElem
End If

xmlAnsc.appendChild xmlHaus
xmlRErs.appendChild xmlAnsc

' Kontakt - Telefon (optional)
TmpStr = S_AdIdx(ManNr, "Telefon1")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "kontakt")
    xmlElem.setAttribute "typ", "beruflich"
    xmlElem.setAttribute "art", "telefonnr"
    xmlElem.Text = Trim$(TmpStr)
    xmlRErs.appendChild xmlElem
End If

' Bankverbindung - IBAN ist in PADNext v2.12 Pflicht
TmpIBA = Trim$(S_AdIdx(ManNr, "IBAN"))
TmpBIC = Trim$(S_AdIdx(ManNr, "BIC"))
TmpBank = Trim$(S_AdIdx(ManNr, "Bankname"))
If TmpIBA <> vbNullString Then
    Set xmlBank = PadXmlEl(xmlDoc, "bankverbindung")
    Set xmlElem = PadXmlEl(xmlDoc, "iban")
    xmlElem.Text = TmpIBA
    xmlBank.appendChild xmlElem
    If TmpBIC <> vbNullString Then
        Set xmlElem = PadXmlEl(xmlDoc, "bic")
        xmlElem.Text = TmpBIC
        xmlBank.appendChild xmlElem
    End If
    If TmpBank <> vbNullString Then
        Set xmlElem = PadXmlEl(xmlDoc, "kreditinstitut")
        xmlElem.Text = Left$(TmpBank, 100)
        xmlBank.appendChild xmlElem
    End If
    xmlRErs.appendChild xmlBank
End If

If GlLog = True Then SLogi "basPAD.S_ReExNAddRErs: rechnungsersteller created successfully"
Set S_ReExNAddRErs = xmlRErs

Exit Function

LiErr:
If GlLog = True Then SLogi "basPAD.S_ReExNAddRErs: ERROR " & Err.Number & " - " & Err.Description
Set S_ReExNAddRErs = Nothing
End Function

' ============================================================================
' Function: S_ReExNAddLErb
' Purpose: Add leistungserbringer element to PADNext XML
' Note: id attribute is required, kundennr is optional but recommended
' ============================================================================
Private Function S_ReExNAddLErb(xmlDoc As Object, ByVal ManNr As Long) As Object
On Error GoTo LiErr

Dim xmlLErb As Object
Dim xmlElem As Object
Dim xmlAnsc As Object
Dim xmlHaus As Object
Dim TmpStr As String
Dim KunNrT As String

If GlLog = True Then SLogi "basPAD.S_ReExNAddLErb: Creating leistungserbringer for ManNr=" & ManNr

If ManNr = 0 Then
    If GlLog = True Then SLogi "basPAD.S_ReExNAddLErb: Invalid ManNr=0"
    Set S_ReExNAddLErb = Nothing
    Exit Function
End If

Set xmlLErb = PadXmlEl(xmlDoc, "leistungserbringer")
xmlLErb.setAttribute "id", "B1"
xmlLErb.setAttribute "aisid", "B1"

' Anrede (optional but valid values: Herr, Frau, Praxis, Labor)
TmpStr = S_AdIdx(ManNr, "Anrede")
If TmpStr <> vbNullString Then
    ' Map to valid PADNext anrede values
    Select Case LCase$(Trim$(TmpStr))
    Case "herr", "hr", "hr.", "m"
        TmpStr = "Herr"
    Case "frau", "fr", "fr.", "w"
        TmpStr = "Frau"
    Case "praxis"
        TmpStr = "Praxis"
    Case "labor"
        TmpStr = "Labor"
    Case Else
        TmpStr = vbNullString
    End Select
    If TmpStr <> vbNullString Then
        Set xmlElem = PadXmlEl(xmlDoc, "anrede")
        xmlElem.Text = TmpStr
        xmlLErb.appendChild xmlElem
    End If
End If

' Titel (optional)
TmpStr = S_AdIdx(ManNr, "Titel")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "titel")
    xmlElem.Text = Trim$(TmpStr)
    xmlLErb.appendChild xmlElem
End If

' Vorname (required)
TmpStr = S_AdIdx(ManNr, "Vorname")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "vorname")
    xmlElem.Text = Trim$(TmpStr)
    xmlLErb.appendChild xmlElem
End If

' Name (required)
TmpStr = S_AdIdx(ManNr, "Name")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "name")
    xmlElem.Text = Trim$(TmpStr)
    xmlLErb.appendChild xmlElem
End If

' Anschrift (optional for leistungserbringer)
Set xmlAnsc = PadXmlEl(xmlDoc, "anschrift")
Set xmlHaus = PadXmlEl(xmlDoc, "hausadresse")

' Land - convert to ISO code
TmpStr = S_AdIdx(ManNr, "Land")
Set xmlElem = PadXmlEl(xmlDoc, "land")
xmlElem.Text = PadLandCd(TmpStr)
xmlHaus.appendChild xmlElem

' PLZ
TmpStr = S_AdIdx(ManNr, "PLZ")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "plz")
    xmlElem.Text = Trim$(TmpStr)
    xmlHaus.appendChild xmlElem
End If

' Ort
TmpStr = S_AdIdx(ManNr, "Ort")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "ort")
    xmlElem.Text = Trim$(TmpStr)
    xmlHaus.appendChild xmlElem
End If

' Strasse - construct field name with special character
TmpStr = S_AdIdx(ManNr, "Stra" & Chr$(223) & "e")
If TmpStr <> vbNullString Then
    Set xmlElem = PadXmlEl(xmlDoc, "strasse")
    xmlElem.Text = Trim$(TmpStr)
    xmlHaus.appendChild xmlElem
End If

xmlAnsc.appendChild xmlHaus
xmlLErb.appendChild xmlAnsc

' Kundennr (optional for leistungserbringer but recommended)
KunNrT = S_AdIdx(ManNr, "Postfach")
If KunNrT <> vbNullString Then
    KunNrT = Trim$(KunNrT)
    If IsNumeric(KunNrT) Then
        KunNrT = CStr(CLng(KunNrT))
    End If
End If

If KunNrT <> vbNullString And KunNrT <> "0" And KunNrT <> "999999" Then
    Set xmlElem = PadXmlEl(xmlDoc, "kundennr")
    xmlElem.Text = KunNrT
    xmlLErb.appendChild xmlElem
End If

If GlLog = True Then SLogi "basPAD.S_ReExNAddLErb: leistungserbringer created successfully"
Set S_ReExNAddLErb = xmlLErb

Exit Function

LiErr:
If GlLog = True Then SLogi "basPAD.S_ReExNAddLErb: ERROR " & Err.Number & " - " & Err.Description
Set S_ReExNAddLErb = Nothing
End Function

' ============================================================================
' Function: PadXmlEl
' Purpose: Create XML element with PADNext namespace
' Note: Uses createNode instead of createElement to set correct namespace
' ============================================================================
Private Function PadXmlEl(xmlDoc As Object, ByVal ElName As String) As Object
    Set PadXmlEl = xmlDoc.createNode(1, ElName, PADX_NS)
End Function

' ============================================================================
' Function: PadLandCd
' Purpose: Convert country name to ISO 3166-1 alpha-1 code for PADNext
' ============================================================================
Private Function PadLandCd(ByVal LandNm As String) As String
Dim TmpLnd As String
TmpLnd = LCase$(Trim$(LandNm))

Select Case TmpLnd
Case "d", "de", "deu", "deutschland", "germany"
    PadLandCd = "D"
Case "a", "at", "aut", "oesterreich", Chr$(246) & "sterreich", "austria"
    PadLandCd = "A"
Case "ch", "che", "schweiz", "switzerland"
    PadLandCd = "CH"
Case "f", "fr", "fra", "frankreich", "france"
    PadLandCd = "F"
Case "i", "it", "ita", "italien", "italy"
    PadLandCd = "I"
Case "nl", "nld", "niederlande", "netherlands"
    PadLandCd = "NL"
Case "b", "be", "bel", "belgien", "belgium"
    PadLandCd = "B"
Case "l", "lu", "lux", "luxemburg", "luxembourg"
    PadLandCd = "L"
Case "pl", "pol", "polen", "poland"
    PadLandCd = "PL"
Case "cz", "cze", "tschechien", "czech"
    PadLandCd = "CZ"
Case "dk", "dnk", "daenemark", Chr$(228) & "nemark", "denmark"
    PadLandCd = "DK"
Case vbNullString, ""
    PadLandCd = "D"
Case Else
    ' If already a short code, return as-is (uppercase)
    If Len(TmpLnd) <= 3 Then
        PadLandCd = UCase$(TmpLnd)
    Else
        PadLandCd = "D"
    End If
End Select
End Function
