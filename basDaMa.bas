Attribute VB_Name = "basDaMa"
Option Explicit

Private FM As Form
Private TxDum As VB.TextBox
Private TxReT As VB.TextBox
Private Lbl01 As XtremeSuiteControls.Label
Private CmTyp As XtremeSuiteControls.ComboBox
Private CaCol As XtremeCalendarControl.CalendarControl
Private DaPro As XtremeCalendarControl.CalendarDataProvider
Private CaLbs As XtremeCalendarControl.CalendarEventLabels
Private CaLbl As XtremeCalendarControl.CalendarEventLabel
Private TxDe3 As XtremeSuiteControls.FlatEdit
Private CoDia As XtremeSuiteControls.CommonDialog
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmPan As XtremeCommandBars.StatusBarPane
Private TrLi2 As XtremeSuiteControls.TreeView
Private TrLi5 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private PrGr1 As XtremePropertyGrid.PropertyGrid
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpHcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpGrw As XtremeReportControl.ReportGroupRow
Private ChRow As XtremeReportControl.ReportRow
Private ChRws As XtremeReportControl.ReportRows
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private TsDia As XtremeSuiteControls.TaskDialog

Private RS150 As ADODB.Recordset
Private RS151 As ADODB.Recordset
Private RS152 As ADODB.Recordset
Private RS153 As ADODB.Recordset
Private RS154 As ADODB.Recordset
Private RS155 As ADODB.Recordset
Private RS156 As ADODB.Recordset
Private RS157 As ADODB.Recordset
Private RS158 As ADODB.Recordset
Private RS159 As ADODB.Recordset
Private RS160 As ADODB.Recordset
Private RS161 As ADODB.Recordset
Private RS162 As ADODB.Recordset
Private RS163 As ADODB.Recordset
Private RS164 As ADODB.Recordset
Private RS165 As ADODB.Recordset
Private RS166 As ADODB.Recordset
Private RS167 As ADODB.Recordset
Private RS168 As ADODB.Recordset
Private RS169 As ADODB.Recordset
Private RS170 As ADODB.Recordset
Private RS171 As ADODB.Recordset
Private RS172 As ADODB.Recordset
Private RS173 As ADODB.Recordset
Private PA101 As ADODB.Parameter
Private CM101 As ADODB.Command
Private FL101 As ADODB.Field

Private Const olContactItem = 2
Private Const olAppointmentItem = 1
Private Const olFolderCalendar = 9
Private Const olFolderContacts = 10

Private clFil As clsFile
Private clAnw As clsAnwend
Private clFen As clsFenster
Private clLis As clsLisLab
Private clDru As clsDruck
Private clNet As clsNetz
Private clWor As clsWord
Private clLiz As clsLizenz
Private clAbd As clsDaAb
Private clICS As clsICS
Public Sub DBWaKl()
On Error GoTo DaErr
'Führt Standardkontollen bei Tabellen durch

DBCmEx1 "qrySimReBeh", "@IdxNr", GlMan(GlSMa, 2)
DoEvents
DBCmEx1 "qrySimReBeM", "@IdxNr", GlMiA(GlSmI, 2)
DoEvents
DBCmEx1 "qryWarWrA2", "@IdxNr", GlWar(1, 0)
DoEvents
DBCmEx1 "qryWarWrB2", "@IdxNr", GlWar(1, 0)
DoEvents
DBCmEx1 "qryWarWrR2", "@IdxNr", GlWar(1, 0)
DoEvents
DBCmEx1 "qryWarWrO2", "@IdxNr", GlWar(1, 0)
DoEvents
DBCmEx1 "qryWarIDP1a", "@IdxNr", GlMan(GlSMa, 2)
DoEvents
DBCmEx1 "qryWarIDP2a", "@IdxNr", GlMan(GlSMa, 2)
DoEvents
'DBCmEx1 "qryWarIDP3a", "@IdxNr", GlMan(GlSMa, 2)
DoEvents
DBCmEx1 "qryWarIDP4a", "@IdxNr", GlMan(GlSMa, 2)
DoEvents
DBCmEx1 "qryWarIDP5a", "@IdxNr", GlMan(GlSMa, 2)
DoEvents

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "DBWaKl " & Err.Number
Resume Next

End Sub
Public Sub S_AnBoA()
On Error GoTo SuErr
'Abfrufen eines Fragebogens

Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim IniNa As String
Dim TmpSt As String
Dim TmpZe As String
Dim TmpFe As String
Dim RetSt As String
Dim ErrSt As String
Dim FiNam As String
Dim ImpOr As String
Dim TmGui As String
Dim AktZe As Integer
Dim AktFe As Integer
Dim GesZe As Integer
Dim GesFe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim AryZe() As String
Dim AryFe() As String

Set FM = frmMain

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Sub
    End If
End If

Screen.MousePointer = vbHourglass
DoEvents

TeTit = "Fragebögen Abrufen"
TeMai = "Soll der Fragebogenabruf jetzt gestartet werden?"
TeInh = "Beim Fragebogenabruf werden die Antworten der von den Patienten ausgefüllten Fragebögen heruntergeladen und Ihren Patienten zugeordnet."
TeFus = "Damit die Fragebogendaten dem richtigen Patienten zugeordnet werden können, ist es notwendig, dass diese bereits in Ihren Stammdaten erfasst wurden."

TmGui = CreateID("D")
FiNam = GlTEx & TmGui & ".csv" 'Termineordner
IniNa = GlTmp & TmGui & ".ini"

ImpOr = GlTEx
If Right$(ImpOr, 1) = "\" Then
    Lange = Len(ImpOr)
    ImpOr = Left$(ImpOr, Lange - 1)
End If

If GlCID <> vbNullString Then 'Cloud-ID
    PrNam = Chr$(34) & PrNam & Chr$(34)

    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
    If GlMes = 33565 Then
        
        PaStr = "download" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & "--csv=" & Chr$(34) & FiNam & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & IniNa & Chr$(34) & Space$(1) & "--zip=" & Chr$(34) & ImpOr & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        With clFil
            If .FilVor(IniNa) = True Then
                .FilPfa IniNa
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If AryZe(AktZe) <> vbNullString Then
                            TmpZe = AryZe(AktZe)
                            Lange = Len(TmpZe)
                            Posit = InStr(1, TmpZe, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmpZe, Posit - 1))
                                Select Case InTyp
                                Case "formsubmissioncount": RetSt = Right$(TmpZe, Lange - Posit)
                                Case "error": ErrSt = Right$(TmpZe, Lange - Posit)
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If

            If RetSt <> vbNullString Then
                If IsNumeric(RetSt) = True Then
                    GlFrD = CInt(RetSt) 'Anzahl Fragebogendownload
                Else
                    GlFrD = 0
                End If
                If GlFrD > 0 Then
                    If .FilVor(FiNam) = True Then
                        .FilPfa FiNam
                        TmpSt = .FilReSt
                        DoEvents
                        If TmpSt <> vbNullString Then
                            AryZe = Split(TmpSt, Chr$(10)) 'Zeilen aufsplitten
                            GesZe = UBound(AryZe)
                            TmpZe = AryZe(0)
                            TmpZe = Replace(TmpZe, vbCrLf, vbNullString, 1)
                            If GesZe > 0 Then
                                AryFe = Split(TmpZe, Chr$(59)) 'Feldnamen aufsplitten
                                GesFe = UBound(AryFe)
                                Set RS164 = New ADODB.Recordset 'neues Recordset erstellen Patienten
                                RS164.CursorLocation = adUseClient
                                For AktFe = 0 To UBound(AryFe)
                                    TmpFe = AryFe(AktFe)
                                    TmpFe = Replace(TmpFe, vbCrLf, vbNullString, 1)
                                    RS164.Fields.Append TmpFe, adVarChar, 250
                                Next AktFe
                                If RS164.State = adStateClosed Then RS164.Open
    
                                For AktZe = 1 To GesZe - 1 'WICHTIG! ab Zeile 1 begindnen
                                    If AryZe(AktZe) <> vbNullString Then
                                        AryFe = Split(AryZe(AktZe), Chr$(59)) 'Felder aufsplitten
                                        RS164.AddNew
                                        RS164.Fields(0).Value = AryFe(0)
                                        RS164.Fields(1).Value = AryFe(1)
                                        RS164.Fields(2).Value = AryFe(2)
                                        RS164.Fields(3).Value = AryFe(3)
                                        RS164.Fields(4).Value = AryFe(4)
                                        RS164.Fields(5).Value = SUmw(AryFe(5), True, False, True, False)
                                        RS164.Update
                                    End If
                                Next AktZe

                                Set frmZuord.FoRST = RS164
                                SAnIn '---------- Fragebogen Zuordnungsdialog ----------
                            Else
                                SPopu "Fragebogenabruf", "Der Fragebogen enthält keine Informationen.", IC48_Information
                            End If
                        End If
                    Else
                        If ErrSt = vbNullString Then
                            SPopu "Fragebogenabruf", "Es liegen keine neuen Fragebögen vor.", IC48_Information
                        Else
                            SPopu "Downloadfehler", ErrSt, IC48_Information
                        End If
                        SAnUm FiNam, IniNa
                    End If
                Else
                    SPopu "Fragebogenabruf", "Es liegen keine neuen Fragebögen vor.", IC48_Information
                    SAnUm FiNam, IniNa
                End If
            Else
                Clipboard.Clear
                Clipboard.SetText PrNam & Space$(1) & PaStr
                If ErrSt = vbNullString Then
                    SPopu "Fragebogenabruf", "Es liegen keine neuen Fragebögen vor.", IC48_Information
                Else
                    SPopu "Downloadfehler", ErrSt, IC48_Information
                End If
                SAnUm FiNam, IniNa
            End If

            If ErrSt <> vbNullString Then
                Clipboard.Clear
                Clipboard.SetText PrNam & Space$(1) & PaStr
            End If
        End With
        DoEvents

        GlNeK = GlKoX 'Protokolleintrag
        With GlNeK
            .PatNr = GlMan(GlSMa, 2)
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = 104
            .TeStr = "Fragebogen abgerufen " & TmGui & " " & ErrSt
            .ZiStr = Format$(Now, "hh:mm") & " Uhr"
            .NeuEi = True
            .KeiAk = True
            .Mitar = GlMiA(GlSmI, 2)
        End With
        S_Prot
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoA " & Err.Number
Resume Next

End Sub
Public Function S_AnBoB(ByVal IdStr As String, ByVal FelNa As String) As String
On Error GoTo SuErr
'Gibt Detailinformationen des Patientenfragebogens zurück

Set RS165 = New ADODB.Recordset
RS165.CursorLocation = adUseClient
Set RS165 = DBCmRe1("qryPatAnGui", "@IdStr", IdStr)
If RS165.RecordCount > 0 Then
    S_AnBoB = RS165.Fields(FelNa).Value
Else
    S_AnBoB = vbNullString
End If
RS165.Close
Set RS165 = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoB " & Err.Number
Resume Next

End Function
Public Sub S_AnBoC()
On Error GoTo SuErr
'Prüft ob noch Fargebogendateien vorliegen

Dim InTyp As String
Dim IniNa As String
Dim FiNam As String
Dim TmpSt As String
Dim TmpZe As String
Dim TmpFe As String
Dim RetSt As String
Dim ErrSt As String
Dim PaStr As String
Dim AnzDa As Integer
Dim AnzIn As Integer
Dim AktZe As Integer
Dim AktFe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim GesZe As Integer
Dim GesFe As Integer
Dim DiNam() As String
Dim InNam() As String
Dim AryZe() As String
Dim AryFe() As String

Set FM = frmMain

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

FiNam = GlTEx & "*.csv" 'Termineordner
IniNa = GlTmp & "*.ini"

With clFil
    If .FilVor(GlTEx & "*.csv") = True Then
        AnzDa = .FilLis(LCase(GlTEx), "*.csv", DiNam)
        If AnzDa > 0 Then
            FiNam = GlTEx & DiNam(1)
        End If
    End If
    If .FilVor(GlTmp & "*.ini") = True Then
        AnzIn = .FilLis(LCase(GlTmp), "*.ini", InNam)
        If AnzIn > 0 Then
            IniNa = GlTmp & InNam(1)
        End If
    End If
End With

If AnzDa = 0 Then 'keine alten Abrufdateien vorhanden
    Set clFil = Nothing
    S_AnBoA 'Fragebögen abrufen
Else
    With clFil
        If .FilVor(IniNa) = True Then
            .FilPfa IniNa
            TmpSt = .FilReSt
            DoEvents
            If TmpSt <> vbNullString Then
                AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                For AktZe = 0 To UBound(AryZe) - 1
                    If AryZe(AktZe) <> vbNullString Then
                        TmpZe = AryZe(AktZe)
                        Lange = Len(TmpZe)
                        Posit = InStr(1, TmpZe, "=", 1)
                        If Posit > 0 Then
                            InTyp = LCase(Left$(TmpZe, Posit - 1))
                            Select Case InTyp
                            Case "formsubmissioncount": RetSt = Right$(TmpZe, Lange - Posit)
                            Case "error": ErrSt = Right$(TmpZe, Lange - Posit)
                            End Select
                        End If
                    End If
                Next AktZe
            End If
        End If

        If RetSt <> vbNullString Then
            If IsNumeric(RetSt) = True Then
                GlFrD = CInt(RetSt) 'Anzahl Fragebogendownload
            Else
                GlFrD = 0
            End If
            If GlFrD > 0 Then
                If .FilVor(FiNam) = True Then
                    .FilPfa FiNam
                    TmpSt = .FilReSt
                    DoEvents
                    If TmpSt <> vbNullString Then
                        AryZe = Split(TmpSt, Chr$(10)) 'Zeilen aufsplitten
                        GesZe = UBound(AryZe)
                        TmpZe = AryZe(0)
                        TmpZe = Replace(TmpZe, vbCrLf, vbNullString, 1)
                        If GesZe > 0 Then
                            AryFe = Split(TmpZe, Chr$(59)) 'Feldnamen aufsplitten
                            GesFe = UBound(AryFe)
                            Set RS164 = New ADODB.Recordset 'neues Recordset erstellen Patienten
                            RS164.CursorLocation = adUseClient
                            For AktFe = 0 To GesFe
                                TmpFe = AryFe(AktFe)
                                TmpFe = Replace(TmpFe, vbCrLf, vbNullString, 1)
                                RS164.Fields.Append TmpFe, adVarChar, 250
                            Next AktFe
                            If RS164.State = adStateClosed Then RS164.Open

                            For AktZe = 1 To GesZe 'WICHTIG! ab Zeile 1 begindnen
                                If AryZe(AktZe) <> vbNullString Then
                                    AryFe = Split(AryZe(AktZe), Chr$(59)) 'Felder aufsplitten
                                    RS164.AddNew
                                    RS164.Fields(0).Value = AryFe(0)
                                    RS164.Fields(1).Value = AryFe(1)
                                    RS164.Fields(2).Value = AryFe(2)
                                    RS164.Fields(3).Value = AryFe(3)
                                    RS164.Fields(4).Value = AryFe(4)
                                    RS164.Fields(5).Value = SUmw(AryFe(5), True, False, True, False)
                                    RS164.Update
                                End If
                            Next AktZe

                            Set frmZuord.FoRST = RS164
                            SAnIn '---------- Fragebogen Zuordnungsdialog ----------
                        Else
                            SPopu "Fragebogenabruf", "Der Fragebogen enthält keine Informationen.", IC48_Information
                            SAnUm FiNam, IniNa 'Fragebogendatei umbenennen
                        End If
                    End If
                Else
                    If ErrSt = vbNullString Then
                        SPopu "Einlesefeher", "Unerwarteter Fehler, bei der Verarbeitung einer Fragebogendatei", IC48_Forbidden
                    Else
                        SPopu "Einlesefeher", ErrSt, IC48_Forbidden
                    End If
                    SAnUm FiNam, IniNa
                End If
            Else
                SPopu "Fragebogenimport", "Es liegen keine neuen Fragebögen vor.", IC48_Information
            End If
        Else
            If ErrSt = vbNullString Then
                SPopu "Einlesefeher", "Unerwarteter Fehler, bei der Verarbeitung einer Fragebogendatei", IC48_Forbidden
            Else
                SPopu "Einlesefeher", ErrSt, IC48_Forbidden
            End If
            SAnUm FiNam, IniNa
        End If
    End With
    
End If

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoC " & Err.Number
Resume Next

End Sub
Public Function S_AnBoD(ByVal IdxNr As Long, ByVal FelNa As String) As String
On Error GoTo SuErr
'Gibt Detailinformationen einer Frage aus

Set RS165 = New ADODB.Recordset
RS165.CursorLocation = adUseClient
Set RS165 = DBCmRe1("qryKat10a", "@IdxNr", IdxNr) 'Fragebogen
If RS165.RecordCount > 0 Then
    S_AnBoD = RS165.Fields(FelNa).Value
Else
    S_AnBoD = vbNullString
End If
RS165.Close
Set RS165 = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoD " & Err.Number
Resume Next

End Function

Public Function S_AnBoF(ByVal IdStr As String, ByVal FelNa As String) As String
On Error GoTo SuErr
'Gibt Detailinformationen des Fragebogens aus

Set RS165 = New ADODB.Recordset
RS165.CursorLocation = adUseClient
Set RS165 = DBCmRe1("qryKat05D", "@IdKey", IdStr) 'Fragebogen
If RS165.RecordCount > 0 Then
    S_AnBoF = RS165.Fields(FelNa).Value
Else
    S_AnBoF = vbNullString
End If
RS165.Close
Set RS165 = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoF " & Err.Number
Resume Next

End Function
Public Function S_AnBoG(ByVal IdxNr As Long, ByVal FelNa As String) As String
On Error GoTo SuErr
'Gibt Detailinformationen einer Frage aus

Set RS165 = New ADODB.Recordset
RS165.CursorLocation = adUseClient
Set RS165 = DBCmRe1("qryKat08A", "@IdxNr", IdxNr) 'Fragebogen
If RS165.RecordCount > 0 Then
    S_AnBoG = RS165.Fields(FelNa).Value
Else
    S_AnBoG = vbNullString
End If
RS165.Close
Set RS165 = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoG " & Err.Number
Resume Next

End Function
Public Sub S_AnBoH()
On Error GoTo SuErr
'Fuegt einen zugeordneten Fragebogen hinzu

Dim SQL1 As String
Dim SQL2 As String
Dim BoDat As Date
Dim BogNr As Long
Dim BoNum As Long
Dim PatNr As Long
Dim ManNr As Long
Dim FrgNr As Long
Dim SubNr As Long
Dim SorNr As Long
Dim BogID As String
Dim BogNa As String
Dim FrgWe As String
Dim FrgID As String
Dim WebID As String
Dim BerTx As String
Dim ImOrd As String
Dim ZipNa As String
Dim TypNr As Integer
Dim Warte As Boolean
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo1 As XtremeReportControl.ReportControl

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpRws = RpCo1.Rows

Set RS160 = FM.FoRST

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatAnBoN WHERE ID5 = -1"
    SQL2 = "SELECT * FROM dbo.qryPatAnNeu WHERE ID1 = -1"
Else
    SQL1 = "SELECT * FROM qryPatAnBoN WHERE [ID5] = -1;"
    SQL2 = "SELECT * FROM qryPatAnNeu WHERE [ID1] = -1;"
End If

Set RS169 = New ADODB.Recordset
With RS169
    .CursorLocation = adUseClient
    .Source = SQL1 'Fragebögen der Patienten
    .ActiveConnection = DB1
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Options:=adCmdText
End With
If RS169.Supports(adAddNew) Then
    For Each RpRow In RpRws
        If RpRow.GroupRow = False Then
            If RpRow.Record(19).CheckboxState = 0 Then 'Wenn Fragebogen nicht gelöscht werden soll
                If RpRow.Record(17).Value <> vbNullString Then
                    PatNr = RpRow.Record(17).Value '[ID0]
                    If PatNr > 0 Then
                        BoNum = 0
                        BogNa = vbNullString
                        BogID = RpRow.Record(0).Value '[BogID]
                        WebID = RpRow.Record(14).Value '[SubMisID]
                        If RpRow.Record(1).Value <> vbNullString Then
                            If IsDate(RpRow.Record(1).Value) = True Then
                                BoDat = RpRow.Record(1).Value '[Datum]
                            End If
                        End If
                        BogNr = Val(S_AnBoF(BogID, "ID3")) 'Fragebogennummer
                        BogNa = S_AnBoF(BogID, "IDKurz") 'Fragebogenname
                        ManNr = S_AdIdx(PatNr, "IDP")
                        If BogNr > 0 Then
                            RS169.AddNew
                            RS169.Fields("ID0").Value = PatNr
                            RS169.Fields("ID3").Value = BogNr
                            RS169.Fields("IDP").Value = ManNr
                            RS169.Fields("Datum").Value = BoDat
                            RS169.Fields("IDKurz").Value = BogNa
                            RS169.Fields("GuiID").Value = WebID
                            RS169.Fields("GuiKey").Value = BogID
                            RS169.Update
                            DoEvents
                            BoNum = CLng(S_AnBoB(WebID, "ID5"))
                            RpRow.Record(18).Value = BoNum 'Patientenbogennummer
                        Else
                            RpRow.Record(18).Value = 0
                        End If
                    End If
                End If
            End If
        End If
    Next RpRow
End If
RS169.Close
Set RS169 = Nothing

Set RS171 = New ADODB.Recordset
With RS171
    .CursorLocation = adUseClient
    .Source = SQL2 'Fragebogenfragen
    .ActiveConnection = DB1
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Options:=adCmdText
End With
If RS171.Supports(adAddNew) Then
    For Each RpRow In RpRws
        If RpRow.GroupRow = False Then
            If RpRow.Record(19).CheckboxState = 0 Then 'Fragebogen löschen
                If RpRow.Record(17).Value <> vbNullString Then
                    SorNr = 0
                    PatNr = RpRow.Record(17).Value
                    If PatNr > 0 Then
                        BogID = RpRow.Record(0).Value '[BogID]
                        BogNr = Val(S_AnBoF(BogID, "ID3")) 'Fragebogennummer
                        BoNum = RpRow.Record(18).Value '[BogenNr]
                        If BoNum > 0 Then
                            Set RS165 = New ADODB.Recordset
                            RS165.CursorLocation = adUseClient 'Fragebogen
                            Set RS165 = DBCmRe1("qryKat08I", "@IdxNr", BogNr)
                            If RS165.RecordCount > 0 Then
                                Do
                                SorNr = SorNr + 1
                                FrgNr = RS165.Fields("ID6").Value 'FragenNr
                                SubNr = RS165.Fields("IDSub").Value
                                FrgID = RS165.Fields("GuiID").Value 'FragenID
                                If RS165.Fields("Type").Value <> vbNullString Then
                                    TypNr = RS165.Fields("Type").Value
                                Else
                                    TypNr = 1
                                End If
                                RS171.AddNew
                                RS171.Fields("ID6").Value = FrgNr
                                RS171.Fields("ID5").Value = BoNum
                                RS171.Fields("ID0").Value = PatNr
                                RS171.Fields("Type").Value = TypNr
                                RS171.Fields("GuiID").Value = FrgID
                                RS171.Fields("Sorter").Value = SorNr
                                If RS165.Fields("Pflicht").Value <> vbNullString Then
                                    RS171.Fields("Pflicht").Value = RS165.Fields("Pflicht").Value
                                Else
                                    RS171.Fields("Pflicht").Value = 0
                                End If
                                RS171.Fields("VorCheck").Value = RS165.Fields("VorCheck").Value
                                RS171.Fields("VorText").Value = RS165.Fields("VorText").Value
                                If BogNr <> SubNr Then
                                    RS171.Fields("IDSub").Value = SubNr
                                    If SubNr > 0 Then
                                        RS171.Fields("Subfeld").Value = -1
                                    End If
                                End If
                                RS165.MoveNext
                                Loop Until RS165.EOF
                                DoEvents
                                RS171.UpdateBatch
                            End If
                            RS165.Close
                            Set RS165 = Nothing
                        End If
                    End If
                End If
            End If
        End If
    Next RpRow
End If
RS171.Close
Set RS171 = Nothing

If RS160.RecordCount > 0 Then
    For Each RpRow In RpRws
        If RpRow.GroupRow = False Then
            If RpRow.Record(19).CheckboxState = 0 Then 'Wenn Fragebogen nicht gelöscht werden soll
                If RpRow.Record(17).Value <> vbNullString Then
                    PatNr = RpRow.Record(17).Value '[ID0]
                    If PatNr > 0 Then
                        BogID = RpRow.Record(0).Value '[BogID]
                        WebID = RpRow.Record(14).Value '[SubmisID]
                        BoNum = RpRow.Record(18).Value '[BogenNr]
                        If BoNum > 0 Then 'Patientenbogennummer

                            RS160.Filter = "[FormSubmissionId] Like '" & WebID & "'"
                            DoEvents
                            If RS160.RecordCount > 0 Then
                                Do
                                If RS160.Fields("QuestionNumber").Value <> vbNullString Then
                                    If IsNumeric(RS160.Fields("QuestionNumber").Value) = True Then
                                        FrgNr = CLng(RS160.Fields("QuestionNumber").Value)
                                    Else
                                        FrgNr = 0
                                    End If
                                Else
                                    FrgNr = 0
                                End If
                                If RS160.Fields("FormElementType").Value <> vbNullString Then
                                    If IsNumeric(RS160.Fields("FormElementType").Value) = True Then
                                        TypNr = CInt(RS160.Fields("FormElementType").Value)
                                    Else
                                        TypNr = 0
                                    End If
                                Else
                                    TypNr = 0
                                End If
                                If RS160.Fields("Value").Value <> vbNullString Then
                                    FrgWe = RS160.Fields("Value").Value
                                Else
                                    FrgWe = vbNullString
                                End If
                                If RS160.Fields("FormElementId").Value <> vbNullString Then
                                    FrgID = RS160.Fields("FormElementId").Value
                                Else
                                    FrgID = vbNullString
                                End If

                                If FrgNr > 23 Then 'WICHTIG!
                                    DBCmEx3 "qryPatAnVal", "@IdStr", "@BogNr", "@FrgNr", FrgWe, BoNum, FrgNr
                                    DoEvents
                                End If

                                RS160.MoveNext
                                Loop Until RS160.EOF
                                DoEvents
                            End If
                            RS160.Filter = adFilterNone

                        End If
                    End If
                End If
            End If
        End If
    Next RpRow
End If

For Each RpRow In RpRws
    If RpRow.GroupRow = False Then
        If RpRow.Record(19).CheckboxState = 0 Then 'Wenn Fragebogen nicht gelöscht werden soll
            If RpRow.Record(17).Value <> vbNullString Then
                PatNr = RpRow.Record(17).Value '[ID0]
                If PatNr > 0 Then
                    WebID = RpRow.Record(14).Value '[SubMisID]
                    BoNum = RpRow.Record(18).Value '[BogenNr]
                    ZipNa = GlTEx & WebID & ".zip" 'Termineordner
                    ImOrd = GlTEx & WebID & "\"
                    If RpRow.Record(15).CheckboxState = 1 Then 'Warteliste hinzuf?gen
                        Warte = True
                    Else
                        Warte = False
                    End If
                    If BoNum > 0 Then
                        BerTx = S_AnTex(BoNum)
                        DoEvents
                        If BerTx <> vbNullString Then
                            GlNeK = GlKoX
                            With GlNeK
                                .PatNr = PatNr
                                .IdxNr = 0
                                .EiDat = Format$(Date, "dd.mm.yyyy")
                                .EiZei = TimeValue(Now)
                                .EiTyp = 21 'Anamnese
                                .KoStr = BerTx
                                .NeuEi = True
                                .Mitar = GlMiA(GlSmI, 2)
                            End With
                            If K_Einf = False Then
                                If GlDbg = True Then MsgBox "Anamnese-Eintrag konnte nicht gespeichert werden", 48, "S_AnBoH"
                            End If
                        End If
                    End If
                    S_AnBoZ PatNr, ZipNa, ImOrd
                    DoEvents 'Dokument- und Bilderimport
                    If Warte = True Then
                        Ter_Edi PatNr, True
                        DoEvents
                        P_List "TeDe", 0, 2 'Warteliste aktualisieren
                    End If
                End If
            End If
        End If
    End If
Next RpRow

Set RpRws = Nothing
Set RpCo1 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoH " & Err.Number
Resume Next

End Sub
Public Sub S_AnBoL()
On Error GoTo SuErr
'Entfernt den Fragebogen vom Webserver

Dim BogNr As Long
Dim BogNa As String
Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim GuiKy As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim IniNa As String
Dim DaIni As String
Dim TmpSt As String
Dim TmZei As String
Dim DoLnk As String
Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim AryZe() As String

Set FM = frmMain
Set TrLi2 = FM.trvList2

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        DoEvents
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
End If

TeTit = "Fragebogen Weblink Entfernen"
TeMai = "Soll dieser Fragebogen jetzt vom Webserver entfernt werden?"
TeInh = "Durch das Entfernen dieses Fragebogens vom Webserver, wird dieser für Ihre Patienten nicht mehr aufrufbar sein."
TeFus = "Wurde dieser Fragebogen vom Webserver entfernt, kann dieser selbstverständlich jederzeit wieder neu publiziert werden."

BogNr = Mid$(GlNod, 2, Len(GlNod) - 1)

If GlCID <> vbNullString Then 'Cloud-ID
    PrNam = Chr$(34) & PrNam & Chr$(34)
    IniNa = CreateID("U") & ".ini"
    DaIni = GlTmp & IniNa

    If BogNr > 0 Then
        For AktZa = 1 To GlBoV
            If GlFrB(AktZa, 0) = BogNr Then
                If GlFrB(AktZa, 2) <> vbNullString Then
                    BogNa = GlFrB(AktZa, 1)
                    GuiKy = GlFrB(AktZa, 2)
                End If
                Exit For
            End If
        Next AktZa

        SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
        If GlMes = 33565 Then
        
            PaStr = "delete" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34)
            WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
            DoEvents

            With clFil
                If .FilVor(DaIni) = True Then
                    .FilPfa DaIni
                    TmpSt = .FilReSt
                    DoEvents
                    If TmpSt <> vbNullString Then
                        AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                        For AktZe = 0 To UBound(AryZe) - 1
                            If AryZe(AktZe) <> vbNullString Then
                                TmZei = AryZe(AktZe)
                                Lange = Len(TmZei)
                                Posit = InStr(1, TmZei, "=", 1)
                                If Posit > 0 Then
                                    InTyp = LCase(Left$(TmZei, Posit - 1))
                                    Select Case InTyp
                                    Case "deletesuccess": DoLnk = Right$(TmZei, Lange - Posit)
                                    End Select
                                End If
                            End If
                        Next AktZe
                    End If
                End If
                DoEvents

                GuiKy = CreateID("F")
                DBCmEx2 "qryKat10b", "@IdStr", "@IdxNr", GuiKy, BogNr
                DBCmEx3 "qryKat10c", "@IdStr", "@IdWeb", "@IdxNr", vbNullString, vbNullString, BogNr
                DoEvents
                If GlBoV > 0 Then 'Fragebogen vorhanden
                    For AktZa = 1 To GlBoV
                        If GlFrB(AktZa, 0) = BogNr Then
                            GlFrB(AktZa, 3) = vbNullString
                            GlFrB(AktZa, 4) = vbNullString
                            Exit For
                        End If
                    Next AktZa
                End If
                Set Knote = TrLi2.Nodes(GlNod)
                Knote.Text = "<TextBlock>" & BogNa & "<Run Foreground='Green' Text='" & vbNullString & "'/></TextBlock>"
                DoEvents
                If DoLnk <> vbNullString Then
                    If LCase(DoLnk) = "true" Then
                        SPopu "Fragebogen Entfernen", "Der Fragebogen wurde erfolgreich vom Webserver entfernt.", IC48_Information
                    End If
                End If
                
                If GlLog = False Then 'General Logging
                    .DaLoe = GlTmp & "*.ini" & vbNullChar
                    .FilLoe
                Else
                    Clipboard.Clear
                    Clipboard.SetText PrNam & Space$(1) & PaStr
                End If
            End With
            DoEvents
            
            GlNeK = GlKoX 'Protokolleintrag
            With GlNeK
                .PatNr = GlMan(GlSMa, 2)
                .IdxNr = 0
                .EiDat = Format$(Date, "dd.mm.yyyy")
                .EiZei = TimeValue(Now)
                .EiTyp = 104
                .TeStr = "Fragebogen entfernt " & BogNa
                .ZiStr = Format$(Now, "hh:mm") & " Uhr"
                .NeuEi = True
                .KeiAk = True
                .Mitar = GlMiA(GlSmI, 2)
            End With
            S_Prot
        End If
    End If
End If

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoL " & Err.Number
Resume Next

End Sub
Public Sub S_AnBoN()
On Error GoTo SuErr
'Fügt einen neuen Fragebogen hinzu

Dim SQL1 As String
Dim SQL3 As String
Dim BogNr As Long
Dim FraNr As Long
Dim SubNr As Long
Dim SorNr As Long
Dim BogID As String
Dim GuiKy As String
Dim FrgID As String
Dim TypNr As Integer
Dim GesZa As Integer

GuiKy = S_AnBoD(GlBoX.BoNum, "GuiID")

If GuiKy = vbNullString Then
    GuiKy = CreateID("F")
    DBCmEx2 "qryKat10b", "@IdStr", "@IdxNr", GuiKy, GlBoX.BoNum
Else
    If Left$(GuiKy, 1) <> "F" Then
        GuiKy = CreateID("F")
        DBCmEx2 "qryKat10b", "@IdStr", "@IdxNr", GuiKy, GlBoX.BoNum
    End If
End If

BogID = CreateID("F")

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatAnBoN WHERE ID5 = -1"
    SQL3 = "SELECT * FROM dbo.qryPatAnNeu WHERE ID1 = -1"
Else
    SQL1 = "SELECT * FROM qryPatAnBoN WHERE [ID5] = -1;"
    SQL3 = "SELECT * FROM qryPatAnNeu WHERE [ID1] = -1;"
End If

Set RS169 = New ADODB.Recordset
With RS169
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Options:=adCmdText
End With

If RS169.Supports(adAddNew) Then
    RS169.AddNew
    RS169.Fields("ID0").Value = GlBoX.PatNr
    RS169.Fields("ID3").Value = GlBoX.BoNum
    RS169.Fields("IDP").Value = GlBoX.BoMan
    RS169.Fields("Datum").Value = GlBoX.BoDat
    RS169.Fields("IDKurz").Value = GlBoX.KoTex
    RS169.Fields("GuiID").Value = BogID
    RS169.Fields("GuiKey").Value = GuiKy
    RS169.Update
End If
RS169.Close
Set RS169 = Nothing
DoEvents

BogNr = CLng(S_AnBoB(BogID, "ID5"))
DoEvents

If BogNr > 0 Then
    Set RS165 = New ADODB.Recordset
    RS165.CursorLocation = adUseClient
    Set RS165 = DBCmRe1("qryKat08I", "@IdxNr", GlBoX.BoNum)
    If RS165.RecordCount > 0 Then
        Set RS171 = New ADODB.Recordset
        With RS171
            .CursorLocation = adUseClient
            .Source = SQL3
            .ActiveConnection = DB1
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open Options:=adCmdText
        End With
        If RS171.Supports(adAddNew) Then
            Do
            SorNr = SorNr + 1
            FraNr = RS165.Fields("ID6").Value
            SubNr = RS165.Fields("IDSub").Value
            If RS165.Fields("GuiID").Value <> vbNullString Then
                If Left$(RS165.Fields("GuiID").Value, 1) = "F" Then
                    FrgID = RS165.Fields("GuiID").Value
                Else
                    FrgID = CreateID("F")
                    DBCmEx3 "qryKat08H", "@IdGui", "@IdStr", "@IdxNr", FrgID, FrgID, FraNr
                End If
            Else
                FrgID = CreateID("F")
                DBCmEx3 "qryKat08H", "@IdGui", "@IdStr", "@IdxNr", FrgID, FrgID, FraNr
            End If
            If RS165.Fields("Type").Value <> vbNullString Then
                TypNr = RS165.Fields("Type").Value
            Else
                TypNr = 1
            End If
            RS171.AddNew
            RS171.Fields("ID5").Value = BogNr
            RS171.Fields("ID0").Value = GlBoX.PatNr
            RS171.Fields("ID6").Value = FraNr
            RS171.Fields("Type").Value = TypNr
            RS171.Fields("GuiID").Value = FrgID
            RS171.Fields("Sorter").Value = SorNr
            If RS165.Fields("Pflicht").Value <> vbNullString Then
                RS171.Fields("Pflicht").Value = RS165.Fields("Pflicht").Value
            Else
                RS171.Fields("Pflicht").Value = 0
            End If
            RS171.Fields("VorCheck").Value = RS165.Fields("VorCheck").Value
            RS171.Fields("VorText").Value = RS165.Fields("VorText").Value
            If GlBoX.BoNum <> SubNr Then
                RS171.Fields("IDSub").Value = SubNr
                If SubNr > 0 Then
                    RS171.Fields("Subfeld").Value = -1
                End If
            End If
            RS165.MoveNext
            Loop Until RS165.EOF
            DoEvents
            RS171.UpdateBatch
        End If
        RS171.Close
        Set RS171 = Nothing
    End If
    RS165.Close
    Set RS165 = Nothing
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoN " & Err.Number
Resume Next

End Sub
Public Sub S_AnBoP()
On Error GoTo KoErr
'Befüllen des PropertyGrid

Dim BogNr As Long
Dim PatNr As Long
Dim FraNr As Long
Dim SubNr As Long
Dim AktZa As Long
Dim TmpDa As Date
Dim VorSt As String
Dim TmpWe As String
Dim BerTx As String
Dim Vorga As String
Dim ZeiZa As Integer
Dim TypNr As Integer
Dim TypSu As Integer
Dim Lange As Integer
Dim GesZa As Integer
Dim TmpBo As Boolean
Dim WeAry() As String

Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrIts As XtremePropertyGrid.PropertyGridItems
Dim PrBol As XtremePropertyGrid.PropertyGridItemBool
Dim PrOpt As XtremePropertyGrid.PropertyGridItemOption
Dim PrDat As XtremePropertyGrid.PropertyGridItemDate

Set FM = frmMain
Set TxDe3 = FM.txtDeta3
Set PrGr1 = FM.prpGrid1
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns
Set RpSel = RpCo5.SelectedRows
Set PrIts = PrGr1.Categories

TxDe3.Text = vbNullString

PrIts.Clear

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    Set RpCol = RpCls.Find(Bog_ID5)
    BogNr = RpRow.Record(RpCol.ItemIndex).Value
Else
    Set RpSel = Nothing
    Set RpCo5 = Nothing
    Set PrGr1 = Nothing
    Exit Sub
End If

PrIts.Clear

If BogNr > 0 Then
    Set RS166 = New ADODB.Recordset
    RS166.CursorLocation = adUseClient
    Set RS166 = DBCmRe1("qryPatAnNum", "@IdxNr", BogNr)
    If RS166.RecordCount > 0 Then
        Do
        Erase WeAry 'Wertearry Löschen
        If RS166.Fields("Vorgaben").Value <> vbNullString Then
            If IsNull(RS166.Fields("Vorgaben").Value) = False Then
                Vorga = LCase(RS166.Fields("Vorgaben").Value)
            Else
                Vorga = vbNullString
            End If
        Else
            Vorga = vbNullString
        End If
        If RS166.Fields("Type").Value <> vbNullString Then
            TypNr = CInt(RS166.Fields("Type").Value)
        Else
            TypNr = 1
        End If
        If RS166.Fields("ID0").Value > 0 Then
            PatNr = RS166.Fields("IDSub").Value
        Else
            PatNr = 0
        End If
        If RS166.Fields("ID6").Value > 0 Then
            FraNr = RS166.Fields("ID6").Value
        Else
            FraNr = 0
        End If
        If RS166.Fields("IDSub").Value > 0 Then
            SubNr = RS166.Fields("IDSub").Value
        Else
            SubNr = 0
        End If
    
        If PrIts.Count > 0 Then
            If PrKat.id <> RS166.Fields("IDG").Value Then
                Set PrKat = PrGr1.AddCategory(RS166.Fields("Gruppe").Value)
                PrKat.id = RS166.Fields("IDG").Value
                PrKat.Expandable = True
                PrKat.Expanded = GlExN
            End If
        Else
            If RS166.Fields("Gruppe").Value <> vbNullString Then
                Set PrKat = PrGr1.AddCategory(RS166.Fields("Gruppe").Value)
                PrKat.id = RS166.Fields("IDG").Value
                PrKat.Expandable = True
                PrKat.Expanded = GlExN
            End If
        End If
    
        Select Case TypNr
        Case 1: 'Textfeld
                If Vorga = "ja;nein" Then
                    If RS166.Fields("Wert").Value <> vbNullString Then
                        TmpWe = LCase(RS166.Fields("Wert").Value)
                        Select Case TmpWe
                        Case "ja": TmpBo = True
                        Case "nein": TmpBo = False
                        Case "yes": TmpBo = True
                        Case "no": TmpBo = False
                        Case "0": TmpBo = False
                        Case "1": TmpBo = True
                        Case "-1": TmpBo = True
                        Case "true": TmpBo = True
                        Case "false": TmpBo = False
                        Case Else:
                            If CBool(RS166.Fields("Wert").Value) = True Then
                                TmpBo = True
                            Else
                                TmpBo = False
                            End If
                        End Select
                    Else
                        TmpBo = False
                    End If
                    Set PrBol = PrKat.AddChildItem(PropertyItemBool, RS166.Fields("IDKurz").Value, TmpBo)
                    PrBol.id = RS166.Fields("ID1").Value
                    PrBol.CheckBoxStyle = True
                    PrBol.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                    If RS166.Fields("Pflicht").Value <> vbNullString Then
                        PrBol.CaptionMetrics.Font.Bold = CBool(RS166.Fields("Pflicht").Value)
                    End If
                Else
                    If RS166.Fields("Wert").Value <> vbNullString Then
                        TmpWe = RS166.Fields("Wert").Value
                    Else
                        If RS166.Fields("VorText").Value <> vbNullString Then
                            TmpWe = RS166.Fields("VorText").Value
                        Else
                            TmpWe = vbNullString
                        End If
                    End If
                    Set PrItm = PrKat.AddChildItem(PropertyItemString, RS166.Fields("IDKurz").Value, TmpWe)
                    PrItm.id = RS166.Fields("ID1").Value
                    PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                    If CBool(RS166.Fields("Selekt").Value) = False Then
                        PrItm.ValueMetrics.ForeColor = 8421504
                    End If
                    If RS166.Fields("Pflicht").Value <> vbNullString Then
                        PrItm.CaptionMetrics.Font.Bold = CBool(RS166.Fields("Pflicht").Value)
                    End If
                End If
    
                If Vorga <> vbNullString Then
                    If Vorga <> "ja;nein" Then
                        WeAry = Split(Vorga, ";")
                        PrItm.flags = ItemHasComboButton
                        For AktZa = 0 To UBound(WeAry)
                            PrItm.Constraints.Add WeAry(AktZa)
                        Next AktZa
                    End If
                Else
                    If RS166.Fields("Zeilen").Value <> vbNullString Then
                        ZeiZa = Val(RS166.Fields("Zeilen").Value)
                        If ZeiZa > 8 Then
                            ZeiZa = 8
                        ElseIf ZeiZa < 1 Then
                            ZeiZa = 1
                        End If
                    Else
                        If RS166.Fields("Kommentar").Value <> vbNullString Then
                            If IsNumeric(RS166.Fields("Kommentar").Value) Then
                                ZeiZa = Val(RS166.Fields("Kommentar").Value)
                                If ZeiZa > 3 Then
                                    ZeiZa = 3
                                ElseIf ZeiZa < 2 Then
                                    ZeiZa = 2
                                End If
                            Else
                                ZeiZa = 2
                            End If
                        Else
                            ZeiZa = 2
                        End If
                    End If
                    If ZeiZa > 1 Then
                        If ZeiZa > 2 Then
                            PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
                        Else
                            PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine
                        End If
                    End If
                    PrItm.MultiLinesCount = ZeiZa
                    If RS166.Fields("Lange").Value <> vbNullString Then
                        Lange = RS166.Fields("Lange").Value
                        PrItm.ValueMetrics.MaxLength = Lange
                    End If
                    If SubNr > 0 Then
                        Set RS167 = New ADODB.Recordset
                        RS167.CursorLocation = adUseClient
                        Set RS167 = DBCmRe2("qryPatAnIdx", "@IdxNr", "@BogNr", SubNr, BogNr)
                        If RS167.RecordCount > 0 Then
                            TypSu = CInt(RS167.Fields("Type").Value)
                            If RS167.Fields("Wert").Value = vbNullString Or IsNull(RS167.Fields("Wert").Value) = True Then
                                PrItm.Hidden = True
                            Else
                                TmpWe = LCase(RS167.Fields("Wert").Value)
                                Select Case TmpWe
                                Case "ja": TmpBo = True
                                Case "nein": TmpBo = False
                                Case "yes": TmpBo = True
                                Case "no": TmpBo = False
                                Case "0": TmpBo = False
                                Case "1": TmpBo = True
                                Case "-1": TmpBo = True
                                Case "true": TmpBo = True
                                Case "false": TmpBo = False
                                Case Else:
                                    If CBool(RS167.Fields("Wert").Value) = True Then
                                        TmpBo = True
                                    Else
                                        TmpBo = False
                                    End If
                                End Select
                                If TmpBo = False Then
                                    PrItm.Hidden = True
                                End If
                            End If
                        End If
                        RS167.Close
                        Set RS167 = Nothing
                    End If
                End If
        Case 2: 'Combobox
                Set RS168 = New ADODB.Recordset
                RS168.CursorLocation = adUseClient
                Set RS168 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                GesZa = RS168.RecordCount
                If GesZa > 0 Then
                    ReDim WeAry(GesZa)
                    AktZa = 1
                    Do
                    WeAry(AktZa) = RS168.Fields("IDKurz").Value
                    AktZa = AktZa + 1
                    RS168.MoveNext
                    Loop Until RS168.EOF
                End If
                RS168.Close
                Set RS168 = Nothing
                
                If RS166.Fields("Wert").Value <> vbNullString Then
                    TmpWe = RS166.Fields("Wert").Value
                Else
                    TmpWe = vbNullString
                End If
                Set PrItm = PrKat.AddChildItem(PropertyItemString, RS166.Fields("IDKurz").Value, TmpWe)
                PrItm.id = RS166.Fields("ID1").Value
                PrItm.flags = ItemHasComboButton
                If GesZa > 0 Then
                    For AktZa = 1 To UBound(WeAry)
                        PrItm.Constraints.Add WeAry(AktZa), AktZa
                    Next AktZa
                End If
                PrItm.ConstraintEdit = False
                PrItm.DropDownItemCount = GesZa
                PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                If RS166.Fields("Pflicht").Value <> vbNullString Then
                    PrItm.CaptionMetrics.Font.Bold = CBool(RS166.Fields("Pflicht").Value)
                End If
                If SubNr > 0 Then
                    PrItm.Tag = SubNr
                    Set RS167 = New ADODB.Recordset
                    RS167.CursorLocation = adUseClient
                    Set RS167 = DBCmRe2("qryPatAnIdx", "@IdxNr", "@BogNr", SubNr, BogNr)
                    If RS167.RecordCount > 0 Then
                        TypSu = CInt(RS167.Fields("Type").Value)
                        If RS167.Fields("Wert").Value = vbNullString Or IsNull(RS167.Fields("Wert").Value) = True Then
                            PrItm.Hidden = True
                        Else
                            TmpWe = LCase(RS167.Fields("Wert").Value)
                            Select Case TmpWe
                            Case "ja": TmpBo = True
                            Case "nein": TmpBo = False
                            Case "yes": TmpBo = True
                            Case "no": TmpBo = False
                            Case "0": TmpBo = False
                            Case "1": TmpBo = True
                            Case "-1": TmpBo = True
                            Case "true": TmpBo = True
                            Case "false": TmpBo = False
                            Case Else:
                                If CBool(RS167.Fields("Wert").Value) = True Then
                                    TmpBo = True
                                Else
                                    TmpBo = False
                                End If
                            End Select
                            If TmpBo = False Then
                                PrItm.Hidden = True
                            End If
                        End If
                    End If
                    RS167.Close
                    Set RS167 = Nothing
                End If
        Case 3: 'Checkbox
                If RS166.Fields("Wert").Value <> vbNullString Then
                    TmpWe = LCase(RS166.Fields("Wert").Value)
                    Select Case TmpWe
                    Case "ja": TmpBo = True
                    Case "nein": TmpBo = False
                    Case "yes": TmpBo = True
                    Case "no": TmpBo = False
                    Case "0": TmpBo = False
                    Case "1": TmpBo = True
                    Case "-1": TmpBo = True
                    Case "true": TmpBo = True
                    Case "false": TmpBo = False
                    Case Else:
                        If CBool(RS166.Fields("Wert").Value) = True Then
                            TmpBo = True
                        Else
                            TmpBo = False
                        End If
                    End Select
                Else
                    TmpBo = False
                End If
                Set PrBol = PrKat.AddChildItem(PropertyItemBool, RS166.Fields("IDKurz").Value, TmpBo)
                PrBol.id = RS166.Fields("ID1").Value
                PrBol.CheckBoxStyle = True
                PrBol.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                If RS166.Fields("Pflicht").Value <> vbNullString Then
                    PrBol.CaptionMetrics.Font.Bold = CBool(RS166.Fields("Pflicht").Value)
                End If
                If SubNr > 0 Then
                    PrBol.Tag = SubNr
                    Set RS167 = New ADODB.Recordset
                    RS167.CursorLocation = adUseClient
                    Set RS167 = DBCmRe2("qryPatAnIdx", "@IdxNr", "@BogNr", SubNr, BogNr)
                    If RS167.RecordCount > 0 Then
                        TypSu = CInt(RS167.Fields("Type").Value)
                        If RS167.Fields("Wert").Value = vbNullString Or IsNull(RS167.Fields("Wert").Value) = True Then
                            PrBol.Hidden = True
                        Else
                            TmpWe = LCase(RS167.Fields("Wert").Value)
                            Select Case TmpWe
                            Case "ja": TmpBo = True
                            Case "nein": TmpBo = False
                            Case "yes": TmpBo = True
                            Case "no": TmpBo = False
                            Case "0": TmpBo = False
                            Case "1": TmpBo = True
                            Case "-1": TmpBo = True
                            Case "true": TmpBo = True
                            Case "false": TmpBo = False
                            Case Else:
                                If CBool(RS167.Fields("Wert").Value) = True Then
                                    TmpBo = True
                                Else
                                    TmpBo = False
                                End If
                            End Select
                            If TmpBo = False Then
                                PrItm.Hidden = True
                            End If
                        End If
                    End If
                    RS167.Close
                    Set RS167 = Nothing
                Else
                    PrBol.Tag = FraNr
                End If
        Case 4: 'Radiobutton
                Set RS168 = New ADODB.Recordset
                RS168.CursorLocation = adUseClient
                Set RS168 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                GesZa = RS168.RecordCount
                If GesZa > 0 Then
                    ReDim WeAry(GesZa)
                    AktZa = 1
                    Do
                    WeAry(AktZa) = RS168.Fields("IDKurz").Value
                    AktZa = AktZa + 1
                    RS168.MoveNext
                    Loop Until RS168.EOF
                End If
                RS168.Close
                Set RS168 = Nothing
                
                If RS166.Fields("Wert").Value <> vbNullString Then
                    TmpWe = RS166.Fields("Wert").Value
                Else
                    TmpWe = vbNullString
                End If
                Set PrOpt = PrKat.AddChildItem(PropertyItemOption, RS166.Fields("IDKurz").Value, Val(TmpWe))
                PrOpt.id = RS166.Fields("ID1").Value
                If GesZa > 0 Then
                    For AktZa = 1 To UBound(WeAry)
                        PrOpt.Constraints.Add WeAry(AktZa), AktZa
                    Next AktZa
                End If
                PrOpt.MultiLinesCount = GesZa
                PrOpt.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                If RS166.Fields("Pflicht").Value <> vbNullString Then
                    PrOpt.CaptionMetrics.Font.Bold = CBool(RS166.Fields("Pflicht").Value)
                End If
                If SubNr > 0 Then
                    PrOpt.Tag = SubNr
                    Set RS167 = New ADODB.Recordset
                    RS167.CursorLocation = adUseClient
                    Set RS167 = DBCmRe2("qryPatAnIdx", "@IdxNr", "@BogNr", SubNr, BogNr)
                    If RS167.RecordCount > 0 Then
                        TypSu = CInt(RS167.Fields("Type").Value)
                        If RS167.Fields("Wert").Value = vbNullString Or IsNull(RS167.Fields("Wert").Value) = True Then
                            PrOpt.Hidden = True
                        Else
                            TmpWe = LCase(RS167.Fields("Wert").Value)
                            Select Case TmpWe
                            Case "ja": TmpBo = True
                            Case "nein": TmpBo = False
                            Case "yes": TmpBo = True
                            Case "no": TmpBo = False
                            Case "0": TmpBo = False
                            Case "1": TmpBo = True
                            Case "-1": TmpBo = True
                            Case "true": TmpBo = True
                            Case "false": TmpBo = False
                            Case Else:
                                If CBool(RS167.Fields("Wert").Value) = True Then
                                    TmpBo = True
                                Else
                                    TmpBo = False
                                End If
                            End Select
                            If TmpBo = False Then
                                PrItm.Hidden = True
                            End If
                        End If
                    End If
                    RS167.Close
                    Set RS167 = Nothing
                End If
        Case 5: 'Checkboxlist
                If RS166.Fields("Wert").Value <> vbNullString Then
                    TmpWe = RS166.Fields("Wert").Value
                Else
                    TmpWe = vbNullString
                End If
                Set PrOpt = PrKat.AddChildItem(PropertyItemOption, RS166.Fields("IDKurz").Value, Val(TmpWe))
                PrOpt.id = RS166.Fields("ID1").Value
                PrOpt.CheckBoxStyle = True

                Set RS168 = New ADODB.Recordset
                RS168.CursorLocation = adUseClient
                Set RS168 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                GesZa = RS168.RecordCount
                If GesZa > 0 Then
                    AktZa = 1
                    Do
                    PrOpt.Constraints.Add Space$(1) & RS168.Fields("IDKurz").Value, AktZa
                    AktZa = AktZa * 2
                    RS168.MoveNext
                    Loop Until RS168.EOF
                End If
                RS168.Close
                Set RS168 = Nothing
                PrOpt.MultiLinesCount = GesZa
                PrOpt.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                If RS166.Fields("Pflicht").Value <> vbNullString Then
                    PrOpt.CaptionMetrics.Font.Bold = CBool(RS166.Fields("Pflicht").Value)
                End If
                If SubNr > 0 Then
                    PrOpt.Tag = SubNr
                    Set RS167 = New ADODB.Recordset
                    RS167.CursorLocation = adUseClient
                    Set RS167 = DBCmRe2("qryPatAnIdx", "@IdxNr", "@BogNr", SubNr, BogNr)
                    If RS167.RecordCount > 0 Then
                        TypSu = CInt(RS167.Fields("Type").Value)
                        If RS167.Fields("Wert").Value = vbNullString Or IsNull(RS167.Fields("Wert").Value) = True Then
                            PrOpt.Hidden = True
                        Else
                            TmpWe = LCase(RS167.Fields("Wert").Value)
                            Select Case TmpWe
                            Case "ja": TmpBo = True
                            Case "nein": TmpBo = False
                            Case "yes": TmpBo = True
                            Case "no": TmpBo = False
                            Case "0": TmpBo = False
                            Case "1": TmpBo = True
                            Case "-1": TmpBo = True
                            Case "true": TmpBo = True
                            Case "false": TmpBo = False
                            Case Else:
                                If CBool(RS167.Fields("Wert").Value) = True Then
                                    TmpBo = True
                                Else
                                    TmpBo = False
                                End If
                            End Select
                            If TmpBo = False Then
                                PrItm.Hidden = True
                            End If
                        End If
                    End If
                    RS167.Close
                    Set RS167 = Nothing
                End If
        Case 6: 'Text
                Set PrItm = PrKat.AddChildItem(PropertyItemString, RS166.Fields("IDKurz").Value, vbNullString)
                PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                PrItm.id = RS166.Fields("ID1").Value
                PrItm.ReadOnly = True
                If SubNr > 0 Then
                    PrItm.Tag = SubNr
                    Set RS167 = New ADODB.Recordset
                    RS167.CursorLocation = adUseClient
                    Set RS167 = DBCmRe2("qryPatAnIdx", "@IdxNr", "@BogNr", SubNr, BogNr)
                    If RS167.RecordCount > 0 Then
                        TypSu = CInt(RS167.Fields("Type").Value)
                        If RS167.Fields("Wert").Value = vbNullString Or IsNull(RS167.Fields("Wert").Value) = True Then
                            PrItm.Hidden = True
                        Else
                            TmpWe = LCase(RS167.Fields("Wert").Value)
                            Select Case TmpWe
                            Case "ja": TmpBo = True
                            Case "nein": TmpBo = False
                            Case "yes": TmpBo = True
                            Case "no": TmpBo = False
                            Case "0": TmpBo = False
                            Case "1": TmpBo = True
                            Case "-1": TmpBo = True
                            Case "true": TmpBo = True
                            Case "false": TmpBo = False
                            Case Else:
                                If CBool(RS167.Fields("Wert").Value) = True Then
                                    TmpBo = True
                                Else
                                    TmpBo = False
                                End If
                            End Select
                            If TmpBo = False Then
                                PrItm.Hidden = True
                            End If
                        End If
                    End If
                    RS167.Close
                    Set RS167 = Nothing
                End If
        Case 7: 'Datefield
                If RS166.Fields("Wert").Value <> vbNullString Then
                    If IsDate(RS166.Fields("Wert").Value) = True Then
                        TmpDa = CDate(RS166.Fields("Wert").Value)
                    Else
                        TmpDa = Date
                    End If
                Else
                    TmpDa = Date
                End If
                Set PrDat = PrKat.AddChildItem(PropertyItemDate, RS166.Fields("IDKurz").Value, Format$(TmpDa, "dd.mm.yyyy"))
                PrDat.defaultValue = Format$(TmpDa, "dd.mm.yyyy")
                PrDat.Format = "%d.%m.%Y"
                PrDat.SetMask "00.00.0000", "__.__.____"
                PrDat.id = RS166.Fields("ID1").Value
                PrDat.CaptionMetrics.DrawTextFormat = DrawTextVcenter
                If CBool(RS166.Fields("Selekt").Value) = False Then
                    PrDat.ValueMetrics.ForeColor = 8421504
                End If
                If RS166.Fields("Pflicht").Value <> vbNullString Then
                    PrDat.CaptionMetrics.Font.Bold = CBool(RS166.Fields("Pflicht").Value)
                End If
                If SubNr > 0 Then
                    PrDat.Tag = SubNr
                    Set RS167 = New ADODB.Recordset
                    RS167.CursorLocation = adUseClient
                    Set RS167 = DBCmRe2("qryPatAnIdx", "@IdxNr", "@BogNr", SubNr, BogNr)
                    If RS167.RecordCount > 0 Then
                        If RS167.Fields("Wert").Value = vbNullString Or IsNull(RS167.Fields("Wert").Value) = True Then
                            PrDat.Hidden = True
                        Else
                            TmpWe = LCase(RS167.Fields("Wert").Value)
                            Select Case TmpWe
                            Case "ja": TmpBo = True
                            Case "nein": TmpBo = False
                            Case "yes": TmpBo = True
                            Case "no": TmpBo = False
                            Case "0": TmpBo = False
                            Case "1": TmpBo = True
                            Case "-1": TmpBo = True
                            Case "true": TmpBo = True
                            Case "false": TmpBo = False
                            Case Else:
                                If CBool(RS167.Fields("Wert").Value) = True Then
                                    TmpBo = True
                                Else
                                    TmpBo = False
                                End If
                            End Select
                            If TmpBo = False Then
                                PrDat.Hidden = True
                            End If
                        End If
                    End If
                    RS167.Close
                    Set RS167 = Nothing
                End If
        End Select
            
        RS166.MoveNext
        Loop Until RS166.EOF
    End If
    
    RS166.Close
    Set RS166 = Nothing

    If GlGrN = True Then
        PrGr1.PropertySort = Categorized
    Else
        PrGr1.PropertySort = NoSort
    End If
    
    If GlSoN = 1 Then 'Sortierung Fragebogen
        PrGr1.PropertySort = Alphabetical
    End If
End If

Set RpSel = Nothing
Set RpCo5 = Nothing
Set PrGr1 = Nothing

DoEvents
If BogNr > 0 Then
    BerTx = S_AnTex(BogNr)
    If BerTx <> vbNullString Then
        TxDe3.Text = BerTx
    End If
End If

Exit Sub

KoErr:
If GlDbg = True Then
    SErLog Err.Description & " S_AnBoP " & Err.Number
    MsgBox Err.Description, 48, "S_AnBoP " & Err.Number
End If
Resume Next

End Sub
Public Sub S_AnBoU(ByVal TmStr As String)
On Error GoTo SuErr
'Unlock Fragebogen nach Submission ID Fragebogen Anfordern

Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim IniNa As String
Dim DaIni As String
Dim TmpSt As String
Dim TmZei As String
Dim SbStr As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim RetWe As Integer
Dim AryZe() As String

Set FM = frmMain

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        DoEvents
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
End If

If GlCID <> vbNullString Then 'Cloud-ID
    If TmStr <> vbNullString Then
        SbStr = Replace(TmStr, Chr$(32), vbNullString, 1)
        PrNam = Chr$(34) & PrNam & Chr$(34)
        IniNa = CreateID("U") & ".ini"
        DaIni = GlTmp & IniNa

        PaStr = "unlock" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & SbStr & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents
    
        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "ReturnCode": RetWe = Val(Right$(TmZei, Lange - Posit))
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
            DoEvents

            TeTit = "Fragebogenanforderung"
            Select Case RetWe
            Case 0:
                TeMai = "Die Einreichungs-ID ist nicht gültig bzw. passt nicht zur Ihrer Anwender-ID"
                TeInh = "Die von Ihnen eingefügte Einreichungs-ID konnte entweder nicht auf dem Server gefunden werden oder ist nicht mit Ihrer Anwender-ID verknüpft."
                TeFus = "Der Fragebogen kann nicht zum erneuten Abruf bereitgestellt werden. Bitte prüfen Sie, ob die Einreichungs-ID korrekt ist."
            Case 1:
                TeMai = "Der Fragebogen konnte nicht bereitgestellt werden"
                TeInh = "Die von Ihnen eingefügte Einreichungs-ID ist gültig und konnte auf dem Server gefunden werden. Bei der erneuten Bereitstellung des Fragebogens gibt es aber ein technisches Problem."
                TeFus = "Der Fragebogen kann nicht zum erneuten Abruf bereitgestellt werden. Bitte wen Sie sich an den technischen Support."
            Case 2:
                TeMai = "Der Fragebogen wurde zum Abruf bereit gestell."
                TeInh = "Die von Ihnen eingefügte Einreichungs-ID ist gültig und kann nun erneut abgerufen werden."
                TeFus = "Bitte achten Sie darauf, dass erst alle zuvor abgerufene Fragebögen zugeordnet oder gelöscht werden müssen, bevor Sie diesen Fragebogen erneut abrufen können."
            End Select
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd

            If GlLog = False Then 'General Logging
                .DaLoe = GlTmp & "*.ini" & vbNullChar
                .FilLoe
            Else
                Clipboard.Clear
                Clipboard.SetText PrNam & Space$(1) & PaStr
            End If
        End With
        DoEvents
        
        GlNeK = GlKoX 'Protokolleintrag
        With GlNeK
            .PatNr = GlMan(GlSMa, 2)
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = 104
            .TeStr = "Fragebogenanforderung " & SbStr
            .ZiStr = Format$(Now, "hh:mm") & " Uhr"
            .NeuEi = True
            .KeiAk = True
            .Mitar = GlMiA(GlSmI, 2)
        End With
        S_Prot
    End If
End If

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoU " & Err.Number
Resume Next

End Sub

Public Function S_AnBoV(ByVal FrSet As Integer, ByVal BogNa As String, ByVal MaEma As String, ByVal MaBrf As String, ByVal DocPf As String, ByVal FrTit As String, ByVal FrTer As Integer, ByVal FrRed As Integer, ByVal FrUpl As Integer, ByVal FrRec As Integer) As String
On Error GoTo SuErr
'Veröffentlicht (Publiziert) den Fragebogen

Dim BogNr As Long
Dim FraNr As Long
Dim GruNr As Long
Dim SubNr As Long
Dim IdSub As Long
Dim SorZa As Long
Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim FiNam As String
Dim DaNam As String
Dim DaNaO As String
Dim GuiID As String
Dim SuGui As String
Dim GuiKy As String
Dim TmStr As String
Dim GruNa As String
Dim IniNa As String
Dim DaIni As String
Dim MaNam As String
Dim TmpSt As String
Dim TmZei As String
Dim FmLnk As String
Dim FrmID As String
Dim ErrSt As String
Dim VorTx As String
Dim BerTx As String
Dim FrgTx As String
Dim TypNr As Integer
Dim SuTyp As Integer
Dim GesZa As Integer
Dim SubZa As Integer
Dim SubAk As Integer
Dim PflFe As Integer
Dim VoChk As Integer
Dim Zeile As Integer
Dim ZeiBr As Integer
Dim ZeiLa As Integer
Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim RetWe As Boolean
Dim AryZe() As String

Set FM = frmMain
Set TrLi2 = FM.trvList2

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

Screen.MousePointer = vbHourglass
DoEvents

DaNam = CreateID("F") & ".csv"
FiNam = GlTEx & DaNam 'Termineordner

If Len(MaBrf) > 200 Then
    MaBrf = Left$(MaBrf, 200)
End If

If Len(BogNa) > 100 Then
    BogNa = Left$(BogNa, 100)
End If

BogNr = Mid$(GlNod, 2, Len(GlNod) - 1)

If FrTit = vbNullString Then
    If DocPf <> vbNullString Then
        If InStr(1, DocPf, Chr$(59), 1) = 0 Then 'Nur ein Dokument
            With clFil
                .FilPfa DocPf
                DaNaO = .DaNaO
            End With
            Lange = Len(DaNaO)
            FrTit = Right$(DaNaO, Lange - 34)
        End If
    End If
End If

If GlCID <> vbNullString Then 'Cloud-ID
    PrNam = Chr$(34) & PrNam & Chr$(34)
    IniNa = CreateID("U") & ".ini"
    DaIni = GlTmp & IniNa

    If BogNr > 0 Then
        GuiKy = S_AnBoD(BogNr, "GuiID")
        
        If GuiKy = vbNullString Then
            GuiKy = CreateID("F")
            DBCmEx2 "qryKat10b", "@IdStr", "@IdxNr", GuiKy, BogNr
        Else
            If Left$(GuiKy, 1) <> "F" Then
                GuiKy = CreateID("F")
                DBCmEx2 "qryKat10b", "@IdStr", "@IdxNr", GuiKy, BogNr
            End If
        End If

        TmStr = "Gruppennummer;Gruppenname;Fragennummer;FragenID;Fragentyp;Sortierzahl;Fragentext;Bericht;Vorlagentext;Pflichtfeld;Checkboxvorgabe;Zeilenzahl;Zeichenbreite;MaxZeichen;SubID" & vbCrLf

        Set RS165 = New ADODB.Recordset
        RS165.CursorLocation = adUseClient
        Set RS165 = DBCmRe1("qryKat08I", "@IdxNr", BogNr)
        GesZa = RS165.RecordCount
        If GesZa > 0 Then

            Do
            FraNr = RS165.Fields("ID6").Value
            GruNr = RS165.Fields("IDG").Value
            If IsNull(RS165.Fields("Gruppe").Value) = True Then
                GruNa = "01 Allgemein"
            ElseIf RS165.Fields("Gruppe").Value = "" Then
                GruNa = "01 Allgemein"
            Else
                GruNa = RS165.Fields("Gruppe").Value
            End If
            If RS165.Fields("GuiID").Value <> vbNullString Then
                If Left$(RS165.Fields("Feldname").Value, 1) = "F" Then
                    GuiID = RS165.Fields("GuiID").Value
                Else
                    GuiID = CreateID("F")
                    DBCmEx3 "qryKat08H", "@IdGui", "@IdStr", "@IdxNr", GuiID, GuiID, FraNr
                End If
            Else
                GuiID = CreateID("F")
                DBCmEx3 "qryKat08H", "@IdGui", "@IdStr", "@IdxNr", GuiID, GuiID, FraNr
            End If
            If RS165.Fields("Type").Value > 0 Then
                TypNr = RS165.Fields("Type").Value
            Else
                TypNr = 1
            End If
            If RS165.Fields("Sorter").Value > 0 Then
                SorZa = RS165.Fields("Sorter").Value
            Else
                SorZa = 1
            End If
            If RS165.Fields("Zeilen").Value <> vbNullString Then
                Zeile = RS165.Fields("Zeilen").Value
            Else
                Zeile = 2
            End If
            If RS165.Fields("Gross").Value <> vbNullString Then
                ZeiBr = RS165.Fields("Gross").Value
            Else
                ZeiBr = 20
            End If
            If RS165.Fields("Lange").Value <> vbNullString Then
                ZeiLa = RS165.Fields("Lange").Value
            Else
                ZeiLa = 250
            End If
            If CBool(RS165.Fields("Pflicht").Value) = True Then
                PflFe = 1
            Else
                PflFe = 0
            End If
            If CBool(RS165.Fields("VorCheck").Value) = True Then
                VoChk = 1
            Else
                VoChk = 0
            End If
            If RS165.Fields("VorText").Value <> vbNullString Then
                VorTx = RS165.Fields("VorText").Value
                VorTx = Replace$(VorTx, vbCrLf, "$$", 1)
                VorTx = SUmw(VorTx, False, False, True, False)
            Else
                VorTx = vbNullString
            End If
            If RS165.Fields("Bericht").Value <> vbNullString Then
                BerTx = RS165.Fields("Bericht").Value
                BerTx = Replace$(BerTx, vbCrLf, "$$", 1)
                BerTx = SUmw(BerTx, False, False, True, False)
            Else
                BerTx = "kein Berichttext"
            End If
            If RS165.Fields("IDKurz").Value <> vbNullString Then
                FrgTx = RS165.Fields("IDKurz").Value
                FrgTx = Replace$(FrgTx, vbCrLf, "$$", 1)
                FrgTx = SUmw(FrgTx, False, False, True, False)
            Else
                FrgTx = "kein Fragentext"
            End If
            If RS165.Fields("IDSub").Value <> vbNullString Then
                If RS165.Fields("IDSub").Value > 0 Then
                    IdSub = RS165.Fields("IDSub").Value
                Else
                    IdSub = 0
                End If
            Else
                IdSub = 0
            End If
            If FraNr = IdSub Then
                IdSub = 0
            End If

            TmStr = TmStr & Format$(GruNr, "0000") & ";" 'Gruppennummer
            TmStr = TmStr & GruNa & ";" 'Gruppenname
            TmStr = TmStr & Format$(FraNr, "0000") & ";" 'Fragennummer
            TmStr = TmStr & GuiID & ";" 'FrgaenGUI
            TmStr = TmStr & Format$(TypNr, "0000") & ";" 'Typnummer
            TmStr = TmStr & Format$(SorZa, "0000") & ";" 'Sortierzahl
            TmStr = TmStr & FrgTx & ";" 'Fragentext
            TmStr = TmStr & BerTx & ";" 'Bericht
            TmStr = TmStr & VorTx & ";" 'Vorlagentext
            TmStr = TmStr & Format$(PflFe, "0000") & ";" 'Pflichtfeld
            TmStr = TmStr & Format$(VoChk, "0000") & ";" 'Checkboxvorgabe
            TmStr = TmStr & Format$(Zeile, "0000") & ";" 'Zeilenzahl
            TmStr = TmStr & Format$(ZeiBr, "0000") & ";" 'Zeichenbreite
            TmStr = TmStr & Format$(ZeiLa, "0000") & ";" 'MaxZeichen
            TmStr = TmStr & Format$(IdSub, "0000") 'SubID
            TmStr = TmStr & vbCrLf

            If TypNr >= 2 And TypNr <= 5 Then
                Set RS172 = New ADODB.Recordset
                RS172.CursorLocation = adUseClient
                Set RS172 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                SubZa = RS172.RecordCount
                If SubZa > 0 Then
                    SubAk = 1
                    Do
                    SubNr = RS172.Fields("ID6").Value
                    If RS172.Fields("GuiID").Value <> vbNullString Then
                        If Left$(RS172.Fields("Feldname").Value, 1) = "S" Then
                            SuGui = RS172.Fields("GuiID").Value
                        Else
                            SuGui = CreateID("S")
                            DBCmEx3 "qryKat08H", "@IdGui", "@IdStr", "@IdxNr", SuGui, SuGui, SubNr
                        End If
                    Else
                        SuGui = CreateID("S")
                        DBCmEx3 "qryKat08H", "@IdGui", "@IdStr", "@IdxNr", SuGui, SuGui, SubNr
                    End If
                    If RS172.Fields("Type").Value > 0 Then
                        SuTyp = RS172.Fields("Type").Value
                    Else
                        SuTyp = 1
                    End If
                    If RS172.Fields("Sorter").Value > 0 Then
                        SorZa = RS172.Fields("Sorter").Value
                    Else
                        SorZa = 1
                    End If
                    If CBool(RS172.Fields("VorCheck").Value) = True Then
                        VoChk = 1
                    Else
                        VoChk = 0
                    End If
                    If RS172.Fields("Bericht").Value <> vbNullString Then
                        BerTx = RS172.Fields("Bericht").Value
                        BerTx = Replace$(BerTx, vbCrLf, "$$", 1)
                        BerTx = SUmw(BerTx, False, False, True, False)
                    Else
                        BerTx = "kein Berichttext"
                    End If
                    If RS172.Fields("IDKurz").Value <> vbNullString Then
                        FrgTx = RS172.Fields("IDKurz").Value
                        FrgTx = Replace$(FrgTx, vbCrLf, "$$", 1)
                        FrgTx = SUmw(FrgTx, False, False, True, False)
                    Else
                        FrgTx = "kein Fragentext"
                    End If
                    If RS172.Fields("IDSub").Value <> vbNullString Then
                        If RS172.Fields("IDSub").Value > 0 Then
                            IdSub = RS172.Fields("IDSub").Value
                        Else
                            IdSub = 0
                        End If
                    Else
                        IdSub = 0
                    End If
                    If SubNr = IdSub Then
                        IdSub = 0
                    End If
                    If SuTyp <> TypNr Then
                        DBCmEx2 "qryKat08K", "@IdTyp", "@IdxNr", TypNr, SubNr
                        SuTyp = TypNr
                    End If
                    If TypNr = 3 Then 'Ankreuzfeld
                        Select Case SubAk
                        Case 1: If LCase(FrgTx) <> "ja" Then FrgTx = "Ja"
                        Case 2: If LCase(FrgTx) <> "nein" Then FrgTx = "Nein"
                        End Select
                    End If

                    TmStr = TmStr & "0000" & ";" 'Gruppennummer
                    TmStr = TmStr & vbNullString & ";" 'Gruppenname
                    TmStr = TmStr & Format$(SubNr, "0000") & ";" 'Fragennummer
                    TmStr = TmStr & SuGui & ";" 'FrgaenGUI
                    TmStr = TmStr & "0000" & ";" 'Typnummer
                    TmStr = TmStr & Format$(SorZa, "0000") & ";" 'Sortierzahl
                    TmStr = TmStr & FrgTx & ";" 'Frage
                    TmStr = TmStr & BerTx & ";" 'Bericht
                    TmStr = TmStr & vbNullString & ";" 'Vorlagentext
                    TmStr = TmStr & "0000" & ";"  'Pflichtfeld
                    TmStr = TmStr & Format$(VoChk, "0000") & ";" 'Checkbox Vorlage
                    TmStr = TmStr & "0000" & ";"  'Zeilenzahl
                    TmStr = TmStr & "0000" & ";"  'Zeichenbreite
                    TmStr = TmStr & "0000" & ";"  'MaxZeichen
                    TmStr = TmStr & Format$(IdSub, "0000") 'SubID
                    TmStr = TmStr & vbCrLf
                    
                    SubAk = SubAk + 1
                    RS172.MoveNext
                    Loop Until RS172.EOF
                End If
                RS172.Close
                Set RS172 = Nothing
            End If
    
            RS165.MoveNext
            Loop Until RS165.EOF
                        
        End If
        RS165.Close
        Set RS165 = Nothing

        If FiNam <> vbNullString Then
            With clFil '-------- Upload --------
                If .FilVor(FiNam) = True Then
                    .DaLoe = FiNam & vbNullChar
                    .FilLoe
                End If
                .FilPfa FiNam
                .StrDa = TmStr
                RetWe = .FilWrSt
            End With
            DoEvents
        End If
    
        If DocPf <> vbNullString Then
            If GlOtL <> vbNullString Then 'Online-Terminbuchungs System Link für Datenschutzerklärung
                If GlOIm <> vbNullString Then 'Online-Terminbuchungs System Link für Impressum
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            Else
                If GlOIm <> vbNullString Then
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            End If
        Else
            If GlOtL <> vbNullString Then 'Online-Terminbuchungs System Link für Datenschutzerklärung
                If GlOIm <> vbNullString Then 'Online-Terminbuchungs System Link für Impressum
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            Else
                If GlOIm <> vbNullString Then
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            End If
        End If

        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If TmpSt <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "formurl": FrmID = Right$(TmZei, (Lange - Posit) - 1)
                                Case "completeformurl": FmLnk = Right$(TmZei, Lange - Posit)
                                Case "error": ErrSt = Right$(TmZei, Lange - Posit)
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
            DoEvents

            If ErrSt = vbNullString Then
                If GlLog = False Then 'General Logging
                    .DaLoe = GlTmp & "*.ini" & vbNullChar
                    .FilLoe
                    If FiNam <> vbNullString Then
                        If FmLnk <> vbNullString Then
                            If .FilVor(FiNam) = True Then
                                .DaLoe = FiNam & vbNullChar
                                .FilLoe
                            End If
                        End If
                    End If
                    If DocPf <> vbNullString Then
                        If .FilVor(DocPf) = True Then
                            .DaLoe = DocPf & vbNullChar
                            .FilLoe
                        End If
                    End If
                Else
                    Clipboard.Clear
                    Clipboard.SetText PrNam & Space$(1) & PaStr
                End If
            End If
        End With
        DoEvents

        If FmLnk <> vbNullString Then
            DBCmEx3 "qryKat10c", "@IdStr", "@IdWeb", "@IdxNr", FmLnk, FrmID, BogNr
            DoEvents
            If GlBoV > 0 Then 'Fragebogen vorhanden
                For AktZa = 1 To GlBoV
                    If GlFrB(AktZa, 0) = BogNr Then
                        GlFrB(AktZa, 3) = FmLnk
                        GlFrB(AktZa, 4) = FmLnk
                        Exit For
                    End If
                Next AktZa
            End If
            Set Knote = TrLi2.Nodes(GlNod)
            Knote.Text = "<TextBlock>" & BogNa & "<Run Foreground='Green' Text='" & "*" & "'/></TextBlock>"
            DoEvents
            Clipboard.Clear
            If GlLog = False Then 'General Logging
                Clipboard.SetText FmLnk
                S_AnBoV = FmLnk
            Else
                S_AnBoV = FmLnk & vbCrLf & vbCrLf & PrNam & Space$(1) & PaStr
            End If
        Else
            Clipboard.Clear
            Clipboard.SetText PrNam & Space$(1) & PaStr
            If ErrSt = vbNullString Then
                SPopu "Uploadfehler", "Unerwarteter Fehler, beim Hochladen des Fragebogens", IC48_Forbidden
            Else
                SPopu "Uploadfehler", ErrSt, IC48_Information
            End If
        End If
        DoEvents

        GlNeK = GlKoX 'Protokolleintrag
        With GlNeK
            .PatNr = GlMan(GlSMa, 2)
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = 104
            .TeStr = "Fragebogen publiziert " & BogNa & " " & DaNam
            .ZiStr = Format$(Now, "hh:mm") & " Uhr"
            .NeuEi = True
            .KeiAk = True
            .Mitar = GlMiA(GlSmI, 2)
        End With
        S_Prot
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoV " & Err.Number
Resume Next

End Function
Public Function S_AnBoY(ByVal BogNa As String, ByVal MaEma As String, ByVal MaBrf As String, ByVal DocPf As String, ByVal FrTit As String, ByVal FrTer As Integer, ByVal FrRed As Integer, ByVal FrUpl As Integer, ByVal FrRec As Integer) As String
On Error GoTo SuErr
'Veröffentlicht (Publiziert) das Neuaufnahmeformular

Dim MitNr As Long
Dim ManNr As Long
Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim SuGui As String
Dim GuiKy As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim IniNa As String
Dim DaIni As String
Dim MaNam As String
Dim DaNaO As String
Dim TmpSt As String
Dim TmZei As String
Dim FmLnk As String
Dim FrmID As String
Dim ErrSt As String
Dim KonZa As Integer
Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim RetWe As Boolean
Dim AryZe() As String

Const FrSet = 0

Set FM = frmMain
Set TrLi2 = FM.trvList2

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

Screen.MousePointer = vbHourglass
DoEvents

If GlCID <> vbNullString Then 'Cloud-ID
    GuiKy = "F" & Right$(GlCID, Len(GlCID) - 1)
Else
    GuiKy = CreateID("F")
End If

If Len(MaBrf) > 200 Then
    MaBrf = Left$(MaBrf, 200)
End If

If Len(BogNa) > 100 Then
    BogNa = Left$(BogNa, 100)
End If

If FrTit = vbNullString Then
    If DocPf <> vbNullString Then
        If InStr(1, DocPf, Chr$(59), 1) = 0 Then 'Nur ein Dokument
            With clFil
                .FilPfa DocPf
                DaNaO = .DaNaO
            End With
            Lange = Len(DaNaO)
            FrTit = Right$(DaNaO, Lange - 34)
        End If
    End If
End If

If GlCID <> vbNullString Then 'Cloud-ID
    PrNam = Chr$(34) & PrNam & Chr$(34)
    IniNa = CreateID("U") & ".ini"
    DaIni = GlTmp & IniNa

    If DocPf <> vbNullString Then
        If GlOtL <> vbNullString Then 'Online-Terminbuchungs System Link für Datenschutzerklärung
            If GlOIm <> vbNullString Then 'Online-Terminbuchungs System Link für Impressum
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            Else
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            End If
        Else
            If GlOIm <> vbNullString Then
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            Else
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & DocPf & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & FrTit & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            End If
        End If
    Else
        If GlOtL <> vbNullString Then 'Online-Terminbuchungs System Link für Datenschutzerklärung
            If GlOIm <> vbNullString Then 'Online-Terminbuchungs System Link für Impressum
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            Else
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            End If
        Else
            If GlOIm <> vbNullString Then
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            Else
                PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
            End If
        End If
    End If

    WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
    DoEvents

    With clFil
        If .FilVor(DaIni) = True Then
            .FilPfa DaIni
            TmpSt = .FilReSt
            DoEvents
            If TmpSt <> vbNullString Then
                AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                For AktZe = 0 To UBound(AryZe) - 1
                    If TmpSt <> vbNullString Then
                        TmZei = AryZe(AktZe)
                        Lange = Len(TmZei)
                        Posit = InStr(1, TmZei, "=", 1)
                        If Posit > 0 Then
                            InTyp = LCase(Left$(TmZei, Posit - 1))
                            Select Case InTyp
                            Case "formurl": FrmID = Right$(TmZei, (Lange - Posit) - 1)
                            Case "completeformurl": FmLnk = Right$(TmZei, Lange - Posit)
                            Case "error": ErrSt = Right$(TmZei, Lange - Posit)
                            End Select
                        End If
                    End If
                Next AktZe
            End If
        End If
        DoEvents

        If ErrSt = vbNullString Then
            If GlLog = False Then 'General Logging
                .DaLoe = GlTmp & "*.ini" & vbNullChar
                .FilLoe
                If DocPf <> vbNullString Then
                    If .FilVor(DocPf) = True Then
                        .DaLoe = DocPf & vbNullChar
                        .FilLoe
                    End If
                End If
            Else
                Clipboard.Clear
                Clipboard.SetText PrNam & Space$(1) & PaStr
            End If
        End If
    End With
    DoEvents

    If FmLnk <> vbNullString Then
        GlSet(1, 82) = FmLnk
        S_SeSe 83, FmLnk
        GlNaf = FmLnk
        DoEvents
        Clipboard.Clear
        If GlLog = True Then
            Clipboard.SetText PrNam & Space$(1) & PaStr
            S_AnBoY = FmLnk & vbCrLf & vbCrLf & PrNam & Space$(1) & PaStr
        Else
            Clipboard.SetText FmLnk
            S_AnBoY = FmLnk
        End If
    Else
        Clipboard.Clear
        Clipboard.SetText PrNam & Space$(1) & PaStr
        If ErrSt = vbNullString Then
            SPopu "Uploadfehler", "Unerwarteter Fehler, beim Hochladen des Neuaufnahmeformulars", IC48_Forbidden
        Else
            SPopu "Uploadfehler", ErrSt, IC48_Information
        End If
    End If
    DoEvents
    
    GlNeK = GlKoX 'Protokolleintrag
    With GlNeK
        .PatNr = GlMan(GlSMa, 2)
        .IdxNr = 0
        .EiDat = Format$(Date, "dd.mm.yyyy")
        .EiZei = TimeValue(Now)
        .EiTyp = 104
        .TeStr = "Fragebogen (Neuaufnahme) publiziert " & BogNa
        .ZiStr = Format$(Now, "hh:mm") & " Uhr"
        .NeuEi = True
        .KeiAk = True
        .Mitar = GlMiA(GlSmI, 2)
    End With
    S_Prot
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoY " & Err.Number
Resume Next

End Function
Private Sub S_AnBoZ(ByVal PatNr As Long, ByVal ZipNa As String, ByVal ImOrd As String)
On Error GoTo CrErr
'Dokument- und Bilderimport aus ZIP-Archiv

Dim DaExt As String
Dim DaNaO As String
Dim PaStr As String
Dim NeuNa As String
Dim NeuDa As String
Dim FiNam As String
Dim PfaNa As String
Dim TyNam As String
Dim AktZ1 As Integer
Dim AktZ2 As Integer
Dim EiTyp As Integer
Dim AnzDa As Integer
Dim ZipOk As Boolean
Dim DiNam() As String

If PatNr = 0 Then
    PatNr = GlAdr
End If

PaStr = "P" & Format$(PatNr, "000000")

Set clFil = New clsFile
clFil.hwnd = frmZuord.hwnd

If clFil.FilVor(ZipNa) = True Then
    If clFil.FilDir(ImOrd) = False Then
        MkDir ImOrd
        DoEvents
    End If
    If GlDbg = True Then
        ZipOk = SZipp(ZipNa, ImOrd, False, True)
    Else
        ZipOk = SZipp(ZipNa, ImOrd, True, True)
    End If
    DoEvents
    If ZipOk = True Then
        If clFil.FilVor(ImOrd & "*.*") = True Then
            AnzDa = clFil.FilLis(LCase(ImOrd), "*.*", DiNam)
            DoEvents
            If AnzDa > 0 Then
                For AktZ1 = 1 To UBound(DiNam)
                    FiNam = ImOrd & DiNam(AktZ1)
                    If SNaPr(FiNam) = False Then 'Dateinamensprüfung
                        clFil.FilPfa FiNam
                        DaExt = clFil.DaExt
                        DaNaO = clFil.DaNaO

                        Select Case LCase(DaExt)
                        Case "pdf":
                            TyNam = "PDF-Dokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "jpg":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "jpeg":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "png":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "psd":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "bmp":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "tif":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "tiff":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "gif":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "wmf":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "emf":
                            TyNam = "Bilddokument"
                            PfaNa = GlBPf
                            EiTyp = 105
                        Case "doc":
                            TyNam = "Textdokument"
                            PfaNa = GlBPf
                            EiTyp = 102
                        Case "docx":
                            TyNam = "Textdokument"
                            PfaNa = GlDox
                            EiTyp = 102
                        Case "rtf":
                            TyNam = "Textdokument"
                            PfaNa = GlDox
                            EiTyp = 102
                        Case "txm":
                            TyNam = "Textdokument"
                            PfaNa = GlDox
                            EiTyp = 24
                        Case "txn":
                            TyNam = "Textdokument"
                            PfaNa = GlDox
                            EiTyp = 24
                        Case "txr":
                            TyNam = "Textdokument"
                            PfaNa = GlDox
                            EiTyp = 24
                        End Select
                        DoEvents

                        For AktZ2 = 1 To UBound(GlFor) 'Importformate
                            If GlFor(AktZ2) = LCase(DaExt) Then

                                NeuDa = PaStr & "_" & DaNaO & "." & LCase(DaExt)
                                NeuNa = PfaNa & NeuDa

                                clFil.DaCop = FiNam & ";" & NeuNa & vbNullChar
                                If clFil.FilCop(2) = True Then
                                    DoEvents
                                    GlNeK = GlKoX
                                    With GlNeK
                                        .PatNr = PatNr
                                        .IdxNr = 0
                                        .EiDat = Format$(Date, "dd.mm.yyyy")
                                        .EiZei = TimeValue(Now)
                                        .EiTyp = EiTyp
                                        .KoStr = NeuDa
                                        .TeStr = TyNam
                                        .NeuEi = True
                                        .Mitar = GlMiA(GlSmI, 2)
                                    End With
                                    If K_Einf = False Then
                                        'DB-Eintrag fehlgeschlagen - kopierte Datei entfernen um Orphan zu verhindern
                                        On Error Resume Next
                                        If clFil.FilVor(NeuNa) = True Then
                                            clFil.DaLoe = NeuNa & vbNullChar
                                            clFil.FilLoe
                                        End If
                                        On Error GoTo CrErr
                                    End If
                                    DoEvents

                                End If
                            End If
                            DoEvents
                        Next AktZ2
                    Else
                        SPopu "Dateiname nicht lesbar", "Der von Ihnen ausgew?hlte Dateiname kann nicht gelesen werden weil er ggf. Sonderzeichen enthält", IC48_Warning
                    End If
                Next AktZ1
            End If
        End If
    End If
End If

Set clFil = Nothing

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnBoZ " & Err.Number
Exit Sub

End Sub
Public Sub S_AnSav(ByVal IdStr As String, ByVal IdxNr As Long)
On Error GoTo LiErr
'Speichert die Daten aus der Tabelle

DBCmEx2 "qryPatAnWer", "@IdStr", "@IdxNr", IdStr, IdxNr

If GlJet = True Then SDBSav

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnSav " & Err.Number
Resume Next

End Sub
Public Sub S_AnSel(ByVal SeTyp As Integer, ByVal ChVal As Boolean)
On Error GoTo PoErr
'Markierungen setzen

Dim IdxNr As Long
Dim RowNr As Long
Dim KrRow As Long
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrIts As XtremePropertyGrid.PropertyGridItems
Dim PrBol As XtremePropertyGrid.PropertyGridItemBool
Dim PrDat As XtremePropertyGrid.PropertyGridItemDate

Set FM = frmMain
Set RpCo5 = FM.repCont5
Set RpSel = RpCo5.SelectedRows
Set PrGr1 = FM.prpGrid1
Set PrIts = PrGr1.Categories

Select Case SeTyp
Case 1:
    For Each PrKat In PrIts
        For Each PrItm In PrKat.Childs
            If PrItm.Selected = True Then
                IdxNr = PrItm.id
                DBCmEx2 "qryPatAnSet", "@IdSet", "@IdxNr", ChVal, IdxNr
            End If
        Next PrItm
    Next PrKat
Case 2:
    For Each PrKat In PrIts
        For Each PrItm In PrKat.Childs
            IdxNr = PrItm.id
            DBCmEx2 "qryPatAnSet", "@IdSet", "@IdxNr", ChVal, IdxNr
        Next PrItm
    Next PrKat
End Select

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpAn RowNr
End If

Set RpCls = Nothing
Set RpRws = Nothing
Set RpCo5 = Nothing
Set PrGr1 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_AnSel " & Err.Number
Resume Next

End Sub
Public Sub S_AnSpl()
On Error GoTo LiErr

Dim AktZa As Integer
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns

Set RpCol = RpCls.Find(Bog_ID3)
If GlBoV > 0 Then 'Fragebogen vorhanden
    For AktZa = 1 To GlBoV
        RpCol.EditOptions.Constraints.Add GlFrB(AktZa, 1), GlFrB(AktZa, 0)
    Next AktZa
End If

Set RpCol = RpCls.Find(Bog_IDP)

For AktZa = 1 To UBound(GlMan)
    RpCol.EditOptions.Constraints.Add GlMan(AktZa, 1), GlMan(AktZa, 2)
Next AktZa

RpCo5.Redraw

Set RpCol = Nothing
Set RpCo5 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then SErLog Err.Description & " S_AnSpl " & Err.Number
Resume Next

End Sub
Public Function S_AnTex(ByVal IdxNr As Long) As String
On Error GoTo PoErr
'Generiert die Zusammenfassung des Fragebogens

Dim FraNr As Long
Dim TmpWe As Long
Dim BeStr As String
Dim WeStr As String
Dim GefWe As String
Dim TmpSt As String
Dim VoStr As String
Dim Gefun As Integer
Dim GeRev As Integer
Dim StaWe As Integer
Dim AktZa As Integer
Dim TypNr As Integer
Dim GesZa As Integer
Dim FrSel As Boolean
Dim TmpBo As Boolean
Dim BeTex() As String
Dim VoTex() As String

Set RS173 = New ADODB.Recordset
RS173.CursorLocation = adUseClient
Set RS173 = DBCmRe1("qryPatAnNum", "@IdxNr", IdxNr)
If RS173.RecordCount > 0 Then
    Do
    StaWe = 1
    If RS173.Fields("Wert").Value <> vbNullString Then
        
        If RS173.Fields("Type").Value <> vbNullString Then
            TypNr = CInt(RS173.Fields("Type").Value)
        Else
            TypNr = 1
        End If
        If RS173.Fields("ID6").Value > 0 Then
            FraNr = RS173.Fields("ID6").Value
        Else
            FraNr = 0
        End If
        If RS173.Fields("Bericht").Value <> vbNullString Then
            BeStr = RS173.Fields("Bericht").Value
        Else
            BeStr = vbNullString
        End If

        WeStr = LTrim$(RS173.Fields("Wert").Value)
        FrSel = CBool(RS173.Fields("Selekt").Value)

        If FrSel = True Then
            Select Case TypNr
            Case 1: 'Textfeld
            
                If RS173.Fields("Vorgaben").Value <> vbNullString Then
                    VoStr = RS173.Fields("Vorgaben").Value
                    VoTex = Split(VoStr, ";")
                    BeTex = Split(BeStr, ";")
                    For AktZa = 0 To UBound(VoTex)
                        If WeStr = "Wahr" Or WeStr = "True" Then
                            If UBound(BeTex) >= AktZa Then
                                TmpSt = TmpSt & BeTex(AktZa) & ". "
                            End If
                            Exit For
                        Else
                            If VoTex(AktZa) = WeStr Then
                                If UBound(BeTex) >= AktZa Then
                                    TmpSt = TmpSt & BeTex(AktZa) & ". "
                                End If
                                Exit For
                            End If
                        End If
                    Next AktZa
                Else
                    TmpSt = TmpSt & BeStr & Chr$(32) & WeStr & ". "
                End If

            Case 2: 'Combobox
                    
                Set RS168 = New ADODB.Recordset
                RS168.CursorLocation = adUseClient
                Set RS168 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                GesZa = RS168.RecordCount
                If GesZa > 0 Then
                    Do
                    If WeStr = RS168.Fields("IDKurz").Value Then
                        TmpSt = TmpSt & RS168.Fields("Bericht").Value & ". "
                        Exit Do
                    End If
                    RS168.MoveNext
                    Loop Until RS168.EOF
                End If
                RS168.Close
                Set RS168 = Nothing

            Case 3: 'Checkbox
            
                If CBool(WeStr) = True Then
                    TmpBo = True
                Else
                    TmpBo = False
                End If

                Set RS168 = New ADODB.Recordset
                RS168.CursorLocation = adUseClient
                Set RS168 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                GesZa = RS168.RecordCount
                If GesZa > 0 Then
                    If TmpBo = True Then
                        TmpSt = TmpSt & RS168.Fields("Bericht").Value & ". "
                    Else
                        RS168.MoveNext
                        TmpSt = TmpSt & RS168.Fields("Bericht").Value & ". "
                    End If
                End If
                RS168.Close
                Set RS168 = Nothing

            Case 4: 'Radiobutton
                If LCase(WeStr) <> "nein" Then
                    If LCase(WeStr) <> "ja" Then
                        Set RS168 = New ADODB.Recordset
                        RS168.CursorLocation = adUseClient
                        Set RS168 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                        GesZa = RS168.RecordCount
                        If GesZa > 0 Then
                            Do
                            If IsNumeric(WeStr) = True Then
                                If CInt(WeStr) = RS168.AbsolutePosition Then
                                    TmpSt = TmpSt & RS168.Fields("Bericht").Value & ". "
                                    Exit Do
                                End If
                            End If
                            RS168.MoveNext
                            Loop Until RS168.EOF
                        End If
                        RS168.Close
                        Set RS168 = Nothing
                    End If
                End If
            Case 5: 'Checkboxlist
                If LCase(WeStr) <> "nein" Then
                    If LCase(WeStr) <> "ja" Then
                        If IsNumeric(WeStr) = True Then
                            VoTex = SModw(Val(WeStr))
                        Else
                            VoTex = SModw(0)
                        End If
                        DoEvents
                        
                        If BeStr <> vbNullString Then
                            TmpSt = TmpSt & BeStr & Chr$(32)
                        End If
                        Set RS168 = New ADODB.Recordset
                        RS168.CursorLocation = adUseClient
                        Set RS168 = DBCmRe1("qryKat08S", "@IdxNr", FraNr)
                        GesZa = RS168.RecordCount
                        If GesZa > 0 Then
                            TmpWe = 1
                            Do
                            For AktZa = 0 To UBound(VoTex) - 1
                                If CLng(VoTex(AktZa)) = TmpWe Then
                                    If BeStr <> vbNullString Then
                                        If GesZa > 1 Then
                                            TmpSt = TmpSt & RS168.Fields("IDKUrz").Value & ", "
                                        Else
                                            TmpSt = TmpSt & RS168.Fields("IDKUrz").Value & ". "
                                        End If
                                    Else
                                        TmpSt = TmpSt & RS168.Fields("Bericht").Value & ". "
                                    End If
                                End If
                            Next AktZa
                            TmpWe = TmpWe * 2
                            RS168.MoveNext
                            Loop Until RS168.EOF
                            If Right$(TmpSt, 2) = ", " Then
                                TmpSt = Left$(TmpSt, Len(TmpSt) - 2) & ". "
                            End If
                        End If
                        RS168.Close
                        Set RS168 = Nothing
                    End If
                End If

            Case 6: 'Text
                
                TmpSt = TmpSt & RS173.Fields("IDKurz").Value & ". "
                
            Case 7: 'Datefield

                TmpSt = TmpSt & BeStr & Chr$(32) & Format$(WeStr, "dd.mm.yyyy") & ". "
                
            End Select
        End If
    End If
    RS173.MoveNext
    Loop Until RS173.EOF
End If
RS173.Close
Set RS173 = Nothing

S_AnTex = TmpSt

Exit Function

PoErr:
If GlDbg = True Then
    SErLog Err.Description & " S_AnTex " & Err.Number
    MsgBox Err.Description, 48, "S_AnTex " & Err.Number
End If
Resume Next

End Function
Public Sub S_Ary0()
On Error GoTo FiErr
'Fuellt die Daten in die Arrays

Dim SQL1 As String
Dim TmSt1 As String
Dim TmSt2 As String
Dim AktZa As Integer
Dim GesZa As Integer
Dim GesLe As Integer
Dim NeuLe As Integer

'-------------------------------------- Formulare ----------------------------------------------

' CRITICAL: Check DB1 connection before accessing
If DB1 Is Nothing Then
    SErLog "FATAL: Datenbankverbindung nicht initialisiert in S_Ary0 (DB1 ist Nothing)"
    Exit Sub
End If

If DB1.State <> adStateOpen Then
    SErLog "FATAL: Datenbankverbindung ist geschlossen in S_Ary0 (State=" & DB1.State & ", erwartet: 1)"
    Exit Sub
End If

AktZa = 0
If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryFormu ORDER BY ID1"
Else
    SQL1 = "SELECT * FROM qryFormu ORDER BY [ID1];"
End If

Set RS152 = New ADODB.Recordset 'Formulare
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount

DoEvents
If GesZa > 0 Then
    If GesZa < 105 Then
        SFoR2
        DoEvents
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx4 "qryFormAd", "@IdTit", "@IdDat", "@IdNam", "@IdAbs", GlFr2(0, AktZa), GlFr2(1, AktZa), GlFr2(2, AktZa), GlFr2(3, AktZa)
        Next AktZa
        ReDim GlFrm(4, GesZa + 4)
        For AktZa = 0 To 99
        If RS152.EOF Then
            Exit For
        End If
        GlFrm(0, AktZa) = RS152.Fields("IdTit").Value
        GlFrm(1, AktZa) = RS152.Fields("IdDat").Value
        GlFrm(2, AktZa) = RS152.Fields("IdNam").Value
        GlFrm(3, AktZa) = RS152.Fields("IdAbs").Value
        GlFrm(4, AktZa) = RS152.Fields("Selekt").Value
        RS152.MoveNext
        Next AktZa
        
        For AktZa = 0 To 4
        GlFrm(0, AktZa + 100) = GlFr2(0, AktZa)
        GlFrm(1, AktZa + 100) = GlFr2(1, AktZa)
        GlFrm(2, AktZa + 100) = GlFr2(2, AktZa)
        GlFrm(3, AktZa + 100) = GlFr2(3, AktZa)
        GlFrm(4, AktZa + 100) = GlFr2(4, AktZa)
        Next AktZa
    Else
        ReDim GlFrm(4, GesZa - 1) 'alle Formulare einlesen
        Do
        GlFrm(0, AktZa) = RS152.Fields("IdTit").Value
        GlFrm(1, AktZa) = RS152.Fields("IdDat").Value
        GlFrm(2, AktZa) = RS152.Fields("IdNam").Value
        GlFrm(3, AktZa) = RS152.Fields("IdAbs").Value
        GlFrm(4, AktZa) = RS152.Fields("Selekt").Value
        AktZa = AktZa + 1
        RS152.MoveNext
        Loop Until RS152.EOF
    End If
Else
    ReDim GlFrm(4, 105) 'WICHTIG!

    For AktZa = 0 To 105 'WICHTIG!
        SFoR1 AktZa, False
    Next AktZa
    DoEvents
    
    For AktZa = 0 To 105 'WICHTIG!
        DBCmEx4 "qryFormAd", "@IdTit", "@IdDat", "@IdNam", "@IdAbs", GlFrm(0, AktZa), GlFrm(1, AktZa), GlFrm(2, AktZa), GlFrm(3, AktZa)
    Next AktZa
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'------------------------------------------ Setup -----------------------------------------------

AktZa = 0
If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qrySetup ORDER BY ID1"
Else
    SQL1 = "SELECT * FROM qrySetup ORDER BY [ID1];"
End If

Set RS153 = New ADODB.Recordset 'Einstellungen
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount

DoEvents
If GesZa > 0 Then
    If GesZa < 20 Then

        If GesZa = 19 Then 'bestehenden Wert anpassen
            S_SeSe 19, , , , False
            S_SeSe 19, , 1
        ElseIf GesZa = 18 Then 'eine fehlende Einstellung hinzufï¿½gen
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", "1108#Standard-Laborkatalog", vbNullString, 1, 0, 0
        End If
        DoEvents
        
        SSet02
        DoEvents
        SSet03
        DoEvents
        SSet04
        DoEvents
        SSet05
        DoEvents
        SSet06
        DoEvents
        SSet07
        DoEvents
        SSet08
        DoEvents
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents

        For AktZa = 0 To 10 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS02(0, AktZa)), CStr(GlS02(1, AktZa)), CLng(GlS02(2, AktZa)), CSng(GlS02(3, AktZa)), CBool(GlS02(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS03(0, AktZa)), CStr(GlS03(1, AktZa)), CLng(GlS03(2, AktZa)), CSng(GlS03(3, AktZa)), CBool(GlS03(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 23 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS04(0, AktZa)), CStr(GlS04(1, AktZa)), CLng(GlS04(2, AktZa)), CSng(GlS04(3, AktZa)), CBool(GlS04(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS05(0, AktZa)), CStr(GlS05(1, AktZa)), CLng(GlS05(2, AktZa)), CSng(GlS05(3, AktZa)), CBool(GlS05(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS06(0, AktZa)), CStr(GlS06(1, AktZa)), CLng(GlS06(2, AktZa)), CSng(GlS06(3, AktZa)), CBool(GlS06(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS07(0, AktZa)), CStr(GlS07(1, AktZa)), CLng(GlS07(2, AktZa)), CSng(GlS07(3, AktZa)), CBool(GlS07(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS08(0, AktZa)), CStr(GlS08(1, AktZa)), CLng(GlS08(2, AktZa)), CSng(GlS08(3, AktZa)), CBool(GlS08(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents
        
        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing
                
    ElseIf GesZa < 40 Then
    
        SSet03
        DoEvents
        SSet04
        DoEvents
        SSet05
        DoEvents
        SSet06
        DoEvents
        SSet07
        DoEvents
        SSet08
        DoEvents
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS03(0, AktZa)), CStr(GlS03(1, AktZa)), CLng(GlS03(2, AktZa)), CSng(GlS03(3, AktZa)), CBool(GlS03(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 23 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS04(0, AktZa)), CStr(GlS04(1, AktZa)), CLng(GlS04(2, AktZa)), CSng(GlS04(3, AktZa)), CBool(GlS04(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS05(0, AktZa)), CStr(GlS05(1, AktZa)), CLng(GlS05(2, AktZa)), CSng(GlS05(3, AktZa)), CBool(GlS05(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS06(0, AktZa)), CStr(GlS06(1, AktZa)), CLng(GlS06(2, AktZa)), CSng(GlS06(3, AktZa)), CBool(GlS06(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS07(0, AktZa)), CStr(GlS07(1, AktZa)), CLng(GlS07(2, AktZa)), CSng(GlS07(3, AktZa)), CBool(GlS07(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS08(0, AktZa)), CStr(GlS08(1, AktZa)), CLng(GlS08(2, AktZa)), CSng(GlS08(3, AktZa)), CBool(GlS08(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents

        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing
        
    ElseIf GesZa < 64 Then
            
        SSet04
        DoEvents
        SSet05
        DoEvents
        SSet06
        DoEvents
        SSet07
        DoEvents
        SSet08
        DoEvents
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents

        For AktZa = 0 To 23 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS04(0, AktZa)), CStr(GlS04(1, AktZa)), CLng(GlS04(2, AktZa)), CSng(GlS04(3, AktZa)), CBool(GlS04(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS05(0, AktZa)), CStr(GlS05(1, AktZa)), CLng(GlS05(2, AktZa)), CSng(GlS05(3, AktZa)), CBool(GlS05(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS06(0, AktZa)), CStr(GlS06(1, AktZa)), CLng(GlS06(2, AktZa)), CSng(GlS06(3, AktZa)), CBool(GlS06(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS07(0, AktZa)), CStr(GlS07(1, AktZa)), CLng(GlS07(2, AktZa)), CSng(GlS07(3, AktZa)), CBool(GlS07(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS08(0, AktZa)), CStr(GlS08(1, AktZa)), CLng(GlS08(2, AktZa)), CSng(GlS08(3, AktZa)), CBool(GlS08(4, AktZa))
        Next AktZa
        DoEvents

        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 3 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents

        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing

    ElseIf GesZa < 70 Then
    
        SSet05
        DoEvents
        SSet06
        DoEvents
        SSet07
        DoEvents
        SSet08
        DoEvents
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS05(0, AktZa)), CStr(GlS05(1, AktZa)), CLng(GlS05(2, AktZa)), CSng(GlS05(3, AktZa)), CBool(GlS05(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS06(0, AktZa)), CStr(GlS06(1, AktZa)), CLng(GlS06(2, AktZa)), CSng(GlS06(3, AktZa)), CBool(GlS06(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS07(0, AktZa)), CStr(GlS07(1, AktZa)), CLng(GlS07(2, AktZa)), CSng(GlS07(3, AktZa)), CBool(GlS07(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS08(0, AktZa)), CStr(GlS08(1, AktZa)), CLng(GlS08(2, AktZa)), CSng(GlS08(3, AktZa)), CBool(GlS08(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 3 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents
        
        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing
        
    ElseIf GesZa < 75 Then
    
        SSet06
        DoEvents
        SSet07
        DoEvents
        SSet08
        DoEvents
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS06(0, AktZa)), CStr(GlS06(1, AktZa)), CLng(GlS06(2, AktZa)), CSng(GlS06(3, AktZa)), CBool(GlS06(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS07(0, AktZa)), CStr(GlS07(1, AktZa)), CLng(GlS07(2, AktZa)), CSng(GlS07(3, AktZa)), CBool(GlS07(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS08(0, AktZa)), CStr(GlS08(1, AktZa)), CLng(GlS08(2, AktZa)), CSng(GlS08(3, AktZa)), CBool(GlS08(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 3 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents
    
        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing

    ElseIf GesZa < 80 Then
    
        SSet07
        DoEvents
        SSet08
        DoEvents
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents
        
        For AktZa = 0 To 4 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS07(0, AktZa)), CStr(GlS07(1, AktZa)), CLng(GlS07(2, AktZa)), CSng(GlS07(3, AktZa)), CBool(GlS07(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS08(0, AktZa)), CStr(GlS08(1, AktZa)), CLng(GlS08(2, AktZa)), CSng(GlS08(3, AktZa)), CBool(GlS08(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 3 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents
    
        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing
        
    ElseIf GesZa < 90 Then
    
        SSet08
        DoEvents
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents

        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS08(0, AktZa)), CStr(GlS08(1, AktZa)), CLng(GlS08(2, AktZa)), CSng(GlS08(3, AktZa)), CBool(GlS08(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 3 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents
        
        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing
        
    ElseIf GesZa < 96 Then
    
        SSet09
        DoEvents
        SSet10
        DoEvents
        SSet11
        DoEvents

        For AktZa = 0 To 5 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS09(0, AktZa)), CStr(GlS09(1, AktZa)), CLng(GlS09(2, AktZa)), CSng(GlS09(3, AktZa)), CBool(GlS09(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 3 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents

        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing

    ElseIf GesZa < 100 Then

        SSet10
        DoEvents
        SSet11
        DoEvents
                
        For AktZa = 0 To 3 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS10(0, AktZa)), CStr(GlS10(1, AktZa)), CLng(GlS10(2, AktZa)), CSng(GlS10(3, AktZa)), CBool(GlS10(4, AktZa))
        Next AktZa
        DoEvents
        
        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents

        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing

    ElseIf GesZa < 110 Then

        SSet11
        DoEvents

        For AktZa = 0 To 9 'WICHTIG!
            DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlS11(0, AktZa)), CStr(GlS11(1, AktZa)), CLng(GlS11(2, AktZa)), CSng(GlS11(3, AktZa)), CBool(GlS11(4, AktZa))
        Next AktZa
        DoEvents
        
        AktZa = 0
        Set RS154 = New ADODB.Recordset 'Read Setup
        With RS154
            .CursorLocation = adUseClient
            .Source = SQL1
            .ActiveConnection = DB1
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open Options:=adCmdText
        End With
        GesZa = RS154.RecordCount
        DoEvents
        ReDim GlSet(5, 110) 'WICHTIG!
        Do 'alle vorhandenen Einstellungen durchlaufen
        If RS154.Fields("IDKurz").Value <> vbNullString Then
            GlSet(0, AktZa) = RS154.Fields("IDKurz").Value
        End If
        If RS154.Fields("SetTex").Value <> vbNullString Then
            GlSet(1, AktZa) = RS154.Fields("SetTex").Value
        End If
        GlSet(2, AktZa) = RS154.Fields("SetInt").Value
        GlSet(3, AktZa) = RS154.Fields("SetDec").Value
        GlSet(4, AktZa) = RS154.Fields("SetBit").Value
        AktZa = AktZa + 1
        RS154.MoveNext
        Loop Until RS154.EOF
        RS154.Close
        Set RS154 = Nothing
        
    ElseIf GesZa = 110 Then
        
        AktZa = 0
        ReDim GlSet(5, 110) 'WICHTIG!

        ' Ensure RS153 is positioned at first record
        If Not RS153.EOF And Not RS153.BOF Then
            RS153.MoveFirst
        End If

        If Not RS153.EOF Then
            Do 'alle vorhandenen Einstellungen durchlaufen
            If RS153.Fields("IDKurz").Value <> vbNullString Then
                GlSet(0, AktZa) = RS153.Fields("IDKurz").Value
            End If
            If RS153.Fields("SetTex").Value <> vbNullString Then
                GlSet(1, AktZa) = RS153.Fields("SetTex").Value
            End If
            GlSet(2, AktZa) = RS153.Fields("SetInt").Value
            GlSet(3, AktZa) = RS153.Fields("SetDec").Value
            GlSet(4, AktZa) = RS153.Fields("SetBit").Value
            AktZa = AktZa + 1
            RS153.MoveNext
            Loop Until RS153.EOF
        End If
        
    End If
Else
    SSet01 'keine zentralen Einstellungen vorhanden
    DoEvents
    For AktZa = 0 To 109 'WICHTIG!
        DBCmEx5 "qrySetAd", "@IdStr", "@SeTex", "@SeInt", "@SeRea", "@SeBit", CStr(GlSet(0, AktZa)), CStr(GlSet(1, AktZa)), CLng(GlSet(2, AktZa)), CSng(GlSet(3, AktZa)), CBool(GlSet(4, AktZa))
    Next AktZa
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'----------------------------------------------------------------------------------------------------

GlStK = CInt(GlSet(2, 0))               'Standardgebï¿½hrenkatalog
GlKe1 = CLng(GlSet(2, 1))               'Standardgebï¿½hrenkette 1
GlKop = CInt(GlSet(2, 2))               'Anzahl Rechnungsausdrucke
If Len(GlSet(1, 3)) > 0 Then
    GlRFm = Right$(GlSet(1, 3), 1)      'Rechnungsnummernformatierung
Else
    GlRFm = "0"
End If
GlReN = CBool(GlSet(4, 4))              'Rechnungsnummern sofort erzeugen
GlRnm = CBool(GlSet(4, 5))              'Rechnungsnummer am Jahrenanfang wieder mit 1 afangen
GlTSH = CBool(GlSet(4, 6))              'DATEV SOLL / HABEN Tausch
GlBGe = CBool(GlSet(4, 7))              'Getrennter Geldkonten Belegnummernkreis
If Len(GlSet(1, 8)) > 0 Then
    GlAdK = Right$(GlSet(1, 8), 1)      'Format der Adressenverkehrsnamens
Else
    GlAdK = "0"
End If
GlRMa = CBool(GlSet(4, 10))             'getrennter Mandentenrechnungsnummernkreis
GlBMa = CBool(GlSet(4, 11))             'Getrennter Mandanten Belegnummernkreis
GlOTS = CBool(GlSet(4, 12))             'Online-Terminbuchungs System aktivieren
GlODi = CBool(GlSet(4, 13))             'Online-Terminbuchungs System Mitarbeiterwahl
GlOTU = CStr(GlSet(1, 14))              'Online-Terminbuchungs Sytem Username
GlOTP = SCrypt(GlSet(1, 15), False)     'Online-Terminbuchungs Sytem Password
GlSeN = CStr(GlSet(1, 16))              'Online-Terminbuchungs Sytem Servername
GlESy = CBool(GlSet(4, 17))             'CalDAV / CardDAV / Exchange Synchronisation
GlStL = CInt(GlSet(2, 18))              'Standardlaborkatalog
If Len(GlSet(1, 19)) > 0 Then
    GlStD = Right$(GlSet(1, 19), 1)     'Standraddezimaltrennzeichen
Else
    GlStD = "0"
End If
GlOtD = CBool(GlSet(4, 20))             'Online-Terminbuchungs System Stornodialog
GlStZ = CInt(GlSet(2, 21))              'Standardzahlungsziel
GlSpl = CBool(GlSet(4, 22))             'Steuerspalte
If Len(GlSet(1, 23)) > 0 Then
    GlKtR = Right$(GlSet(1, 23), 1)     'Standardkontenrahmen
Else
    GlKtR = "0"
End If
GlGkB = CInt(GlSet(2, 26))              'Standardgeldkonto (Bank)
GlGkK = CInt(GlSet(2, 27))              'Standardgeldkonto (Kasse)
GlSSt = CBool(GlSet(4, 28))             'Starre Termintaktung
GlMPl = CBool(GlSet(4, 29))             'Mitarbeiterplan anstelle von Mandantenplam
GlReT = CStr(GlSet(1, 30))              'Standard-Belegtyp
GlStW = CInt(GlSet(2, 31))              'Standardwï¿½hrung
GlKnF = CBool(GlSet(4, 32))             'Sachkontenformatierung sechsstellig
GlPzn = CBool(GlSet(4, 33))             'PZN Einfï¿½gen
GlMVo = CBool(GlSet(4, 34))             'mandantenbezogene Vorgaben verwenden
GlRSo = CBool(GlSet(4, 35))             'Die Raumzuordnung numerisch sortiert anzeigen
GlAcI = CStr(GlSet(1, 36))              'SMS Account-ID
GlAbs = CStr(GlSet(1, 37))              'SMS Absenderkennung
GlOIm = CStr(GlSet(1, 38))              'Online-Terminbuchungs System Link fï¿½r Impressum
GlTok = CStr(GlSet(1, 39))              'SMS Produkt Token
GlPin = CBool(GlSet(4, 40))             'Online-Terminbuchungs System PIN
GlPxV = CBool(GlSet(4, 41))             'Proxyserver verwenden
GlPxN = CStr(GlSet(1, 42))              'Proxyserver Name
GlGbK = CBool(GlSet(4, 43))             'PAD Gebï¿½hrenkatalog benennen
GlIgL = CBool(GlSet(4, 44))             'Keine Preisberechnung bei IgL
GlPvM = CBool(GlSet(4, 45))             'Keine Positionskennzeichen bei Medikamenten und Begrï¿½ndungen
GlSpB = CBool(GlSet(4, 46))             'Umsatzsteuer Splittbuchungen
GldKt = CBool(GlSet(4, 47))             'DATEV Sachkonten vierstellig
GlDvB = CLng(GlSet(2, 48))              'DATEV Beraternummer
GlDvM = CLng(GlSet(2, 49))              'DATEV Mandantennummer
GlMaR = CStr(GlSet(1, 50))              'Mandant neue(s) Rechnung/Rezept
GlTeZ = CBool(GlSet(4, 51))             'Terminzeit aus dem Terminbetreff verwenden
GlBel = CBool(GlSet(4, 52))             'Online-Terminbuchungs System zeige belegte Buchungszeiten
GlOIC = CBool(GlSet(4, 53))             'Online-Terminbuchungs System ICS Datei
GlGWF = CStr(GlSet(1, 54))              'Online-Terminbuchungs System Google Web Font
GlOTr = CLng(GlSet(2, 55))              'Online-Terminbuchungs System allgemeine Textfarbe
GlOHF = CLng(GlSet(2, 56))              'Online-Terminbuchungs System allgemeine Hintergrundfarbe
GlOTG = CLng(GlSet(2, 57))              'Online-Terminbuchungs System allgemeine Textgrï¿½ï¿½e
GlOBH = CLng(GlSet(2, 58))              'Online-Terminbuchungs System Button Hintergrundfarbe
GlOBT = CLng(GlSet(2, 59))              'Online-Terminbuchungs System Button Textfarbe
GlOBO = CLng(GlSet(2, 60))              'Online-Terminbuchungs System Button Hooverfarbe
GlOBD = CLng(GlSet(2, 61))              'Online-Terminbuchungs System Button Deaktiviertfarbe
GlOSe = CStr(GlSet(1, 62))              'Online-Terminbuchungs System Link Anschlussseite
GlOtL = CStr(GlSet(1, 63))              'Online-Terminbuchungs System Link Datenschutzerklï¿½rung
GlStS = CInt(GlSet(2, 64))              'Standard-Steuersatz
GlRst = CBool(GlSet(4, 65))             'Mandantenbezogene Datenbegrenzung
GlIFo = CStr(GlSet(1, 66))              'LDT Import-Zeichensatz
GlLaF = CStr(GlSet(3, 67))              'Steigerungsfaktor Laborparameter
GlICS = CBool(GlSet(4, 68))             'Terminnachricht mit ICS Dateiversand
GlTeO = CBool(GlSet(4, 69))             'Mitarbeitername in Terminort speichern
GlEKr = CBool(GlSet(4, 70))             'Dokumentiert Emails in Krankenblatt
GlGDM = CBool(GlSet(4, 71))             'Mitarbeiternummer als GDT-Dateiname
GlGDD = CBool(GlSet(4, 72))             'GDT-Speicherung ohne Speichern-Dialog
GlGDn = CStr(GlSet(1, 73))              'Dateiname der GDT Exportdatei
GlKe2 = CLng(GlSet(2, 74))              'Standardgebï¿½hrenkette 2
GlSpT = CBool(GlSet(4, 75))             'Starre oder flexible Sprechzeiten verwenden
GlBuc = CBool(GlSet(4, 76))             'einfache Buchfï¿½hrung verwenden
GlNoM = CBool(GlSet(4, 77))             'Datenbankscripting aktivieren
GlSKo = CLng(GlSet(1, 78))              'Standardsteuerkonto
GlOTA = CBool(GlSet(4, 79))             'Online-Terminbuchungs System Adressenerfassung
GlTeB = CBool(GlSet(4, 80))             'Terminnachricht auch an BCC
GlTlA = CBool(GlSet(4, 81))             'Terminland WebCAL (ICS) aktivieren
GlNaf = CStr(GlSet(1, 82))              'Neuaufnahmeformular-Webadresse
GlTlB = CStr(GlSet(1, 83))              'Terminland Benutzername
GlTlP = SCrypt(GlSet(1, 84), False)     'Terminland Password
GlDoL = CBool(GlSet(4, 85))             'Dokument nach einfï¿½gen lï¿½schen
GlKrS = CBool(GlSet(4, 86))             'Konstante Krankenblattsortierung
GlDPr = CStr(GlSet(1, 87))              'Prï¿½fung bereits vorhandener Diagnosen
GlTeE = CBool(GlSet(4, 88))             'Email-Termin-Erinnerung
GlOtW = CBool(GlSet(4, 89))             'Online-Terminbuchungs System Warteliste
GlASM = CBool(GlSet(4, 90))             'Automatische SMS Terminerinnerung
GlOTK = CBool(GlSet(4, 91))             'Online-Terminbuchungs System autom. Aktualisierung
GlDeT = CBool(GlSet(4, 92))             'Stornierte Termine Termindetails
GlTSN = CStr(GlSet(1, 93))              'TSE Kennung
GLTSL = CStr(GlSet(1, 94))              'TSE Laufwerk
GlCID = CStr(GlSet(1, 95))              'Cloud-ID
If Len(GlSet(1, 96)) > 0 Then
    GlTSe = CInt(Right$(GlSet(1, 96), 1)) 'TSE Verfahren
Else
    GlTSe = 0
End If
GlTSK = CStr(GlSet(1, 97))              'TSE Organisation Key
GlTSS = CStr(GlSet(1, 98))              'TSE Organisation Secret
GlTSI = CStr(GlSet(1, 99))              'TSE Organisation ID
GlVrw = CInt(GlSet(2, 100))             'Verweildauer der Downloadlinks
If Len(GlSet(1, 101)) > 0 Then
    GlRVs = Right$(GlSet(1, 101), 1)    'Standard-Rechnungsversandweg
Else
    GlRVs = "0"                          'Default wenn leer
End If
GlMaB = CBool(GlSet(4, 102))            'E-Mail Bewertung und Markierung
GlRKr = CBool(GlSet(4, 103))            'Rechnungsvermerk im Krankenblatt
GlSPn = CStr(GlSet(1, 104))             'SMTP Praxisname
GlSIP = CStr(GlSet(1, 105))             'SMTP IP-Adresse
GlSSp = CStr(GlSet(1, 106))             'SMTP SocksProxyServer
GlSPo = CStr(GlSet(1, 107))             'SMTP SocksProxyPort
GlOtE = CBool(GlSet(4, 108))            'Online-Terminbuchungs System Storno Entfernen
GlTrD = CBool(GlSet(4, 109))            'Termindetails mit Mitarbeitername

If GlKnF = True Then 'Sachkontenformatierung sechsstellig
    TmSt1 = Val(Left$(GlSet(1, 24), 6))
    TmSt2 = Val(Left$(GlSet(1, 25), 6))
Else
    TmSt1 = Val(Left$(GlSet(1, 24), 4))
    TmSt2 = Val(Left$(GlSet(1, 25), 4))
End If

GlSE1 = SBuFo(CLng(TmSt1)) 'Standarderlï¿½skonto (Kasse)
GlSE2 = SBuFo(CLng(TmSt2)) 'Standarderlï¿½skonto (Bankkonto)

GesLe = Len(GlOSe) + Len(GlOtL)
If GesLe > 180 Then 'Online-Terminbuchungs System Beschriftung und Link Datenschutzerklï¿½rung
    NeuLe = GesLe - 180
    GlOSe = Left(GlOSe, NeuLe)
End If

If GlStK = 0 Then
    GlStK = 1
End If

If GlKe2 = 0 Then
    GlKe2 = 1
End If

If Right$(GLTSL, 1) <> "\" Then
    GLTSL = GLTSL & "\"
End If

If GlAbs = vbNullString Then 'SMS Absenderkennung
    GlAbs = "SMSAbsender"
Else
    GlAbs = SNaFi(GlAbs)
    If Len(GlAbs) > 11 Then
        GlAbs = Left$(GlAbs, 11)
    End If
End If

If GlDPr = vbNullString Then
    GlSet(1, 87) = "M6"
    GlDPr = "M6"
ElseIf Len(GlDPr) > 2 Then
    GlSet(1, 87) = "M6"
    GlDPr = "M6"
End If

GlSeE = True 'Setupdaten eingelesen

'--------------------------------------------------

Exit Sub

FiErr:
Dim ErrMsg As String
Dim ErrNum As Long
Dim ErrSrc As String
ErrNum = Err.Number
ErrSrc = Err.Source

' Build detailed error message
ErrMsg = "FEHLER in S_Ary0: " & Err.Description & " (Num=" & ErrNum & ", Source=" & ErrSrc & ")"

' Categorize by severity - critical errors exit, others continue
Select Case ErrNum
    Case -2147217865, -2147217900
        ' Invalid object name / syntax error - CRITICAL
        SErLog "CRITICAL: Datenbank-Schema Fehler in S_Ary0: " & Err.Description
        Exit Sub

    Case 3704, 3705, 91
        ' Recordset closed / Object not set - CRITICAL
        SErLog "CRITICAL: Datenbankverbindung verloren in S_Ary0: " & Err.Description
        Exit Sub

    Case 9
        ' Subscript out of range - CRITICAL
        SErLog "CRITICAL: Array-Fehler in S_Ary0: " & Err.Description
        Exit Sub

    Case -2147467259
        ' Unspecified error (often connection) - CRITICAL
        SErLog "CRITICAL: Verbindungsfehler in S_Ary0: " & Err.Description
        Exit Sub

    Case Else
        ' Unknown errors - log and try to continue
        If GlDbg = True Then
            SErLog "WARNING: in S_Ary0: " & Err.Description & " (Err.Number=" & ErrNum & ")"
        End If
        Resume Next
End Select

End Sub
Public Sub S_Ary1()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

Dim SQL1 As String
Dim MaAbs As String
Dim MaBkN As String
Dim MaBkR As String
Dim EmBet As String
Dim AktZa As Integer
Dim GesZa As Integer
Dim AktNu As Integer

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qrySimKrTyp ORDER BY IDT"
Else
    SQL1 = "SELECT * FROM qrySimKrTyp ORDER BY [IDT];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Krankenblatttypen
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    GlTyV = True 'Krankenblattypen vorhanden
    ReDim GlKrA(GesZa + 8, 6)
    Do
    GlKrA(AktZa, 0) = RS152.Fields("IDT").Value
    If RS152.Fields("IDK").Value <> vbNullString Then
        GlKrA(AktZa, 1) = RS152.Fields("IDK").Value
    Else
        GlKrA(AktZa, 1) = "XX"
    End If
    If RS152.Fields("IDKurz").Value <> vbNullString Then
        GlKrA(AktZa, 2) = RS152.Fields("IDKurz").Value
    Else
        GlKrA(AktZa, 2) = "XXX"
    End If
    GlKrA(AktZa, 3) = RS152.Fields("Farbe").Value
    If RS152.Fields("Selekt").Value <> vbNullString Then
        GlKrA(AktZa, 4) = CBool(RS152.Fields("Selekt").Value)
    Else
        GlKrA(AktZa, 4) = 0
    End If
    If RS152.Fields("Betrag").Value <> vbNullString Then
        GlKrA(AktZa, 5) = CBool(RS152.Fields("Betrag").Value)
    Else
        GlKrA(AktZa, 5) = 0
    End If
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop Until RS152.EOF
    GlKrA(AktZa, 0) = 101
    GlKrA(AktZa, 1) = "RP"
    GlKrA(AktZa, 2) = "Beleg"
    GlKrA(AktZa, 3) = 0
    GlKrA(AktZa, 4) = 0
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
    GlKrA(AktZa, 0) = 102
    GlKrA(AktZa, 1) = "DA"
    GlKrA(AktZa, 2) = "Datei"
    GlKrA(AktZa, 3) = 0
    GlKrA(AktZa, 4) = 0
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
    GlKrA(AktZa, 0) = 103
    GlKrA(AktZa, 1) = "KD"
    GlKrA(AktZa, 2) = "Diagnose"
    GlKrA(AktZa, 3) = 16512
    GlKrA(AktZa, 4) = -1
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
    GlKrA(AktZa, 0) = 104
    GlKrA(AktZa, 1) = "PR"
    GlKrA(AktZa, 2) = "Protokoll"
    GlKrA(AktZa, 3) = 8421504
    GlKrA(AktZa, 4) = 0
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
    GlKrA(AktZa, 0) = 105
    GlKrA(AktZa, 1) = "BI"
    GlKrA(AktZa, 2) = "Bilddatei"
    GlKrA(AktZa, 3) = 8388672
    GlKrA(AktZa, 4) = 0
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
    GlKrA(AktZa, 0) = 106
    GlKrA(AktZa, 1) = "RE"
    GlKrA(AktZa, 2) = "Rechnung"
    GlKrA(AktZa, 3) = 0
    GlKrA(AktZa, 4) = 0
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
    GlKrA(AktZa, 0) = 107
    GlKrA(AktZa, 1) = "KM"
    GlKrA(AktZa, 2) = "Medikament"
    GlKrA(AktZa, 3) = 16711935
    GlKrA(AktZa, 4) = 0
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
    GlKrA(AktZa, 0) = 108
    GlKrA(AktZa, 1) = "EM"
    GlKrA(AktZa, 2) = "Emails"
    GlKrA(AktZa, 3) = 8421440
    GlKrA(AktZa, 4) = 0
    GlKrA(AktZa, 5) = 0
    AktZa = AktZa + 1
Else
    S_Ary1a
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryLabGrupp ORDER BY Sorter"
Else
    SQL1 = "SELECT * FROM qryLabGrupp ORDER BY [Sorter];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Laborgruppierung
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlLGr(GesZa, 2)
    Do
    GlLGr(AktZa, 0) = RS153.Fields("IDG").Value
    If RS153.Fields("IDKurz").Value <> vbNullString Then
        GlLGr(AktZa, 1) = RS153.Fields("IDKurz").Value
    Else
        GlLGr(AktZa, 1) = "keine Gruppierung"
    End If
    GlLGr(AktZa, 2) = RS153.Fields("Sorter").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop Until RS153.EOF
Else
    ReDim GlLGr(5, 2)
    GlLGr(1, 0) = 1
    GlLGr(1, 1) = "Allgemein"
    GlLGr(1, 2) = 10
    GlLGr(2, 0) = 2
    GlLGr(2, 1) = "Hämatologischer Status"
    GlLGr(2, 2) = 20
    GlLGr(3, 0) = 3
    GlLGr(3, 1) = "Biochemischer Status"
    GlLGr(3, 2) = 30
    GlLGr(4, 0) = 4
    GlLGr(4, 1) = "Urinstatus"
    GlLGr(4, 2) = 40
    GlLGr(5, 0) = 5
    GlLGr(5, 1) = "Stuhluntersuchung"
    GlLGr(5, 2) = 50
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatBeh ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryPatBeh ORDER BY [IDKurz];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Mandanten
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    GlMaV = True
    ReDim GlThe(GesZa, 48)
    Do
    GlThe(AktZa, 0) = CLng(RS152.Fields("ID0").Value)
    If RS152.Fields("Vorname").Value <> vbNullString Then GlThe(AktZa, 1) = RS152.Fields("Vorname").Value
    If RS152.Fields("Name").Value <> vbNullString Then GlThe(AktZa, 2) = RS152.Fields("Name").Value
    If RS152.Fields("Straße").Value <> vbNullString Then GlThe(AktZa, 3) = RS152.Fields("Straße").Value
    If RS152.Fields("PLZ").Value <> vbNullString Then GlThe(AktZa, 4) = RS152.Fields("PLZ").Value
    If RS152.Fields("Ort").Value <> vbNullString Then GlThe(AktZa, 5) = RS152.Fields("Ort").Value
    If RS152.Fields("Telefon2").Value <> vbNullString Then GlThe(AktZa, 6) = RS152.Fields("Telefon2").Value
    If RS152.Fields("Telefon3").Value <> vbNullString Then GlThe(AktZa, 7) = RS152.Fields("Telefon3").Value
    If RS152.Fields("Bank").Value <> vbNullString Then GlThe(AktZa, 8) = RS152.Fields("Bank").Value
    If RS152.Fields("BLZ").Value <> vbNullString Then GlThe(AktZa, 9) = RS152.Fields("BLZ").Value
    If RS152.Fields("Konto").Value <> vbNullString Then GlThe(AktZa, 10) = RS152.Fields("Konto").Value
    If RS152.Fields("Abteilung").Value <> vbNullString Then GlThe(AktZa, 11) = RS152.Fields("Abteilung").Value
    If RS152.Fields("Beruf").Value <> vbNullString Then GlThe(AktZa, 12) = RS152.Fields("Beruf").Value
    If RS152.Fields("IDKurz").Value <> vbNullString Then GlThe(AktZa, 13) = RS152.Fields("IDKurz").Value
    If RS152.Fields("Titel").Value <> vbNullString Then GlThe(AktZa, 14) = RS152.Fields("Titel").Value
    If RS152.Fields("KVNummer").Value <> vbNullString Then GlThe(AktZa, 15) = RS152.Fields("KVNummer").Value 'LANR
    If RS152.Fields("Telefon5").Value <> vbNullString Then
        GlThe(AktZa, 16) = RS152.Fields("Telefon5").Value
    Else
        GlThe(AktZa, 16) = "keine@emailadresse.de"
    End If
    If RS152.Fields("Internet").Value <> vbNullString Then GlThe(AktZa, 17) = RS152.Fields("Internet").Value
    If RS152.Fields("IBAN").Value <> vbNullString Then GlThe(AktZa, 18) = RS152.Fields("IBAN").Value
    If RS152.Fields("R_Firma1").Value <> vbNullString Then GlThe(AktZa, 19) = RS152.Fields("R_Firma1").Value
    If RS152.Fields("Bank2").Value <> vbNullString Then GlThe(AktZa, 20) = RS152.Fields("Bank2").Value
    If RS152.Fields("BLZ2").Value <> vbNullString Then GlThe(AktZa, 21) = RS152.Fields("BLZ2").Value
    If RS152.Fields("Konto2").Value <> vbNullString Then GlThe(AktZa, 22) = RS152.Fields("Konto2").Value
    If RS152.Fields("IBAN2").Value <> vbNullString Then GlThe(AktZa, 23) = RS152.Fields("IBAN2").Value
    GlThe(AktZa, 25) = CBool(RS152.Fields("Passiv").Value)
    If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
        If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
            GlThe(AktZa, 24) = GlSZe 'Sprechzietenstring
        Else
            GlThe(AktZa, 24) = RS152.Fields("Sprechzeiten").Value
        End If
    Else
        GlThe(AktZa, 24) = GlSZe 'Sprechzietenstring
    End If
    If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
        GlThe(AktZa, 37) = RS152.Fields("Buchungszeiten").Value
    Else
        GlThe(AktZa, 37) = GlSZe 'Sprechzietenstring
    End If
    If RS152.Fields("OnlRas").Value <> vbNullString Then
        GlThe(AktZa, 26) = Format$(RS152.Fields("OnlRas").Value, "00")
    Else
        GlThe(AktZa, 26) = GlZeR 'Zeitrasterindex
    End If
    If RS152.Fields("OnlRa2").Value <> vbNullString Then
        GlThe(AktZa, 27) = Format$(RS152.Fields("OnlRa2").Value, "00")
    Else
        GlThe(AktZa, 27) = GlZeR 'Zeitrasterindex
    End If
    If RS152.Fields("Größe").Value <> vbNullString Then GlThe(AktZa, 28) = RS152.Fields("Größe").Value
    If RS152.Fields("Gewicht").Value <> vbNullString Then GlThe(AktZa, 29) = RS152.Fields("Gewicht").Value
    If RS152.Fields("Gesperrt").Value <> vbNullString Then GlThe(AktZa, 30) = RS152.Fields("Gesperrt").Value
    If RS152.Fields("BIC").Value <> vbNullString Then GlThe(AktZa, 31) = RS152.Fields("BIC").Value
    If RS152.Fields("BIC2").Value <> vbNullString Then GlThe(AktZa, 32) = RS152.Fields("BIC2").Value
    If RS152.Fields("GID").Value <> vbNullString Then GlThe(AktZa, 33) = RS152.Fields("GID").Value
    If RS152.Fields("Em_User").Value <> vbNullString Then GlThe(AktZa, 34) = RS152.Fields("Em_User").Value
    If RS152.Fields("Em_Pass").Value <> vbNullString Then GlThe(AktZa, 35) = RS152.Fields("Em_Pass").Value
    
    If GlThe(AktZa, 19) <> vbNullString Then 'R_Firma1
        If InStr(1, GlThe(AktZa, 2), GlThe(AktZa, 19), 1) > 0 Then 'Name
            If GlThe(AktZa, 14) <> vbNullString Then 'Titel
                MaAbs = GlThe(AktZa, 14) & Chr$(32) & GlThe(AktZa, 1) & Chr$(32) & GlThe(AktZa, 2)
            Else
                MaAbs = GlThe(AktZa, 1) & Chr$(32) & GlThe(AktZa, 2)
            End If
        Else
            MaAbs = GlThe(AktZa, 19) & " - " & GlThe(AktZa, 1) & " " & GlThe(AktZa, 2)
        End If
    Else
        If GlThe(AktZa, 14) <> vbNullString Then
            MaAbs = GlThe(AktZa, 14) & Chr$(32) & GlThe(AktZa, 1) & Chr$(32) & GlThe(AktZa, 2)
        Else
            MaAbs = GlThe(AktZa, 1) & Chr$(32) & GlThe(AktZa, 2)
        End If
    End If
    EmBet = MaAbs
    MaAbs = MaAbs & " - " & GlThe(AktZa, 3)
    MaAbs = MaAbs & " - " & GlThe(AktZa, 4)
    MaAbs = MaAbs & Chr$(32) & GlThe(AktZa, 5)
    GlThe(AktZa, 36) = MaAbs 'Briefabsenezeile
    GlThe(AktZa, 37) = EmBet 'Emailbetreff
    If RS152.Fields("R_Firma1").Value <> vbNullString Then
        MaBkN = RS152.Fields("R_Firma1").Value & vbCrLf
    End If
    If RS152.Fields("Titel").Value <> vbNullString Then
        MaBkN = MaBkN & RS152.Fields("Titel").Value & Chr$(32) & RS152.Fields("Vorname").Value & Chr$(32) & RS152.Fields("Name").Value
    Else
        MaBkN = MaBkN & RS152.Fields("Vorname").Value & Chr$(32) & RS152.Fields("Name").Value
    End If
    If RS152.Fields("Beruf").Value <> vbNullString Then
        MaBkN = MaBkN & vbCrLf & RS152.Fields("Beruf").Value
    End If
    MaBkR = MaBkR & vbCrLf & RS152.Fields("Straße").Value
    MaBkR = MaBkR & vbCrLf & RS152.Fields("PLZ").Value
    MaBkR = MaBkR & Chr$(32) & RS152.Fields("Ort").Value
    If RS152.Fields("Telefon2").Value <> vbNullString Then
        MaBkR = MaBkR & vbCrLf & vbCrLf & "Telefon: " & RS152.Fields("Telefon2").Value
    End If
    If RS152.Fields("Telefon3").Value <> vbNullString Then
        MaBkR = MaBkR & vbCrLf & "Telefax: " & RS152.Fields("Telefon3").Value
    End If
    GlThe(AktZa, 38) = MaBkN & vbCrLf & MaBkR
    If RS152.Fields("ID3").Value <> vbNullString Then
        GlThe(AktZa, 39) = RS152.Fields("ID3").Value
    Else
        GlThe(AktZa, 39) = GlFri 'Fachrichtung
    End If
    If RS152.Fields("GLN").Value <> vbNullString Then
        GlThe(AktZa, 40) = RS152.Fields("GLN").Value 'GLN
    Else
        GlThe(AktZa, 40) = "000000"
    End If
    If RS152.Fields("ZSR").Value <> vbNullString Then
        GlThe(AktZa, 41) = RS152.Fields("ZSR").Value 'ZSR
    Else
        GlThe(AktZa, 41) = "00000"
    End If
    If RS152.Fields("BunLan").Value <> vbNullString Then 'Bundesland
        For AktNu = 0 To UBound(GlBsl)
            If CInt(RS152.Fields("BunLan").Value) = CInt(GlBsl(AktNu, 2)) Then
                GlThe(AktZa, 42) = GlBsl(AktNu, 0)
                Exit For
            End If
        Next AktNu
    Else
        GlThe(AktZa, 42) = GlBsl(0, 0)
    End If
    If RS152.Fields("KVBez").Value <> vbNullString Then 'KVBezirk
        For AktNu = 0 To UBound(GlKVB)
            If CInt(RS152.Fields("KVBez").Value) = CInt(GlKVB(AktNu, 2)) Then
                GlThe(AktZa, 43) = GlKVB(AktNu, 0)
                Exit For
            End If
        Next AktNu
    Else
        GlThe(AktZa, 43) = GlKVB(0, 0)
    End If
    If RS152.Fields("Kanton").Value <> vbNullString Then 'Kantone
        For AktNu = 0 To UBound(GlKtn)
            If CInt(RS152.Fields("Kanton").Value) = CInt(GlKtn(AktNu, 2)) Then
                GlThe(AktZa, 44) = GlKtn(AktNu, 0)
                Exit For
            End If
        Next AktNu
    Else
        GlThe(AktZa, 44) = GlKtn(0, 0)
    End If
    If RS152.Fields("OnlRa1").Value <> vbNullString Then
        GlThe(AktZa, 46) = Format$(RS152.Fields("OnlRa1").Value, "00")
    Else
        GlThe(AktZa, 46) = "12"
    End If
    If RS152.Fields("AbrBereich").Value <> vbNullString Then
        GlThe(AktZa, 47) = RS152.Fields("AbrBereich").Value
    End If
    If RS152.Fields("Abteilung").Value <> vbNullString Then
        GlThe(AktZa, 48) = RS152.Fields("Abteilung").Value
    End If
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop Until RS152.EOF
Else
    GlMaV = False
    S_Ary2e
End If
RS152.Close
Set RS152 = Nothing
DoEvents

If GlSMa > UBound(GlThe) Then
    GlSMa = 1
End If

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatArz ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryPatArz ORDER BY [IDKurz];"
End If
AktZa = 2
Set RS153 = New ADODB.Recordset 'Verordner
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    GlArV = True 'Verordner vorhanden
    ReDim GlArz(GesZa + 1, 17)
    GlArz(1, 0) = 1
    GlArz(1, 1) = "Vorname"
    GlArz(1, 2) = "Name"
    GlArz(1, 3) = "Straße"
    GlArz(1, 4) = "PLZ"
    GlArz(1, 5) = "Ort"
    GlArz(1, 6) = "Telefon"
    GlArz(1, 7) = "Telefax"
    GlArz(1, 8) = "kein Verordner"
    GlArz(1, 9) = "Titel"
    GlArz(1, 10) = "Email"
    GlArz(1, 11) = "Beruf"
    GlArz(1, 12) = "LANR"
    GlArz(1, 13) = "GLN"
    GlArz(1, 14) = "ZSR"
    GlArz(1, 15) = GlBsl(0, 0)
    GlArz(1, 16) = GlKVB(0, 0)
    GlArz(1, 17) = GlKtn(0, 0)
    Do
    GlArz(AktZa, 0) = CLng(RS153.Fields("ID0").Value)
    If RS153.Fields("Vorname").Value <> vbNullString Then GlArz(AktZa, 1) = RS153.Fields("Vorname").Value
    If RS153.Fields("Name").Value <> vbNullString Then GlArz(AktZa, 2) = RS153.Fields("Name").Value
    If RS153.Fields("Straße").Value <> vbNullString Then GlArz(AktZa, 3) = RS153.Fields("Straße").Value
    If RS153.Fields("PLZ").Value <> vbNullString Then GlArz(AktZa, 4) = RS153.Fields("PLZ").Value
    If RS153.Fields("Ort").Value <> vbNullString Then GlArz(AktZa, 5) = RS153.Fields("Ort").Value
    If RS153.Fields("Telefon2").Value <> vbNullString Then GlArz(AktZa, 6) = RS153.Fields("Telefon2").Value
    If RS153.Fields("Telefon3").Value <> vbNullString Then GlArz(AktZa, 7) = RS153.Fields("Telefon3").Value
    If RS153.Fields("IDKurz").Value <> vbNullString Then GlArz(AktZa, 8) = RS153.Fields("IDKurz").Value
    If RS153.Fields("Titel").Value <> vbNullString Then GlArz(AktZa, 9) = RS153.Fields("Titel").Value
    If RS153.Fields("Telefon5").Value <> vbNullString Then GlArz(AktZa, 10) = RS153.Fields("Telefon5").Value
    If RS153.Fields("Beruf").Value <> vbNullString Then GlArz(AktZa, 11) = RS153.Fields("Beruf").Value
    If RS153.Fields("KVNummer").Value <> vbNullString Then GlArz(AktZa, 12) = RS153.Fields("KVNummer").Value 'LANR
    If RS153.Fields("GLN").Value <> vbNullString Then
        GlArz(AktZa, 13) = RS153.Fields("GLN").Value 'GLN
    Else
        GlArz(AktZa, 13) = "000000"
    End If
    If RS153.Fields("ZSR").Value <> vbNullString Then
        GlArz(AktZa, 14) = RS153.Fields("ZSR").Value 'ZSR
    Else
        GlArz(AktZa, 14) = "000000"
    End If
    If RS153.Fields("BunLan").Value <> vbNullString Then 'Bundesland
        For AktNu = 0 To UBound(GlBsl)
            If CInt(RS153.Fields("BunLan").Value) = CInt(GlBsl(AktNu, 2)) Then
                GlArz(AktZa, 15) = GlBsl(AktNu, 0)
                Exit For
            End If
        Next AktNu
    Else
        GlArz(AktZa, 15) = GlBsl(0, 0)
    End If
    If RS153.Fields("KVBez").Value <> vbNullString Then 'KVBezirk
        For AktNu = 0 To UBound(GlKVB)
            If CInt(RS153.Fields("KVBez").Value) = CInt(GlKVB(AktNu, 2)) Then
                GlArz(AktZa, 16) = GlKVB(AktNu, 0)
                Exit For
            End If
        Next AktNu
    Else
        GlArz(AktZa, 16) = GlKVB(0, 0)
    End If
    If RS153.Fields("Kanton").Value <> vbNullString Then 'Kantone
        For AktNu = 0 To UBound(GlKtn)
            If CInt(RS153.Fields("Kanton").Value) = CInt(GlKtn(AktNu, 2)) Then
                GlArz(AktZa, 17) = GlKtn(AktNu, 0)
                Exit For
            End If
        Next AktNu
    Else
        GlArz(AktZa, 17) = GlKtn(0, 0)
    End If
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop Until RS153.EOF
Else
    GlArV = True
    ReDim GlArz(1, 17)
    GlArz(1, 0) = 1
    GlArz(1, 1) = "Vorname"
    GlArz(1, 2) = "Name"
    GlArz(1, 3) = "Straße"
    GlArz(1, 4) = "PLZ"
    GlArz(1, 5) = "Ort"
    GlArz(1, 6) = "Telefon"
    GlArz(1, 7) = "Telefax"
    GlArz(1, 8) = "Verordnername"
    GlArz(1, 9) = "Titel"
    GlArz(1, 10) = "Email"
    GlArz(1, 11) = "Beruf"
    GlArz(1, 12) = "LANR"
    GlArz(1, 13) = "GLN"
    GlArz(1, 14) = "ZSR"
    GlArz(1, 15) = GlBsl(0, 0)
    GlArz(1, 16) = GlKVB(0, 0)
    GlArz(1, 17) = GlKtn(0, 0)
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary1 " & Err.Number
Resume Next

End Sub
Private Sub S_Ary1a()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

Dim AktZa As Integer

GlTyV = True 'Krankenblattypen vorhanden

ReDim GlKrA(27, 6)
GlKrA(1, 0) = 1
GlKrA(1, 1) = "DI"
GlKrA(1, 2) = "Diagnosen"
GlKrA(1, 3) = 8421376
GlKrA(1, 4) = 0
GlKrA(1, 5) = 0
GlKrA(2, 0) = 2
GlKrA(2, 1) = "GE"
GlKrA(2, 2) = "Gebührenziffer"
GlKrA(2, 3) = -2147483640
GlKrA(2, 4) = 0
GlKrA(2, 5) = 0
GlKrA(3, 0) = 3
GlKrA(3, 1) = "LA"
GlKrA(3, 2) = "Laborleistung"
GlKrA(3, 3) = 25800
GlKrA(3, 4) = 0
GlKrA(3, 5) = 0
GlKrA(4, 0) = 4
GlKrA(4, 1) = "ME"
GlKrA(4, 2) = "Medikament"
GlKrA(4, 3) = 16711935
GlKrA(4, 4) = 0
GlKrA(4, 5) = 0
GlKrA(5, 0) = 5
GlKrA(5, 1) = "BE"
GlKrA(5, 2) = "Begründung"
GlKrA(5, 3) = 16711680
GlKrA(5, 4) = 0
GlKrA(5, 5) = 0
GlKrA(6, 0) = 6
GlKrA(6, 1) = "ZA"
GlKrA(6, 2) = "Zahlung"
GlKrA(6, 3) = 52224
GlKrA(6, 4) = 0
GlKrA(6, 5) = 0
GlKrA(7, 0) = 7
GlKrA(7, 1) = "PR"
GlKrA(7, 2) = "Provisionseintrag"
GlKrA(7, 3) = 32768
GlKrA(7, 4) = 0
GlKrA(7, 5) = 0
GlKrA(8, 0) = 8
GlKrA(8, 1) = "IG"
GlKrA(8, 2) = "IGeL Leistungen"
GlKrA(8, 3) = 8421504
GlKrA(8, 4) = 0
GlKrA(8, 5) = 0
GlKrA(9, 0) = 9
GlKrA(9, 1) = "GL"
GlKrA(9, 2) = "Gewerbeleistung"
GlKrA(9, 3) = -2147483640
GlKrA(9, 4) = 0
GlKrA(9, 5) = 0
GlKrA(10, 0) = 21
GlKrA(10, 1) = "BuAn"
GlKrA(10, 2) = "Anamnese"
GlKrA(10, 3) = 255
GlKrA(10, 4) = 0
GlKrA(10, 5) = 0
GlKrA(11, 0) = 22
GlKrA(11, 1) = "NO"
GlKrA(11, 2) = "Notizen"
GlKrA(11, 3) = 10053171
GlKrA(11, 4) = 0
GlKrA(11, 5) = 0
GlKrA(12, 0) = 23
GlKrA(12, 1) = "BF"
GlKrA(12, 2) = "Befunde"
GlKrA(12, 3) = 10040319
GlKrA(12, 4) = 0
GlKrA(12, 5) = 0
GlKrA(13, 0) = 24
GlKrA(13, 1) = "TxPh"
GlKrA(13, 2) = "Textdokument"
GlKrA(13, 3) = 6710784
GlKrA(13, 4) = 0
GlKrA(13, 5) = 0
GlKrA(14, 0) = 25
GlKrA(14, 1) = "LB"
GlKrA(14, 2) = "Laborbefund"
GlKrA(14, 3) = 2162853
GlKrA(14, 4) = 0
GlKrA(14, 5) = 0
GlKrA(15, 0) = 26
GlKrA(15, 1) = "TH"
GlKrA(15, 2) = "Therapiekonzept"
GlKrA(15, 3) = 6736896
GlKrA(15, 4) = 0
GlKrA(15, 5) = 0
GlKrA(16, 0) = 27
GlKrA(16, 1) = "PR"
GlKrA(16, 2) = "Prozedere"
GlKrA(16, 3) = 10053375
GlKrA(16, 4) = 0
GlKrA(16, 5) = 0
GlKrA(17, 0) = 28
GlKrA(17, 1) = "BL"
GlKrA(17, 2) = "Blutdruck"
GlKrA(17, 3) = 39372
GlKrA(17, 4) = 0
GlKrA(17, 5) = 0
GlKrA(18, 0) = 29
GlKrA(18, 1) = "GW"
GlKrA(18, 2) = "Gewicht"
GlKrA(18, 3) = 16750848
GlKrA(18, 4) = 0
GlKrA(18, 5) = 0
GlKrA(19, 0) = 30
GlKrA(19, 1) = "GR"
GlKrA(19, 2) = "Größe"
GlKrA(19, 3) = 16737945
GlKrA(19, 4) = 0
GlKrA(19, 5) = 0
GlKrA(20, 0) = 101
GlKrA(20, 1) = "RP"
GlKrA(20, 2) = "Beleg"
GlKrA(20, 3) = 0
GlKrA(20, 4) = 0
GlKrA(20, 5) = 0
GlKrA(21, 0) = 102
GlKrA(21, 1) = "DA"
GlKrA(21, 2) = "Datei"
GlKrA(21, 3) = 0
GlKrA(21, 4) = 0
GlKrA(21, 5) = 0
GlKrA(22, 0) = 103
GlKrA(22, 1) = "KD"
GlKrA(22, 2) = "Diagnose"
GlKrA(22, 3) = 16512
GlKrA(22, 4) = -1
GlKrA(22, 5) = 0
GlKrA(23, 0) = 104
GlKrA(23, 1) = "PR"
GlKrA(23, 2) = "Protokoll"
GlKrA(23, 3) = 8421504
GlKrA(23, 4) = 0
GlKrA(23, 5) = 0
GlKrA(24, 0) = 105
GlKrA(24, 1) = "BI"
GlKrA(24, 2) = "Bilddatei"
GlKrA(24, 3) = 8388672
GlKrA(24, 4) = 0
GlKrA(24, 5) = 0
GlKrA(25, 0) = 106
GlKrA(25, 1) = "RE"
GlKrA(25, 2) = "Rechnung"
GlKrA(25, 3) = 0
GlKrA(25, 4) = 0
GlKrA(25, 5) = 0
GlKrA(26, 0) = 107
GlKrA(26, 1) = "KM"
GlKrA(26, 2) = "Medikament"
GlKrA(26, 3) = 16711935
GlKrA(26, 4) = 0
GlKrA(26, 5) = 0
GlKrA(27, 0) = 108
GlKrA(27, 1) = "EM"
GlKrA(27, 2) = "Emails"
GlKrA(27, 3) = 8421440
GlKrA(27, 4) = 0
GlKrA(27, 5) = 0

For AktZa = 1 To 19
    If GlKrA(AktZa, 0) = 6 Then
        If GlTyp < 2 Then
            DBCmEx3 "qrySimKrTyAb", "@IdKur", "@IdStr", "@IdFar", GlKrA(AktZa, 1), GlKrA(AktZa, 2), GlKrA(AktZa, 3)
        Else
            DBCmEx4 "qrySimKrTyAd", "@IdxNr", "@IdKur", "@IdStr", "@IdFar", GlKrA(AktZa, 0), GlKrA(AktZa, 1), GlKrA(AktZa, 2), GlKrA(AktZa, 3)
        End If
    End If
Next AktZa

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary1a " & Err.Number
Resume Next
    
End Sub
Public Sub S_Ary2()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

Dim IdxWe As Long
Dim KtoNr As Long
Dim KtoZu As Long
Dim SQL1 As String
Dim KtoSt As String
Dim KtoBe As String
Dim GuiSt As String
Dim ReStu As Single
Dim AktZa As Integer
Dim AktKo As Integer
Dim GesZa As Integer
Dim Lange As Integer
Dim Kasse As Boolean
Dim Verre As Boolean
Dim AlKon As Boolean

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTermTyp ORDER BY IDTyp"
Else
    SQL1 = "SELECT * FROM qryTermTyp ORDER BY [IDTyp];"
End If
AktZa = 1
Set RS154 = New ADODB.Recordset 'Kalendermarker
With RS154
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlTep(GesZa, 4)
    Do
    GlTep(AktZa, 0) = RS154.Fields("IDTyp").Value
    GlTep(AktZa, 1) = RS154.Fields("IDKurz").Value
    GlTep(AktZa, 2) = RS154.Fields("Farbe").Value
    If RS154.Fields("Selekt").Value <> vbNullString Then
        GlTep(AktZa, 3) = RS154.Fields("Selekt").Value
    Else
        GlTep(AktZa, 3) = 0
    End If
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop Until RS154.EOF
Else
    ReDim GlTep(1, 4)
    GlTep(1, 0) = 1
    GlTep(1, 1) = "Termintyp"
    GlTep(1, 2) = vbBlack
    GlTep(1, 3) = 0
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTermOnl ORDER BY ID5"
Else
    SQL1 = "SELECT * FROM qryTermOnl ORDER BY [ID5];"
End If
AktZa = 1
Set RS151 = New ADODB.Recordset 'OTS-Betreffs
With RS151
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS151.RecordCount
If GesZa > 0 Then
    GlOBV = True 'OTS-Betreffs vorhanden
    ReDim GlOTB(GesZa, 9)
    Do
    GlOTB(AktZa, 0) = RS151.Fields("ID5").Value
    GlOTB(AktZa, 1) = RS151.Fields("IDM").Value
    GlOTB(AktZa, 2) = RS151.Fields("IDP").Value
    If RS151.Fields("IDKurz").Value <> vbNullString Then GlOTB(AktZa, 3) = RS151.Fields("IDKurz").Value
    If RS151.Fields("GuiID").Value <> vbNullString Then GlOTB(AktZa, 4) = RS151.Fields("GuiID").Value
    If RS151.Fields("Email").Value <> vbNullString Then GlOTB(AktZa, 5) = RS151.Fields("Email").Value
    If RS151.Fields("DauMin").Value <> vbNullString Then GlOTB(AktZa, 6) = RS151.Fields("DauMin").Value
    If RS151.Fields("WoTage").Value <> vbNullString Then GlOTB(AktZa, 7) = RS151.Fields("WoTage").Value
    If RS151.Fields("Vorlauf").Value <> vbNullString Then GlOTB(AktZa, 8) = RS151.Fields("Vorlauf").Value
    If RS151.Fields("Selekt").Value <> vbNullString Then
        GlOTB(AktZa, 9) = RS151.Fields("Selekt").Value
    Else
        GlOTB(AktZa, 9) = 0
    End If
    AktZa = AktZa + 1
    RS151.MoveNext
    Loop Until RS151.EOF
End If

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatZah ORDER BY IDZ"
Else
    SQL1 = "SELECT * FROM qryPatZah ORDER BY [IDZ];"
End If
AktZa = 1
Set RS150 = New ADODB.Recordset 'Zahlungsziele
With RS150
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS150.RecordCount
If GesZa > 0 Then
    ReDim GlZah(GesZa, 6)
    Do
    GlZah(AktZa, 0) = RS150.Fields("IDZ").Value
    GlZah(AktZa, 1) = RS150.Fields("Text").Value
    GlZah(AktZa, 2) = RS150.Fields("Ziel").Value
    GlZah(AktZa, 3) = RS150.Fields("Mahnbar").Value
    GlZah(AktZa, 4) = RS150.Fields("Intervall").Value
    If IsNull(RS150.Fields("IDB").Value) = False Then
        GlZah(AktZa, 5) = RS150.Fields("IDB").Value
    Else
        GlZah(AktZa, 5) = 0
    End If
    AktZa = AktZa + 1
    RS150.MoveNext
    Loop Until RS150.EOF
Else
    ReDim GlZah(1, 6)
    GlZah(1, 0) = 1
    GlZah(1, 1) = "14 Tage"
    GlZah(1, 2) = 14
    GlZah(1, 3) = -1
    GlZah(1, 4) = 14
    GlZah(1, 5) = 1
End If
RS150.Close
Set RS150 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKatVers ORDER BY ID3"
Else
    SQL1 = "SELECT * FROM qryKatVers ORDER BY [ID3];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Gebührenkataloge
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    ReDim GlGKa(GesZa, 5)
    Do
    GlGKa(AktZa, 0) = RS152.Fields("ID3").Value
    GlGKa(AktZa, 1) = RS152.Fields("Versicherer").Value
    If RS152.Fields("IDV").Value <> vbNullString Then
        GlGKa(AktZa, 2) = RS152.Fields("IDV").Value
    Else
        GlGKa(AktZa, 2) = 0
    End If
    If RS152.Fields("Farbe").Value <> vbNullString Then
        GlGKa(AktZa, 3) = RS152.Fields("Farbe").Value
    Else
        GlGKa(AktZa, 3) = vbBlack
    End If
    If RS152.Fields("IDKey").Value <> vbNullString Then
        GlGKa(AktZa, 4) = RS152.Fields("IDKey").Value
    Else
        GlGKa(AktZa, 4) = vbNullString
    End If
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop Until RS152.EOF
Else
    ReDim GlGKa(30, 4)
    For AktZa = 1 To 30
        GlGKa(AktZa, 0) = AktZa
        GlGKa(AktZa, 1) = "Katalog " & AktZa
        GlGKa(AktZa, 2) = AktZa
        GlGKa(AktZa, 3) = vbBlack
        GlGKa(AktZa, 4) = vbNullString
    Next AktZa
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qryPatTarf WHERE Kommentar Like 'CH' ORDER BY IDKurz"
    Else
        SQL1 = "SELECT * FROM qryPatTarf WHERE [Kommentar] Like 'CH' ORDER BY [IDKurz];"
    End If
Else
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qryPatTarf ORDER BY IDKurz"
    Else
        SQL1 = "SELECT * FROM qryPatTarf ORDER BY [IDKurz];"
    End If
End If
AktZa = 1
Set RS154 = New ADODB.Recordset 'Versicherungstarife
With RS154
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlTar(GesZa, 3)
    Do
    GlTar(AktZa, 0) = RS154.Fields("IDV").Value
    GlTar(AktZa, 1) = RS154.Fields("ID3").Value
    GlTar(AktZa, 2) = RS154.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop Until RS154.EOF
Else
    ReDim GlTar(1, 3)
    GlTar(1, 0) = 1
    GlTar(1, 1) = 1
    GlTar(1, 2) = "Standardtarif"
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qrySimBuSatz ORDER BY IDS"
Else
    SQL1 = "SELECT * FROM qrySimBuSatz ORDER BY [IDS];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Steuersätze / Steuermix
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlStu(GesZa, 3)
    Do
    GlStu(AktZa, 0) = RS153.Fields("IDS").Value
    ReStu = CSng(RS153.Fields("Satz").Value)
    If ReStu > 0 And ReStu < 1 Then
        ReStu = ReStu * 100
    End If
    GlStu(AktZa, 1) = Format$(ReStu, GlWa1)
    GlStu(AktZa, 2) = RS153.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop Until RS153.EOF
Else
    ReDim GlStu(3, 3)
    GlStu(1, 0) = 1
    GlStu(1, 1) = Format$(0, GlWa1)
    GlStu(1, 2) = "ohne Steuer"
    GlStu(2, 0) = 2
    GlStu(2, 1) = Format$(19, GlWa1)
    GlStu(2, 2) = "UmSt. 19%"
    GlStu(3, 0) = 3
    GlStu(3, 1) = Format$(7, GlWa1)
    GlStu(3, 2) = "UmSt. 7%"
End If
RS153.Close
Set RS153 = Nothing
GlStV = True
DoEvents

'--------------------------------------------------

AktZa = 1
Set RS155 = New ADODB.Recordset 'Erlöskonten
RS155.CursorLocation = adUseClient
Set RS155 = DBCmRe2("qrySimBuKtt", "@IdTyp", "@IdxNr", 2, GlKtR)
GesZa = RS155.RecordCount
If GesZa > 0 Then
    ReDim GlErK(GesZa, 4)

    Set RS163 = New ADODB.Recordset 'neues Recordset erstellen
    With RS163
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockBatchOptimistic
    End With

    For Each FL101 In RS155.Fields
        RS163.Fields.Append FL101.Name, FL101.Type, FL101.DefinedSize
    Next FL101
    If RS163.State = adStateClosed Then RS163.Open
    
    Do Until RS155.EOF
        RS163.AddNew
        For Each FL101 In RS163.Fields
            If IsNull(RS155.Fields(FL101.Name).Value) = False Then
                If FL101.Name = "IDK" Then
                    If RS155.Fields("IDK").Value <> vbNullString Then
                        If RS155.Fields("IDK").Value > 0 Then
                            KtoNr = RS155.Fields("IDK").Value
                            KtoSt = SBuFo(KtoNr) 'Sachkontenformatierung
                            RS163.Fields("IDK").Value = KtoSt
                        End If
                    End If
                Else
                    RS163.Fields(FL101.Name).Value = RS155.Fields(FL101.Name).Value
                End If
            End If
        Next FL101
        RS155.MoveNext
    Loop
    RS163.UpdateBatch
    
    RS163.MoveFirst
    RS163.Sort = "IDK ASC"
    
    Do Until RS163.EOF
    KtoSt = RS163.Fields("IDK").Value
    IdxWe = RS163.Fields("IDI").Value
    If RS163.Fields("IDKurz").Value <> vbNullString Then
        KtoBe = RS163.Fields("IDKurz").Value
    Else
        KtoBe = "keine Sachkontenbezeichnung"
    End If
    GlErK(AktZa, 0) = KtoSt
    GlErK(AktZa, 1) = KtoSt & Chr$(32) & KtoBe
    GlErK(AktZa, 2) = KtoBe
    GlErK(AktZa, 3) = IdxWe '[IDI]
    AktZa = AktZa + 1
    RS163.MoveNext
    Loop

    RS163.Close
    Set RS163 = Nothing
Else
    ReDim GlErK(2, 4)
    GlErK(1, 0) = "420000"
    GlErK(1, 1) = "420000 Erlöse"
    GlErK(1, 2) = "Erlöse"
    GlErK(1, 3) = 1
    GlErK(2, 0) = "420000"
    GlErK(2, 1) = "420000 Erlöse"
    GlErK(2, 2) = "Erlöse"
    GlErK(2, 3) = 2
End If
RS155.Close
Set RS155 = Nothing
DoEvents

'--------------------------------------------------

AktZa = 1
Set RS153 = New ADODB.Recordset 'Sachkonten mit Geldkontenzuordnung
RS153.CursorLocation = adUseClient
Set RS153 = DBCmRe2("qrySimBuKtG", "@IdGel", "@IdxNr", -1, GlKtR) 'Standardkontenrahmen
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlSaK(GesZa, 7)

    Set RS163 = New ADODB.Recordset 'neues Recordset erstellen
    With RS163
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockBatchOptimistic
    End With

    For Each FL101 In RS153.Fields
        RS163.Fields.Append FL101.Name, FL101.Type, FL101.DefinedSize
    Next FL101
    If RS163.State = adStateClosed Then RS163.Open
    
    Do Until RS153.EOF
        RS163.AddNew
        For Each FL101 In RS163.Fields
            If IsNull(RS153.Fields(FL101.Name).Value) = False Then
                If FL101.Name = "IDK" Then
                    If RS153.Fields("IDK").Value <> vbNullString Then
                        If RS153.Fields("IDK").Value > 0 Then
                            KtoNr = RS153.Fields("IDK").Value
                            KtoSt = SBuFo(KtoNr) 'Sachkontenformatierung
                            RS163.Fields("IDK").Value = KtoSt
                        End If
                    End If
                Else
                    RS163.Fields(FL101.Name).Value = RS153.Fields(FL101.Name).Value
                End If
            End If
        Next FL101
        RS153.MoveNext
    Loop
    RS163.UpdateBatch
    
    RS163.MoveFirst
    RS163.Sort = "IDK ASC"

    Do Until RS163.EOF
    IdxWe = RS163.Fields("IDI").Value
    KtoSt = RS163.Fields("IDK").Value
    KtoZu = RS163.Fields("IDB").Value
    If RS163.Fields("IDKurz").Value <> vbNullString Then
        KtoBe = RS163.Fields("IDKurz").Value
    Else
        KtoBe = "keine Sachkontenbezeichnung"
    End If
    GlSaK(AktZa, 0) = IdxWe
    GlSaK(AktZa, 1) = KtoBe
    GlSaK(AktZa, 2) = KtoSt
    GlSaK(AktZa, 3) = KtoSt & Chr$(32) & KtoBe
    GlSaK(AktZa, 4) = Format$(AktZa, "00") & " " & KtoBe
    GlSaK(AktZa, 5) = AktZa
    GlSaK(AktZa, 6) = KtoZu
    AktZa = AktZa + 1
    RS163.MoveNext
    Loop

    RS163.Close
    Set RS163 = Nothing
Else
    ReDim GlSaK(4, 6)
    GlSaK(1, 0) = 1
    GlSaK(1, 1) = "Bankkonto 1"
    GlSaK(1, 2) = "180000"
    GlSaK(1, 3) = "180000 Bankkonto 1"
    GlSaK(1, 4) = "01 Bankkonto 1"
    GlSaK(1, 5) = 1
    GlSaK(1, 6) = 1
    GlSaK(2, 0) = 2
    GlSaK(2, 1) = "Kasse 1"
    GlSaK(2, 2) = "160000"
    GlSaK(2, 3) = "160000 Kasse 1"
    GlSaK(2, 4) = "02 Kasse 1"
    GlSaK(2, 5) = 2
    GlSaK(2, 6) = 2
    GlSaK(3, 0) = 3
    GlSaK(3, 1) = "Bankkonto 2"
    GlSaK(3, 2) = "180000"
    GlSaK(3, 3) = "180000 Bankkonto 2"
    GlSaK(3, 4) = "03 Bankkonto 2"
    GlSaK(3, 5) = 3
    GlSaK(3, 6) = 3
    GlSaK(4, 0) = 4
    GlSaK(4, 1) = "Bankkonto 3"
    GlSaK(4, 2) = "180000"
    GlSaK(4, 3) = "180000 Bankkonto 3"
    GlSaK(4, 4) = "04 Bankkonto 3"
    GlSaK(4, 5) = 4
    GlSaK(4, 6) = 4
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

AktZa = 1
Set RS153 = New ADODB.Recordset 'Sachkonten mit Steuerkontenzuordnung
RS153.CursorLocation = adUseClient
Set RS153 = DBCmRe2("qrySimBuKtU", "@IdStu", "@IdxNr", -1, GlKtR)
GesZa = RS153.RecordCount

If GesZa > 0 Then
    ReDim GlSaU(GesZa, 7)

    Set RS163 = New ADODB.Recordset 'neues Recordset erstellen
    With RS163
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockBatchOptimistic
    End With

    For Each FL101 In RS153.Fields
        RS163.Fields.Append FL101.Name, FL101.Type, FL101.DefinedSize
    Next FL101
    If RS163.State = adStateClosed Then RS163.Open

    Do Until RS153.EOF
        RS163.AddNew
        For Each FL101 In RS163.Fields
            If IsNull(RS153.Fields(FL101.Name).Value) = False Then
                If FL101.Name = "IDK" Then
                    If RS153.Fields("IDK").Value <> vbNullString Then
                        If RS153.Fields("IDK").Value > 0 Then
                            KtoNr = RS153.Fields("IDK").Value
                            KtoSt = SBuFo(KtoNr) 'Sachkontenformatierung
                            RS163.Fields("IDK").Value = KtoSt
                        End If
                    End If
                Else
                    RS163.Fields(FL101.Name).Value = RS153.Fields(FL101.Name).Value
                End If
            End If
        Next FL101
        RS153.MoveNext
    Loop
    RS163.UpdateBatch
    
    RS163.MoveFirst
    RS163.Sort = "IDK ASC"

    Do Until RS163.EOF
    IdxWe = RS163.Fields("IDI").Value
    KtoSt = RS163.Fields("IDK").Value
    KtoZu = RS163.Fields("IDB").Value
    If RS163.Fields("IDKurz").Value <> vbNullString Then
        KtoBe = RS163.Fields("IDKurz").Value
    Else
        KtoBe = "keine Sachkontenbezeichnung"
    End If
    GlSaU(AktZa, 0) = IdxWe
    GlSaU(AktZa, 1) = KtoBe
    GlSaU(AktZa, 2) = KtoSt
    GlSaU(AktZa, 3) = KtoSt & Chr$(32) & KtoBe
    GlSaU(AktZa, 4) = Format$(AktZa, "00") & " " & KtoBe
    GlSaU(AktZa, 5) = AktZa
    GlSaU(AktZa, 6) = KtoZu '[IDI]
    AktZa = AktZa + 1
    RS163.MoveNext
    Loop
    
    RS163.Close
    Set RS163 = Nothing
Else
    ReDim GlSaU(3, 7)
    GlSaU(1, 0) = 1
    GlSaU(1, 1) = "Umsatzsteuer 19%"
    GlSaU(1, 2) = "380600"
    GlSaU(1, 3) = "380600 Umsatzsteuer 19%"
    GlSaU(1, 4) = "01 Umsatzsteuer 19%"
    GlSaU(1, 5) = 1
    GlSaU(1, 6) = 1
    GlSaU(2, 0) = 2
    GlSaU(2, 1) = "Umsatzsteuer 07%"
    GlSaU(2, 2) = "380100"
    GlSaU(2, 3) = "380100 Umsatzsteuer 07%"
    GlSaU(2, 4) = "02 Umsatzsteuer 07%"
    GlSaU(2, 5) = 2
    GlSaU(2, 6) = 2
    GlSaU(3, 0) = 3
    GlSaU(3, 1) = "Umsatzsteuer 15%"
    GlSaU(3, 2) = "380050"
    GlSaU(3, 3) = "380050 Umsatzsteuer 15%"
    GlSaU(3, 4) = "03 Umsatzsteuer 15%"
    GlSaU(3, 5) = 3
    GlSaU(3, 6) = 3
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qrySimBuBa ORDER BY IDB"
Else
    SQL1 = "SELECT * FROM qrySimBuBa ORDER BY [IDB];"
End If
AktZa = 1
Set RS154 = New ADODB.Recordset 'Geldkonten
With RS154
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlGeK(GesZa, 9)
    Do Until RS154.EOF
    AktKo = AktKo + 1
    IdxWe = RS154.Fields("IDB").Value
    If IsNull(RS154.Fields("Kasse").Value) = False Then
        If RS154.Fields("Kasse").Value <> vbNullString Then
            If CBool(RS154.Fields("Kasse").Value) = True Then
                Kasse = True
            Else
                Kasse = False
            End If
        Else
            Kasse = False
        End If
    Else
        Kasse = False
    End If
    If IsNull(RS154.Fields("Verrechnung").Value) = False Then
        If RS154.Fields("Verrechnung").Value <> vbNullString Then
            If CBool(RS154.Fields("Verrechnung").Value) = True Then
                Verre = True
            Else
                Verre = False
            End If
        Else
            Verre = False
        End If
    Else
        Verre = False
    End If
    If RS154.Fields("Konto").Value <> vbNullString Then
        KtoNr = RS154.Fields("Konto").Value
    Else
        If Kasse = True Then
            KtoNr = 170000
        Else
            KtoNr = 180000
        End If
    End If
    If IsNull(RS154.Fields("IDI").Value) = False Then
        If RS154.Fields("IDI").Value <> vbNullString Then
            If RS154.Fields("IDI").Value > 0 Then
                KtoZu = RS154.Fields("IDI").Value
            Else
                KtoZu = AktKo
            End If
        Else
            KtoZu = AktKo
        End If
    Else
        KtoZu = AktKo
    End If
    KtoSt = SBuFo(KtoNr) 'Sachkontenformatierung
    If RS154.Fields("IDKurz").Value <> vbNullString Then
        KtoBe = RS154.Fields("IDKurz").Value
    Else
        KtoBe = "keine Sachkontenbezeichnung"
    End If
    If RS154.Fields("TSEClient").Value <> vbNullString Then
        GuiSt = RS154.Fields("TSEClient").Value
    Else
        GuiSt = vbNullString
    End If
    GlGeK(AktZa, 0) = IdxWe '[IDB]
    GlGeK(AktZa, 1) = KtoBe
    GlGeK(AktZa, 2) = KtoSt
    GlGeK(AktZa, 3) = KtoSt & Space$(1) & KtoBe
    GlGeK(AktZa, 4) = Format$(IdxWe, "00") & Space$(1) & KtoBe
    GlGeK(AktZa, 5) = Kasse
    GlGeK(AktZa, 6) = KtoZu
    GlGeK(AktZa, 7) = Verre
    GlGeK(AktZa, 8) = GuiSt
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop
Else
    ReDim GlGeK(2, 9)
    GlGeK(1, 0) = 1
    GlGeK(1, 1) = "Bankkonto"
    GlGeK(1, 2) = "120000"
    GlGeK(1, 3) = "120000 Bankkonto"
    GlGeK(1, 4) = "01 Bankkonto"
    GlGeK(1, 5) = 0
    GlGeK(1, 6) = 1
    GlGeK(1, 7) = 0
    GlGeK(1, 8) = vbNullString
    GlGeK(2, 0) = 2
    GlGeK(2, 1) = "Kasse"
    GlGeK(2, 2) = "100000"
    GlGeK(2, 2) = "100000 Kasse"
    GlGeK(2, 4) = "02 Kasse"
    GlGeK(2, 5) = -1
    GlGeK(2, 6) = 2
    GlGeK(2, 7) = 0
    GlGeK(2, 8) = vbNullString
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qrySimBuTex ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qrySimBuTex ORDER BY [IDKurz];"
End If
AktZa = 1
Set RS155 = New ADODB.Recordset 'Buchungstexte
With RS155
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS155.RecordCount
If GesZa > 0 Then
    ReDim GlBTe(GesZa, 2)
    Do Until RS155.EOF
    GlBTe(AktZa, 0) = RS155.Fields("ID1").Value
    GlBTe(AktZa, 1) = RS155.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS155.MoveNext
    Loop
Else
    ReDim GlBTe(1, 2)
    GlBTe(1, 0) = 1
    GlBTe(1, 1) = "Bankkonto"
End If
RS155.Close
Set RS155 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatWar ORDER BY IDW"
Else
    SQL1 = "SELECT * FROM qryPatWar ORDER BY [IDW];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Wärungen
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    ReDim GlWar(GesZa, 3)
    Do Until RS152.EOF
    GlWar(AktZa, 0) = RS152.Fields("IDW").Value
    GlWar(AktZa, 1) = RS152.Fields("Währung").Value
    GlWar(AktZa, 2) = RS152.Fields("Symbol").Value
    GlWar(AktZa, 3) = RS152.Fields("Faktor").Value
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop
Else
    ReDim GlWar(2, 3)
    GlWar(1, 0) = 1
    GlWar(1, 1) = "Euro (€)"
    GlWar(1, 2) = "€"
    GlWar(1, 3) = 1
    GlWar(2, 0) = 2
    GlWar(2, 1) = "Euro (€)"
    GlWar(2, 2) = "€"
    GlWar(2, 3) = 1
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

If GlTBe = True Then
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qryKontBetr ORDER BY Betreff"
    Else
        SQL1 = "SELECT * FROM qryKontBetr ORDER BY [Betreff];"
    End If
Else
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qryKontBetr ORDER BY ID4"
    Else
        SQL1 = "SELECT * FROM qryKontBetr ORDER BY [ID4];"
    End If
End If

AktZa = 1
Set RS154 = New ADODB.Recordset 'Terminbetreffs
With RS154
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlBtr(GesZa, 6)
    Do Until RS154.EOF
    GlBtr(AktZa, 0) = RS154.Fields("ID4").Value
    GlBtr(AktZa, 1) = RS154.Fields("Betreff").Value
    GlBtr(AktZa, 2) = RS154.Fields("IDR").Value
    GlBtr(AktZa, 3) = RS154.Fields("Zeit").Value
    If RS154.Fields("Farbe").Value <> vbNullString Then
        If RS154.Fields("Farbe").Value > 0 Then
            GlBtr(AktZa, 4) = RS154.Fields("Farbe").Value
        Else
            GlBtr(AktZa, 4) = vbWhite
        End If
    Else
        GlBtr(AktZa, 4) = vbWhite
    End If
    If RS154.Fields("IDM").Value <> vbNullString Then
        If RS154.Fields("IDM").Value > 0 Then
            GlBtr(AktZa, 5) = RS154.Fields("IDM").Value
        Else
            GlBtr(AktZa, 5) = 0
        End If
    Else
        GlBtr(AktZa, 5) = 0
    End If
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop
Else
    ReDim GlBtr(1, 6)
    GlBtr(1, 0) = 1
    GlBtr(1, 1) = "Terminbetreff 01"
    GlBtr(1, 2) = 0
    GlBtr(1, 3) = 0
    GlBtr(1, 4) = vbWhite
    GlBtr(1, 5) = 0
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlRSo = True Then 'Die Raumzuordnung numerisch sortiert anzeigen
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qryKontOrt ORDER BY ID4"
    Else
        SQL1 = "SELECT * FROM qryKontOrt ORDER BY [ID4];"
    End If
Else
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qryKontOrt ORDER BY Ort"
    Else
        SQL1 = "SELECT * FROM qryKontOrt ORDER BY [Ort];"
    End If
End If

AktZa = 1
Set RS155 = New ADODB.Recordset 'Raumzuordnung
With RS155
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS155.RecordCount
If GesZa > 0 Then
    ReDim GlRmu(GesZa, 5)
    Do Until RS155.EOF
    GlRmu(AktZa, 0) = AktZa
    GlRmu(AktZa, 1) = RS155.Fields("Ort").Value
    GlRmu(AktZa, 2) = RS155.Fields("ID4").Value
    If RS155.Fields("Typ").Value <> vbNullString Then
        GlRmu(AktZa, 3) = RS155.Fields("Typ").Value
    Else
        GlRmu(AktZa, 3) = 0
    End If
    If RS155.Fields("IDM").Value <> vbNullString Then
        GlRmu(AktZa, 4) = RS155.Fields("IDM").Value
    Else
        GlRmu(AktZa, 4) = 0
    End If
    AktZa = AktZa + 1
    RS155.MoveNext
    Loop
Else
    ReDim GlRmu(1, 5)
    GlRmu(1, 0) = 1
    GlRmu(1, 1) = "Raum#1"
    GlRmu(1, 2) = 1
    GlRmu(1, 3) = 0
    GlRmu(1, 4) = 0
End If
RS155.Close
Set RS155 = Nothing
GlRaV = True
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKat03 ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryKat03 ORDER BY [IDKurz];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Diagnosegruppen
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    ReDim GlDia(GesZa, 2)
    Do Until RS152.EOF
    GlDia(AktZa, 0) = RS152.Fields("ID3").Value
    GlDia(AktZa, 1) = RS152.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop
Else
    ReDim GlDia(1, 2)
    GlDia(1, 0) = 1
    GlDia(1, 1) = " Meine Diagnoseauswahl"
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKat04 ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryKat04 ORDER BY [IDKurz];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Arzneigruppen
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlMed(GesZa, 2)
    Do Until RS153.EOF
    GlMed(AktZa, 0) = RS153.Fields("ID3").Value
    GlMed(AktZa, 1) = RS153.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop
Else
    ReDim GlMed(1, 2)
    GlMed(1, 0) = 1
    GlMed(1, 1) = "_Meine Arzneiauswahl"
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKat12 ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryKat12 ORDER BY [IDKurz];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Artikelgruppen
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlArt(GesZa, 2)
    Do Until RS153.EOF
    GlArt(AktZa, 0) = RS153.Fields("ID3").Value
    GlArt(AktZa, 1) = RS153.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop
Else
    ReDim GlArt(1, 2)
    GlArt(1, 0) = 1
    GlArt(1, 1) = "_Meine Artikelauswahl"
    DBCmEx2 "qryKat12Ad", "@IdKurz", "@IdSel", "_Meine Artikelauswahl", -1
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKat05 ORDER BY ID3"
Else
    SQL1 = "SELECT * FROM qryKat05 ORDER BY [ID3];"
End If
AktZa = 1
Set RS154 = New ADODB.Recordset 'Fragebogengruppe
With RS154
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlAnG(GesZa, 2)
    Do Until RS154.EOF
    GlAnG(AktZa, 0) = RS154.Fields("ID3").Value
    GlAnG(AktZa, 1) = RS154.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop
Else
    ReDim GlAnG(1, 2)
    GlAnG(1, 0) = 6
    GlAnG(1, 1) = "Fragebogengruppe"
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKat10 ORDER BY ID3"
Else
    SQL1 = "SELECT * FROM qryKat10 ORDER BY [ID3];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Fragebogen
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlFrB(GesZa, 4)
    Do Until RS153.EOF
    GlFrB(AktZa, 0) = RS153.Fields("ID3").Value
    GlFrB(AktZa, 1) = RS153.Fields("IDKurz").Value
    GlFrB(AktZa, 2) = RS153.Fields("GuiID").Value
    GlFrB(AktZa, 3) = RS153.Fields("Weblink").Value
    GlFrB(AktZa, 4) = RS153.Fields("WebID").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop
Else
    ReDim GlFrB(1, 4)
    GlFrB(1, 0) = 1
    GlFrB(1, 1) = "Fragebogengruppe"
    GlFrB(1, 2) = vbNullString
    GlFrB(1, 3) = vbNullString
    GlFrB(1, 4) = vbNullString
End If
RS153.Close
Set RS153 = Nothing
GlBoV = GesZa 'Fragebogen vorhanden
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKat09 ORDER BY ID3"
Else
    SQL1 = "SELECT * FROM qryKat09 ORDER BY [ID3];"
End If
AktZa = 1
Set RS154 = New ADODB.Recordset 'Laborkataloge
With RS154
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlLab(GesZa, 2)
    Do Until RS154.EOF
    GlLab(AktZa, 0) = RS154.Fields("ID3").Value
    GlLab(AktZa, 1) = RS154.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop
Else
    ReDim GlLab(1, 2)
    GlLab(1, 0) = 6
    GlLab(1, 1) = "Laborkatalog"
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTerFar ORDER BY IDTyp"
Else
    SQL1 = "SELECT * FROM qryTerFar ORDER BY [IDTyp];"
End If
AktZa = 1
Set RS150 = New ADODB.Recordset 'Terminfarben
With RS150
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS150.RecordCount
If GesZa > 0 Then
    ReDim GlTmF(GesZa, 3)
    Do Until RS150.EOF
    GlTmF(AktZa, 0) = RS150.Fields("IDKurz").Value
    GlTmF(AktZa, 1) = RS150.Fields("Farbe").Value
    GlTmF(AktZa, 2) = RS150.Fields("IDTyp").Value
    AktZa = AktZa + 1
    RS150.MoveNext
    Loop
Else
    ReDim GlTmF(20, 3)
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
        GlTmF(AktZa, 2) = AktZa
        GlTmF(AktZa, 0) = "Terminfarbe " & Format$(AktZa, "00")
        DBCmEx2 "qryTerFaAd", "@IdStr", "@IdFar", GlTmF(AktZa, 0), GlTmF(AktZa, 1)
    Next AktZa
End If
RS150.Close
Set RS150 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTerHin ORDER BY IDTyp"
Else
    SQL1 = "SELECT * FROM qryTerHin ORDER BY [IDTyp];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Terminhintergrund
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    If GesZa < 19 Then
        ReDim GlTmH(19, 6) 'Terminhintergrund
        Do Until RS152.EOF
        GlTmH(AktZa, 0) = RS152.Fields("IDTyp").Value
        GlTmH(AktZa, 1) = RS152.Fields("Farbe1").Value
        GlTmH(AktZa, 2) = RS152.Fields("Farbe2").Value
        GlTmH(AktZa, 3) = RS152.Fields("Farbe3").Value
        GlTmH(AktZa, 4) = RS152.Fields("Farbe4").Value
        GlTmH(AktZa, 5) = RS152.Fields("Farbe5").Value
        AktZa = AktZa + 1
        RS152.MoveNext
        Loop
        For AktZa = 12 To 19
            GlTmH(AktZa, 0) = AktZa
            GlTmH(AktZa, 1) = GlFaS(AktZa, 1)
            GlTmH(AktZa, 2) = GlFaS(AktZa, 2)
            GlTmH(AktZa, 3) = GlFaS(AktZa, 3)
            GlTmH(AktZa, 4) = GlFaS(AktZa, 4)
            GlTmH(AktZa, 5) = GlFaS(AktZa, 5)
            DBCmEx5 "qryTerHiAd", "@IdFa1", "@IdFa2", "@IdFa3", "@IdFa4", "@IdFa5", GlFaS(AktZa, 1), GlFaS(AktZa, 2), GlFaS(AktZa, 3), GlFaS(AktZa, 4), GlFaS(AktZa, 5)
        Next AktZa
    Else
        ReDim GlTmH(GesZa, 6) 'Terminhintergrund
        Do Until RS152.EOF
        GlTmH(AktZa, 0) = RS152.Fields("IDTyp").Value
        GlTmH(AktZa, 1) = RS152.Fields("Farbe1").Value
        GlTmH(AktZa, 2) = RS152.Fields("Farbe2").Value
        GlTmH(AktZa, 3) = RS152.Fields("Farbe3").Value
        GlTmH(AktZa, 4) = RS152.Fields("Farbe4").Value
        GlTmH(AktZa, 5) = RS152.Fields("Farbe5").Value
        AktZa = AktZa + 1
        RS152.MoveNext
        Loop
    End If
Else
    ReDim GlTmH(19, 6) 'Terminhintergrund
    For AktZa = 1 To 19
        GlTmH(AktZa, 0) = AktZa
        GlTmH(AktZa, 1) = GlFaS(AktZa, 1)
        GlTmH(AktZa, 2) = GlFaS(AktZa, 2)
        GlTmH(AktZa, 3) = GlFaS(AktZa, 3)
        GlTmH(AktZa, 4) = GlFaS(AktZa, 4)
        GlTmH(AktZa, 5) = GlFaS(AktZa, 5)
        DBCmEx5 "qryTerHiAd", "@IdFa1", "@IdFa2", "@IdFa3", "@IdFa4", "@IdFa5", GlFaS(AktZa, 1), GlFaS(AktZa, 2), GlFaS(AktZa, 3), GlFaS(AktZa, 4), GlFaS(AktZa, 5)
    Next AktZa
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qrySimReKm ORDER BY IDK"
Else
    SQL1 = "SELECT * FROM qrySimReKm ORDER BY [IDK];"
End If
AktZa = 1
Set RS155 = New ADODB.Recordset 'Rechnungskommentare
With RS155
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS155.RecordCount
If GesZa > 0 Then
    ReDim GlReK(GesZa, 2)
    Do Until RS155.EOF
    GlReK(AktZa, 0) = RS155.Fields("IDK").Value
    If IsNull(RS155.Fields("IDKurz").Value) Then
        GlReK(AktZa, 1) = Chr$(32)
    Else
        GlReK(AktZa, 1) = RS155.Fields("IDKurz").Value
    End If
    AktZa = AktZa + 1
    RS155.MoveNext
    Loop
Else
    ReDim GlReK(1, 2)
    GlReK(1, 0) = 1
    GlReK(1, 1) = vbNullString
End If
RS155.Close
Set RS155 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then 'Adressengruppen (Patientengruppen)
    SQL1 = "SELECT * FROM dbo.qryPatGrup ORDER BY TreKey"
Else
    SQL1 = "SELECT * FROM qryPatGrup ORDER BY [TreKey];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlPaG(GesZa, 2)
    Do
    GlPaG(AktZa, 0) = RS153.Fields("ID1").Value
    If IsNull(RS153.Fields("IDKurz").Value) Then
        GlPaG(AktZa, 1) = Space$(1)
    Else
        GlPaG(AktZa, 1) = RS153.Fields("IDKurz").Value
    End If
    If IsNull(RS153.Fields("TreKey").Value) Then
        GlPaG(AktZa, 2) = 1
    Else
        GlPaG(AktZa, 2) = RS153.Fields("TreKey").Value
    End If
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop Until RS153.EOF
Else
    ReDim GlPaG(1, 2)
    GlPaG(1, 0) = 0
    GlPaG(1, 1) = "Adressgruppen"
    GlPaG(1, 2) = "001"
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

AktZa = 1
Set RS154 = New ADODB.Recordset 'Emailgruppen
RS154.CursorLocation = adUseClient
Set RS154 = DBCmRe1("qryMailGrMi", "@IdMit", GlMiA(GlSmI, 2))
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlEmG(GesZa, 3)
    Do
    GlEmG(AktZa, 0) = RS154.Fields("ID1").Value
    If IsNull(RS154.Fields("IDKurz").Value) Then
        GlEmG(AktZa, 1) = Space$(1)
    Else
        GlEmG(AktZa, 1) = RS154.Fields("IDKurz").Value
    End If
    If IsNull(RS154.Fields("TreKey").Value) Then
        GlEmG(AktZa, 2) = 1
    Else
        GlEmG(AktZa, 2) = RS154.Fields("TreKey").Value
    End If
    If IsNull(RS154.Fields("IDM").Value) Then
        GlEmG(AktZa, 3) = 0
    Else
        GlEmG(AktZa, 3) = RS154.Fields("IDM").Value
    End If
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop Until RS154.EOF
Else
    ReDim GlEmG(1, 3)
    GlEmG(1, 0) = 0
    GlEmG(1, 1) = "Emailgruppen"
    GlEmG(1, 2) = "001"
    GlEmG(1, 3) = 0
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryLand ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryLand ORDER BY [IDKurz];"
End If
AktZa = 2
Set RS153 = New ADODB.Recordset 'Länder
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlLan(GesZa + 1, 3)
    GlLan(1, 0) = 1
    GlLan(1, 1) = vbNullString
    GlLan(1, 2) = vbNullString
    Do
    GlLan(AktZa, 0) = RS153.Fields("IDL").Value
    GlLan(AktZa, 1) = RS153.Fields("IDKurz").Value
    GlLan(AktZa, 2) = RS153.Fields("Sprache").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop Until RS153.EOF
Else
    ReDim GlLan(4, 3)
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
End If
RS153.Close
Set RS153 = Nothing
DoEvents

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryKat11 ORDER BY IDG"
Else
    SQL1 = "SELECT * FROM qryKat11 ORDER BY [IDG];"
End If
AktZa = 1
Set RS154 = New ADODB.Recordset 'Behinderungsgrade
With RS154
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlBeG(GesZa, 2)
    Do Until RS154.EOF
    GlBeG(AktZa, 0) = RS154.Fields("IDG").Value
    GlBeG(AktZa, 1) = RS154.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop
Else
    ReDim GlBeG(1, 2)
    GlBeG(1, 0) = 1
    GlBeG(1, 1) = "Keine Beeinträchtigungen"
End If
RS154.Close
Set RS154 = Nothing
GlBgV = True
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTermSta ORDER BY IDS"
Else
    SQL1 = "SELECT * FROM qryTermSta ORDER BY [IDS];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Terminstatus
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlTeS(GesZa, 2)
    Do Until RS153.EOF
    GlTeS(AktZa, 0) = RS153.Fields("IDS").Value
    GlTeS(AktZa, 1) = RS153.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop
Else
    ReDim GlTeS(4, 2)
    GlTeS(1, 0) = 1
    GlTeS(1, 1) = "Abgesagt"
    GlTeS(2, 0) = 2
    GlTeS(2, 1) = "Vorläufig"
    GlTeS(3, 0) = 3
    GlTeS(3, 1) = "Ordentlich"
    GlTeS(4, 0) = 4
    GlTeS(4, 1) = "Verpasst"
    For AktZa = 1 To UBound(GlTeS)
        DBCmEx1 "qryTermStAd", "@IdStr", GlTeS(AktZa, 1)
    Next AktZa
End If
RS153.Close
Set RS153 = Nothing
GlTrV = True
DoEvents

'--------------------------------------------------

AktZa = 1
If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryMailKonta ORDER BY IDK"
Else
    SQL1 = "SELECT * FROM qryMailKonta ORDER BY [IDK];"
End If
Set RS152 = New ADODB.Recordset 'Emailkonten
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    GlEKV = True 'Emailkonten vorhanden
    ReDim GlMkt(GesZa, 21)
    Do Until RS152.EOF
    AlKon = False
    GlMkt(AktZa, 0) = RS152.Fields("IDK").Value
    If RS152.Fields("IDM").Value <> vbNullString Then
        GlMkt(AktZa, 1) = RS152.Fields("IDM").Value 'Mitarbeiter
    Else
        GlMkt(AktZa, 1) = 0
    End If
    If RS152.Fields("Em_Verf").Value <> vbNullString Then
        If RS152.Fields("Em_Verf").Value > 0 Then
            GlMkt(AktZa, 2) = RS152.Fields("Em_Verf").Value 'IMAP/'POP3
        Else
            GlMkt(AktZa, 2) = 2 'POP3
            AlKon = True
        End If
    Else
        GlMkt(AktZa, 2) = 2 'POP3
        AlKon = True
    End If
    If RS152.Fields("Em_TLS1").Value <> vbNullString Then
        If AlKon = False Then
            GlMkt(AktZa, 3) = RS152.Fields("Em_TLS1").Value
        Else
            GlMkt(AktZa, 3) = GlPro(14, 4)
        End If
    Else
        GlMkt(AktZa, 3) = GlPro(14, 4)
    End If
    If RS152.Fields("Em_TLS2").Value <> vbNullString Then
        If AlKon = False Then
            GlMkt(AktZa, 4) = RS152.Fields("Em_TLS2").Value
        Else
            GlMkt(AktZa, 4) = GlPro(14, 5)
        End If
    Else
        GlMkt(AktZa, 4) = GlPro(14, 5)
    End If
    If RS152.Fields("Em_IMAP").Value <> vbNullString Then GlMkt(AktZa, 5) = RS152.Fields("Em_IMAP").Value
    If RS152.Fields("Em_POP").Value <> vbNullString Then GlMkt(AktZa, 6) = RS152.Fields("Em_POP").Value
    If RS152.Fields("Em_SMTP").Value <> vbNullString Then GlMkt(AktZa, 7) = RS152.Fields("Em_SMTP").Value
    If RS152.Fields("Em_Port1").Value <> vbNullString Then
        If RS152.Fields("Em_Port1").Value > 0 Then
            GlMkt(AktZa, 8) = CInt(RS152.Fields("Em_Port1").Value)
        Else
            GlMkt(AktZa, 8) = 110
        End If
    Else
        GlMkt(AktZa, 8) = 110
    End If
    If RS152.Fields("Em_Port2").Value <> vbNullString Then
        If RS152.Fields("Em_Port2").Value > 0 Then
            GlMkt(AktZa, 9) = CInt(RS152.Fields("Em_Port2").Value)
        Else
            GlMkt(AktZa, 9) = 587
        End If
    Else
        GlMkt(AktZa, 9) = 587
    End If
    If RS152.Fields("Em_User").Value <> vbNullString Then GlMkt(AktZa, 10) = RS152.Fields("Em_User").Value
    If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMkt(AktZa, 11) = RS152.Fields("Em_Pass").Value
    If RS152.Fields("Em_SeNam").Value <> vbNullString Then GlMkt(AktZa, 12) = RS152.Fields("Em_SeNam").Value
    If RS152.Fields("Em_Adresse").Value <> vbNullString Then GlMkt(AktZa, 13) = RS152.Fields("Em_Adresse").Value
    If RS152.Fields("Em_Repl").Value <> vbNullString Then GlMkt(AktZa, 14) = RS152.Fields("Em_Repl").Value
    If RS152.Fields("Em_Aut").Value <> vbNullString Then
        GlMkt(AktZa, 15) = CBool(RS152.Fields("Em_Aut").Value)
    Else
        GlMkt(AktZa, 15) = -1
    End If
    If RS152.Fields("Selekt").Value <> vbNullString Then
        GlMkt(AktZa, 16) = CBool(RS152.Fields("Selekt").Value)
    Else
        GlMkt(AktZa, 16) = 0 'Emails auf Server belassen
    End If
    If RS152.Fields("Em_Chk").Value <> vbNullString Then
        GlMkt(AktZa, 17) = CBool(RS152.Fields("Em_Chk").Value)
    Else
        GlMkt(AktZa, 17) = 0 'Als gelesen kennzeichnen
    End If
    If RS152.Fields("Em_Back").Value <> vbNullString Then
        GlMkt(AktZa, 18) = CInt(RS152.Fields("Em_Back").Value)
    Else
        GlMkt(AktZa, 18) = 0 'Emailabrufalter
    End If
    If RS152.Fields("Em_Neu").Value <> vbNullString Then
        GlMkt(AktZa, 19) = CBool(RS152.Fields("Em_Neu").Value)
    Else
        GlMkt(AktZa, 19) = 0 'Nur neue Emails abrufen
    End If
    If RS152.Fields("Em_Mas").Value <> vbNullString Then 'Standardmeilkonto
        GlMkt(AktZa, 20) = CBool(RS152.Fields("Em_Mas").Value)
    Else
        GlMkt(AktZa, 20) = 0
    End If
    If RS152.Fields("Anonym").Value <> vbNullString Then
        GlMkt(AktZa, 21) = CBool(RS152.Fields("Anonym").Value)
    Else
        GlMkt(AktZa, 21) = 0
    End If
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

AktZa = 1
Set RS153 = New ADODB.Recordset 'Emailfilter
RS153.CursorLocation = adUseClient
Set RS153 = DBCmRe1("qryMailSpMi", "@IdMit", GlMiA(GlSmI, 2))
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlMft(GesZa, 6)
    Do Until RS153.EOF
    GlMft(AktZa, 0) = RS153.Fields("IDA").Value
    If RS153.Fields("ID1").Value <> vbNullString Then GlMft(AktZa, 1) = RS153.Fields("ID1").Value
    If RS153.Fields("Subject").Value <> vbNullString Then GlMft(AktZa, 2) = RS153.Fields("Subject").Value
    If RS153.Fields("SenderMail").Value <> vbNullString Then GlMft(AktZa, 3) = RS153.Fields("SenderMail").Value
    If RS153.Fields("Selekt").Value <> vbNullString Then
        GlMft(AktZa, 4) = RS153.Fields("Selekt").Value
    Else
        GlMft(AktZa, 4) = 0
    End If
    If RS153.Fields("TreKey").Value <> vbNullString Then GlMft(AktZa, 5) = RS153.Fields("TreKey").Value
    If RS153.Fields("IDKurz").Value <> vbNullString Then GlMft(AktZa, 6) = RS153.Fields("IDKurz").Value
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop
Else
    ReDim GlMft(1, 6)
    GlMft(1, 0) = 1
    GlMft(1, 1) = 801
    GlMft(1, 2) = vbNullString
    GlMft(1, 3) = vbNullString
    GlMft(1, 4) = -1
    GlMft(1, 5) = vbNullString
    GlMft(1, 6) = "Testfilter"
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

If GlStK > UBound(GlGKa) Then
    GlStK = 1
End If

AktZa = 1
Set RS156 = New ADODB.Recordset
RS156.CursorLocation = adUseClient
Set RS156 = DBCmRe1("qryKat01H", "@IdxNr", GlGKa(GlStK, 0)) 'Gebührenketten
GesZa = RS156.RecordCount
If GesZa > 0 Then
    ReDim GlKet(GesZa, 3)
    Do Until RS156.EOF
    GlKet(AktZa, 0) = RS156.Fields("ID1").Value
    If RS156.Fields("GOID").Value <> vbNullString Then GlKet(AktZa, 1) = RS156.Fields("GOID").Value
    If RS156.Fields("IDKurz").Value <> vbNullString Then GlKet(AktZa, 2) = RS156.Fields("IDKurz").Value
    If RS156.Fields("Preis").Value <> vbNullString Then GlKet(AktZa, 3) = RS156.Fields("Preis").Value
    AktZa = AktZa + 1
    RS156.MoveNext
    Loop
Else
    ReDim GlKet(1, 3)
    GlKet(1, 0) = 1
    GlKet(1, 1) = "KET1"
    GlKet(1, 2) = "Gebührenkette"
    GlKet(1, 3) = "0,00"
End If
RS156.Close
Set RS156 = Nothing
DoEvents

'--------------------------------------------------

AktZa = 1
Set RS154 = New ADODB.Recordset 'Zugehörige Adressen Emailadressen
RS154.CursorLocation = adUseClient
Set RS154 = DBCmRe1("qryMailAdFil", "@IdStr", "%@%")
GesZa = RS154.RecordCount
If GesZa > 0 Then
    ReDim GlZAd(GesZa, 2)
    Do Until RS154.EOF
    GlZAd(AktZa, 0) = RS154.Fields("IDA").Value
    If RS154.Fields("IDKurz").Value <> vbNullString Then GlZAd(AktZa, 1) = RS154.Fields("IDKurz").Value
    If RS154.Fields("Telefon5").Value <> vbNullString Then GlZAd(AktZa, 2) = RS154.Fields("Telefon5").Value
    AktZa = AktZa + 1
    RS154.MoveNext
    Loop
Else
    ReDim GlZAd(1, 2)
    GlZAd(1, 0) = 1
    GlZAd(1, 1) = "Mustermann"
    GlZAd(1, 2) = "Mustermann@Mustermann.de"
End If
RS154.Close
Set RS154 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryMailBeEdi ORDER BY ID1"
Else
    SQL1 = "SELECT * FROM qryMailBeEdi ORDER BY [ID1];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Emailtextvorlagen
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    If GesZa > 16 Then 'WICHTIG!
        If GlRDP = False Then
            S_Ary2j
        Else
            ReDim GlEmT(16, 3)
            For AktZa = 1 To 16
            GlEmT(AktZa, 0) = RS153.Fields("ID1").Value
            If RS153.Fields("IDKurz").Value <> vbNullString Then GlEmT(AktZa, 1) = RS153.Fields("IDKurz").Value
            If RS153.Fields("IDKey").Value <> vbNullString Then GlEmT(AktZa, 2) = RS153.Fields("IDKey").Value
            If RS153.Fields("Betreff").Value <> vbNullString Then GlEmT(AktZa, 3) = RS153.Fields("Betreff").Value
            RS153.MoveNext
            Next AktZa
        End If
    Else
        ReDim GlEmT(GesZa, 3)
        Do Until RS153.EOF
        GlEmT(AktZa, 0) = RS153.Fields("ID1").Value
        If RS153.Fields("IDKurz").Value <> vbNullString Then GlEmT(AktZa, 1) = RS153.Fields("IDKurz").Value
        If RS153.Fields("IDKey").Value <> vbNullString Then GlEmT(AktZa, 2) = RS153.Fields("IDKey").Value
        If RS153.Fields("Betreff").Value <> vbNullString Then GlEmT(AktZa, 3) = RS153.Fields("Betreff").Value
        AktZa = AktZa + 1
        RS153.MoveNext
        Loop
    End If
Else
    S_Ary2j
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryMailTxEdi ORDER BY ID1"
Else
    SQL1 = "SELECT * FROM qryMailTxEdi ORDER BY [ID1];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Terminnachrichtentexte
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    ReDim GlEmN(GesZa, 3)
    Do Until RS152.EOF
    GlEmN(AktZa, 0) = AktZa
    If RS152.Fields("IDKurz").Value <> vbNullString Then GlEmN(AktZa, 1) = RS152.Fields("IDKurz").Value
    If RS152.Fields("IDKey").Value <> vbNullString Then GlEmN(AktZa, 2) = RS152.Fields("IDKey").Value
    If RS152.Fields("Betreff").Value <> vbNullString Then GlEmN(AktZa, 3) = RS152.Fields("Betreff").Value
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop
Else
    S_Ary2i
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

AktZa = 1
Set RS156 = New ADODB.Recordset
RS156.CursorLocation = adUseClient
Set RS156 = DBCmRe1("qryKat04E", "@IdxNr", 9) 'Textphrasen
GesZa = RS156.RecordCount
If GesZa > 0 Then
    ReDim GlTxP(GesZa, 3)
    Do Until RS156.EOF
    GlTxP(AktZa, 0) = RS156.Fields("ID0").Value
    If RS156.Fields("GOID").Value <> vbNullString Then GlTxP(AktZa, 1) = RS156.Fields("GOID").Value
    If RS156.Fields("IDKurz").Value <> vbNullString Then GlTxP(AktZa, 2) = RS156.Fields("IDKurz").Value
    If RS156.Fields("Preis1").Value <> vbNullString Then GlTxP(AktZa, 3) = RS156.Fields("Preis1").Value
    AktZa = AktZa + 1
    RS156.MoveNext
    Loop
Else
    ReDim GlTxP(1, 3)
    GlTxP(1, 0) = 1
    GlTxP(1, 1) = "mfg"
    GlTxP(1, 2) = "Mit freundlichen Grüßen"
    GlTxP(1, 3) = "0,00"
End If
RS156.Close
Set RS156 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryTerZeNe ORDER BY Datum"
Else
    SQL1 = "SELECT * FROM qryTerZeNe ORDER BY [Datum];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Terminnachrichtenvorlagen
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    ReDim GlSpr(GesZa, 4)
    Do Until RS152.EOF
    GlSpr(AktZa, 0) = RS152.Fields("IDZ").Value
    If RS152.Fields("ID0").Value <> vbNullString Then GlSpr(AktZa, 1) = RS152.Fields("ID0").Value 'Mandanten / Mitarbeiternummer
    If RS152.Fields("Sprechzeiten").Value <> vbNullString Then GlSpr(AktZa, 2) = RS152.Fields("Sprechzeiten").Value
    If RS152.Fields("Buchungszeiten").Value <> vbNullString Then GlSpr(AktZa, 3) = RS152.Fields("Buchungszeiten").Value
    If RS152.Fields("Datum").Value <> vbNullString Then GlSpr(AktZa, 4) = RS152.Fields("Datum").Value
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop
Else
    ReDim GlSpr(1, 4)
    GlSpr(1, 0) = 1
    GlSpr(1, 1) = 1
    GlSpr(1, 2) = GlSZe
    GlSpr(1, 3) = GlSZe
    GlSpr(1, 4) = Date
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryZahlText ORDER BY IDS"
Else
    SQL1 = "SELECT * FROM qryZahlText ORDER BY [IDS];"
End If
AktZa = 1
Set RS153 = New ADODB.Recordset 'Zahlungstexte
With RS153
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS153.RecordCount
If GesZa > 0 Then
    ReDim GlZTe(GesZa, 4)
    Do Until RS153.EOF
    GlZTe(AktZa, 0) = RS153.Fields("IDS").Value
    If RS153.Fields("IDB").Value <> vbNullString Then GlZTe(AktZa, 1) = RS153.Fields("IDB").Value 'Geldkonto
    If RS153.Fields("IDZ").Value <> vbNullString Then GlZTe(AktZa, 2) = RS153.Fields("IDZ").Value 'Zahlungsziel
    If RS153.Fields("IDKurz").Value <> vbNullString Then GlZTe(AktZa, 3) = RS153.Fields("IDKurz").Value
    GlZTe(AktZa, 4) = CBool(RS153.Fields("Selekt").Value)
    AktZa = AktZa + 1
    RS153.MoveNext
    Loop
Else
    ReDim GlZTe(7, 4)
    GlZTe(1, 0) = 1
    GlZTe(1, 1) = 2
    GlZTe(1, 2) = 0
    GlZTe(1, 3) = "Barzahlung erhalten"
    GlZTe(1, 4) = 0
    GlZTe(2, 0) = 2
    GlZTe(2, 1) = 2
    GlZTe(2, 2) = 0
    GlZTe(2, 3) = "Anzahlung erhalten"
    GlZTe(2, 4) = -1
    GlZTe(3, 0) = 3
    GlZTe(3, 1) = 4
    GlZTe(3, 2) = 0
    GlZTe(3, 3) = "Bankkartenzahlung"
    GlZTe(3, 4) = 0
    GlZTe(4, 0) = 4
    GlZTe(4, 1) = 0
    GlZTe(4, 2) = 0
    GlZTe(4, 3) = "Betrag Verrechnet"
    GlZTe(4, 4) = 0
    GlZTe(5, 0) = 5
    GlZTe(5, 1) = 0
    GlZTe(5, 2) = 0
    GlZTe(5, 3) = "Guthaben verrechnet"
    GlZTe(5, 4) = 0
    GlZTe(6, 0) = 6
    GlZTe(6, 1) = 1
    GlZTe(6, 2) = 0
    GlZTe(6, 3) = "Kreditkartenzahlung"
    GlZTe(6, 4) = 0
    GlZTe(7, 0) = 7
    GlZTe(7, 1) = 0
    GlZTe(7, 2) = 0
    GlZTe(7, 3) = "Onlinebezahlsystem"
    GlZTe(7, 4) = 0
    DoEvents
    For AktZa = 1 To UBound(GlZTe)
        DBCmEx4 "qryZahlAnf", "@IdBnk", "@IdZah", "@IdStr", "@IdSel", GlZTe(AktZa, 1), GlZTe(AktZa, 2), GlZTe(AktZa, 3), GlZTe(AktZa, 4)
    Next AktZa
End If
RS153.Close
Set RS153 = Nothing
DoEvents

'--------------------------------------------------

If GlStL > UBound(GlLab) Then
    GlStL = 1
End If

GlStP = Right$(IniGetVal("Vorgabe", "StaEmp"), 1)  'Standardempfängertyp

Select Case GlStD 'Standard-Dezimaltrennzeichen
Case 0: GlFak = CSng(IniGetVal("System", "AbrFak"))
Case 1: GlFak = Replace(IniGetVal("System", "AbrFak"), ",", ".", 1)
End Select

Select Case GlStD
Case 0:
    GlWa1 = "#,##0.00" 'Deutschland
    GlWa2 = "0,00"
    GlWa3 = "1,00"
Case 1:
    GlWa1 = "0.00" 'Schweiz
    GlWa2 = "0.00"
    GlWa3 = "1.00"
End Select

Select Case GlBut
Case RibTab_Kat_Ketten:
    GlNod = "D" & GlGKa(GlStK, 0)
Case RibTab_Kat_Frage:
    GlNod = "N1"
Case Else:
    GlNod = "A" & GlGKa(GlStK, 0)
End Select
DoEvents

If GlApp = False Then 'AppMode
    S_Ary2k
End If

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2 " & Err.Number
Resume Next

End Sub
Private Sub S_Ary2a()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

ReDim GlMiK(1, 33) 'Alle Mitarbeiter
GlMiK(1, 0) = 1
GlMiK(1, 1) = "Mitarbeitername"
GlMiK(1, 2) = 1
GlMiK(1, 3) = "Name"
GlMiK(1, 4) = "Vorname"
GlMiK(1, 5) = False
GlMiK(1, 6) = GlSZe 'Sprechzietenstring
GlMiK(1, 7) = 0
GlMiK(1, 8) = GlZeR 'Zeitrasterindex
GlMiK(1, 9) = vbNullString
GlMiK(1, 10) = vbNullString
GlMiK(1, 11) = "Emailsignatur"
GlMiK(1, 12) = False
GlMiK(1, 13) = vbNullString
GlMiK(1, 14) = vbNullString
GlMiK(1, 15) = False
GlMiK(1, 16) = 0
GlMiK(1, 17) = 2
GlMiK(1, 18) = vbNullString
GlMiK(1, 19) = GlStR 'Rechtestring
GlMiK(1, 20) = vbNullString
GlMiK(1, 21) = 0
GlMiK(1, 22) = "Emailadresse"
GlMiK(1, 23) = "Titel"
GlMiK(1, 24) = GlSZe 'Sprechzietenstring
GlMiK(1, 25) = "Signaturdatei"
GlMiK(1, 26) = GlZeR 'Onlinezeitrasterindex
GlMiK(1, 27) = vbNullString
GlMiK(1, 28) = "01"
GlMiK(1, 29) = "Heilpraktiker"
GlMiK(1, 30) = vbNullString
GlMiK(1, 31) = "GLN"
GlMiK(1, 32) = "ZSR"
GlMiK(1, 33) = "Briefabsendezeile"

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2a " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2b()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
ReDim GlMiA(1, 33) 'Aktive Mitarbeier
GlMiA(1, 0) = 1
GlMiA(1, 1) = "Mitarbeitername"
GlMiA(1, 2) = 1
GlMiA(1, 3) = "Name"
GlMiA(1, 4) = "Vorname"
GlMiA(1, 5) = False
GlMiA(1, 6) = GlSZe 'Sprechzietenstring
GlMiA(1, 7) = 0
GlMiA(1, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
GlMiA(1, 9) = vbNullString
GlMiA(1, 10) = vbNullString
GlMiA(1, 11) = "Emailsignatur"
GlMiA(1, 12) = False
GlMiA(1, 13) = vbNullString
GlMiA(1, 14) = vbNullString
GlMiA(1, 15) = False
GlMiA(1, 16) = 0
GlMiA(1, 17) = 2
GlMiA(1, 18) = vbNullString
GlMiA(1, 19) = GlStR 'Rechtestring
GlMiA(1, 20) = vbNullString
GlMiA(1, 21) = 0
GlMiA(1, 22) = "Emailadresse"
GlMiA(1, 23) = "Titel"
GlMiA(1, 24) = GlSZe 'Sprechzietenstring
GlMiA(1, 25) = "Signaturdatei"
GlMiA(1, 26) = GlZeR 'Onlinezeitrasterindex
GlMiA(1, 27) = vbNullString
GlMiA(1, 28) = "01"
GlMiA(1, 29) = "Heilpraktiker"
GlMiA(1, 30) = vbNullString
GlMiA(1, 31) = "GLN"
GlMiA(1, 32) = "ZSR"
GlMiA(1, 33) = "Briefabsendezeile"

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Aryb " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2c()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
ReDim GlMiT(1, 33) 'Aktive Mitarbeier + Terminspalte
GlMiT(1, 0) = 1
GlMiT(1, 1) = "Mitarbeitername"
GlMiT(1, 2) = 1
GlMiT(1, 3) = "Name"
GlMiT(1, 4) = "Vorname"
GlMiT(1, 5) = False
GlMiT(1, 6) = GlSZe 'Sprechzietenstring
GlMiT(1, 7) = 0
GlMiT(1, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
GlMiT(1, 9) = vbNullString
GlMiT(1, 10) = vbNullString
GlMiT(1, 11) = "Emailsignatur"
GlMiT(1, 12) = False
GlMiT(1, 13) = vbNullString
GlMiT(1, 14) = vbNullString
GlMiT(1, 15) = False
GlMiT(1, 16) = 0
GlMiT(1, 17) = 2
GlMiT(1, 18) = vbNullString
GlMiT(1, 19) = GlStR 'Rechtestring
GlMiT(1, 20) = vbNullString
GlMiT(1, 21) = 0
GlMiT(1, 22) = "Emailadresse"
GlMiT(1, 23) = "Titel"
GlMiT(1, 24) = GlSZe 'Sprechzietenstring
GlMiT(1, 25) = "Signaturdatei"
GlMiT(1, 26) = GlZeR 'Onlinezeitrasterindex
GlMiT(1, 27) = vbNullString
GlMiT(1, 28) = "01"
GlMiT(1, 29) = "Heilpraktiker"
GlMiT(1, 30) = vbNullString
GlMiT(1, 31) = "GLN"
GlMiT(1, 32) = "ZSR"
GlMiT(1, 33) = "Briefabsendezeile"
    
Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2c " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2d()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
ReDim GlMiO(1, 33) 'Aktive Mitarbeier + Terminspalte + OTS
GlMiO(1, 0) = 1
GlMiO(1, 1) = "Mitarbeitername"
GlMiO(1, 2) = 1
GlMiO(1, 3) = "Name"
GlMiO(1, 4) = "Vorname"
GlMiO(1, 5) = False
GlMiO(1, 6) = GlSZe 'Sprechzietenstring
GlMiO(1, 7) = 0
GlMiO(1, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
GlMiO(1, 9) = vbNullString
GlMiO(1, 10) = vbNullString
GlMiO(1, 11) = "Emailsignatur"
GlMiO(1, 12) = False
GlMiO(1, 13) = vbNullString
GlMiO(1, 14) = vbNullString
GlMiO(1, 15) = False
GlMiO(1, 16) = 0
GlMiO(1, 17) = 2
GlMiO(1, 18) = vbNullString
GlMiO(1, 19) = GlStR 'Rechtestring
GlMiO(1, 20) = vbNullString
GlMiO(1, 21) = 0
GlMiO(1, 22) = "Emailadresse"
GlMiO(1, 23) = "Titel"
GlMiO(1, 24) = GlSZe 'Sprechzietenstring
GlMiO(1, 25) = "Signaturdatei"
GlMiO(1, 26) = GlZeR 'Onlinezeitrasterindex
GlMiO(1, 27) = vbNullString
GlMiO(1, 28) = "01"
GlMiO(1, 29) = "Heilpraktiker"
GlMiO(1, 30) = vbNullString
GlMiO(1, 31) = "GLN"
GlMiO(1, 32) = "ZSR"
GlMiO(1, 33) = "Briefabsendezeile"
    
Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2d " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2e()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
ReDim GlThe(1, 48)
GlThe(1, 0) = 1
GlThe(1, 1) = "Vorname"
GlThe(1, 2) = "Name"
GlThe(1, 3) = "Straße"
GlThe(1, 4) = "PLZ"
GlThe(1, 5) = "Ort"
GlThe(1, 6) = "Telefon"
GlThe(1, 7) = "Telefax"
GlThe(1, 8) = "Bankname"
GlThe(1, 9) = "BLZ"
GlThe(1, 10) = "Konto"
GlThe(1, 11) = "999999"
GlThe(1, 12) = "Beruf"
GlThe(1, 13) = "Mandantenname"
GlThe(1, 14) = "Titel"
GlThe(1, 15) = "LANR"
GlThe(1, 16) = "Keine@Emailadresse.com"
GlThe(1, 17) = "Internet"
GlThe(1, 18) = "IBAN"
GlThe(1, 19) = "Praxis"
GlThe(1, 20) = "Bank2"
GlThe(1, 21) = "BLZ2"
GlThe(1, 22) = "Konto2"
GlThe(1, 23) = "IBAN2"
GlThe(1, 24) = GlSZe 'Sprechzietenstring
GlThe(1, 25) = False
GlThe(1, 26) = GlZeR 'Zeitrasterindex
GlThe(1, 27) = GlZeR 'Zeitrasterindex
GlThe(1, 28) = "180"
GlThe(1, 29) = "80"
GlThe(1, 30) = False
GlThe(1, 31) = "BIC"
GlThe(1, 32) = "BIC2"
GlThe(1, 33) = "GID"
GlThe(1, 34) = "Username"
GlThe(1, 35) = "Passwort"
GlThe(1, 36) = "Absendezeile"
GlThe(1, 37) = "Emailbetreff"
GlThe(1, 38) = "-"
GlThe(1, 39) = GlFri 'Fachrichtung
GlThe(1, 40) = "123456"
GlThe(1, 41) = "123456"
GlThe(1, 42) = GlBsl(0, 0)
GlThe(1, 43) = GlKVB(0, 0)
GlThe(1, 44) = GlKtn(0, 0)
    
Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2e " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2f()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
ReDim GlMan(1, 38)
GlMan(1, 0) = 1
GlMan(1, 1) = "Mandantenname"
GlMan(1, 2) = 1
GlMan(1, 3) = "Name"
GlMan(1, 4) = "Vorname"
GlMan(1, 5) = False
GlMan(1, 6) = GlSZe 'Sprechzietenstring
GlMan(1, 21) = GlSZe 'Sprechzietenstring
GlMan(1, 7) = "LANR"
GlMan(1, 8) = GlZeR 'Zeitrasterindex
GlMan(1, 9) = vbNullString
GlMan(1, 10) = vbNullString
GlMan(1, 11) = False
GlMan(1, 12) = vbNullString
GlMan(1, 13) = vbNullString
GlMan(1, 14) = False
GlMan(1, 15) = vbNullString
GlMan(1, 16) = vbNullString
GlMan(1, 17) = vbNullString
GlMan(1, 18) = False
GlMan(1, 19) = False
GlMan(1, 20) = "Titel"
GlMan(1, 22) = CInt(GlSet(2, 0))
GlMan(1, 23) = CLng(GlSet(2, 1)) 'Standardgebührenkette 1
GlMan(1, 24) = GlStS
GlMan(1, 25) = Right$(GlSet(1, 23), 1)
GlMan(1, 26) = CLng(GlSet(1, 24))
GlMan(1, 27) = CLng(GlSet(1, 25))
GlMan(1, 28) = CInt(GlSet(2, 26))
GlMan(1, 29) = CInt(GlSet(2, 27))
GlMan(1, 30) = Left$(GlSet(1, 30), 1)
GlMan(1, 31) = GlZeR 'Zeitrasterindex
GlMan(1, 35) = 2
GlMan(1, 38) = CLng(GlSet(2, 74)) 'Standardgebührenkette 2
    
Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2f " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2g()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
ReDim GlMaT(1, 38)
GlMaT(1, 0) = 1
GlMaT(1, 1) = "Mandantenname"
GlMaT(1, 2) = 1
GlMaT(1, 3) = "Name"
GlMaT(1, 4) = "Vorname"
GlMaT(1, 5) = False
GlMaT(1, 6) = GlSZe 'Sprechzietenstring
GlMaT(1, 21) = GlSZe 'Sprechzietenstring
GlMaT(1, 7) = "LANR"
GlMaT(1, 8) = GlZeR 'Zeitrasterindex
GlMaT(1, 9) = vbNullString
GlMaT(1, 10) = vbNullString
GlMaT(1, 11) = False
GlMaT(1, 12) = vbNullString
GlMaT(1, 13) = vbNullString
GlMaT(1, 14) = False
GlMaT(1, 15) = vbNullString
GlMaT(1, 16) = vbNullString
GlMaT(1, 17) = vbNullString
GlMaT(1, 18) = False
GlMaT(1, 19) = False
GlMaT(1, 20) = "Titel"
GlMaT(1, 22) = CInt(GlSet(2, 0))
GlMaT(1, 23) = CLng(GlSet(2, 1)) 'Standardgebührenkette 1
GlMaT(1, 24) = GlStS
GlMaT(1, 25) = Right$(GlSet(1, 23), 1)
GlMaT(1, 26) = CLng(GlSet(1, 24))
GlMaT(1, 27) = CLng(GlSet(1, 25))
GlMaT(1, 28) = CInt(GlSet(2, 26))
GlMaT(1, 29) = CInt(GlSet(2, 27))
GlMaT(1, 30) = Left$(GlSet(1, 30), 1)
GlMaT(1, 31) = GlZeR 'Zeitrasterindex
GlMaT(1, 35) = 2
GlMaT(1, 38) = CLng(GlSet(2, 74)) 'Standardgebührenkette 2
    
Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2g " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2h()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
ReDim GlMaO(1, 38)
GlMaO(1, 0) = 1
GlMaO(1, 1) = "Mandantenname"
GlMaO(1, 2) = 1
GlMaO(1, 3) = "Name"
GlMaO(1, 4) = "Vorname"
GlMaO(1, 5) = False
GlMaO(1, 6) = GlSZe 'Sprechzietenstring
GlMaO(1, 21) = GlSZe 'Sprechzietenstring
GlMaO(1, 7) = "LANR"
GlMaO(1, 8) = GlZeR 'Zeitrasterindex
GlMaO(1, 9) = vbNullString
GlMaO(1, 10) = vbNullString
GlMaO(1, 11) = False
GlMaO(1, 12) = vbNullString
GlMaO(1, 13) = vbNullString
GlMaO(1, 14) = False
GlMaO(1, 15) = vbNullString
GlMaO(1, 16) = vbNullString
GlMaO(1, 17) = vbNullString
GlMaO(1, 18) = False
GlMaO(1, 19) = False
GlMaO(1, 20) = "Titel"
GlMaO(1, 22) = CInt(GlSet(2, 0))
GlMaO(1, 23) = CLng(GlSet(2, 1)) 'Standardgebührenkette 1
GlMaO(1, 24) = GlStS
GlMaO(1, 25) = Right$(GlSet(1, 23), 1)
GlMaO(1, 26) = CLng(GlSet(1, 24))
GlMaO(1, 27) = CLng(GlSet(1, 25))
GlMaO(1, 28) = CInt(GlSet(2, 26))
GlMaO(1, 29) = CInt(GlSet(2, 27))
GlMaO(1, 30) = Left$(GlSet(1, 30), 1)
GlMaO(1, 31) = GlZeR 'Zeitrasterindex
GlMaO(1, 35) = 2
GlMaO(1, 38) = CLng(GlSet(2, 74)) 'Standardgebührenkette 2

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2h " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2i()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

Dim AktZa As Integer

ReDim GlEmN(16, 3)
GlEmN(1, 0) = 1
GlEmN(1, 1) = "hiermit bescheinige ich, dass: {PATIENT} am {DATUM} in der Zeit von {STARTZEIT} bis {ENDZEIT} Uhr, in meiner Praxis behandelt wurde."
GlEmN(1, 2) = "Terminbescheinigung (Brief)"
GlEmN(1, 3) = ""
GlEmN(2, 0) = 2
GlEmN(2, 1) = "hiermit bestätige ich die Terminbuchung für: {PATIENT} am {DATUM} in der Zeit von {STARTZEIT} bis {ENDZEIT} Uhr. Sollten Sie diesen Termin nicht wahrnehmen können,  so möchte ich Sie bitten, diesen bis spätestens 24 Stunden vorher abzusagen."
GlEmN(2, 2) = "Terminbestätigung (Brief)"
GlEmN(2, 3) = ""
GlEmN(3, 0) = 3
GlEmN(3, 1) = "hiermit möchte ich Sie an die Terminbuchung für: {PATIENT} am {DATUM} von {STARTZEIT} bis {ENDZEIT} Uhr erinnern. Sollten Sie diesen Termin nicht wahrnehmen können,  so möchte ich Sie bitten, diesen bis spätestens 24 Stunden vorher abzusagen."
GlEmN(3, 2) = "Terminerinnerung (Brief)"
GlEmN(3, 3) = ""
GlEmN(4, 0) = 4
GlEmN(4, 1) = "hiermit möchte ich Ihnen mitteilen, dass der Termin für: {PATIENT} am {DATUM} um {STARTZEIT} Uhr bedauerlicherweise nicht stattfinden kann. Ich möchte Sie daher bitten, in den nächsten Tagen einen neuen Termin zu vereinbaren."
GlEmN(4, 2) = "Terminabsage (Brief)"
GlEmN(4, 3) = ""
GlEmN(5, 0) = 5
GlEmN(5, 1) = "hiermit möchte ich Ihnen {DATUM} von {STARTZEIT} bis {ENDZEIT} Uhr als nächsten Termin für: {PATIENT} vorschlagen. Ich würde mich freuen wenn Sie diesen Termin kurzfristig bestätigen würden."
GlEmN(5, 2) = "Terminvorschlag (Brief)"
GlEmN(5, 3) = ""
GlEmN(6, 0) = 6
GlEmN(6, 1) = "hiermit bestätige ich die Terminbuchung für: {PATIENT} am {DATUM} in der Zeit von {STARTZEIT} bis {ENDZEIT} Uhr. Sollten Sie diesen Termin nicht wahrnehmen können,  so möchte ich Sie bitten, diesen bis spätestens 24 Stunden vorher abzusagen."
GlEmN(6, 2) = "Terminbestätigung (Email)"
GlEmN(6, 3) = "Terminbestätigung - {Mitarbeiter}"
GlEmN(7, 0) = 7
GlEmN(7, 1) = "hiermit möchte ich Sie an die Terminbuchung für: {PATIENT} am {DATUM} von {STARTZEIT} bis {ENDZEIT} Uhr erinnern. Sollten Sie diesen Termin nicht wahrnehmen können,  so möchte ich Sie bitten, diesen bis spätestens 24 Stunden vorher abzusagen."
GlEmN(7, 2) = "Terminerinnerung (Email)"
GlEmN(7, 3) = "Terminerinnerung - {Mitarbeiter}"
GlEmN(8, 0) = 8
GlEmN(8, 1) = "hiermit möchte ich Ihnen mitteilen, dass der Termin für: {PATIENT} am {DATUM} um {STARTZEIT} Uhr bedauerlicherweise nicht stattfinden kann. Ich möchte Sie daher bitten, in den nächsten Tagen einen neuen Termin zu vereinbaren."
GlEmN(8, 2) = "Terminabsage (Email)"
GlEmN(8, 3) = "Terminabsage - {Mitarbeiter}"
GlEmN(9, 0) = 9
GlEmN(9, 1) = "hiermit möchte ich Ihnen {DATUM} von {STARTZEIT} bis {ENDZEIT} Uhr als nächsten Termin für: {PATIENT} vorschlagen. Ich würde mich freuen wenn Sie diesen Termin kurzfristig bestätigen würden."
GlEmN(9, 2) = "Terminvorschlag (Email)"
GlEmN(9, 3) = "Terminvorschlag - {Mitarbeiter}"
GlEmN(10, 0) = 10
GlEmN(10, 1) = "hiermit bestätige ich die Terminbuchung für: {PATIENT} am {DATUM} von {STARTZEIT} bis {ENDZEIT} Uhr."
GlEmN(10, 2) = "Terminbestätigung (SMS)"
GlEmN(10, 3) = ""
GlEmN(11, 0) = 11
GlEmN(11, 1) = "hiermit möchte ich Sie an die Terminbuchung für: {PATIENT} am {DATUM} von {STARTZEIT} bis {ENDZEIT} Uhr erinnern."
GlEmN(11, 2) = "Terminerinnerung (SMS)"
GlEmN(11, 3) = ""
GlEmN(12, 0) = 12
GlEmN(12, 1) = "hiermit möchte ich Ihnen mitteilen, dass der Termin für: {PATIENT} am {DATUM} um {STARTZEIT} Uhr bedauerlicherweise nicht stattfinden kann."
GlEmN(12, 2) = "Terminabsage (SMS)"
GlEmN(12, 3) = ""
GlEmN(13, 0) = 13
GlEmN(13, 1) = "hiermit möchte ich Ihnen den {DATUM} von {STARTZEIT} bis {ENDZEIT} Uhr als nächsten Termin für: {PATIENT} vorschlagen."
GlEmN(13, 2) = "Terminvorschlag (SMS)"
GlEmN(13, 3) = ""
GlEmN(14, 0) = 14
GlEmN(14, 1) = "hiermit bestätige ich die Terminstornierung für: {PATIENT} am {DATUM} in der Zeit von {STARTZEIT} bis {ENDZEIT} Uhr."
GlEmN(14, 2) = "Terminstornierung (Brief)"
GlEmN(14, 3) = ""
GlEmN(15, 0) = 15
GlEmN(15, 1) = "hiermit bestätige ich die Terminstornierung für: {PATIENT} am {DATUM} in der Zeit von {STARTZEIT} bis {ENDZEIT} Uhr."
GlEmN(15, 2) = "Terminstornierung (Email)"
GlEmN(15, 3) = "Terminstornierung - {Mitarbeiter}"
GlEmN(16, 0) = 16
GlEmN(16, 1) = "hiermit bestätige ich die Terminstornierung für: {PATIENT} am {DATUM} in der Zeit von {STARTZEIT} bis {ENDZEIT} Uhr."
GlEmN(16, 2) = "Terminstornierung (SMS)"
GlEmN(16, 3) = ""

For AktZa = 1 To 16 'WICHTIG!
    DBCmEx3 "qryMailTxAd", "@IdStr", "@IdKey", "@IdBet", GlEmN(AktZa, 1), GlEmN(AktZa, 2), GlEmN(AktZa, 3)
Next AktZa

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2i " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2j()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

Dim AktZa As Integer

If GlTyp < 2 Then
    DBCmEx0 "TRUNCATE TABLE dbo.Tabelle_Mail_Betr", True
Else
    DBCmEx0 "DELETE * FROM Tabelle_Mail_Betr;", True
End If

ReDim GlEmT(16, 3) 'WICHTIG!
GlEmT(1, 0) = 1
GlEmT(1, 1) = "anbei übersende ich Ihnen den Beleg: {Rechnung} vom: {Datum} als PDF-Dokument."
GlEmT(1, 2) = "Belegversand"
GlEmT(1, 3) = "{Mandant} - Beleg: {Rechnung}"
GlEmT(2, 0) = 2
GlEmT(2, 1) = "anbei übersende ich Ihnen die Zahlungserinnerung: {Rechnung} als PDF-Dokument"
GlEmT(2, 2) = "Mahnungsversand"
GlEmT(2, 3) = "{Mandant} - Beleg: {Rechnung}"
GlEmT(3, 0) = 3
GlEmT(3, 1) = "anbei übersende ich Ihnen den Fragebogen: {Dokument} als PDF-Dokument."
GlEmT(3, 2) = "Fragebogenversand"
GlEmT(3, 3) = "{Mitarbeiter} - Fragebogenversand"
GlEmT(4, 0) = 4
GlEmT(4, 1) = "anbei übersende ich Ihnen Ihren Laborbericht vom {Datum} als PDF-Dokument."
GlEmT(4, 2) = "Laborberichtversand"
GlEmT(4, 3) = "{Mitarbeiter} - Laborbericht"
GlEmT(5, 0) = 5
GlEmT(5, 1) = "anbei übersende ich Ihnen Ihre aktuellen Termine als PDF-Dokument."
GlEmT(5, 2) = "Terminzettelversand"
GlEmT(5, 3) = "{Mitarbeiter} - Termine"
GlEmT(6, 0) = 6
GlEmT(6, 1) = "anbei übersende ich Ihnen die Ihr Rezept vom: {Datum} als PDF-Dokument."
GlEmT(6, 2) = "Rezeptversand"
GlEmT(6, 3) = "{Mandant} - Rezept"
GlEmT(7, 0) = 7
GlEmT(7, 1) = "anbei übersende ich Ihnen die Ihr Beleg vom: {Datum} als PDF-Dokument."
GlEmT(7, 2) = "Belegversand"
GlEmT(7, 3) = "{Mandant} - Beleg"
GlEmT(8, 0) = 8
GlEmT(8, 1) = "anbei übersende ich Ihnen die Ihr Rezept vom: {Datum} als PDF-Dokument."
GlEmT(8, 2) = "Langrezeptversand"
GlEmT(8, 3) = "{Mandant} - Rezept"
GlEmT(9, 0) = 9
GlEmT(9, 1) = "anbei übersende ich Ihnen das Dokument: {Dokument} vom: {Datum} als PDF-Dokument"
GlEmT(9, 2) = "Dokumentenversand"
GlEmT(9, 3) = "{Mitarbeiter} - Dokumentenversand"
GlEmT(10, 0) = 10
GlEmT(10, 1) = "anbei übersende ich Ihnen Ihre aktuelle Patientenauskunft als PDF-Dokument."
GlEmT(10, 2) = "Krankenblattversand"
GlEmT(10, 3) = "{Mandant} - Patientenauskunft"
GlEmT(11, 0) = 11
GlEmT(11, 1) = "Ihr Dokumenten Entschlüsselungskennwort lautet: {Passwort}"
GlEmT(11, 2) = "Entschlüsselungskennwort"
GlEmT(11, 3) = "{Mitarbeiter} - E-Mail-Versand"
GlEmT(12, 0) = 12
GlEmT(12, 1) = "Ihre Rechnung: {Rechnung} vom: {Datum} liegt für Sie in den nächsten {Gültigkeit} Tagen unter dem folgenden Link zum Download bereit: {CR}{Direktlink}{CR}{CR}Ältere Rechnungen finden Sie unter dem folgenden Link:{CR}{Portallink}"
GlEmT(12, 2) = "Rechnungsdownload"
GlEmT(12, 3) = "{Mandant} - Beleg: {Rechnung}"
GlEmT(13, 0) = 13
GlEmT(13, 1) = "Ihr aktuelles Dokument vom: {Datum} liegt für Sie in den nächsten {Gültigkeit} Tagen unter dem folgenden Link zum Download bereit: {CR}{Direktlink}{CR}{CR}Ältere Dokumente finden Sie unter dem folgenden Link:{CR}{Portallink}"
GlEmT(13, 2) = "Dokumentendownload"
GlEmT(13, 3) = "{Mitarbeiter} - Praxisdokument"
GlEmT(14, 0) = 14
GlEmT(14, 1) = "anbei übersende ich Ihnen ein zu signierendes Dokument. Dieses liegt für Sie in den nächsten {Gültigkeit} Tagen unter dem folgenden Link zur Unterschrift bereit:{CR}{Direktlink}"
GlEmT(14, 2) = "Digitale Unterschrift"
GlEmT(14, 3) = "{Mandant} - Digitale Unterschrift"
GlEmT(15, 0) = 15
GlEmT(15, 1) = "Ihr Rezept: {Rechnung} vom: {Datum} liegt für Sie in den nächsten {Gültigkeit} Tagen unter dem folgenden Link zum Download bereit: {CR}{Direktlink}{CR}{CR}Ältere Rezepte finden Sie unter dem folgenden Link:{CR}{Portallink}"
GlEmT(15, 2) = "Rezepzdownload"
GlEmT(15, 3) = "{Mandant} - Rezept"
GlEmT(16, 0) = 16
GlEmT(16, 1) = "Ihr Beleg: {Rechnung} vom: {Datum} liegt für Sie in den nächsten {Gültigkeit} Tagen unter dem folgenden Link zum Download bereit: {CR}{Direktlink}{CR}{CR}Ältere Belege finden Sie unter dem folgenden Link:{CR}{Portallink}"
GlEmT(16, 2) = "Belegdownload"
GlEmT(16, 3) = "{Mandant} - Beleg"

For AktZa = 1 To 16 'WICHTIG!
    DBCmEx3 "qryMailBeAd", "@IdStr", "@IdKey", "@IdBet", GlEmT(AktZa, 1), GlEmT(AktZa, 2), GlEmT(AktZa, 3)
Next AktZa

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2j " & Err.Number
Resume Next
    
End Sub
Private Sub S_Ary2k()
On Error GoTo FiErr
'Füllt die Daten in die Arrays
    
Dim AktZa As Integer
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim EvCat As XtremeCalendarControl.CalendarEventCategory
Dim EvCas As XtremeCalendarControl.CalendarEventCategories
    
Set FM = frmMain
Set CaCol = FM.calCont1
Set DaPro = CaCol.DataProvider
Set CaLbs = DaPro.LabelList
Set EvCas = DaPro.EventCategories
    
With CaLbs 'Terminfarben / Terminbetreffs
    .removeAll
    .AddLabel 1, GlTmF(1, 1), vbNullString   'weiss
    .AddLabel 2, GlTmF(2, 1), vbNullString   'hellrot
    .AddLabel 3, GlTmF(3, 1), vbNullString    'hellorange
    .AddLabel 4, GlTmF(4, 1), vbNullString    'hellgelb
    .AddLabel 5, GlTmF(5, 1), vbNullString    'hellgrün
    .AddLabel 6, GlTmF(6, 1), vbNullString    'helltürkis
    .AddLabel 7, GlTmF(7, 1), vbNullString    'hellblau
    .AddLabel 8, GlTmF(8, 1), vbNullString    'hellrosa
    .AddLabel 9, GlTmF(9, 1), vbNullString    'rot
    .AddLabel 10, GlTmF(10, 1), vbNullString  'orange
    .AddLabel 11, GlTmF(11, 1), vbNullString  'gelb
    .AddLabel 12, GlTmF(12, 1), vbNullString  'mosgrün
    .AddLabel 13, GlTmF(13, 1), vbNullString  'türkis
    .AddLabel 14, GlTmF(14, 1), vbNullString  'graublau
    .AddLabel 15, GlTmF(15, 1), vbNullString  'mangenta
    .AddLabel 16, GlTmF(16, 1), vbNullString  'hellgraublau
    .AddLabel 17, GlTmF(17, 1), vbNullString  'gelborange
    .AddLabel 18, GlTmF(18, 1), vbNullString  'orange
    .AddLabel 19, GlTmF(19, 1), vbNullString  'grün
    .AddLabel 20, GlTmF(20, 1), vbNullString  'grau
    For AktZa = 1 To UBound(GlBtr)
        .AddLabel AktZa + 20, GlBtr(AktZa, 4), vbNullString
    Next AktZa
End With

Set CaLbs = Nothing
Set EvCas = Nothing
Set DaPro = Nothing
Set CaCol = Nothing
    
Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary2k " & Err.Number
Resume Next
    
End Sub

Public Sub S_Ary3()
On Error GoTo FiErr
'Füllt die Daten in die Arrays

Dim SQL1 As String
Dim MaAbs As String
Dim EmBet As String
Dim AktZa As Integer
Dim GesZa As Integer
Dim AnMiA As Integer
Dim AnMiT As Integer
Dim AnMiO As Integer
Dim AnMaA As Integer
Dim AnMaT As Integer
Dim AnMaO As Integer

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatBeh ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryPatBeh ORDER BY [IDKurz];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Therapeuten Index
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    ReDim GlMan(GesZa, 38)
    Do
    GlMan(AktZa, 0) = AktZa
    If RS152.Fields("IDKurz").Value <> vbNullString Then GlMan(AktZa, 1) = RS152.Fields("IDKurz").Value
    If RS152.Fields("ID0").Value > 0 Then GlMan(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
    If RS152.Fields("Name").Value <> vbNullString Then GlMan(AktZa, 3) = RS152.Fields("Name").Value
    If RS152.Fields("Vorname").Value <> vbNullString Then GlMan(AktZa, 4) = RS152.Fields("Vorname").Value
    GlMan(AktZa, 5) = CBool(RS152.Fields("Passiv").Value)
    If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
        If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
            GlMan(AktZa, 6) = GlSZe 'Sprechzietenstring
        Else
            GlMan(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
        End If
    Else
        GlMan(AktZa, 6) = GlSZe 'Sprechzietenstring
    End If
    If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
        GlMan(AktZa, 21) = RS152.Fields("Buchungszeiten").Value
    Else
        GlMan(AktZa, 21) = GlSZe 'Buchungszeitenstring
    End If
    If RS152.Fields("OnlRas").Value <> vbNullString Then
        GlMan(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
    Else
        GlMan(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
    End If
    If RS152.Fields("KVNummer").Value <> vbNullString Then GlMan(AktZa, 7) = RS152.Fields("KVNummer").Value
    If RS152.Fields("Größe").Value <> vbNullString Then GlMan(AktZa, 9) = RS152.Fields("Größe").Value
    If RS152.Fields("Gewicht").Value <> vbNullString Then GlMan(AktZa, 10) = RS152.Fields("Gewicht").Value
    If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMan(AktZa, 11) = RS152.Fields("Gesperrt").Value
    If RS152.Fields("Em_User").Value <> vbNullString Then GlMan(AktZa, 12) = RS152.Fields("Em_User").Value
    If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMan(AktZa, 13) = RS152.Fields("Em_Pass").Value
    If RS152.Fields("Versand").Value <> vbNullString Then GlMan(AktZa, 14) = RS152.Fields("Versand").Value
    If RS152.Fields("OnlMax").Value <> vbNullString Then GlMan(AktZa, 15) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
    If RS152.Fields("OnlVor").Value <> vbNullString Then GlMan(AktZa, 16) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
    If RS152.Fields("GuiID").Value <> vbNullString Then GlMan(AktZa, 17) = RS152.Fields("GuiID").Value 'UserID
    If RS152.Fields("OnlTer").Value <> vbNullString Then GlMan(AktZa, 18) = RS152.Fields("OnlTer").Value 'Online-Terminbuchungs System
    If RS152.Fields("Telefon5").Value <> vbNullString Then
        GlMan(AktZa, 19) = RS152.Fields("Telefon5").Value
    Else
        GlMan(AktZa, 19) = "keine@emailadresse.de"
    End If
    If RS152.Fields("Titel").Value <> vbNullString Then GlMan(AktZa, 20) = RS152.Fields("Titel").Value
    If RS152.Fields("StaGeb").Value <> vbNullString Then 'Standardgebührenkatalog
        If RS152.Fields("StaGeb").Value > 0 Then
            GlMan(AktZa, 22) = CInt(RS152.Fields("StaGeb").Value)
        Else
            GlMan(AktZa, 22) = CInt(GlSet(2, 0))
        End If
    Else
        GlMan(AktZa, 22) = CInt(GlSet(2, 0))
    End If
    If RS152.Fields("StaKet").Value <> vbNullString Then 'Standardgebührenkette 1
        If RS152.Fields("StaKet").Value > 0 Then
            GlMan(AktZa, 23) = CLng(RS152.Fields("StaKet").Value)
        Else
            GlMan(AktZa, 23) = CLng(GlSet(2, 1))
        End If
    Else
        GlMan(AktZa, 23) = CLng(GlSet(2, 1))
    End If
    If RS152.Fields("StaStu").Value <> vbNullString Then 'Standardsteuersatz
        If RS152.Fields("StaStu").Value > 0 Then
            GlMan(AktZa, 24) = CInt(RS152.Fields("StaStu").Value)
        Else
            GlMan(AktZa, 24) = CInt(GlSet(2, 64))
        End If
    Else
        GlMan(AktZa, 24) = CInt(GlSet(2, 64))
    End If
    If RS152.Fields("StaRam").Value <> vbNullString Then 'Kontenragmen
        If RS152.Fields("StaRam").Value > 0 Then
            GlMan(AktZa, 25) = CInt(RS152.Fields("StaRam").Value)
        Else
            GlMan(AktZa, 25) = GlKtR
        End If
    Else
        GlMan(AktZa, 25) = GlKtR
    End If
    If RS152.Fields("StaKon").Value <> vbNullString Then 'Standarderlöskonto (Bank)
        If RS152.Fields("StaKon").Value > 0 Then
            GlMan(AktZa, 26) = RS152.Fields("StaKon").Value
        Else
            GlMan(AktZa, 26) = GlSE2
        End If
    Else
        GlMan(AktZa, 26) = GlSE2
    End If
    If RS152.Fields("StaKo2").Value <> vbNullString Then 'Standarderlöskonto (Kasse)
        If RS152.Fields("StaKo2").Value > 0 Then
            GlMan(AktZa, 27) = RS152.Fields("StaKo2").Value
        Else
            GlMan(AktZa, 27) = GlSE1
        End If
    Else
        GlMan(AktZa, 27) = GlSE1
    End If
    If RS152.Fields("StaGk1").Value <> vbNullString Then 'Standardgeldkonto (Bank)
        If RS152.Fields("StaGk1").Value > 0 Then
            GlMan(AktZa, 28) = RS152.Fields("StaGk1").Value
        Else
            GlMan(AktZa, 28) = GlGkB
        End If
    Else
        GlMan(AktZa, 28) = GlGkB
    End If
    If RS152.Fields("StaGk2").Value <> vbNullString Then 'Standardgeldkonto (Kasse)
        If RS152.Fields("StaGk2").Value > 0 Then
            GlMan(AktZa, 29) = RS152.Fields("StaGk2").Value
        Else
            GlMan(AktZa, 29) = GlGkK
        End If
    Else
        GlMan(AktZa, 29) = GlGkK
    End If
    If RS152.Fields("SteRet").Value <> vbNullString Then 'Standardbelegtyp
        If IsNull(RS152.Fields("SteRet").Value) = False Then
            GlMan(AktZa, 30) = Left$(RS152.Fields("SteRet").Value, 1)
        Else
            GlMan(AktZa, 30) = Left$(GlSet(1, 30), 1)
        End If
    Else
        GlMan(AktZa, 30) = Left$(GlSet(1, 30), 1)
    End If
    If RS152.Fields("OnlRa2").Value <> vbNullString Then
        GlMan(AktZa, 31) = Format$(RS152.Fields("OnlRa2").Value, "00")
    Else
        GlMan(AktZa, 31) = Format$(GlZeR, "00") 'Zeitrasterindex
    End If
    If RS152.Fields("Firma2").Value <> vbNullString Then GlMan(AktZa, 32) = RS152.Fields("Firma2").Value 'Verkehrsname
    If RS152.Fields("Objekt").Value <> vbNullString Then GlMan(AktZa, 33) = RS152.Fields("Objekt").Value 'Signatur
    If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMan(AktZa, 34) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
    If RS152.Fields("ID3").Value <> vbNullString Then
        GlMan(AktZa, 35) = RS152.Fields("ID3").Value
    Else
        GlMan(AktZa, 35) = GlFri 'Fachrichtung
    End If
    If RS152.Fields("OnlRa1").Value <> vbNullString Then
        GlMan(AktZa, 36) = Format$(RS152.Fields("OnlRa1").Value, "00")
    Else
        GlMan(AktZa, 36) = "12"
    End If
    If RS152.Fields("StStKt").Value <> vbNullString Then 'Standardsteuerkonto
        If RS152.Fields("StStKt").Value > 0 Then
            GlMan(AktZa, 37) = RS152.Fields("StStKt").Value
        Else
            GlMan(AktZa, 37) = GlSKo
        End If
    Else
        GlMan(AktZa, 37) = GlSKo
    End If
    If CBool(RS152.Fields("Passiv").Value) = False Then
        AnMaA = AnMaA + 1 'Anzahl aktiver Mandanten
        If CBool(RS152.Fields("Gesperrt").Value) = False Then
            AnMaT = AnMaT + 1 'Anzahl aktiver Mandanten + Terminspalte
            If RS152.Fields("OnlTer").Value <> vbNullString Then
                If CBool(RS152.Fields("OnlTer").Value) = True Then
                    AnMaO = AnMaO + 1 'Anzahl aktiver Mandanten + Terminspalte + OTS
                End If
            End If
        End If
    End If
    If RS152.Fields("Kanton").Value <> vbNullString Then 'Standardgebührenkette 2
        If RS152.Fields("Kanton").Value > 0 Then
            GlMan(AktZa, 38) = CLng(RS152.Fields("Kanton").Value)
        Else
            GlMan(AktZa, 38) = CLng(GlSet(2, 74))
        End If
    Else
        GlMan(AktZa, 38) = CLng(GlSet(2, 74))
    End If
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop Until RS152.EOF

    AktZa = 1
    RS152.MoveFirst
    
    If AnMaA > 0 Then
        ReDim GlMaA(AnMaA, 38)
        Do
        If CBool(RS152.Fields("Passiv").Value) = False Then
            GlMaA(AktZa, 0) = AktZa
            If RS152.Fields("IDKurz").Value <> vbNullString Then GlMaA(AktZa, 1) = RS152.Fields("IDKurz").Value
            If RS152.Fields("ID0").Value > 0 Then GlMaA(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
            If RS152.Fields("Name").Value <> vbNullString Then GlMaA(AktZa, 3) = RS152.Fields("Name").Value
            If RS152.Fields("Vorname").Value <> vbNullString Then GlMaA(AktZa, 4) = RS152.Fields("Vorname").Value
            GlMaA(AktZa, 5) = CBool(RS152.Fields("Passiv").Value)
            If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
                If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
                    GlMaA(AktZa, 6) = GlSZe 'Sprechzietenstring
                Else
                    GlMaA(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
                End If
            Else
                GlMaA(AktZa, 6) = GlSZe 'Sprechzietenstring
            End If
            If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
                GlMaA(AktZa, 21) = RS152.Fields("Buchungszeiten").Value
            Else
                GlMaA(AktZa, 21) = GlSZe 'Buchungszeitenstring
            End If
            If RS152.Fields("OnlRas").Value <> vbNullString Then
                GlMaA(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
            Else
                GlMaA(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
            End If
            If RS152.Fields("KVNummer").Value <> vbNullString Then GlMaA(AktZa, 7) = RS152.Fields("KVNummer").Value
            If RS152.Fields("Größe").Value <> vbNullString Then GlMaA(AktZa, 9) = RS152.Fields("Größe").Value
            If RS152.Fields("Gewicht").Value <> vbNullString Then GlMaA(AktZa, 10) = RS152.Fields("Gewicht").Value
            If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMaA(AktZa, 11) = RS152.Fields("Gesperrt").Value
            If RS152.Fields("Em_User").Value <> vbNullString Then GlMaA(AktZa, 12) = RS152.Fields("Em_User").Value
            If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMaA(AktZa, 13) = RS152.Fields("Em_Pass").Value
            If RS152.Fields("Versand").Value <> vbNullString Then GlMaA(AktZa, 14) = RS152.Fields("Versand").Value
            If RS152.Fields("OnlMax").Value <> vbNullString Then GlMaA(AktZa, 15) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
            If RS152.Fields("OnlVor").Value <> vbNullString Then GlMaA(AktZa, 16) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
            If RS152.Fields("GuiID").Value <> vbNullString Then GlMaA(AktZa, 17) = RS152.Fields("GuiID").Value 'UserID
            If RS152.Fields("OnlTer").Value <> vbNullString Then GlMaA(AktZa, 18) = RS152.Fields("OnlTer").Value 'Online-Terminbuchungs System
            If RS152.Fields("Telefon5").Value <> vbNullString Then
                GlMaA(AktZa, 19) = RS152.Fields("Telefon5").Value
            Else
                GlMaA(AktZa, 19) = "keine@emailadresse.de"
            End If
            If RS152.Fields("Titel").Value <> vbNullString Then GlMaA(AktZa, 20) = RS152.Fields("Titel").Value
            If RS152.Fields("StaGeb").Value <> vbNullString Then 'Standardgebührenkatalog
                If RS152.Fields("StaGeb").Value > 0 Then
                    GlMaA(AktZa, 22) = CInt(RS152.Fields("StaGeb").Value)
                Else
                    GlMaA(AktZa, 22) = CInt(GlSet(2, 0))
                End If
            Else
                GlMaA(AktZa, 22) = CInt(GlSet(2, 0))
            End If
            If RS152.Fields("StaKet").Value <> vbNullString Then 'Standardgebührenkette
                If RS152.Fields("StaKet").Value > 0 Then
                    GlMaA(AktZa, 23) = CLng(RS152.Fields("StaKet").Value)
                Else
                    GlMaA(AktZa, 23) = CLng(GlSet(2, 1))
                End If
            Else
                GlMaA(AktZa, 23) = CLng(GlSet(2, 1))
            End If
            If RS152.Fields("StaStu").Value <> vbNullString Then 'Standardsteuersatz
                If RS152.Fields("StaStu").Value > 0 Then
                    GlMaA(AktZa, 24) = CInt(RS152.Fields("StaStu").Value)
                Else
                    GlMaA(AktZa, 24) = CInt(GlSet(2, 64))
                End If
            Else
                GlMaA(AktZa, 24) = CInt(GlSet(2, 64))
            End If
            If RS152.Fields("StaRam").Value <> vbNullString Then 'Standardkontenrahmen
                If RS152.Fields("StaRam").Value > 0 Then
                    GlMaA(AktZa, 25) = CInt(RS152.Fields("StaRam").Value)
                Else
                    GlMaA(AktZa, 25) = GlKtR
                End If
            Else
                GlMaA(AktZa, 25) = GlKtR
            End If
            If RS152.Fields("StaKon").Value <> vbNullString Then 'Standarderlöskonto (Bankkonto)
                If RS152.Fields("StaKon").Value > 0 Then
                    GlMaA(AktZa, 26) = RS152.Fields("StaKon").Value
                Else
                    GlMaA(AktZa, 26) = GlSE2
                End If
            Else
                GlMaA(AktZa, 26) = GlSE2
            End If
            If RS152.Fields("StaKo2").Value <> vbNullString Then 'Standarderlöskonto (Kasse)
                If RS152.Fields("StaKo2").Value > 0 Then
                    GlMaA(AktZa, 27) = RS152.Fields("StaKo2").Value
                Else
                    GlMaA(AktZa, 27) = GlSE1
                End If
            Else
                GlMaA(AktZa, 27) = GlSE1
            End If
            If RS152.Fields("StaGk1").Value <> vbNullString Then 'Standardgeldkonto (Bankkonto)
                If RS152.Fields("StaGk1").Value > 0 Then
                    GlMaA(AktZa, 28) = CLng(RS152.Fields("StaGk1").Value)
                Else
                    GlMaA(AktZa, 28) = GlGkB
                End If
            Else
                GlMaA(AktZa, 28) = GlGkB
            End If
            If RS152.Fields("StaGk2").Value <> vbNullString Then 'Standardgeldkonto (Kasse)
                If RS152.Fields("StaGk2").Value > 0 Then
                    GlMaA(AktZa, 29) = CLng(RS152.Fields("StaGk2").Value)
                Else
                    GlMaA(AktZa, 29) = GlGkK
                End If
            Else
                GlMaA(AktZa, 29) = GlGkK
            End If
            If RS152.Fields("SteRet").Value <> vbNullString Then 'Standardbelegtyp
                If IsNull(RS152.Fields("SteRet").Value) = False Then
                    GlMaA(AktZa, 30) = Left$(RS152.Fields("SteRet").Value, 1)
                Else
                    GlMaA(AktZa, 30) = Left$(GlSet(1, 30), 1)
                End If
            Else
                GlMaA(AktZa, 30) = Left$(GlSet(1, 30), 1)
            End If
            If RS152.Fields("OnlRa2").Value <> vbNullString Then
                GlMaA(AktZa, 31) = Format$(RS152.Fields("OnlRa2").Value, "00")
            Else
                GlMaA(AktZa, 31) = Format$(GlZeR, "00") 'Zeitrasterindex
            End If
            If RS152.Fields("Firma2").Value <> vbNullString Then GlMaA(AktZa, 32) = RS152.Fields("Firma2").Value 'Verkehrsname
            If RS152.Fields("Objekt").Value <> vbNullString Then GlMaA(AktZa, 33) = RS152.Fields("Objekt").Value 'Signatur
            If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMaA(AktZa, 34) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
            If RS152.Fields("ID3").Value <> vbNullString Then
                GlMaA(AktZa, 35) = RS152.Fields("ID3").Value
            Else
                GlMaA(AktZa, 35) = GlFri 'Fachrichtung
            End If
            If RS152.Fields("OnlRa1").Value <> vbNullString Then
                GlMaA(AktZa, 36) = Format$(RS152.Fields("OnlRa1").Value, "00")
            Else
                GlMaA(AktZa, 36) = "12"
            End If
            If RS152.Fields("StStKt").Value <> vbNullString Then 'Standarderlöskonto (Bankkonto)
                If RS152.Fields("StStKt").Value > 0 Then
                    GlMaA(AktZa, 37) = SBuFo(CLng(RS152.Fields("StStKt").Value))
                Else
                    GlMaA(AktZa, 37) = GlSKo
                End If
            Else
                GlMaA(AktZa, 37) = GlSKo
            End If
            If RS152.Fields("Kanton").Value <> vbNullString Then 'Standardgebührenkette 2
                If RS152.Fields("Kanton").Value > 0 Then
                    If CLng(RS152.Fields("Kanton").Value) < 1000 Then
                        GlMaA(AktZa, 38) = CInt(RS152.Fields("Kanton").Value)
                    Else
                        GlMaA(AktZa, 38) = CLng(GlSet(2, 74))
                    End If
                Else
                    GlMaA(AktZa, 38) = CLng(GlSet(2, 74))
                End If
            Else
                GlMaA(AktZa, 38) = CLng(GlSet(2, 74))
            End If
            AktZa = AktZa + 1
        End If
        RS152.MoveNext
        Loop Until RS152.EOF
    Else
        S_Ary2g
    End If

    AktZa = 1
    RS152.MoveFirst
    
    If AnMaT > 0 Then
        ReDim GlMaT(AnMaT, 38)
        Do
        If CBool(RS152.Fields("Passiv").Value) = False Then
            If CBool(RS152.Fields("Gesperrt").Value) = False Then
                GlMaT(AktZa, 0) = AktZa
                If RS152.Fields("IDKurz").Value <> vbNullString Then GlMaT(AktZa, 1) = RS152.Fields("IDKurz").Value
                GlMaT(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
                If RS152.Fields("Name").Value <> vbNullString Then GlMaT(AktZa, 3) = RS152.Fields("Name").Value
                If RS152.Fields("Vorname").Value <> vbNullString Then GlMaT(AktZa, 4) = RS152.Fields("Vorname").Value
                GlMaT(AktZa, 5) = CBool(RS152.Fields("Passiv").Value)
                If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
                    If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
                        GlMaT(AktZa, 6) = GlSZe 'Sprechzietenstring
                    Else
                        GlMaT(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
                    End If
                Else
                    GlMaT(AktZa, 6) = GlSZe 'Sprechzietenstring
                End If
                If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
                    GlMaT(AktZa, 21) = RS152.Fields("Buchungszeiten").Value
                Else
                    GlMaT(AktZa, 21) = GlSZe 'Sprechzietenstring
                End If
                If RS152.Fields("OnlRas").Value <> vbNullString Then
                    GlMaT(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
                Else
                    GlMaT(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
                End If
                If RS152.Fields("KVNummer").Value <> vbNullString Then GlMaT(AktZa, 7) = RS152.Fields("KVNummer").Value
                If RS152.Fields("Größe").Value <> vbNullString Then GlMaT(AktZa, 9) = RS152.Fields("Größe").Value
                If RS152.Fields("Gewicht").Value <> vbNullString Then GlMaT(AktZa, 10) = RS152.Fields("Gewicht").Value
                If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMaT(AktZa, 11) = RS152.Fields("Gesperrt").Value
                If RS152.Fields("Em_User").Value <> vbNullString Then GlMaT(AktZa, 12) = RS152.Fields("Em_User").Value
                If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMaT(AktZa, 13) = RS152.Fields("Em_Pass").Value
                If RS152.Fields("Versand").Value <> vbNullString Then GlMaT(AktZa, 14) = RS152.Fields("Versand").Value
                If RS152.Fields("OnlMax").Value <> vbNullString Then GlMaT(AktZa, 15) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
                If RS152.Fields("OnlVor").Value <> vbNullString Then GlMaT(AktZa, 16) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
                If RS152.Fields("GuiID").Value <> vbNullString Then GlMaT(AktZa, 17) = RS152.Fields("GuiID").Value 'UserID
                If RS152.Fields("OnlTer").Value <> vbNullString Then GlMaT(AktZa, 18) = RS152.Fields("OnlTer").Value 'Online-Terminbuchungs System
                If RS152.Fields("Telefon5").Value <> vbNullString Then
                    GlMaT(AktZa, 19) = RS152.Fields("Telefon5").Value
                Else
                    GlMaT(AktZa, 19) = "keine@emailadresse.de"
                End If
                If RS152.Fields("Titel").Value <> vbNullString Then GlMaT(AktZa, 20) = RS152.Fields("Titel").Value
                If RS152.Fields("OnlRa2").Value <> vbNullString Then
                    GlMaT(AktZa, 22) = Format$(RS152.Fields("OnlRa2").Value, "00")
                Else
                    GlMaT(AktZa, 22) = Format$(GlZeR, "00") 'Zeitrasterindex
                End If
                If RS152.Fields("StaKet").Value <> vbNullString Then 'Standardgebührenkette 1
                    If RS152.Fields("StaKet").Value > 0 Then
                        GlMaT(AktZa, 23) = CLng(RS152.Fields("StaKet").Value)
                    Else
                        GlMaT(AktZa, 23) = CLng(GlSet(2, 1))
                    End If
                Else
                    GlMaT(AktZa, 23) = CLng(GlSet(2, 1))
                End If
                If RS152.Fields("OnlRa1").Value <> vbNullString Then
                    GlMaT(AktZa, 24) = Format$(RS152.Fields("OnlRa1").Value, "00")
                Else
                    GlMaT(AktZa, 24) = "12"
                End If
                If RS152.Fields("OnlTmp").Value <> vbNullString Then
                    If CInt(RS152.Fields("OnlTmp").Value) > 0 Then
                        GlMaT(AktZa, 25) = RS152.Fields("OnlTmp").Value
                    Else
                        GlMaT(AktZa, 25) = 24
                    End If
                Else
                    GlMaT(AktZa, 25) = 24
                End If
                If RS152.Fields("StaKon").Value <> vbNullString Then 'Standarderlöskonto (Bankkonto)
                    If RS152.Fields("StaKon").Value > 0 Then
                        GlMaT(AktZa, 26) = RS152.Fields("StaKon").Value
                    Else
                        GlMaT(AktZa, 26) = GlSE2
                    End If
                Else
                    GlMaT(AktZa, 26) = GlSE2
                End If
                If RS152.Fields("StaKo2").Value <> vbNullString Then 'Standarderlöskonto (Kasse)
                    If RS152.Fields("StaKo2").Value > 0 Then
                        GlMaT(AktZa, 27) = RS152.Fields("StaKo2").Value
                    Else
                        GlMaT(AktZa, 27) = GlSE1
                    End If
                Else
                    GlMaT(AktZa, 27) = GlSE1
                End If
                If RS152.Fields("StaGk1").Value <> vbNullString Then 'Standardgeldkonto (Bankkonto)
                    If RS152.Fields("StaGk1").Value > 0 Then
                        GlMaT(AktZa, 28) = CLng(RS152.Fields("StaGk1").Value)
                    Else
                        GlMaT(AktZa, 28) = GlGkB
                    End If
                Else
                    GlMaT(AktZa, 28) = GlGkB
                End If
                If RS152.Fields("StaGk2").Value <> vbNullString Then 'Standardgeldkonto (Kasse)
                    If RS152.Fields("StaGk2").Value > 0 Then
                        GlMaT(AktZa, 29) = CLng(RS152.Fields("StaGk2").Value)
                    Else
                        GlMaT(AktZa, 29) = GlGkK
                    End If
                Else
                    GlMaT(AktZa, 29) = GlGkK
                End If
                If RS152.Fields("SteRet").Value <> vbNullString Then 'Standardbelegtyp
                    If IsNull(RS152.Fields("SteRet").Value) = False Then
                        GlMaT(AktZa, 30) = Left$(RS152.Fields("SteRet").Value, 1)
                    Else
                        GlMaT(AktZa, 30) = Left$(GlSet(1, 30), 1)
                    End If
                Else
                    GlMaT(AktZa, 30) = Left$(GlSet(1, 30), 1)
                End If
                If RS152.Fields("OnlRa2").Value <> vbNullString Then
                    GlMaT(AktZa, 31) = Format$(RS152.Fields("OnlRa2").Value, "00")
                Else
                    GlMaT(AktZa, 31) = Format$(GlZeR, "00") 'Zeitrasterindex
                End If
                If RS152.Fields("Firma2").Value <> vbNullString Then GlMaT(AktZa, 32) = RS152.Fields("Firma2").Value 'Verkehrsname
                If RS152.Fields("Objekt").Value <> vbNullString Then GlMaT(AktZa, 33) = RS152.Fields("Objekt").Value 'Signatur
                If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMaT(AktZa, 34) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
                If RS152.Fields("ID3").Value <> vbNullString Then
                    GlMaT(AktZa, 35) = RS152.Fields("ID3").Value
                Else
                    GlMaT(AktZa, 35) = GlFri 'Fachrichtung
                End If
                If RS152.Fields("OnlRa1").Value <> vbNullString Then
                    GlMaT(AktZa, 36) = Format$(RS152.Fields("OnlRa1").Value, "00")
                Else
                    GlMaT(AktZa, 36) = "12"
                End If
                If RS152.Fields("StStKt").Value <> vbNullString Then 'Standarderlöskonto (Bankkonto)
                    If RS152.Fields("StStKt").Value > 0 Then
                        GlMaT(AktZa, 37) = SBuFo(CLng(RS152.Fields("StStKt").Value))
                    Else
                        GlMaT(AktZa, 37) = GlSKo
                    End If
                Else
                    GlMaT(AktZa, 37) = GlSKo
                End If
                If RS152.Fields("Kanton").Value <> vbNullString Then 'Standardgebührenkette 2
                    If RS152.Fields("Kanton").Value > 0 Then
                        If CLng(RS152.Fields("Kanton").Value) < 1000 Then
                            GlMaT(AktZa, 38) = CInt(RS152.Fields("Kanton").Value)
                        Else
                            GlMaT(AktZa, 38) = CLng(GlSet(2, 74))
                        End If
                    Else
                        GlMaT(AktZa, 38) = CLng(GlSet(2, 74))
                    End If
                Else
                    GlMaT(AktZa, 38) = CLng(GlSet(2, 74))
                End If
                AktZa = AktZa + 1
            End If
        End If
        RS152.MoveNext
        Loop Until RS152.EOF
    Else
        S_Ary2g
    End If
    
    AktZa = 1
    RS152.MoveFirst

    If AnMaO > 0 Then
        ReDim GlMaO(AnMaO, 37)
        Do
        If CBool(RS152.Fields("Passiv").Value) = False Then
            If CBool(RS152.Fields("Gesperrt").Value) = False Then
                If RS152.Fields("OnlTer").Value <> vbNullString Then
                    If CBool(RS152.Fields("OnlTer").Value) = True Then
                        GlMaO(AktZa, 0) = AktZa
                        If RS152.Fields("IDKurz").Value <> vbNullString Then GlMaO(AktZa, 1) = RS152.Fields("IDKurz").Value
                        GlMaO(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
                        If RS152.Fields("Name").Value <> vbNullString Then GlMaO(AktZa, 3) = RS152.Fields("Name").Value
                        If RS152.Fields("Vorname").Value <> vbNullString Then GlMaO(AktZa, 4) = RS152.Fields("Vorname").Value
                        GlMaO(AktZa, 5) = CBool(RS152.Fields("Passiv").Value)
                        If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
                            If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
                                GlMaO(AktZa, 6) = GlSZe 'Sprechzietenstring
                            Else
                                GlMaO(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
                            End If
                        Else
                            GlMaO(AktZa, 6) = GlSZe 'Sprechzietenstring
                        End If
                        If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
                            GlMaO(AktZa, 21) = RS152.Fields("Buchungszeiten").Value
                        Else
                            GlMaO(AktZa, 21) = GlSZe 'Sprechzietenstring
                        End If
                        If RS152.Fields("OnlRas").Value <> vbNullString Then
                            GlMaO(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
                        Else
                            GlMaO(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
                        End If
                        If RS152.Fields("KVNummer").Value <> vbNullString Then GlMaO(AktZa, 7) = RS152.Fields("KVNummer").Value
                        If RS152.Fields("Größe").Value <> vbNullString Then GlMaO(AktZa, 9) = RS152.Fields("Größe").Value
                        If RS152.Fields("Gewicht").Value <> vbNullString Then GlMaO(AktZa, 10) = RS152.Fields("Gewicht").Value
                        If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMaO(AktZa, 11) = RS152.Fields("Gesperrt").Value
                        If RS152.Fields("Em_User").Value <> vbNullString Then GlMaO(AktZa, 12) = RS152.Fields("Em_User").Value
                        If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMaO(AktZa, 13) = RS152.Fields("Em_Pass").Value
                        If RS152.Fields("Versand").Value <> vbNullString Then GlMaO(AktZa, 14) = RS152.Fields("Versand").Value
                        If RS152.Fields("OnlMax").Value <> vbNullString Then GlMaO(AktZa, 15) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
                        If RS152.Fields("OnlVor").Value <> vbNullString Then GlMaO(AktZa, 16) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
                        If RS152.Fields("GuiID").Value <> vbNullString Then GlMaO(AktZa, 17) = RS152.Fields("GuiID").Value 'UserID
                        If RS152.Fields("OnlTer").Value <> vbNullString Then GlMaO(AktZa, 18) = RS152.Fields("OnlTer").Value 'Online-Terminbuchungs System
                        If RS152.Fields("Telefon5").Value <> vbNullString Then
                            GlMaO(AktZa, 19) = RS152.Fields("Telefon5").Value
                        Else
                            GlMaO(AktZa, 19) = "keine@emailadresse.de"
                        End If
                        If RS152.Fields("Titel").Value <> vbNullString Then GlMaO(AktZa, 20) = RS152.Fields("Titel").Value
                        If RS152.Fields("OnlRa2").Value <> vbNullString Then
                            GlMaO(AktZa, 22) = Format$(RS152.Fields("OnlRa2").Value, "00")
                        Else
                            GlMaO(AktZa, 22) = Format$(GlZeR, "00") 'Zeitrasterindex
                        End If
                        If RS152.Fields("ID3").Value <> vbNullString Then
                            GlMaO(AktZa, 23) = RS152.Fields("ID3").Value
                        Else
                            GlMaO(AktZa, 23) = GlFri 'Fachrichtung
                        End If
                        If RS152.Fields("OnlRa1").Value <> vbNullString Then
                            GlMaO(AktZa, 24) = Format$(RS152.Fields("OnlRa1").Value, "00")
                        Else
                            GlMaO(AktZa, 24) = "12"
                        End If
                        If RS152.Fields("OnlTmp").Value <> vbNullString Then
                            If CInt(RS152.Fields("OnlTmp").Value) > 0 Then
                                GlMaO(AktZa, 25) = RS152.Fields("OnlTmp").Value
                            Else
                                GlMaO(AktZa, 25) = 24
                            End If
                        Else
                            GlMaO(AktZa, 25) = 24
                        End If
                        If RS152.Fields("StaKon").Value <> vbNullString Then 'Standarderlöskonto (Bankkonto)
                            If RS152.Fields("StaKon").Value > 0 Then
                                GlMaO(AktZa, 26) = RS152.Fields("StaKon").Value
                            Else
                                GlMaO(AktZa, 26) = GlSE2
                            End If
                        Else
                            GlMaO(AktZa, 26) = GlSE2
                        End If
                        If RS152.Fields("StaKo2").Value <> vbNullString Then 'Standarderlöskonto (Kasse)
                            If RS152.Fields("StaKo2").Value > 0 Then
                                GlMaO(AktZa, 27) = RS152.Fields("StaKo2").Value
                            Else
                                GlMaO(AktZa, 27) = GlSE1
                            End If
                        Else
                            GlMaO(AktZa, 27) = GlSE1
                        End If
                        If RS152.Fields("StaGk1").Value <> vbNullString Then 'Standardgeldkonto (Bankkonto)
                            If RS152.Fields("StaGk1").Value > 0 Then
                                GlMaO(AktZa, 28) = CLng(RS152.Fields("StaGk1").Value)
                            Else
                                GlMaO(AktZa, 28) = GlGkB
                            End If
                        Else
                            GlMaO(AktZa, 28) = GlGkB
                        End If
                        If RS152.Fields("StaGk2").Value <> vbNullString Then 'Standardgeldkonto (Kasse)
                            If RS152.Fields("StaGk2").Value > 0 Then
                                GlMaO(AktZa, 29) = CLng(RS152.Fields("StaGk2").Value)
                            Else
                                GlMaO(AktZa, 29) = GlGkK
                            End If
                        Else
                            GlMaO(AktZa, 29) = GlGkK
                        End If
                        If RS152.Fields("SteRet").Value <> vbNullString Then 'Standardbelegtyp
                            If IsNull(RS152.Fields("SteRet").Value) = False Then
                                GlMaO(AktZa, 30) = Left$(RS152.Fields("SteRet").Value, 1)
                            Else
                                GlMaO(AktZa, 30) = Left$(GlSet(1, 30), 1)
                            End If
                        Else
                            GlMaO(AktZa, 30) = Left$(GlSet(1, 30), 1)
                        End If
                        If RS152.Fields("OnlRa2").Value <> vbNullString Then
                            GlMaO(AktZa, 31) = Format$(RS152.Fields("OnlRa2").Value, "00")
                        Else
                            GlMaO(AktZa, 31) = Format$(GlZeR, "00") 'Zeitrasterindex
                        End If
                        If RS152.Fields("Firma2").Value <> vbNullString Then GlMaO(AktZa, 32) = RS152.Fields("Firma2").Value 'Verkehrsname
                        If RS152.Fields("Objekt").Value <> vbNullString Then GlMaO(AktZa, 33) = RS152.Fields("Objekt").Value 'Signatur
                        If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMaO(AktZa, 34) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
                        If RS152.Fields("ID3").Value <> vbNullString Then
                            GlMaO(AktZa, 35) = RS152.Fields("ID3").Value
                        Else
                            GlMaO(AktZa, 35) = GlFri 'Fachrichtung
                        End If
                        If RS152.Fields("OnlRa1").Value <> vbNullString Then
                            GlMaO(AktZa, 36) = Format$(RS152.Fields("OnlRa1").Value, "00")
                        Else
                            GlMaO(AktZa, 36) = "12"
                        End If
                        If RS152.Fields("StStKt").Value <> vbNullString Then 'Standarderlöskonto (Bankkonto)
                            If RS152.Fields("StStKt").Value > 0 Then
                                GlMaO(AktZa, 37) = SBuFo(CLng(RS152.Fields("StStKt").Value))
                            Else
                                GlMaO(AktZa, 37) = GlSKo
                            End If
                        Else
                            GlMaO(AktZa, 37) = GlSKo
                        End If
                        AktZa = AktZa + 1
                    End If
                End If
            End If
        End If
        RS152.MoveNext
        Loop Until RS152.EOF
    Else
        S_Ary2h
    End If

Else
    S_Ary2f
    S_Ary2g
    S_Ary2h
End If
RS152.Close
Set RS152 = Nothing
DoEvents

'--------------------------------------------------

If GlTyp < 2 Then
    SQL1 = "SELECT * FROM dbo.qryPatMit ORDER BY IDKurz"
Else
    SQL1 = "SELECT * FROM qryPatMit ORDER BY [IDKurz];"
End If
AktZa = 1
Set RS152 = New ADODB.Recordset 'Mitarbeiter Index
With RS152
    .CursorLocation = adUseClient
    .Source = SQL1
    .ActiveConnection = DB1
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open Options:=adCmdText
End With
GesZa = RS152.RecordCount
If GesZa > 0 Then
    GlMiV = True 'Mitarbeiter vorhanden
    ReDim GlMiK(GesZa, 40) 'Alle Mitarbeiter
    Do
    GlMiK(AktZa, 0) = AktZa
    GlMiK(AktZa, 1) = RS152.Fields("IDKurz").Value
    GlMiK(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
    If RS152.Fields("Name").Value <> vbNullString Then GlMiK(AktZa, 3) = RS152.Fields("Name").Value
    If RS152.Fields("Vorname").Value <> vbNullString Then GlMiK(AktZa, 4) = RS152.Fields("Vorname").Value
    GlMiK(AktZa, 5) = RS152.Fields("Passiv").Value
    If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
        If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
            GlMiK(AktZa, 6) = GlSZe 'Sprechzietenstring
        Else
            GlMiK(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
        End If
    Else
        GlMiK(AktZa, 6) = GlSZe 'Sprechzietenstring
    End If
    If RS152.Fields("IDP").Value <> vbNullString Then
        GlMiK(AktZa, 7) = RS152.Fields("IDP").Value 'zugeordneter Mandant
    Else
        GlMiK(AktZa, 7) = 0
    End If
    If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
        GlMiK(AktZa, 24) = RS152.Fields("Buchungszeiten").Value
    Else
        GlMiK(AktZa, 24) = GlSZe 'Sprechzietenstring
    End If
    If RS152.Fields("OnlRas").Value <> vbNullString Then
        If IsNumeric(RS152.Fields("OnlRas").Value) = True Then
            GlMiK(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
        Else
            GlMiK(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
        End If
    Else
        GlMiK(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
    End If
    If RS152.Fields("Größe").Value <> vbNullString Then GlMiK(AktZa, 9) = RS152.Fields("Größe").Value
    If RS152.Fields("Gewicht").Value <> vbNullString Then GlMiK(AktZa, 10) = RS152.Fields("Gewicht").Value
    If RS152.Fields("Objekt").Value <> vbNullString Then GlMiK(AktZa, 11) = RS152.Fields("Objekt").Value 'Signatur
    If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMiK(AktZa, 12) = RS152.Fields("Gesperrt").Value
    If RS152.Fields("Em_User").Value <> vbNullString Then GlMiK(AktZa, 13) = RS152.Fields("Em_User").Value
    If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMiK(AktZa, 14) = RS152.Fields("Em_Pass").Value
    If RS152.Fields("Versand").Value <> vbNullString Then GlMiK(AktZa, 15) = RS152.Fields("Versand").Value
    If RS152.Fields("OnlMax").Value <> vbNullString Then GlMiK(AktZa, 16) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
    If RS152.Fields("OnlVor").Value <> vbNullString Then GlMiK(AktZa, 17) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
    If RS152.Fields("Blutgruppe").Value <> vbNullString Then GlMiK(AktZa, 18) = RS152.Fields("Blutgruppe").Value 'Passwort
    If RS152.Fields("Kontoinhaber").Value <> vbNullString Then GlMiK(AktZa, 19) = RS152.Fields("Kontoinhaber").Value 'Rechte
    If RS152.Fields("GuiID").Value <> vbNullString Then GlMiK(AktZa, 20) = RS152.Fields("GuiID").Value 'UserID
    If RS152.Fields("OnlTer").Value <> vbNullString Then GlMiK(AktZa, 21) = RS152.Fields("OnlTer").Value    'Online-Terminbuchungs System
    If RS152.Fields("Telefon5").Value <> vbNullString Then
        GlMiK(AktZa, 22) = RS152.Fields("Telefon5").Value
    Else
        GlMiK(AktZa, 22) = "keine@emailadresse.de"
    End If
    If RS152.Fields("Titel").Value <> vbNullString Then GlMiK(AktZa, 23) = RS152.Fields("Titel").Value
    If RS152.Fields("Telefon6").Value <> vbNullString Then GlMiK(AktZa, 25) = RS152.Fields("Telefon6").Value 'Signaturdatei
    If RS152.Fields("OnlRa2").Value <> vbNullString Then
        If IsNumeric(RS152.Fields("OnlRa2").Value) = True Then
            GlMiK(AktZa, 26) = Format$(RS152.Fields("OnlRa2").Value, "00")
        Else
            GlMiK(AktZa, 26) = Format$(GlZeR, "00") 'Zeitrasterindex
        End If
    Else
        GlMiK(AktZa, 26) = Format$(GlZeR, "00") 'Zeitrasterindex
    End If
    If RS152.Fields("Firma2").Value <> vbNullString Then GlMiK(AktZa, 27) = RS152.Fields("Firma2").Value 'Anzeigename
    If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMiK(AktZa, 28) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
    If RS152.Fields("Beruf").Value <> vbNullString Then GlMiK(AktZa, 29) = RS152.Fields("Beruf").Value
    If RS152.Fields("Telefon1").Value <> vbNullString Then GlMiK(AktZa, 30) = RS152.Fields("Telefon1").Value
    If RS152.Fields("GLN").Value <> vbNullString Then
        GlMiK(AktZa, 31) = RS152.Fields("GLN").Value 'GLN
    Else
        GlMiK(AktZa, 31) = "000000"
    End If
    If RS152.Fields("ZSR").Value <> vbNullString Then
        GlMiK(AktZa, 32) = RS152.Fields("ZSR").Value 'ZSR
    Else
        GlMiK(AktZa, 32) = "000000"
    End If
    If RS152.Fields("R_Firma1").Value <> vbNullString Then
        If InStr(1, RS152.Fields("Name").Value, RS152.Fields("R_Firma1").Value, 1) > 0 Then
            If RS152.Fields("Titel").Value <> vbNullString Then
                MaAbs = RS152.Fields("Titel").Value & Chr$(32) & RS152.Fields("Vorname").Value & Chr$(32) & RS152.Fields("Name").Value
            Else
                MaAbs = RS152.Fields("Vorname").Value & Chr$(32) & RS152.Fields("Name").Value
            End If
        Else
            If RS152.Fields("Vorname").Value <> vbNullString Then
                MaAbs = RS152.Fields("R_Firma1").Value & " - " & Left$(RS152.Fields("Vorname").Value, 1) & ". " & RS152.Fields("Name").Value
            Else
                MaAbs = RS152.Fields("R_Firma1").Value & " - " & RS152.Fields("Name").Value
            End If
        End If
    Else
        If RS152.Fields("Titel").Value <> vbNullString Then
            MaAbs = RS152.Fields("R_Firma1").Value & Chr$(32) & RS152.Fields("Vorname").Value & Chr$(32) & RS152.Fields("Name").Value
        Else
            MaAbs = RS152.Fields("Vorname").Value & Chr$(32) & RS152.Fields("Name").Value
        End If
    End If
    EmBet = MaAbs
    MaAbs = MaAbs & " - " & RS152.Fields("Straße").Value
    MaAbs = MaAbs & " - " & RS152.Fields("PLZ").Value
    MaAbs = MaAbs & Chr$(32) & RS152.Fields("Ort").Value
    GlMiK(AktZa, 33) = MaAbs 'Briefabsenezeile
    If RS152.Fields("Straße").Value <> vbNullString Then GlMiK(AktZa, 34) = RS152.Fields("Straße").Value
    If RS152.Fields("PLZ").Value <> vbNullString Then GlMiK(AktZa, 35) = RS152.Fields("PLZ").Value
    If RS152.Fields("Ort").Value <> vbNullString Then GlMiK(AktZa, 36) = RS152.Fields("Ort").Value
    If RS152.Fields("OnlRa1").Value <> vbNullString Then
        If IsNumeric(RS152.Fields("OnlRa1").Value) = True Then
            GlMiK(AktZa, 37) = Format$(RS152.Fields("OnlRa1").Value, "00")
        Else
            GlMiK(AktZa, 37) = "12"
        End If
    Else
        GlMiK(AktZa, 37) = "24"
    End If
    If RS152.Fields("Kurativ").Value <> vbNullString Then
        GlMiK(AktZa, 38) = RS152.Fields("Kurativ").Value
    Else
        GlMiK(AktZa, 38) = 0
    End If
    If RS152.Fields("OnlTmp").Value <> vbNullString Then
        If CInt(RS152.Fields("OnlTmp").Value) > 0 Then
            GlMiK(AktZa, 39) = RS152.Fields("OnlTmp").Value
        Else
            GlMiK(AktZa, 39) = 24
        End If
    Else
        GlMiK(AktZa, 39) = 24
    End If
    If CBool(RS152.Fields("Passiv").Value) = False Then
        AnMiA = AnMiA + 1 'Anzahl Aktiver Mitarbeiter
        If CBool(RS152.Fields("Gesperrt").Value) = False Then
            AnMiT = AnMiT + 1 'Anzahl Aktiver Mitarbeiter + Terminspalte
            If RS152.Fields("OnlTer").Value <> vbNullString Then
                If CBool(RS152.Fields("OnlTer").Value) = True Then
                    AnMiO = AnMiO + 1 'Anzahl Aktiver Mitarbeiter + Terminspalte + OTS
                End If
            End If
        End If
    End If
    AktZa = AktZa + 1
    RS152.MoveNext
    Loop Until RS152.EOF
    
    AktZa = 1
    RS152.MoveFirst

    If AnMiA > 0 Then
        ReDim GlMiA(AnMiA, 40) 'Aktive Mitarbeier
        Do
        If CBool(RS152.Fields("Passiv").Value) = False Then
            GlMiA(AktZa, 0) = AktZa
            GlMiA(AktZa, 1) = RS152.Fields("IDKurz").Value
            GlMiA(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
            If RS152.Fields("Name").Value <> vbNullString Then GlMiA(AktZa, 3) = RS152.Fields("Name").Value
            If RS152.Fields("Vorname").Value <> vbNullString Then GlMiA(AktZa, 4) = RS152.Fields("Vorname").Value
            GlMiA(AktZa, 5) = RS152.Fields("Passiv").Value
            If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
                If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
                    GlMiA(AktZa, 6) = GlSZe 'Sprechzietenstring
                Else
                    GlMiA(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
                End If
            Else
                GlMiA(AktZa, 6) = GlSZe 'Sprechzietenstring
            End If
            If RS152.Fields("IDP").Value <> vbNullString Then
                GlMiA(AktZa, 7) = RS152.Fields("IDP").Value 'zugeordneter Mandant
            Else
                GlMiA(AktZa, 7) = 0
            End If
            If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
                GlMiA(AktZa, 24) = RS152.Fields("Buchungszeiten").Value
            Else
                GlMiA(AktZa, 24) = GlSZe 'Sprechzietenstring
            End If
            If RS152.Fields("OnlRas").Value <> vbNullString Then
                If IsNumeric(RS152.Fields("OnlRas").Value) = True Then
                    GlMiA(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
                Else
                    GlMiA(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
                End If
            Else
                GlMiA(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
            End If
            If RS152.Fields("Größe").Value <> vbNullString Then GlMiA(AktZa, 9) = RS152.Fields("Größe").Value
            If RS152.Fields("Gewicht").Value <> vbNullString Then GlMiA(AktZa, 10) = RS152.Fields("Gewicht").Value
            If RS152.Fields("Objekt").Value <> vbNullString Then GlMiA(AktZa, 11) = RS152.Fields("Objekt").Value 'Emaisignatur
            If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMiA(AktZa, 12) = RS152.Fields("Gesperrt").Value
            If RS152.Fields("Em_User").Value <> vbNullString Then GlMiA(AktZa, 13) = RS152.Fields("Em_User").Value
            If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMiA(AktZa, 14) = RS152.Fields("Em_Pass").Value
            If RS152.Fields("Versand").Value <> vbNullString Then GlMiA(AktZa, 15) = RS152.Fields("Versand").Value 'Multiterminbetreffauswahl
            If RS152.Fields("OnlMax").Value <> vbNullString Then GlMiA(AktZa, 16) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
            If RS152.Fields("OnlVor").Value <> vbNullString Then GlMiA(AktZa, 17) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
            If RS152.Fields("Blutgruppe").Value <> vbNullString Then GlMiA(AktZa, 18) = RS152.Fields("Blutgruppe").Value 'Passwort
            If RS152.Fields("Kontoinhaber").Value <> vbNullString Then GlMiA(AktZa, 19) = RS152.Fields("Kontoinhaber").Value 'Rechte
            If RS152.Fields("GuiID").Value <> vbNullString Then GlMiA(AktZa, 20) = RS152.Fields("GuiID").Value 'UserID
            If RS152.Fields("OnlTer").Value <> vbNullString Then GlMiA(AktZa, 21) = RS152.Fields("OnlTer").Value  'Online-Terminbuchungs System
            If RS152.Fields("Telefon5").Value <> vbNullString Then
                GlMiA(AktZa, 22) = RS152.Fields("Telefon5").Value
            Else
                GlMiA(AktZa, 22) = "keine@emailadresse.de"
            End If
            If RS152.Fields("Titel").Value <> vbNullString Then GlMiA(AktZa, 23) = RS152.Fields("Titel").Value
            If RS152.Fields("Telefon6").Value <> vbNullString Then GlMiA(AktZa, 25) = RS152.Fields("Telefon6").Value
            If RS152.Fields("OnlRa2").Value <> vbNullString Then
                If IsNumeric(RS152.Fields("OnlRa2").Value) = True Then
                    GlMiA(AktZa, 26) = Format$(RS152.Fields("OnlRa2").Value, "00")
                Else
                    GlMiA(AktZa, 26) = Format$(GlZeR, "00") 'Onlinezeitrasterindex
                End If
            Else
                GlMiA(AktZa, 26) = Format$(GlZeR, "00") 'Onlinezeitrasterindex
            End If
            If RS152.Fields("Firma2").Value <> vbNullString Then GlMiA(AktZa, 27) = RS152.Fields("Firma2").Value
            If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMiA(AktZa, 28) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
            If RS152.Fields("Beruf").Value <> vbNullString Then GlMiA(AktZa, 29) = RS152.Fields("Beruf").Value
            If RS152.Fields("Telefon1").Value <> vbNullString Then GlMiA(AktZa, 30) = RS152.Fields("Telefon1").Value
            If RS152.Fields("GLN").Value <> vbNullString Then
                GlMiA(AktZa, 31) = RS152.Fields("GLN").Value 'GLN
            Else
                GlMiA(AktZa, 31) = "000000"
            End If
            If RS152.Fields("ZSR").Value <> vbNullString Then
                GlMiA(AktZa, 32) = RS152.Fields("ZSR").Value 'ZSR
            Else
                GlMiA(AktZa, 32) = "000000"
            End If
            GlMiA(AktZa, 33) = vbNullString 'Briefabsenezeile
            If RS152.Fields("Straße").Value <> vbNullString Then GlMiA(AktZa, 34) = RS152.Fields("Straße").Value
            If RS152.Fields("PLZ").Value <> vbNullString Then GlMiA(AktZa, 35) = RS152.Fields("PLZ").Value
            If RS152.Fields("Ort").Value <> vbNullString Then GlMiA(AktZa, 36) = RS152.Fields("Ort").Value
            If RS152.Fields("OnlRa1").Value <> vbNullString Then
                If IsNumeric(RS152.Fields("OnlRa1").Value) = True Then
                    GlMiA(AktZa, 37) = Format$(RS152.Fields("OnlRa1").Value, "00")
                Else
                    GlMiA(AktZa, 37) = "12"
                End If
            Else
                GlMiA(AktZa, 37) = "24"
            End If
            If RS152.Fields("Kurativ").Value <> vbNullString Then
                GlMiA(AktZa, 38) = RS152.Fields("Kurativ").Value
            Else
                GlMiA(AktZa, 38) = 0
            End If
            If RS152.Fields("OnlTmp").Value <> vbNullString Then
                If CInt(RS152.Fields("OnlTmp").Value) > 0 Then
                    GlMiA(AktZa, 39) = RS152.Fields("OnlTmp").Value
                Else
                    GlMiA(AktZa, 39) = 24
                End If
            Else
                GlMiA(AktZa, 39) = 24
            End If
            AktZa = AktZa + 1
        End If
        RS152.MoveNext
        Loop Until RS152.EOF
    Else
        S_Ary2b
    End If
    
    AktZa = 1
    RS152.MoveFirst
    
    If AnMiT > 0 Then
        ReDim GlMiT(AnMiT, 40) 'aktive Mitarbeier + Terminspalte
        Do
        If CBool(RS152.Fields("Passiv").Value) = False Then
            If CBool(RS152.Fields("Gesperrt").Value) = False Then
                GlMiT(AktZa, 0) = AktZa
                GlMiT(AktZa, 1) = RS152.Fields("IDKurz").Value
                GlMiT(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
                If RS152.Fields("Name").Value <> vbNullString Then GlMiT(AktZa, 3) = RS152.Fields("Name").Value
                If RS152.Fields("Vorname").Value <> vbNullString Then GlMiT(AktZa, 4) = RS152.Fields("Vorname").Value
                GlMiT(AktZa, 5) = RS152.Fields("Passiv").Value
                If RS152.Fields("IDP").Value <> vbNullString Then
                    GlMiT(AktZa, 7) = RS152.Fields("IDP").Value 'zugeordneter Mandant
                Else
                    GlMiT(AktZa, 7) = 0
                End If
                If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
                    If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
                        GlMiT(AktZa, 6) = GlSZe 'Sprechzietenstring
                    Else
                        GlMiT(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
                    End If
                Else
                    GlMiT(AktZa, 6) = GlSZe 'Sprechzietenstring
                End If
                If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
                    GlMiT(AktZa, 24) = RS152.Fields("Buchungszeiten").Value
                Else
                    GlMiT(AktZa, 24) = GlSZe 'Sprechzietenstring
                End If
                If RS152.Fields("OnlRas").Value <> vbNullString Then
                    If IsNumeric(RS152.Fields("OnlRas").Value) = True Then
                        GlMiT(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
                    Else
                        GlMiT(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
                    End If
                Else
                    GlMiT(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
                End If
                If RS152.Fields("Größe").Value <> vbNullString Then GlMiT(AktZa, 9) = RS152.Fields("Größe").Value 'Username
                If RS152.Fields("Gewicht").Value <> vbNullString Then GlMiT(AktZa, 10) = RS152.Fields("Gewicht").Value 'Password
                If RS152.Fields("Objekt").Value <> vbNullString Then GlMiT(AktZa, 11) = RS152.Fields("Objekt").Value 'Emailsignatur
                If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMiT(AktZa, 12) = RS152.Fields("Gesperrt").Value
                If RS152.Fields("Em_User").Value <> vbNullString Then GlMiT(AktZa, 13) = RS152.Fields("Em_User").Value
                If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMiT(AktZa, 14) = RS152.Fields("Em_Pass").Value
                If RS152.Fields("Versand").Value <> vbNullString Then GlMiT(AktZa, 15) = RS152.Fields("Versand").Value 'Multiterminbetreffauswahl
                If RS152.Fields("OnlMax").Value <> vbNullString Then GlMiT(AktZa, 16) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
                If RS152.Fields("OnlVor").Value <> vbNullString Then GlMiT(AktZa, 17) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
                If RS152.Fields("Blutgruppe").Value <> vbNullString Then GlMiT(AktZa, 18) = RS152.Fields("Blutgruppe").Value 'Passwort
                If RS152.Fields("Kontoinhaber").Value <> vbNullString Then GlMiT(AktZa, 19) = RS152.Fields("Kontoinhaber").Value 'Rechte
                If RS152.Fields("GuiID").Value <> vbNullString Then GlMiT(AktZa, 20) = RS152.Fields("GuiID").Value 'UserID
                If RS152.Fields("OnlTer").Value <> vbNullString Then GlMiT(AktZa, 21) = RS152.Fields("OnlTer").Value  'Online-Terminbuchungs System
                If RS152.Fields("Telefon5").Value <> vbNullString Then
                    GlMiT(AktZa, 22) = RS152.Fields("Telefon5").Value
                Else
                    GlMiT(AktZa, 22) = "keine@emailadresse.de"
                End If
                If RS152.Fields("Titel").Value <> vbNullString Then GlMiT(AktZa, 23) = RS152.Fields("Titel").Value
                If RS152.Fields("Telefon6").Value <> vbNullString Then GlMiT(AktZa, 25) = RS152.Fields("Telefon6").Value
                If RS152.Fields("OnlRa2").Value <> vbNullString Then
                    If IsNumeric(RS152.Fields("OnlRa2").Value) = True Then
                        GlMiT(AktZa, 26) = Format$(RS152.Fields("OnlRa2").Value, "00")
                    Else
                        GlMiT(AktZa, 26) = Format$(GlZeR, "00") 'Onlinezeitrasterindex
                    End If
                Else
                    GlMiT(AktZa, 26) = Format$(GlZeR, "00") 'Onlinezeitrasterindex
                End If
                If RS152.Fields("Firma2").Value <> vbNullString Then GlMiT(AktZa, 27) = RS152.Fields("Firma2").Value 'Verkehrsname
                If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMiT(AktZa, 28) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
                If RS152.Fields("Beruf").Value <> vbNullString Then GlMiT(AktZa, 29) = RS152.Fields("Beruf").Value
                If RS152.Fields("Telefon1").Value <> vbNullString Then GlMiT(AktZa, 30) = RS152.Fields("Telefon1").Value
                If RS152.Fields("GLN").Value <> vbNullString Then
                    GlMiT(AktZa, 31) = RS152.Fields("GLN").Value 'GLN
                Else
                    GlMiT(AktZa, 31) = "000000"
                End If
                If RS152.Fields("ZSR").Value <> vbNullString Then
                    GlMiT(AktZa, 32) = RS152.Fields("ZSR").Value 'ZSR
                Else
                    GlMiT(AktZa, 32) = "000000"
                End If
                GlMiT(AktZa, 33) = vbNullString 'Briefabsenezeile
                If RS152.Fields("Straße").Value <> vbNullString Then GlMiT(AktZa, 34) = RS152.Fields("Straße").Value
                If RS152.Fields("PLZ").Value <> vbNullString Then GlMiT(AktZa, 35) = RS152.Fields("PLZ").Value
                If RS152.Fields("Ort").Value <> vbNullString Then GlMiT(AktZa, 36) = RS152.Fields("Ort").Value
                If RS152.Fields("OnlRa1").Value <> vbNullString Then
                    If IsNumeric(RS152.Fields("OnlRa1").Value) = True Then
                        GlMiT(AktZa, 37) = Format$(RS152.Fields("OnlRa1").Value, "00")
                    Else
                        GlMiT(AktZa, 37) = "12"
                    End If
                Else
                    GlMiT(AktZa, 37) = "24"
                End If
                If RS152.Fields("Kurativ").Value <> vbNullString Then
                    GlMiT(AktZa, 38) = RS152.Fields("Kurativ").Value
                Else
                    GlMiT(AktZa, 38) = 0
                End If
                If RS152.Fields("OnlTmp").Value <> vbNullString Then
                    If CInt(RS152.Fields("OnlTmp").Value) > 0 Then
                        GlMiT(AktZa, 39) = RS152.Fields("OnlTmp").Value
                    Else
                        GlMiT(AktZa, 39) = 24
                    End If
                Else
                    GlMiT(AktZa, 39) = 24
                End If
                AktZa = AktZa + 1
            End If
        End If
        RS152.MoveNext
        Loop Until RS152.EOF
    Else
        S_Ary2c
    End If
    
    AktZa = 1
    RS152.MoveFirst
'üüü
    If AnMiO > 0 Then
        ReDim GlMiO(AnMiO, 40) 'aktive Mitarbeier + Terminspalte + OTS
        Do
        If CBool(RS152.Fields("Passiv").Value) = False Then
            If CBool(RS152.Fields("Gesperrt").Value) = False Then
                If RS152.Fields("OnlTer").Value <> vbNullString Then
                    If CBool(RS152.Fields("OnlTer").Value) = True Then
                        GlMiO(AktZa, 0) = AktZa
                        GlMiO(AktZa, 1) = RS152.Fields("IDKurz").Value
                        GlMiO(AktZa, 2) = CLng(RS152.Fields("ID0").Value)
                        If RS152.Fields("Name").Value <> vbNullString Then GlMiO(AktZa, 3) = RS152.Fields("Name").Value
                        If RS152.Fields("Vorname").Value <> vbNullString Then GlMiO(AktZa, 4) = RS152.Fields("Vorname").Value
                        GlMiO(AktZa, 5) = RS152.Fields("Passiv").Value
                        If RS152.Fields("IDP").Value <> vbNullString Then
                            GlMiO(AktZa, 7) = RS152.Fields("IDP").Value 'zugeordneter Mandant
                        Else
                            GlMiO(AktZa, 7) = 0
                        End If
                        If RS152.Fields("Sprechzeiten").Value <> vbNullString Then
                            If Len(RS152.Fields("Sprechzeiten").Value) < 100 Then
                                GlMiO(AktZa, 6) = GlSZe 'Sprechzietenstring
                            Else
                                GlMiO(AktZa, 6) = RS152.Fields("Sprechzeiten").Value
                            End If
                        Else
                            GlMiO(AktZa, 6) = GlSZe 'Sprechzietenstring
                        End If
                        If RS152.Fields("Buchungszeiten").Value <> vbNullString Then
                            GlMiO(AktZa, 24) = RS152.Fields("Buchungszeiten").Value
                        Else
                            GlMiO(AktZa, 24) = GlSZe 'Sprechzietenstring
                        End If
                        If RS152.Fields("OnlRas").Value <> vbNullString Then
                            If IsNumeric(RS152.Fields("OnlRas").Value) = True Then
                                GlMiO(AktZa, 8) = Format$(RS152.Fields("OnlRas").Value, "00")
                            Else
                                GlMiO(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
                            End If
                        Else
                            GlMiO(AktZa, 8) = Format$(GlZeR, "00") 'Zeitrasterindex
                        End If
                        If RS152.Fields("Größe").Value <> vbNullString Then GlMiO(AktZa, 9) = RS152.Fields("Größe").Value 'Username
                        If RS152.Fields("Gewicht").Value <> vbNullString Then GlMiO(AktZa, 10) = RS152.Fields("Gewicht").Value 'Password
                        If RS152.Fields("Objekt").Value <> vbNullString Then GlMiO(AktZa, 11) = RS152.Fields("Objekt").Value 'Emailsignatur
                        If RS152.Fields("Gesperrt").Value <> vbNullString Then GlMiO(AktZa, 12) = RS152.Fields("Gesperrt").Value
                        If RS152.Fields("Em_User").Value <> vbNullString Then GlMiO(AktZa, 13) = RS152.Fields("Em_User").Value
                        If RS152.Fields("Em_Pass").Value <> vbNullString Then GlMiO(AktZa, 14) = RS152.Fields("Em_Pass").Value
                        If RS152.Fields("Versand").Value <> vbNullString Then GlMiO(AktZa, 15) = RS152.Fields("Versand").Value 'Multiterminbetreffauswahl
                        If RS152.Fields("OnlMax").Value <> vbNullString Then GlMiO(AktZa, 16) = Format$(RS152.Fields("OnlMax").Value, "00") 'Anzahl maximal buchbarer Termine pro Tag
                        If RS152.Fields("OnlVor").Value <> vbNullString Then GlMiO(AktZa, 17) = Format$(RS152.Fields("OnlVor").Value, "00") 'Vorlaufzeit für Terminbuchung
                        If RS152.Fields("Blutgruppe").Value <> vbNullString Then GlMiO(AktZa, 18) = RS152.Fields("Blutgruppe").Value 'Passwort
                        If RS152.Fields("Kontoinhaber").Value <> vbNullString Then GlMiO(AktZa, 19) = RS152.Fields("Kontoinhaber").Value 'Rechte
                        If RS152.Fields("GuiID").Value <> vbNullString Then GlMiO(AktZa, 20) = RS152.Fields("GuiID").Value 'UserID
                        If RS152.Fields("OnlTer").Value <> vbNullString Then GlMiO(AktZa, 21) = RS152.Fields("OnlTer").Value  'Online-Terminbuchungs System
                        If RS152.Fields("Telefon5").Value <> vbNullString Then
                            GlMiO(AktZa, 22) = RS152.Fields("Telefon5").Value
                        Else
                            GlMiO(AktZa, 22) = "keine@emailadresse.de"
                        End If
                        If RS152.Fields("Titel").Value <> vbNullString Then GlMiO(AktZa, 23) = RS152.Fields("Titel").Value
                        If RS152.Fields("Telefon6").Value <> vbNullString Then GlMiO(AktZa, 25) = RS152.Fields("Telefon6").Value
                        If RS152.Fields("OnlRa2").Value <> vbNullString Then
                            If IsNumeric(RS152.Fields("OnlRa2").Value) = True Then
                                GlMiO(AktZa, 26) = Format$(RS152.Fields("OnlRa2").Value, "00")
                            Else
                                GlMiO(AktZa, 26) = Format$(GlZeR, "00") 'Onlinezeitrasterindex
                            End If
                        Else
                            GlMiO(AktZa, 26) = Format$(GlZeR, "00") 'Onlinezeitrasterindex
                        End If
                        If RS152.Fields("Firma2").Value <> vbNullString Then GlMiO(AktZa, 27) = RS152.Fields("Firma2").Value
                        If RS152.Fields("OnlMa2").Value <> vbNullString Then GlMiO(AktZa, 28) = Format$(RS152.Fields("OnlMa2").Value, "00") 'Anzahl maximal buchbarer Termine pro Patient / Tag
                        If RS152.Fields("Beruf").Value <> vbNullString Then GlMiO(AktZa, 29) = RS152.Fields("Beruf").Value
                        If RS152.Fields("Telefon1").Value <> vbNullString Then GlMiO(AktZa, 30) = RS152.Fields("Telefon1").Value
                        If RS152.Fields("GLN").Value <> vbNullString Then
                            GlMiO(AktZa, 31) = RS152.Fields("GLN").Value 'GLN
                        Else
                            GlMiO(AktZa, 31) = "000000"
                        End If
                        If RS152.Fields("ZSR").Value <> vbNullString Then
                            GlMiO(AktZa, 32) = RS152.Fields("ZSR").Value 'ZSR
                        Else
                            GlMiO(AktZa, 32) = "000000"
                        End If
                        GlMiO(AktZa, 33) = vbNullString 'Briefabsenezeile
                        If RS152.Fields("Straße").Value <> vbNullString Then GlMiO(AktZa, 34) = RS152.Fields("Straße").Value
                        If RS152.Fields("PLZ").Value <> vbNullString Then GlMiO(AktZa, 35) = RS152.Fields("PLZ").Value
                        If RS152.Fields("Ort").Value <> vbNullString Then GlMiO(AktZa, 36) = RS152.Fields("Ort").Value
                        If RS152.Fields("OnlRa1").Value <> vbNullString Then
                            If IsNumeric(RS152.Fields("OnlRa1").Value) = True Then
                                GlMiO(AktZa, 37) = Format$(RS152.Fields("OnlRa1").Value, "00")
                            Else
                                GlMiO(AktZa, 37) = "12"
                            End If
                        Else
                            GlMiO(AktZa, 37) = "24"
                        End If
                        If RS152.Fields("Kurativ").Value <> vbNullString Then
                            GlMiO(AktZa, 38) = RS152.Fields("Kurativ").Value
                        Else
                            GlMiO(AktZa, 38) = 0
                        End If
                        If RS152.Fields("OnlTmp").Value <> vbNullString Then
                            If CInt(RS152.Fields("OnlTmp").Value) > 0 Then
                                GlMiO(AktZa, 39) = RS152.Fields("OnlTmp").Value
                            Else
                                GlMiO(AktZa, 39) = 24
                            End If
                        Else
                            GlMiO(AktZa, 39) = 24
                        End If
                        AktZa = AktZa + 1
                    End If
                End If
            End If
        End If
        RS152.MoveNext
        Loop Until RS152.EOF
    Else
        S_Ary2d
    End If

Else
    GlMiV = False 'Mitarbeiter vorhanden
    S_Ary2a
    S_Ary2b
    S_Ary2c
    S_Ary2d
End If
RS152.Close
Set RS152 = Nothing
DoEvents

If GlSmI > UBound(GlMiA) Then
    GlSmI = 1 'Aktive Mitarbeiter (WICHTIG!)
End If

If UBound(GlMiO) > 0 Then 'Standardmitarbeiter Online-Terminbuchungs Sytem
    GlSMo = 1
End If

'--------------------------------------------------

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_Ary3 " & Err.Number
Resume Next

End Sub
Public Sub S_BaCm(Optional ByVal MaKon As Boolean = False)
On Error GoTo FiErr
'füllt Combobox in Suchleiste mit Geldkonten

Dim ManNr As Long
Dim StaRa As Long
Dim AktZa As Integer
Dim AktKo As Integer
Dim AktSa As Integer
Dim GesZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmGeg As XtremeCommandBars.CommandBarComboBox
Dim CmMan As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01

Set CmMan = CmBrs.FindControl(CmMan, SY_SuMan, , True)
Set CmGeg = CmBrs.FindControl(CmGeg, SY_SuBuh, , True)

ManNr = CmMan.ItemData(CmMan.ListIndex)

If ManNr > 0 Then
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            StaRa = GlMan(AktZa, 25) 'Standardkontenrahmen
            Exit For
        End If
    Next AktZa
Else
    StaRa = GlMan(GlSMa, 25) 'Standardkontenrahmen
End If

CmGeg.Clear
DoEvents

If MaKon = True Then
    Set RS159 = New ADODB.Recordset 'Sachkonten mit Geldkontenzuordnung
    RS159.CursorLocation = adUseClient
    Set RS159 = DBCmRe2("qrySimBuKtE", "@IdGel", "@IdxNr", -1, StaRa)
    GesZa = RS159.RecordCount
    If GesZa > 0 Then
        AktZa = 1
        Do Until RS159.EOF
        With CmGeg
            .AddItem SBuFo(RS159.Fields("IDK").Value) & Chr$(32) & RS159.Fields("IDKurz").Value
            .ItemData(AktZa) = RS159.Fields("IDB").Value
        End With
        AktZa = AktZa + 1
        RS159.MoveNext
        Loop
        With CmGeg
            .AddItem "Alle Geldkonten"
            .ItemData(AktZa) = 0
        End With
    End If
    RS159.Close
    Set RS159 = Nothing
    CmGeg.ListIndex = AktZa
Else
    With CmGeg
        If GlBuc = True Then 'einfache Buchhaltung verwenden
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(AktZa) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
            .AddItem "Alle Geldkonten"
            .ItemData(AktZa) = 0
            If .ListCount > 0 Then
                .ListIndex = AktZa
            Else
                .ListIndex = 0
            End If
        Else
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                    If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                        AktSa = AktSa + 1
                        .AddItem GlSaK(AktKo, 3)
                        .ItemData(AktSa) = GlSaK(AktKo, 6) '[IDB]
                        Exit For
                    End If
                Next AktKo
            Next AktZa
            If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
                For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                    .AddItem GlGeK(AktZa, 3)
                    .ItemData(AktZa) = GlGeK(AktZa, 0)
                Next AktZa
                .AddItem "Alle Geldkonten"
                .ItemData(AktZa) = 0
                .ListIndex = AktZa
            Else
                AktSa = AktSa + 1
                .AddItem "Alle Geldkonten"
                .ItemData(AktSa) = 0
                .ListIndex = AktSa
            End If
        End If
    End With
End If

Set CmMan = Nothing
Set CmGeg = Nothing
Set CmBrs = Nothing

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_BaCm " & Err.Number
Resume Next

End Sub
Public Sub S_Expor(ByVal ExTyp As String, Optional ByVal EmVer As Integer, Optional ByVal ManNr As Long, Optional ByVal ReAbs As Boolean = False)
On Error GoTo SuErr
'Exportiert die markierten Buchungen

Dim Datu1 As Date
Dim Datu2 As Date
Dim Zeit1 As Date
Dim Zeit2 As Date
Dim TerNr As Long
Dim IdxNr As Long
Dim MitNr As Long
Dim BlgNa As String
Dim BlgNe As String
Dim DaNam As String
Dim DaNaO As String
Dim DaPfa As String
Dim DaExt As String
Dim ZipNa As String
Dim ZipBl As String
Dim TmpSt As String
Dim CapSt As String
Dim FiNam As String
Dim FilNa As String
Dim FiExt As String
Dim ManNa As String
Dim TmStr As String
Dim TeBet As String
Dim TePat As String
Dim DaSt1 As String
Dim DaSt2 As String
Dim DaSt3 As String
Dim DaSt4 As String
Dim ExSep As String
Dim CSVep As String
Dim DaKop As String
Dim BerSt As String
Dim MaNam As String
Dim MiNam As String
Dim ExOrd As String
Dim RetWe As Boolean
Dim GesPo As Integer
Dim AktPo As Integer
Dim AnzPo As Integer
Dim AktZa As Integer
Dim BlgPo As Integer
Dim AktRe As Integer
Dim RmuNr As Integer
Dim AnzKo As Integer
Dim AkZa1 As Integer
Dim AkZa2 As Integer
Dim Frage As Integer
Dim OpnWo As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Dim NaSpa As Object
Dim MapFo As Object
Dim TeIts As Object
Dim TeItm As Object

Set clICS = New clsICS
Set clFil = New clsFile
Set clAnw = New clsAnwend

ExSep = Chr$(59) 'DATEV Seperator
CSVep = Chr$(34) & Chr$(59) & Chr$(34)

BerSt = Left$(Format$(GlDvB, "00000"), 5)
MaNam = Left$(Format$(GlDvM, "00000"), 5)

If GldKt = True Then
    AnzKo = 4
Else
    AnzKo = 6
End If

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

DaKop = "Umsatz;Soll/Haben;WKZ_Umsatz;Kurs;Basisumsatz;WKZ_Basisumsatz;Konto;Gegenkonto;BU-Schlüssel;Belegdatum;Belegfeld1;Belegfeld2;Skonto;Buchungstext;" & _
"Postensperre;Adressnummer;Geschäftspartnerbank;Sachverhalt;Zinssperre;Beleglink;Beleginfo1a;Beleginfo1b;Beleginfo2a;Beleginfo2b;Beleginfo3a;Beleginfo3b;" & _
"Beleginfo4a;Beleginfo4b;Beleginfo5a;Beleginfo5b;Beleginfo6a;Beleginfo6b;Beleginfo7a;Beleginfo7b;Beleginfo8a;Beleginfo8b;KOST1;KOST2;Menge;UStID;" & _
"Steuersatz;Versteuerungsart;L+L1;L+L2;BU49a;BU49b;BU49c;Info1a;Infor1b;Info2a;Info2b;Zusatzin3a;Info3b;Info4a;Info4b;Info5a;Info5b;Info6a;Info6b;" & _
"Info7a;Info7b;Info8a;Info8b;Info9a;Info9b;Info10a;Info10b;Info11a;Info11b;Info12a;Info12b;Info13a;Info13b;Info14a;Info14b;Info15a;Info15b;Info16a;" & _
"Info16b;Info17a;Info17b;Info18a;Info18b;Info19a;Info19b;Info20a;Info20b;Stück;Gewicht;Zahlweise;Forderungsart;Veranlagungsjahr;Fälligkeit;Skontotyp;" & _
"Auftragsnummer;Buchungstyp;Ust-Schlüssel;EU-Land;L+L3;Steuersatz;Erlöskonto;Herkunft;GUID;KOST-Datum;SEPA-Mandatsreferenz;Skontosperre;Gesellschaftername;" & _
"Beteiligtennummer;Identifikationsnummer;Zeichnernummer;Postensperre;SoBil-Sachverhalt;SoBil-Buchung;Festschreibung"

Set FM = frmMain
Set CoDia = FM.comDialo

Select Case GlBut
Case RibTab_Adressen:
        Set RpCo2 = FM.repCont2
        Set RpCls = RpCo2.Columns
        Set RpSel = RpCo2.SelectedRows
Case RibTab_Rechnungen:
        Set RpCo4 = FM.repCont4
        Set RpCls = RpCo4.Columns
        Set RpSel = RpCo4.SelectedRows
Case RibTab_Mahnwesen:
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
Case RibTab_Buchungen:
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
Case RibTab_Ter_Listen:
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
Case RibTab_HomeBanki:
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
End Select

If ManNr > 0 Then
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            ManNa = GlMan(AktZa, 3)
            Exit For
        End If
    Next AktZa
Else
    ManNa = GlMan(GlSMa, 3)
End If
If Len(ManNa) > 2 Then
    ManNa = SUmw(ManNa, False, True, True)
    If Len(ManNa) > 25 Then
        ManNa = Left$(ManNa, 25)
    End If
Else
    ManNa = "Admin"
End If

If InStr(1, ManNa, Chr$(32), 1) > 0 Then
    ManNa = Replace(ManNa, Chr$(32), vbNullString, 1)
End If

If InStr(1, ManNa, Chr$(58), 1) > 0 Then
    ManNa = Replace(ManNa, Chr$(58), vbNullString, 1)
End If

AktZa = 0
GesPo = RpSel.Count

Select Case LCase(ExTyp)
Case "pst":
    If GesPo > 0 Then
    
        SOuOp 'Outlook
        
        Set NaSpa = OutOb.GetNamespace("MAPI") 'NameSpace
        
        If GlMaF = True Then
            Set MapFo = NaSpa.GetDefaultFolder(olFolderContacts)
        Else
            Set MapFo = NaSpa.PickFolder
        End If
        
        If TypeName(MapFo) = "Nothing" Then
            Exit Sub
        ElseIf MapFo.DefaultItemType <> olAppointmentItem Then
            Exit Sub
        End If
    
        Set TeIts = MapFo.Items
        
        Screen.MousePointer = vbHourglass
        
        frmStatus.Show
        DoEvents
        frmStatus.Caption = "Outlookexport"
        Set Lbl01 = frmStatus.lblLab01
        Set PrBr1 = frmStatus.prbStat1
        Set PrBr2 = frmStatus.prbStat2
        Set TxDum = frmStatus.txtDummy
        Lbl01.Caption = "Bitte warten..."
        PrBr1.Min = 0
        PrBr1.Max = GesPo
        
        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then
                Set TeItm = TeIts.Add(olAppointmentItem)
                With TeItm
                    Set RpCol = RpCls.Find(Ter_ID2)
                    TerNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_VonDat)
                    Datu1 = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                    Set RpCol = RpCls.Find(Ter_BisDat)
                    Datu2 = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                    Set RpCol = RpCls.Find(Ter_ZeiVon)
                    Zeit1 = TimeValue(RpRow.Record(RpCol.ItemIndex).Value)
                    Set RpCol = RpCls.Find(Ter_ZeiBis)
                    Zeit2 = TimeValue(RpRow.Record(RpCol.ItemIndex).Value)
                    Set RpCol = RpCls.Find(Ter_IDKurz)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                         TeBet = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        TeBet = "Termin"
                    End If
                    Set RpCol = RpCls.Find(Ter_Patient)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TePat = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    .start = Format$(Datu1, "dd.mm.yyyy") & Chr$(32) & Format$(Zeit1, "hh:mm:ss")
                    .End = Format$(Datu2, "dd.mm.yyyy") & Chr$(32) & Format$(Zeit2, "hh:mm:ss")
                    Set RpCol = RpCls.Find(Ter_Farbtyp)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .BusyStatus = RpRow.Record(RpCol.ItemIndex).Value - 1
                    End If
                    Set RpCol = RpCls.Find(Ter_Selekt)
                    If RpRow.Record(RpCol.ItemIndex).Value = -1 Then
                        .AllDayEvent = True
                    End If
                    Set RpCol = RpCls.Find(Ter_Kommentar)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .Body = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    Set RpCol = RpCls.Find(Ter_Raum)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .Location = GlRmu(RpRow.Record(RpCol.ItemIndex).Value, 1)
                    End If
                    Set RpCol = RpCls.Find(Ter_Priorität)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .Importance = RpRow.Record(RpCol.ItemIndex).Value - 1
                    End If
                    Set RpCol = RpCls.Find(Ter_GuiID) 'GuiID
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .BillingInformation = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    If TeBet = vbNullString Then
                        If TePat = vbNullString Then
                            .Subject = "Termin"
                        Else
                            .Subject = TePat
                        End If
                    Else
                        If TePat = vbNullString Then
                            .Subject = TeBet
                        Else
                            .Subject = TePat & Chr$(32) & TeBet
                        End If
                    End If
                    DoEvents
                    .Save
                End With
                DoEvents
                If GlESy = True Then 'CalDAV / CardDAV / Exchange Synchronisation
                    DBCmEx1 "qryTerRep1", "@IdxNr", TerNr
                End If
                DoEvents
                AktZa = AktZa + 1
                If TxDum.Text = "B" Then Exit For 'Abbrechen
                PrBr1.Value = AktZa
            End If
        Next RpRow

        Unload frmStatus
        DoEvents
        
        Screen.MousePointer = vbNormal
        
        WindowMess GesPo & " Termine wurden erfolgreich exportiert", Dial2, "Outlookexport", FM.hwnd
        
        Set TeItm = Nothing
        Set NaSpa = Nothing
        Set MapFo = Nothing
        Set TeIts = Nothing
        Set OutOb = Nothing
    End If
    
Case "ics":

    If GesPo > 0 Then
                
        DaNam = "EXTF_DATEV_" & Format$(Now, "YYYYMMDD_HHMM") & ".ics"

        With CoDia
            .CancelError = True
            .DialogStyle = 1
            .Filter = "iCalendar Datei (*.ics)|*.ics|Alle Dateien (*.*)|*.*"
            .DefaultExt = "*.ics"
            .DialogTitle = "Bitte Name und Ordner der Exportdatei angeben"
            .FileName = ExOrd & DaNam
            .InitDir = ExOrd
            .ShowSave
            FiNam = .FileName
            If .FileTitle = vbNullString Then
                Set CoDia = Nothing
                Set RpSel = Nothing
                Set RpCls = Nothing
                Set RpCo1 = Nothing
                Set RpCo2 = Nothing
                Set clFil = Nothing
                Set clAnw = Nothing
                Set clICS = Nothing
                Exit Sub
            End If
        End With
        If Right$(FiNam, 4) <> ".ics" Then
            FiNam = FiNam & ".ics"
        End If

        With clFil
            .FilPfa FiNam
            DaPfa = .DaPfa & "\"
            DaNaO = .DaNaO
            ZipNa = DaPfa & DaNaO & ".zip"
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

        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Ter_IDM)
            MitNr = RpRow.Record(RpCol.ItemIndex).Value
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If MitNr = GlMiT(AktZa, 2) Then
                    If GlMiT(AktZa, 27) <> vbNullString Then
                        MiNam = GlMiT(AktZa, 27) 'Anzeigename
                    Else
                        MiNam = GlMiT(AktZa, 1)
                    End If
                    Exit For
                End If
            Next AktZa
            Set RpCol = RpCls.Find(Ter_Raum)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                RmuNr = RpRow.Record(RpCol.ItemIndex).Value
                If RmuNr > UBound(GlRmu) Then
                    RmuNr = 1
                End If
            Else
                RmuNr = 1
            End If
        End If

        Screen.MousePointer = vbHourglass
        
        frmStatus.Show
        DoEvents
        frmStatus.Caption = "ICalendar Export"
        Set Lbl01 = frmStatus.lblLab01
        Set PrBr1 = frmStatus.prbStat1
        Set PrBr2 = frmStatus.prbStat2
        Set TxDum = frmStatus.txtDummy
        Lbl01.Caption = "Bitte warten..."
        PrBr1.Min = 0
        PrBr1.Max = GesPo

        With clICS
            .moKaNa = MiNam
            TmStr = .ICSHead
        End With

        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then
                With clICS
                    Set RpCol = RpCls.Find(Ter_GuiID) 'GuiID
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moTeID = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    Set RpCol = RpCls.Find(Ter_Datum)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moDaAd = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    Set RpCol = RpCls.Find(Ter_VonDat)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moDaSt = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    Set RpCol = RpCls.Find(Ter_BisDat)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moDaEn = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    Set RpCol = RpCls.Find(Ter_ZeiVon)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moZeSt = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    Set RpCol = RpCls.Find(Ter_ZeiBis)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moZeEn = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                    Set RpCol = RpCls.Find(Ter_IDKurz)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TeBet = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        TeBet = "Termin"
                    End If
                    Set RpCol = RpCls.Find(Ter_Patient)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TePat = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        TePat = vbNullString
                    End If
                    Set RpCol = RpCls.Find(Ter_IDKurz)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moTeBe = TeBet & " " & TePat
                    End If
                    Set RpCol = RpCls.Find(Ter_Kommentar)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moTeKo = Left$(RpRow.Record(RpCol.ItemIndex).Value, 62)
                    End If
                    Set RpCol = RpCls.Find(Ter_Priorität)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        .moTePr = RpRow.Record(RpCol.ItemIndex).Value - 1
                    End If
                    Set RpCol = RpCls.Find(Ter_Selekt)
                    If RpRow.Record(RpCol.ItemIndex).Value = -1 Then
                        .moTeGa = True
                    Else
                        .moTeGa = False
                    End If
                    .moTeOr = GlRmu(RmuNr, 1)
                    TmStr = TmStr & .ICSBody
                End With
            End If
            DoEvents
            AktZa = AktZa + 1
            If TxDum.Text = "B" Then Exit For 'Abbrechen
            PrBr1.Value = AktZa
        Next RpRow
        
        TmStr = TmStr & clICS.ICSFoot

        Call clFil.FilCnWr(FiNam, TmStr)
        DoEvents

        Unload frmStatus
        DoEvents
        
        Screen.MousePointer = vbNormal
    End If

Case Else:

    If LCase(ExTyp) = "xls" Then
        DaExt = "csv"
    Else
        DaExt = LCase(ExTyp)
    End If
    
    DaSt4 = Format$(Now, "YYYYMMDD") & Space$(1) & Format$(Now, "HHMMSS")
    DaNam = "EXTF_DATEV_" & Format$(Now, "YYYYMMDD_HHMM") & "." & DaExt

    With clFil
        .hwnd = FM.hwnd
        .StaVe = ExOrd
        .DaExt = DaExt
        .DaNam = ExOrd & DaNam
        .DaTit = "Bitte Name und Ordner der Exportdatei angeben"
        Select Case ExTyp
        Case "xls": .DaStr = "Microsoft Excel (*.csv)" & Chr(0) & "*.xls" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        Case "csv": .DaStr = "DATEV 4.0 Dateien (*.csv)" & Chr(0) & "*.csv" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        Case "txt": .DaStr = "Lexware-Dateien (*.txt)" & Chr(0) & "*.txt" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        Case "xml": .DaStr = "XML-Dateien (*.xml)" & Chr(0) & "*.xml" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        End Select
         FiNam = .FilSav
    End With
    If FiNam = vbNullString Then
        Set CoDia = Nothing
        Set RpSel = Nothing
        Set RpCls = Nothing
        Set RpCo1 = Nothing
        Set RpCo2 = Nothing
        Set clFil = Nothing
        Set clAnw = Nothing
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

    If GesPo > 0 Then
        Screen.MousePointer = vbHourglass
        
        Set Lbl01 = frmStatus.lblLab01
        Set PrBr1 = frmStatus.prbStat1
        Set PrBr2 = frmStatus.prbStat2
        Set TxDum = frmStatus.txtDummy
                
        PrBr1.Min = 0
        PrBr1.Max = GesPo
        PrBr2.Min = 0
        PrBr2.Max = GesPo
        frmStatus.Show
        Lbl01.Caption = "Bitte warten..."
        frmStatus.Caption = "Exportieren"
        DoEvents

        Set RS161 = New ADODB.Recordset 'Kopfspalten
        For Each RpCol In RpCls
            CapSt = RpCol.Caption
            If CapSt <> vbNullString Then
                RS161.Fields.Append CapSt, adVariant
            Else
                RS161.Fields.Append "S" & RpCol.Index, adVariant
            End If
        Next RpCol
        If RS161.State = adStateClosed Then
            RS161.Open
        End If
        DoEvents

        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then

                Select Case GlBut 'Berichtdatum setzen
                Case RibTab_Adressen:
                        IdxNr = AdAry(Adr_ID0, RpRow.Index)
                Case RibTab_Rechnungen:
                        Set RpCol = RpCls.Find(Rec_ID1)
                        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                        DBCmEx2 "qrySimReBer", "@IdDat", "@IdxNr", Date, IdxNr
                Case RibTab_Mahnwesen:
                        Set RpCol = RpCls.Find(OPo_ID1)
                        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                        DBCmEx2 "qrySimOPBer", "@IdDat", "@IdxNr", Date, IdxNr
                Case RibTab_Buchungen:
                        Set RpCol = RpCls.Find(Buh_ID0)
                        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                        DBCmEx2 "qrySimBuBer", "@IdDat", "@IdxNr", Date, IdxNr
                Case RibTab_HomeBanki:
                        Set RpCol = RpCls.Find(Ban_ID2)
                        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                End Select

                RS161.AddNew
                For Each RpCol In RpCls
                    CapSt = RpCol.Caption
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        If RpRow.Record(RpCol.ItemIndex).Caption = "GuiID" Then
                            TmStr = CreateID("B")
                        Else
                            TmStr = RpRow.Record(RpCol.ItemIndex).Value
                        End If
                        For AkZa2 = 1 To 10 'das Sonderzeichen darf mnax. 10x im String vorkommen
                            For AkZa1 = 30 To 180
                                If (AkZa1 < 34 Or AkZa1 > 122) Then 'WICHTIG # muss enthlaten bleiben
                                    If AkZa1 <> 32 Then 'Leerzeichen ausklammern
                                        TmStr = Replace(TmStr, Chr$(AkZa1), vbNullString, 1)
                                    End If
                                End If
                            Next AkZa1
                        Next AkZa2
                        If RS161.Fields(RpCol.ItemIndex).Type = adBoolean Then
                            RS161.Fields(RpCol.ItemIndex).Value = CBool(TmStr)
                        Else
                            If TmStr <> vbNullString Then
                                RS161.Fields(RpCol.ItemIndex).Value = TmStr
                            Else
                                RS161.Fields(RpCol.ItemIndex).Value = 0
                            End If
                        End If
                    End If
                Next RpCol
                RS161.Update
            End If
            
            PrBr1.Value = AktPo
            AktPo = AktPo + 1
            DoEvents
        Next RpRow

        AktPo = 0 'WICHTIG!
        AnzPo = RS161.RecordCount

        If AnzPo > 0 Then
            RS161.MoveFirst
            If GlBut = RibTab_Ter_Listen Then
                DaSt1 = Format$(RS161.Fields("Hinzugefügt").Value, "YYYYMMDD") 'Buchungsdatum erste Buchung
                If IsDate(RS161.Fields("Hinzugefügt").Value) = True Then
                    DaSt3 = DatePart("yyyy", CDate(RS161.Fields("Hinzugefügt").Value), vbMonday) & "0101" 'Wirtschaftsjahresbeginn
                Else
                    DaSt3 = DatePart("yyyy", Date, vbMonday) & "0101" 'Wirtschaftsjahresbeginn
                End If
            Else
                DaSt1 = Format$(RS161.Fields("Datum").Value, "YYYYMMDD") 'Buchungsdatum erste Buchung
                If IsDate(RS161.Fields("Datum").Value) = True Then
                    DaSt3 = DatePart("yyyy", CDate(RS161.Fields("Datum").Value), vbMonday) & "0101" 'Wirtschaftsjahresbeginn
                Else
                    DaSt3 = DatePart("yyyy", Date, vbMonday) & "0101" 'Wirtschaftsjahresbeginn
                End If
            End If
            DoEvents
            RS161.MoveLast
            If GlBut = RibTab_Ter_Listen Then
                DaSt2 = Format$(RS161.Fields("Hinzugefügt").Value, "YYYYMMDD") 'Buchungsdatum letzte Buchung
            Else
                DaSt2 = Format$(RS161.Fields("Datum").Value, "YYYYMMDD") 'Buchungsdatum letzte Buchung
            End If
            DoEvents
            RS161.MoveFirst

            Select Case LCase(ExTyp)
            Case "xls":
                For Each RpCol In RpCls
                    If TmpSt = vbNullString Then
                        Select Case RpCol.Caption
                        Case "Sachkontenbezeichnung": TmpSt = Chr$(34) & "IDKurz" & Chr$(34) & Chr$(59)
                        Case "Buchungstext": TmpSt = Chr$(34) & "Buchtext" & Chr$(34) & Chr$(59)
                        Case "Bericht": TmpSt = Chr$(34) & "Berichtdatum" & Chr$(34) & Chr$(59)
                        Case "Belegzeichen": TmpSt = Chr$(34) & "RechNr" & Chr$(34) & Chr$(59)
                        Case "Nummer": TmpSt = Chr$(34) & "Beleg" & Chr$(34) & Chr$(59)
                        Case "Sachkonto": TmpSt = Chr$(34) & "IDK" & Chr$(34) & Chr$(59)
                        Case "Mandant": TmpSt = Chr$(34) & "IDT" & Chr$(34) & Chr$(59)
                        Case "Mitarbeiter": TmpSt = Chr$(34) & "IDM" & Chr$(34) & Chr$(59)
                        Case Else: TmpSt = Chr$(34) & RpCol.Caption & Chr$(34) & Chr$(59)
                        End Select
                    Else
                        Select Case RpCol.Caption
                        Case "Sachkontenbezeichnung": TmpSt = TmpSt & Chr$(34) & "IDKurz" & Chr$(34) & Chr$(59)
                        Case "Buchungstext": TmpSt = TmpSt & Chr$(34) & "Buchtext" & Chr$(34) & Chr$(59)
                        Case "Bericht": TmpSt = TmpSt & Chr$(34) & "Berichtdatum" & Chr$(34) & Chr$(59)
                        Case "Belegzeichen": TmpSt = TmpSt & Chr$(34) & "RechNr" & Chr$(34) & Chr$(59)
                        Case "Nummer": TmpSt = TmpSt & Chr$(34) & "Beleg" & Chr$(34) & Chr$(59)
                        Case "Sachkonto": TmpSt = TmpSt & Chr$(34) & "IDK" & Chr$(34) & Chr$(59)
                        Case "Mandant": TmpSt = TmpSt & Chr$(34) & "IDT" & Chr$(34) & Chr$(59)
                        Case "Mitarbeiter": TmpSt = TmpSt & Chr$(34) & "IDM" & Chr$(34) & Chr$(59)
                        Case Else: TmpSt = TmpSt & Chr$(34) & RpCol.Caption & Chr$(34) & Chr$(59)
                        End Select
                    End If
                Next RpCol
                TmpSt = TmpSt & vbCrLf
                For Each RpRow In RpSel
                    If RpRow.GroupRow = False Then
                        For Each RpCol In RpCls
                            TmpSt = TmpSt & Chr$(34) & RpRow.Record(RpCol.ItemIndex).Value & Chr$(34) & Chr$(59)
                            If GlBlE = True Then 'DATEV Belegexport
                                If RpCol.ItemIndex = 33 Then 'Datei
                                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                                        BlgNa = RpRow.Record(RpCol.ItemIndex).Value
                                        With clFil
                                            If .FilVor(GlBPf & BlgNa) = True Then
                                                If Len(BlgNa) > 46 Then
                                                    FiExt = Right$(BlgNa, 4)
                                                    BlgNa = Left$(BlgNa, 41) & FiExt
                                                    .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNe & vbNullChar
                                                Else
                                                    .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNa & vbNullChar
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                                If RpCol.ItemIndex = 8 Then 'IDR
                                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                                        AktRe = AktRe + 1
                                        ReDim Preserve GloDr(AktRe)
                                        GloDr(AktRe) = RpRow.Record(RpCol.ItemIndex).Value
                                        Select Case GlBut
                                        Case RibTab_Rechnungen: Set RpCol = RpCls.Find(Rec_RechNr)
                                        Case RibTab_Mahnwesen: Set RpCol = RpCls.Find(OPo_RechNr)
                                        Case RibTab_Buchungen: Set RpCol = RpCls.Find(Buh_RechNr)
                                        Case RibTab_HomeBanki: Set RpCol = RpCls.Find(Ban_RechNr1)
                                        End Select
                                        BlgNa = "Rechnung_Beleg_" & RpRow.Record(RpCol.ItemIndex).Value & ".zip"
                                        BlgPo = BlgPo + 1
                                        ReDim Preserve GlZip(BlgPo)
                                        GlZip(BlgPo) = BlgNa
                                    End If
                                End If
                            End If
                        Next RpCol
                        TmpSt = TmpSt & vbCrLf
                    End If
                    PrBr2.Value = AktPo
                    AktPo = AktPo + 1
                    DoEvents
                Next RpRow
                DoEvents
                With clFil
                    .StrDa = TmpSt
                    RetWe = .FilWrSt
                End With
                DoEvents
                
            Case "csv":

                TmpSt = Chr$(34) & "EXTF" & Chr$(34) & ExSep & "300" & ExSep & "21" & ExSep & Chr$(34) & "Buchungsstapel" & Chr$(34) & ExSep & "4" & ExSep & DaSt4 & "000" & ExSep & ExSep & Chr$(34) & "SV" & CSVep & ManNa & CSVep & Chr$(34) & ExSep & BerSt & ExSep & MaNam & ExSep & DaSt3 & ExSep & AnzKo & ExSep & DaSt1 & ExSep & DaSt2 & ExSep & Chr$(34) & "Buchungen" & CSVep & Chr$(34) & ExSep & "1" & ExSep & "0" & ExSep & Chr$(34) & "0" & Chr$(34) & ExSep & Chr$(34) & "EUR" & Chr$(34) & ExSep & ExSep & ExSep & ExSep & vbCrLf
                TmpSt = TmpSt & DaKop & vbCrLf

                Select Case GlBut
                Case RibTab_Mahnwesen: TmpSt = TmpSt + S_DaExP(RS161, LCase(ExTyp))
                Case RibTab_Buchungen: TmpSt = TmpSt + S_DaExB(RS161, LCase(ExTyp))
                Case RibTab_Rechnungen: TmpSt = TmpSt + S_DaExP(RS161, LCase(ExTyp), , ReAbs)
                End Select

                Select Case GlBut
                Case RibTab_Buchungen:
                    RS161.MoveFirst
                    DoEvents
                    Do
                    If GlBlE = True Then 'DATEV Belegexport
                        If RS161.Fields("Datei").Value <> vbNullString Then
                            BlgNa = RS161.Fields("Datei").Value
                            With clFil
                                If .FilVor(GlBPf & BlgNa) = True Then
                                    If Len(BlgNa) > 46 Then
                                        FiExt = Right$(BlgNa, 4)
                                        BlgNe = Left$(BlgNa, 41) & FiExt
                                        .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNe & vbNullChar
                                    Else
                                        .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNa & vbNullChar
                                    End If
                                End If
                            End With
                        End If
                        If RS161.Fields("IDR").Value <> vbNullString Then
                            If RS161.Fields("IDR").Value > 0 Then
                                AktRe = AktRe + 1
                                ReDim Preserve GloDr(AktRe)
                                GloDr(AktRe) = RS161.Fields("IDR").Value
                                BlgNa = "Rechnung_Beleg_" & RS161.Fields("Belegzeichen").Value & ".pdf"
                                BlgPo = BlgPo + 1
                                ReDim Preserve GlZip(BlgPo)
                                GlZip(BlgPo) = BlgNa
                            End If
                        End If
                    End If
                    RS161.MoveNext
                    PrBr2.Value = AktPo
                    AktPo = AktPo + 1
                    DoEvents
                    Loop Until RS161.EOF
                    DoEvents
                Case RibTab_Rechnungen:
                    RS161.MoveFirst
                    DoEvents
                    Do
                    If GlBlE = True Then 'DATEV Belegexport
                        If RS161.Fields("Zähler").Value <> vbNullString Then
                            If RS161.Fields("Zähler").Value > 0 Then
                                AktRe = AktRe + 1
                                ReDim Preserve GloDr(AktRe)
                                GloDr(AktRe) = RS161.Fields("Zähler").Value
                                BlgNa = "Rechnung_Beleg_" & RS161.Fields("Rechnung").Value & ".pdf"
                                BlgPo = BlgPo + 1
                                ReDim Preserve GlZip(BlgPo)
                                GlZip(BlgPo) = BlgNa
                            End If
                        End If
                    End If
                    RS161.MoveNext
                    PrBr2.Value = AktPo
                    AktPo = AktPo + 1
                    DoEvents
                    Loop Until RS161.EOF
                    DoEvents
                End Select
                                                
                With clFil
                    .StrDa = TmpSt
                    RetWe = .FilWrSt
                End With
                DoEvents

                If GlBlE = True Then 'DATEV Belegexport
                    FilNa = DaPfa & "document.xml"
                    RS161.MoveFirst
                    DoEvents
                    Select Case GlBut
                    Case RibTab_Mahnwesen: S_DaExX RS161, FilNa, 2, DaNam
                    Case RibTab_Buchungen: S_DaExX RS161, FilNa, 2, DaNam
                    Case RibTab_Rechnungen: S_DaExX RS161, FilNa, 3, DaNam
                    End Select
                End If
                
            Case "txt":

                Select Case GlBut
                Case RibTab_Mahnwesen: TmpSt = TmpSt + S_DaExP(RS161, LCase(ExTyp))
                Case RibTab_Buchungen: TmpSt = TmpSt + S_DaExB(RS161, LCase(ExTyp))
                Case RibTab_Rechnungen: TmpSt = TmpSt + S_DaExP(RS161, LCase(ExTyp), , ReAbs)
                End Select
                                
                If GlBut = RibTab_Buchungen Then
                    RS161.MoveFirst
                    DoEvents
                    Do
                    If GlBlE = True Then 'DATEV Belegexport
                        If RS161.Fields("Datei").Value <> vbNullString Then
                            BlgNa = RS161.Fields("Datei").Value
                            With clFil
                                If .FilVor(GlBPf & BlgNa) = True Then
                                    If Len(BlgNa) > 46 Then
                                        FiExt = Right$(BlgNa, 4)
                                        BlgNe = Left$(BlgNa, 41) & FiExt
                                        .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNe & vbNullChar
                                    Else
                                        .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNa & vbNullChar
                                    End If
                                End If
                            End With
                        End If
                        If RS161.Fields("IDR").Value <> vbNullString Then
                            If RS161.Fields("IDR").Value > 0 Then
                                AktRe = AktRe + 1
                                ReDim Preserve GloDr(AktRe)
                                GloDr(AktRe) = RS161.Fields("IDR").Value
                                Select Case GlBut
                                Case RibTab_Rechnungen: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Rechnung").Value & ".zip"
                                Case RibTab_Mahnwesen: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Rechnung").Value & ".zip"
                                Case RibTab_Buchungen: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Belegzeichen").Value & ".zip"
                                Case RibTab_HomeBanki: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Rechnung1").Value & ".zip"
                                End Select
                                BlgNa = "Rechnung_Beleg_" & RpRow.Record(RpCol.ItemIndex).Value & ".zip"
                                BlgPo = BlgPo + 1
                                ReDim Preserve GlZip(BlgPo)
                                GlZip(BlgPo) = BlgNa
                            End If
                        End If
                    End If
                    RS161.MoveNext
                    PrBr2.Value = AktPo
                    AktPo = AktPo + 1
                    DoEvents
                    Loop Until RS161.EOF
                    DoEvents
                End If
                
                With clFil
                    .StrDa = TmpSt
                    RetWe = .FilWrSt
                End With
                
            Case "xml":
                Select Case GlBut
                Case RibTab_Mahnwesen:
                
                        S_XML RS161, FiNam
                        
                Case RibTab_Buchungen:
                        
                        Do
                        If GlBlE = True Then 'DATEV Belegexport
                            If RS161.Fields("Datei").Value <> vbNullString Then
                                BlgNa = RS161.Fields("Datei").Value
                                With clFil
                                    If .FilVor(GlBPf & BlgNa) = True Then
                                        If Len(BlgNa) > 46 Then
                                            FiExt = Right$(BlgNa, 4)
                                            BlgNa = Left$(BlgNa, 41) & FiExt
                                            .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNe & vbNullChar
                                        Else
                                            .DaCop = GlBPf & BlgNa & ";" & DaPfa & BlgNa & vbNullChar
                                        End If
                                    End If
                                End With
                            End If
                            If RS161.Fields("IDR").Value <> vbNullString Then
                                If RS161.Fields("IDR").Value > 0 Then
                                    AktRe = AktRe + 1
                                    ReDim Preserve GloDr(AktRe)
                                    GloDr(AktRe) = RS161.Fields("IDR").Value
                                    Select Case GlBut
                                    Case RibTab_Rechnungen: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Rechnung").Value & ".zip"
                                    Case RibTab_Mahnwesen: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Rechnung").Value & ".zip"
                                    Case RibTab_Buchungen: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Belegzeichen").Value & ".zip"
                                    Case RibTab_HomeBanki: BlgNa = "Rechnung_Beleg_" & RS161.Fields("Rechnung1").Value & ".zip"
                                    End Select
                                    BlgPo = BlgPo + 1
                                    ReDim Preserve GlZip(BlgPo)
                                    GlZip(BlgPo) = BlgNa
                                End If
                            End If
                        End If
                        RS161.MoveNext
                        PrBr2.Value = AktPo
                        AktPo = AktPo + 1
                        DoEvents
                        Loop Until RS161.EOF
                        DoEvents

                        RS161.MoveFirst
                        DoEvents
                        FilNa = DaPfa & "document.xml"
                        Select Case GlBut
                        Case RibTab_Mahnwesen: S_DaExX RS161, FilNa, 2, DaNam
                        Case RibTab_Buchungen: S_DaExX RS161, FilNa, 2, DaNam
                        Case RibTab_Rechnungen: S_DaExX RS161, FilNa, 3, DaNam
                        End Select
                End Select
                
            End Select
        End If

        RS161.Close
        Set RS161 = Nothing

        Unload frmStatus
        DoEvents

        If GlBut = RibTab_Buchungen Then
            If GlBlE = True Then 'DATEV Belegexport
                Set clLis = New clsLisLab
                With clLis
                    .ForNam = "Rechnu"
                    .FilNam = GlFrO & S_FoCh("Rechnu")
                    .PfaTmp = GlTmp
                    .ExpFmt = "PDF"
                    .StaVer = DaPfa
                    .DatNam = DaPfa & DaNaO & ".pdf"
                    .DruDia = False
                    .DruVor = False
                    .MitaVo = GlMiV
                    .ArztVo = GlArV
                    .MandVo = GlMaV
                    .LLExBl
                End With
                Set clLis = Nothing
            End If
        End If
        
        ReDim GloDr(0)
        
        Screen.MousePointer = vbNormal
        DoEvents

        If EmVer > 0 Then
            SMaNe 0, , , DaNam, DaNam, FiNam
        End If
    End If
End Select

DoEvents
GlNeK = GlKoX 'Protokolleintrag
With GlNeK
    .PatNr = GlMan(GlSMa, 2)
    .IdxNr = 0
    .EiDat = Format$(Date, "dd.mm.yyyy")
    .EiZei = TimeValue(Now)
    .EiTyp = 104
    .ZiStr = Format$(Now, "hh:mm") & " Uhr"
    .NeuEi = True
    .KeiAk = True
    .Mitar = GlMiA(GlSmI, 2)
    If LCase(ExTyp) = "pst" Then
        .TeStr = "Exportiert Outlooktermine - " & ExTyp & " - (" & GesPo & " Positionen)"
    Else
        Select Case GlBut
        Case RibTab_Adressen: .TeStr = "Exportiert Adressen - " & ExTyp & " - (" & GesPo & " Positionen)"
        Case RibTab_Mahnwesen: .TeStr = "Exportiert Offene Posten - " & ExTyp & " - (" & GesPo & " Positionen)"
        Case RibTab_Ter_Listen: .TeStr = "Exportiert Termine - " & ExTyp & " - (" & GesPo & " Positionen)"
        Case RibTab_Ter_Akont: .TeStr = "Exportiert Termine - " & ExTyp & " - (" & GesPo & " Positionen)"
        Case RibTab_Ter_Warte: .TeStr = "Exportiert Termine - " & ExTyp & " - (" & GesPo & " Positionen)"
        Case Else: .TeStr = "Exportiert Buchungen - " & ExTyp & " - (" & GesPo & " Positionen)"
        End Select
    End If
End With
S_Prot

SPopu "Datenexport", "Der Exportvorgang wurde abgeschlossen", IC48_Information

Set CoDia = Nothing
Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo4 = Nothing

Set clFil = Nothing
Set clAnw = Nothing
Set clICS = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_Expor " & Err.Number
Exit Sub

End Sub
Public Sub S_MeCm()
On Error GoTo FiErr
'füllt Comboboxen der Mandantencomboboxen

Dim AktZa As Integer
Dim AktKo As Integer
Dim AktSa As Integer
Dim GesZa As Integer
Dim CmLei As XtremeSuiteControls.ComboBox
Dim CmMta As XtremeSuiteControls.ComboBox
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpRws As XtremeReportControl.ReportRows
Dim RpCoT As XtremeReportControl.ReportControl
Dim RpRec As XtremeReportControl.ReportRecord
Dim RpRcs As XtremeReportControl.ReportRecords
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmFra As XtremeCommandBars.CommandBarControl
Dim CmKr1 As XtremeCommandBars.CommandBarControl
Dim CmKr2 As XtremeCommandBars.CommandBarComboBox
Dim CmMa1 As XtremeCommandBars.CommandBarComboBox
Dim CmMa2 As XtremeCommandBars.CommandBarComboBox
Dim CmMi1 As XtremeCommandBars.CommandBarComboBox
Dim CmMi2 As XtremeCommandBars.CommandBarComboBox
Dim CmRau As XtremeCommandBars.CommandBarComboBox
Dim CmGeg As XtremeCommandBars.CommandBarComboBox
Dim CmCo1 As XtremeCommandBars.CommandBarComboBox
Dim CmCo2 As XtremeCommandBars.CommandBarComboBox
Dim CmTSt As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set FM = frmMain
Set RpCoT = FM.repContT
Set CmMta = FM.cmbMitar
Set CmTyp = FM.cmbTypen
Set TrLi5 = FM.trvList5
Set CmLei = FM.cmbLeiTe
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set RpRcs = RpCoT.Records
Set RpRws = RpCoT.Rows

Set CmCo1 = CmBrs.FindControl(CmCo1, SY_TE_Termin_FiltTyp, , True)
Set CmCo2 = CmBrs.FindControl(CmCo2, SY_TE_Termin_FiltIdx, , True)
Set CmMa1 = CmBrs.FindControl(CmMa1, Sta_CmMan, , True)
Set CmMa2 = CmBrs.FindControl(CmMa2, SY_SuMan, , True)
Set CmMi1 = CmBrs.FindControl(CmMi1, SY_SuMit, , True)
Set CmRau = CmBrs.FindControl(CmRau, SY_SuRau, , True)
Set CmGeg = CmBrs.FindControl(CmGeg, SY_SuBuh, , True)
Set CmTSt = CmBrs.FindControl(CmTSt, SY_SuTSt, , True)
Set CmMi2 = CmBrs.FindControl(CmMi2, SY_VB_Vorbe_Mitar, , True)
Set CmKr1 = CmBrs.FindControl(CmKr1, SY_KB_KraBla_Hinzufueg, , True)
Set CmKr2 = CmBrs.FindControl(CmKr2, SY_KB_KraBla_Typen, , True)
Set CmFra = CmBrs.FindControl(CmFra, KA_Frage_Hinzufuegen, , True)

CmMa2.Clear
DoEvents
CmMi1.Clear
DoEvents
CmMi2.Clear
DoEvents
CmMa1.Clear
DoEvents
CmRau.Clear
DoEvents
CmTSt.Clear
DoEvents
CmLei.Clear
DoEvents
CmGeg.Clear
DoEvents
CmKr2.Clear
DoEvents

CmCo1.ListIndex = GlCaF

With TrLi5
    .Nodes.Clear
    Set Knote = .Nodes.Add(, xtpTreeViewFirst, "P900", "Krankenblatttypen", IC16_Doc_View)
    Knote.Expanded = True
    For AktZa = 1 To UBound(GlKrA)
        Set Knote = .Nodes.Add("P900", 4, "P" & Format$(GlKrA(AktZa, 0), "000"), GlKrA(AktZa, 2), IC16_Doc_Norm)
        Knote.ForeColor = GlKrA(AktZa, 3)
        If GlKrA(AktZa, 0) = 104 Then Knote.Checked = True
    Next AktZa
    .Nodes(1).Expanded = True
    Set Knote = .Nodes.Add("P900", 4, "P800", "Terminprotokoll", IC16_Doc_Norm)
End With
DoEvents

With CmMa2 'Mandantenwahl
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa) = GlThe(AktZa, 0)
    Next AktZa
    .ListIndex = 1
    DoEvents
End With

If GlMiV = True Then 'Mitarbeiter vorhanden
    With CmMi1
        For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
            .AddItem GlMiA(AktZa, 1)
            .ItemData(AktZa) = GlMiA(AktZa, 2)
        Next AktZa
        .ListIndex = 1
        DoEvents
    End With
End If

With CmMi2
    .AddItem "Alle aktiven Mitarbeiter"
    .ItemData(1) = 0
    If GlMiV = True Then 'Mitarbeiter vorhanden
        For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
            .AddItem GlMiA(AktZa, 1)
            .ItemData(AktZa + 1) = GlMiA(AktZa, 2)
        Next AktZa
        .ListIndex = 1
        DoEvents
    End If
End With

With CmMa1
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa) = GlThe(AktZa, 0)
    Next AktZa
    If GlRst = True Then 'Mandantenbezogene Datenbegrenzung
        .ListIndex = 1
    Else
        .AddItem "für alle Mandanten"
        .ItemData(AktZa) = 0
        .ListIndex = AktZa
    End If
End With
DoEvents

If GlRaV = True Then 'Räume
    With CmRau
        For AktZa = 1 To UBound(GlRmu)
            .AddItem GlRmu(AktZa, 1)
            .ItemData(AktZa) = GlRmu(AktZa, 2)
        Next AktZa
        .ListIndex = 1
    End With
End If
DoEvents

With CmTSt
    For AktZa = 1 To UBound(GlTeS)
        .AddItem GlTeS(AktZa, 1)
        .ItemData(AktZa) = GlTeS(AktZa, 0)
    Next AktZa
    .ListIndex = 1
    DoEvents
End With

With CmLei
    For AktZa = 1 To UBound(GlBtr)
        .AddItem GlBtr(AktZa, 1)
        .ItemData(.NewIndex) = GlBtr(AktZa, 0)
    Next AktZa
End With
DoEvents

With CmGeg
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            .AddItem GlGeK(AktZa, 3)
            .ItemData(AktZa) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
        .AddItem "Alle Geldkonten"
        .ItemData(AktZa) = 0
        If .ListCount > 0 Then
            .ListIndex = AktZa
        Else
            .ListIndex = 0
        End If
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    AktSa = AktSa + 1
                    .AddItem GlSaK(AktKo, 3)
                    .ItemData(AktSa) = GlSaK(AktKo, 6) '[IDB]
                    Exit For
                End If
            Next AktKo
        Next AktZa
        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(AktZa) = GlGeK(AktZa, 0)
            Next AktZa
            .AddItem "Alle Geldkonten"
            .ItemData(AktZa) = 0
            .ListIndex = AktZa
        Else
            AktSa = AktSa + 1
            .AddItem "Alle Geldkonten"
            .ItemData(AktSa) = 0
            .ListIndex = AktSa
        End If
    End If
End With
DoEvents

If CmTyp.ListCount = 0 Then 'Krankenblatttypen
    With CmTyp
        For AktZa = 1 To UBound(GlKrA)
            If GlKrA(AktZa, 0) < 10 Then
                .AddItem GlKrA(AktZa, 1) & " - " & GlKrA(AktZa, 2)
                .ItemData(AktZa - 1) = GlKrA(AktZa, 0)
            End If
        Next AktZa
        If GlStS > 1 Then 'Standardsteuersatz
            If .ListCount >= 8 Then
                .ListIndex = 8
            Else
                .ListIndex = 0
            End If
        Else
            If CmTyp.ListCount > 1 Then
                .ListIndex = 1
            Else
                .ListIndex = 0
            End If
        End If
    End With
End If
DoEvents

If CmMta.ListCount = 0 Then
    With CmMta
        If GlMsp = True Then 'Mandantenspalte anstelle von Mitarbeiterspalte in Abrechnung
            For AktZa = 1 To UBound(GlMan)
                .AddItem GlMan(AktZa, 1)
                .ItemData(AktZa - 1) = GlMan(AktZa, 2)
            Next AktZa
            .ListIndex = GlSMa - 1
        Else
            If GlMiV = True Then 'Mitarbeiter vorhanden
                For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
                    .AddItem GlMiA(AktZa, 1)
                    .ItemData(AktZa - 1) = GlMiA(AktZa, 2)
                Next AktZa
                .ListIndex = GlSmI - 1
            End If
        End If
    End With
End If
DoEvents

Select Case GlCaF
Case 1:
    CmCo2.Clear
    CmAcs(SY_TE_Termin_FiltIdx).Enabled = False
Case 2:
    If GlRaV = True Then 'Räume
        CmCo2.Clear
        For AktZa = 1 To UBound(GlRmu)
            CmCo2.AddItem GlRmu(AktZa, 1)
            CmCo2.ItemData(AktZa) = GlRmu(AktZa, 2)
        Next AktZa
        CmCo2.ListIndex = GlCaS 'Kalenderfilterinhalt
    End If
    CmAcs(SY_TE_Termin_FiltIdx).Enabled = True
Case 3:
    CmCo2.Clear
    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
            CmCo2.AddItem GlMiT(AktZa, 1)
            CmCo2.ItemData(AktZa) = GlMiT(AktZa, 2)
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlMaT)
            CmCo2.AddItem GlMaT(AktZa, 1)
            CmCo2.ItemData(AktZa) = GlMaT(AktZa, 2)
        Next AktZa
    End If
    CmCo2.ListIndex = GlCaS 'Kalenderfilterinhalt
    CmAcs(SY_TE_Termin_FiltIdx).Enabled = True
End Select
DoEvents

With RpCoT
    .EditItem Nothing, Nothing
    If .Records.Count > 0 Then .Records.DeleteAll
    .Populate
End With

'Kalenderspalte ausblenden

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter + Terminspalte
        Set RpRec = RpRcs.Add()
        Set RpItm = RpRec.AddItem(vbNullString)
        With RpItm
            .HasCheckbox = True
            .Focusable = True
            If CBool(GlMiA(AktZa, 12)) = False Then 'Gespert
                .Checked = True
            End If
        End With
        Set RpItm = RpRec.AddItem(GlMiA(AktZa, 1))
        With RpItm
            .Focusable = False
        End With
        Set RpItm = RpRec.AddItem(GlMiA(AktZa, 2))
    Next AktZa
Else
    For AktZa = 1 To UBound(GlMan)
        Set RpRec = RpRcs.Add()
        Set RpItm = RpRec.AddItem(vbNullString)
        With RpItm
            .HasCheckbox = True
            .Focusable = True
            If CBool(GlMan(AktZa, 11)) = False Then
                .Checked = True
            End If
        End With
        Set RpItm = RpRec.AddItem(GlMan(AktZa, 1))
        With RpItm
            .Focusable = False
        End With
        Set RpItm = RpRec.AddItem(GlMan(AktZa, 2))
    Next AktZa
End If
DoEvents

RpCoT.Populate
Set RpRws = RpCoT.Rows
RpRws.Row(0).Selected = False
DoEvents

For AktZa = 1 To UBound(GlFrT)
    CmFra.CommandBar.Controls.Add xtpControlButton, 1200 + AktZa, GlFrT(AktZa)
Next AktZa
DoEvents

If GlMPl = True Then
    If GlMiV = True Then 'Mitarbeiter vorhanden
        If GlCaS > UBound(GlMiT) Then 'Kalenderfilterinhalt
            GlCaS = 1
        End If
    Else
        GlCaS = 1
    End If
Else
    If GlMaV = True Then 'Mandanten vorhanden
        If GlCaS > UBound(GlMaT) Then 'Kalenderfilterinhalt
            GlCaS = 1
        End If
    Else
        GlCaS = 1
    End If
End If

For AktZa = 1 To UBound(GlKrA) 'Krankenblatttypen
    If GlKrA(AktZa, 0) > 9 Then
        Select Case GlKrA(AktZa, 0)
        Case 24:
        Case 101:
        Case 102:
        Case 104:
        Case 105:
        Case 106:
        Case 108:
        Case Else:
            If CmKr1.CommandBar.Controls.Count <= UBound(GlKrA) Then
                CmKr1.CommandBar.Controls.Add xtpControlButton, 1000 + GlKrA(AktZa, 0), GlKrA(AktZa, 2)
            End If
        End Select
    End If
Next AktZa
DoEvents

With CmKr2
    .AddItem "Alle Eintragstypen", 1
    .AddItem "Krankenblatttypen", 2
    .AddItem "Dokumentationstypen", 3
End With
For AktZa = 1 To UBound(GlKrA) 'Krankenblatttypen
    If CmKr2.CommandBar.Controls.Count <= UBound(GlKrA) Then
        CmKr2.AddItem GlKrA(AktZa, 2), AktZa + 3
    End If
Next AktZa
DoEvents
CmKr2.ListIndex = GlUm2

If CmMa2.Enabled = False Then
    CmMa2.Enabled = True
End If

If CmMa1.Enabled = False Then
    CmMa1.Enabled = True
End If

Set RpRcs = Nothing
Set CmBrs = Nothing
Set RpCoT = Nothing

Exit Sub

FiErr:
If GlDbg = True Then SErLog Err.Description & " S_MeCm " & Err.Number
Resume Next

End Sub
Public Sub S_XML(ByRef RST As ADODB.Recordset, ByVal FiNam As String)
On Error GoTo KoErr
'generiert eine XML Datei für StarMoney Import

Dim PatNr As Long
Dim MsgID As String
Dim CreDa As String
Dim RolDa As String
Dim MaNam As String
Dim GlIDn As String
Dim PaDat As String
Dim PaStr As String
Dim BICSt As String
Dim IBANs As String
Dim PaLan As String
Dim PaAdr As String
Dim NaSpa As String
Dim RecNr As String
Dim GeSum As Double
Dim BetOf As Double
Dim AktZa As Integer
Dim GesZa As Integer

Dim objDOM As New MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMNode
Dim objChildNode As MSXML2.IXMLDOMNode
Dim objGrandChildNode As MSXML2.IXMLDOMNode
Dim objGGClildNode As MSXML2.IXMLDOMNode
Dim objGGGChildNode As MSXML2.IXMLDOMNode
Dim objAttribute As MSXML2.IXMLDOMAttribute
Dim objElement As MSXML2.IXMLDOMElement

Dim Levl0 As MSXML2.IXMLDOMNode
Dim Levl1 As MSXML2.IXMLDOMNode
Dim Levl2 As MSXML2.IXMLDOMNode
Dim Levl3 As MSXML2.IXMLDOMNode
Dim Levl4 As MSXML2.IXMLDOMNode
Dim Levl5 As MSXML2.IXMLDOMNode
Dim Levl6 As MSXML2.IXMLDOMNode
Dim Levl7 As MSXML2.IXMLDOMNode

NaSpa = "urn:iso:std:iso:20022:tech:xsd:pain.008.001.02"

GesZa = RST.RecordCount
MsgID = CreateID("O") 'd284182e-56447187-18420318-20180618
CreDa = Format$(Date, "yyyy-mm-dd") & "T" & Format$(Now, "hh:mm:ss") & "Z" '"2018-06-11T09:23:48Z"
RolDa = Format$(Date + 1, "yyyy-mm-dd")
If GlMan(GlSMa, 32) <> vbNullString Then 'Verkehrsname
    MaNam = SUmw(GlMan(GlSMa, 32), False, True)
Else
    MaNam = SUmw(GlThe(GlSMa, 13), False, True)
End If
If GlThe(GlSMa, 33) <> vbNullString Then '
    GlIDn = GlThe(GlSMa, 33) 'Gläubiger Identifikationsnummer
Else
    GlIDn = vbNullString
End If

If GesZa > 0 Then
    Do
    BetOf = Round(RST.Fields(OPo_OffBetrag).Value, 2)
    GeSum = Round(GeSum + BetOf, 2)
    RST.MoveNext
    Loop Until RST.EOF
    RST.MoveFirst
End If

'Create the main xml node
'Set objNode = objDOM.createNode(NODE_PROCESSING_INSTRUCTION, "xml", "version='1.0' Encoding='UTF-8'")
Set objNode = objDOM.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

objDOM.appendChild objNode

'Create the an attribute name is "name" and value is "nguyen" Ccy="EUR"
'Setting Namespace for all childs

'XML Begin HERE
Set Levl0 = objDOM.CreateNode(NODE_ELEMENT, "Document", vbNullString)

Set objElement = Levl0
Set objAttribute = objDOM.createAttribute("xsi:schemaLocation")
objAttribute.Text = "urn:iso:std:iso:20022:tech:xsd:pain.008.001.02 pain.008.001.02.xsd"
objElement.setAttributeNode objAttribute
Set objAttribute = Nothing
'Second Attribute
Set objAttribute = objDOM.createAttribute("xmlns:xsi")
objAttribute.Text = "http://www.w3.org/2001/XMLSchema-instance"
objElement.setAttributeNode objAttribute
Set objAttribute = Nothing
'Third Attribute
Set objAttribute = objDOM.createAttribute("xmlns")
objAttribute.Text = "urn:iso:std:iso:20022:tech:xsd:pain.008.001.02"
objElement.setAttributeNode objAttribute
Set objAttribute = Nothing
' Create the Parent Node - "CstmrDrctDbtInitn"
Set objNode = objDOM.CreateNode(NODE_ELEMENT, "CstmrDrctDbtInitn", NaSpa)

' Create a child node - "GrpHdr"
Set objChildNode = objDOM.CreateNode(NODE_ELEMENT, "GrpHdr", NaSpa)

' Create a Grand Child Node - "MsgId"
Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "MsgId", NaSpa)
' Set this node value
objGrandChildNode.Text = MsgID

' Append "GrandChildofRoot" to "ChildofRoot"
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "CreDtTm", NaSpa)
objGrandChildNode.Text = CreDa
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "NbOfTxs", NaSpa)
objGrandChildNode.Text = GesZa
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "CtrlSum", NaSpa)
objGrandChildNode.Text = Replace(Format$(GeSum, "####0.00"), ",", ".", 1)
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGGClildNode = objDOM.CreateNode(NODE_ELEMENT, "InitgPty", NaSpa)

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "Nm", NaSpa)
objGrandChildNode.Text = SUmw(GlMan(GlSMa, 1), False, True)
objGGClildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

objChildNode.appendChild objGGClildNode
Set objGGClildNode = Nothing

' Append "GrpHdr"
objNode.appendChild objChildNode
Set objChildNode = Nothing

'First Part of XML Ends Here
Set objChildNode = objDOM.CreateNode(NODE_ELEMENT, "PmtInf", NaSpa)

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "PmtInfId", NaSpa)
objGrandChildNode.Text = MsgID
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "PmtMtd", NaSpa)
objGrandChildNode.Text = "DD"
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "BtchBookg", NaSpa)
objGrandChildNode.Text = "false"
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "NbOfTxs", NaSpa)
objGrandChildNode.Text = GesZa
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "CtrlSum", NaSpa)
objGrandChildNode.Text = Replace(Format$(GeSum, "####0.00"), ",", ".", 1)
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGGGChildNode = objDOM.CreateNode(NODE_ELEMENT, "PmtTpInf", NaSpa)
Set objGGClildNode = objDOM.CreateNode(NODE_ELEMENT, "SvcLvl", NaSpa)

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "Cd", NaSpa)
objGrandChildNode.Text = "SEPA"
objGGClildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

objGGGChildNode.appendChild objGGClildNode
Set objGGClildNode = Nothing

Set objGGClildNode = objDOM.CreateNode(NODE_ELEMENT, "LclInstrm", NaSpa)

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "Cd", NaSpa)
objGrandChildNode.Text = "CORE"
objGGClildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

objGGGChildNode.appendChild objGGClildNode
Set objGGClildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "SeqTp", NaSpa)
objGrandChildNode.Text = "RCUR"
objGGGChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

objChildNode.appendChild objGGGChildNode
Set objGGGChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "ReqdColltnDt", NaSpa)
objGrandChildNode.Text = RolDa
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

Set objGGClildNode = objDOM.CreateNode(NODE_ELEMENT, "Cdtr", NaSpa)

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "Nm", NaSpa)
objGrandChildNode.Text = MaNam
objGGClildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

objChildNode.appendChild objGGClildNode
Set objGGClildNode = Nothing

Set objGGGChildNode = objDOM.CreateNode(NODE_ELEMENT, "CdtrAcct", NaSpa)
Set objGGClildNode = objDOM.CreateNode(NODE_ELEMENT, "Id", NaSpa)

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "IBAN", NaSpa)
objGrandChildNode.Text = GlThe(GlSMa, 18)

objGGClildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing
    
objGGGChildNode.appendChild objGGClildNode
Set objGGClildNode = Nothing
  
objChildNode.appendChild objGGGChildNode
Set objGGGChildNode = Nothing

Set objGGGChildNode = objDOM.CreateNode(NODE_ELEMENT, "CdtrAgt", NaSpa)
Set objGGClildNode = objDOM.CreateNode(NODE_ELEMENT, "FinInstnId", NaSpa)

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "BIC", NaSpa)
objGrandChildNode.Text = GlThe(GlSMa, 31)
objGGClildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing
    
objGGGChildNode.appendChild objGGClildNode
Set objGGClildNode = Nothing
  
objChildNode.appendChild objGGGChildNode
Set objGGGChildNode = Nothing

Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "ChrgBr", NaSpa)
objGrandChildNode.Text = "SLEV"
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing


'--------------------------------------------------

Do
PaStr = SUmw(RST.Fields(OPo_Patient).Value, False, True)
PatNr = RST.Fields(OPo_ID0).Value
RecNr = RST.Fields(OPo_RechNr).Value
BICSt = RST.Fields(OPo_BIC).Value
IBANs = RST.Fields(OPo_IBAN).Value
PaLan = Left$(IBANs, 2)

S_AdDe PatNr 'Adressendetails
With GlADt
    PaDat = Format$(.AdDat, "yyyy-mm-dd")
    PaAdr = SUmw(.AdStr & ", " & .AdPLZ & Space$(1) & .AdOrt, False, True)
End With

'Start of Loop Element
Set objGrandChildNode = objDOM.CreateNode(NODE_ELEMENT, "DrctDbtTxInf", NaSpa)

Set Levl1 = objDOM.CreateNode(NODE_ELEMENT, "PmtId", NaSpa)

Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "EndToEndId", NaSpa)
Levl2.Text = "NOTPROVIDED"
Levl1.appendChild Levl2
Set Levl2 = Nothing

objGrandChildNode.appendChild Levl1
Set Levl1 = Nothing

Set objAttribute = objDOM.createAttribute("Ccy")
objAttribute.Text = "EUR"

Set Levl1 = objDOM.CreateNode(NODE_ELEMENT, "InstdAmt", NaSpa)
Levl1.Text = Replace(Format$(RST.Fields(OPo_OffBetrag).Value, "####0.00"), ",", ".", 1)
Set objElement = Levl1
objElement.setAttributeNode objAttribute
Set objAttribute = Nothing
objGrandChildNode.appendChild Levl1
Set Levl1 = Nothing

Set Levl1 = objDOM.CreateNode(NODE_ELEMENT, "DrctDbtTx", NaSpa)

Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "MndtRltdInf", NaSpa)

Set Levl3 = objDOM.CreateNode(NODE_ELEMENT, "MndtId", NaSpa)
Levl3.Text = Format$(PatNr, "00000000")
Levl2.appendChild Levl3
Set Levl3 = Nothing

Set Levl3 = objDOM.CreateNode(NODE_ELEMENT, "DtOfSgntr", NaSpa)
Levl3.Text = PaDat 'Anlegedatum
Levl2.appendChild Levl3
Set Levl3 = Nothing

Levl1.appendChild Levl2
Set Levl2 = Nothing

Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "CdtrSchmeId", NaSpa)

Set Levl3 = objDOM.CreateNode(NODE_ELEMENT, "Id", NaSpa)

Set Levl4 = objDOM.CreateNode(NODE_ELEMENT, "PrvtId", NaSpa)

Set Levl5 = objDOM.CreateNode(NODE_ELEMENT, "Othr", NaSpa)

Set Levl6 = objDOM.CreateNode(NODE_ELEMENT, "Id", NaSpa)
Levl6.Text = GlIDn
Levl5.appendChild Levl6
Set Levl6 = Nothing

Set Levl6 = objDOM.CreateNode(NODE_ELEMENT, "SchmeNm", NaSpa)

Set Levl7 = objDOM.CreateNode(NODE_ELEMENT, "Prtry", NaSpa)
Levl7.Text = "SEPA"
Levl6.appendChild Levl7
Set Levl7 = Nothing

Levl5.appendChild Levl6
Set Levl6 = Nothing

Levl4.appendChild Levl5
Set Levl5 = Nothing

Levl3.appendChild Levl4
Set Levl4 = Nothing

Levl2.appendChild Levl3
Set Levl3 = Nothing

Levl1.appendChild Levl2
Set Levl2 = Nothing
      
objGrandChildNode.appendChild Levl1
Set Levl1 = Nothing

Set Levl1 = objDOM.CreateNode(NODE_ELEMENT, "DbtrAgt", NaSpa)
Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "FinInstnId", NaSpa)

Set Levl3 = objDOM.CreateNode(NODE_ELEMENT, "BIC", NaSpa)
Levl3.Text = BICSt 'BIC
Levl2.appendChild Levl3
Set Levl3 = Nothing

Levl1.appendChild Levl2
Set Levl2 = Nothing

objGrandChildNode.appendChild Levl1
Set Levl1 = Nothing

Set Levl1 = objDOM.CreateNode(NODE_ELEMENT, "Dbtr", NaSpa)

Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "Nm", NaSpa)
Levl2.Text = PaStr 'Patientenname
Levl1.appendChild Levl2
Set Levl2 = Nothing

Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "PstlAdr", NaSpa)

Set Levl3 = objDOM.CreateNode(NODE_ELEMENT, "Ctry", NaSpa)
Levl3.Text = PaLan 'Land
Levl2.appendChild Levl3
Set Levl3 = Nothing

Set Levl3 = objDOM.CreateNode(NODE_ELEMENT, "AdrLine", NaSpa)
Levl3.Text = PaAdr 'Adresse
Levl2.appendChild Levl3
Set Levl3 = Nothing

Levl1.appendChild Levl2
Set Levl2 = Nothing

objGrandChildNode.appendChild Levl1
Set Levl1 = Nothing

Set Levl1 = objDOM.CreateNode(NODE_ELEMENT, "DbtrAcct", NaSpa)
Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "Id", NaSpa)

Set Levl3 = objDOM.CreateNode(NODE_ELEMENT, "IBAN", NaSpa)
Levl3.Text = IBANs 'IBAN
Levl2.appendChild Levl3
Set Levl3 = Nothing

Levl1.appendChild Levl2
Set Levl2 = Nothing

objGrandChildNode.appendChild Levl1
Set Levl1 = Nothing

Set Levl1 = objDOM.CreateNode(NODE_ELEMENT, "RmtInf", NaSpa)

Set Levl2 = objDOM.CreateNode(NODE_ELEMENT, "Ustrd", NaSpa)
Levl2.Text = RecNr & " (" & MaNam & ")"
Levl1.appendChild Levl2
Set Levl2 = Nothing

objGrandChildNode.appendChild Levl1
Set Levl1 = Nothing

'End of Loop Element
objChildNode.appendChild objGrandChildNode
Set objGrandChildNode = Nothing

RST.MoveNext
Loop Until RST.EOF

'--------------------------------------------------

' Append "GrpHdr"
objNode.appendChild objChildNode
Set objChildNode = Nothing

'XML Ends HERE
' Append "CstmrDrctDbtInitn" to the XML Dom Document

'Levl0
Levl0.appendChild objNode
Set objChildNode = Nothing

objDOM.appendChild Levl0
Set objNode = Nothing

'Set Path here
objDOM.Save FiNam
Set objDOM = Nothing

SPopu "Datenexport", "Die XML Daten wurden erfolgreich exportiert", IC48_Information

Exit Sub

KoErr:
Exit Sub

End Sub
Public Sub S_ZeKo()
On Error GoTo KoErr
'Lädt den Kontenplan und die Gegenkonten

Dim CmEiK As XtremeSuiteControls.ComboBox
Dim CmGeg As XtremeSuiteControls.ComboBox
Dim KtoNr As Long
Dim KtoSt As String
Dim AnzKo As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim Lange As Integer

Set FM = frmZeitraum
Set CmEiK = FM.cmbKonto
Set CmGeg = FM.cmbGegen

Set RS155 = New ADODB.Recordset
RS155.CursorLocation = adUseClient
Set RS155 = DBCmRe1("qrySimBuKtR", "@IdxNr", GlKtR)
If RS155.AbsolutePosition > adPosBOF Then

    Set RS163 = New ADODB.Recordset 'neues Recordset erstellen
    With RS163
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockBatchOptimistic
    End With

    For Each FL101 In RS155.Fields
        RS163.Fields.Append FL101.Name, FL101.Type, FL101.DefinedSize
    Next FL101
    If RS163.State = adStateClosed Then RS163.Open
    
    Do Until RS155.EOF
        RS163.AddNew
        For Each FL101 In RS163.Fields
            If IsNull(RS155.Fields(FL101.Name).Value) = False Then
                If FL101.Name = "IDK" Then
                    If RS155.Fields("IDK").Value <> vbNullString Then
                        If RS155.Fields("IDK").Value > 0 Then
                            KtoNr = RS155.Fields("IDK").Value
                            KtoSt = SBuFo(KtoNr) 'Sachkontenformatierung
                            RS163.Fields("IDK").Value = KtoSt
                        End If
                    End If
                Else
                    RS163.Fields(FL101.Name).Value = RS155.Fields(FL101.Name).Value
                End If
            End If
        Next FL101
        RS155.MoveNext
    Loop
    RS163.UpdateBatch
    
    RS163.MoveFirst
    RS163.Sort = "IDK ASC"

    Do Until RS163.EOF
        If RS163.Fields("IDK").Value > 0 Then
            CmEiK.AddItem RS163.Fields("IDK").Value & Chr$(32) & RS163.Fields("IDKurz").Value
            CmEiK.ItemData(CmEiK.NewIndex) = RS163.Fields("IDK").Value
        End If
        RS163.MoveNext
    Loop

    RS163.Close
    Set RS163 = Nothing
    
    CmEiK.ListIndex = 0
End If
RS155.Close
Set RS155 = Nothing

With CmGeg
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            .AddItem GlGeK(AktZa, 3)
            .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    .AddItem GlSaK(AktKo, 3)
                    .ItemData(AktZa - 1) = GlSaK(AktKo, 6) '[IDB]
                End If
            Next AktKo
        Next AktZa
        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
        End If
    End If
End With

CmGeg.ListIndex = 0

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_ZeKo " & Err.Number
Resume Next

End Sub
Public Function S_ZuIdx(ByVal IdxNr As Long) As Long
On Error GoTo LoErr
'Gibt Zugehörigendetails wieder

Dim AktZa As Long
Dim GesZa As Long

Set RS163 = New ADODB.Recordset
RS163.CursorLocation = adUseClient
Set RS163 = DBCmRe1("qryMailAdPat", "@IdxNr", IdxNr)
GesZa = RS163.RecordCount
If GesZa > 0 Then
    AktZa = 1
    ReDim GlZug(GesZa, 5) 'ist immer mind. 1 Array groß
    Do
    GlZug(AktZa, 0) = RS163.Fields("IDKurz").Value
    GlZug(AktZa, 1) = RS163.Fields("Telefon5").Value
    GlZug(AktZa, 2) = RS163.Fields("Briefanrede").Value
    GlZug(AktZa, 3) = RS163.Fields("Vorname").Value
    GlZug(AktZa, 4) = RS163.Fields("Name").Value
    RS163.MoveNext
    AktZa = AktZa + 1
    Loop Until RS163.EOF
End If
RS163.Close
Set RS163 = Nothing

S_ZuIdx = GesZa

Exit Function

LoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "S_ZuIdx " & Err.Number
Resume Next

End Function

