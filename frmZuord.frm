VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmZuord 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Berichtzuordnung"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   2500
      Left            =   100
      TabIndex        =   1
      Top             =   800
      Width           =   11700
      _Version        =   1048579
      _ExtentX        =   20637
      _ExtentY        =   4410
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont2 
      Height          =   2500
      Left            =   100
      TabIndex        =   2
      Top             =   3600
      Width           =   11700
      _Version        =   1048579
      _ExtentX        =   20637
      _ExtentY        =   4410
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkAdrVe 
      Height          =   240
      Left            =   10000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   1700
      _Version        =   1048579
      _ExtentX        =   2999
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Adressenvergleich"
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   5
      Top             =   6700
      Width           =   12000
      _Version        =   1048579
      _ExtentX        =   21167
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   10000
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schlieﬂen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   8600
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Weiter"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnFunkt 
         Height          =   400
         Left            =   7200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Lˆschen"
         UseVisualStyle  =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   5900
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hilfe"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ProgressBar prbStat1 
      Height          =   350
      Left            =   100
      TabIndex        =   3
      Top             =   6300
      Visible         =   0   'False
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   93
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   80
   End
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   495
      Left            =   210
      TabIndex        =   10
      Top             =   160
      Width           =   11200
      _Version        =   1048579
      _ExtentX        =   19756
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   $"frmZuord.frx":0000
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmZuord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PuBu1 As XtremeSuiteControls.PushButton
Private Rahm0 As XtremeSuiteControls.GroupBox
Private RpRow As XtremeReportControl.ReportRow
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRcs As XtremeReportControl.ReportRecords
Private ChAdr As XtremeSuiteControls.CheckBox

Private Declare Function MoveAbruf Lib "kernel32" Alias "MoveFileExA" (ByVal OldName As String, ByVal NewName As String, ByVal Flags As Long) As Long
Private ImportRan As Boolean
Private ImportBusy As Boolean
Private ImportBad As Boolean
Public FoCsv As String
Public FoIni As String
Public FoHead As String
Public FoSaved As Boolean
Public FoRST As ADODB.Recordset
Public Function SaveAbruf() As Boolean
On Error GoTo SuErr
'Nur erfolgreich zugeordnete oder bewusst verworfene Einreichungen entfernen.
Dim DoneIds As Object
Dim Zeile As XtremeReportControl.ReportRecord
Dim CsvWrite As ADODB.Stream
Dim RestCount As Long
Dim TmpName As String
Dim ErrNum As Long
Dim CsvDone As Boolean
Dim Behalten As Boolean
Dim ArchivDir As String
Dim ArchivNa As String
Dim SubmisID As String

If FoCsv = vbNullString Or FoIni = vbNullString Then Exit Function
Set DoneIds = CreateObject("Scripting.Dictionary")
DoneIds.CompareMode = vbTextCompare
For Each Zeile In Me.repCont1.Records
    Behalten = False
    If Zeile.Item(19).CheckboxState = 0 Then
        Behalten = (Val(Zeile.Item(17).Value) <= 0)
        If Zeile.Item(0).Value <> "FN" Then
            If Val(Zeile.Item(18).Value) <= 0 Then Behalten = True
        End If
    End If
    If Not Behalten Then DoneIds(Zeile.Item(14).Value) = True
Next Zeile

Set CsvWrite = New ADODB.Stream
CsvWrite.Type = adTypeText
CsvWrite.Charset = "windows-1252"
CsvWrite.Open
CsvWrite.WriteText FoHead & vbLf
FoRST.Filter = adFilterNone
If FoRST.RecordCount > 0 Then
    FoRST.MoveFirst
    Do Until FoRST.EOF
        SubmisID = FoRST.Fields("FormSubmissionId").Value
        If Not DoneIds.Exists(SubmisID) Then
            CsvWrite.WriteText FoRST.Fields("CsvRow").Value & vbLf
            If CLng(FoRST.Fields("QuestionNumber").Value) = 1 Then RestCount = RestCount + 1
        End If
        FoRST.MoveNext
    Loop
End If
If GlLog Then 'Nur das aktuelle Dateipaar archivieren.
    ArchivDir = GlTEx & "Archiv\"
    If Dir$(ArchivDir, vbDirectory) = vbNullString Then MkDir ArchivDir
    ArchivNa = ArchivDir & Mid$(FoCsv, InStrRev(FoCsv, "\") + 1)
    If Dir$(ArchivNa) = vbNullString Then FileCopy FoCsv, ArchivNa
    If Dir$(FoIni) <> vbNullString Then
        ArchivDir = GlTmp & "Archiv\"
        If Dir$(ArchivDir, vbDirectory) = vbNullString Then MkDir ArchivDir
        ArchivNa = ArchivDir & Mid$(FoIni, InStrRev(FoIni, "\") + 1)
        If Dir$(ArchivNa) = vbNullString Then FileCopy FoIni, ArchivNa
    End If
End If
TmpName = FoCsv & "." & CreateID("R") & ".tmp"
CsvWrite.SaveToFile TmpName, adSaveCreateNotExist
CsvWrite.Close
If MoveAbruf(TmpName, FoCsv, 9) = 0 Then Err.Raise vbObjectError + 2101, "SaveAbruf", "Restdatei konnte nicht ersetzt werden"
CsvDone = True
TmpName = vbNullString

CsvWrite.Open
CsvWrite.WriteText "[Result]" & vbCrLf & "FormSubmissionCount=" & RestCount & vbCrLf
TmpName = FoIni & "." & CreateID("R") & ".tmp"
CsvWrite.SaveToFile TmpName, adSaveCreateNotExist
CsvWrite.Close
If MoveAbruf(TmpName, FoIni, 9) = 0 Then Err.Raise vbObjectError + 2101, "SaveAbruf", "INI konnte nicht ersetzt werden"
TmpName = vbNullString
If RestCount = 0 Then
    Kill FoCsv
    Kill FoIni
End If
Set CsvWrite = Nothing
Set DoneIds = Nothing
SaveAbruf = True
Exit Function
SuErr:
ErrNum = Err.Number
If GlLog = True Then SLogi "SaveAbruf Err: Restdatei nicht gespeichert #" & ErrNum
On Error Resume Next 'Nur lokale Ressourcen und eigene temporaere Datei bereinigen.
If Not CsvWrite Is Nothing Then
    If CsvWrite.State <> adStateClosed Then CsvWrite.Close
End If
If TmpName <> vbNullString Then Kill TmpName
Set CsvWrite = Nothing
Set DoneIds = Nothing
On Error GoTo 0
If CsvDone Then
    SaveAbruf = True 'Die CSV ist verbindlich; die INI wird zum Lesen nicht benoetigt.
    WindowMess "Die Restfrageboegen sind gespeichert. Die Begleitdatei konnte nicht aktualisiert werden: " & FoIni & vbCrLf & "Fehler #" & ErrNum, Dial2, "Fragebogenimport", Me.hwnd
Else
    WindowMess "Die Restfrageboegen konnten nicht gespeichert werden: " & FoCsv & vbCrLf & "Die Originaldatei bleibt erhalten. Bitte erneut Weiter waehlen oder den Support informieren. Fehler #" & ErrNum, Dial2, "Fragebogenimport", Me.hwnd
End If
End Function

Private Function FePru() As Boolean
On Error GoTo SuErr
Dim Zeile As XtremeReportControl.ReportRecord
Dim RestCount As Long
Dim GutCount As Long
Dim LoeschAnz As Long
Dim OhnePat As Long
Dim Meldung As String

For Each Zeile In Me.repCont1.Records
    If Zeile.Item(19).CheckboxState <> 0 Then
        LoeschAnz = LoeschAnz + 1
    ElseIf Val(Zeile.Item(17).Value) <= 0 Or Zeile.Item(0).Value = vbNullString Then
        RestCount = RestCount + 1
        If Val(Zeile.Item(17).Value) <= 0 Then OhnePat = OhnePat + 1
    Else
        GutCount = GutCount + 1
    End If
Next Zeile
If RestCount > 0 And GutCount = 0 And LoeschAnz = 0 Then
    'Nichts uebernehmbar: kein Import und keine neue CSV/INI, die Abrufdateien bleiben unveraendert liegen
    If OhnePat = RestCount Then
        Meldung = "Es ist noch keinem Fragebogen ein Patient zugeordnet."
    Else
        Meldung = "Es kann noch kein Fragebogen uebernommen werden (Zuordnung oder Vorlage fehlt)."
    End If
    Meldung = Meldung & vbCrLf & "Bitte die Zeilen mit ""?"" per Doppelklick einem Patienten zuordnen oder mit ""Neuer Patient"" anlegen." & vbCrLf & "Um es spaeter zu tun, das Fenster schliessen; die Frageboegen bleiben gespeichert."
    WindowMess Meldung, Dial2, "Fragebogenimport", Me.hwnd
    FePru = True
ElseIf RestCount > 0 Then
    Meldung = GutCount & " Frageboegen werden uebernommen, " & RestCount & " Frageboegen bleiben fuer spaeter gespeichert (Zuordnung oder Vorlage fehlt)."
    If LoeschAnz > 0 Then Meldung = Meldung & vbCrLf & LoeschAnz & " Frageboegen werden bewusst verworfen."
    If OhnePat > 0 Then Meldung = Meldung & vbCrLf & "Zeilen mit ""?"" lassen sich vorher per Doppelklick einem Patienten zuordnen."
    Meldung = Meldung & vbCrLf & "Fortfahren?"
    FePru = (WindowMess(Meldung, Dial1, "Fragebogenimport", Me.hwnd) <> 6)
End If
Exit Function
SuErr:
FePru = True
If GlLog = True Then SLogi "FePru Err: Zuordnung nicht geprueft #" & Err.Number
WindowMess "Die Zuordnung konnte nicht geprueft werden. Die Abrufdateien bleiben erhalten.", Dial2, "Fragebogenimport", Me.hwnd
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo SuErr
Dim Meldung As String
If ImportBusy Then
    Cancel = 1
    Exit Sub
End If
If FoCsv = vbNullString Or FoSaved Then Exit Sub
If ImportBad Then
    WindowMess "Der Import wurde abgebrochen. Die Originaldatei bleibt erhalten. Bitte vor dem erneuten Import den Support informieren.", Dial2, "Fragebogenimport", Me.hwnd
    Exit Sub
End If
If ImportRan Then
    SAnLo
    If Not FoSaved Then Cancel = 1
Else
    Meldung = "Die Frageboegen bleiben fuer spaeter gespeichert. Beim naechsten Abruf wird dieser Abruf erneut angezeigt."
    WindowMess Meldung, Dial2, "Fragebogenimport", Me.hwnd
End If
Exit Sub
SuErr:
Cancel = 1
If GlLog = True Then SLogi "Form_QueryUnload Err: Abruf bleibt erhalten #" & Err.Number
End Sub

Private Sub FeLoe()
On Error GoTo InErr

Dim IdxNr As Long
Dim AnzPo As Long
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

AnzPo = RpSel.Count

Tit1 = "Eintr‰ge Entfernen"
If AnzPo > 1 Then
    Mld1 = "Mˆchten Sie die " & AnzPo & " markierten Eintr‰ge wirklich entfernen?"
Else
    Mld1 = "Mˆchten Sie den aktuellen Eintrag wirklich entfernen?"
End If

If AnzPo > 0 Then
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then
                IdxNr = RpRow.Record(0).Value
                DBCmEx1 "qryLabBeLo", "@IdxNr", IdxNr
            End If
        Next RpRow
        T_FeB
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FeLoe " & Err.Number
Resume Next

End Sub

Private Sub btnFunkt_Click()
    Select Case GlBut
    Case RibTab_LabBericht: FeLoe
    Case RibTab_LabBerichte: FeLoe
    Case Else: F_An3
    End Select
End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case GlBut
Case RibTab_LabBericht:
    TeTit = IniGetOpt("Hilfe", 50671)
    TeMai = IniGetOpt("Hilfe", 50672)
    TeInh = IniGetOpt("Hilfe", 50673)
    TeFus = IniGetOpt("Hilfe", 50674)
Case RibTab_LabBerichte:
    TeTit = IniGetOpt("Hilfe", 50671)
    TeMai = IniGetOpt("Hilfe", 50672)
    TeInh = IniGetOpt("Hilfe", 50673)
    TeFus = IniGetOpt("Hilfe", 50674)
Case Else:
    TeTit = IniGetOpt("Hilfe", 50681)
    TeMai = IniGetOpt("Hilfe", 50682)
    TeInh = IniGetOpt("Hilfe", 50683)
    TeFus = IniGetOpt("Hilfe", 50684)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
On Error GoTo SuErr
If ImportBusy Or ImportBad Then Exit Sub
If FoCsv <> vbNullString Then ImportBusy = True
    
Screen.MousePointer = vbHourglass
DoEvents

Select Case GlBut
Case RibTab_LabBericht:
            On Error GoTo 0
            T_FeA
            DoEvents
            S_LaPV
            DoEvents
            SUpLa
            DoEvents
            Unload Me
Case RibTab_LabBerichte:
            On Error GoTo 0
            T_FeA
            DoEvents
            S_LaPV
            DoEvents
            SUpLa
            DoEvents
            Unload Me
Case Else:
            If ImportRan = False Then
                If FePru Then GoTo Fertig
                ImportBusy = True
                Me.repCont1.Enabled = False
                Me.btnFunkt.Enabled = False
                Me.chkAdrVe.Enabled = False
                Me.btnWeiter.Enabled = False 'Zuordnungspr¸fung
                F_An4 'Adressenvergleich
                DoEvents
                S_AnBoH 'F¸gt einen zugeordneten Fragebogen hinzu
                DoEvents
                S_AnBoP 'Bef¸llen des PropertyGrid
                DoEvents
                SUpAn
                DoEvents
                ImportRan = True
            End If
            SAnLo
            ImportBusy = False
            Me.btnWeiter.Enabled = True
            If FoSaved Then Unload Me
End Select

Fertig:
ImportBusy = False
DoEvents
Screen.MousePointer = vbNormal
Exit Sub
SuErr:
ImportBusy = False
ImportBad = True
Screen.MousePointer = vbNormal
If GlLog = True Then SLogi "btnWeiter_Click Err: Import abgebrochen #" & Err.Number
WindowMess "Der Import wurde abgebrochen. Die Originaldatei bleibt erhalten. Bitte den Support informieren.", Dial2, "Fragebogenimport", Me.hwnd
End Sub

Private Sub chkAdrVe_Click()
On Error Resume Next

Dim AdrVe As Boolean

Set ChAdr = Me.chkAdrVe

If ChAdr.Value = xtpChecked Then
    AdrVe = True
End If

Screen.MousePointer = vbHourglass
DoEvents
    
Select Case GlBut
Case RibTab_LabBericht: T_FeP
Case RibTab_LabBerichte: T_FeP
Case Else:
        FaVe1
        If AdrVe = True Then
            FaVe2
        Else
            F_An2
        End If
End Select

DoEvents
Screen.MousePointer = vbNormal

End Sub
Private Sub Form_Load()
On Error Resume Next

Set ChAdr = Me.chkAdrVe

Me.BackColor = GlBak
AFont Me
SFrame 1, Me.hwnd

Select Case GlBut
Case RibTab_LabBericht:
Case RibTab_LabBerichte:
Case Else: ChAdr.Visible = True
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case GlBut
    Case RibTab_LabBericht:
    Case RibTab_LabBerichte:
    Case Else:
        If Not FoRST Is Nothing Then
            If FoRST.State = adStateOpen Then FoRST.Close
            Set FoRST = Nothing
        End If
    End Select
    Set frmZuord = Nothing
End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next
    
Select Case GlBut
Case RibTab_LabBericht:
Case RibTab_LabBerichte:
Case Else:
    If Row.Record(0).Value = vbNullString Then
        Metrics.ForeColor = 8421504
        Metrics.Font.Strikethrough = True
    End If
End Select

End Sub
Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim AdrVe As Boolean

Set ChAdr = Me.chkAdrVe

If ChAdr.Value = xtpChecked Then
    AdrVe = True
End If

Screen.MousePointer = vbHourglass
DoEvents
    
Select Case GlBut
Case RibTab_LabBericht: T_FeP
Case RibTab_LabBerichte: T_FeP
Case Else:
        FaVe1
        If AdrVe = True Then
            FaVe2
        Else
            F_An2
        End If
End Select

DoEvents
Screen.MousePointer = vbNormal

End Sub
Private Sub repCont1_MouseUp(Button As Integer, Shift As Integer, x As Long, Y As Long)
On Error Resume Next

Dim AdrVe As Boolean

Set ChAdr = Me.chkAdrVe

If ChAdr.Value = xtpChecked Then
    AdrVe = True
End If

Screen.MousePointer = vbHourglass
DoEvents
    
Select Case GlBut
Case RibTab_LabBericht: T_FeP
Case RibTab_LabBerichte: T_FeP
Case Else:
        FaVe1
        If AdrVe = True Then
            FaVe2
        Else
            F_An2
        End If
End Select

DoEvents
Screen.MousePointer = vbNormal

End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error GoTo SuErr

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo1 As XtremeReportControl.ReportControl
Dim ZuKur As String

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpRws = RpCo1.Rows
Set RpSel = RpCo1.SelectedRows

TeTit = "Patientenzuordnung"
TeMai = "Soll die besthende Patientenzuordnung entfernt werden?"
TeInh = "Aufgrund der vom Patienten hinterlassenen E-Mail-Adresse wurde diesem Fragebogen ein Patient zugeordnet. Diese Zuordnung kann entfernt werden."
TeFus = "Nachdem die Zuordnung entfernt wurde, ist es mˆglich einen anderen Patienten zuzuordnen oder einen neuen Patienten hinzuzuf¸gen."

Select Case GlBut
Case RibTab_LabBericht: frmAdrSuch.Show vbModal
Case RibTab_LabBerichte: frmAdrSuch.Show vbModal
Case Else:
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.Record(17).Value <> vbNullString Then
            If RpRow.Record(17).Value > 0 Then
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
                If GlMes = 33565 Then
                    RpRow.Record(16).Value = "?"
                    RpRow.Record(17).Value = 0
                End If
            End If
        End If
        'E-Mail-Treffer ohne Zuordnung: Patientensuche mit diesem Patienten als Treffer oeffnen
        ZuKur = RpRow.Record(16).Value & vbNullString
        If Left$(ZuKur, 11) = "? (E-Mail: " And Right$(ZuKur, 1) = ")" And ZuKur <> "? (E-Mail: mehrere Patienten)" Then
            Load frmAdrSuch
            frmAdrSuch.FVorb Mid$(ZuKur, 12, Len(ZuKur) - 12)
        End If
        frmAdrSuch.Show vbModal
    End If
End Select

Screen.MousePointer = vbNormal

Set RpSel = Nothing
Set RpRws = Nothing
Set RpCo1 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "repCont1_RowDblClick " & Err.Number
Resume Next

End Sub
Private Sub repCont2_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FaVe3
End Sub
