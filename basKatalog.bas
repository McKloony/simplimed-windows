Attribute VB_Name = "basKatalog"
Option Explicit

Private FM As Form
Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Rahm7 As XtremeSuiteControls.GroupBox
Private Rahm8 As XtremeSuiteControls.GroupBox
Private Rahm9 As XtremeSuiteControls.GroupBox
Private CmZif As XtremeSuiteControls.ComboBox
Private CmBez As XtremeSuiteControls.ComboBox
Private CmSta As XtremeCommandBars.StatusBar
Private TbBar As XtremeCommandBars.TabToolBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private ChRow As XtremeReportControl.ReportRow
Private ChRws As XtremeReportControl.ReportRows
Private MoKal As XtremeCalendarControl.DatePicker
Private TabCo As XtremeSuiteControls.TabControl
Private TrLi2 As XtremeSuiteControls.TreeView
Private TrLi3 As XtremeSuiteControls.TreeView
Private LiVw5 As XtremeSuiteControls.ListView
Private Knote As XtremeSuiteControls.TreeViewNode
Private TxDe7 As XtremeSuiteControls.FlatEdit
Private LiFld As FolderViewControl.FolderView
Private LiNod As FolderViewControl.TreeNode
Private LiFi1 As FileViewControl.FileView
Private LiFi2 As FileViewControl.FileView
Private LiFit As FileViewControl.ListItem

Global GlFav As Boolean 'Favoriten Anziegen
Global GlFAE As Boolean 'Favoriten Anamnese
Global GlFBE As Boolean 'Favoriten Begründung
Global GlFDE As Boolean 'Favoriten Diagnose
Global GlFGE As Boolean 'Favoriten Gebühren
Global GlFKD As Boolean 'Favoriten Krankenblattdiagnose
Global GlFKM As Boolean 'Favoriten Krankenblattmedikament
Global GlFLE As Boolean 'Favoriten Laboreinträge
Global GlFME As Boolean 'Favoriten Arzneimittel
Global GlFAR As Boolean 'Favoriten Artikelliste
Global GlFPE As Boolean 'Favoriten Laborparameterauftrag
Global GlFRE As Boolean 'Favoriten Rechnungen
Global GlFTE As Boolean 'Favoriten Termine
Global GlFTX As Boolean 'Favoriten Textverarbeitung
Global GlFBA As Boolean 'Favoriten Banking
Global GlKaN As Boolean 'Neuer Eintrag
Global GlKeN As Boolean 'Neue Kette
Global GlKaE As Boolean 'Expertensystem
Global GlZeM As Boolean 'Zeilenmarker
Global GlNod As String  'Aktueller TreeNode
Global GlKSt As String  'Aktueller TreeNode
Global GlNoT As String  'Nodetext
Global GlKtg As String  'Aktueller TreeNode
Global GlKPo() As Long  'Expertensystem

Private clFil As clsFile
Private clFen As clsFenster
Private clAnw As clsAnwend
Private clLis As clsLisLab

Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Sub FaMain(ByVal IdxNr As Long)
On Error GoTo MeErr

GlKal = True

FaReg

Load frmFragen

Set FM = frmFragen

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    .FenDsk 2
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (600 / 2)
        .FeObn = (GlyGr / 2) - (340 / 2)
    Else
        .FeLin = IniGetVal("KatFrag", "FenLin")
        .FeObn = IniGetVal("KatFrag", "FenObe")
    End If
End With

FaSpl
F_CmL IdxNr
DoEvents
If IdxNr > 0 Then
    F_Lad IdxNr
    DoEvents
End If
DoEvents

clFen.FenMov
DoEvents

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

frmFragen.Show
DoEvents
GlKal = False

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaMain " & Err.Number
Resume Next

End Sub
Public Sub FaVe1()
On Error GoTo LiErr
'Vergleicht die Adressen mit den Stammdaten

Dim PatNr As Long
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows
Dim PuBu1 As XtremeSuiteControls.PushButton

Set FM = frmZuord
Set PuBu1 = FM.btnFunkt
Set RpCo1 = FM.repCont1
Set RpRws = RpCo1.Rows
Set RpSel = RpCo1.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
ElseIf RpRws.Count > 0 Then
    Set RpRow = RpRws(0)
End If

If RpRow.GroupRow = False Then
    If RpRow.Record(19).CheckboxState = 0 Then 'Wenn nicht gelöscht werden soll
        If RpRow.Record(17).Value = vbNullString Then 'PatNr
            PuBu1.Enabled = True
        ElseIf CInt(RpRow.Record(17).Value) = 0 Then
            PuBu1.Enabled = True
        Else
            PuBu1.Enabled = False
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaVe3 " & Err.Number
Resume Next

End Sub

Public Sub FaVe2()
On Error GoTo LiErr
'Vergleicht die Adressen mit den Stammdaten

Dim PatNr As Long
Dim PaGes As String
Dim PaNam As String
Dim PaGeb As String
Dim PaVor As String
Dim PaTel As String
Dim PaEma As String
Dim PaStr As String
Dim PaPLZ As String
Dim PaOrt As String
Dim PaLan As String
Dim AdGes As String
Dim AdNam As String
Dim AdGeb As String
Dim AdVor As String
Dim AdAnr As String
Dim AdTel As String
Dim AdEma As String
Dim AdStr As String
Dim AdPLZ As String
Dim AdOrt As String
Dim AdLan As String
Dim AdKur As String
Dim TmStr As String
Dim PaVer As Integer
Dim AdVer As Integer
Dim Versi(6) As String
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpRcs = RpCo2.Records
Set RpSel = RpCo1.SelectedRows

With RpCo2
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Set RpCls = RpCo2.Columns
With RpCls
    Set RpCol = .Add(0, "Feldname", 120, False)
    Set RpCol = .Add(1, "Besthende Adresse", 250, False)
    Set RpCol = .Add(2, "Neue Adresse", 250, True)
    RpCol.AutoSize = True
    Set RpCol = .Add(3, "Übernehmen", 90, False)
    RpCol.HeaderAlignment = xtpAlignmentCenter
End With

For Each RpCol In RpCls
    RpCol.Alignment = xtpAlignmentLeft
    RpCol.Editable = False
    RpCol.Groupable = True
    RpCol.Sortable = False
Next RpCol

With RpCls(3)
    .Editable = True
    .Alignment = xtpAlignmentCenter
End With

Versi(0) = "keine Angabe"
Versi(1) = "Private Vollversicherung"
Versi(2) = "Private Zusatzversicherung"
Versi(3) = "Beihilfeversichert"
Versi(4) = "Postbeihilfe"
Versi(5) = "Gesetzlich Versichert"

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If RpRow.Record(19).CheckboxState = 0 Then 'Wenn nicht gelöscht werden soll
            If RpRow.Record(17).Value <> vbNullString Then
                PatNr = RpRow.Record(17).Value '[ID0]
                If PatNr > 0 Then
                    PaVer = Left$(RpRow.Record(3).Value, 1)
                    PaGes = Trim$(RpRow.Record(4).Value)
                    PaVor = Trim$(RpRow.Record(5).Value)
                    PaNam = Trim$(RpRow.Record(6).Value)
                    PaGeb = Trim$(RpRow.Record(7).Value)
                    PaEma = Trim$(RpRow.Record(8).Value)
                    PaTel = Trim$(RpRow.Record(9).Value)
                    PaPLZ = Trim$(RpRow.Record(10).Value)
                    PaOrt = Trim$(RpRow.Record(11).Value)
                    PaStr = Trim$(RpRow.Record(12).Value)
                    PaLan = Trim$(RpRow.Record(13).Value)
                    If RpRow.Record(20).Value <> vbNullString Then
                        TmStr = RpRow.Record(20).Value
                    End If

                    S_AdDe PatNr 'Adressendetails
                    With GlADt
                        AdGes = .AdGes
                        AdGeb = .AdGeb
                        AdNam = .AdNam
                        AdVor = .AdVor
                        AdAnr = .AdAnr
                        AdTel = .AdTe1
                        AdEma = .AdTe5
                        AdStr = .AdStr
                        AdPLZ = .AdPLZ
                        AdOrt = .AdOrt
                        AdLan = .AdLan
                        AdVer = .AdVty
                        AdKur = .AdKur
                    End With
                    DoEvents
                    If AdVer = 0 Then
                        AdVer = 1
                    End If

                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Vorname:")
                    Set RpItm = RpRec.AddItem(AdVor)
                    Set RpItm = RpRec.AddItem(PaVor)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 1, 1) = "1" Then
                        RpItm.Checked = True
                    End If

                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Nachname:")
                    Set RpItm = RpRec.AddItem(AdNam)
                    Set RpItm = RpRec.AddItem(PaNam)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 2, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                    
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Geburtstag:")
                    Set RpItm = RpRec.AddItem(AdGeb)
                    Set RpItm = RpRec.AddItem(PaGeb)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 3, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                    
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Straße:")
                    Set RpItm = RpRec.AddItem(AdStr)
                    Set RpItm = RpRec.AddItem(PaStr)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 4, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                    
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("PLZ")
                    Set RpItm = RpRec.AddItem(AdPLZ)
                    Set RpItm = RpRec.AddItem(PaPLZ)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 5, 1) = "1" Then
                        RpItm.Checked = True
                    End If
            
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Ort:")
                    Set RpItm = RpRec.AddItem(AdOrt)
                    Set RpItm = RpRec.AddItem(PaOrt)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 6, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                    
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Land:")
                    Set RpItm = RpRec.AddItem(AdLan)
                    Set RpItm = RpRec.AddItem(PaLan)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 7, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                                                                                                                                 
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("E-Mail:")
                    Set RpItm = RpRec.AddItem(AdEma)
                    Set RpItm = RpRec.AddItem(PaEma)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 8, 1) = "1" Then
                        RpItm.Checked = True
                    End If
            
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Telefon:")
                    Set RpItm = RpRec.AddItem(SRufn(AdTel))
                    Set RpItm = RpRec.AddItem(SRufn(PaTel))
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 9, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                                                            
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Geschlecht:")
                    Set RpItm = RpRec.AddItem(AdGes)
                    Set RpItm = RpRec.AddItem(PaGes)
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 10, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                    
                    Set RpRec = RpRcs.Add()
                    Set RpItm = RpRec.AddItem("Versichert:")
                    Set RpItm = RpRec.AddItem(GlVeA(AdVer - 1, 0))
                    Set RpItm = RpRec.AddItem(Versi(PaVer))
                    Set RpItm = RpRec.AddItem(vbNullString)
                    RpItm.HasCheckbox = True
                    RpItm.Alignment = xtpAlignmentCenter
                    RpItm.Editable = True
                    If Mid$(TmStr, 11, 1) = "1" Then
                        RpItm.Checked = True
                    End If
                        
                End If
            End If
        End If
    End If
End If

RpCo2.Populate

Set RpSel = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaVe2 " & Err.Number
Resume Next

End Sub
Public Sub FaVe3()
On Error GoTo LiErr
'Vergleicht die Adressen mit den Stammdaten

Dim PatNr As Long
Dim TmStr As String
Dim RpRow As XtremeReportControl.ReportRow
Dim RpRo2 As XtremeReportControl.ReportRow
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpSel = RpCo1.SelectedRows
Set RpRws = RpCo2.Rows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If RpRow.Record(19).CheckboxState = 0 Then 'Wenn nicht gelöscht werden soll
            If RpRow.Record(17).Value <> vbNullString Then
                PatNr = RpRow.Record(17).Value '[ID0]
                If PatNr > 0 Then
                    For Each RpRo2 In RpRws
                        TmStr = TmStr & RpRo2.Record(3).CheckboxState
                    Next RpRo2
                End If
            End If
        End If
        If TmStr <> vbNullString Then
            RpRow.Record(20).Value = TmStr
            RpCo1.Populate
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaVe3 " & Err.Number
Resume Next

End Sub
Public Sub FMov(ByVal Flag As Boolean)
On Error GoTo OrErr
'Verschiebt den Eintrag in der Kette

Dim IdxNr As Long
Dim RowNr As Long
Dim AnzPo As Long
Dim GesZa As Long
Dim TxIdx As XtremeSuiteControls.FlatEdit
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpNav As XtremeReportControl.ReportNavigator

Set FM = frmFragen
Set TxIdx = FM.txtIdxNr
Set RpCo1 = FM.repCont1
Set RpSel = RpCo1.SelectedRows
Set RpNav = RpCo1.Navigator

If TxIdx.Text <> vbNullString Then
    If IsNumeric(TxIdx.Text) Then
        If CLng(TxIdx.Text) > 0 Then
            IdxNr = CLng(TxIdx)
        Else
            IdxNr = 0
        End If
    Else
        IdxNr = 0
    End If
Else
    IdxNr = 0
End If

AnzPo = RpSel.Count

If AnzPo = 0 Then Exit Sub

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If Flag = True Then
            RowNr = RpRow.Index - 1
            F_Mov IdxNr, Flag
            F_Rec IdxNr
            If RowNr > 0 Then
                RpNav.MoveToRow RowNr
            Else
                RpNav.MoveFirstRow
            End If
        Else
            RowNr = RpRow.Index + 1
            GesZa = RpCo1.Records.Count
            F_Mov IdxNr, Flag
            F_Rec IdxNr
            If RowNr <= GesZa Then
                RpNav.MoveToRow RowNr
            Else
                RpNav.MoveLastRow
            End If
        End If
    End If
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RpSel = Nothing
Set RpRow = Nothing
Set RpNav = Nothing
Set RpCo1 = Nothing

Set clFen = Nothing

Exit Sub

OrErr:
If GlDbg = True Then SErLog Err.Description & " FMov " & Err.Number
Resume Next

End Sub
Public Sub FPubl(Optional NeuAu As Boolean = False)
On Error GoTo OrErr

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim Posit As Integer

Set FM = frmPublizieren

If GlEKV = False Then 'Emailkonten vorhanden
    TeTit = "E-Mail-Versand"
    TeMai = "Es ist kein E-Mail-Konto vorhanden"
    TeInh = "Um eine E-Mail-Bestätigung für ein ausgefüllten Aufnahmeformular zu erhalten, ist es notwendig mind. ein E-Mail-Konto hinzuzufügen."
    TeFus = "Um ein E-Mail-Konto hinzuzufügen, wechseln Sie in das Modul: Textverarbeitung und dann oben auf Emails. Dort klicken Sie auf die Schaltfläche Emailkonten."
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, frmMain.hwnd
Else
    If GlNaf <> vbNullString Then
        Posit = InStr(1, GlNaf, "terminland", 1)
        If Posit > 0 Then
            S_SeSe 83, vbNullString
            FM.mNeAu = NeuAu
            FM.Show
        Else
            TeTit = "Neuaufnahmeformular vorhanden"
            TeMai = "Das Neuaufnahmeformular wurde bereits publiziert, soll es überschrieben werden?"
            TeInh = "Es ist bereits ein publiziertes Neuaufnahmeformular verfügbar, Sie finden die Adresse jetzt in Ihrer Zwischenablage sowie im Modul Fragebogen."
            TeFus = "Sollte es noch Eingaben Ihrer Patienten geben, die noch nicht abgerufen wurden, gehen diese mit dem erneuten Publizieren des Neuaufnahmeformular verlogen. Bitte rufen sie diese vorher sicherheitshalber einmal ab."
            Clipboard.Clear
            Clipboard.SetText GlNaf
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, frmMain.hwnd
            If GlMes = 33565 Then
                FM.mNeAu = NeuAu
                FM.Show
            End If
        End If
    Else
        FM.mNeAu = NeuAu
        FM.Show
    End If
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPubl " & Err.Number
Resume Next

End Sub

Public Sub FaNeu(Optional ByVal TypNr As Integer = 0, Optional ByVal FrNeu As Boolean = True, Optional ByVal IdxNr As Long, Optional ByVal FaRes As Boolean = False)
On Error GoTo InErr
'Fragebogenfrage hinzufügen

Dim GesZa As Long
Dim FrgID As String
Dim AktCo As VB.Control
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmLab As XtremeCommandBars.CommandBarControl
Dim CmBu1 As XtremeCommandBars.CommandBarControl
Dim CmBu2 As XtremeCommandBars.CommandBarControl
Dim CmBu3 As XtremeCommandBars.CommandBarControl
Dim CmBu4 As XtremeCommandBars.CommandBarControl
Dim CmBu5 As XtremeCommandBars.CommandBarControl
Dim CmTyp As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim ChPfl As XtremeSuiteControls.CheckBox

If FrNeu = True Then
    GlKaN = True
    FaMain 0
    FrgID = CreateID("K")
    FM.txtFelNa.Text = FrgID
Else
    GlKaN = False
    If IdxNr > 0 Then
        FaMain IdxNr
    End If
End If
DoEvents

Set FM = frmFragen
Set ChPfl = FM.chkPflch
Set RpCon = FM.repCont1
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RpCo8 = frmMain.repCont8

Set CmLab = CmBrs.FindControl(CmLab, SY_Cap02, , True)
Set CmTyp = CmBrs.FindControl(CmTyp, SY_SuCm1, , True)
Set CmBu1 = CmBrs.FindControl(CmBu1, SY_OP_Sub_Neu, , True)
Set CmBu2 = CmBrs.FindControl(CmBu2, SY_OP_Sub_Loe, , True)
Set CmBu3 = CmBrs.FindControl(CmBu3, SY_OP_Nav_Vor, , True)
Set CmBu4 = CmBrs.FindControl(CmBu4, SY_OP_Nav_Zuru, , True)
Set CmBu5 = CmBrs.FindControl(CmBu5, SY_OP_Sub_Sav, , True)

If FaRes = True Then
    For Each AktCo In FM.Controls
        Select Case TypeName(AktCo)
        Case "FlatEdit": AktCo.Text = vbNullString
        Case "TextBox": AktCo.Text = vbNullString
        Case "CheckBox": AktCo.Value = 0
        Case "ComboBox": If AktCo.ListCount > 0 Then AktCo.ListIndex = 0
        End Select
    Next AktCo
    With RpCon
        .EditItem Nothing, Nothing
        If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
        If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
        If .Records.Count > 0 Then .Records.DeleteAll
        .Populate
    End With
    FM.txtZeile.Text = "2"
    FM.txtZeich.Text = "20"
    FM.txtMaxZe.Text = "250"
    FM.txtSorte.Text = GesZa + 1
    ChPfl.Value = xtpUnchecked
End If

GesZa = RpCo8.Records.Count

If FrNeu = True Then
    CmTyp.ListIndex = TypNr
Else
    If TypNr > 0 Then
        CmTyp.ListIndex = TypNr
    Else
        TypNr = CmTyp.ListIndex
    End If
End If

Select Case TypNr
Case 1: 'Textfeld
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = True
    FM.txtVorga.Enabled = True
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = True
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = True
    CmBu1.Enabled = False
    CmBu2.Enabled = False
    CmBu3.Enabled = False
    CmBu4.Enabled = False
    CmBu5.Enabled = False
    RpCon.Enabled = False
Case 2: 'Auswahlfeld
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 3: 'Ankreuzfeld
    FM.chkVorga.Enabled = True
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = False
    FM.txtZeich.Enabled = False
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = False
    FM.updCont2.Enabled = False
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 4: 'Einfachauswahl
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = False
    FM.txtZeich.Enabled = False
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = False
    FM.updCont2.Enabled = False
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 5: 'Mehrfachauswahl
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = True
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = False
    FM.txtZeich.Enabled = False
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = False
    FM.updCont2.Enabled = False
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 6: 'Zwischentext
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = False
    CmBu1.Enabled = False
    CmBu2.Enabled = False
    CmBu3.Enabled = False
    CmBu4.Enabled = False
    CmBu5.Enabled = False
    RpCon.Enabled = False
    ChPfl.Enabled = False
Case 7: 'Datumsfeld
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = True
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = True
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = True
    CmBu1.Enabled = False
    CmBu2.Enabled = False
    CmBu3.Enabled = False
    CmBu4.Enabled = False
    CmBu5.Enabled = False
    RpCon.Enabled = False
End Select

If FrNeu = False Then
    FM.txtDummy.SetFocus
Else
    FM.txtSorte.Text = GesZa + 1
End If

Set RpCon = Nothing
Set RpCo8 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaNeu " & Err.Number
Resume Next

End Sub
Public Sub FaPos()
On Error GoTo PoErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim TxFra As XtremeSuiteControls.FlatEdit
Dim TxBer As XtremeSuiteControls.FlatEdit
Dim TxVor As XtremeSuiteControls.FlatEdit
Dim TxFel As XtremeSuiteControls.FlatEdit
Dim CmAbh As XtremeSuiteControls.ComboBox
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmFragen
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set TxFra = FM.txtBezei
Set TxBer = FM.txtBeTex
Set TxFel = FM.txtFelNa
Set TxVor = FM.txtVorga
Set CmAbh = FM.cmbAbhen
Set RpCon = FM.repCont1
Set CmBrs = FM.comBar02

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    Rahm1.Move ClLin, ClObn, ClBre, ClHoh
    Rahm2.Move ClLin, ClObn, ClBre, ClHoh
    RpCon.Move 0, 0, ClBre, ClHoh
    TxFra.Width = ClBre - 1700
    TxBer.Width = ClBre - 1700
    TxVor.Width = ClBre - 1700
    CmAbh.Width = ClBre - 1700
    TxFel.Width = ClBre - 7200
End If

Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaPos " & Err.Number
Resume Next

End Sub
Private Sub FaReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

Set FM = frmFragen

If IniGetSek(GlINI, "KatFrag") = False Then
    xGro = 721
    yGro = 471

    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)

    IniSetSek "KatFrag"
    IniSetVal "KatFrag", "FenLin", xPos
    IniSetVal "KatFrag", "FenObe", yPos
    IniSetVal "KatFrag", "FenBre", xGro
    IniSetVal "KatFrag", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaReg " & Err.Number
Resume Next

End Sub
Public Sub FaSpl()
On Error GoTo SpErr
'Formratieren der Spalten

Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmFragen
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns

If RpCls.Count = 0 Then
    With RpCls
        Set RpCol = .Add(0, "ID6", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "IDSub", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Antwortauswahl", 300, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = True
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = False
        End With
        Set RpCol = .Add(3, "Antworttext", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = True
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        Set RpCol = .Add(4, "", 30, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentIconCenter
            .Alignment = xtpAlignmentIconCenter
            .Icon = IC16_Check
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(5, "Cookie", 0, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentIconCenter
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(6, "Sorter", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaSpl " & Err.Number
Resume Next

End Sub
Public Sub KaMain(ByVal IdxNr As Long)
On Error GoTo MeErr

GlKal = True

Screen.MousePointer = vbHourglass
DoEvents

KaReg
DoEvents

Load frmKaEdit

Set FM = frmKaEdit

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    .FenDsk 2
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (600 / 2)
        .FeObn = (GlyGr / 2) - (340 / 2)
    Else
        .FeLin = IniGetVal("KatBear", "FenLin")
        .FeObn = IniGetVal("KatBear", "FenObe")
    End If
End With

If IdxNr > 0 Then
    K_Lad IdxNr
    DoEvents
End If
K_Com
DoEvents

With clFen
    .FenMov
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

frmKaEdit.Show
DoEvents

Screen.MousePointer = vbNormal
DoEvents

GlKal = False

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaMain " & Err.Number
Resume Next

End Sub
Public Sub KaNeu()
On Error GoTo OrErr
'Lädt den Detailbereich

Dim KatNr As Long
Dim GrpNr As Long
Dim GesZa As Long
Dim TreKy As String
Dim AktZa As Integer
Dim AktCo As Control
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmTyp As XtremeCommandBars.CommandBarComboBox
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl

TreKy = Left$(GlNod, 1)
KatNr = Mid$(GlNod, 2, Len(GlNod) - 1)

Set RpCo8 = frmMain.repCont8

GesZa = RpCo8.Records.Count

Select Case GlBut
Case RibTab_Kat_Ketten:
    GlKeN = True
    GlKeE = False
    EMain 0
    Set FM = frmKetten
    Set RpCo4 = FM.repCont4
    Set RpCo5 = FM.repCont5
    Set CmBrs = FM.comBar02
    Set CmAcs = CmBrs.Actions
    Set CmSta = CmBrs.StatusBar

    RpCo4.Enabled = False
    RpCo5.Enabled = False
    
    CmAcs(KA_Edit_Einfuegen).Enabled = False
    CmAcs(KA_Edit_Entfernen).Enabled = False
    CmAcs(KA_Edit_NachOben).Enabled = False
    CmAcs(KA_Edit_NachUnten).Enabled = False
    
    Set CmEdt = CmBrs.FindControl(CmEdt, KA_KeKur, , True)
    CmEdt.Execute
    
    SPopu "Kettenname", "Bitte geben Sie nun der Kette ein Kürzel und eine Bezeichnung", IC48_Information
Case RibTab_Kat_Eintrg:
    GlKaN = True
    KaMain 0
    Set FM = frmKaEdit
    DoEvents
    Select Case TreKy
    Case "A": 'Gebührenkatalog
        For AktZa = 1 To UBound(GlGKa)
            If GlGKa(AktZa, 0) = KatNr Then
                GrpNr = GlGKa(AktZa, 2)
                Exit For
            End If
        Next AktZa
        FM.txtGrupe.Text = GrpNr
        FM.chkFakPr.Visible = True
        FM.txtPrei2.Visible = True
        FM.txtSteue.Visible = False
        FM.lblLabl8.Visible = False
        FM.lblLabl4.Caption = "Diagnose :"
        FM.txtPrei1.Text = "0,0"
        If FM.cmbGrupe.ListCount = 0 Then
            K_CmL 4
            K_CmL 8
        End If
    Case "C": 'Diagnosen
        FM.txtGrupe.Text = 1
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.txtPrei1.Visible = False
        FM.txtPrei2.Visible = False
        FM.txtMulti.Visible = False
        FM.txtSteue.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.lblLabl4.Visible = False
        FM.lblLabl2.Visible = False
        FM.lblLabl3.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl8.Visible = False
        FM.lblLabl1.Caption = "ICD-10 Code :"
        If FM.cmbGrupe.ListCount = 0 Then
            K_CmL 5
        End If
    Case "G": 'Laborparameter
        For AktZa = 1 To UBound(GlGKa)
            If GlGKa(AktZa, 0) = KatNr Then
                GrpNr = GlGKa(AktZa, 2)
                Exit For
            End If
        Next AktZa
        FM.txtPrei1.Text = "0,0"
        FM.txtPrei2.Text = "0,0"
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl8.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.txtSteue.Visible = False
        FM.lblLabl2.Caption = "Patientenpreis:"
        FM.lblLabl4.Caption = "LGM-Preis:"
        For Each AktCo In frmKaEdit.Controls
             Select Case TypeName(AktCo)
             Case "FlatEdit": AktCo.Text = vbNullString
             Case "TextBox": AktCo.Text = vbNullString
             Case "CheckBox": AktCo.Value = 0
             Case "ComboBox": If AktCo.ListCount > 0 Then AktCo.ListIndex = 0
             End Select
        Next AktCo
        FM.txtGrupe.Text = GrpNr
        If FM.cmbStatu.ListCount = 0 Then K_CmL 1
        If FM.cmbGrupp.ListCount = 0 Then K_CmL 2
        If FM.cmbProbe.ListCount = 0 Then K_CmL 3
        If FM.cmbGrupe.ListCount = 0 Then K_CmL 2
    Case "I": 'Arzneimittel
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.cmbGrupe.Visible = False
        FM.lblLabl6.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl1.Caption = "PZN :"
        FM.lblLabl2.Caption = "Packungspreis :"
        FM.txtPrei1.Text = "0,0"
        FM.txtPrei2.Text = "0,0"
    Case "K": 'Begründungen
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.cmbGrupe.Visible = False
        FM.txtPrei1.Visible = False
        FM.txtPrei2.Visible = False
        FM.txtMulti.Visible = False
        FM.txtMinut.Visible = False
        FM.txtSteue.Visible = False
        FM.lblLabl4.Visible = False
        FM.lblLabl5.Visible = False
        FM.lblLabl2.Visible = False
        FM.lblLabl3.Visible = False
        FM.lblLabl6.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl8.Visible = False
        FM.lblLab30.Visible = False
        FM.UpDown1.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.lblLabl1.Caption = "Suchkürzel :"
    Case "L": 'Anamnesetexte
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.cmbGrupe.Visible = False
        FM.txtPrei1.Visible = False
        FM.txtPrei2.Visible = False
        FM.txtMulti.Visible = False
        FM.txtMinut.Visible = False
        FM.txtSteue.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.lblLabl4.Visible = False
        FM.lblLabl5.Visible = False
        FM.lblLabl2.Visible = False
        FM.lblLabl3.Visible = False
        FM.lblLabl6.Visible = False
        FM.updCont1.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl8.Visible = False
        FM.lblLabl1.Caption = "Suchkürzel :"
    Case "M": 'Terminbetreffs
        If FM.cmbMitar.ListCount = 0 Then
            K_CmL 7
        End If
    Case "N": 'Fragebogen
        FM.txtGrupe.Text = 1
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.txtPrei1.Visible = False
        FM.txtPrei2.Visible = False
        FM.txtMulti.Visible = False
        FM.txtMinut.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.txtSteue.Visible = False
        FM.lblLabl4.Visible = False
        FM.lblLabl5.Visible = False
        FM.lblLabl2.Visible = False
        FM.lblLabl3.Visible = False
        FM.lblLabl6.Visible = False
        FM.lblLabl1.Visible = False
        FM.txtZiff1.Visible = False
        FM.updCont1.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl8.Visible = False
        FM.Label17.Caption = "Antworttext :"
        FM.Label18.Caption = "Auswahlvorgaben :"
        FM.Label3.Caption = "Anamnese :"
        If FM.cmbGrupe.ListCount = 0 Then K_CmL 6
    Case "O": 'Textphrasen
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.cmbGrupe.Visible = False
        FM.txtPrei1.Visible = False
        FM.txtPrei2.Visible = False
        FM.txtMulti.Visible = False
        FM.txtMinut.Visible = False
        FM.txtSteue.Visible = False
        FM.lblLabl4.Visible = False
        FM.lblLabl5.Visible = False
        FM.lblLabl2.Visible = False
        FM.lblLabl3.Visible = False
        FM.lblLabl6.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl8.Visible = False
        FM.lblLab30.Visible = False
        FM.UpDown1.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.lblLabl1.Caption = "Suchkürzel :"
    Case "P": 'Artikelliste
        FM.chkAnalo.Visible = False
        FM.chkFakPr.Visible = False
        FM.cmpEiTyp.Visible = False
        FM.cmbGrupe.Visible = False
        FM.lblLabl6.Visible = False
        FM.lblLabl7.Visible = False
        FM.lblLabl1.Caption = "Artikelnummer :"
        FM.lblLabl2.Caption = "Packungspreis :"
        FM.txtPrei1.Text = "0,0"
        FM.txtPrei2.Text = "0,0"
        FM.txtBeSol.Text = "0,0"
        FM.txtBeMin.Text = "0,0"
        FM.txtBeMax.Text = "0,0"
        FM.txtBeMel.Text = "0,0"
        FM.txtBeIst.Text = "0,0"
        FM.txtMeBes.Text = "0,0"
        FM.txtMeEin.Text = "0,0"
        FM.txtDaBes.Text = Format$(Date, "dd.mm.yyyy")
    End Select
    DoEvents
    FM.txtSorte.Text = GesZa + 1
End Select

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo8 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaNeu " & Err.Number
Resume Next

End Sub
Public Sub KaPos()
On Error GoTo PoErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmKaEdit
Set CmBez = FM.cmbBezei
Set FTex1 = FM.txtKomme
Set FTex2 = FM.txtZusat
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set Rahm8 = FM.frmRahm8
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set CmBrs = FM.comBar02

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    Rahm1.Move ClLin, ClObn, ClBre, ClHoh
    Rahm2.Move ClLin, ClObn, ClBre, ClHoh
    Rahm3.Move ClLin, ClObn, ClBre, ClHoh
    Rahm4.Move ClLin, ClObn, ClBre, ClHoh
    Rahm5.Move ClLin, ClObn, ClBre, ClHoh
    Rahm6.Move ClLin, ClObn, ClBre, ClHoh
    Rahm7.Move ClLin, ClObn, ClBre, ClHoh
    Rahm8.Move ClLin, ClObn, ClBre, ClHoh
    FTex1.Width = ClBre - 210
    FTex2.Move 100, 1900, ClBre - 210, ClHoh - 2000
    RpCo1.Move 100, 1900, ClBre - 220, ClHoh - 2000
    RpCo2.Move 100, 100, ClBre - 220, ClHoh - 200
    CmBez.Width = ClBre - 1450
End If

Set CmBrs = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaPos " & Err.Number
Resume Next

End Sub
Private Sub KaReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

Set FM = frmKaEdit

If IniGetSek(GlINI, "KatBear") = False Then
    xGro = 600
    yGro = 340

    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)

    IniSetSek "KatBear"
    IniSetVal "KatBear", "FenLin", xPos
    IniSetVal "KatBear", "FenObe", yPos
    IniSetVal "KatBear", "FenBre", xGro
    IniSetVal "KatBear", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaReg " & Err.Number
Resume Next

End Sub
Public Sub KaRes()
On Error GoTo InErr

Dim AktCo As VB.Control
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CoPic As XtremeSuiteControls.ColorPicker

Set FM = frmKaEdit
Set CoPic = FM.colPick1
Set FTex1 = FM.txtMnute
Set FTex2 = FM.txtNachl
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

For Each AktCo In FM.Controls
    Select Case TypeName(AktCo)
    Case "FlatEdit": AktCo.Text = vbNullString
    Case "TextBox": AktCo.Text = vbNullString
    Case "CheckBox": AktCo.Value = 0
    Case "ComboBox": If AktCo.ListCount > 0 Then AktCo.ListIndex = 0
    End Select
Next AktCo

CoPic.SelectedColor = vbWhite
FTex1.Text = "0"
FTex2.Text = "0"

CmAcs(SY_OP_Loeschen).Enabled = True

GlKaN = True

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaRes " & Err.Number
Resume Next

End Sub


Public Sub KaEdi()
On Error GoTo OrErr
'Lädt den Detailbereich

Dim IdxNr As Long
Dim KetNa As String
Dim KetKu As String
Dim TreKy As String
Dim TmStr As String
Dim TypNr As Integer
Dim AktZa As Integer
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set RpCls = RpCo8.Columns
Set RpSel = RpCo8.SelectedRows

If GlBut = RibTab_Kat_Frage Then
    GlKaN = False
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Kat_ID0)
            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Kat_Preis1)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                TmStr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                TmStr = GlFrT(1)
            End If
            For AktZa = 1 To UBound(GlFrT)
                If GlFrT(AktZa) = TmStr Then
                   TypNr = AktZa
                    Exit For
                End If
            Next AktZa
            FaNeu TypNr, False, IdxNr
        End If
    End If
Else
    TreKy = Left$(GlNod, 1)
    If TreKy = "D" Or TreKy = "F" Or TreKy = "H" Or TreKy = "J" Or TreKy = "R" Or TreKy = "Q" Then
        GlKeN = False
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Kat_ID0)
                IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Kat_GOID)
                KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Kat_IDKurz)
                KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                GlKeE = False
                EMain IdxNr, KetNa, KetKu
            End If
        End If
    Else
        GlKaN = False
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Kat_ID0)
                IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                KaMain IdxNr
            End If
        End If
    End If
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo8 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaEdi " & Err.Number
Resume Next

End Sub
Public Sub KKata(Optional ByVal NoSel As Boolean = False)
On Error GoTo PeErr
'Wählt den korrekten Eintrag in der Kategorieauswahl

Dim TreKy As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmKat As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01

TreKy = Left$(GlNod, 1)

Select Case GlBut
Case RibTab_Kat_Eintrg: Set CmKat = CmBrs.FindControl(CmKat, KA_Eint_KatCombo, , True)
Case RibTab_Kat_Ketten: Set CmKat = CmBrs.FindControl(CmKat, KA_Kett_KatCombo, , True)
End Select

Select Case GlBut
Case RibTab_Kat_Eintrg:
    Select Case TreKy
    Case "A": CmKat.ListIndex = 1 'Gebührenkataloge
    Case "C": CmKat.ListIndex = 2 'Diagnosekataloge
    Case "G": CmKat.ListIndex = 3 'Laborparameter
    Case "I": CmKat.ListIndex = 4 'Arzneikataloge
    Case "K": CmKat.ListIndex = 5 'Begründungskatalog
    Case "L": CmKat.ListIndex = 6 'Anamnesekataloge
    Case "M": CmKat.ListIndex = 7 'Terminbetreffs
    Case "N": CmKat.ListIndex = 8 'Fragebogenkataloge
    Case "O": CmKat.ListIndex = 9 'Textphrasenkataloge
    Case "P": CmKat.ListIndex = 10 'Artikelkatalog
    End Select
Case RibTab_Kat_Ketten
    Select Case TreKy
    Case "D": CmKat.ListIndex = 1 'Gebührenketten
    Case "F": CmKat.ListIndex = 2 'Diagnoseketten
    Case "H": CmKat.ListIndex = 3 'Laborketten
    Case "J": CmKat.ListIndex = 4 'Arzneiketten
    Case "R": CmKat.ListIndex = 5 'Terminketten
    Case "Q": CmKat.ListIndex = 6 'Artikelketten
    End Select
End Select

Exit Sub

PeErr:
If GlDbg = True Then SErLog Err.Description & " KKata " & Err.Number
Resume Next

End Sub

Public Sub KaSpl()
On Error GoTo SpErr
'Formratieren der Spalten

Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmKaEdit
Set RpCo2 = FM.repCont2
Set RpCls = RpCo2.Columns

If RpCls.Count = 0 Then
    With RpCls
        Set RpCol = .Add(0, "ID1", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID4", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "ID3", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(3, "Diagnosegruppe", 10, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If RpCo2.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
    End With
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo2 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaSpl " & Err.Number
Resume Next

End Sub
Public Sub KaSpr()
On Error GoTo SpErr
'Formratieren der Spalten

Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmKaEdit
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

If RpCls.Count = 0 Then
    With RpCls
        Set RpCol = .Add(KaB_ID1, "ID1", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(KaB_IDx, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(KaB_IDy, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(KaB_GOID, "Nummer/Code", 80, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(KaB_IDKurz, "Darf nicht am selben Tag abgerechnet werden mit:", 10, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If RpCo1.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(KaB_Anz, "x", 0, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(KaB_Multi, "Faktor", 0, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(KaB_Preis1, "Preis", 0, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(KaB_Sorter, "Sorter", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KaSpr " & Err.Number
Resume Next

End Sub
Public Sub KButt()
On Error GoTo LaErr
'wird ausgeführt, wenn Katalogcombo geklickt wird

Dim TreKy As String
Dim KeyNr As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmKat As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set TrLi2 = FM.trvList2
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Select Case GlBut
Case RibTab_Kat_Eintrg: Set CmKat = CmBrs.FindControl(CmKat, KA_Eint_KatCombo, , True)
Case RibTab_Kat_Ketten: Set CmKat = CmBrs.FindControl(CmKat, KA_Kett_KatCombo, , True)
End Select

GlAkt = True
GlKaG = True
GlGrD = False

TreKy = Chr$(CmKat.ItemData(CmKat.ListIndex))

For Each Knote In TrLi2.Nodes
    If Left$(Knote.Key, 1) = TreKy Then
        If IsNumeric(Mid$(Knote.Key, 2, 1)) = True Then
            KeyNr = CInt(Mid$(Knote.Key, 2, 1))
            If KeyNr > 0 Then
                GlNod = TreKy & KeyNr
                Exit For
            End If
        End If
    End If
Next Knote

With TrLi2
    For Each Knote In .Nodes
        Knote.Expanded = False
    Next Knote
    For Each Knote In .Nodes
        If Knote.Key = GlNod Then
            .Nodes(GlNod).EnsureVisible
            .Nodes(GlNod).Expanded = True
            .Nodes(GlNod).Selected = True
            .Nodes(GlNod).Image = IC16_Folder_Open
            Exit For
        End If
    Next Knote
End With

SSpLaK
DoEvents
K_List

Select Case GlBut
Case RibTab_Kat_Eintrg:
    If Left$(GlNod, 1) = "P" Then
        CmAcs(KA_Eint_Druck01).Enabled = True
        CmAcs(KA_Eint_Druck02).Enabled = True
    Else
        CmAcs(KA_Eint_Druck01).Enabled = False
        CmAcs(KA_Eint_Druck02).Enabled = False
    End If
Case RibTab_Kat_Ketten:
    CmAcs(KA_Eint_Druck01).Enabled = False
    CmAcs(KA_Eint_Druck02).Enabled = False
End Select

Set CmSta = Nothing
Set CmBrs = Nothing

GlKaG = False
GlAkt = False

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KButt " & Err.Number
Resume Next

End Sub
Public Function KCoSu(LiCon As VB.Control, ByVal SuStr As String, Optional ByVal StaZa As Long = 0, Optional ByVal Exakt As Boolean = False) As Long
On Error GoTo LaErr
'Durchsuchen einer Combobox oder einer Listbox
 
Dim RetWe As Long
Dim AktZa As Long
Dim RetZa As Long
Dim Lange As Long
Dim wMsg As Long

' Kann die schnelle API-Funktion verwendet werden?
' Voraussetzung:
' exakte Suche, oder Suche nach Teilbegriff von links
' UND Groß-/Kleinschreibung ignorieren

If Left$(SuStr, 1) <> "*" And Exakt = True Then
    ' Voraussetzungen erfüllt!
  
    If TypeOf LiCon Is ListBox Then
        wMsg = IIf(Right$(SuStr, 1) = "*", LB_FINDSTRING, LB_FINDSTRINGEXACT)
    Else
        wMsg = IIf(Right$(SuStr, 1) = "*", CB_FINDSTRING, CB_FINDSTRINGEXACT)
    End If

    If Right$(SuStr, 1) = "*" Then SuStr = Left$(SuStr, Len(SuStr) - 1)

    If StaZa = 0 Then StaZa = -1
  
    RetWe = SendMessage(LiCon.hwnd, wMsg, StaZa, SuStr)
  
    If RetWe < StaZa Then RetWe = -1

ElseIf Left$(SuStr, 1) <> "*" And Right$(SuStr, 1) <> "*" And Exakt = False Then
  ' exakte Suche unter Berücksichtigung der Groß-/Kleinschreibung
  
    RetWe = -1
    With LiCon
        RetZa = .ListCount - 1
        For AktZa = StaZa To RetZa
            If .List(AktZa) = SuStr Then
                RetWe = AktZa: Exit For
            End If
        Next AktZa
    End With

ElseIf Left$(SuStr, 1) <> "*" And Right$(SuStr, 1) = "*" And Exakt = False Then
    ' Suche beginnend von links unter Berücksichtigung der Groß-/Kleinschreibung
  
    SuStr = Left$(SuStr, Len(SuStr) - 1)
    Lange = Len(SuStr)
    RetWe = -1
    With LiCon
        RetZa = .ListCount - 1
        For AktZa = StaZa To RetZa
            If Left$(.List(AktZa), Lange) = SuStr Then
                RetWe = AktZa: Exit For
            End If
        Next AktZa
    End With

ElseIf Left$(SuStr, 1) = "*" And Right$(SuStr, 1) <> "*" And Exakt = False Then
  ' Suche beginnend von rechts unter Berücksichtigung der Groß-/Kleinschreibung
  
    SuStr = Mid$(SuStr, 2)
    Lange = Len(SuStr)
    RetWe = -1
    With LiCon
        RetZa = .ListCount - 1
        For AktZa = StaZa To RetZa
            If Right$(.List(AktZa), Lange) = SuStr Then
                RetWe = AktZa: Exit For
            End If
        Next AktZa
    End With

Else
  ' Globale Suche oder Suche von rechts beginnend, Groß-/Kleinschreibung ignorieren
  
    RetWe = -1
    With LiCon
        RetZa = .ListCount - 1
        For AktZa = StaZa To RetZa
            If .List(AktZa) Like SuStr Then
                RetWe = AktZa: Exit For
            End If
        Next AktZa
    End With
    
End If

KCoSu = RetWe

Exit Function

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KCoSu " & Err.Number
Resume Next

End Function
Public Sub KDatei(ByVal FiFnk As Integer)
On Error GoTo AnErr
'FileView Datei Operationen

Dim NoNam As String
Dim FiNam As String
Dim DaExt As String
Dim DaNam As String
Dim DaPfa As String
Dim FiStr As String
Dim AusZa As Integer
Dim AktZa As Integer
Dim GesZa As Integer
Dim SelZa As Integer
Dim Frage As Integer
Dim LiIdx As Integer
Dim RetWe As Boolean
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmDaS As XtremeCommandBars.CommandBarComboBox
Dim CmDaA As XtremeCommandBars.CommandBarComboBox
Dim CmDat As XtremeCommandBars.CommandBarComboBox
Dim CmBiS As XtremeCommandBars.CommandBarComboBox
Dim CmBiA As XtremeCommandBars.CommandBarComboBox
Dim CmbIt As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01
Set LiFld = FM.fldView1
Set LiFi1 = FM.filView1
Set LiFi2 = FM.filView2

Set clFil = New clsFile

Set CmDaS = CmBrs.FindControl(CmDaS, SY_EX_Datei_Sor, , True)
Set CmDaA = CmBrs.FindControl(CmDaA, SY_EX_Datei_Ans, , True)
Set CmDat = CmBrs.FindControl(CmDat, SY_EX_Datei_Thm, , True)

Set CmBiS = CmBrs.FindControl(CmBiS, SY_BI_Bild_Sorter, , True)
Set CmBiA = CmBrs.FindControl(CmBiA, SY_BI_Bild_Ansicht, , True)
Set CmbIt = CmBrs.FindControl(CmbIt, SY_BI_Bild_Thumbna, , True)

Select Case FiFnk
Case 2:
    Set LiNod = LiFld.SelectedNode
    NoNam = LiNod.DisplayName
    LiNod.CreateNewFolder "Neuer Ordner", True
Case 3:
    Set LiNod = LiFld.SelectedNode
    NoNam = LiNod.DisplayName
    LiNod.BeginLabelEdit
Case 4:
    Set LiNod = LiFld.SelectedNode
    NoNam = LiNod.DisplayName
    Select Case NoNam
    Case "Praxisdaten": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Backup": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Dokumente": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Bilder": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Vorlagen": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Formulare": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case Else: LiNod.ExecuteShellCommand cmdDelete
    End Select
    LiNod.Delete
Case 5:
    Set LiNod = LiFld.SelectedNode
    NoNam = LiNod.DisplayName
    Select Case NoNam
    Case "Praxisdaten": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Backup": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Dokumente": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Bilder": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Vorlagen": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Formulare": SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case Else: LiNod.ExecuteShellCommand cmdProperties
    End Select
    LiNod.RefreshTree
Case 6:
    Set clFil = New clsFile
    GesZa = LiFi2.ItemCount
    If GesZa > 0 Then
        SelZa = LiFi2.SelectedCount
        If SelZa > 0 Then
            If SelZa > 1 Then
                Tit1 = "Dateien Entfernen"
                Mld1 = "Möchten Sie die " & SelZa & " Dateien wirklich löschen?"
            Else
                Tit1 = "Datei Entfernen"
                Mld1 = "Möchten Sie die Datei wirklich löschen?"
            End If
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                For AktZa = 1 To GesZa
                    Set LiFit = LiFi2.ListItem(AktZa)
                    If LiFit.Selected = True Then
                        FiNam = LiFit.DisplayName
                        If LiFit.Attributes(Folder) And Folder Then
                            Select Case FiNam
                            Case "Praxisdaten": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "Backup": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "Dokumente": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "Bilder": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "Vorlagen": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "Formulare": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case Else:
                                    With clFil
                                        .DaLoe = LiFit.Path & vbNullChar
                                        .FilLoe
                                    End With
                            End Select
                        Else
                            Select Case LCase(Right$(FiNam, 3))
                            Case "ini": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "mdb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "ldb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "dbx": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "dbv": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case "dax": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
                            Case Else:
                                With clFil
                                    .DaLoe = LiFit.Path & vbNullChar
                                    .FilLoe
                                End With
                            End Select
                        End If
                    End If
                Next AktZa
            End If
        End If
    End If
    LiFi2.RefreshViewFast
    Set clFil = Nothing
Case 7:
    If LiFi2.SelectedCount > 0 Then
        Set LiFit = LiFi2.FirstSelectedItem
        FiNam = LiFit.DisplayName
        If LiFit.Attributes(Folder) And Folder Then
            Select Case FiNam
            Case "Praxisdaten": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "Backup": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "Dokumente": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "Bilder": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "Vorlagen": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "Formulare": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case Else: LiFit.ExecuteShellCommand cmdRename
            End Select
        Else
            Select Case LCase(Right$(FiNam, 3))
            Case "ini": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "mdb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "ldb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "dbx": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "dbv": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case "dax": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht umbenannt werden", IC48_Forbidden
            Case Else: LiFit.ExecuteShellCommand cmdRename
            End Select
            LiFi2.RefreshViewFast
        End If
    End If
Case 8:
    GesZa = LiFi2.ItemCount
    If GesZa > 0 Then
        If LiFi2.SelectedCount > 0 Then
            For AktZa = 1 To GesZa
                Set LiFit = LiFi2.ListItem(AktZa)
                If LiFit.Selected = True Then
                    FiNam = LiFit.DisplayName
                    If LiFit.Attributes(Folder) And Folder Then
                        Select Case FiNam
                        Case "Praxisdaten": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "Backup": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "Dokumente": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "Bilder": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "Vorlagen": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "Formulare": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case Else: LiFit.ExecuteShellCommand cmdCut
                        End Select
                    Else
                        Select Case LCase(Right$(FiNam, 3))
                        Case "ini": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "mdb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "ldb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "dbx": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "dbv": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case "dax": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht ausgeschnitten werden", IC48_Forbidden
                        Case Else: LiFit.ExecuteShellCommand cmdCut
                        End Select
                    End If
                End If
            Next AktZa
        End If
    End If
    LiFi2.RefreshViewFast
Case 9:
    GesZa = LiFi2.ItemCount
    If GesZa > 0 Then
        If LiFi2.SelectedCount > 0 Then
            For AktZa = 1 To GesZa
                Set LiFit = LiFi2.ListItem(AktZa)
                If LiFit.Selected = True Then
                    FiNam = LiFit.DisplayName
                    If LiFit.Attributes(Folder) And Folder Then
                        LiFit.ExecuteShellCommand cmdCopy
                    Else
                        LiFit.ExecuteShellCommand cmdCopy
                    End If
                End If
            Next AktZa
        End If
    End If
    LiFi2.RefreshView
Case 10:
    Set LiNod = LiFld.SelectedNode
    LiNod.ExecuteShellCommand cmdPaste
    LiNod.RefreshTree
    LiFi2.RefreshView
Case 11:
    If LiFi2.SelectedCount > 0 Then
        Set LiFit = LiFi2.FirstSelectedItem
        DaNam = LiFit.DisplayName
        FiNam = LiFit.Path
        If Not LiFit.Attributes(Folder) And Folder Then
            Select Case LCase(Right$(FiNam, 3))
            Case "ini": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "mdb": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "ldb": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "dbx": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "dbv": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "dax": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "blg": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "crd": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "lst": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "lsv": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "lsp": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case "bat": SPopu "Systemdatei", "Die Datei: " & DaNam & " darf nicht geöffnet werden", IC48_Forbidden
            Case Else:
                If GlRDP = True Then
                    Set clFil = New clsFile
                    With clFil
                        .FilPfa FiNam
                        DaPfa = .DaPfa
                        DaExt = .DaExt
                    End With
                    Select Case LCase(DaExt)
                    Case "pdf": SImage FiNam
                    Case "jpg": SImage FiNam
                    Case "png": SImage FiNam
                    Case "bmp": SImage FiNam
                    Case "tif": SImage FiNam
                    Case "gif": SImage FiNam
                    Case "wmf": SImage FiNam
                    Case "emf": SImage FiNam
                    Case "jpeg": SImage FiNam
                    Case "tiff": SImage FiNam
                    Case "doc": VoTxMa FiNam, DaNam, 9
                    Case "dot": VoTxMa FiNam, DaNam, 9
                    Case "rtf": VoTxMa FiNam, DaNam, 5
                    Case "txt": VoTxMa FiNam, DaNam, 1
                    Case "csv": VoTxMa FiNam, DaNam, 1
                    Case "docx": VoTxMa FiNam, DaNam, 13
                    Case Else: SPopu "Ungültiger Dateityp", "Dieser Dateityp darf nicht geöffnet werden", IC48_Warning
                    End Select
                    Set clFil = Nothing
                Else
                    LiFit.ExecuteShellCommand cmdDefault
                End If
            End Select
            LiFi2.RefreshViewFast
        End If
    End If
    
Case 12:
    If LiFi2.SelectedCount > 0 Then
        Set LiFit = LiFi2.FirstSelectedItem
        LiFit.ExecuteShellCommand cmdProperties
    End If
Case 13:
    LiFld.RefreshTree
    LiFi2.RefreshView
Case 14:
    Tit1 = "Textdatei hinzufügen"
    Mld1 = "Möchten Sie jetzt eine neue Textdatei hinzufügen"
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
        Set clFil = New clsFile
        Set LiNod = LiFld.SelectedNode
        FiNam = LiNod.Path & "\Neues Dokument.txt"
        With clFil
            If .FilVor(FiNam) = True Then
                .DaLoe = FiNam & vbNullChar
                .FilLoe
            End If
            .FilPfa FiNam
            .StrDa = vbNullString
            .FilWrSt
        End With
        Set clFil = Nothing
        DoEvents
        LiFi2.RefreshView
        DoEvents
        GesZa = LiFi2.ItemCount
        If GesZa > 0 Then
            For AktZa = 1 To GesZa
                Set LiFit = LiFi2.ListItem(AktZa)
                FiNam = LiFit.DisplayName
                If FiNam = "Neues Dokument.txt" Then
                    LiFit.Selected = True
                    LiFit.ExecuteShellCommand cmdRename
                    Exit For
                End If
            Next AktZa
        End If
    End If
Case 15:
    If clFil.FilDir(GlDpf) = True Then
        LiFi2.CurrentFolder = GlDpf
        Set LiNod = LiFld.SelectedNode
        LiNod.Expanded = True
    End If
Case 16:
    FiNam = IniGetVal("System", "EmalPf")
    If clFil.FilDir(FiNam) = True Then
        LiFi2.CurrentFolder = FiNam
        Set LiNod = LiFld.SelectedNode
        LiNod.Expanded = True
    End If
Case 17:
    FiNam = IniGetVal("System", "ImpPfa")
    If clFil.FilDir(FiNam) = True Then
        LiFi2.CurrentFolder = FiNam
        Set LiNod = LiFld.SelectedNode
        LiNod.Expanded = True
    End If
Case 18:
    FiNam = IniGetVal("System", "ExpPfa")
    If clFil.FilDir(FiNam) = True Then
        LiFi2.CurrentFolder = FiNam
        Set LiNod = LiFld.SelectedNode
        LiNod.Expanded = True
    End If
Case 19:
    FiNam = IniGetVal("System", "DockPf")
    If clFil.FilDir(FiNam) = True Then
        LiFi2.CurrentFolder = FiNam
        Set LiNod = LiFld.SelectedNode
        LiNod.Expanded = True
    End If
Case 20:
    FiNam = GlFrO
    If clFil.FilDir(FiNam) = True Then
        LiFi2.CurrentFolder = FiNam
        Set LiNod = LiFld.SelectedNode
        LiNod.Expanded = True
    End If
Case 21:
    LiIdx = CmDaA.ListIndex - 1
    LiFi2.ViewStyle = LiIdx
    IniSetVal "Layout", "ViwSt2", LiIdx
Case 22:
    LiIdx = CmBiA.ListIndex - 1
    LiFi1.ViewStyle = LiIdx
    IniSetVal "Layout", "ViwSt1", LiIdx
Case 23:
    LiIdx = CmDaS.ListIndex - 1
    RetWe = LiFi2.SortByColumn(vbNullString, LiIdx, Ascending)
    IniSetVal "Layout", "SorSp2", LiIdx
Case 24:
    LiIdx = CmBiS.ListIndex - 1
    RetWe = LiFi1.SortByColumn(vbNullString, LiIdx, Ascending)
    IniSetVal "Layout", "SorSp1", LiIdx
Case 25:
    LiIdx = CmDat.ListIndex
    Select Case LiIdx
    Case 1: LiFi2.SetThumbnailSize 120, 80
    Case 2: LiFi2.SetThumbnailSize 120, 120
    Case 3: LiFi2.SetThumbnailSize 80, 120
    Case 4: LiFi2.SetThumbnailSize 140, 100
    Case 5: LiFi2.SetThumbnailSize 140, 140
    Case 6: LiFi2.SetThumbnailSize 100, 140
    Case 7: LiFi2.SetThumbnailSize 180, 140
    Case 8: LiFi2.SetThumbnailSize 180, 180
    Case 9: LiFi2.SetThumbnailSize 140, 180
    Case 10: LiFi2.SetThumbnailSize 200, 160
    Case 11: LiFi2.SetThumbnailSize 200, 200
    Case 12: LiFi2.SetThumbnailSize 160, 200
    End Select
    IniSetVal "Layout", "ThmGr2", LiIdx
Case 26:
    LiIdx = CmbIt.ListIndex
    Select Case LiIdx
    Case 1: LiFi1.SetThumbnailSize 120, 80
    Case 2: LiFi1.SetThumbnailSize 120, 120
    Case 3: LiFi1.SetThumbnailSize 80, 120
    Case 4: LiFi1.SetThumbnailSize 140, 100
    Case 5: LiFi1.SetThumbnailSize 140, 140
    Case 6: LiFi1.SetThumbnailSize 100, 140
    Case 7: LiFi1.SetThumbnailSize 180, 140
    Case 8: LiFi1.SetThumbnailSize 180, 180
    Case 9: LiFi1.SetThumbnailSize 140, 180
    Case 10: LiFi1.SetThumbnailSize 200, 160
    Case 11: LiFi1.SetThumbnailSize 200, 200
    Case 12: LiFi1.SetThumbnailSize 160, 200
    End Select
    IniSetVal "Layout", "ThmGr1", LiIdx
Case 27:
    If LiFi2.SelectedCount > 0 Then
        Set LiFit = LiFi2.FirstSelectedItem
        FiNam = LiFit.GetDisplayNameEx(FIVForParsing)
        If LiFit.Attributes(Folder) And Folder Then
            SPopu "Windows Ordner", "Ein Windows Ordner kann nicht importiert werden", IC48_Forbidden
        Else
            Select Case LCase(Right$(FiNam, 3))
            Case "ini": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "mdb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "ldb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "dbx": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "dbv": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "dax": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "blg": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "crd": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "lst": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "lsv": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case "lsp": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht importiert werden", IC48_Forbidden
            Case Else:
                frmAdrSuch.FiNam = FiNam
                frmAdrSuch.Show vbModal
            End Select
        End If
    End If
Case 28:
    FiStr = InputBox("Bitte Filterkriterien eingeben:", "Filterkriterien", "*.*")
    LiFi2.FilePattern = FiStr
    LiFi2.RefreshView
Case 29:
    GesZa = LiFi2.ItemCount
    If GesZa > 0 Then
        If LiFi2.SelectedCount > 0 Then
            For AktZa = 1 To GesZa
                Set LiFit = LiFi2.ListItem(AktZa)
                If LiFit.Selected = True Then
                    DaNam = LiFit.DisplayName
                    FiNam = LiFit.Path
                    SMaNe 0, , , DaNam, DaNam, FiNam
                End If
            Next AktZa
        End If
    End If
End Select

Set clFil = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KDatei " & Err.Number
Resume Next

End Sub
Public Sub KFrKo(Optional ByVal EmSen As Boolean = False)
On Error GoTo AnErr
'Versendet einen Fragebogen oder das Neuaufnahmeformular

Dim BogNr As Long
Dim KmStr As String
Dim EmTex As String
Dim EmBet As String
Dim AktZa As Integer

BogNr = Mid$(GlNod, 2, Len(GlNod) - 1)

If GlBoV > 0 Then 'Fragebogen vorhanden
    For AktZa = 1 To GlBoV
        If GlFrB(AktZa, 0) = BogNr Then
            If GlFrB(AktZa, 3) <> vbNullString Then
                KmStr = GlFrB(AktZa, 3)
            End If
            Exit For
        End If
    Next AktZa
End If

If BogNr = 0 Then
    KmStr = GlNaf
End If

If KmStr <> vbNullString Then
    Clipboard.Clear
    Clipboard.SetText KmStr
    If EmSen = False Then
        SPopu "Fragebogenpublizierung", "Die URL des Fragebogens wurde in die Zwischenablage kopiert.", IC48_Information
    End If
    
    If EmSen = True Then
        EmBet = "Fragebogenlink"
        EmTex = vbCrLf & KmStr
        SMaNe 0, , , EmTex, EmBet
    End If
Else
    SPopu "Fragebogenpublizierung", "Dieser Fragebogen wurde noch nicht publiziert!", IC48_Forbidden
End If

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KDatei " & Err.Number
Resume Next

End Sub
Public Sub KGeKa(Optional ByVal NeKat As Long)
On Error GoTo PeErr
'Gebührenkatalog aktualisieren

Dim KatNr As Long
Dim AktZa As Integer
Dim KaIdx As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox

Set FM = frmKatGE
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set TxDe7 = frmMain.txtDeta7

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

If NeKat > 0 Then
    KatNr = NeKat
Else
    If TxDe7.Text <> vbNullString Then
        KatNr = TxDe7.Text
    Else
        If GlStK <= UBound(GlGKa) Then
            KatNr = GlStK
        Else
            KatNr = GlGKa(1, 0)
        End If
    End If
End If

For AktZa = 1 To UBound(GlGKa)
    If GlGKa(AktZa, 0) = KatNr Then
        KaIdx = AktZa
        Exit For
    End If
Next AktZa

CmSu1.ListIndex = KaIdx
CmSu2.ListIndex = KaIdx
DoEvents

Select Case RbTab.id
Case RibTab_Kat_EinGeb: P_List "GbEi", KatNr, 1, GlFGE
Case RibTab_Kat_KetGeb: P_List "GbEi", KatNr, 2
End Select

Set CmAcs = Nothing
Set CmSu1 = Nothing
Set CmSu2 = Nothing
Set CmBrs = Nothing

Exit Sub

PeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KGeKa " & Err.Number
Resume Next

End Sub
Public Sub KGrKa(ByVal KaFor As String)
On Error GoTo OpErr
'Stellt bestimmte Formatierungen ein

Dim GrCap As String
Dim KetMe As Boolean
Dim PopKa As Boolean
Dim GrAnz As Integer
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

GlAkt = True

For AktZa = 1 To 17
    Select Case AktZa
    Case 1: Set FM = frmKatBE 'Begründungen
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 2: Set FM = frmKatDE 'Diagnosen
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 3: Set FM = frmKatGE 'Gebühren
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 4: Set FM = frmKatLE 'Laqborleistungen
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 5: Set FM = frmKatME 'Arzneimittel
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 6: Set FM = frmKatTE 'Termine
            Set CmBrs = FM.comBar02
            Set RpCon = FM.repCont7
            PopKa = False
    Case 7: Set FM = frmKatRE 'Rechnungen
            Set CmBrs = FM.comBar02
            Set RpCon = FM.repCont7
            PopKa = False
    Case 8: Set FM = frmKatPE 'Laborparameter
            Set CmBrs = FM.comBar02
            Set RpCon = FM.repCont7
            PopKa = False
    Case 9: Set FM = frmKatKD 'Krankenblattdiagnosen
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 10: Set FM = frmKatAE 'Anamnesen
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 11: Set FM = frmKatKM 'Krankenblattmedikamente
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 12: Set FM = frmKatBU 'Buchhaltung
            Set CmBrs = FM.comBar02
            Set RpCon = FM.repCont7
            PopKa = False
    Case 13: Set FM = frmKatRC
            Set CmBrs = FM.comBar02
            Set RpCon = FM.repCont7
            PopKa = False
    Case 14: Set FM = frmKatBA 'Kontoumsätze
            Set CmBrs = FM.comBar02
            Set RpCon = FM.repCont7
            PopKa = False
    Case 15: Set FM = frmKatAR 'Artikelliste
            Set CmBrs = FM.comBar02
            Set MoKal = FM.dtpDatu1
            Set RpCon = FM.repCont7
            PopKa = True
    Case 16: Set FM = frmKatTX 'Textverarbeitung
            'Set CmBrs = FM.comBar02
            'Set RpCon = FM.repCont7
            'PopKa = False
    Case 17: Set FM = frmKatBV 'Banking
            Set CmBrs = FM.comBar02
            Set RpCon = FM.repCont7
            PopKa = False
    End Select

    Set CmAcs = CmBrs.Actions
    Set RpCls = RpCon.Columns

    With RpCon
        Select Case KaFor
        Case "PopKal":
            If PopKa = True Then
                Set CmBar = CmBrs.Item(2)
                If CmAcs(KM_Popupkalender).Checked = False Then
                    CmAcs(KM_Popupkalender).Checked = True
                    CmAcs(KM_Multimarker).Checked = False
                    CmAcs(KM_Multimarker).Enabled = False
                    IniSetVal "Layout", "PopKal", True
                    CmBar.Visible = True
                    With MoKal
                        .AllowNoncontinuousSelection = False
                        If GlSty = 8 Then 'Office 2013
                            .BorderStyle = xtpDatePickerBorderStatic
                        ElseIf GlSty = 7 Then 'Office 2013
                            .BorderStyle = xtpDatePickerBorderStatic
                        Else
                            .BorderStyle = xtpDatePickerBorderOffice
                        End If
                        .MaxSelectionCount = 1
                        .ShowNoneButton = True
                        .ShowTodayButton = True
                        .ToolTipText = "Markieren Sie bitte hier den Behandlungstag des Patienten"
                        .Visible = False
                    End With
                    GlPoK = True 'Popupkalender
                Else
                    CmAcs(KM_Popupkalender).Checked = False
                    CmAcs(KM_Multimarker).Enabled = True
                    IniSetVal "Layout", "PopKal", False
                    CmBar.Visible = False
                    With MoKal
                        .AllowNoncontinuousSelection = True
                        .BorderStyle = xtpDatePickerBorderNone
                        .MaxSelectionCount = GlMxK 'Maximal slektierbare Kalendertage
                        .ShowNoneButton = False
                        .ShowTodayButton = False
                        .ToolTipText = "Markieren Sie bitte hier die Behandlungstage des Patienten"
                        .Visible = True
                    End With
                    GlPoK = False 'Popupkalender
                End If
            End If
        Case "MulMar":
            Select Case AktZa
            Case 1: GlM01 = Not GlM01
                    MoKal.MultiSelectionMode = GlM01
                    CmAcs(KM_Multimarker).Checked = GlM01
                    IniSetVal "Layout", "MulM01", GlM01
            Case 2: GlM02 = Not GlM02
                    MoKal.MultiSelectionMode = GlM02
                    CmAcs(KM_Multimarker).Checked = GlM02
                    IniSetVal "Layout", "MulM02", GlM02
            Case 3: GlM03 = Not GlM03
                    MoKal.MultiSelectionMode = GlM03
                    CmAcs(KM_Multimarker).Checked = GlM03
                    IniSetVal "Layout", "MulM03", GlM03
            Case 4: GlM04 = Not GlM04
                    MoKal.MultiSelectionMode = GlM04
                    CmAcs(KM_Multimarker).Checked = GlM04
                    IniSetVal "Layout", "MulM04", GlM04
            Case 5: GlM05 = Not GlM05
                    MoKal.MultiSelectionMode = GlM05
                    CmAcs(KM_Multimarker).Checked = GlM05
                    IniSetVal "Layout", "MulM05", GlM05
            Case 6: GlM06 = Not GlM06
                    MoKal.MultiSelectionMode = GlM06
                    CmAcs(KM_Multimarker).Checked = GlM06
                    IniSetVal "Layout", "MulM06", GlM06
            Case 7: GlM07 = Not GlM07
                    CmAcs(KM_Multimarker).Checked = GlM07
                    IniSetVal "Layout", "MulM07", GlM07
            Case 8: GlM08 = Not GlM08
                    MoKal.MultiSelectionMode = GlM08
                    CmAcs(KM_Multimarker).Checked = GlM08
                    IniSetVal "Layout", "MulM08", GlM08
            Case 9: GlM09 = Not GlM09
                    MoKal.MultiSelectionMode = GlM09
                    CmAcs(KM_Multimarker).Checked = GlM09
                    IniSetVal "Layout", "MulM09", GlM09
            Case 10: GlM10 = Not GlM10
                    MoKal.MultiSelectionMode = GlM10
                    CmAcs(KM_Multimarker).Checked = GlM10
                    IniSetVal "Layout", "MulM10", GlM10
            Case 11: GlM11 = Not GlM11
                    CmAcs(KM_Multimarker).Checked = GlM11
                    IniSetVal "Layout", "MulM11", GlM11
            Case 12: GlM12 = Not GlM12
                    MoKal.MultiSelectionMode = GlM12
                    CmAcs(KM_Multimarker).Checked = GlM12
                    IniSetVal "Layout", "MulM12", GlM12
            End Select
        Case "GrdZei":
            If CmAcs(KM_Zeilenumbruch).Checked = False Then
                .PaintManager.FixedRowHeight = False
                IniSetVal "Katalog", "KatZei", -1
                CmAcs(KM_Zeilenumbruch).Checked = True
                Set RpCol = RpCls.Find(Kat_IDKurz)
                RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            Else
                .PaintManager.FixedRowHeight = True
                IniSetVal "Katalog", "KatZei", 0
                CmAcs(KM_Zeilenumbruch).Checked = False
            End If
        Case "GrdMkr":
            If CmAcs(KM_Zeilenmarker).Checked = True Then
                IniSetVal "Katalog", "KatMkr", 0
                CmAcs(KM_Zeilenmarker).Checked = False
                .PaintManager.UseAlternativeBackground = False
                GlKaZ = False
            Else
                IniSetVal "Katalog", "KatMkr", -1
                CmAcs(KM_Zeilenmarker).Checked = True
                .PaintManager.UseAlternativeBackground = True
                GlKaZ = True
            End If
        Case "GrdGrl":
            If CmAcs(KM_Gitternetz).Checked = True Then
                .PaintManager.HorizontalGridStyle = xtpGridNoLines
                .PaintManager.VerticalGridStyle = xtpGridNoLines
                IniSetVal "Katalog", "KatGrl", 0
                CmAcs(KM_Gitternetz).Checked = False
            Else
                .PaintManager.HorizontalGridStyle = xtpGridSolid
                .PaintManager.VerticalGridStyle = xtpGridSolid
                IniSetVal "Katalog", "KatGrl", -1
                CmAcs(KM_Gitternetz).Checked = True
            End If
        Case "GrdPrv":
            If AktZa = 4 Then
                If CmAcs(KM_Vorschauzeile).Checked = True Then
                    .PreviewMode = False
                    IniSetVal "Katalog", "KatVor", 0
                    CmAcs(KM_Vorschauzeile).Checked = False
                Else
                    .PreviewMode = True
                    IniSetVal "Katalog", "KatVor", -1
                    CmAcs(KM_Vorschauzeile).Checked = True
                End If
            End If
        Case "GrdGrp":
                Select Case AktZa
                Case 2:
                    If GlGrD = False Then
                        .SortOrder.Add .Columns(Kat_Gruppe)
                        .GroupsOrder.Add .Columns(Kat_Gruppe)
                        .GroupsOrder(0).SortAscending = True
                        GlGrD = True
                    Else
                        .SortOrder.DeleteAll
                        .GroupsOrder.DeleteAll
                        GlGrD = False
                    End If
                    .Populate
                    Set RpRws = .Rows
                    For Each RpRow In RpRws
                        If RpRow.GroupRow = True Then
                            Set RpGrw = RpRow
                            Set ChRws = RpGrw.Childs
                            GrAnz = ChRws.Count
                            If Len(RpGrw.GroupCaption) > 0 Then
                                GrCap = Right$(RpGrw.GroupCaption, Len(RpGrw.GroupCaption) - 8)
                                RpGrw.GroupCaption = GrCap & " (" & GrAnz & ")"
                            End If
                        End If
                    Next RpRow
                    CmAcs(KM_Gruppierung).Checked = GlGrD
                    IniSetVal "Katalog", "KaDeGr", GlGrD
                Case 3:
                    If GlKaG = False Then
                        .SortOrder.Add .Columns(Kat_Gruppe)
                        .GroupsOrder.Add .Columns(Kat_Gruppe)
                        .GroupsOrder(0).SortAscending = True
                        GlKaG = True
                    Else
                        .SortOrder.DeleteAll
                        .GroupsOrder.DeleteAll
                        GlKaG = False
                    End If
                    .Populate
                    Set RpRws = .Rows
                    For Each RpRow In RpRws
                        If RpRow.GroupRow = True Then
                            Set RpGrw = RpRow
                            Set ChRws = RpGrw.Childs
                            GrAnz = ChRws.Count
                            If Len(RpGrw.GroupCaption) > 0 Then
                                GrCap = Right$(RpGrw.GroupCaption, Len(RpGrw.GroupCaption) - 8)
                                RpGrw.GroupCaption = GrCap & " (" & GrAnz & ")"
                            End If
                        End If
                    Next RpRow
                    CmAcs(KM_Gruppierung).Checked = GlKaG
                    IniSetVal "Katalog", "KaGeGr", GlKaG
                Case 9:
                    If GlGrD = True Then
                        .SortOrder.Add .Columns(Kat_Gruppe)
                        .GroupsOrder.Add .Columns(Kat_Gruppe)
                        .GroupsOrder(0).SortAscending = True
                    Else
                        .SortOrder.DeleteAll
                        .GroupsOrder.DeleteAll
                    End If
                    .Populate
                    Set RpRws = .Rows
                    For Each RpRow In RpRws
                        If RpRow.GroupRow = True Then
                            Set RpGrw = RpRow
                            Set ChRws = RpGrw.Childs
                            GrAnz = ChRws.Count
                            If Len(RpGrw.GroupCaption) > 0 Then
                                GrCap = Right$(RpGrw.GroupCaption, Len(RpGrw.GroupCaption) - 8)
                                RpGrw.GroupCaption = GrCap & " (" & GrAnz & ")"
                            End If
                        End If
                    Next RpRow
                    CmAcs(KM_Gruppierung).Checked = GlGrD
                End Select
        End Select
        DoEvents
        .Redraw
    End With
Next AktZa

GlAkt = False

Set CmAct = Nothing
Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KGrKa " & Err.Number
Resume Next

End Sub
Public Sub KGrLa()
On Error GoTo OpErr
'Stellt bestimmte Formatierungen ein

Dim GrPre As Single
Dim GrCap As String
Dim KetMe As Boolean
Dim GrVor As Boolean
Dim GrAnz As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set RpCls = RpCo8.Columns

GlAkt = True

Screen.MousePointer = vbHourglass
DoEvents

Select Case Left$(GlNod, 1)
Case "A": GrVor = True
Case "C": GrVor = True
Case "D": KetMe = True
Case "F": KetMe = True
Case "G": GrVor = True
Case "H": KetMe = True
Case "I":
Case "J": KetMe = True
Case "K":
Case "L":
Case "M":
Case "N": GrVor = True
Case "O":
Case "P":
Case "Q": KetMe = True
End Select

With RpCo8
    If KetMe = False Then
        If GrVor = True Then
            If GlGrG = False Then 'Gruppierung Kataloge
                .SortOrder.Add .Columns(Kat_Gruppe)
                .GroupsOrder.Add .Columns(Kat_Gruppe)
                .GroupsOrder(0).SortAscending = False
                GlGrG = True
            Else
                .SortOrder.DeleteAll
                .GroupsOrder.DeleteAll
                GlGrG = False
            End If
            DoEvents
            KList
        End If
    End If
End With

CmAcs(SY_EI_Gruppierung).Checked = GlGrG
IniSetVal "Layout", "GruEin", GlGrG

DoEvents
Screen.MousePointer = vbNormal

Set CmAct = Nothing
Set CmBrs = Nothing
Set RpCo8 = Nothing

GlAkt = False

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KGrLa " & Err.Number
Resume Next

End Sub
Public Sub KGrNe()
On Error GoTo ReErr
'Legt eine neue Gruppe an

Dim KatNr As Long
Dim NeNam As String
Dim TreKy As String

Set FM = frmMain
Set TrLi2 = FM.trvList2

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

TreKy = Left$(GlNod, 1)

Select Case GlBut
Case RibTab_Kat_Eintrg:
    Select Case TreKy
    Case "A": NeNam = InputBox("Bitte Namen des neuen Gebührenkataloges eingeben:", "Neuer Gebührenkatalog", "Neuer Gebührenkatalog") 'Gebührenkataloge
    Case "C": NeNam = InputBox("Bitte Namen des neuen Diagnosekataloges eingeben:", "Neuer Diagnosekatalog", "Neuer Diagnosekatalog") 'Diagnosekataloge
    Case "G": NeNam = InputBox("Bitte Namen des neuen Laborkataloges eingeben:", "Neuer Laborkatalog", "Neuer Laborkatalog") 'Laborkataloge
    Case "I": NeNam = InputBox("Bitte Namen des neuen Arzneikataloges eingeben:", "Neuer Arzneikatalog", "Neuer Arzneikatalog") 'Arzneikataloge
    Case "K": Exit Sub 'Begründungskatalog
    Case "L": NeNam = InputBox("Bitte Namen des neuen Anamnesekataloges eingeben:", "Neuer Anamnesekatalog", "Neuer Anamnesekatalog") 'Anamnesekataloge
    Case "M": Exit Sub 'Terminbetreffs
    Case "N": NeNam = InputBox("Bitte Namen des neuen Fragebogens eingeben:", "Neuer Fragebogen", "Neuer Fragebogen") 'Fragebogenkataloge
    Case "O": Exit Sub 'Textphrasenkataloge
    Case "P": NeNam = InputBox("Bitte Namen des neuen Artikelkataloges eingeben:", "Neuer Artikelkatalog", "Neuer Artikelkatalog") 'Artikelkataloge
    End Select
Case RibTab_Kat_Ketten:
    Select Case TreKy
    Case "D": NeNam = InputBox("Bitte Namen des neuen Gebührenkataloges eingeben:", "Neuer Gebührenkatalog", "Neuer Gebührenkatalog") 'Gebührenketten
    Case "F": NeNam = InputBox("Bitte Namen des neuen Diagnosekataloges eingeben:", "Neuer Diagnosekatalog", "Neuer Diagnosekatalog") 'Diagnoseketten
    Case "H": NeNam = InputBox("Bitte Namen des neuen Laborkataloges eingeben:", "Neuer Laborkatalog", "Neuer Laborkatalog") 'Laborketten
    Case "J": NeNam = InputBox("Bitte Namen des neuen Arzneikataloges eingeben:", "Neuer Arzneikatalog", "Neuer Arzneikatalog") 'Arzneiketten
    Case "R": Exit Sub 'Terminketten
    Case "Q": NeNam = InputBox("Bitte Namen des neuen Artikelkataloges eingeben:", "Neuer Artikelkatalog", "Neuer Artikelkatalog") 'Artikelketten
    End Select
Case RibTab_Kat_Frage:
    NeNam = InputBox("Bitte Namen des neuen Fragebogens eingeben:", "Neuer Fragebogen", "Neuer Fragebogen") 'Fragebogenkataloge
End Select

If NeNam <> vbNullString Then
    GlAkt = True
    If Len(NeNam) > 40 Then
        NeNam = Left$(NeNam, 40)
    End If
    NeNam = SNaFi(NeNam, False, False, True, True)
    
    Screen.MousePointer = vbHourglass
    DoEvents
    clFen.FenDsk 2

    TrLi2.Nodes.Clear
    KatNr = K_GrNe(NeNam)
    GlNod = TreKy & KatNr
    K_GrLa
    KTree
    K_List
    
    clFen.FenDsk 3
    DoEvents
    Screen.MousePointer = vbNormal

    Set clFen = Nothing
    
    GlAkt = False
End If

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KGrNe " & Err.Number
Resume Next

End Sub
Public Sub KGrUm()
On Error GoTo ReErr
'Benennt eine Gruppe um

Dim NeNam As String
Dim AlNam As String
Dim Mld1, Tit1 As String

Set FM = frmMain
Set TrLi2 = FM.trvList2

If TrLi2.SelectedItem.Text <> vbNullString Then
    If Left$(GlNod, 1) <> "Z" Then
        AlNam = TrLi2.SelectedItem.Text
        NeNam = InputBox("Bitte geben Sie den neuen Namen ein:", "Umbenennen", AlNam)
        If NeNam <> vbNullString Then
            If Len(NeNam) > 40 Then
                NeNam = Left$(NeNam, 40)
            End If
            NeNam = SNaFi(NeNam, False, False, True, False)
            K_GrUm NeNam
            K_GrLa
            TrLi2.Nodes.Clear
            KTree
            K_List
        End If
    Else
        Mld1 = "Diese Gruppe kann nicht umbenannt werden"
        Tit1 = "Umbenennen"
        WindowMess Mld1, Dial2, Tit1, FM.hwnd
    End If
End If

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KGrUm " & Err.Number
Resume Next

End Sub

Public Sub KList()
On Error GoTo LaErr
'Wird ausgeführt, wenn im TreeView geklickt wird

Dim TreKy As String
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmKat As XtremeCommandBars.CommandBarComboBox
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set TrLi2 = FM.trvList2
Set TrLi3 = FM.trvList3
Set RpCo8 = FM.repCont8
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RpCls = RpCo8.Columns

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Select Case GlBut
Case RibTab_Kat_Eintrg: Set CmKat = CmBrs.FindControl(CmKat, KA_Eint_KatCombo, , True)
Case RibTab_Kat_Ketten: Set CmKat = CmBrs.FindControl(CmKat, KA_Kett_KatCombo, , True)
End Select

TreKy = Left$(GlNod, 1)

GlAkt = True

SSuAu 'Hebt die markierten Suchbuchstaben wieder auf
DoEvents

Select Case GlBut
Case RibTab_Kat_Eintrg:
    Select Case TreKy
    Case "A": CmKat.ListIndex = 1 'Gebührenkataloge
    Case "C": CmKat.ListIndex = 2 'Diagnosekataloge
    Case "G": CmKat.ListIndex = 3 'Laborparameter
    Case "I": CmKat.ListIndex = 4 'Arzneikataloge
    Case "K": CmKat.ListIndex = 5 'Begründungskatalog
    Case "L": CmKat.ListIndex = 6 'Anamnesekataloge
    Case "M": CmKat.ListIndex = 7 'Terminbetreffs
    Case "N": CmKat.ListIndex = 8 'Fragebogenkataloge
    Case "O": CmKat.ListIndex = 9 'Textphrasenkataloge
    Case "P": CmKat.ListIndex = 10 'Artikelkatalog
    End Select
Case RibTab_Kat_Ketten:
    Select Case TreKy
    Case "D": CmKat.ListIndex = 1 'Gebührenketten
    Case "F": CmKat.ListIndex = 2 'Diagnoseketten
    Case "H": CmKat.ListIndex = 3 'Laborketten
    Case "J": CmKat.ListIndex = 4 'Arzneiketten
    Case "R": CmKat.ListIndex = 5 'Terminketten
    Case "Q": CmKat.ListIndex = 6 'Artikelketten
    End Select
End Select

If GlBut = RibTab_Tex_Email Then
    For Each Knote In TrLi3.Nodes
        Knote.Image = IC16_Folder_Close
    Next Knote
    TrLi3.SelectedItem.Image = IC16_Folder_Open
    TrLi3.Nodes(1).Image = IC16_Folder_View
Else
    For Each Knote In TrLi2.Nodes
        If Knote.Key <> "N0" Then
            Knote.Image = IC16_Folder_Close
        End If
    Next Knote
    If TrLi2.SelectedItem.Key <> "N0" Then
        TrLi2.SelectedItem.Image = IC16_Folder_Open
    End If
    TrLi2.Nodes(1).Image = IC16_Folder_View
    TrLi2.SelectedItem.Selected = True
End If

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

SSpLaK
DoEvents
K_List

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpCls = Nothing
Set RpCo8 = Nothing

Set clFen = Nothing

GlAkt = False

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KList " & Err.Number
Resume Next

End Sub
Public Sub KMnAE()
On Error GoTo LaErr
'Menue Anmanese

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatAE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinAna, "Anamnesetexte")
With RbTab
    .id = RibTab_Kat_EinAna
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Anamnese Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Anamnesetexte in das Krankenblatt ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 220
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Anamnesetexte, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Anamnesetexte an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinAna).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFAE

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM01
End If

'---

DoEvents
KMnPa "AnEi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnAE " & Err.Number
Resume Next

End Sub
Public Sub KMnAR()
On Error GoTo LaErr
'Menue Arzneimittel

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatAR
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinMed, "Artikelliste")
With RbTab
    .id = RibTab_Kat_EinMed
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Artikel Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Artikel in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Artikel, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Artikel an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetMed, "Artikelketten")
With RbTab
    .id = RibTab_Kat_KetMed
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Artikelkette in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneiketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinMed).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFME

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM06
End If

'---

DoEvents
KMnPa "ArLi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnAR " & Err.Number
Resume Next

End Sub

Public Sub KMnBA()
On Error GoTo LaErr
'Menue Kontoumsätze

Dim AktZa As Integer
Dim IdxZa As Integer
Dim BuJah As Integer
Dim IdxNr As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatBA
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinBan, "Offene Posten")
With RbTab
    .id = RibTab_Kat_EinBan
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Posten Zuordnen")
With CmCon
    .ToolTipText = "Ordnet dem Umsatz den markierten offenen Posten zu ohne diese auszugleichen"
    .IconId = IC32_Clipboard_Left
    .Width = GlRib
End With

If GlBuL = False Then 'echtes Löschen erlauben
    Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Loeschen, "Posten Stornieren")
Else
    Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Loeschen, "Posten Entfernen")
End If
With CmCon
    .ToolTipText = "Löscht den markierten offenen Posten"
    .IconId = IC32_Clipboard_Del
    .Width = GlRib
    .BeginGroup = True
End With

Set RbGrp = RbGps.AddGroup("Suche", RibGrp_Kat_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Alle Posten anzeigen")
With CmCon
    .ToolTipText = "Zeigt wieder alle offene Posten an"
    .IconId = IC32_Clipboard_Eye
    .Width = GlRib
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinBan, "Rechnungen")
With RbTab
    .id = RibTab_Kat_KetBan
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Rechnung Zuordnen")
With CmCon
    .ToolTipText = "Ordnet der Kontoauszugs-Buchung die markierte Rechnung zu"
    .IconId = IC32_Mail_Clip
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Alle Einträge anzeigen")
With CmCon
    .IconId = IC32_Mail_Eye
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Rechnungsfilter", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .ToolTipText = "Die Rechnungen aus welchem Jahr sollen angezeigt werden?"
    .Style = xtpButtonAutomatic
    .ThemedItems = True
    .Width = 130
    .DropDownItemCount = 18
    IdxZa = 1
    For BuJah = Year(Date) - 15 To Year(Date) + 2
        .AddItem BuJah
        .ItemData(IdxZa) = BuJah
        IdxZa = IdxZa + 1
    Next BuJah
    IdxNr = SCom(CmCom, Year(Date))
    .ListIndex = IdxNr
End With

Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .BeginGroup = True
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .ToolTipText = "Welche Belegtypen sollen angeziegt werden?"
    .Style = xtpButtonAutomatic
    .ThemedItems = True
    .Width = 130
    .AddItem "Standardrechnungen", 1
    .AddItem "Kostenvoranschläge", 2
    .AddItem "Laborrechnungen", 3
    .AddItem "Abrechnungsstelle", 4
    .AddItem "Gutschriften", 5
    .AddItem "Rechnungsaufträge", 6
    .AddItem "Gewerberechnungen", 7
    .AddItem "Importrechnungen", 8
    .AddItem "Rechnungsbelegtypen", 9
    .AddItem "Alle Belegtypen", 10
    .ListIndex = 9
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Suche in :")
    With CmCon
        .ToolTipText = "Geben Sie bitte hier Ihre Suchanfrage ein"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With

    Set CmCom = .Add(xtpControlComboBox, SY_SuCm4, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Welches Datenfeld soll durchsucht werden?"
        .IconId = IC16_View
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 80
        .AddItem "Rechn.-Nummer", 1
        .AddItem "Patientenname", 2
        .AddItem "Rechn.-Datum", 3
        .AddItem "Rechn.-Betrag", 4
        .ListIndex = 2
    End With

    Set CmCon = .Add(xtpControlLabel, SY_Cap04, " nach :")
    With CmCon
        .ToolTipText = "Tragen Sie hier Ihre Suchkriterien ein"
        .Style = xtpButtonCaption
    End With

    Set CmEdt = .Add(xtpControlEdit, SY_SuTex, vbNullString)
    With CmEdt
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Suchkriterium..."
        .ToolTipText = "Geben Sie bitte hier das Suchkriterium ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonAutomatic
        .Width = 110
    End With
End With

'ABC Leiste

RbBar.FindTab(RibTab_Kat_EinBan).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM08
End If


DoEvents
KMnPa "BaPo"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnBA " & Err.Number
Resume Next

End Sub
Public Sub KMnBE()
On Error GoTo LaErr
'Menue Begründungen

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatBE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinBeg, "Begründungen")
With RbTab
    .id = RibTab_Kat_EinBeg
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Begründung Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Begründungen in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Begründungen, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Begründungen an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinBeg).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFBE

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM02
End If

'---

DoEvents
KMnPa "BeEi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnBE " & Err.Number
Resume Next

End Sub
Public Sub KMnBU()
On Error GoTo LaErr
'Menue Buchhlatung

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatBU
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinBuc, "Buchungsvorl.")
With RbTab
    .id = RibTab_Kat_EinBuc
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Vorlage", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Buchung Hinzufügen")
With CmCon
    .ToolTipText = "Legt eine neue Buchung auf der Basis der Buchungsvorlage an"
    .IconId = IC32_Ordner_Add
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Kat_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Hinzufuegen, "Vorlage Hinzufügen")
With CmCon
    .ToolTipText = "Legt eine neue Buchungsvorlage an"
    .IconId = IC32_Dollar_Add
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Bearbeiten, "Vorlage Bearbeiten")
With CmCon
    .ToolTipText = "Öffnet den Dialog zum Bearbeiten der Buchungsvorlage"
    .IconId = IC32_Dollar_Edit
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Loeschen, "Vorlage Entfernen")
With CmCon
    .ToolTipText = "Entfernt die markierte Buchungsvorlage"
    .IconId = IC32_Dollar_Del
    .Width = GlRib
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetBuc, "Serienbuchung")
With RbTab
    .id = RibTab_Kat_KetBuc
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Serienbuchung", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Buchung Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Serienbuchungen in die Buchführung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Suchen, "Buchungs- Serie planen")
With CmCon
    .ToolTipText = "Öffnet den Assistenten zum generieren von Serienbuchungen"
    .IconId = IC32_Calendar_Folder
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ket_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Bearbeiten, "Eintrag Bearbeiten")
With CmCon
    .ToolTipText = "Öffnet den Dialog zum Bearbeiten der Serienbuchung"
    .IconId = IC32_Calendar_Edit
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Loeschen, "Eintrag Entfernen")
With CmCon
    .ToolTipText = "Entfernt die markierte Serienbuchung"
    .IconId = IC32_Calendar_Del
    .Width = GlRib
End With

'---

'ABC Leiste

RbBar.FindTab(RibTab_Kat_EinBuc).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM08
End If

'---

DoEvents
KMnPa "BuVo"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnBU " & Err.Number
Resume Next

End Sub
Public Sub KMnBV()
On Error GoTo LaErr
'Menue Buchhlatung

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatBV
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinBuc, "Buchungsvorl.")
With RbTab
    .id = RibTab_Kat_EinBuc
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Umsätze", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Vorlage Zuordnen")
With CmCon
    .ToolTipText = "Ordnet dem markierten Umsatz eine Buchungsvorlage zu"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Vorlagen", RibGrp_Kat_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Hinzufuegen, "Vorlage Hinzufügen")
With CmCon
    .ToolTipText = "Legt eine neue Buchungsvorlage an"
    .IconId = IC32_Dollar_Add
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Bearbeiten, "Vorlage Bearbeiten")
With CmCon
    .ToolTipText = "Öffnet den Dialog zum Bearbeiten der Buchungsvorlage"
    .IconId = IC32_Dollar_Edit
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Loeschen, "Vorlage Entfernen")
With CmCon
    .ToolTipText = "Löscht die markierte Buchungsvorlage aus der Auflistung"
    .IconId = IC32_Dollar_Del
    .Width = GlRib
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetBuc, "Buchungsregeln")
With RbTab
    .id = RibTab_Kat_KetBuc
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Umsätze", RibGrp_Ket_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Suchen, "Regel Zuordnen")
With CmCon
    .ToolTipText = "Ordnet die markierte Zuordnungsregel einem Umsatz zu."
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Zuordnungsregeln", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Hinzufuegen, "Regel Hinzufügen")
With CmCon
    .ToolTipText = "Fügt eine neue Zuordnungsregel hinzu"
    .IconId = IC32_Link_Add
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Bearbeiten, "Regel Bearbeiten")
With CmCon
    .ToolTipText = "Öffnet den Dialog zum Bearbeiten einer Zuordnungsregel"
    .IconId = IC32_Link_Edit
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Loeschen, "Regel Entfernen")
With CmCon
    .ToolTipText = "Löscht die markierte Zuordnungsregel aus der Auflistung"
    .IconId = IC32_Link_Del
    .Width = GlRib
End With

'---

'ABC Leiste

RbBar.FindTab(RibTab_Kat_EinBuc).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM08
End If

'---

DoEvents
KMnPa "BaVo"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnBV " & Err.Number
Resume Next

End Sub
Public Sub KMnDE()
On Error GoTo LaErr
'Menue Diagnosen

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatDE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Gruppierung, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Gruppierung, "Gruppierung")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinDia, "Diagnosen")
With RbTab
    .id = RibTab_Kat_EinDia
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, KA_Eint_Einfuegen, "Diagnose Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Diagnosen in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagDau, "Als Dauerdiagnose")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagTag, "Als Tagesdiagnose")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagVeAu, "Als Verdachtsdiagnose")
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagZuNa, "Als Zustandsdiagnose")
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 230
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Diagnosen, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Diagnosen an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetDia, "Diagnoseketten")
With RbTab
    .id = RibTab_Kat_KetDia
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Diagnoseketten in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagDau, "Als Dauerdiagnose")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagTag, "Als Tagesdiagnose")
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 230
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Diagnoseketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinDia).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFDE

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM03
End If

'---

DoEvents
KMnPa "DiEi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnDE " & Err.Number
Resume Next

End Sub

Public Sub KMnGE()
On Error GoTo LaErr
'Menue Gebühren

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatGE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Gruppierung, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Vorschauzeile, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Gruppierung, "Gruppierung")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinGeb, "Gebühren")
With RbTab
    .id = RibTab_Kat_EinGeb
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Gebühren Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Gebührenziffern in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 220
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Ansicht)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vorschlag, "Vorschlag")
With CmCon
    .ToolTipText = "Zeigt die zur gestellten ICD-10 Diagnose abrechnungsfähigen Gebührenziffern"
    .IconId = IC16_Lightpulb
    .Style = xtpButtonIconAndCaption
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Gebührenziffern, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Gebührenziffern an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetGeb, "Gebührenketten")
With RbTab
    .id = RibTab_Kat_KetGeb
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Gebührenketten in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 220
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Gebührenketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinGeb).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFGE

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM04
End If

'---

DoEvents
KMnPa "GbEi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnGE " & Err.Number
Resume Next

End Sub
Public Sub KMnKD()
On Error GoTo LaErr
'Menue Krankenblattdiagnosen

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatKD
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Gruppierung, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Gruppierung, "Gruppierung")
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinDiK, "Diagnosen")
With RbTab
    .id = RibTab_Kat_EinDiK
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, KA_Eint_Einfuegen, "Diagnose Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Diagnosen in das Krankenblatt ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagDau, "Als Aufnahmediagnose")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagTag, "Als Krankenblattdiagnose")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagVeAu, "Als Verdachtsdiagnose")
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagZuNa, "Als Zustandsdiagnose")
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 230
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Diagnosen, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Diagnosen an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetDiK, "Diagnoseketten")
With RbTab
    .id = RibTab_Kat_KetDiK
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Diagnoseketten in das Krankenblatt ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagDau, "Als Aufnahmediagnose")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_DiagTag, "Als Krankenblattdiagnose")
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 230
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Diagnoseketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinDiK).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFKD

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM03
End If

'---

DoEvents
KMnPa "KrDi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnKD " & Err.Number
Resume Next

End Sub
Public Sub KMnKM()
On Error GoTo LaErr
'Menue Krankenblattmedikamente

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatKM
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinMeK, "Arzneimittel")
With RbTab
    .id = RibTab_Kat_EinMeK
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, KA_Eint_Einfuegen, "Arzneimittel Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Arzneimittel in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_MediAuf, "Als Aufnahmemedikament")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_MediKra, "Als Krankenblattmedikament")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_MediThe, "Als Therapiekonzept")
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Arzneimittel, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneimittel an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetMeK, "Arzneiketten")
With RbTab
    .id = RibTab_Kat_KetMeK
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Arzneiketten in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_MediAuf, "Als Aufnahmemedikament")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_MediKra, "Als Krankenblattmedikament")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_MediThe, "Als Therapiekonzept")
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneiketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinMeK).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFKM

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM06
End If

'---

DoEvents
KMnPa "KrMe"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnKM " & Err.Number
Resume Next

End Sub
Public Sub KMnLE()
On Error GoTo LaErr
'Menue Laborleistungen

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatLE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinLab, "Laborparameter")
With RbTab
    .id = RibTab_Kat_EinLab
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Parameter Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Laborparameter in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Laborparamater, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Laborparameter an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetLab, "Laborketten")
With RbTab
    .id = RibTab_Kat_KetLab
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Ketten Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Laborketten in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Laborketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinLab).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFLE

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM05
End If

'---

DoEvents
KMnPa "LaEi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnLE " & Err.Number
Resume Next

End Sub
Public Sub KMnLv(ByVal PaFrm As String, Optional ByVal Kalen As Boolean = False)
On Error GoTo LaErr
'ListView Settings

Dim LiVw4 As XtremeSuiteControls.ListView
Dim LiIts As XtremeSuiteControls.ListViewItems
Dim ImMan As XtremeCommandBars.ImageManager

Select Case PaFrm
Case "TxPh": Set FM = frmKatTX
Case Else: Exit Sub
End Select

Set LiVw4 = FM.lstView4
Set LiIts = LiVw4.ListItems
If Kalen = True Then Set MoKal = FM.dtpDatu1
Set ImMan = frmMain.imgManag

With LiVw4
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = True
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = False
    .FlatScrollBar = False
    .Font.SIZE = 10
    .Font.Name = GlTFt.Name
    .ForeColor = vbBlack
    .FullRowSelect = True
    .GridLines = False
    .HideColumnHeaders = True
    .HideSelection = False
    .HotTracking = False
    .HoverSelection = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpListViewLabelManual
    .LabelWrap = True
    .MultiSelect = True
    .OLEDropMode = xtpOLEDropManual
    .View = xtpListViewReport
End With

If Kalen = True Then 'Enthält das Formular einen Kalender?
    If GlPoK = True Then 'Popupkalender
        With MoKal
            .AllowNoncontinuousSelection = False
            If GlSty = 8 Then 'Office 2013
                .BorderStyle = xtpDatePickerBorderStatic
            ElseIf GlSty = 7 Then 'Office 2013
                .BorderStyle = xtpDatePickerBorderStatic
            Else
                .BorderStyle = xtpDatePickerBorderOffice
            End If
            .MaxSelectionCount = 1
            .ShowNoneButton = True
            .ShowTodayButton = True
            .ToolTipText = "Markieren Sie bitte hier den Behandlungstag des Patienten"
            .Visible = False
        End With
    Else
        With MoKal
            .AllowNoncontinuousSelection = True
            .BorderStyle = xtpDatePickerBorderNone
            .MaxSelectionCount = GlMxK 'Maximal slektierbare Kalendertage
            .ShowNoneButton = False
            .ShowTodayButton = False
            .ToolTipText = "Markieren Sie bitte hier die Behandlungstage des Patienten"
            .Visible = True
        End With
    End If
    With MoKal
        .AskDayMetrics = True
        .AutoSizeRowCol = True
        .Enabled = True
        .FirstDayOfWeek = 2
        .FirstWeekOfYearDays = 4
        .HighlightToday = True
        .MultiSelectionMode = GlM04
        .RightToLeft = False
        .ShowNonMonthDays = True
        .ShowWeekNumbers = False
        .TextNoneButton = "Keine"
        .TextTodayButton = "Heute"
        .MonthDelta = 1
        .YearsTriangle = False
        Select Case GlSty
        Case 8: .VisualTheme = xtpCalendarThemeResource
        Case 7: .VisualTheme = xtpCalendarThemeResource
        Case Else: .VisualTheme = xtpCalendarThemeResource
        End Select
        .PaintManager.ButtonTextColor = -2147483640
        .PaintManager.ControlBackColor = -2147483643
        .PaintManager.DayBackColor = -2147483643
        .PaintManager.DayTextColor = -2147483640
        .PaintManager.DaysOfWeekBackColor = -2147483643
        .PaintManager.DaysOfWeekTextColor = -2147483640
        .PaintManager.ListControlBackColor = -2147483643
        .PaintManager.ListControlTextColor = -2147483640
        .PaintManager.NonMonthDayBackColor = -2147483643
        .PaintManager.NonMonthDayTextColor = -2147483640
        .PaintManager.SelectedDayBackColor = GlFac
        .PaintManager.SelectedDayTextColor = -2147483640
        .PaintManager.WeekNumbersBackColor = -2147483643
        .PaintManager.WeekNumbersTextColor = -2147483640
        .PaintManager.MonthHeaderBackColor = GlMoB
        .EnsureVisible DateAdd("m", -1, Date)
        .SelectRange Date, Date
        .Select Date
    End With
End If

Set LiIts = Nothing
Set LiVw4 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnLv " & Err.Number
Resume Next

End Sub

Public Sub KMnME()
On Error GoTo LaErr
'Menue Arzneimittel

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatME
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Popupkalender, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Popupkalender, "Popupkalender")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinMed, "Arzneimittel")
With RbTab
    .id = RibTab_Kat_EinMed
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Arzneimittel Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Arzneimittel in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Arzneimittel, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneimittel an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetMed, "Arzneiketten")
With RbTab
    .id = RibTab_Kat_KetMed
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Arzneiketten in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneiketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Suchleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(13))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Datum :")
    With CmCon
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_Kalen, vbNullString)
    With CmEdt
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag eingefügt werden soll"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .Width = 80
        .Text = Date
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinMed).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFME

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM06
End If

'---

DoEvents
KMnPa "MeEi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnME " & Err.Number
Resume Next

End Sub
Public Sub KMnPa(ByVal PaFrm As String)
On Error GoTo LaErr
'Menue Paintings

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Select Case PaFrm
Case "AnEi": Set FM = frmKatAE
Case "KrDi": Set FM = frmKatKD
Case "KrMe": Set FM = frmKatKM
Case "BeEi": Set FM = frmKatBE
Case "DiEi": Set FM = frmKatDE
Case "GbEi": Set FM = frmKatGE
Case "LaEi": Set FM = frmKatLE
Case "LaPa": Set FM = frmKatPE
Case "MeEi": Set FM = frmKatME
Case "ReEi": Set FM = frmKatRE
Case "TeDe": Set FM = frmKatTE
Case "TxPh": Set FM = frmKatTX
Case "BuVo": Set FM = frmKatBU
Case "ReSe": Set FM = frmKatRC
Case "BaPo": Set FM = frmKatBA
Case "BaVo": Set FM = frmKatBV
Case "BaRe": Set FM = frmKatBV
Case "ArLi": Set FM = frmKatAR
Case Else: Exit Sub
End Select

Set CmBrs = FM.comBar02
Set CmOpt = CmBrs.Options
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
End With

With CmOpt
    .AltDragCustomization = False
    .AlwaysShowFullMenus = True
    .AutoHideUnusedPopups = False
    .ExpandDelay = 100
    .ExpandHoverDelay = 100
    .FloatToolbarsByDoubleClick = False
    .IconsWithShadow = False
    .KeyboardCuesShow = xtpKeyboardCuesShowAlways
    .KeyboardCuesUse = xtpKeyboardCuesUseMenuOnly
    .LargeIcons = False
    .LunaColors = GlLun
    .MaxPopupWidth = 0.5
    .OfficeStyleDisabledIcons = True
    .SetIconSize True, 32, 32
    .ShowExpandButtonAlways = False
    .ShowFullAfterDelay = True
    .ShowPopupBarToolTips = False
    .ShowTextBelowIcons = False
    .ShowKeyboardTips = True
    .SyncFloatingToolbars = True
    .ToolBarAccelTips = True
    .ToolBarScreenTips = True
    .UpdatePeriod = 100
    .UseAltNumPadKeys = False
    .UseDisabledIcons = True
    .UseFadedIcons = False
    .UseSharedImageList = False
    .UseSystemSaveBitsStyle = False
    .Animation = xtpAnimateWindowsDefault
End With

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7: .VisualTheme = xtpThemeOffice2013
    Case 8: .VisualTheme = xtpThemeOffice2013
    Case Else: .VisualTheme = xtpThemeRibbon
    End Select
    If GlSty = 8 Then 'Office 2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    ElseIf GlSty = 7 Then 'Office 2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Else
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End If
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = True
    .ToolTipContext.ShowOfficeBorder = True
    .ToolTipContext.ShowShadow = True
    .ToolTipContext.ShowTitleAndDescription False, xtpToolTipIconNone
    .ToolTipContext.Style = xtpToolTipResource
    .PaintManager.AutoResizeIcons = False
    .PaintManager.ClearTypeTextQuality = GlCle
    .PaintManager.EnableAnimation = GlMeA
    .PaintManager.FlatMenuBar = False
    .PaintManager.FlatToolBar = False
    .PaintManager.SelectImageInPopupBar = True
    .PaintManager.ShowShadow = True
    .PaintManager.ThemedCheckBox = True
    .PaintManager.ThemedStatusBar = True
    .PaintManager.ThickCheckMark = False
    .KeyBindings.Add 0, VK_F1, KY_F1
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
End With

Set RbBar = CmBrs.Item(1)
With RbBar
    .AllowMinimize = False
    .AllowQuickAccessCustomization = False
    .AllowQuickAccessDuplicates = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .EnableAnimation = GlMeA
    .FontHeight = 11 'WICHTIG!
    .GroupsVisible = True
    .MinimumVisibleWidth = 100
    .RibbonPaintManager.HotTrackingGroups = True
    .RibbonPaintManager.CaptionFont.SIZE = 8
    .RibbonPaintManager.CaptionFont.Name = GlTFt.Name
    .RibbonPaintManager.WindowCaptionFont.SIZE = 8
    .RibbonPaintManager.WindowCaptionFont.Name = GlTFt.Name
    .ShowQuickAccess = False
    .ShowQuickAccessBelowRibbon = False
    .ShowCaptionAlways = False
    .Position = xtpBarTop
    .SetIconSize 16, 16
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
            .TabPaintManager.Appearance = xtpTabAppearanceVisualStudio2005
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 6
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
        Else
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 2
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        End If
    End Select
    .TabPaintManager.ClientFrame = xtpTabFrameNone
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    If GlSty < 7 Then
        .TabPaintManager.ButtonMargin.Top = 6
        .TabPaintManager.ButtonMargin.Bottom = 0
        .TabPaintManager.ButtonMargin.Left = 0
        .TabPaintManager.ButtonMargin.Right = 0
        .TabPaintManager.ClientMargin.Top = 0
        .TabPaintManager.ClientMargin.Bottom = 0
        .TabPaintManager.ClientMargin.Left = 0
        .TabPaintManager.ClientMargin.Right = 0
        .TabPaintManager.ControlMargin.Top = 0
        .TabPaintManager.ControlMargin.Bottom = 0
        .TabPaintManager.ControlMargin.Left = 0
        .TabPaintManager.ControlMargin.Right = 0
        .TabPaintManager.HeaderMargin.Top = 0
        .TabPaintManager.HeaderMargin.Bottom = 0
        .TabPaintManager.HeaderMargin.Left = 0
        .TabPaintManager.HeaderMargin.Right = 0
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
    Else
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        .TabPaintManager.HeaderMargin.Left = 0
    End If
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = False
    .TabPaintManager.HotTracking = True
    .TabPaintManager.Layout = xtpTabLayoutAutoSize
    .TabPaintManager.MinTabWidth = 100
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.Font.SIZE = 8
    .TabPaintManager.Font.Name = GlTFt.Name
End With

DoEvents
CmBrs.RecalcLayout
DoEvents
CmBrs.PaintManager.RefreshMetrics

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnPa " & Err.Number
Resume Next

End Sub
Public Sub KMnPE()
On Error GoTo LaErr
'Menue Labormodul

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatPE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinLaP, "Laborparameter")
With RbTab
    .id = RibTab_Kat_EinLaP
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Parameter Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Laborparameter in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Laborparamater, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Laborparameter an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetLaP, "Laborketten")
With RbTab
    .id = RibTab_Kat_KetLaP
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Laborketten in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Laborketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinLaP).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFPE

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM07
End If

'---

DoEvents
KMnPa "LaPa"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnPE " & Err.Number
Resume Next

End Sub
Public Sub KMnRC()
On Error GoTo LaErr
'Menue Rechnungen

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatRC
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
End With

'------------------------------ Einträge ------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinRec, "Serienrechnungen")
With RbTab
    .id = RibTab_Kat_EinRec
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Serienrechnung", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Suchen, "Rechnungen Erstellen")
With CmCon
    .ToolTipText = "Erstellt neue Rechnungen auf Basis der markieren Serienrechnungen"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, KA_Eint_Vorschlag, "Rechnungs- Serie planen")
With CmCon
    .ToolTipText = "Öffnet den Assistenten zum generieren von Serienrechnungen"
    .IconId = IC32_Calendar_Folder
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_Vorschlag, "Rechnungsserie Planen")
    CmCon.IconId = IC16_Calendar_Month
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, KA_Eint_Kopieren, "Rechnungsserie Kopieren")
    CmCon.IconId = IC16_Calendar_Copy
End With

Set RbGrp = RbGps.AddGroup("Eintrag", RibGrp_Ket_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Bearbeiten, "Bearbeiten")
With CmCon
    .ToolTipText = "Öffnet den Dialog zum Bearbeiten der Serienrechnung"
    .IconId = IC16_Doc_Edit
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneimittel an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Loeschen, "Entfernen")
With CmCon
    .ToolTipText = "Entfernt die markierte Serienrechnung"
    .IconId = IC16_Doc_Del
    .Style = xtpButtonIconAndCaption
End With

'------------------------------ Suchleiste ------------------------------

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Suche in :")
    With CmCon
        .ToolTipText = "Geben Sie bitte hier Ihre Suchanfrage ein"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With

    Set CmCom = .Add(xtpControlComboBox, SY_SuCm4, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Welches Datenfeld soll durchsucht werden?"
        .IconId = IC16_View
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 80
        .AddItem "Patientenname", 1
        .AddItem "Terminbetreff", 2
        .ListIndex = 1
    End With

    Set CmCon = .Add(xtpControlLabel, SY_Cap04, " nach :")
    With CmCon
        .ToolTipText = "Tragen Sie hier Ihre Suchkriterien ein"
        .Style = xtpButtonCaption
    End With

    Set CmEdt = .Add(xtpControlEdit, SY_SuTex, vbNullString)
    With CmEdt
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Suchkriterium..."
        .ToolTipText = "Geben Sie bitte hier das Suchkriterium ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonAutomatic
        .Width = 110
    End With
End With

'------------------------------ ABC-Leiste ------------------------------

RbBar.FindTab(RibTab_Kat_EinRec).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM08
End If

'---

DoEvents
KMnPa "ReSe"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnRC " & Err.Number
Resume Next

End Sub

Public Sub KMnRp(ByVal PaFrm As String, Optional ByVal Kalen As Boolean = False)
On Error GoTo LaErr
'Reportcontrol Settings

Dim RetWe As Long
Dim MuSel As Boolean
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Select Case PaFrm
Case "AnEi": Set FM = frmKatAE
Case "KrDi": Set FM = frmKatKD
Case "KrMe": Set FM = frmKatKM
Case "BeEi": Set FM = frmKatBE
Case "DiEi": Set FM = frmKatDE
Case "GbEi": Set FM = frmKatGE
Case "LaEi": Set FM = frmKatLE
Case "LaPa": Set FM = frmKatPE
Case "MeEi": Set FM = frmKatME
Case "ReEi": Set FM = frmKatRE
Case "TeDe": Set FM = frmKatTE
Case "TxPh": Set FM = frmKatTX
Case "BuVo": Set FM = frmKatBU
Case "BuSe": Set FM = frmKatBU
Case "ReSe": Set FM = frmKatRC
Case "BaPo": Set FM = frmKatBA
Case "BaVo": Set FM = frmKatBV
Case "BaRe": Set FM = frmKatBV
Case "ArLi": Set FM = frmKatAR
Case Else: Exit Sub
End Select

Set RpCon = FM.repCont7
If Kalen = True Then Set MoKal = FM.dtpDatu1
Set ImMan = frmMain.imgManag

Select Case PaFrm
Case "BuVo": MuSel = True
Case "BuSe": MuSel = False
Case "BaPo": MuSel = False
Case "BaVo": MuSel = False
Case Else: MuSel = True
End Select

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = MuSel
    .MultiSelectionMode = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = -2147483641
    .PaintManager.MaxPreviewLines = 5
    .PaintManager.ThemedInplaceButtons = True
    If GlKLi = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = Not GlKaU 'Zeilenumbruch der Kataloge
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlKaZ
    .ShowGroupBox = False
    If PaFrm = "TeDe" Then
        .PreviewMode = True
    Else
        .PreviewMode = False
    End If
    .ShowHeader = True
    .SortedDragDrop = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    .UnrestrictedDragDrop = True
    RetWe = .EnableDragDrop("Katalog", xtpReportAllowDragCopy)
End With

If Kalen = True Then 'Enthält das Formular einen Kalender?
    If GlPoK = True Then 'Popupkalender
        With MoKal
            .AllowNoncontinuousSelection = False
            If GlSty = 8 Then 'Office 2013
                .BorderStyle = xtpDatePickerBorderStatic
            ElseIf GlSty = 7 Then 'Office 2013
                .BorderStyle = xtpDatePickerBorderStatic
            Else
                .BorderStyle = xtpDatePickerBorderOffice
            End If
            .MaxSelectionCount = 1
            .ShowNoneButton = True
            .ShowTodayButton = True
            .ToolTipText = "Markieren Sie bitte hier den Behandlungstag des Patienten"
            .Visible = False
        End With
    Else
        With MoKal
            .AllowNoncontinuousSelection = True
            .BorderStyle = xtpDatePickerBorderNone
            .MaxSelectionCount = GlMxK 'Maximal slektierbare Kalendertage
            .ShowNoneButton = False
            .ShowTodayButton = False
            .ToolTipText = "Markieren Sie bitte hier die Behandlungstage des Patienten"
            .Visible = True
        End With
    End If
    With MoKal
        .AskDayMetrics = True
        .AutoSizeRowCol = True
        .Enabled = True
        .FirstDayOfWeek = 2
        .FirstWeekOfYearDays = 4
        .HighlightToday = True
        .MultiSelectionMode = GlM04
        .RightToLeft = False
        .ShowNonMonthDays = True
        .ShowWeekNumbers = False
        .TextNoneButton = "Keine"
        .TextTodayButton = "Heute"
        .MonthDelta = 1
        .YearsTriangle = False
        Select Case GlSty
        Case 8: .VisualTheme = xtpCalendarThemeResource
        Case 7: .VisualTheme = xtpCalendarThemeResource
        Case Else: .VisualTheme = xtpCalendarThemeResource
        End Select
        .PaintManager.ButtonTextColor = -2147483640
        .PaintManager.ControlBackColor = -2147483643
        .PaintManager.DayBackColor = -2147483643
        .PaintManager.DayTextColor = -2147483640
        .PaintManager.DaysOfWeekBackColor = -2147483643
        .PaintManager.DaysOfWeekTextColor = -2147483640
        .PaintManager.ListControlBackColor = -2147483643
        .PaintManager.ListControlTextColor = -2147483640
        .PaintManager.NonMonthDayBackColor = -2147483643
        .PaintManager.NonMonthDayTextColor = -2147483640
        .PaintManager.SelectedDayBackColor = GlFac
        .PaintManager.SelectedDayTextColor = -2147483640
        .PaintManager.WeekNumbersBackColor = -2147483643
        .PaintManager.WeekNumbersTextColor = -2147483640
        .PaintManager.MonthHeaderBackColor = GlMoB
        .EnsureVisible DateAdd("m", -1, Date)
        .SelectRange Date, Date
        .Select Date
    End With
End If

Set RpCon = Nothing
Set ImMan = Nothing
Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnRp " & Err.Number
Resume Next

End Sub
Public Sub KMnRE()
On Error GoTo LaErr
'Menue Langrezepte

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatRE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
End With

'Einträge

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinRez, "Arzneimittel")
With RbTab
    .id = RibTab_Kat_EinRez
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Arzneimittel Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Arzneimittel in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Arzneimittel, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneimittel an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Ketten

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetRez, "Arzneiketten")
With RbTab
    .id = RibTab_Kat_KetRez
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Kette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Arzneiketten in die Rechnung ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 160
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Arzneiketten an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Dosierungsleiste

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = GlPoK 'Popupkalender
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(1))
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "MO:")
    With CmCon
        .ToolTipText = "Dosierungsangabe für Morgens"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, SY_DoMor, vbNullString)
    With CmEdt
        .ToolTipText = "Dosierungsangabe für Morgens"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .ShowSpinButtons = True
        .Width = 40
        .Text = 0
    End With
    Set CmCon = .Add(xtpControlLabel, SY_Plac2, Space$(1))
    Set CmCon = .Add(xtpControlLabel, SY_Cap01, "MI:")
    With CmCon
        .ToolTipText = "Dosierungsangabe für Mittags"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, SY_DoMit, vbNullString)
    With CmEdt
        .ToolTipText = "Dosierungsangabe für Mittags"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .ShowSpinButtons = True
        .Width = 40
        .Text = 0
    End With
    Set CmCon = .Add(xtpControlLabel, SY_Plac3, Space$(1))
    Set CmCon = .Add(xtpControlLabel, SY_Cap03, "AB:")
    With CmCon
        .ToolTipText = "Dosierungsangabe für Abends"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, SY_DoAbe, vbNullString)
    With CmEdt
        .ToolTipText = "Dosierungsangabe für Abends"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .ShowSpinButtons = True
        .Width = 40
        .Text = 0
    End With
    Set CmCon = .Add(xtpControlLabel, SY_Plac4, Space$(1))
    Set CmCon = .Add(xtpControlLabel, SY_Cap04, "NA:")
    With CmCon
        .ToolTipText = "Dosierungsangabe für Nachts"
        .Style = xtpButtonCaption
    End With
    Set CmEdt = .Add(xtpControlEdit, SY_DoNac, vbNullString)
    With CmEdt
        .ToolTipText = "Dosierungsangabe für Nachts"
        .Style = xtpButtonCaption
        .EditStyle = xtpEditStyleCenter
        .ShowSpinButtons = True
        .Width = 40
        .Text = 0
    End With
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinRez).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Kett_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFRE

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM08
End If

'---

DoEvents
KMnPa "ReEi"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnRE " & Err.Number
Resume Next

End Sub
Public Sub KMnTE()
On Error GoTo LaErr
'Menue Terminplaner

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatTE
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuVon, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuBis, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei01, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei02, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei03, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei04, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei05, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei06, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei07, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei08, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei09, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei10, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei11, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei12, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei13, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei14, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei15, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei16, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei17, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei18, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei19, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei20, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Multimarker, "Multiselektion")
    CmCon.BeginGroup = True
End With

'------------------------------ Einträge ------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinTer, "Warteliste")
With RbTab
    .id = RibTab_Kat_EinTer
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Termin Einfügen")
With CmCon
    .ToolTipText = "Legt für den markierten Patienten einen Termin an"
    .IconId = IC32_Calendar_Add
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_WarNeu, "Wartenden Hinzufügen")
With CmCon
    .IconId = IC32_Clipboard_Patient
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_WarDel, "Wartenden Entfernen")
With CmCon
    .IconId = IC32_Clipboard_Del
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButtonPopup, AD_Termin_WarMai, "Termin Nachricht")
With CmCon
    .IconId = IC32_Calendar_Phone
    .Width = GlRib
    .BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_DocVrs, "Termin-Vorschlag")
    CmCon.IconId = IC16_Doc_Edit
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_EmlVrs, "Email-Vorschlag")
    CmCon.IconId = IC16_Earth_Mail
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_SMSVrs, "SMS-Vorschlag")
    CmCon.IconId = IC16_Phone_Mobil
    CmCon.BeginGroup = True
End With

'------------------------------ Ketten ------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetTer, "Terminketten")
With RbTab
    .id = RibTab_Kat_KetTer
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Terminkette Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierte Terminkette in den Kalender ein"
    .IconId = IC32_Calendar_Add
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Gruppe wählen..."
    .ToolTipText = "Bitte wählen Sie die gewünschte Gruppe"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Terminvorgaben, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Terminvorgaben an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'------------------------------ Suchleiste ------------------------------

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Suche in :")
    With CmCon
        .ToolTipText = "Geben Sie bitte hier Ihre Suchanfrage ein"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With

    Set CmCom = .Add(xtpControlComboBox, KA_SuCo2, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Welches Datenfeld soll durchsucht werden?"
        .IconId = IC16_View
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 80
        .AddItem "Patientenname", 1
        .AddItem "Bemerkung", 2
        .ListIndex = 1
    End With

    Set CmCon = .Add(xtpControlLabel, SY_Cap04, " nach :")
    With CmCon
        .ToolTipText = "Tragen Sie hier Ihre Suchkriterien ein"
        .Style = xtpButtonCaption
    End With

    Set CmEdt = .Add(xtpControlEdit, KA_SuFe2, vbNullString)
    With CmEdt
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Suchkriterium..."
        .ToolTipText = "Geben Sie bitte hier das Suchkriterium ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonAutomatic
        .Width = 110
    End With
End With

'------------------------------ ABC Leiste ------------------------------

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .Visible = False
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinTer).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KM_Multimarker).Checked = False
CmAcs(KM_Multimarker).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFTE

'---

DoEvents
KMnPa "TeDe"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnTE " & Err.Number
Resume Next

End Sub
Public Sub KMnTX()
On Error GoTo LaErr
'Menue Textphrasen

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKatTX
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    CmCon.BeginGroup = True
End With

'Textvorlagen

Set RbTab = RbBar.InsertTab(RibTab_Kat_EinTex, "Arzneimittel")
With RbTab
    .id = RibTab_Kat_EinTex
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Einfuegen, "Arzneimittel Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Arzneimittel in das Dokument ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Textvorlagen, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe1, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Textvorlagen an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'Textphrasen

Set RbTab = RbBar.InsertTab(RibTab_Kat_KetTex, "Textphrasen")
With RbTab
    .id = RibTab_Kat_KetTex
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Einfuegen, "Textphrasen &Einfügen")
With CmCon
    .ToolTipText = "Fügt die markierten Textphrasen in das Dokument ein"
    .IconId = IC32_Nav_Left
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 15
    .DropDownWidth = 140
    .ThemedItems = True
    .EditHint = "Hier Katalog wählen..."
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .Style = xtpButtonAutomatic
    .Width = 140
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Textvorlagen, die als Favoriten gekennzeichnet sind"
    .IconId = IC16_Doc_Check
    .Style = xtpButtonIconAndCaption
End With
Set CmEdt = RbGrp.Add(xtpControlEdit, KA_SuFe2, vbNullString)
With CmEdt
    .EditStyle = xtpEditStyleLeft
    .EditHint = "Hier Suchbegriff eingeben"
    .ShowLabel = False
    .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
    .Width = 140
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Vollst, "Vollständig")
With CmCon
    .ToolTipText = "Hebt das Suchergebnis auf und zeigt wieder alle Textphrasen an"
    .IconId = IC16_Doc_View
    .Style = xtpButtonIconAndCaption
End With

'ABC Leiste

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'---

RbBar.FindTab(RibTab_Kat_EinTex).Selected = True

CmAcs(KA_Eint_Vollst).Enabled = False
CmAcs(KA_Eint_Favoriten).Checked = GlFTX

If GlPoK = True Then 'Popupkalender
    CmAcs(KM_Multimarker).Checked = False
    CmAcs(KM_Multimarker).Enabled = False
Else
    CmAcs(KM_Multimarker).Checked = GlM11
End If

'---

DoEvents
KMnPa "TxPh"
DoEvents

Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " KMnTX " & Err.Number
Resume Next

End Sub
Public Sub KPrint(ByVal ForNa As String, ByVal DruVo As Boolean)
On Error GoTo LaErr
'Druckeinleitung

Dim IdxNr As Long
Dim FiNam As String
Dim LoNam As String
Dim Formu As Boolean
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set RpCls = RpCo8.Columns
Set RpSel = RpCo8.SelectedRows

Set clLis = New clsLisLab
Set clFil = New clsFile

FiNam = GlFrO & S_FoCh(ForNa) 'Formulardaten auslesen

If clFil.FilVor(FiNam) = True Then
    Formu = True
Else
    Formu = False
    SMeFr GlMeT, GlMeM, GlMeI, GlMeF, False, 1, True, FM.hwnd
End If

If Formu = True Then
    Select Case ForNa
    Case "KatLis":
            IdxNr = Mid$(GlNod, 2, Len(GlNod) - 1)
            If IdxNr = 0 Then Exit Sub
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .PfaTmp = GlTmp
                .IndxNr = IdxNr
                .DruDia = True
                .DruVor = GlDrV
                .LLPrKa
            End With
    Case "BesLis":
            IdxNr = Mid$(GlNod, 2, Len(GlNod) - 1)
            If IdxNr = 0 Then Exit Sub
            If Left$(GlNod, 1) <> "P" Then Exit Sub
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .PfaTmp = GlTmp
                .IndxNr = IdxNr
                .DruDia = True
                .DruVor = GlDrV
                .LLPrKa
            End With
    Case "InvLis":
            IdxNr = Mid$(GlNod, 2, Len(GlNod) - 1)
            If IdxNr = 0 Then Exit Sub
            If Left$(GlNod, 1) <> "P" Then Exit Sub
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .PfaTmp = GlTmp
                .IndxNr = IdxNr
                .DruDia = True
                .DruVor = GlDrV
                .LLPrKa
            End With
    Case "KetLis"
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Kat_ID0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    With clLis
                        .ForNam = ForNa
                        .FilNam = FiNam
                        .PfaTmp = GlTmp
                        .IndxNr = IdxNr
                        .DruDia = True
                        .DruVor = GlDrV
                        .LLPrKa
                    End With
                End If
            End If
    End Select
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo8 = Nothing

Set clFil = Nothing
Set clLis = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KPrint " & Err.Number
Resume Next

End Sub

Public Sub KSuAu(ByVal PaFrm As String)
On Error GoTo OrErr
'Hebt die markierten Suchbuchstaben wieder auf

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Select Case PaFrm
Case "AnEi": Set FM = frmKatAE
             Set CmBrs = FM.comBar02
Case "KrDi": Set FM = frmKatKD
             Set CmBrs = FM.comBar02
Case "KrMe": Set FM = frmKatKM
             Set CmBrs = FM.comBar02
Case "BeEi": Set FM = frmKatBE
             Set CmBrs = FM.comBar02
Case "DiEi": Set FM = frmKatDE
             Set CmBrs = FM.comBar02
Case "GbEi": Set FM = frmKatGE
             Set CmBrs = FM.comBar02
Case "LaEi": Set FM = frmKatLE
             Set CmBrs = FM.comBar02
Case "LaPa": Set FM = frmKatPE
             Set CmBrs = FM.comBar02
Case "MeEi": Set FM = frmKatME
             Set CmBrs = FM.comBar02
Case "ReEi": Set FM = frmKatRE
             Set CmBrs = FM.comBar02
Case "TeDe": Set FM = frmKatTE
             Set CmBrs = FM.comBar02
Case "BuVo": Set FM = frmKatBU
             Set CmBrs = FM.comBar02
Case "BuSe": Set FM = frmKatBU
             Set CmBrs = FM.comBar02
Case "ReSe": Set FM = frmKatRC
             Set CmBrs = FM.comBar02
Case "BaPo": Set FM = frmKatBA
             Set CmBrs = FM.comBar02
Case "TxPh": Set FM = frmKatTX
             Set CmBrs = FM.comBar02
Case "ArLi": Set FM = frmKatAR
             Set CmBrs = FM.comBar02
End Select

Set CmAcs = CmBrs.Actions

CmAcs(142).Checked = False
CmAcs(153).Checked = False
CmAcs(154).Checked = False
For AktZa = 65 To 90
    CmAcs(AktZa).Checked = False
Next AktZa

Select Case PaFrm
Case "AnEi":
    CmAcs(KA_Eint_Vollst).Enabled = False
Case "KrDi":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "KrMe":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "BeEi":
    CmAcs(KA_Eint_Vollst).Enabled = False
Case "DiEi":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "GbEi":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "LaEi":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "LaPa":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "MeEi":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "ReEi":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "TeDe":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "TxPh":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "BuVo":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "BuSe":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
Case "ReSe":
    CmAcs(KA_Eint_Vollst).Enabled = False
Case "BaPo":
    CmAcs(KA_Eint_Vollst).Enabled = False
Case "BaVo":
    CmAcs(KA_Eint_Vollst).Enabled = False
Case "ArLi":
    CmAcs(KA_Eint_Vollst).Enabled = False
    CmAcs(KA_Kett_Vollst).Enabled = False
End Select

CmAcs(KA_Eint_Favoriten).Checked = False
GlFav = False

Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KSuAu " & Err.Number
Resume Next

End Sub

Public Sub KSuch(ByVal PaFrm As String, ByVal IdxNr As Long, ByVal SuTyp As Integer, Optional ByVal SuJah As String)
On Error GoTo LaErr
'Filtert Einträge nach bestimmten Kriterien

Dim RetWe As Boolean
Dim SuIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit

Select Case PaFrm
Case "AnEi": Set FM = frmKatAE
             Set CmBrs = FM.comBar02
Case "KrDi": Set FM = frmKatKD
             Set CmBrs = FM.comBar02
Case "KrMe": Set FM = frmKatKM
             Set CmBrs = FM.comBar02
Case "BeEi": Set FM = frmKatBE
             Set CmBrs = FM.comBar02
Case "DiEi": Set FM = frmKatDE
             Set CmBrs = FM.comBar02
Case "GbEi": Set FM = frmKatGE
             Set CmBrs = FM.comBar02
Case "LaEi": Set FM = frmKatLE
             Set CmBrs = FM.comBar02
Case "LaPa": Set FM = frmKatPE
             Set CmBrs = FM.comBar02
Case "MeEi": Set FM = frmKatME
             Set CmBrs = FM.comBar02
Case "ReEi": Set FM = frmKatRE
             Set CmBrs = FM.comBar02
Case "TeDe": Set FM = frmKatTE
             Set CmBrs = FM.comBar02
Case "TxPh": Set FM = frmKatTX
             Set CmBrs = FM.comBar02
Case "BuVo": Set FM = frmKatBU
             Set CmBrs = FM.comBar02
Case "BuSe": Set FM = frmKatBU
             Set CmBrs = FM.comBar02
Case "ReSe": Set FM = frmKatRC
             Set CmBrs = FM.comBar02
Case "BaPo": Set FM = frmKatBA
             Set CmBrs = FM.comBar02
Case "ArLi": Set FM = frmKatAR
             Set CmBrs = FM.comBar02
End Select

Set CmAcs = CmBrs.Actions

GlAkt = True

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case SuTyp
Case 1: SuIdx = GlSuE.SuIdx
Case 2: SuIdx = GlSuN.SuIdx
Case 3: SuIdx = GlSuG.SuIdx
Case 4: SuIdx = GlSuI.SuIdx
Case 5: SuIdx = GlSuE.SuIdx
End Select

Select Case SuIdx
Case -1: 'Neuanlegen eines Eintrags
    P_List PaFrm, IdxNr, SuTyp, False, SuJah
Case 0: 'Suche aufheben
    Select Case SuTyp
    Case 1:
        CmAcs(KA_Eint_Vollst).Enabled = False
        GlSuE = GlSuX
        P_List PaFrm, IdxNr, SuTyp
    Case 2:
        CmAcs(KA_Kett_Vollst).Enabled = False
        GlSuN = GlSuX
        P_List PaFrm, IdxNr, SuTyp
    Case 5:
        CmAcs(KA_Kett_Vollst).Enabled = False
        GlSuN = GlSuX
        P_List PaFrm, IdxNr, SuTyp, False, SuJah
    End Select
Case Else:
    Select Case SuTyp
    Case 1:
        CmAcs(KA_Eint_Vollst).Enabled = True
        Select Case PaFrm
        Case "AnEi":
            If GlFAE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_LiSu PaFrm, IdxNr, SuTyp
            End If
        Case "BeEi":
            If GlFBE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_LiSu PaFrm, IdxNr, SuTyp
            End If
        Case "DiEi":
            If GlFDE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_Such PaFrm, IdxNr, SuTyp
            End If
        Case "GbEi":
            If GlFGE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_LiSu PaFrm, IdxNr, SuTyp
            End If
        Case "KrDi":
            If GlFKD = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_Such PaFrm, IdxNr, SuTyp
            End If
        Case "KrMe":
            If GlFKM = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_Such PaFrm, IdxNr, SuTyp
            End If
        Case "LaEi":
            If GlFLE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_LiSu PaFrm, IdxNr, SuTyp
            End If
        Case "MeEi":
            If GlFME = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_Such PaFrm, IdxNr, SuTyp
            End If
        Case "LaPa":
            If GlFPE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_LiSu PaFrm, IdxNr, SuTyp
            End If
        Case "ReEi":
            If GlFRE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_Such PaFrm, IdxNr, SuTyp
            End If
        Case "TeDe":
            If GlFTE = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_LiSu PaFrm, IdxNr, SuTyp
            End If
        Case "TxPh":
            If GlFTX = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_Such PaFrm, IdxNr, SuTyp
            End If
        Case "ArLi":
            If GlFAR = True Then
                P_LiSu PaFrm, IdxNr, SuTyp
            Else
                P_Such PaFrm, IdxNr, SuTyp
            End If
        Case Else:
            P_LiSu PaFrm, IdxNr, SuTyp
        End Select
    Case 2:
        CmAcs(KA_Kett_Vollst).Enabled = True
        P_LiSu PaFrm, IdxNr, SuTyp
    Case 5:
        CmAcs(KA_Kett_Vollst).Enabled = True
        P_LiSu PaFrm, IdxNr, SuTyp
    End Select
End Select

Select Case PaFrm
Case "AnEi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    CmEd1.Text = vbNullString
Case "KrDi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "KrMe":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "BeEi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    CmEd1.Text = vbNullString
Case "DiEi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "GbEi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "LaEi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "LaPa":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "MeEi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "ReEi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "TeDe":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu1, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd1, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "TxPh":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
Case "ReSe":
    Set CmEd1 = CmBrs.FindControl(CmEd1, SY_SuTex, , True)
    CmEd1.Text = vbNullString
Case "BaPo":
    Set CmEd1 = CmBrs.FindControl(CmEd1, SY_SuTex, , True)
    CmEd1.Text = vbNullString
Case "BaVo":
    Set CmEd1 = CmBrs.FindControl(CmEd1, SY_SuTex, , True)
    CmEd1.Text = vbNullString
Case "ArLi":
    Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
    Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
    Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
    Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)
    CmEd1.Text = vbNullString
    CmEd2.Text = vbNullString
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmAct = Nothing
Set CmBrs = Nothing

Set clFen = Nothing

GlAkt = False

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KSuch " & Err.Number
Resume Next

End Sub
Public Sub KSuGr()
On Error GoTo PeErr
'Lädt Daten in die Reportcontrols der Flyoutfenster

Dim KatNr As Long
Dim AktZa As Integer
Dim KaIdx As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmPop As XtremeCommandBars.CommandBarPopup
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set TxDe7 = frmMain.txtDeta7

If TxDe7.Text <> vbNullString Then
    KatNr = TxDe7.Text
Else
    KatNr = GlStK
End If

For AktZa = 1 To UBound(GlGKa)
    If GlGKa(AktZa, 0) = KatNr Then
        KaIdx = AktZa
        Exit For
    End If
Next AktZa

Set FM = frmKatAE 'Anamnese
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlAnG) 'Fragebogengruppen
    CmSu1.AddItem GlAnG(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlAnG(AktZa, 0)
Next AktZa
CmSu1.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "AnEi", 6, 1, GlFAE


Set FM = frmKatKD 'Krankenblattdiagnosen
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlDia)
    CmSu1.AddItem GlDia(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlDia(AktZa, 0)
Next AktZa
CmSu2.AddItem "Alle Diagnosegruppen", 1
CmSu1.ListIndex = 1
CmSu2.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Gruppierung).Checked = GlGrD
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "KrDi", 1, 1, GlFKD


Set FM = frmKatKM 'Krankenblattmedikamente
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlMed)
    CmSu1.AddItem GlMed(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlMed(AktZa, 0)
Next AktZa
CmSu2.AddItem "Alle Arzneigruppen", 1
CmSu1.ListIndex = 1
CmSu2.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "KrMe", 1, 1, GlFKM


Set FM = frmKatBE 'Begründungen
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
CmSu1.AddItem "Alle Begründungen"
CmSu1.ItemData(1) = 1
CmSu1.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "BeEi", 3, 1, GlFBE


Set FM = frmKatDE 'Diagnosen
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlDia)
    CmSu1.AddItem GlDia(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlDia(AktZa, 0)
Next AktZa
CmSu2.AddItem "Alle Diagnosegruppen", 1
CmSu1.ListIndex = 1
CmSu2.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Gruppierung).Checked = GlGrD
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "DiEi", 1, 1, GlFDE


Set FM = frmKatGE 'Gebühren
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlGKa)
    CmSu1.AddItem GlGKa(AktZa, 1)
    CmSu2.AddItem GlGKa(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlGKa(AktZa, 0)
    CmSu2.ItemData(AktZa) = GlGKa(AktZa, 0)
Next AktZa
CmSu1.ListIndex = KaIdx
CmSu2.ListIndex = KaIdx
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Gruppierung).Checked = GlKaG
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender


Set FM = frmKatLE 'Laborparameter
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
If CmSu2.ListCount > 0 Then CmSu2.Clear
For AktZa = 1 To UBound(GlLab)
    CmSu1.AddItem GlLab(AktZa, 1)
    CmSu2.AddItem GlLab(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlLab(AktZa, 0)
    CmSu2.ItemData(AktZa) = GlLab(AktZa, 0)
Next AktZa
CmSu1.ListIndex = GlStL
CmSu2.ListIndex = GlStL
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "LaEi", CmSu1.ListIndex, 1, GlFLE


Set FM = frmKatPE 'Labormodul
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
If CmSu2.ListCount > 0 Then CmSu2.Clear
For AktZa = 1 To UBound(GlLab)
    CmSu1.AddItem GlLab(AktZa, 1)
    CmSu2.AddItem GlLab(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlLab(AktZa, 0)
    CmSu2.ItemData(AktZa) = GlLab(AktZa, 0)
Next AktZa
CmSu1.ListIndex = GlStL
CmSu2.ListIndex = GlStL
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
P_List "LaPa", CmSu1.ListIndex, 1, GlFPE


Set FM = frmKatME 'Arzneimittel
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlMed)
    CmSu1.AddItem GlMed(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlMed(AktZa, 0)
Next AktZa
CmSu2.AddItem "Alle Arzneigruppen", 1
CmSu1.ListIndex = 1
CmSu2.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "MeEi", 1, 1, GlFME


Set FM = frmKatAR 'Artikelliste
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlArt)
    CmSu1.AddItem GlArt(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlArt(AktZa, 0)
Next AktZa
CmSu2.AddItem "Alle Artikelgruppen", 1
CmSu1.ListIndex = 1
CmSu2.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
CmAcs(KM_Popupkalender).Checked = GlPoK 'Popupkalender
P_List "ArLi", 1, 1, GlFAR


Set FM = frmKatRE 'Rechnungen
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlMed)
    CmSu1.AddItem GlMed(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlMed(AktZa, 0)
Next AktZa
CmSu2.AddItem "Alle Arzneigruppen", 1
CmSu1.ListIndex = 1
CmSu2.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
P_List "ReEi", 1, 1, GlFRE


Set FM = frmKatTX 'Textverarbeitung
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
If CmSu1.ListCount > 0 Then CmSu1.Clear
For AktZa = 1 To UBound(GlMed)
    CmSu1.AddItem GlMed(AktZa, 1)
    CmSu1.ItemData(AktZa) = GlMed(AktZa, 0)
Next AktZa
With CmSu2
    .AddItem "Alle Textphrasen", 1
    .ItemData(1) = 9
End With
CmSu1.ListIndex = 1
CmSu2.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Enabled = False
CmAcs(KM_Gitternetz).Enabled = False
P_List "TxPh", 1, 1, GlFTX


Set FM = frmKatBU 'Buchungsvorlagen
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
K_BuVpl "BuVo"
P_List "BuVo", 0, 1


Set FM = frmKatRC 'Langrezepte
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
P_TeSpl
P_List "ReSe", 1, 1


Set FM = frmKatBA 'Kontoumsätze
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
P_List "BaPo", 1, 1


Set FM = frmKatBV 'Kontoumsätze
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
K_BuVpl "BaVo"
P_List "BaVo", 0, 1


Set FM = frmKatTE 'Terminplaner
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmPop = CmBrs.FindControl(CmPop, SY_SuFar, , True)
CmSu1.AddItem "Alle Terminketten", 1
CmSu1.ListIndex = 1
CmAcs(KM_Zeilenumbruch).Checked = GlKaU
CmAcs(KM_Zeilenmarker).Checked = GlKaZ
CmAcs(KM_Gitternetz).Checked = GlKLi
P_List "TeDe", 0, 2, GlFTE

Set CmAcs = Nothing
Set CmSu1 = Nothing
Set CmSu2 = Nothing
Set CmBrs = Nothing

Exit Sub

PeErr:
If GlDbg = True Then SErLog Err.Description & " KSuGr " & Err.Number
Resume Next

End Sub
Public Sub KTree(Optional ByVal NoSel As Boolean = False)
On Error GoTo PeErr
'Stellt Baumstruktur im Katalogexplorer zusammen

Dim TmNod As String
Dim AktZa As Integer

Set FM = frmMain
Set TrLi2 = FM.trvList2

Select Case GlBut
Case RibTab_Kat_Eintrg:

    With TrLi2
        Set Knote = .Nodes.Add(, , "Z00", "Kataloge", IC16_Folder_View)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "A00", "Gebührenkataloge", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "C00", "Diagnosekataloge", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "G00", "Laborkataloge", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "I00", "Arzneikataloge", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "K00", "Begründungstexte", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "L00", "Anamnesetexte", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "M00", "Terminbetreffs", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "O00", "Textphrasenkatalog", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "P00", "Artikelkatalog", IC16_Folder_Close)
        Knote.Bold = True
        
        For AktZa = 1 To UBound(GlGKa) 'Gebührenkataloge
            Set Knote = TrLi2.Nodes.Add("A00", 4, "A" & GlGKa(AktZa, 0), GlGKa(AktZa, 1), IC16_Folder_Close)
        Next AktZa
        For AktZa = 1 To UBound(GlDia) 'Diagnosekataloge
            Set Knote = TrLi2.Nodes.Add("C00", 4, "C" & GlDia(AktZa, 0), GlDia(AktZa, 1), IC16_Folder_Close)
        Next AktZa
        For AktZa = 1 To UBound(GlMed) 'Arzneikataloge
            Set Knote = TrLi2.Nodes.Add("I00", 4, "I" & GlMed(AktZa, 0), GlMed(AktZa, 1), IC16_Folder_Close)
        Next AktZa
        For AktZa = 1 To UBound(GlLab) 'Laborkataloge
            Set Knote = TrLi2.Nodes.Add("G00", 4, "G" & GlLab(AktZa, 0), GlLab(AktZa, 1), IC16_Folder_Close)
        Next AktZa
        For AktZa = 1 To UBound(GlAnG) 'Fragebogengruppen
            Set Knote = TrLi2.Nodes.Add("L00", 4, "L" & GlAnG(AktZa, 0), GlAnG(AktZa, 1), IC16_Folder_Close)
        Next AktZa
        For AktZa = 1 To UBound(GlArt) 'Artikelkataloge
            Set Knote = TrLi2.Nodes.Add("P00", 4, "P" & GlArt(AktZa, 0), GlArt(AktZa, 1), IC16_Folder_Close)
        Next AktZa

        Set Knote = .Nodes.Add("K00", 4, "K3", "Alle Begründungen", IC16_Folder_Close)
        Set Knote = .Nodes.Add("M00", 4, "M5", "Alle Terminbetreffs", IC16_Folder_Close)
        Set Knote = .Nodes.Add("O00", 4, "O9", "Alle Textphrasen", IC16_Folder_Close)

        If NoSel = False Then
            .Nodes(GlNod).Expanded = True
            .Nodes(GlNod).Selected = True
            .Nodes(GlNod).Image = IC16_Folder_Open
            .Nodes(GlNod).EnsureVisible
        End If
    End With

Case RibTab_Kat_Ketten

    With TrLi2
        Set Knote = .Nodes.Add(, , "Z00", "Ketten", IC16_Folder_View)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "D00", "Gebührenketten", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "F00", "Diagnoseketten", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "H00", "Laborketten", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "J00", "Arzneiketten", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "R00", "Terminketten", IC16_Folder_Close)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "Q00", "Artikelketten", IC16_Folder_Close)
        Knote.Bold = True
        For AktZa = 1 To UBound(GlGKa) 'Gebührenkataloge
            Set Knote = TrLi2.Nodes.Add("D00", 4, "D" & GlGKa(AktZa, 0), GlGKa(AktZa, 1), IC16_Folder_Close)
        Next AktZa
        For AktZa = 1 To UBound(GlLab) 'Laborkataloge
            Set Knote = TrLi2.Nodes.Add("H00", 4, "H" & GlLab(AktZa, 0), GlLab(AktZa, 1), IC16_Folder_Close)
        Next AktZa
        
        Set Knote = .Nodes.Add("F00", 4, "F1", "Alle Diagnoseketten", IC16_Folder_Close)
        Set Knote = .Nodes.Add("J00", 4, "J1", "Alle Arzneiketten", IC16_Folder_Close)
        Set Knote = .Nodes.Add("R00", 4, "R1", "Alle Terminketten", IC16_Folder_Close)
        Set Knote = .Nodes.Add("Q00", 4, "Q1", "Alle Artikelketten", IC16_Folder_Close)
        
        TmNod = "D" & Mid$(GlNod, 2, Len(GlNod) - 1)
        If NoSel = False Then
            .Nodes(TmNod).Expanded = True
            .Nodes(TmNod).Selected = True
            .Nodes(TmNod).Image = IC16_Folder_Open
            .Nodes(TmNod).EnsureVisible
        End If
    End With
    
Case RibTab_Kat_Frage:

    With TrLi2
        Set Knote = .Nodes.Add(, , "Z00", "Kataloge", IC16_Folder_View)
        Knote.Bold = True
        Set Knote = .Nodes.Add("Z00", 4, "N00", "Fragebogenkataloge", IC16_Folder_Close)
        If GlBoV > 0 Then 'Fragebogen vorhanden
            Set Knote = TrLi2.Nodes.Add("N00", 4, "N0", "Neuaufnahmeformular", IC16_Folder_Lock)
            If GlNaf <> vbNullString Then
                Knote.Text = "<TextBlock>" & "Neuaufnahmeformular" & "<Run Foreground='Green' Text='" & "*" & "'/></TextBlock>"
            End If
            For AktZa = 1 To GlBoV 'Fragebogen vorhanden
                Set Knote = TrLi2.Nodes.Add("N00", 4, "N" & GlFrB(AktZa, 0), GlFrB(AktZa, 1), IC16_Folder_Close)
                If GlFrB(AktZa, 3) <> vbNullString Then
                    Knote.Text = "<TextBlock>" & GlFrB(AktZa, 1) & "<Run Foreground='Green' Text='" & "*" & "'/></TextBlock>"
                End If
            Next AktZa
        End If
        If NoSel = False Then
            .Nodes(GlNod).Expanded = True
            .Nodes(GlNod).Selected = True
            .Nodes(GlNod).Image = IC16_Folder_Open
            .Nodes(GlNod).EnsureVisible
        End If
    End With

End Select

Exit Sub

PeErr:
If GlDbg = True Then SErLog Err.Description & " KTree " & Err.Number
Resume Next

End Sub
Public Sub KUpKa(Optional ByVal RowNr As Long)
On Error GoTo OrErr
'Lädt den Detailbereich

Dim RowFi As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set CmBrs = FM.comBar01
Set RpRws = RpCo8.Rows
Set CmAcs = CmBrs.Actions

GlAkt = True

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If RowNr > 0 Then
    RowFi = RpCo8.TopRowIndex
End If

SSuch

If RowNr > 0 Then
    If RowNr >= RpRws.Count Then
        RowNr = RpRws.Count - 1
    End If
    RpCo8.TopRowIndex = RowFi
    If RpRws.Count > 0 Then
        RpRws.Row(0).Selected = False
        RpRws.Row(RowNr).EnsureVisible
        RpRws.Row(RowNr).Selected = True
        If GlFoc = True Then
            Set RpCo8.FocusedRow = RpRws.Row(RowNr)
        End If
    End If
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmBrs = Nothing
Set RpRws = Nothing
Set RpCo8 = Nothing

Set clFen = Nothing

GlAkt = False

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KUpKa " & Err.Number
Resume Next

End Sub
Public Sub SKale(ByVal KaAnz As Integer)
On Error GoTo SpErr
'Wechselt die Kalenderansicht

Dim DaSta As Date
Dim DaEnd As Date

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents

Set FM = frmMain
Set CaCol = FM.calCont1
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Set ViEvs = CaCol.ActiveView.GetSelectedEvents

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

CmAcs(SY_TE_Termin_Tag).Checked = False
CmAcs(SY_TE_Termin_ArWoche).Checked = False
CmAcs(SY_TE_Termin_ErWoche).Checked = False
CmAcs(SY_TE_Termin_Woche).Checked = False
CmAcs(SY_TE_Termin_Monat).Checked = False

DaSta = GlDFi 'Kalender Anfangsdatum

If ViEvs.Count > 0 Then
    For Each ViEvt In ViEvs
        If ViEvt.Selected = True Then
            ViEvt.Selected = False
        End If
    Next ViEvt
End If

Select Case KaAnz 'Kalenderanzeige
Case 1:
    DaEnd = DaSta
Case 2:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 4, DaSta)
Case 3:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 6, DaSta)
Case 4:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 6, DaSta)
Case 5:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 29, DaSta)
End Select

GlDFi = DaSta
GlDLa = DaEnd

With CaCol
    Select Case KaAnz 'Kalenderanzeige
    Case 1:
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayAllWeek
        .ViewType = xtpCalendarDayView
        CmAcs(SY_TE_Termin_Tag).Checked = True
    Case 2:
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayMo_Fr
        .ViewType = xtpCalendarWorkWeekView
        CmAcs(SY_TE_Termin_ArWoche).Checked = True
    Case 3:
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayAllWeek
        .ViewType = xtpCalendarFullWeekView
        CmAcs(SY_TE_Termin_ErWoche).Checked = True
    Case 4:
        .UseMultiColumnWeekMode = False
        .ViewType = xtpCalendarWeekView
        CmAcs(SY_TE_Termin_Woche).Checked = True
    Case 5:
        .UseMultiColumnWeekMode = False
        .ViewType = xtpCalendarMonthView
        CmAcs(SY_TE_Termin_Monat).Checked = True
    End Select
End With

IniSetVal "Layout", "KalAnz", KaAnz

GlCal = KaAnz 'Kalenderanzeige

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmBrs = Nothing
Set CaCol = Nothing

Set clFen = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SKale " & Err.Number
Resume Next

End Sub

Public Sub SKaSc(ByVal TiSca As Integer)
On Error GoTo LaErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CaCol As XtremeCalendarControl.CalendarControl

Set FM = frmMain
Set CaCol = FM.calCont1
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

GlSca = TiSca

CmAcs(ME_Termin_Minuten_01).Checked = False
CmAcs(ME_Termin_Minuten_05).Checked = False
CmAcs(ME_Termin_Minuten_10).Checked = False
CmAcs(ME_Termin_Minuten_15).Checked = False
CmAcs(ME_Termin_Minuten_20).Checked = False
CmAcs(ME_Termin_Minuten_30).Checked = False
CmAcs(ME_Termin_Minuten_60).Checked = False
CmAcs(ME_Termin_Minuten_120).Checked = False

CaCol.DayView.TimeScale = GlSca

CaCol.Populate

IniSetVal "TerSys", "TimSca", GlSca

Set CaCol = Nothing
Set CmAct = Nothing
Set CmBrs = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SKaSc " & Err.Number
Resume Next

End Sub

Public Function SKaSh(ByVal KalLi As Long, ByVal KalOb As Long, ByVal NeuDa As Date, ByVal mHwnd As Long, Optional ByVal Flag As Boolean = False) As Date
On Error GoTo LaErr

Dim RpCo6 As XtremeReportControl.ReportControl
Dim DaPi3 As XtremeCalendarControl.DatePicker

Dim Datu1 As Date
Dim DayFi As Date
Dim DayLa As Date
Dim KaBre As Long
Dim KaHoh As Long
Dim RetWe As Boolean

Set FM = frmMain
Set DaPi3 = FM.dtpDatu3

DayFi = NeuDa - 30
DayLa = NeuDa + 30

With DaPi3
    .RedrawControl
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    S_AbTe DayFi, DayLa
    .GetMinReqRect KaBre, KaHoh, 1, 1
    If Flag = True Then
        RetWe = .ShowModalEx(KalLi - KaBre - 4 - 4, KalOb, KaBre + 4, KaHoh + 4, mHwnd)
    Else
        RetWe = .ShowModalEx(-1, 20, KaBre + 4, KaHoh + 4, mHwnd)
    End If
End With

If RetWe = True Then
    If DaPi3.Selection.BlocksCount > 0 Then
        Datu1 = DaPi3.Selection.Blocks(0).DateBegin()
        If IsDate(Datu1) Then
            SKaSh = CDate(Datu1)
        Else
            SKaSh = NeuDa
        End If
    Else
        SKaSh = NeuDa
    End If
Else
    SKaSh = NeuDa
End If

Set DaPi3 = Nothing

Exit Function

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SKaSh " & Err.Number
Resume Next

End Function

Public Sub TrButt()
On Error GoTo OpErr
'Zeigt im TreeView auf einen anderen Knoten

Dim IdxNr As Long
Dim TrKey As String
Dim KetMe As Boolean
Dim GrVor As Boolean
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmGrp As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTerKat
Set CmBrs = FM.comBar03
Set RpCon = FM.repCont9
Set RpCls = RpCon.Columns
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmGrp = CmBrs.FindControl(CmGrp, KA_SuCo2, , True)

TrKey = Chr$(CmCom.ItemData(CmCom.ListIndex))
IdxNr = CLng(CmGrp.ItemData(CmGrp.ListIndex))

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case TrKey
Case "A": GlKtg = "A" & IdxNr
Case "C": GlKtg = "C" & IdxNr
Case "D": GlKtg = "D" & IdxNr
Case "F": GlKtg = "F1"
Case "G": GlKtg = "G2"
Case "H": GlKtg = "H2"
Case "I": GlKtg = "I" & IdxNr
Case "J": GlKtg = "J1"
Case "K": GlKtg = "K3"
Case "L": GlKtg = "L" & IdxNr
Case "M": GlKtg = "M1" 'Diagnosegruppen
Case "P": GlKtg = "P" & IdxNr
Case "Q": GlKtg = "Q1"
End Select

Select Case Left$(GlKtg, 1)
Case "A": GrVor = True
Case "C": GrVor = True
Case "D": KetMe = True
Case "F": KetMe = True
Case "G": KetMe = False
Case "H": KetMe = True
Case "I": KetMe = False
Case "J": KetMe = True
Case "K": KetMe = False
Case "L": KetMe = False
Case "M": KetMe = False
Case "P": KetMe = False
Case "Q": KetMe = True
End Select

CmAcs(KA_KaBu2).Enabled = False

If KetMe = True Then
    CmAcs(KA_KaBu1).Enabled = False
Else
    CmAcs(KA_KaBu1).Enabled = True
End If

CmAcs(142).Checked = False
CmAcs(153).Checked = False
CmAcs(154).Checked = False
For AktZa = 65 To 90
    CmAcs(AktZa).Checked = False
Next AktZa

RpCls.DeleteAll
DoEvents
TrSpla
DoEvents
Tr_List

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrButt " & Err.Number
Resume Next

End Sub
Public Sub TrGrla(ByVal Flag As String)
On Error GoTo OpErr
'Stellt bestimmte Formatierungen im GridEx ein

Dim IdxNr As Long
Dim TreKy As String
Dim KetMe As Boolean
Dim GrVor As Boolean
Dim GrAnz As Integer
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmGrp As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmTerKat
Set RpCon = FM.repCont9
Set CmBrs = FM.comBar03
Set CmAcs = CmBrs.Actions
Set RpCls = RpCon.Columns

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmGrp = CmBrs.FindControl(CmGrp, KA_SuCo2, , True)

Tit1 = "Keine Listenansicht"
Mld1 = "Diese Einstellung kann nur vorgenommen werden, wenn sich die Tabelle in der Listenansicht befindet"

TreKy = Chr$(CmCom.ItemData(CmCom.ListIndex))
IdxNr = CLng(CmGrp.ItemData(CmGrp.ListIndex))

Select Case TreKy
Case "A": GrVor = True
Case "C": GrVor = True
Case "D": KetMe = True
Case "F": KetMe = True
Case "G":
Case "H": KetMe = True
Case "I":
Case "J": KetMe = True
Case "K":
Case "L":
Case "P":
Case "Q": KetMe = True
End Select

With RpCon
    Select Case Flag
    Case "GrdGkp":
        If GrVor = True Then
            If .GroupsOrder.Count > 0 Then
                IniSetVal "TerKat", "GrdGkp", 0
                CmAcs(KM_Gruppierung).Checked = False
            Else
                IniSetVal "TerKat", "GrdGkp", -1
                CmAcs(KM_Gruppierung).Checked = True
            End If
        End If
    Case "GrdPrv":
        If KetMe = False Then
            If .PreviewMode = True Then
                IniSetVal "TerKat", "GrdPrv", 0
                CmAcs(KM_Vorschauzeile).Checked = False
            Else
                IniSetVal "TerKat", "GrdPrv", -1
                CmAcs(KM_Vorschauzeile).Checked = True
            End If
        End If
    Case "GrdZei":
        If .PaintManager.FixedRowHeight = True Then
            IniSetVal "TerKat", "GrdZei", -1
            CmAcs(KM_Zeilenumbruch).Checked = True
        Else
            IniSetVal "TerKat", "GrdZei", 0
            CmAcs(KM_Zeilenumbruch).Checked = False
        End If
    Case "GrdMkr":
        If GlZeM = True Then
            IniSetVal "TerKat", "GrdMkr", 0
            CmAcs(KM_Zeilenmarker).Checked = False
            GlZeM = False
        Else
            IniSetVal "TerKat", "GrdMkr", -1
            CmAcs(KM_Zeilenmarker).Checked = True
            GlZeM = True
        End If
    Case "GrdGrl":
        If CmAcs(KM_Gitternetz).Checked = True Then
            IniSetVal "TerKat", "GrdGrl", 0
            CmAcs(KM_Gitternetz).Checked = False
        Else
            IniSetVal "TerKat", "GrdGrl", -1
            CmAcs(KM_Gitternetz).Checked = True
        End If
    End Select
End With

With RpCon
    Select Case Flag
    Case "GrdGkp":
        If KetMe = False Then
            If GrVor = True Then
                If .GroupsOrder.Count = 0 Then
                    .SortOrder.Add .Columns(Kat_Gruppe)
                    .GroupsOrder.Add .Columns(Kat_Gruppe)
                    .GroupsOrder(0).SortAscending = True
                Else
                    .SortOrder.DeleteAll
                    .GroupsOrder.DeleteAll
                End If
                TrButt
            End If
        End If
    Case "GrdPrv":
            If KetMe = False Then
                If .PreviewMode = True Then
                    .PreviewMode = False
                Else
                    .PreviewMode = True
                End If
            End If
    Case "GrdZei":
            If .PaintManager.FixedRowHeight = True Then
                .PaintManager.FixedRowHeight = False
                Set RpCol = RpCls.Find(Kat_IDKurz)
                RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            Else
                .PaintManager.FixedRowHeight = True
            End If
    Case "GrdMkr":
            If .PaintManager.UseAlternativeBackground = True Then
                .PaintManager.UseAlternativeBackground = False
            Else
                .PaintManager.UseAlternativeBackground = True
            End If
    Case "GrdGrl":
            If .PaintManager.HorizontalGridStyle = xtpGridSolid Then
                .PaintManager.HorizontalGridStyle = xtpGridNoLines
                .PaintManager.VerticalGridStyle = xtpGridNoLines
            Else
                .PaintManager.HorizontalGridStyle = xtpGridSolid
                .PaintManager.VerticalGridStyle = xtpGridSolid
            End If
    End Select
    .Populate
End With

Set CmBrs = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrGrla " & Err.Number
Resume Next

End Sub
Public Sub TrGrp(Optional ByVal KatNr As Long)
On Error GoTo MeErr

Dim TrKey As String
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmGrp As XtremeCommandBars.CommandBarComboBox

Set FM = frmTerKat
Set TxDe7 = FM.txtDummy
Set CmBrs = FM.comBar03
Set CmAcs = CmBrs.Actions

Set CmCon = CmBrs.FindControl(CmCom, KA_KaBu2, , True)
Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmGrp = CmBrs.FindControl(CmGrp, KA_SuCo2, , True)

TrKey = Chr$(CmCom.ItemData(CmCom.ListIndex))

If KatNr > 0 Then TxDe7.Text = KatNr

If TxDe7.Text <> vbNullString Then KatNr = CLng(TxDe7.Text)

CmGrp.Clear

Select Case TrKey
Case "A": 'Gebührenkataloge
    For AktZa = 1 To UBound(GlGKa)
        CmGrp.AddItem GlGKa(AktZa, 1), AktZa
        CmGrp.ItemData(AktZa) = GlGKa(AktZa, 0)
    Next AktZa
    If KatNr > 0 Then
        For AktZa = 1 To UBound(GlGKa)
            If CmGrp.ItemData(AktZa) = KatNr Then
                CmGrp.ListIndex = AktZa
                Exit For
            End If
        Next AktZa
    Else
        CmGrp.ListIndex = 1
    End If
Case "C": 'Diagnosen
    For AktZa = 1 To UBound(GlDia)
        CmGrp.AddItem GlDia(AktZa, 1), AktZa
        CmGrp.ItemData(AktZa) = GlDia(AktZa, 0)
    Next AktZa
    CmGrp.ListIndex = 1
Case "D": 'Gebührenketten
    For AktZa = 1 To UBound(GlGKa)
        CmGrp.AddItem GlGKa(AktZa, 1), AktZa
        CmGrp.ItemData(AktZa) = GlGKa(AktZa, 0)
    Next AktZa
    If KatNr > 0 Then
        For AktZa = 1 To UBound(GlGKa)
            If CmGrp.ItemData(AktZa) = KatNr Then
                CmGrp.ListIndex = AktZa
                Exit For
            End If
        Next AktZa
    Else
        CmGrp.ListIndex = 1
    End If
Case "F": 'Diagnoseketten
    CmGrp.AddItem "Alle Diagnosegruppen", 1
    CmGrp.ItemData(1) = 1
    CmGrp.ListIndex = 1
Case "G": 'Laborparameter
    For AktZa = 1 To UBound(GlLab)
        CmGrp.AddItem GlLab(AktZa, 1), AktZa
        CmGrp.ItemData(AktZa) = GlLab(AktZa, 0)
    Next AktZa
    Select Case GlFri
    Case 1: CmGrp.ListIndex = 1 'Arzt (GOÄ)
    Case 2: CmGrp.ListIndex = 2 'Heilpraktiker (GebüH)
    Case 3: CmGrp.ListIndex = 1 'Zahnarzt (GOZ)
    Case 4: CmGrp.ListIndex = 1 'Veterinär (GOT)
    Case 5: CmGrp.ListIndex = 2 'Naturheilpraktiker (Tarif 590)
    Case 6: CmGrp.ListIndex = 2 'Physiotherapeut
    Case 7: CmGrp.ListIndex = 2 'Wahlarzt (AT)
    End Select
Case "H": 'Laborketten
    For AktZa = 1 To UBound(GlLab)
        CmGrp.AddItem GlLab(AktZa, 1), AktZa
        CmGrp.ItemData(AktZa) = GlLab(AktZa, 0)
    Next AktZa
    Select Case GlFri
    Case 1: CmGrp.ListIndex = 1 'Arzt (GOÄ)
    Case 2: CmGrp.ListIndex = 2 'Heilpraktiker (GebüH)
    Case 3: CmGrp.ListIndex = 1 'Zahnarzt (GOZ)
    Case 4: CmGrp.ListIndex = 1 'Veterinär (GOT)
    Case 5: CmGrp.ListIndex = 2 'Naturheilpraktiker (Tarif 590)
    Case 6: CmGrp.ListIndex = 2 'Physiotherapeut
    Case 7: CmGrp.ListIndex = 2 'Wahlarzt (AT)
    End Select
Case "I": 'Arzneimittel
    For AktZa = 1 To UBound(GlMed)
        CmGrp.AddItem GlMed(AktZa, 1), AktZa
        CmGrp.ItemData(AktZa) = GlMed(AktZa, 0)
    Next AktZa
    CmGrp.ListIndex = 1
Case "J": 'Arzneiketten
    CmGrp.AddItem "Alle Arzneigruppen", 1
    CmGrp.ItemData(1) = 1
    CmGrp.ListIndex = 1
Case "K": 'Begründungen
    CmGrp.AddItem "Alle Begründungen", 1
    CmGrp.ItemData(1) = 3
    CmGrp.ListIndex = 1
Case "M": 'Diagnosegruppen
    CmGrp.AddItem "Alle Diagnosegruppen", 1
    CmGrp.ItemData(1) = 1
    CmGrp.ListIndex = 1
Case "P": 'Artikelliste
    For AktZa = 1 To UBound(GlArt)
        CmGrp.AddItem GlArt(AktZa, 1), AktZa
        CmGrp.ItemData(AktZa) = GlArt(AktZa, 0)
    Next AktZa
    CmGrp.ListIndex = 1
Case "R": 'Terminketten
Case "Q": 'Artikelketten
    CmGrp.AddItem "Alle Artikelgruppen", 1
    CmGrp.ItemData(1) = 1
    CmGrp.ListIndex = 1
End Select

CmAcs(KA_KaBu2).Enabled = False

DoEvents
TrButt

Set CmAct = Nothing
Set CmBrs = Nothing

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrGrp " & Err.Number
Resume Next

End Sub
Public Sub TrMain(ByVal KatNr As Long)
On Error GoTo LaErr

Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmTerKat") = True Then
    frmTerKat.ZOrder 0
    Exit Sub
End If

TrLad = True

TrReg

Load frmTerKat

Set FM = frmTerKat

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    .FenDsk 2
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (460 / 2)
        .FeObn = (GlyGr / 2) - (600 / 2)
        .FeBre = 460
        .FeHoh = 600
    Else
        .FeLin = IniGetVal("TerKat", "FenLin")
        .FeObn = IniGetVal("TerKat", "FenObe")
        .FeBre = IniGetVal("TerKat", "FenBre")
        .FeHoh = IniGetVal("TerKat", "FenHoh")
    End If
End With

TrMnu
TrOpn
DoEvents

With clFen
    .FenMov
    Set CmBrs = FM.comBar03
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    TrPosi
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

DoEvents
TrGrp KatNr

Set clFen = Nothing

frmTerKat.Show
DoEvents
TrLad = False

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrMain " & Err.Number
Resume Next

End Sub
Private Sub TrMnu()
On Error GoTo MeErr
'Erstellt alle Menü- und Toolleisten

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmTerKat
Set CmBrs = FM.comBar03
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(KM_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Gitternetz, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Gruppierung, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Vorschauzeile, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Multimarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_KaBu1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_KaBu2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_KaBu3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Text = vbNullString
    CmPan.Width = 140
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Style = SBPS_STRETCH
    .Visible = True
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmPop = RbBar.Controls.Add(xtpControlPopup, KA_TabAn, "Ansicht")
With CmPop
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
    .flags = xtpFlagRightAlign
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, KM_Zeilenmarker, "Zeilenmarker")
    Set CmCon = .Add(xtpControlButton, KM_Zeilenumbruch, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, KM_Gitternetz, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, KM_Gruppierung, "Gruppierung")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, KM_Vorschauzeile, "Vorschauzeile")
End With
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmCon
    .ToolTipText = "Öffnet die Kurzhilfe"
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
    .flags = xtpFlagRightAlign
End With
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_KaBu4, "Schließen")
With CmCon
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .ToolTipText = "Schließt den Dialog"
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

'----------

Set RbTab = RbBar.InsertTab(RibTab_Kat_Eintrg, "Kataloge")
With RbTab
    .id = RibTab_Kat_Eintrg
End With
Set RbGps = RbTab.Groups

If GlBut = RibTab_Kat_Eintrg Then
    Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, KA_KaBu3, "Diagnosen Zuordnen")
    With CmCon
        .ToolTipText = "Ordnen die markierten Diagnosen den ausgewählten Gebührenpositionen zu"
        .IconId = IC32_Doc_Check
        .Width = GlRib
    End With
Else
    Set RbGrp = RbGps.AddGroup("", RibGrp_Kat_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, KA_KaBu3, "Positionen Einfügen")
    With CmCon
        .ToolTipText = "Fügt die markierten Einträge in den Termin ein"
        .IconId = IC32_Nav_Left
        .Width = GlRib
    End With
End If

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Kat_Suchen)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .EditHint = "Hier Katalog wählen..."
    .Style = xtpButtonAutomatic
    .ToolTipText = "Bitte wählen Sie den gewünschten Katalog"
    .ThemedItems = True
    .Width = 180
    If GlBut = RibTab_Kat_Eintrg Then
        .AddItem "Diagnosegruppen", 1
        .ItemData(1) = Asc("M")
    Else
        .AddItem "Gebührenkataloge", 1
        .ItemData(1) = Asc("A")
        .AddItem "Diagnosen", 2
        .ItemData(2) = Asc("C")
        .AddItem "Gebührenketten", 3
        .ItemData(3) = Asc("D")
        .AddItem "Diagnoseketten", 4
        .ItemData(4) = Asc("F")
        .AddItem "Laborkataloge", 5
        .ItemData(5) = Asc("G")
        .AddItem "Laborketten", 6
        .ItemData(6) = Asc("H")
        .AddItem "Arzneimittel", 7
        .ItemData(7) = Asc("I")
        .AddItem "Arzneiketten", 8
        .ItemData(8) = Asc("J")
        .AddItem "Begründungen", 9
        .ItemData(9) = Asc("K")
        .AddItem "Artikelliste", 10
        .ItemData(10) = Asc("P")
        .AddItem "Artikelketten", 11
        .ItemData(11) = Asc("Q")
    End If
    .ListIndex = 1
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_KaBu1, "Favoriten")
With CmCon
    .ToolTipText = "Zeigt nur die Einträge, die als Favoriten gekennzeichnet sind"
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Check
    If GlBut = RibTab_Kat_Eintrg Then .Enabled = False
End With
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_SuCo2, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .EditHint = "Hier Unterkategorie wählen..."
    .Style = xtpButtonAutomatic
    .ToolTipText = "Bitte wählen Sie die gewünschte Unterkategorie"
    .ThemedItems = True
    .Width = 180
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_KaBu2, "Vollständig")
With CmCon
    .ToolTipText = "Klicken Sie hier, um wieder alle Einträge anzuzeigen"
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_View
End With

'----------

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(4))
    Set CmCon = .Add(xtpControlLabel, KA_Capt3, "Suche:")
    With CmCon
        .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonIconAndCaption
        .flags = xtpFlagRightAlign
        .IconId = IC16_View
    End With
    Set CmEdt = .Add(xtpControlEdit, KA_SuFe1, vbNullString)
    With CmEdt
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Hier Suchbegriff eingeben"
        .ShowLabel = False
        .ToolTipText = "Geben Sie bitte hier den Begriff ein und bestätigen mit der ENTER-Taste"
        .Width = 180
    End With
End With

'----------

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, 42, Chr$(42))
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge in der Auswahl"
    End With
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
End With

'----------

With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
End With

With CmOpt
    .AltDragCustomization = False
    .AlwaysShowFullMenus = True
    .AutoHideUnusedPopups = False
    .ExpandDelay = 100
    .ExpandHoverDelay = 100
    .FloatToolbarsByDoubleClick = False
    .IconsWithShadow = False
    .KeyboardCuesShow = xtpKeyboardCuesShowAlways
    .KeyboardCuesUse = xtpKeyboardCuesUseMenuOnly
    .LargeIcons = False
    .LunaColors = GlLun
    .MaxPopupWidth = 0.5
    .OfficeStyleDisabledIcons = True
    .SetIconSize True, 32, 32
    .ShowExpandButtonAlways = False
    .ShowFullAfterDelay = True
    .ShowPopupBarToolTips = False
    .ShowTextBelowIcons = False
    .ShowKeyboardTips = True
    .SyncFloatingToolbars = True
    .ToolBarAccelTips = True
    .ToolBarScreenTips = True
    .UpdatePeriod = 100
    .UseAltNumPadKeys = False
    .UseDisabledIcons = True
    .UseFadedIcons = False
    .UseSharedImageList = False
    .UseSystemSaveBitsStyle = False
    .Animation = xtpAnimateWindowsDefault
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case 8:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case Else:
        If GlRah = True Then 'Office EnableThemeframe
            .VisualTheme = xtpThemeRibbon
        Else
            If GlFRg = True Then 'farbige Register
                .VisualTheme = xtpThemeResource
            Else
                .VisualTheme = xtpThemeRibbon
            End If
        End If
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End Select
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = True
    .ToolTipContext.ShowOfficeBorder = True
    .ToolTipContext.ShowShadow = True
    .ToolTipContext.ShowTitleAndDescription False, xtpToolTipIconNone
    .ToolTipContext.Style = xtpToolTipResource
    .PaintManager.AutoResizeIcons = False
    .PaintManager.ClearTypeTextQuality = GlCle
    .PaintManager.EnableAnimation = GlMeA
    .PaintManager.FlatMenuBar = False
    .PaintManager.FlatToolBar = False
    .PaintManager.SelectImageInPopupBar = True
    .PaintManager.ShowShadow = True
    .PaintManager.ThemedCheckBox = True
    .PaintManager.ThemedStatusBar = True
    .PaintManager.ThickCheckMark = False
    .KeyBindings.Add 0, VK_F1, KY_F1
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
End With

Set RbBar = CmBrs.Item(1)
With RbBar
    .AllowMinimize = False
    .AllowQuickAccessCustomization = False
    .AllowQuickAccessDuplicates = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .EnableAnimation = GlMeA
    .FontHeight = GlToF
    .GroupsVisible = True
    .MinimumVisibleWidth = 100
    .RibbonPaintManager.HotTrackingGroups = True
    .RibbonPaintManager.CaptionFont.SIZE = 8
    .RibbonPaintManager.CaptionFont.Name = GlTFt.Name
    .RibbonPaintManager.WindowCaptionFont.SIZE = 8
    .RibbonPaintManager.WindowCaptionFont.Name = GlTFt.Name
    .ShowQuickAccess = False
    .ShowQuickAccessBelowRibbon = False
    .ShowCaptionAlways = True
    .Position = xtpBarTop
    .SetIconSize 16, 16
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
            .TabPaintManager.Appearance = xtpTabAppearanceVisualStudio2005
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 6
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
        Else
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        End If
    End Select
    .TabPaintManager.Layout = xtpTabLayoutAutoSize
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.MinTabWidth = 100
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ClientFrame = xtpTabFrameNone
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = False
    .TabPaintManager.HotTracking = True
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.Font.SIZE = 8
    .TabPaintManager.Font.Name = GlTFt.Name
    If GlRDP = True Then
        .EnableFrameTheme
    Else
        If GlRah = True Then
            .EnableFrameTheme
        End If
    End If
End With

RbBar.FindTab(RibTab_Kat_Eintrg).Selected = True
CmAcs(KA_KaBu2).Enabled = False

Set CmPan = Nothing
Set CmSta = Nothing
Set CmPop = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrMnu " & Err.Number
Resume Next

End Sub

Private Sub TrOpn()
On Error GoTo PoErr

Dim RetWe As Long
Dim IdxNr As Long
Dim TreKy As String
Dim LiGrp As Boolean
Dim KetMe As Boolean
Dim ZeiMa As Boolean
Dim GrVor As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmGrp As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmTerKat
Set RpCon = FM.repCont9
Set CmBrs = FM.comBar03
Set CmAcs = CmBrs.Actions
Set RpCls = RpCon.Columns
Set ImMan = frmMain.imgManag

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmGrp = CmBrs.FindControl(CmGrp, KA_SuCo2, , True)

TreKy = Chr$(CmCom.ItemData(CmCom.ListIndex))
IdxNr = CLng(CmGrp.ItemData(CmGrp.ListIndex))

Select Case TreKy
Case "A": GrVor = True
Case "C": GrVor = True
Case "D": KetMe = True
Case "F": KetMe = True
Case "G": KetMe = False
Case "H": KetMe = True
Case "I": KetMe = False
Case "J": KetMe = True
Case "K": KetMe = False
Case "L": KetMe = False
Case "M": KetMe = False
Case "P": KetMe = False
Case "Q": KetMe = True
End Select

LiGrp = CBool(IniGetVal("TerKat", "GrdGkp"))
ZeiMa = CBool(IniGetVal("TerKat", "GrdMkr"))

GlZeM = ZeiMa

If GrVor = True Then
    CmAcs(KM_Gruppierung).Checked = LiGrp
Else
    CmAcs(KM_Gruppierung).Checked = False
End If

If KetMe = False Then
    CmAcs(KM_Vorschauzeile).Checked = GlGrV
Else
    CmAcs(KM_Vorschauzeile).Checked = False
End If

CmAcs(KM_Zeilenumbruch).Checked = GlGZe
CmAcs(KM_Zeilenmarker).Checked = ZeiMa
CmAcs(KM_Gitternetz).Checked = GlGrL

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = True
    If GlBut = RibTab_Kat_Eintrg Then
        .MultiSelectionMode = True
    Else
        .MultiSelectionMode = False
    End If
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind keine Leistungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Leistungen vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = GlAnZ
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = Not GlGZe
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 82, -2, 10, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZeM
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .SortedDragDrop = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    '.UnrestrictedDragDrop = True
    RetWe = .EnableDragDrop("Katalog", xtpReportAllowDrag)
End With

Set RpCon = Nothing
Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrOpn " & Err.Number
Resume Next

End Sub
Public Sub TrPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerKat
Set RpCon = FM.repCont9
Set CmBrs = FM.comBar03

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    RpCon.Move 0, ClObn, ClBre, ClHoh
End If

Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrPosi " & Err.Number
Resume Next

End Sub
Private Sub TrReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "TerKat") = False Then
    xPos = GlxGr - 470
    xGro = 460
    yGro = 600
     
    yPos = (GlyGr / 2) - (yGro / 2)
     
    IniSetSek "TerKat"
    IniSetVal "TerKat", "FenLin", xPos
    IniSetVal "TerKat", "FenObe", yPos
    IniSetVal "TerKat", "FenBre", xGro
    IniSetVal "TerKat", "FenHoh", yGro
    IniSetVal "TerKat", "GrdGkp", -1
    IniSetVal "TerKat", "GrdPrv", 0
    IniSetVal "TerKat", "GrdZei", 0
    IniSetVal "TerKat", "GrdMkr", 0
    IniSetVal "TerKat", "GrdGrl", -1
End If

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrReg " & Err.Number
Resume Next

End Sub
Private Sub TrSpla()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim TrKey As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmGrp As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTerKat
Set CmBrs = FM.comBar03
Set RpCon = FM.repCont9
Set RpCls = RpCon.Columns

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmGrp = CmBrs.FindControl(CmGrp, KA_SuCo2, , True)

TrKey = Chr$(CmCom.ItemData(CmCom.ListIndex))

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Kat_ID0, "ID0", 0, False)
    If TrKey = "M" Then
        Set RpCol = .Add(Kat_GOID, "Ziffer", 0, False)
    Else
        Set RpCol = .Add(Kat_GOID, "Ziffer", 80, False)
    End If
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        If RpCon.PaintManager.FixedRowHeight = False Then
            .Alignment = xtpAlignmentIconLeft Or xtpAlignmentIconTop
        Else
            .Alignment = xtpAlignmentIconLeft
        End If
    End With
    Set RpCol = .Add(Kat_IDKurz, "Bezeichnung", 200, False)
    If RpCon.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Kat_Gruppe, "Gruppe", 0, False)
    Select Case TrKey
    Case "C": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
    Case "F": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
    Case "K": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
    Case "L": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
    Case "M": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
    Case Else: Set RpCol = .Add(Kat_Preis1, "Preis", 60, False)
    End Select
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Kat_Sorter, "Sorter", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Resizable = False
    RpCol.Sortable = True
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(Kat_IDKurz).AutoSize = True

RpCon.AutoColumnSizing = True

Set CmBrs = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrSpla " & Err.Number
Resume Next

End Sub
Public Sub TrSuch(ByVal SuOpt As Integer, Optional ByVal SuStr As String, Optional ByVal SuPar As String)
On Error GoTo OrErr
'Filtert bestimmte Einträge heraus

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTerKat
Set CmBrs = FM.comBar03
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If SuOpt < 5 Then CmAcs(KA_KaBu1).Checked = False

If SuOpt > 0 And SuOpt < 5 Then CmAcs(KA_KaBu2).Enabled = True

CmAcs(142).Checked = False
CmAcs(153).Checked = False
CmAcs(154).Checked = False
For AktZa = 65 To 90
    CmAcs(AktZa).Checked = False
Next AktZa

Select Case SuOpt
Case 0: CmAcs(KA_KaBu2).Enabled = False
        Tr_Filt 0
Case 1: Select Case Left$(GlKtg, 1)
        Case "C": Tr_Such SuStr
        Case "I": Tr_Such SuStr
        Case Else: Tr_Filt 1, SuStr
        End Select
Case 2: Tr_Filt 2, SuStr
Case 3: Tr_Filt 3, SuStr, SuPar
Case 4: Tr_Filt 4, SuStr
        Select Case SuStr
        Case "Ä": CmAcs(142).Checked = True
        Case "Ö": CmAcs(153).Checked = True
        Case "Ü": CmAcs(154).Checked = True
        Case Else: CmAcs(Asc(SuStr)).Checked = True
        End Select
Case 5:
        If CmAcs(KA_KaBu1).Checked = False Then
            Tr_Filt 5
            CmAcs(KA_KaBu1).Checked = True
        Else
            Tr_Filt 0
            CmAcs(KA_KaBu1).Checked = False
        End If
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmSta = Nothing
Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TrSuch " & Err.Number
Resume Next

End Sub
