VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatRC 
   BorderStyle     =   0  'Kein
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   1815
      Left            =   1560
      TabIndex        =   0
      Top             =   2760
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   3201
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKatRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form

Private AktCo As VB.Control
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TbBar As XtremeCommandBars.TabToolBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmSta As XtremeCommandBars.StatusBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private Sub FEdit()
On Error GoTo AnErr

Dim IdxNr As Long
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Set CmBrs = Me.comBar02
Set RpCo7 = Me.repCont7
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    Set RpCol = RpCls.Find(Ter_ID2)
    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
    TeMain IdxNr
End If

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEdit " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50511)
TeMai = IniGetOpt("Hilfe", 50512)
TeInh = IniGetOpt("Hilfe", 50513)
TeFus = IniGetOpt("Hilfe", 50514)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub

Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

Select Case RbTab.id
Case RibTab_Kat_EinRec:
        RpCls(1).Width = 110
        RpCo7.PaintManager.NoFieldsAvailableText = "Es sind keine Serienrechnungen vorhanden"
        RpCo7.PaintManager.NoItemsText = "Es sind keine Serienrechnungen vorhanden"
Case RibTab_Kat_KetRec:
        RpCls(1).Width = 0
        RpCo7.PaintManager.NoFieldsAvailableText = "Es sind keine Rechnungsvorlagen vorhanden"
        RpCo7.PaintManager.NoItemsText = "Es sind keine Rechnungsvorlagen vorhanden"
End Select

Select Case RbTab.id
Case RibTab_Kat_EinRec: P_List "ReSe", 0, 1
Case RibTab_Kat_KetRec: P_List "ReSe", 0, 2
End Select

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim AktZa As Integer
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(Ter_ID0, "ID0", 0, False)
    Set RpCol = .Add(Ter_ID2, "ID2", 0, False)
    Set RpCol = .Add(Ter_IDR, "IDR", 0, False)
    Set RpCol = .Add(Ter_IDSer, "IDSer", 0, False)
    Set RpCol = .Add(Ter_Icon, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Calendar_Day
    End With
    Set RpCol = .Add(Ter_Aufgabe, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Mail_Close
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_Status, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Pin_Gray
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_VonDat, "Startdatum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Ter_BisDat, "BisDat", 0, False)
    Set RpCol = .Add(Ter_ZeiVon, "Von", 0, True)
    Set RpCol = .Add(Ter_ZeiBis, "Bis", 0, True)
    Set RpCol = .Add(Ter_ZeiVor, "ZeiVor", 0, False)
    Set RpCol = .Add(Ter_Priorität, "Prio.", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ter_Vorwarn, "Vorwarn", 0, False)
    Set RpCol = .Add(Ter_Farbe, "Farbe", 0, False)
    Set RpCol = .Add(Ter_Anzahl, "Anzahl", 0, False)
    Set RpCol = .Add(Ter_Abgehakt, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Check
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_Erledigt, "Erledigt", 0, True)
    Set RpCol = .Add(Ter_Patient, "Patient", 0, True)
    Set RpCol = .Add(Ter_IDKurz, "Betreff", 0, True)
    Set RpCol = .Add(Ter_Datei, "Datei", 0, False)
    Set RpCol = .Add(Ter_Datum, "Hinzugefügt", 0, False)
    Set RpCol = .Add(Ter_Change, "Geändert", 0, False)
    Set RpCol = .Add(Ter_Farbtyp, "Status", 0, False)
    Set RpCol = .Add(Ter_Folge, "Folge", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Ter_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Ter_Raum, "Raum", 0, True)
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        If GlRaV = True Then
            For AktZa = 1 To UBound(GlRmu)
                .EditOptions.Constraints.Add GlRmu(AktZa, 1), GlRmu(AktZa, 2)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(Ter_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Ter_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Ter_Wiederholung, "Wiederholung", 0, False)
    Set RpCol = .Add(Ter_Selekt, "G", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Editable = False
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_Wochentag, "Tag", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Ter_MasTer, "Serie", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_AbrKom, "Abgerechnet", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_TerBet, "Terminbetrag", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_Monat, "Monat", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        .EditOptions.Constraints.Add "Januar", 1
        .EditOptions.Constraints.Add "Februar", 2
        .EditOptions.Constraints.Add "März", 3
        .EditOptions.Constraints.Add "April", 4
        .EditOptions.Constraints.Add "Mai", 5
        .EditOptions.Constraints.Add "Juni", 6
        .EditOptions.Constraints.Add "Juli", 7
        .EditOptions.Constraints.Add "August", 8
        .EditOptions.Constraints.Add "September", 9
        .EditOptions.Constraints.Add "Oktober", 10
        .EditOptions.Constraints.Add "November", 11
        .EditOptions.Constraints.Add "Dezember", 12
    End With
    Set RpCol = .Add(Ter_SerBet, "Serienbetrag", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_BezBet, "Bezahlt", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_BezBet2, "Bezahlt2", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_BetOff, "Offen", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_Fallig1, "Fälligkeit", 0, False)
    Set RpCol = .Add(Ter_Fallig2, "Fälligkeit2", 0, False)
    Set RpCol = .Add(Ter_Passiv, vbNullString, 0, False)
End With

For Each RpCol In RpCls
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = True
        .Sortable = True
        .AutoSize = False
        .AutoSortWhenGrouped = False
    End With
Next RpCol

RpCls(Ter_ID0).Width = 0
RpCls(Ter_ID2).Width = 0
RpCls(Ter_IDR).Width = 0
RpCls(Ter_IDSer).Width = 0
RpCls(Ter_Icon).Width = 20
RpCls(Ter_Aufgabe).Width = 20
RpCls(Ter_Status).Width = 20
RpCls(Ter_VonDat).Width = 80
RpCls(Ter_BisDat).Width = 0
RpCls(Ter_ZeiVon).Width = 0
RpCls(Ter_ZeiBis).Width = 0
RpCls(Ter_ZeiVor).Width = 0
RpCls(Ter_Priorität).Width = 0
RpCls(Ter_Vorwarn).Width = 0
RpCls(Ter_Farbe).Width = 0
RpCls(Ter_Anzahl).Width = 0
RpCls(Ter_Abgehakt).Width = 0
RpCls(Ter_Erledigt).Width = 0
RpCls(Ter_Patient).Width = 200
RpCls(Ter_IDKurz).Width = 180
RpCls(Ter_Datei).Width = 0
RpCls(Ter_Datum).Width = 120
RpCls(Ter_Change).Width = 0
RpCls(Ter_Farbtyp).Width = 0
RpCls(Ter_Folge).Width = 60
RpCls(Ter_IDP).Width = 180
RpCls(Ter_IDM).Width = 180
RpCls(Ter_Raum).Width = 110
RpCls(Ter_GuiID).Width = 0
RpCls(Ter_Kommentar).Width = 0
RpCls(Ter_Wiederholung).Width = 0
RpCls(Ter_Selekt).Width = 20
RpCls(Ter_Wochentag).Width = 30
RpCls(Ter_MasTer).Width = 60
RpCls(Ter_AbrKom).Width = 150
RpCls(Ter_TerBet).Width = 80
RpCls(Ter_Monat).Width = 0
RpCls(Ter_SerBet).Width = 80
RpCls(Ter_BezBet).Width = 0
RpCls(Ter_BezBet2).Width = 0
RpCls(Ter_BetOff).Width = 0
RpCls(Ter_Fallig1).Width = 0
RpCls(Ter_Fallig2).Width = 0
RpCls(Ter_Passiv).Width = 0

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpal " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo OpErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long

Set CmBrs = Me.comBar02
Set RpCon = Me.repCont7

CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
ClHoh = ClHoh - ClObn

If ClHoh - 2200 > 0 Then
    RpCon.Move ClLin, ClObn, ClBre - ClLin, ClHoh
End If

Set RpCon = Nothing
Set CmBrs = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FLoe()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

K_RcLo RbTab.id

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoe " & Err.Number
Resume Next

End Sub

Private Sub repCont7_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Select Case RbTab.id
Case RibTab_Kat_EinRec:
        If IsDate(Row.Record(Ter_VonDat).Value) = True Then
            If CDate(Row.Record(Ter_VonDat).Value) <= Date Then
                Metrics.Font.Bold = True
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
        End If
Case RibTab_Kat_KetRec:
        
End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatRC = Nothing
End Sub
Private Sub Form_Load()
    KMnRp "ReSe"
    FSpal
End Sub
Private Sub comBar02_Resize()
    If GlDcP = False Then
        FPosi
    End If
End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlAkt = False Then
        If Control.id = XTP_ID_RIBBONCONTROLTAB Then
            FTabu
        Else
            FTool Control.id
        End If
    End If
End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FEdit
End Sub

Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KA_Hilfe: FHilfe
Case KM_Zeilenumbruch: KGrKa "GrdZei"
Case KM_Zeilenmarker: KGrKa "GrdMkr"
Case KM_Gitternetz: KGrKa "GrdGrl"
Case KM_Multimarker: KGrKa "MulMar"
Case KA_Eint_Suchen: frmReErs.Show vbModal
Case KA_Eint_Vorschlag: TeVoMa
Case KA_Eint_Kopieren: TerAs
Case KA_Eint_Bearbeiten: FEdit
Case KA_Eint_Loeschen: FLoe
Case KA_Kett_Bearbeiten: FEdit
Case KA_Kett_Loeschen: FLoe
Case KA_Eint_Vollst: FSuAu
Case SY_SuCm4: FSuGr
Case SY_SuTex: FSuch
End Select

GlToo = False

End Sub
Private Sub FSuGr()
On Error GoTo OrErr
'Favoriten Knopf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmSu1 = CmBrs.FindControl(CmSu1, SY_SuCm4, , True)

KSuAu "ReSe"

Select Case RbTab.id
Case RibTab_Kat_EinRec:
        With GlSuE
            .SuIdx = 0
        End With
Case RibTab_Kat_KetRec:
        With GlSuN
            .SuIdx = 0
        End With
End Select

KSuch "ReSe", 3, 1
DoEvents

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuGr " & Err.Number
Resume Next

End Sub
Private Sub FSuch()
On Error GoTo OrErr
'Sucheingabe

Dim SuStr As String
Dim TyIdx As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmEd1 As XtremeCommandBars.CommandBarEdit

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmSu1 = CmBrs.FindControl(CmSu1, SY_SuCm4, , True)
Set CmEd1 = CmBrs.FindControl(CmEd1, SY_SuTex, , True)

TyIdx = CmSu1.ListIndex

KSuAu "ReSe"

SuStr = CmEd1.Text
With GlSuE
    .SuIdx = TyIdx
    .SuStr = SuStr
End With

KSuch "ReSe", 3, 1
DoEvents

If RpCo7.Records.Count = 0 Then
    CmEd1.Text = vbNullString
    SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuch " & Err.Number
Resume Next

End Sub
Private Sub FSuAu()
On Error GoTo OrErr
'Hebt die markierten Suchbuchstaben wieder auf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCo7 = Me.repCont7

Set CmSu1 = CmBrs.FindControl(CmSu1, SY_SuCm4, , True)

GlSuE = GlSuX

KSuAu "ReSe"
DoEvents

KSuch "ReSe", 3, 1

DoEvents
RpCo7.SetFocus

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuAu " & Err.Number
Resume Next

End Sub
