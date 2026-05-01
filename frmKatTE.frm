VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatTE 
   BorderStyle     =   0  'Kein
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   3201
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   240
      Top             =   360
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKatTE"
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

Private TabId As Integer
Private Sub FEinf()
On Error Resume Next

Dim IdxNr As Long
Dim MitNr As Long
Dim ManNr As Long
Dim PatNa As String
Dim TeBet As String
Dim TeKom As String
Dim KetNa As String
Dim KetKu As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpSel As XtremeReportControl.ReportSelectedRows

Set CmBrs = Me.comBar02
Set RpCo7 = Me.repCont7
Set RbBar = CmBrs.Item(1)
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows
Set RbTab = RbBar.SelectedTab

Select Case RbTab.id
Case RibTab_Kat_EinTer:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    TeBet = vbNullString
                    TeKom = RpRow.Record.PreviewText
                    Set RpCol = RpCls.Find(Kat_ID0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Kat_GOID)
                    MitNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Kat_IDKurz)
                    PatNa = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Kat_Preis1)
                    ManNr = RpRow.Record(RpCol.ItemIndex).Value
                    STerm True, False, 0, IdxNr, PatNa, TeBet, TeKom, MitNr, ManNr
                End If
            End If
Case RibTab_Kat_KetTer:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Kat_ID0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Kat_GOID)
                    KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                    Set RpCol = RpCls.Find(Kat_IDKurz)
                    KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                    GlNod = "R1"
                    GlKSt = "TeDe"
                    GlKeE = True
                    EMain IdxNr, KetNa, KetKu
                End If
            End If
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing
Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case TabId
Case RibTab_Kat_EinTer:
    TeTit = IniGetOpt("Hilfe", 50531)
    TeMai = IniGetOpt("Hilfe", 50532)
    TeInh = IniGetOpt("Hilfe", 50533)
    TeFus = IniGetOpt("Hilfe", 50534)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
Case RibTab_Kat_KetTer:
    If WindowLoad("frmKetten") = False Then
        TeTit = IniGetOpt("Hilfe", 50541)
        TeMai = IniGetOpt("Hilfe", 50542)
        TeInh = IniGetOpt("Hilfe", 50543)
        TeFus = IniGetOpt("Hilfe", 50544)
        SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
    End If
End Select

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

RpCon.Move ClLin, ClObn, ClBre - ClLin, ClHoh

Set RpCon = Nothing
Set CmBrs = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Public Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCo7
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Kat_ID0, "ID0", 0, False)
    Set RpCol = .Add(Kat_GOID, "Nummer", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        If RpCo7.PaintManager.FixedRowHeight = False Then
            .Alignment = xtpAlignmentLeft Or xtpAlignmentIconTop
        Else
            .Alignment = xtpAlignmentLeft
        End If
    End With
    Select Case TabId
    Case RibTab_Kat_EinTer:
            Set RpCol = .Add(Kat_IDKurz, "Patient", 100, True)
            If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
                Set RpCol = .Add(Kat_Gruppe, "Mitarbeiter", 110, False)
            Else
                Set RpCol = .Add(Kat_Gruppe, "Mandant", 110, False)
            End If
            Set RpCol = .Add(Kat_Preis1, vbNullString, 0, False)
            RpCol.Alignment = xtpAlignmentCenter
            RpCol.HeaderAlignment = xtpAlignmentCenter
            Set RpCol = .Add(Kat_Sorter, "Datum", 45, True)
    Case RibTab_Kat_KetTer:
            Set RpCol = .Add(Kat_IDKurz, "Terminkette", 200, True)
            Set RpCol = .Add(Kat_Gruppe, "Gruppe", 0, False)
            Set RpCol = .Add(Kat_Preis1, "Zeit", 80, False)
            RpCol.Alignment = xtpAlignmentCenter
            RpCol.HeaderAlignment = xtpAlignmentCenter
            Set RpCol = .Add(Kat_Sorter, vbNullString, 0, False)
    End Select
End With

For Each RpCol In RpCls
    With RpCol
        .Editable = False
        .Resizable = False
        .Sortable = True
        .AutoSortWhenGrouped = False
    End With
Next RpCol

RpCls(Kat_IDKurz).AutoSize = True

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpal " & Err.Number
Resume Next

End Sub
Private Sub FSuAu()
On Error GoTo OrErr
'Hebt die markierten Suchbuchstaben wieder auf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCo7 = Me.repCont7

Select Case RbTab.id
Case RibTab_Kat_EinTer: GlSuN = GlSuX
Case RibTab_Kat_KetTer: GlSuE = GlSuX
End Select

If GlFTE = True Then
    GlFTE = False
    IniSetVal "Layout", "FavoTE", GlFTE
End If

KSuAu "TeDe"
DoEvents

Select Case RbTab.id
Case RibTab_Kat_EinTer: KSuch "TeDe", 0, 2
Case RibTab_Kat_KetTer: KSuch "TeDe", 5, 1
End Select

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
Private Sub FSuch()
On Error GoTo OrErr
'Sucheingabe

Dim SuStr As String
Dim TyIdx As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)

TyIdx = CmSu2.ListIndex

KSuAu "TeDe"

Select Case RbTab.id
Case RibTab_Kat_EinTer:
            SuStr = CmEd2.Text
            With GlSuN
                .SuIdx = TyIdx
                .SuStr = SuStr
            End With
            KSuch "TeDe", 3, 2
            DoEvents
            If RpCo7.Records.Count = 0 Then
                CmEd2.Text = vbNullString
                SPopu "Patient nicht gefunden", "Der von Ihnen gesuchte Patient, konnte nicht gefunden werden", IC48_Forbidden
            Else
                RpCo7.SetFocus
            End If
Case RibTab_Kat_KetTer:
            SuStr = CmEd1.Text
            With GlSuE
                .SuIdx = 1
                .SuStr = SuStr
            End With
            KSuch "TeDe", 5, 1
            DoEvents
            If RpCo7.Records.Count = 0 Then
                CmEd1.Text = vbNullString
                SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
            Else
                RpCo7.SetFocus
            End If
End Select

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuch " & Err.Number
Resume Next

End Sub
Private Sub FSuFa()
On Error GoTo OrErr
'Favoriten Knopf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

KSuAu "TeDe"

If GlFTE = False Then
    CmAcs(KA_Eint_Favoriten).Checked = True
    GlFTE = True
    Select Case RbTab.id
    Case RibTab_Kat_EinTer:
            With GlSuN
                .SuIdx = 5
            End With
    Case RibTab_Kat_KetTer:
            With GlSuE
                .SuIdx = 5
            End With
    End Select
Else
    CmAcs(KA_Eint_Favoriten).Checked = False
    GlFTE = False
    Select Case RbTab.id
    Case RibTab_Kat_EinTer:
            With GlSuN
                .SuIdx = 0
            End With
    Case RibTab_Kat_KetTer:
            With GlSuE
                .SuIdx = 0
            End With
    End Select
End If

IniSetVal "Layout", "FavoTE", GlFTE

Select Case RbTab.id
Case RibTab_Kat_EinTer: KSuch "TeDe", 0, 2
Case RibTab_Kat_KetTer: KSuch "TeDe", 5, 1
End Select

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFa " & Err.Number
Resume Next

End Sub
Private Sub FSuGr()
On Error GoTo OrErr
'Favoriten Knopf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

KSuAu "TeDe"

Select Case RbTab.id
Case RibTab_Kat_EinTer:
        With GlSuN
            .SuIdx = 0
        End With
Case RibTab_Kat_KetTer:
        With GlSuE
            .SuIdx = 0
        End With
End Select

Select Case RbTab.id
Case RibTab_Kat_EinTer: KSuch "TeDe", 0, 2
Case RibTab_Kat_KetTer: KSuch "TeDe", 5, 1
End Select

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
Private Sub FSuLe(ByVal SuStr As String, ByVal TolId As Long)
On Error GoTo OrErr
'ABC Leiste

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

KSuAu "TeDe"

CmAcs(TolId).Checked = True

Select Case RbTab.id
Case RibTab_Kat_EinTer:
        With GlSuN
            .SuIdx = 4
            .SuStr = SuStr
        End With
Case RibTab_Kat_KetTer:
        With GlSuE
            .SuIdx = 4
            .SuStr = SuStr
        End With
End Select

DoEvents
Select Case RbTab.id
Case RibTab_Kat_EinTer: KSuch "TeDe", 0, 2
Case RibTab_Kat_KetTer: KSuch "TeDe", 5, 1
End Select
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
If GlDbg = True Then MsgBox Err.Description, 48, "FSuLe " & Err.Number
Resume Next

End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

TabId = RbTab.id

Screen.MousePointer = vbHourglass

FSpal
DoEvents

Select Case TabId
Case RibTab_Kat_EinTer:
            P_List "TeDe", 0, 2
            CmBrs.Item(2).Visible = True
            CmBrs.Item(3).Visible = False
Case RibTab_Kat_KetTer:
            P_List "TeDe", 5, 1
            CmBrs.Item(2).Visible = False
            CmBrs.Item(3).Visible = True
End Select

DoEvents
FPosi

Screen.MousePointer = vbNormal

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F2:
Case KY_F3:
Case KY_F4:
Case KY_F5:
Case KY_F6:
Case KY_F7:
Case KY_F8: SSave
Case KY_F9:
Case KY_F10:
Case KY_F11: Unload frmMain
Case KA_Hilfe: FHilfe
Case KM_Zeilenumbruch: KGrKa "GrdZei"
Case KM_Zeilenmarker: KGrKa "GrdMkr"
Case KM_Gitternetz: KGrKa "GrdGrl"
Case KM_Multimarker: KGrKa "MulMar"
Case KA_Eint_Einfuegen: FEinf
Case KA_Eint_Favoriten: FSuFa
Case KA_Eint_Vollst: FSuAu
Case KA_Kett_Einfuegen: FEinf
Case AD_Termin_WarNeu: frmAdrSuch.Show vbModal
Case AD_Termin_WarDel: FWaDe
Case AD_Termin_WarErf: FWaNe
Case AD_Termin_WarBea: FWaBe
Case AD_Termin_WarMai:
Case SY_TE_Termin_DocVrs: STeNa 5, True
Case SY_TE_Termin_EmlVrs: STeNa 9, True
Case SY_TE_Termin_SMSVrs: STeNa 13, True
Case KA_SuFe1: FSuch
Case KA_SuFe2: FSuch
Case KA_SuCo1: FSuGr
Case KA_SuCo2: FSuGr
Case 142: FSuLe "Ä", TolId
Case 153: FSuLe "Ö", TolId
Case 154: FSuLe "Ü", TolId
Case Else: If TolId >= 65 And TolId <= 90 Then FSuLe Chr$(TolId), TolId
End Select

GlToo = False

End Sub
Private Sub FWaBe()
On Error GoTo KoErr
'Bearbeitet den markierten Patienten

Dim PatNr As Long
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow
Dim RpSel As XtremeReportControl.ReportSelectedRows
    
Set FM = frmKatTE
Set RpCo7 = FM.repCont7
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Kat_ID0)
        PatNr = RpRow.Record(RpCol.ItemIndex).Value
        GlWaN = True
        AMain PatNr
    End If
End If

Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWaBe " & Err.Number
Resume Next

End Sub

Private Sub FWaDe()
On Error GoTo KoErr
'Entfernt einen Eintrag aus der Warteliste

Dim PatNr As Long
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow
Dim RpSel As XtremeReportControl.ReportSelectedRows
    
Set FM = frmKatTE
Set RpCo7 = FM.repCont7
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Kat_ID0)
        PatNr = RpRow.Record(RpCol.ItemIndex).Value
        Ter_Edi PatNr, False 'aus Warteliste entfernen
        DoEvents
        P_List "TeDe", 0, 2
    End If
End If

Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWaDe " & Err.Number
Resume Next

End Sub
Private Sub FWaNe()
On Error GoTo KoErr
'Erfasst einen neuen Wartenden

GlWaN = True
SAdre 1

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWaDe " & Err.Number
Resume Next

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
Private Sub comBar02_Resize()
    If GlDcP = False Then
        FPosi
    End If
End Sub
Private Sub Form_Load()
    TabId = RibTab_Kat_EinTer
    KMnRp "TeDe"
    FSpal
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatTE = Nothing
End Sub
Private Sub repCont7_KeyUp(KeyCode As Integer, Shift As Integer)

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCo7 = Me.repCont7

If GlAkt = False Then
    If RbTab.id = RibTab_Kat_KetTer Then
        If Shift = 0 Then
            If RpCo7.Records.Count > 0 Then
                Select Case KeyCode
                Case 65: FSuLe "A", KeyCode
                Case 66: FSuLe "B", KeyCode
                Case 67: FSuLe "C", KeyCode
                Case 68: FSuLe "D", KeyCode
                Case 69: FSuLe "E", KeyCode
                Case 70: FSuLe "F", KeyCode
                Case 71: FSuLe "G", KeyCode
                Case 72: FSuLe "H", KeyCode
                Case 73: FSuLe "I", KeyCode
                Case 74: FSuLe "J", KeyCode
                Case 75: FSuLe "K", KeyCode
                Case 76: FSuLe "L", KeyCode
                Case 77: FSuLe "M", KeyCode
                Case 78: FSuLe "N", KeyCode
                Case 79: FSuLe "O", KeyCode
                Case 80: FSuLe "P", KeyCode
                Case 81: FSuLe "Q", KeyCode
                Case 82: FSuLe "R", KeyCode
                Case 83: FSuLe "S", KeyCode
                Case 84: FSuLe "T", KeyCode
                Case 85: FSuLe "U", KeyCode
                Case 86: FSuLe "V", KeyCode
                Case 87: FSuLe "W", KeyCode
                Case 88: FSuLe "X", KeyCode
                Case 89: FSuLe "Y", KeyCode
                Case 90: FSuLe "Z", KeyCode
                Case 132: FSuLe "Ä", KeyCode
                Case 142: FSuLe "Ä", KeyCode
                Case 129: FSuLe "Ü", KeyCode
                Case 154: FSuLe "Ü", KeyCode
                Case 148: FSuLe "Ö", KeyCode
                Case 153: FSuLe "Ö", KeyCode
                End Select
            End If
        End If
    End If
End If
    
Set RpCo7 = Nothing
Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

If RbTab.id = RibTab_Kat_KetTer Then
    FEinf
Else
    FWaBe
End If

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing
    
End Sub
