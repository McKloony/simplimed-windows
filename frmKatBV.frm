VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatBV 
   BorderStyle     =   0  'Kein
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5880
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
Attribute VB_Name = "frmKatBV"
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

Private TabId As Integer
Private Sub FEdit()
On Error GoTo AnErr

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

If RpSel.Count > 0 Then
    Select Case TabId
    Case RibTab_Kat_EinBuc: frmBaEdVo.Show
    Case RibTab_Kat_KetBuc: frmBaEdRe.Show
    End Select
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo7 = Nothing

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

Select Case TabId
Case RibTab_Kat_EinBuc:
    TeTit = IniGetOpt("Hilfe", 50081)
    TeMai = IniGetOpt("Hilfe", 50082)
    TeInh = IniGetOpt("Hilfe", 50083)
    TeFus = IniGetOpt("Hilfe", 50084)
Case RibTab_Kat_KetBuc:
    TeTit = IniGetOpt("Hilfe", 50091)
    TeMai = IniGetOpt("Hilfe", 50092)
    TeInh = IniGetOpt("Hilfe", 50093)
    TeFus = IniGetOpt("Hilfe", 50094)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FLoe()
On Error GoTo AnErr

Select Case TabId
Case RibTab_Kat_EinBuc: K_BuLo "BaVo"
Case RibTab_Kat_KetBuc: K_BaLo
End Select

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoe " & Err.Number
Resume Next

End Sub

Private Sub FNeu()
On Error GoTo AnErr

GlNeB = True 'neue Buchung

Select Case TabId
Case RibTab_Kat_EinBuc: frmBaEdVo.Show
Case RibTab_Kat_KetBuc: frmBaEdRe.Show
End Select

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeu " & Err.Number
Resume Next

End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

TabId = RbTab.id

Screen.MousePointer = vbHourglass

FSpal
DoEvents

K_BuVpl "BaVo"
DoEvents

Select Case TabId
Case RibTab_Kat_EinBuc:
        P_List "BaVo", 0, 1
Case RibTab_Kat_KetBuc:
        P_List "BaVo", 0, 2
End Select

Screen.MousePointer = vbNormal

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
'Formratieren der Spalten

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
    Set RpCol = .Add(Buh_ID0, "ID0", 0, False)
    Set RpCol = .Add(Buh_Datum, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Buh_Buchtext, "Buchungstext", 0, True)
    Set RpCol = .Add(Buh_Einnahme, "Betrag", 0, True)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Buh_Ausgabe, "Brutto", 0, True)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    RpCol.Alignment = xtpAlignmentRight
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        Set RpCol = .Add(Buh_Sachkonto, "Sachkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Geldkonto", 0, True)
    Else
        Set RpCol = .Add(Buh_Sachkonto, "Sachkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Geldkonto", 0, True)
    End If
    Set RpCol = .Add(Buh_RechNr, "Belegzeichen", 0, True)
    Set RpCol = .Add(Buh_IDR, "IDR", 0, False)
    Set RpCol = .Add(Buh_Beleg, "Nummer", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Buh_Sachkontenbez, "Sachkontenbezeichnung", 0, True)
    Select Case TabId
    Case RibTab_Kat_EinBuc: Set RpCol = .Add(Buh_Geldkontenbez, "Geldkontenbezeichnung", 0, True)
    Case RibTab_Kat_KetBuc: Set RpCol = .Add(Buh_Geldkontenbez, "Sorter", 0, True)
    End Select
    Set RpCol = .Add(Buh_Steuer, "Steuer", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Buh_W, "W", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Buh_Privat, "Privat", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_Abziehbar, "Abziehbar", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDB, "IDB", 0, False)
    Set RpCol = .Add(Buh_IDA, "IDA", 0, False)
    Set RpCol = .Add(Buh_Währung, "Währung", 0, False)
    Set RpCol = .Add(Buh_Ermittlung, "KE", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Select Case TabId
    Case RibTab_Kat_EinBuc: Set RpCol = .Add(Buh_Dokument, "DK", 0, False)
    Case RibTab_Kat_KetBuc: Set RpCol = .Add(Buh_Dokument, "Anz.", 0, False)
    End Select
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Buh_IDP, "IDP", 0, False)
    Set RpCol = .Add(Buh_IDArt, "IDArt", 0, False)
    Set RpCol = .Add(Buh_IDBank, "IDBank", 0, False)
    Set RpCol = .Add(Buh_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Buh_IDT, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Buh_Berichtdatum, "Berichtdatum", 0, True)
    Set RpCol = .Add(Buh_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Buh_Monat, "Monat", 0, False)
    Set RpCol = .Add(Buh_Storniert, "Storniert", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Buh_Zuordnung, "ZU", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_User_Norm
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Lock, "Lock", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconLeft
        .Icon = IC16_Lock
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Datei, "Datei", 0, False)
    Set RpCol = .Add(Buh_Doppelt, "DO", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
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

If GlTFt.SIZE > 10 Then
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 0
    RpCls(Buh_Buchtext).Width = 200
    RpCls(Buh_Einnahme).Width = 0
    RpCls(Buh_Ausgabe).Width = 0
Else
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 0
    RpCls(Buh_Buchtext).Width = 180
    RpCls(Buh_Einnahme).Width = 0
    RpCls(Buh_Ausgabe).Width = 0
End If
RpCls(Buh_Sachkonto).Width = 80
RpCls(Buh_Gegenkonto).Width = 0
RpCls(Buh_RechNr).Width = 0
RpCls(Buh_IDR).Width = 0
RpCls(Buh_Beleg).Width = 0
RpCls(Buh_Sachkontenbez).Width = 180
Select Case TabId
Case RibTab_Kat_EinBuc: RpCls(Buh_Geldkontenbez).Width = 0
Case RibTab_Kat_KetBuc: RpCls(Buh_Geldkontenbez).Width = 60
End Select
RpCls(Buh_Steuer).Width = 75
RpCls(Buh_W).Width = 40
RpCls(Buh_Privat).Width = 0
RpCls(Buh_Abziehbar).Width = 0
RpCls(Buh_IDB).Width = 0
RpCls(Buh_IDA).Width = 0
RpCls(Buh_Währung).Width = 0
RpCls(Buh_Ermittlung).Width = 25
Select Case TabId
Case RibTab_Kat_EinBuc: RpCls(Buh_Dokument).Width = 0
Case RibTab_Kat_KetBuc: RpCls(Buh_Dokument).Width = 50
End Select
RpCls(Buh_IDP).Width = 0
RpCls(Buh_IDArt).Width = 0
RpCls(Buh_IDBank).Width = 0
RpCls(Buh_Kommentar).Width = 0
RpCls(Buh_IDT).Width = 150
RpCls(Buh_Berichtdatum).Width = 0
RpCls(Buh_GuiID).Width = 0
RpCls(Buh_Monat).Width = 0
RpCls(Buh_Storniert).Width = 0
RpCls(Buh_IDM).Width = 150
RpCls(Buh_Zuordnung).Width = 18
RpCls(Buh_Lock).Width = 18
RpCls(Buh_Datei).Width = 0
RpCls(Buh_Doppelt).Width = 0

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
Private Sub repCont7_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim GeKto As Long

If Row.Record(Buh_IDB).Value <> vbNullString Then
    If Row.Record(Buh_IDB).Value > 0 Then
        GeKto = Row.Record(Buh_IDB).Value
    Else
        GeKto = 0
    End If
Else
    GeKto = 0
End If
If GeKto > 0 Then
    If GeKto <= UBound(GlGeK) Then
        If CBool(GlGeK(Row.Record(Buh_IDB).Value, 5)) = True Then
            Metrics.ForeColor = 16711680
        End If
    End If
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatBV = Nothing
End Sub
Private Sub Form_Load()
    KMnRp "BaVo"
    TabId = RibTab_Kat_EinBuc
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
Case KA_Eint_Einfuegen: S_BaKon
Case KA_Eint_Hinzufuegen: FNeu
Case KA_Eint_Bearbeiten: FEdit
Case KA_Eint_Loeschen: FLoe
Case KA_Kett_Suchen: S_BaZuo
Case KA_Kett_Hinzufuegen: FNeu
Case KA_Kett_Bearbeiten: FEdit
Case KA_Kett_Loeschen: FLoe
End Select

GlToo = False

End Sub

Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FEdit
End Sub
