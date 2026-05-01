VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatME 
   BorderStyle     =   0  'Kein
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   1815
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   3201
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
      _Version        =   1048579
      _ExtentX        =   4895
      _ExtentY        =   4260
      _StockProps     =   64
      Show3DBorder    =   0
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   360
      Top             =   360
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKatME"
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
Private MoKal As XtremeCalendarControl.DatePicker
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow

Private TabId As Integer
Private Sub FDaFo()
On Error GoTo OrErr
'Läßt den Kalender aufklappen

Dim ItmLi As Long
Dim ItmOb As Long
Dim ItmRe As Long
Dim ItmHo As Long
Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim NeuDa As Date
Dim DayFi As Date
Dim DayLa As Date
Dim DaChk As Date
Dim AnzTa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim DaPi4 As XtremeCalendarControl.DatePicker

Set CmBrs = Me.comBar02
Set MoKal = Me.dtpDatu1

DaChk = DateAdd("yyyy", -10, Date)

Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)

CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
CmEdt.GetRect ItmLi, ItmOb, ItmRe, ItmHo

If IsDate(CmEdt.Text) Then
    NeuDa = CmEdt.Text
Else
    NeuDa = Date
End If

If NeuDa < DaChk Then
    NeuDa = Date
End If

DayFi = NeuDa - 30
DayLa = NeuDa + 30

With MoKal
    .RedrawControl
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    S_AbTe DayFi, DayLa
    .Left = ItmLi
    .Top = ClObn
    If .ShowModal(2, 1) Then
        If .Selection.BlocksCount > 0 Then
            CmEdt.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set MoKal = Nothing
Set CmEdt = Nothing
Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaFo " & Err.Number
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
Case KY_F5: frmAdrSuch.Show vbModal
Case KY_F6:
Case KY_F7:
Case KY_F8: SSave
Case KY_F9:
Case KY_F10: SDrLis 2
Case KY_F11: Unload frmMain
Case KA_KaBu1: FDaFo
Case KA_Hilfe: FHilfe
Case KM_Zeilenumbruch: KGrKa "GrdZei"
Case KM_Zeilenmarker: KGrKa "GrdMkr"
Case KM_Gitternetz: KGrKa "GrdGrl"
Case KM_Multimarker: KGrKa "MulMar"
Case KM_Popupkalender: KGrKa "PopKal"
Case KA_Eint_Einfuegen: FEinf
Case KA_Kett_Einfuegen: FEinf
Case KA_Eint_Favoriten: FSuFa
Case KA_Eint_Vollst: FSuAu
Case KA_Kett_Vollst: FSuAu
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
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

TabId = RbTab.id

Select Case TabId
Case RibTab_Kat_EinMed: P_List "MeEi", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetMed: P_List "MeEi", 1, 2
End Select

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FSuLe(ByVal SuStr As String, ByVal TolId As Long)
On Error GoTo OrErr
'ABC Leiste

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

KSuAu "MeEi"

CmAcs(TolId).Checked = True

Select Case TabId
Case RibTab_Kat_EinMed:
        With GlSuE
            .SuIdx = 4
            .SuStr = SuStr
        End With
Case RibTab_Kat_KetMed:
        With GlSuN
            .SuIdx = 4
            .SuStr = SuStr
        End With
End Select

DoEvents
Select Case TabId
Case RibTab_Kat_EinMed: KSuch "MeEi", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetMed: KSuch "MeEi", 1, 2
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
Private Sub FSuGr()
On Error GoTo OrErr
'Favoriten Knopf

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

Screen.MousePointer = vbHourglass

KSuAu "MeEi"

Select Case TabId
Case RibTab_Kat_EinMed:
        With GlSuE
            .SuIdx = 0
        End With
Case RibTab_Kat_KetMed:
        With GlSuN
            .SuIdx = 0
        End With
End Select

Select Case TabId
Case RibTab_Kat_EinMed: KSuch "MeEi", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetMed: KSuch "MeEi", 1, 2
End Select

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Screen.MousePointer = vbNormal

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuGr " & Err.Number
Resume Next

End Sub
Private Sub FSuFa()
On Error GoTo OrErr
'Favoriten Knopf

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

Screen.MousePointer = vbHourglass

KSuAu "MeEi"

If GlFME = False Then
    CmAcs(KA_Eint_Favoriten).Checked = True
    GlFME = True
    Select Case TabId
    Case RibTab_Kat_EinMed:
            With GlSuE
                .SuIdx = 5
            End With
    Case RibTab_Kat_KetMed:
            With GlSuN
                .SuIdx = 5
            End With
    End Select
Else
    CmAcs(KA_Eint_Favoriten).Checked = False
    GlFME = False
    Select Case TabId
    Case RibTab_Kat_EinMed:
            With GlSuE
                .SuIdx = 0
            End With
    Case RibTab_Kat_KetMed:
            With GlSuN
                .SuIdx = 0
            End With
    End Select
End If

IniSetVal "Layout", "FavoME", GlFME

Select Case TabId
Case RibTab_Kat_EinMed: KSuch "MeEi", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetMed: KSuch "MeEi", 1, 2
End Select

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Screen.MousePointer = vbNormal

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFa " & Err.Number
Resume Next

End Sub
Private Sub FSuch()
On Error GoTo OrErr
'Sucheingabe

Dim SuStr As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)

Screen.MousePointer = vbHourglass

KSuAu "MeEi"

Select Case TabId
Case RibTab_Kat_EinMed:
            SuStr = CmEd1.Text
            With GlSuE
                .SuIdx = 1
                .SuStr = SuStr
            End With
            KSuch "MeEi", GlMed(CmSu1.ListIndex, 0), 1
            DoEvents
            If RpCo7.Records.Count = 0 Then
                CmEd1.Text = vbNullString
                SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
            Else
                RpCo7.SetFocus
            End If
Case RibTab_Kat_KetMed:
            SuStr = CmEd2.Text
            With GlSuN
                .SuIdx = 1
                .SuStr = SuStr
            End With
            KSuch "MeEi", 1, 2
            DoEvents
            If RpCo7.Records.Count = 0 Then
                CmEd1.Text = vbNullString
                SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
            Else
                RpCo7.SetFocus
            End If
End Select

Screen.MousePointer = vbNormal

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

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set RpCo7 = Me.repCont7

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

Screen.MousePointer = vbHourglass

Select Case TabId
Case RibTab_Kat_EinMed: GlSuE = GlSuX
Case RibTab_Kat_KetMed: GlSuN = GlSuX
End Select

If GlFKM = True Then
    GlFKM = False
    IniSetVal "Layout", "FavoKM", GlFKM
End If

KSuAu "MeEi"
DoEvents

Select Case TabId
Case RibTab_Kat_EinMed: KSuch "MeEi", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetMed: KSuch "MeEi", 1, 2
End Select

Screen.MousePointer = vbNormal

DoEvents
RpCo7.SetFocus

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuAu " & Err.Number
Resume Next

End Sub
Private Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition Medikamente ein

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(Kat_ID0, "ID0", 0, False)
    Set RpCol = .Add(Kat_GOID, "PZN", 80, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        If RpCo7.PaintManager.FixedRowHeight = False Then
            .Alignment = xtpAlignmentLeft Or xtpAlignmentIconTop
        Else
            .Alignment = xtpAlignmentLeft
        End If
    End With
    Set RpCol = .Add(Kat_IDKurz, "Heilmitteltext", 400, False)
    If RpCo7.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Kat_Gruppe, "Gruppe", 0, False)
    Set RpCol = .Add(Kat_Preis1, "Preis", 60, False)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Kat_Sorter, "Sorter", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Resizable = False
    RpCol.Sortable = True
    RpCol.AutoSortWhenGrouped = False
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
Set MoKal = Me.dtpDatu1

CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
ClHoh = ClHoh - ClObn

If GlPoK = True Then 'Popupkalender
    RpCon.Move ClLin, ClObn, ClBre - ClLin, ClHoh
Else
    MoKal.Move ClLin, ClObn, ClBre - ClLin, 2200
    If ClHoh - 2200 > 0 Then
        RpCon.Move ClLin, ClObn + 2200, ClBre - ClLin, ClHoh - 2200
    End If
End If

Set MoKal = Nothing
Set RpCon = Nothing
Set CmBrs = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FEinf()
On Error Resume Next

Dim IdxNr As Long
Dim DayFi As Date
Dim DayLa As Date
Dim EiDat As Date
Dim KetNa As String
Dim KetKu As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim CmEdt As XtremeCommandBars.CommandBarEdit

Set MoKal = Me.dtpDatu1
Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)

With MoKal
    DayFi = .FirstDayOfWeek
    DayLa = .LastVisibleDay
End With

If IsDate(CmEdt.Text) = True Then
    EiDat = CDate(CmEdt.Text)
Else
    EiDat = Date
End If

Select Case GlBut
Case RibTab_Abrechnung:
    Select Case TabId
    Case RibTab_Kat_EinMed:
            If GlPoK = True Then 'Popupkalender
                K_Kat1 "MeEi", , , EiDat
            Else
                K_Kat1 "MeEi", , , GlTag(1)
            End If
    Case RibTab_Kat_KetMed:
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Kat_ID0)
                IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Kat_GOID)
                KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Kat_IDKurz)
                KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                GlNod = "J1"
                GlKSt = "MeEi"
                GlKeE = True
                EMain IdxNr, KetNa, KetKu
            End If
        End If
    End Select
Case RibTab_Rezeptmodul:
    Select Case TabId
    Case RibTab_Kat_EinMed: K_RzEi "MeEi"
    Case RibTab_Kat_KetMed: K_RzEi "MeEi", True
    End Select
End Select

S_AbTe DayFi, DayLa

Set MoKal = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case TabId
Case RibTab_Kat_EinMed:
    TeTit = IniGetOpt("Hilfe", 50481)
    TeMai = IniGetOpt("Hilfe", 50482)
    TeInh = IniGetOpt("Hilfe", 50483)
    TeFus = IniGetOpt("Hilfe", 50484)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
Case RibTab_Kat_KetMed:
    If WindowLoad("frmKetten") = False Then
        TeTit = IniGetOpt("Hilfe", 50491)
        TeMai = IniGetOpt("Hilfe", 50492)
        TeInh = IniGetOpt("Hilfe", 50493)
        TeFus = IniGetOpt("Hilfe", 50494)
        SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
    End If
End Select

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim DaChk As Date
Dim AnzBl As Long
Dim AktBl As Long
Dim AktTa As Long
Dim BloTa As Long

Set MoKal = Me.dtpDatu1
Set TxDa1 = frmMain.txtDatu1

AktTa = 0
AnzBl = MoKal.Selection.BlocksCount

DaChk = DateAdd("yyyy", -10, Date)

If AnzBl = 0 Then
    ReDim Preserve GlTag(1)
    GlTag(1) = Date
ElseIf AnzBl = 1 Then
    DaBeg = MoKal.Selection(0).DateBegin
    DaEnd = MoKal.Selection(0).DateEnd
    If DaBeg < DaChk Then
        DaBeg = Date
        DaEnd = Date
        With MoKal
            .EnsureVisible DaBeg - 30
            .Select DaBeg
            .SelectRange DaBeg, DaEnd
        End With
    End If
    If DaEnd > DaBeg Then
        Do
        DaAkt = DaBeg + AktTa
        AktTa = AktTa + 1
        ReDim Preserve GlTag(AktTa)
        GlTag(AktTa) = DaAkt
        Loop Until DaAkt >= DaEnd
    Else
        ReDim Preserve GlTag(1)
        GlTag(1) = DaBeg
    End If
ElseIf AnzBl > 1 Then
    For AktBl = 0 To AnzBl - 1
        DaBeg = MoKal.Selection.Blocks(AktBl).DateBegin
        DaEnd = MoKal.Selection.Blocks(AktBl).DateEnd
        If DaBeg < DaChk Then
            DaBeg = Date
            DaEnd = Date
            With MoKal
                .EnsureVisible DaBeg - 30
                .Select DaBeg
                .SelectRange DaBeg, DaEnd
            End With
        End If
        If DaEnd > DaBeg Then
            BloTa = 0
            Do
            DaAkt = DaBeg + BloTa
            AktTa = AktTa + 1
            BloTa = BloTa + 1
            ReDim Preserve GlTag(AktTa)
            GlTag(AktTa) = DaAkt
            Loop Until DaAkt >= DaEnd
        Else
            AktTa = AktTa + 1
            ReDim Preserve GlTag(AktTa)
            GlTag(AktTa) = DaBeg
        End If
    Next AktBl
End If

TxDa1.Text = GlTag(1)

If GlPop = True Then
    S_AbTa
    S_AbDo
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatME = Nothing
End Sub
Private Sub Form_Load()
    KMnRp "MeEi", True
    TabId = RibTab_Kat_EinMed
    FSpal
End Sub
Private Sub dtpDatu1_SelectionChanged()
    If GlDcP = False Then
        FDatu
    End If
End Sub

Private Sub dtpDatu1_MonthChanged()

Dim DayFi As Date
Dim DayLa As Date

Set MoKal = Me.dtpDatu1

With MoKal
    DayFi = .FirstDayOfWeek
    DayLa = .LastVisibleDay
End With

If GlDcP = False Then
    S_AbTe DayFi, DayLa
End If

Set MoKal = Nothing

End Sub

Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
On Error Resume Next

Dim AktTa As Long

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTaV > 0 Then
    For AktTa = 1 To GlTaV
        If Day = GlBet(AktTa) Then
            Metrics.BackColor = GlMkr
        End If
    Next AktTa
End If

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

Private Sub repCont7_KeyUp(KeyCode As Integer, Shift As Integer)

Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7

If GlAkt = False Then
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
    
Set RpCo7 = Nothing

End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FEinf
End Sub
