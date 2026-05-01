VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatAE 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   1815
      Left            =   1200
      TabIndex        =   0
      Top             =   3600
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
      TabIndex        =   1
      Top             =   840
      Width           =   2775
      _Version        =   1048579
      _ExtentX        =   4895
      _ExtentY        =   4260
      _StockProps     =   64
      Show3DBorder    =   0
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   600
      Top             =   240
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKatAE"
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

Private TabId As Integer
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50301)
TeMai = IniGetOpt("Hilfe", 50302)
TeInh = IniGetOpt("Hilfe", 50303)
TeFus = IniGetOpt("Hilfe", 50304)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

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
Case KA_Hilfe: FHilfe
Case KA_KaBu1: FDaFo
Case KM_Gruppierung: KGrKa "GrdGrp"
Case KM_Zeilenumbruch: KGrKa "GrdZei"
Case KM_Zeilenmarker: KGrKa "GrdMkr"
Case KM_Gitternetz: KGrKa "GrdGrl"
Case KM_Multimarker: KGrKa "MulMar"
Case KM_Popupkalender: KGrKa "PopKal"
Case KA_Eint_Einfuegen: FEinf
Case KA_Eint_Favoriten: FSuFa
Case KA_Eint_Vollst: FSuAu
Case KA_SuFe1: FSuch
Case KA_SuCo1: FSuGr
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

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)

TabId = RbTab.id

Select Case TabId 'Fragebogengruppen
Case RibTab_Kat_EinAna: P_List "AnEi", GlAnG(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetAna:
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
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)

KSuAu "AnEi"

CmAcs(TolId).Checked = True

Select Case TabId
Case RibTab_Kat_EinAna:
        With GlSuE
            .SuIdx = 4
            .SuStr = SuStr
        End With
Case RibTab_Kat_KetAna:
        With GlSuN
            .SuIdx = 4
            .SuStr = SuStr
        End With
End Select

DoEvents
Select Case TabId 'Anamnesegruppen
Case RibTab_Kat_EinAna: KSuch "AnEi", GlAnG(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetAna:
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
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)

KSuAu "AnEi"

Select Case TabId
Case RibTab_Kat_EinAna:
        With GlSuE
            .SuIdx = 0
        End With
Case RibTab_Kat_KetAna:
        With GlSuN
            .SuIdx = 0
        End With
End Select

Select Case TabId 'Anamnesegruppen
Case RibTab_Kat_EinAna: KSuch "AnEi", GlAnG(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetAna:
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
Private Sub FSuFa()
On Error GoTo OrErr
'Favoriten Knopf

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)

KSuAu "AnEi"

If GlFAE = False Then
    CmAcs(KA_Eint_Favoriten).Checked = True
    GlFAE = True
    Select Case TabId
    Case RibTab_Kat_EinAna:
            With GlSuE
                .SuIdx = 5
            End With
    Case RibTab_Kat_KetAna:
            With GlSuN
                .SuIdx = 5
            End With
    End Select
Else
    CmAcs(KA_Eint_Favoriten).Checked = False
    GlFAE = False
    Select Case TabId
    Case RibTab_Kat_EinAna:
            With GlSuE
                .SuIdx = 0
            End With
    Case RibTab_Kat_KetAna:
            With GlSuN
                .SuIdx = 0
            End With
    End Select
End If

IniSetVal "Layout", "FavoAE", GlFAE

Select Case TabId 'Anamnesegruppen
Case RibTab_Kat_EinAna: KSuch "AnEi", GlAnG(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetAna:
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
Private Sub FSuch()
On Error GoTo OrErr
'Sucheingabe

Dim SuStr As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmEd1 As XtremeCommandBars.CommandBarEdit

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)

KSuAu "AnEi"

Select Case TabId
Case RibTab_Kat_EinAna:
            SuStr = CmEd1.Text
            With GlSuE
                .SuIdx = 1
                .SuStr = SuStr
            End With
            KSuch "AnEi", GlAnG(CmSu1.ListIndex, 0), 1
            DoEvents
            If RpCo7.Records.Count = 0 Then
                CmEd1.Text = vbNullString
                SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
            Else
                RpCo7.SetFocus
            End If
Case RibTab_Kat_KetAna:
            
End Select

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
Dim RpCo7 As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set RpCo7 = Me.repCont7

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)

Select Case TabId
Case RibTab_Kat_EinAna: GlSuE = GlSuX
Case RibTab_Kat_KetAna: GlSuN = GlSuX
End Select

If GlFAE = True Then
    GlFAE = False
    IniSetVal "Layout", "FavoAE", GlFAE
End If

KSuAu "AnEi"
DoEvents

Select Case TabId 'Anamnesegruppen
Case RibTab_Kat_EinAna: KSuch "AnEi", GlAnG(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetAna:
End Select

DoEvents
RpCo7.SetFocus

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuAu " & Err.Number
Resume Next

End Sub
Public Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition Anamnese ein

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(Kat_ID0, "ID0", 0, False)
    Set RpCol = .Add(Kat_GOID, "Nummer", 80, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        If RpCo7.PaintManager.FixedRowHeight = False Then
            .Alignment = xtpAlignmentLeft Or xtpAlignmentIconTop
        Else
            .Alignment = xtpAlignmentLeft
        End If
    End With
    Set RpCol = .Add(Kat_IDKurz, "Anamnesetext", 400, False)
    If RpCo7.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Kat_Gruppe, "Gruppe", 0, False)
    Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
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
Private Sub FEinf(Optional ByVal DroFe As Integer)

Dim DayFi As Date
Dim DayLa As Date
Dim EiDat As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit

Set MoKal = Me.dtpDatu1
Set CmBrs = Me.comBar02

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

Select Case TabId
Case RibTab_Kat_EinAna:
        If GlPoK = True Then 'Popupkalender
            K_Kat1 "AnEi", , DroFe, EiDat
        Else
            K_Kat1 "AnEi", , DroFe, GlTag(1)
        End If
Case RibTab_Kat_KetAna:

End Select
S_AbTe DayFi, DayLa

Set MoKal = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim AnzBl As Long
Dim AktBl As Long
Dim AktTa As Long
Dim BloTa As Long

Set MoKal = Me.dtpDatu1
Set TxDa1 = frmMain.txtDatu1

AktTa = 0
AnzBl = MoKal.Selection.BlocksCount

If AnzBl = 0 Then
    ReDim Preserve GlTag(1)
    GlTag(1) = Date
ElseIf AnzBl = 1 Then
    DaBeg = MoKal.Selection(0).DateBegin
    DaEnd = MoKal.Selection(0).DateEnd
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
Dim AnzTa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim DaPi4 As XtremeCalendarControl.DatePicker

Set CmBrs = Me.comBar02
Set MoKal = Me.dtpDatu1

Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)

CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
CmEdt.GetRect ItmLi, ItmOb, ItmRe, ItmHo

If IsDate(CmEdt.Text) Then
    NeuDa = CmEdt.Text
Else
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
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatAE = Nothing
End Sub
Private Sub Form_Load()
    KMnRp "AnEi", True
    TabId = RibTab_Kat_EinAna
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

Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FEinf
End Sub

