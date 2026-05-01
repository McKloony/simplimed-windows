VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKetten 
   Caption         =   "Ketten"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   ControlBox      =   0   'False
   Icon            =   "frmKetten.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8550
   Begin XtremeReportControl.ReportControl repCont4 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
      _Version        =   1048579
      _ExtentX        =   3836
      _ExtentY        =   2990
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont5 
      Height          =   1575
      Left            =   4680
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
      _Version        =   1048579
      _ExtentX        =   3625
      _ExtentY        =   2778
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin VB.TextBox txtPatNr 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   500
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   13000
      Width           =   80
   End
   Begin VB.TextBox txtDopFe 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   300
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   13000
      Width           =   80
   End
   Begin VB.TextBox txtIdxNr 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   13000
      Width           =   80
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   720
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
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
Attribute VB_Name = "frmKetten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxIdx As VB.TextBox
Private TxPat As VB.TextBox
Private TxDro As VB.TextBox
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
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4

Private Const MinBr = 2400
Private KeSav As Boolean
Private KyMov As Boolean
Private TabId As Integer

Private clFil As clsFile
Private clFen As clsFenster

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Sub FAktu(Optional ByVal PrSav As Boolean = False)
On Error GoTo OrErr

Dim KetNr As Long
Dim RowNr As Long
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set TxIdx = Me.txtIdxNr
Set RpCo8 = FM.repCont8
Set RpRws = RpCo8.Rows
Set RpSel = RpCo8.SelectedRows

If TxIdx.Text <> vbNullString Then
    KetNr = TxIdx.Text
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            E_Akt KetNr, PrSav
            If PrSav = True Then
                KUpKa RowNr
            End If
        End If
    End If
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpRws = Nothing
Set RpCo8 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAktu " & Err.Number
Resume Next

End Sub

Private Sub FAusw()
On Error GoTo OrErr
'Kettenkennzeichnung

Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, KA_Kett_Auswahl, , True)

LiIdx = CmCom.ListIndex

IniSetVal "System", "KetKen", LiIdx

GlKeK = LiIdx

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAusw " & Err.Number
Resume Next

End Sub
Private Sub FClos()
On Error GoTo OrErr

Set FM = frmKetten

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If KeSav = True Then
    FAktu True
End If

If GlRes = False Then 'Reset der Einstellungen
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "Ketten", "FenLin", clFen.FeLin
        IniSetVal "Ketten", "FenObe", clFen.FeObn
        IniSetVal "Ketten", "FenBre", clFen.FeBre
        IniSetVal "Ketten", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub

Private Sub FDrop(ByVal IdxNr As Long, ByVal KyAdd As Boolean)
On Error GoTo OrErr
'Speichern

Dim KetNr As Long

Set FM = frmKetten
Set TxIdx = FM.txtIdxNr

If TxIdx.Text <> vbNullString Then
    KetNr = TxIdx.Text
    If KyAdd = True Then
        E_KeSa KetNr
        E_Ket KetNr, IdxNr
    Else
        E_KeDe IdxNr
        E_Ket KetNr, IdxNr
    End If
    FAktu True
    KeSav = True
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDrop " & Err.Number
Resume Next

End Sub
Private Sub FEinf()
On Error GoTo OrErr

Dim KetNr As Long
Dim PatNr As Long
Dim KetKe As String
Dim PaStr As String
Dim DroFe As Integer
Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKetten
Set TxIdx = FM.txtIdxNr
Set TxPat = FM.txtPatNr
Set TxDro = FM.txtDopFe
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

Set CmCom = CmBrs.FindControl(CmCom, KA_Kett_Auswahl, , True)

If TxIdx.Text <> vbNullString Then
    If IsNumeric(TxIdx.Text) = True Then
        KetNr = TxIdx.Text
    End If
End If

If TxPat.Text <> vbNullString Then
    If IsNumeric(TxPat.Text) = True Then
        PatNr = TxPat.Text
        If CmSta.Pane(1).Text <> vbNullString Then
            PaStr = CmSta.Pane(1).Text
        End If
    End If
End If

LiIdx = CmCom.ListIndex
DroFe = Val(TxDro.Text)
KetKe = CmCom.Text

If GlBut = RibTab_Abrechnung Or GlBut = RibTab_Krankenbla Or GlBut = RibTab_Rezeptmodul Then
    If LiIdx > 1 Then
        K_Kat1 GlKSt, True, DroFe, GlTag(1), , , KetKe
    Else
        K_Kat1 GlKSt, True, DroFe, GlTag(1)
    End If
Else
    S_TeKe KetNr, PatNr
End If

Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEinf " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TreKy As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TreKy = Left$(GlNod, 1)

If GlBut = RibTab_Kat_Ketten Then
    If GlNod <> vbNullString Then
        Select Case TreKy
        Case "D": 'Gebührenketten
                TeTit = IniGetOpt("Hilfe", 50581)
                TeMai = IniGetOpt("Hilfe", 50582)
                TeInh = IniGetOpt("Hilfe", 50583)
                TeFus = IniGetOpt("Hilfe", 50584)
        Case "F": 'Diagnoseketten
                TeTit = IniGetOpt("Hilfe", 50571)
                TeMai = IniGetOpt("Hilfe", 50572)
                TeInh = IniGetOpt("Hilfe", 50573)
                TeFus = IniGetOpt("Hilfe", 50574)
        Case "H": 'Laborketten
                TeTit = IniGetOpt("Hilfe", 50611)
                TeMai = IniGetOpt("Hilfe", 50612)
                TeInh = IniGetOpt("Hilfe", 50613)
                TeFus = IniGetOpt("Hilfe", 50614)
        Case "J": 'Arzneiketten
                TeTit = IniGetOpt("Hilfe", 50591)
                TeMai = IniGetOpt("Hilfe", 50592)
                TeInh = IniGetOpt("Hilfe", 50593)
                TeFus = IniGetOpt("Hilfe", 50594)
        Case "R": 'Terminketten
                TeTit = IniGetOpt("Hilfe", 50641)
                TeMai = IniGetOpt("Hilfe", 50642)
                TeInh = IniGetOpt("Hilfe", 50643)
                TeFus = IniGetOpt("Hilfe", 50644)
        Case "Q": 'Artikelketten
                TeTit = IniGetOpt("Hilfe", 50601)
                TeMai = IniGetOpt("Hilfe", 50602)
                TeInh = IniGetOpt("Hilfe", 50603)
                TeFus = IniGetOpt("Hilfe", 50604)
        End Select
        SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
    End If
Else
    If GlKSt <> vbNullString Then
        Select Case GlKSt
        Case "DiEi":
            TeTit = IniGetOpt("Hilfe", 50571)
            TeMai = IniGetOpt("Hilfe", 50572)
            TeInh = IniGetOpt("Hilfe", 50573)
            TeFus = IniGetOpt("Hilfe", 50574)
        Case "GbEi":
            TeTit = IniGetOpt("Hilfe", 50581)
            TeMai = IniGetOpt("Hilfe", 50582)
            TeInh = IniGetOpt("Hilfe", 50583)
            TeFus = IniGetOpt("Hilfe", 50584)
        Case "MeEi":
            TeTit = IniGetOpt("Hilfe", 50591)
            TeMai = IniGetOpt("Hilfe", 50592)
            TeInh = IniGetOpt("Hilfe", 50593)
            TeFus = IniGetOpt("Hilfe", 50594)
        Case "ArLi":
            TeTit = IniGetOpt("Hilfe", 50601)
            TeMai = IniGetOpt("Hilfe", 50602)
            TeInh = IniGetOpt("Hilfe", 50603)
            TeFus = IniGetOpt("Hilfe", 50604)
        Case "LaEi":
            TeTit = IniGetOpt("Hilfe", 50611)
            TeMai = IniGetOpt("Hilfe", 50612)
            TeInh = IniGetOpt("Hilfe", 50613)
            TeFus = IniGetOpt("Hilfe", 50614)
        Case "KrDi":
            TeTit = IniGetOpt("Hilfe", 50621)
            TeMai = IniGetOpt("Hilfe", 50622)
            TeInh = IniGetOpt("Hilfe", 50623)
            TeFus = IniGetOpt("Hilfe", 50624)
        Case "KrMe":
            TeTit = IniGetOpt("Hilfe", 50631)
            TeMai = IniGetOpt("Hilfe", 50632)
            TeInh = IniGetOpt("Hilfe", 50633)
            TeFus = IniGetOpt("Hilfe", 50634)
        Case "TeDe":
            TeTit = IniGetOpt("Hilfe", 50641)
            TeMai = IniGetOpt("Hilfe", 50642)
            TeInh = IniGetOpt("Hilfe", 50643)
            TeFus = IniGetOpt("Hilfe", 50644)
        End Select
        SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
    End If
End If

End Sub
Private Sub FKop(ByVal KyAdd As Boolean)
On Error GoTo OrErr
'Fügt einen Eintrag in die Kette ein

Dim KetNr As Long

Set FM = frmKetten
Set TxIdx = FM.txtIdxNr

If TxIdx.Text <> vbNullString Then
    KetNr = TxIdx.Text
    If KyAdd = True Then
        E_Kop KetNr
        E_Ket KetNr
    Else
        E_Del
        E_Ket KetNr
    End If
    KeSav = True
    FAktu True
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKop " & Err.Number
Resume Next

End Sub
Private Sub FLoad()

GlKeL = True
KeSav = False

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 12000
    .ClientMaxWidth = 17000
    .ClientMinHeight = 7000
    .ClientMinWidth = 11000
    .TopMost = True
End With

TabId = RibTab_Ket_Edit

Set FrmEx = Nothing

End Sub
Private Sub FSave()
On Error GoTo OrErr
'Speichern

Dim IdxNr As Long
Dim KetKu As Variant
Dim KetNa As Variant
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmKetten
Set TxIdx = FM.txtIdxNr
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmEdt = CmBrs.FindControl(CmEdt, KA_KeKur, , True)
KetKu = CmEdt.Text
Set CmEdt = CmBrs.FindControl(CmEdt, KA_KeNam, , True)
KetNa = CmEdt.Text

If KetKu = vbNullString Then
    Mld1 = "Sie müssen eine Bezeichnung eingeben, um die Kette speichern zu können"
    Tit1 = "Speichern"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
    Exit Sub
ElseIf KetNa = vbNullString Then
    Mld1 = "Sie müssen eine Ziffer bzw. Suchkürzel eingeben, um die Kette speichern zu können"
    Tit1 = "Speichern"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
    Exit Sub
End If

If CmAcs(KA_Edit_Einfuegen).Enabled = False Then
    If GlKeN = True Then
        E_Sav KetKu, KetNa
    Else
        If TxIdx.Text <> vbNullString Then
            IdxNr = TxIdx.Text
            E_Sav KetKu, KetNa, IdxNr
        End If
    End If
    RpCo4.Enabled = True
    RpCo5.Enabled = True
    CmAcs(KA_Edit_Einfuegen).Enabled = True
    CmAcs(KA_Edit_Entfernen).Enabled = True
    CmAcs(KA_Edit_NachOben).Enabled = True
    CmAcs(KA_Edit_NachUnten).Enabled = True
Else
    If TxIdx.Text <> vbNullString Then
        IdxNr = TxIdx.Text
        E_Sav KetKu, KetNa, IdxNr
    End If
End If

FAktu True
KeSav = False

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSele(ByVal SeTyp As Integer)
On Error GoTo OrErr
'Selektieren & Desleektieren

Dim EiPre As Single
Dim Multi As Single
Dim GePre As Single
Dim TreKy As String
Dim TeMin As Integer
Dim GeMin As Integer
Dim Anzal As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set RpCo5 = FM.repCont5
Set RpRcs = RpCo5.Records
Set CmSta = CmBrs.StatusBar

TreKy = Left$(GlNod, 1)

For Each RpRec In RpRcs
    If TreKy = "R" Then
        TeMin = Left$(RpRec(Ket_Preis).Value, 3)
    Else
        Anzal = RpRec(Ket_Anz).Value
        Multi = Round(RpRec(Ket_Fakto).Value, 2)
        EiPre = Round(RpRec(Ket_Preis).Value, 2)
    End If

    Select Case SeTyp
    Case 1: RpRec(Ket_Selekt).Checked = True
            RpCo5.Populate
            If TreKy = "R" Then
                GeMin = GeMin + TeMin
            Else
                GePre = GePre + (Anzal * Multi * EiPre)
            End If
    Case 2: RpRec(Ket_Selekt).Checked = False
            RpCo5.Populate
            GeMin = 0
            GePre = 0
    Case Else:
            If RpRec(Ket_Selekt).Checked = True Then
                If TreKy = "R" Then
                    GeMin = GeMin + TeMin
                Else
                    GePre = GePre + (Anzal * Multi * EiPre)
                End If
            End If
    End Select
Next RpRec

If TreKy = "R" Then
    CmSta.Pane(2).Text = "Dauer: " & Format$(GeMin, "000" & " Min")
Else
    CmSta.Pane(2).Text = "Gesamt: " & Format$(GePre, GlWa1)
End If

Set RpCo5 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSele " & Err.Number
Resume Next

End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim TreKy As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set CmAcs = CmBrs.Actions
Set RpCls = RpCo5.Columns

TreKy = Left$(GlNod, 1)

TabId = RbTab.id

Select Case TabId
Case RibTab_Ket_Edit:
        RpCo4.Visible = True
        RpCo5.Visible = True
        RpCls(Ket_Selekt).Width = 0
        If TreKy = "R" Then
            RpCls(Ket_Anz).Width = 0
            RpCls(Ket_Fakto).Width = 0
            RpCls(Ket_Zeit).Width = 0
        End If
        CmBrs.Item(2).Visible = True
        CmBrs.Item(3).Visible = True
Case RibTab_Ket_Anwe:
        RpCo4.Visible = False
        RpCo5.Visible = True
        RpCls(Ket_Selekt).Width = 40
        If TreKy = "R" Then
            RpCls(Ket_Anz).Width = 120
            RpCls(Ket_Fakto).Width = 180
            RpCls(Ket_Zeit).Width = 80
        End If
        CmBrs.Item(2).Visible = False
        CmBrs.Item(3).Visible = False
End Select

DoEvents
EPosi

Set RpCo4 = Nothing
Set RpCo5 = Nothing
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
Case KY_F4: EFilt 5
Case KY_F5: frmKaSuch.Show vbModal
Case KY_F6: EFilt 0
Case KY_F7: FEinf
Case KY_F8: FSave
Case KY_F10: EPrint "KetLis", False
Case KY_F11: Unload Me
Case KA_Edit_Hilfe: FHilfe
Case KA_Kett_Speichern: FSave
Case KA_Eint_Favoriten: EFilt 5
Case KA_Eint_Suchen: frmKaSuch.Show vbModal
Case KA_Eint_Vollst: EFilt 0
Case KA_Edit_Einfuegen: FKop True
Case KA_Edit_Entfernen: FKop False
Case KA_Edit_NachOben: EMov True
Case KA_Edit_NachUnten: EMov False
Case KA_Kett_Drucken: EPrint "KetLis", False
Case KA_Kett_Ubernehmen: FEinf
Case KA_Kett_Selekt: FSele 1
Case KA_Kett_Deselekt: FSele 2
Case KA_Kett_Auswahl: FAusw
Case KA_Kett_Patient: frmAdrSuch.Show vbModal
Case AM_Beenden: Unload Me
Case 42: EFilt 0
Case 142: EFilt 4, "Ä"
Case 153: EFilt 4, "Ö"
Case 154: EFilt 4, "Ü"
Case Else: If TolId >= 65 And TolId <= 90 Then EFilt 4, Chr$(TolId)
End Select

GlToo = False

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlKeL = False Then
        If Control.id = XTP_ID_RIBBONCONTROLTAB Then
            FTabu
        Else
            FTool Control.id
        End If
    End If
End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlKeL = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    EPosi
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub Form_Load()
    FLoad
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmKetten = Nothing
End Sub

Private Sub repCont4_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim TreKy As String
Dim KrTyp As Integer
Dim AktZa As Integer

TreKy = Left$(GlNod, 1)

If TreKy = "D" Then 'Gebühren
    If Row.GroupRow = False Then
        If Item.Record(Ket_IDA).Value <> vbNullString Then
            If IsNumeric(Item.Record(Ket_IDA).Value) = True Then
                KrTyp = Item.Record(Ket_IDA).Value
                For AktZa = 1 To UBound(GlKrA)
                    If KrTyp = GlKrA(AktZa, 0) Then
                        Metrics.ForeColor = GlKrA(AktZa, 3)
                        Exit For
                    End If
                Next AktZa
            End If
        End If
    End If
End If

End Sub
Private Sub repCont4_KeyUp(KeyCode As Integer, Shift As Integer)

Dim RpCo4 As XtremeReportControl.ReportControl

Set RpCo4 = Me.repCont4

If RpCo4.Records.Count > 0 Then
    Set RpSel = RpCo4.SelectedRows
    If RpSel.Count > 0 Then
        If KeyCode >= 65 And KeyCode <= 90 Then
            EFilt 4, Chr$(KeyCode)
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo4 = Nothing

End Sub
Private Sub repCont4_RecordsDropped(ByVal TargetRecord As XtremeReportControl.IReportRecord, ByVal Records As XtremeReportControl.IReportRecords, ByVal Above As Boolean)
    If Records.Count > 0 Then
        If Records(0).Item(Ket_IDA).Value <> vbNullString Then
            If Records(0).Item(Ket_IDA).Value > 0 Then
                FDrop Records(0).Item(Ket_IDA).Value, False
            End If
        End If
    End If
End Sub

Private Sub repCont5_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim TreKy As String
Dim KrTyp As Integer
Dim AktZa As Integer

TreKy = Left$(GlNod, 1)

If TreKy = "D" Then 'Gebühren
    If Row.GroupRow = False Then
        If Item.Record(Ket_Zeit).Value <> vbNullString Then
            If IsNumeric(Item.Record(Ket_Zeit).Value) = True Then
                KrTyp = Item.Record(Ket_Zeit).Value
                For AktZa = 1 To UBound(GlKrA)
                    If KrTyp = GlKrA(AktZa, 0) Then
                        Metrics.ForeColor = GlKrA(AktZa, 3)
                        Exit For
                    End If
                Next AktZa
            End If
        End If
    End If
End If

End Sub
Private Sub repCont5_BeginDrag(ByVal Records As XtremeReportControl.IReportRecords)
    If Records.Count > 0 Then
        If Records(0).Item(Ket_ID0).Value > 0 Then
            KyMov = True
        End If
    End If
End Sub
Private Sub repCont5_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Item.Index = Ket_Selekt Then
        FSele 0
    End If
End Sub
Private Sub repCont5_RecordsDropped(ByVal TargetRecord As XtremeReportControl.IReportRecord, ByVal Records As XtremeReportControl.IReportRecords, ByVal Above As Boolean)
On Error Resume Next

Dim KetNr As Long
Dim IdxNr As Long

Set FM = frmKetten
Set TxIdx = FM.txtIdxNr

Screen.MousePointer = vbHourglass

If TxIdx.Text <> vbNullString Then
    KetNr = TxIdx.Text
    If Records.Count > 0 Then
        IdxNr = Records(0).Item(Ket_ID0).Value
        If KyMov = True Then
            E_Sor
            DoEvents
            E_Ket KetNr, IdxNr
        Else
            FDrop IdxNr, True
        End If
        KyMov = False
    End If
End If

Screen.MousePointer = vbNormal

End Sub
Private Sub repCont5_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error GoTo OrErr

Dim KetNr As Long
Dim IdxNr As Long
Dim TmTag As String
Dim TmpTg As String

Set TxIdx = Me.txtIdxNr

KetNr = TxIdx.Text
TmpTg = Item.Tag
TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
IdxNr = Item.Record(Ket_ID0).Value

If TabId = RibTab_Ket_Edit Then
    Item.Tag = "@" & TmTag
    E_Anz KetNr
    DoEvents
    E_Ket KetNr, IdxNr
    DoEvents
    Item.Tag = TmpTg
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAnza " & Err.Number
Resume Next

End Sub
