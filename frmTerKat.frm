VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmTerKat 
   Caption         =   "Kataloge"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   ControlBox      =   0   'False
   Icon            =   "frmTerKat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont9 
      Height          =   1575
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
      _Version        =   1048579
      _ExtentX        =   4471
      _ExtentY        =   2778
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   5000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars comBar03 
      Left            =   600
      Top             =   480
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTerKat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private Kale1 As XtremeCalendarControl.DatePicker
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
Private WithEvents TbBar As XtremeCommandBars.TabToolBar
Attribute TbBar.VB_VarHelpID = -1
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Const CB_SHOWDROPDOWN = &H14F
Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4

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
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = "Einträge Einfügen"
TeMai = "Einträge auswählen, um diese in den Termin einzutragen"
TeInh = "Mit Hilfe dieser Funktion ist es möglich, Einträge oder Ketten, seien es Gebührenleistungen, Diagnosen oder Gebührenketten, in das Krankenblatt des Termins einzufügen. Diese Einträge werden dann später im Rahmen der Funktion: Rechnung Erstellen, in die Rechnung des Patienten übertragen."
TeFus = "Über das Auswahlfeld kann der Katalogtyp gewählt und über das Suchfeld einzelne Einträge oder Ketten gesucht werden. Beim Einfügen von Gebührenketten werden immer die Preise des zugrunde liegenden Kataloges und nicht die angepasste Preise in der Kette verwendet."

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FSum()
On Error GoTo OrErr
'Summiert die markierten Einträge

Dim GeSum As Single
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTerKat
Set RpCon = FM.repCont9
Set CmBrs = FM.comBar03
Set CmSta = CmBrs.StatusBar
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

If RpSel.Count > 0 Then
    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Kat_Preis1)
            GeSum = GeSum + RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Kat_IDKurz)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                CmSta.Pane(1).Text = RpRow.Record(RpCol.ItemIndex).Value
            End If
        End If
    Next RpRow
    CmSta.Pane(1).Text = "Gesamt: " & Format$(GeSum, GlWa1)
End If

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpSel = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSum " & Err.Number
Resume Next

End Sub
Private Sub FSuch()
On Error GoTo MeErr

Dim SuStr As Variant
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim RpCon As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmTerKat
Set RpCon = FM.repCont9
Set CmBrs = FM.comBar03
Set RpRws = RpCon.Rows
Set CmAcs = CmBrs.Actions

Set CmEdt = CmBrs.FindControl(CmEdt, KA_SuFe1, , True)
SuStr = CmEdt.Text

If SuStr <> vbNullString Then
    TrSuch 1, SuStr
Else
    TrSuch 5
End If

CmEdt.Text = vbNullString

If RpRws.Count > 0 Then
    RpCon.SetFocus
    CmAcs(KA_KaBu2).Checked = True
End If

Set CmBrs = Nothing
Set RpCon = Nothing
Set RpRws = Nothing

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuch " & Err.Number
Resume Next

End Sub
Private Sub FClos()
On Error GoTo PoErr

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlIdi = False Then 'Idiotenmodus
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "TerKat", "FenLin", clFen.FeLin
        IniSetVal "TerKat", "FenObe", clFen.FeObn
        IniSetVal "TerKat", "FenBre", clFen.FeBre
        IniSetVal "TerKat", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KM_Zeilenmarker: TrGrla "GrdMkr":
Case KM_Zeilenumbruch: TrGrla "GrdZei":
Case KM_Gitternetz: TrGrla "GrdGrl":
Case KM_Gruppierung: TrGrla "GrdGkp":
Case KM_Vorschauzeile: TrGrla "GrdPrv":
Case KY_F1: FHilfe
Case KY_F2:
Case KY_F3:
Case KY_F4:
Case KY_F5:
Case KY_F6:
Case KY_F7:
Case KY_F9:
Case KY_F10:
Case KY_F11: Unload Me
Case KA_Hilfe: FHilfe
Case KA_SuCo1: TrGrp
Case KA_SuCo2: TrButt
Case KA_SuFe1: FSuch
Case KA_KaBu1: TrSuch 5
Case KA_KaBu2: TrButt
Case KA_KaBu3: FEinf
Case KA_KaBu4: Unload Me
Case 42: TrButt
Case 142: TrSuch 4, "Ä"
Case 153: TrSuch 4, "Ö"
Case 154: TrSuch 4, "Ü"
Case Else: If TolId >= 65 And TolId <= 90 Then TrSuch 4, Chr$(TolId)
End Select

GlToo = False

End Sub
Private Sub comBar03_Resize()
On Error Resume Next

Dim ClRe As RECT

If TrLad = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    TrPosi
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub FEinf()
On Error GoTo OrErr

Dim RowNr As Integer
Dim KrRow As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

If WindowLoad("frmTermin") = True Then
    Set RpCon = frmTermin.repCont1
    Set CmBrs = frmTermin.comBar02
    Set CmAcs = CmBrs.Actions
    DoEvents
    Tr_Einf
Else
    If GlBut = RibTab_Kat_Eintrg Then
        K_Diag
    Else
        Set RpCon = frmTermVo.repCont1
        Set CmBrs = frmTermVo.comBar02
        Set CmAcs = CmBrs.Actions
        DoEvents
        Tr_VoEi
    End If
End If
DoEvents

If GlBut <> RibTab_Kat_Eintrg Then
    Set RpSel = RpCo1.SelectedRows
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        RowNr = RpRow.Index
        SUpTe RowNr
    End If
    
    Set RpSel = RpCon.SelectedRows
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        KrRow = RpRow.Index
    Else
        KrRow = 1
    End If
    
    If WindowLoad("frmTermin") = True Then
        TUpAb KrRow
        CmAcs(AD_Termin_Abrechnen).Enabled = True
        CmAcs(AD_Termin_EintLoe).Enabled = True
    Else
        TVoUp KrRow
    End If
End If

GlTSa = True

Set RpCon = Nothing
Set RpCo1 = Nothing
Set RpCls = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEinf " & Err.Number
Resume Next

End Sub
Private Sub comBar03_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If TrLad = False Then FTool Control.id
End Sub

Private Sub Form_Load()

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 14000
    .ClientMaxWidth = 11000
    .ClientMinHeight = 6000
    .ClientMinWidth = 5300
    .TopMost = True
End With

Set FrmEx = Nothing

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FClos
    Set frmTerKat = Nothing
End Sub
Private Sub repCont9_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            FEinf
        End If
    End If
End Sub
Private Sub repCont9_KeyUp(KeyCode As Integer, Shift As Integer)

Dim RpCon As XtremeReportControl.ReportControl

Set RpCon = Me.repCont9

If GlAkt = False Then
    If RpCon.Records.Count > 0 Then
        Set RpSel = RpCon.SelectedRows
        If RpSel.Count > 0 Then
            If KeyCode >= 65 And KeyCode <= 90 Then
                TrSuch 4, Chr$(KeyCode)
            Else
                FSum
            End If
        End If
    End If
End If

Set RpSel = Nothing
Set RpCon = Nothing

End Sub
Private Sub repCont9_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    FSum
End Sub
Private Sub repCont9_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FEinf
End Sub
