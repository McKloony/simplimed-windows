VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Begin VB.Form frmKomment 
   Caption         =   "Texteingabe"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   ControlBox      =   0   'False
   Icon            =   "frmKomment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6045
   Begin XtremeSuiteControls.FlatEdit txtKomme 
      Height          =   1335
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   2355
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      MaxLength       =   14000
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   1215
      Left            =   3480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      _Version        =   1048579
      _ExtentX        =   2778
      _ExtentY        =   2143
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   14000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtNeuEi 
      Height          =   195
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   14000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtIdxNr 
      Height          =   195
      Left            =   840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   14000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtFiNam 
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   14000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   960
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKomment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAcs As XtremeCommandBars.CommandBarActions
Private FTex1 As XtremeSuiteControls.FlatEdit
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1
Private CoDia As XtremeSuiteControls.CommonDialog

Private SuLei As Boolean

Private clFen As clsFenster
Private clDru As clsDruck
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = ""
TeMai = ""
TeInh = ""
TeFus = ""

'SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FText()
On Error GoTo PoErr

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmKomment
Set CmBrs = FM.comBar02
Set FTex1 = FM.txtKomme
Set CmSta = CmBrs.StatusBar

CmSta.Pane(0).Text = "Anzahl Zeichen : " & Len(FTex1.Text)

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FText " & Err.Number
Resume Next

End Sub
Private Sub KFarb()
On Error GoTo PoErr
'Ändert die Farbe im Texteditor

Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKomment
Set FTex1 = FM.txtKomme
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

LiIdx = CmCom.ListIndex + 9

FTex1.ForeColor = GlKrA(LiIdx, 3)

Select Case CmCom.ListIndex
Case 1:
    CmAcs(SY_OP_Suchen).Enabled = True
    K_Kom
Case 13:
    CmAcs(SY_OP_Suchen).Enabled = True
    K_Kom
Case Else:
    CmAcs(SY_OP_Suchen).Enabled = False
    SuLei = True
    KSuLe
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KFarb " & Err.Number
Resume Next

End Sub
Private Sub KSuLe()
On Error GoTo PoErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKomment
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo2, , True)

SuLei = Not SuLei

CmBrs.Item(4).Visible = SuLei

CmAcs(SY_OP_Suchen).Checked = SuLei

If CmCom.ListIndex = 0 Then K_Kom

If SuLei = True Then
    CmCom.SetFocus
    CmCom.Execute
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmAcs = Nothing
Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KSuLe " & Err.Number
Resume Next

End Sub
Private Sub FClos()
On Error GoTo LiErr

Dim RetWe As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

CmAcs(SY_KB_KraBla_Hinzufueg).Enabled = True
CmAcs(SY_KB_KraBla_Loeschen).Enabled = True

If GlRes = False Then 'Reset der Einstellungen
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "Kommentar", "FenLin", clFen.FeLin
        IniSetVal "Kommentar", "FenObe", clFen.FeObn
        IniSetVal "Kommentar", "FenBre", clFen.FeBre
        IniSetVal "Kommentar", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim DaPi1 As XtremeCalendarControl.DatePicker
Dim CmEdi As XtremeCommandBars.CommandBarEdit

Set CmBrs = Me.comBar02
Set DaPi1 = Me.dtpDatu1

Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)

If DaPi1.Selection.BlocksCount > 0 Then
    NeuDa = DaPi1.Selection.Blocks(0).DateBegin
    CmEdi.Text = NeuDa
End If

Set DaPi1 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub

Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim DaPi1 As XtremeCalendarControl.DatePicker
Dim CmEdi As XtremeCommandBars.CommandBarEdit

Set FM = frmKomment
Set CmBrs = FM.comBar02
Set DaPi1 = FM.dtpDatu1
Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)

If IsDate(CmEdi.Text) Then
    NeuDa = CmEdi.Text
Else
    NeuDa = Date
End If

With DaPi1
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Left = 3940
    .Top = 860
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            CmEdi.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set DaPi1 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Public Sub FDruk()
On Error GoTo InErr
'Zeigt den Druckdiolog

Dim RetWe As Long
Dim DrNam As String

Set FM = frmKomment
Set FTex1 = FM.txtKomme
Set CoDia = frmMain.comDialo

Set clDru = New clsDruck

If FTex1.Text <> vbNullString Then
    RetWe = clDru.DruDia()
    If RetWe = 1 Then
        clDru.DruTex FTex1.Text
    End If
End If

Set clDru = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDruk " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F5: KSuLe
Case KY_F8: KoSav
            Unload Me
Case KY_F10: FDruk
Case KY_F11: Unload Me
Case SY_OP_Hilfe: FHilfe
Case SY_OP_Speichern: KoSav
                      Unload Me
Case SY_OP_Suchen: KSuLe
Case SY_OP_Hilfe: FHilfe
Case SY_OP_Abbruch: Unload Me
Case SY_OP_Drucken: FDruk
Case KA_KaBu1: FKale
Case KA_SuCo1: KFarb
Case KA_SuCo2: KEing
End Select

GlToo = False

End Sub
Private Sub KEing()
On Error GoTo PoErr
'Ändert die Farbe im Texteditor

Dim MeStr As Variant

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKomment
Set FTex1 = FM.txtKomme
Set CmBrs = FM.comBar02

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo2, , True)

MeStr = FTex1.Text

If MeStr <> vbNullString Then
    MeStr = MeStr & "; " & CmCom.Text
Else
    MeStr = CmCom.Text
End If

FTex1.Text = MeStr

CmCom.Text = vbNullString

CmCom.SetFocus
CmCom.Execute

Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KEing " & Err.Number
Resume Next

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
Private Sub comBar02_Resize()
    KoAus
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
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_Activate()
    AFont Me
    Me.txtKomme.SetFocus
    KoAus
    FText
End Sub
Private Sub Form_Load()

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 8000
    .ClientMaxWidth = 12000
    .ClientMinHeight = 4000
    .ClientMinWidth = 8800
    .TopMost = True
End With

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmKomment = Nothing
End Sub
Private Sub txtKomme_Change()
    FText
End Sub

