VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Begin VB.Form frmKraEd 
   Caption         =   "Krankenblatteintrag"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   ControlBox      =   0   'False
   Icon            =   "frmKraEd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7020
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      _Version        =   1048579
      _ExtentX        =   1508
      _ExtentY        =   1508
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   195
      Left            =   1320
      TabIndex        =   1
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
      Left            =   600
      TabIndex        =   2
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
      Left            =   960
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
   Begin XtremeSuiteControls.FlatEdit txtFiNam 
      Height          =   195
      Left            =   1680
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
   Begin XtremeSuiteControls.FlatEdit txtKomme 
      Height          =   1335
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   2355
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      MaxLength       =   14000
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtForma 
      Height          =   195
      Left            =   1680
      TabIndex        =   6
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
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   960
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   240
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKraEd"
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
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private TxFor As XtremeSuiteControls.FlatEdit
Private FTex1 As XtremeSuiteControls.FlatEdit
Private CoDia As XtremeSuiteControls.CommonDialog
Private TxCoN As Tx4oleLib.TXTextControl

Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private FoLad As Boolean

Private clFen As clsFenster
Private clDru As clsDruck
Private clFil As clsFile

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const GWL_WNDPROC = (-4)
Private Const KEYEVENTF_KEYUP = &H2
Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_EMPTYUNDOBUFFER = &HCD

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Sub FClos()
On Error GoTo LiErr

Dim RetWe As Long
Dim KrLok As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set CmBrs = FM.comBar01
Set RpCoK = FM.repContK
Set CmAcs = CmBrs.Actions

Set RpSel = RpCoK.SelectedRows
Set RpCls = RpCoK.Columns

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Kra_Lock)
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            KrLok = True
        Else
            KrLok = False
        End If
    End If
End If

If KrLok = False Then
    If GlKaS = True Then 'Krankenblatteintrag Sepichern
        KrSav True
    End If
End If

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

CmAcs(SY_KB_KraBla_Hinzufueg).Enabled = True
CmAcs(SY_KB_KraBla_Loeschen).Enabled = True

If GlIdi = False Then 'Idiotenmodus
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "Krankenblatt", "FenLin", clFen.FeLin
        IniSetVal "Krankenblatt", "FenObe", clFen.FeObn
        IniSetVal "Krankenblatt", "FenBre", clFen.FeBre
        IniSetVal "Krankenblatt", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCoK = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FEin3()
    K_Kat1 "AnEi"
End Sub
Public Sub FExpo()
On Error GoTo InErr

Dim TmpSt As String
Dim DaNam As String 'Dateinamen
Dim FiNam As String 'Filename
Dim PrNam As String
Dim RetWe As Boolean

Set FM = frmKraEd
Set FTex1 = FM.txtKomme
Set CoDia = frmMain.comDialo

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

DaNam = S_AdIdx(GlAdr, "IDKurz") & ".txt"

TmpSt = FTex1.Text

If TmpSt <> vbNullString Then
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.txt"
        .DialogTitle = "Bitte Name und Ordner der Exportdatei angeben"
        .FileName = GlEPf & DaNam
        .Filter = "Windows Ansi-Text Format (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
        .InitDir = GlEPf
        .ShowSave
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Sub
    End With
    If Right$(FiNam, 4) <> ".txt" Then
        FiNam = FiNam & ".txt"
    End If
    
    If Not IsNull(FiNam) And Not FiNam = vbNullString Then
        With clFil
            .FilPfa FiNam
            If .FilVor(FiNam) = True Then
                .DaLoe = FiNam & vbNullChar
                .FilLoe
            End If
            .StrDa = TmpSt
            RetWe = .FilWrSt
            DoEvents
            PrNam = .FilAnw(FiNam)
            DoEvents
            WindowStart Chr$(34) & PrNam & Chr$(34) & Chr$(32) & Chr$(34) & FiNam & Chr$(34), vbNormalNoFocus, False, False
        End With
    End If
    Set clFil = Nothing
    
    DoEvents
    Unload FM
Else
    Set clFil = Nothing
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FExpo " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50791)
TeMai = IniGetOpt("Hilfe", 50792)
TeInh = IniGetOpt("Hilfe", 50793)
TeFus = IniGetOpt("Hilfe", 50794)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FTxUn()
On Error GoTo PoErr

Dim RetWe As Long

Set FM = frmKraEd
Set FTex1 = FM.txtKomme

RetWe = SendMessage(FTex1.hwnd, EM_CANUNDO, 0&, 0&)

If RetWe <> 0 Then
  Call SendMessage(FTex1.hwnd, EM_UNDO, 0&, 0&)
  Call SendMessage(FTex1.hwnd, EM_EMPTYUNDOBUFFER, 0&, 0&)
End If

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KFarb " & Err.Number
Resume Next

End Sub
Private Sub KFarb()
On Error GoTo PoErr
'Ändert die Farbe im Texteditor

Dim TmKTF As String
Dim AktZa As Integer
Dim EiTyp As Integer
Dim Lange As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKraEd
Set FTex1 = FM.txtKomme
Set TxFor = FM.txtForma
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

If TxFor.Text <> vbNullString Then
    TmKTF = TxFor.Text
Else
    TmKTF = "0000L000000001677721510Arial"
End If

Lange = Len(TmKTF)

EiTyp = CmCom.ItemData(CmCom.ListIndex)

For AktZa = 1 To UBound(GlKrA)
    If GlKrA(AktZa, 0) > 9 Then
        Select Case GlKrA(AktZa, 0)
        Case 24:
        Case 101:
        Case 102:
        Case 104:
        Case 105:
        Case Else:
            If EiTyp = GlKrA(AktZa, 0) Then
                FTex1.ForeColor = GlKrA(AktZa, 3)
                TmKTF = Left$(TmKTF, 5) & Format$(GlKrA(AktZa, 3), "00000000") & Mid$(TmKTF, 14, Lange - 8)
                Exit For
            End If
        End Select
    End If
Next AktZa

TxFor.Text = TmKTF

If GlKrA(AktZa, 5) <> vbNullString Then
    If CBool(GlKrA(AktZa, 5)) = True Then
        CmBrs.Item(3).Visible = True
    Else
        CmBrs.Item(3).Visible = False
    End If
Else
    CmBrs.Item(3).Visible = False
End If

Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "KFarb " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim ItmLi As Long
Dim ItmOb As Long
Dim ItmRe As Long
Dim ItmHo As Long
Dim ItmBr As Long
Dim ItmTo As Long
Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim DaPi1 As XtremeCalendarControl.DatePicker
Dim CmEdi As XtremeCommandBars.CommandBarEdit

Set FM = frmKraEd
Set CmBrs = FM.comBar02
Set DaPi1 = FM.dtpDatu1
Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)

CmEdi.GetRect ItmLi, ItmOb, ItmRe, ItmHo

If IsDate(CmEdi.Text) Then
    NeuDa = CmEdi.Text
Else
    NeuDa = Date
End If

With DaPi1
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Left = ItmLi
    .Top = ItmHo
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
Private Sub FSuFe(Optional ByVal SuKey As Boolean = False)
On Error GoTo OrErr
'Suchleiste einblenden oder Suchformular anzeigen

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKraEd
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set FTex1 = FM.txtKomme

Set CmCom = CmBrs.FindControl(CmCom, SY_SuCm1, , True)
            
If SuKey = False Then
    GlSPh = Not GlSPh 'Suchleiste Textphrase
    CmBrs.Item(2).Visible = GlSPh
    CmAcs(Tex_Suchen).Checked = GlSPh
    IniSetVal "Layout", "SuPhra", GlSPh
    If GlSPh = False Then
        If FTex1.Enabled = True Then
            FTex1.SetFocus
        End If
    Else
        With CmCom
            If .Enabled = True Then
                .SetFocus
                .Execute
            End If
        End With
    End If
Else
    If GlSPh = False Then 'Suchleiste Textphrase
        GlSPh = Not GlSPh 'Suchleiste Textphrase
        CmBrs.Item(2).Visible = GlSPh
        CmAcs(Tex_Suchen).Checked = GlSPh
        IniSetVal "Layout", "SuPhra", GlSPh
    End If
    With CmCom
        .SetFocus
        .Execute
    End With
End If

Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFe " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long, Optional ByVal ColID As Long, Optional ByVal CoTex As String)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KA_Kalen:
Case KA_KaBu1: FKale
Case KA_Uhrze:
Case AM_Beenden: Unload Me
Case KA_SuCo1: KFarb
Case KA_SuCo3:
Case KA_Hilfe: FHilfe
Case SY_SuCm1: K_KrEd 1, 1
Case SY_SuCm2: K_KrEd 1, 2
Case Tex_Suchen: FSuFe
Case Tex_DatSpe: KrSav
Case Tex_TexCut: FTxCo TolId
Case Tex_TexCop: FTxCo TolId
Case Tex_TexEin: FTxCo TolId
Case Tex_DatSpV: FExpo
Case Tex_DocDru: FDruk
Case Tex_ForFet: FTxCo TolId
Case Tex_ForKur: FTxCo TolId
Case Tex_ForUnt: FTxCo TolId
Case Tex_ForDur: FTxCo TolId
Case Tex_AusrLi: FTxCo TolId
Case Tex_AusrRe: FTxCo TolId
Case Tex_AusrZe: FTxCo TolId
Case Tex_FntAu6: FTxCo TolId, , CoTex
Case Tex_FntGr6: FTxCo TolId, , CoTex
Case Tex_FaVor1: FTxCo TolId, ColID
Case Tex_FaVor2: FTxCo TolId
Case Tex_FaHin1: FTxCo TolId, ColID
Case Tex_FaHin2: FTxCo TolId
Case Tex_EdUndo: FTxCo TolId
Case KY_F1: FHilfe
Case KY_F5: FSuFe True
Case KY_F8: KrSav
Case KY_F10: FDruk
Case KY_F11: Unload Me
Case KY_CT_A: FTxEd TolId
Case KY_CT_S: FTxEd TolId
Case KY_CT_V: FTxEd TolId
Case KY_CT_BS: FTxEd TolId
End Select

GlToo = False

End Sub
Private Sub FTxCo(ByVal TxFun As Integer, Optional ByVal ColID As Long, Optional ByVal TxStr As String)
On Error GoTo PoErr

Dim CoIdx As Long
Dim TmpSt As String
Dim TmKTF As String
Dim Lange As Integer
Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKraEd
Set TxFor = FM.txtForma
Set CmBrs = FM.comBar02
Set FTex1 = FM.txtKomme
Set CmAcs = CmBrs.Actions
Set CoDia = frmMain.comDialo
Set TxCoN = frmMain.TexCont1

If TxFor.Text <> vbNullString Then
    TmKTF = TxFor.Text
Else
    TmKTF = "0000L000000001677721510Arial"
End If

Lange = Len(TmKTF)

Select Case TxFun
Case Tex_TexCut:
        TmpSt = FTex1.SelText
        FTex1.SelText = vbNullString
        Clipboard.Clear
        Clipboard.SetText TmpSt
Case Tex_TexCop:
        TmpSt = FTex1.SelText
        Clipboard.Clear
        Clipboard.SetText TmpSt
Case Tex_TexEin:
        With Clipboard
            TmpSt = .GetText
            .Clear
            .SetText TmpSt
        End With
        With TxCoN
            If .CanPaste = True Then
                .ResetContents
                .Paste 5
                .SelText = TmpSt
                .Clip 1
            End If
        End With
        FTex1.SelText = Clipboard.GetText
Case Tex_ForFet:
    If Mid$(TmKTF, 1, 1) = "0" Then
        FTex1.Font.Bold = True
        TmKTF = "1" & Mid$(TmKTF, 2, Lange - 1)
        CmAcs(Tex_ForFet).Checked = True
    Else
        FTex1.Font.Bold = False
        TmKTF = "0" & Mid$(TmKTF, 2, Lange - 1)
        CmAcs(Tex_ForFet).Checked = False
    End If
Case Tex_ForKur:
    If Mid$(TmKTF, 2, 1) = "0" Then
        FTex1.Font.Italic = True
        TmKTF = Left$(TmKTF, 1) & "1" & Mid$(TmKTF, 3, Lange - 1)
        CmAcs(Tex_ForKur).Checked = True
    Else
        FTex1.Font.Italic = False
        TmKTF = Left$(TmKTF, 1) & "0" & Mid$(TmKTF, 3, Lange - 1)
        CmAcs(Tex_ForKur).Checked = False
    End If
Case Tex_ForUnt:
    If Mid$(TmKTF, 3, 1) = "0" Then
        FTex1.Font.Underline = True
        TmKTF = Left$(TmKTF, 2) & "1" & Mid$(TmKTF, 4, Lange - 1)
        CmAcs(Tex_ForUnt).Checked = True
    Else
        FTex1.Font.Underline = False
        TmKTF = Left$(TmKTF, 2) & "0" & Mid$(TmKTF, 4, Lange - 1)
        CmAcs(Tex_ForUnt).Checked = False
    End If
Case Tex_ForDur:
    If Mid$(TmKTF, 4, 1) = "0" Then
        FTex1.Font.Strikethrough = True
        TmKTF = Left$(TmKTF, 3) & "1" & Mid$(TmKTF, 5, Lange - 1)
        CmAcs(Tex_ForDur).Checked = True
    Else
        FTex1.Font.Strikethrough = False
        TmKTF = Left$(TmKTF, 3) & "0" & Mid$(TmKTF, 5, Lange - 1)
        CmAcs(Tex_ForDur).Checked = False
    End If
Case Tex_AusrLi:
        FTex1.Alignment = xtpEditAlignLeft
        TmKTF = Left$(TmKTF, 4) & "L" & Mid$(TmKTF, 6, Lange - 1)
        CmAcs(Tex_AusrLi).Checked = True
        CmAcs(Tex_AusrRe).Checked = False
        CmAcs(Tex_AusrZe).Checked = False
Case Tex_AusrRe:
        FTex1.Alignment = xtpEditAlignRight
        TmKTF = Left$(TmKTF, 4) & "R" & Mid$(TmKTF, 6, Lange - 1)
        CmAcs(Tex_AusrLi).Checked = False
        CmAcs(Tex_AusrRe).Checked = True
        CmAcs(Tex_AusrZe).Checked = False
Case Tex_AusrZe:
        FTex1.Alignment = xtpEditAlignCenter
        TmKTF = Left$(TmKTF, 4) & "Z" & Mid$(TmKTF, 6, Lange - 1)
        CmAcs(Tex_AusrLi).Checked = False
        CmAcs(Tex_AusrRe).Checked = False
        CmAcs(Tex_AusrZe).Checked = True
Case Tex_FaVor1:
        FTex1.ForeColor = ColID
        TmKTF = Left$(TmKTF, 5) & Format$(ColID, "00000000") & Mid$(TmKTF, 14, Lange - 8)
Case Tex_FaVor2:
        FTex1.ForeColor = 0
        TmKTF = Left$(TmKTF, 5) & "00000000" & Mid$(TmKTF, 14, Lange - 8)
Case Tex_FaHin1:
        FTex1.BackColor = ColID
        TmKTF = Left$(TmKTF, 13) & Format$(ColID, "00000000") & Mid$(TmKTF, 22, Lange - 8)
Case Tex_FaHin2:
        FTex1.BackColor = 16777215
        TmKTF = Left$(TmKTF, 13) & "16777215" & Mid$(TmKTF, 22, Lange - 8)
Case Tex_FntAu6:
        FTex1.Font.Name = TxStr
        TmKTF = Left$(TmKTF, 23) & TxStr
Case Tex_FntGr6:
        FTex1.Font.SIZE = CLng(TxStr)
        TmKTF = Left$(TmKTF, 21) & Format$(CLng(TxStr), "00") & Mid$(TmKTF, 24, Lange - 2)
Case IC16_FarVor:
        With CoDia
            .CancelError = True
            .DialogStyle = 1
            .Color = FTex1.ForeColor
            .ShowColor
            CoIdx = .Color
        End With
        FTex1.ForeColor = CoIdx
Case IC16_FarHin:
        With CoDia
            .CancelError = True
            .DialogStyle = 1
            .Color = FTex1.BackColor
            .ShowColor
            CoIdx = .Color
        End With
        FTex1.BackColor = CoIdx
Case Tex_EdUndo:
        FTxUn
End Select

TxFor.Text = TmKTF

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxCO " & Err.Number
Resume Next

End Sub
Private Sub FTxEd(ByVal TxFun As Integer)
On Error GoTo PoErr

Dim TmTex As String
Dim TmpSt As String
Dim Lange As Integer
Dim CuPos As Integer
Dim LePos As Integer
Dim EnPos As Integer

Set FM = frmKraEd
Set FTex1 = FM.txtKomme
Set TxCoN = frmMain.TexCont1

TmTex = FTex1.Text
Lange = Len(TmTex)

Select Case TxFun
Case KY_CT_A:
        With FTex1
            .SelStart = 0
            .SelLength = Lange
        End With
Case KY_CT_S:
        KrSav
Case KY_CT_V:
        With Clipboard
            TmpSt = .GetText
            .Clear
            .SetText TmpSt
        End With
        FTex1.SelText = Clipboard.GetText
Case KY_CT_BS:
    If Lange > 0 Then
        CuPos = FTex1.SelStart
        LePos = InStrRev(TmTex, Chr$(32), CuPos, 1)
        If LePos > 0 Then
            EnPos = InStr(LePos + 1, TmTex, Chr$(32), 1)
            With FTex1
                .SelStart = LePos
                .SelLength = EnPos - (LePos + 1)
            End With
        End If
    End If
End Select

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxEd " & Err.Number
Resume Next

End Sub


Private Sub FText()
On Error GoTo PoErr

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmKraEd
Set CmBrs = FM.comBar02
Set FTex1 = FM.txtKomme
Set CmSta = CmBrs.StatusBar

CmSta.Pane(0).Text = "Anzahl Zeichen : " & Len(FTex1.Text)

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FText " & Err.Number
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
Public Sub FDruk()
On Error GoTo InErr
'Zeigt den Druckdiolog

Dim RetWe As Long
Dim DrNam As String

Set FM = frmKraEd
Set FTex1 = FM.txtKomme
Set CoDia = frmMain.comDialo

Set clDru = New clsDruck

If FTex1.Text <> vbNullString Then
    If GlKaS = True Then 'Krankenblatteintrag Sepichern
        KrSav True
        DoEvents
    End If
    
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
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlAkt = False Then
        Select Case Control.id
        Case Tex_FaVor1: FTool Control.id, Control.Color
        Case Tex_FaHin1: FTool Control.id, Control.Color
        Case Tex_FntAu6: FTool Control.id, , Control.Text
        Case Tex_FntGr6: FTool Control.id, , Control.Text
        Case Else: FTool Control.id
        End Select
    End If
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
    KrPos
    FoLad = False
End Sub

Private Sub Form_Load()

Set FrmEx = Me.frmExtde

FoLad = True

With FrmEx
    .ClientMaxHeight = 11000
    .ClientMaxWidth = 18000
    .ClientMinHeight = 5000
    .ClientMinWidth = 12400
    .TopMost = GlKrF 'Krankenblattdialog im Vordergund
End With

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmKraEd = Nothing
End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlAkK = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    KrPos
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub txtKomme_Change()
    If FoLad = False Then
        FText
        GlKaS = True 'Krankenblatteintrag Sepichern
    End If
End Sub

Private Sub txtKomme_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Set FM = frmKraEd
Set FTex1 = FM.txtKomme

If Shift = vbCtrlMask Then
    If KeyCode = vbKeyV Then
        KeyCode = 0
    End If
End If

End Sub


