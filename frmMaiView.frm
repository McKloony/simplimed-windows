VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{621DDB00-A516-11E8-A658-0013D350667C}#3.2#0"; "tx4ole26.ocx"
Begin VB.Form frmMaiView 
   Caption         =   "Email"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   ControlBox      =   0   'False
   Icon            =   "frmMaiView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows-Standard
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   1900
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   5000
      _Version        =   1048579
      _ExtentX        =   8819
      _ExtentY        =   3351
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbEmEmp 
         Height          =   315
         Left            =   1000
         TabIndex        =   3
         Top             =   120
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnEmEmp 
         Height          =   330
         Left            =   100
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "An :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtEmCCM 
         Height          =   350
         Left            =   1000
         TabIndex        =   4
         Top             =   560
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtEmBet 
         Height          =   350
         Left            =   1000
         TabIndex        =   6
         Top             =   1440
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnEmCCM 
         Height          =   330
         Left            =   100
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   560
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Cc :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnEMBCC 
         Height          =   330
         Left            =   100
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1000
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Bcc :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbEmBCC 
         Height          =   315
         Left            =   1000
         TabIndex        =   5
         Top             =   1000
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   250
         Left            =   200
         TabIndex        =   10
         Top             =   1480
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "Betreff :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.WebBrowser WebBrow1 
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   4680
      Width           =   2775
      _Version        =   1048579
      _ExtentX        =   4895
      _ExtentY        =   4683
      _StockProps     =   173
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   15000
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
   Begin Tx4oleLib.TXTextControl TexCont3 
      Height          =   2715
      Left            =   3480
      TabIndex        =   11
      Top             =   4680
      Width           =   2715
      _Version        =   196610
      _ExtentX        =   4789
      _ExtentY        =   4789
      _StockProps     =   73
      BackColor       =   16777215
      Language        =   49
      BorderStyle     =   1
      BackStyle       =   1
      ControlChars    =   0   'False
      EditMode        =   0
      HideSelection   =   -1  'True
      InsertionMode   =   -1  'True
      MousePointer    =   0
      ZoomFactor      =   100
      ViewMode        =   0
      ClipChildren    =   0   'False
      ClipSiblings    =   -1  'True
      SizeMode        =   0
      TabKey          =   -1  'True
      FormatSelection =   0   'False
      VTSpellDictionary=   "C:\PROGRA~1\THEIMA~1\TXTEXT~2.0\Bin\AMERICAN.VTD"
      ScrollBars      =   3
      PageWidth       =   12240
      PageHeight      =   15840
      PageMarginL     =   1440
      PageMarginT     =   1440
      PageMarginR     =   1440
      PageMarginB     =   1440
      PrintZoom       =   100
      PrintOffset     =   0   'False
      PrintColors     =   -1  'True
      FontName        =   "Arial"
      FontSize        =   12
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Baseline        =   0
      TextBkColor     =   16777215
      Alignment       =   0
      LineSpacing     =   100
      LineSpacingT    =   0
      FrameStyle      =   32
      FrameDistance   =   0
      FrameLineWidth  =   20
      IndentL         =   0
      IndentR         =   0
      IndentFL        =   0
      IndentT         =   0
      IndentB         =   0
      Text            =   ""
      WordWrapMode    =   1
      AllowUndo       =   -1  'True
      TextFrameMarkerLines=   -1  'True
      FieldLinkTargetMarkers=   0   'False
      PageOrientation =   0
      PageViewStyle   =   1
      FontSettings    =   0
      AllowDrag       =   0   'False
      AllowDrop       =   0   'False
      SelectionViewMode=   1
      SectionRestartPageNumbering=   0
      PermanentControlChars=   16
      RightToLeft     =   0   'False
      TextDirection   =   2
      Locale          =   1031
      Justification   =   1
      FrameColor      =   16777215
      FrameLineColor  =   0
      DocumentPermissions=   31
      SelectObjects   =   -1  'True
      IsTrackChangesEnabled=   0   'False
      IsFormulaCalculationEnabled=   -1  'True
      FormulaReferenceStyle=   0
   End
   Begin XtremeSuiteControls.FlatEdit txtMaiTx 
      Height          =   1215
      Left            =   6480
      TabIndex        =   12
      Top             =   4800
      Width           =   1575
      _Version        =   1048579
      _ExtentX        =   2778
      _ExtentY        =   2143
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      MultiLine       =   -1  'True
      ScrollBars      =   2
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
      Left            =   720
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMaiView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmSwi As XtremeCommandBars.StatusBarSwitchPane
Private CmPgs As XtremeCommandBars.StatusBarProgressPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CoDia As XtremeSuiteControls.CommonDialog
Private WeBr1 As XtremeSuiteControls.WebBrowser
Private TxCoN As Tx4oleLib.TXTextControl
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow

Private MaiMail As EAGetMailObjLib.Mail

Private WithEvents CmSta As XtremeCommandBars.StatusBar
Attribute CmSta.VB_VarHelpID = -1
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Public mPaNr As Long
Private mAnza As Integer

Private Const OLECMDID_PRINT = 6
Private Const OLECMDID_PRINT2 = 49
Private Const OLECMDID_PRINTPREVIEW = 7
Private Const OLECMDID_PRINTPREVIEW2 = 50
Private Const OLECMDID_SAVE = 3
Private Const OLECMDID_SAVEAS = 4
Private Const OLECMDID_SAVECOPYAS = 5

Private Const OLECMDEXECOPT_DODEFAULT = 0
Private Const OLECMDEXECOPT_PROMPTUSER = 1
Private Const OLECMDEXECOPT_DONTPROMPTUSER = 2
Private Const OLECMDEXECOPT_SHOWHELP = 3

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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Sub FChek(ByVal ChTyp As Integer)
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set CmAcs = CmBrs.Actions

Select Case ChTyp
Case 1: CmAcs(TX_Mail_Prioritaet).Checked = Not CmAcs(TX_Mail_Prioritaet).Checked
Case 2: CmAcs(TX_Mail_Notific).Checked = Not CmAcs(TX_Mail_Notific).Checked
Case 3: CmAcs(TX_Mail_NoHTML).Checked = Not CmAcs(TX_Mail_NoHTML).Checked
End Select

Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FChek " & Err.Number
Resume Next

End Sub
Private Sub FClip(ByVal AnTyp As Integer)
On Error GoTo SaErr
'Kopiert die GuiID in die Zwischenablage

Dim RowNr As Long
Dim MaGui As String
Dim FiNam As String
Dim DaNam As String
Dim TmMes As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMaiView
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set RpCo0 = frmMain.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

Select Case AnTyp
Case 2:
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            MaGui = MaAry(Mai_GuiID, RowNr)
        End If
    End If
    Clipboard.Clear
    Clipboard.SetText MaGui
Case 3:
    Clipboard.Clear
    Clipboard.SetText GlTxE
End Select

Set RpCo0 = Nothing
Set RpCls = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClip " & Err.Number
Exit Sub

End Sub
Private Sub FCopy()
On Error GoTo SaErr

Dim DaNam As String
Dim FiNam As String
Dim NeNam As String
Dim ExpOr As String 'Exportordner
Dim AktZa As Integer
Dim AnzAt As Integer
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMaiView
Set RpCo0 = frmMain.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

If GlRDP = True Then
    If LCase(GlEPf) = LCase(GlIPf) Then
        ExpOr = GlDpf & "Import\"
    Else
        ExpOr = GlEPf
    End If
Else
    ExpOr = GlEPf
End If

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        DaNam = MaAry(Mai_MailFile, RpRow.Index)
        FiNam = GlDpf & "Emails\" & DaNam
    End If
End If

NeNam = ExpOr & DaNam

With clFil
    If .FilVor(FiNam) = True Then
        If .FilVor(NeNam) = True Then
            .DaLoe = NeNam & vbNullChar
            .FilLoe
        End If
        .DaCop = FiNam & ";" & NeNam & vbNullChar
        .FilCop 1
    End If
End With

Set clFil = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FCopy " & Err.Number
Exit Sub

End Sub
Private Sub FNeu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set CmAcs = CmBrs.Actions
Set RbTab = RbBar.SelectedTab

Select Case RbTab.id
Case RibTab_Tex_Dokumt:
            SMaAn 0
Case RibTab_Tex_Vorlag:
            GlMaE = 1
            frmAdrSuch.Show vbModal
End Select

Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeu " & Err.Number
Resume Next

End Sub
Private Sub FClos()
On Error GoTo LiErr

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlIdi = False Then 'Idiotenmodus
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "MailView", "FenLin", clFen.FeLin
        IniSetVal "MailView", "FenObe", clFen.FeObn
        IniSetVal "MailView", "FenBre", clFen.FeBre
        IniSetVal "MailView", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FDele()
On Error GoTo SaErr

Dim DaNam As String
Dim FiNam As String
Dim TmPfa As String
Dim TmpNa As String
Dim AktZa As Integer
Dim AnzAt As Integer
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMaiView
Set RpCo0 = frmMain.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        DaNam = MaAry(Mai_MailFile, RpRow.Index)
        FiNam = GlDpf & "Emails\" & DaNam
        TmPfa = GlDpf & "Emails\Temp\"
        TmpNa = TmPfa & Left$(DaNam, Len(DaNam) - 3) & "htm"
    End If
End If

If GlAtV = True Then
    AnzAt = UBound(GlAtt)
End If

With clFil
    If .FilVor(TmpNa) = True Then
        .DaLoe = TmpNa & vbNullChar
        .FilLoe
    End If
    If AnzAt > 0 Then
        For AktZa = 1 To AnzAt
            .FilPfa GlAtt(AktZa)
            If LCase(.DaOrd) = "export" Then
                .DaLoe = GlAtt(AktZa) & vbNullChar
                .FilLoe
            ElseIf LCase(.DaOrd) = "temp" Then
                .DaLoe = GlAtt(AktZa) & vbNullChar
                .FilLoe
            End If
        Next AktZa
        .DaLoe = GlTmp & "*.ll"
        .FilLoe
    End If
End With

Set clFil = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDele " & Err.Number
Exit Sub

End Sub
Private Sub FDial()
On Error GoTo SaErr
'Dateianhang hinzufügen

Dim DaNam As String
Dim FiNam As String
Dim GesZa As Integer
Dim AktZa As Integer
Dim AryIt() As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmAtt As XtremeCommandBars.CommandBarComboBox

Set FM = frmMaiView
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RbBar = CmBrs.Item(1)

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

Set CmAtt = CmBrs.FindControl(CmAtt, SY_SuCm3, , True)

mAnza = CmAtt.ListCount

With clFil
    .hwnd = FM.hwnd
    .StaVe = GlIPf
    .DaTit = "Bitte Name und Ordner der Datei angeben"
    .DaStr = "Unterstützte Formate (*.pdf;*.jpg;*.bmp;*.png;*.tif;*.wmf;*.zip)" & Chr(0) & "*.pdf;*.jpg;*.bmp;*.png;*.tif;*.wmf;*.zip" & Chr(0) & "Adobe-Acrobat Dokument (*.pdf)" & Chr(0) & "*.pdf" & Chr(0) & "Joint Photographic Experts Group (.jpg)" & Chr(0) & "*.jpg" & Chr(0) & "Windows Bitmap (.bmp)" & Chr(0) & "*.bmp" & Chr(0) & "Portable Network Graphics (.png)" & Chr(0) & "*.png" & Chr(0) & "Tagged Image Format (.tif)" & Chr(0) & "*.tif" & Chr(0) & "Windows-Meta-File (.wmf)" & Chr(0) & "*.wmf" & Chr(0) & "Microsoft Word 2002/2003 (.doc)" & Chr(0) & "*.doc" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0)
    FiNam = .FilOpn
End With

If FiNam <> vbNullString Then
    FiNam = ";" & FiNam
    AryIt = Split(FiNam, Chr$(59))
    GesZa = UBound(AryIt)

    For AktZa = 1 To GesZa
        mAnza = mAnza + 1
        With clFil
            .FilPfa AryIt(AktZa)
            DaNam = .DaNam
        End With
        With CmAtt
            .AddItem DaNam, mAnza
            .ListIndex = mAnza
        End With
        ReDim Preserve GlAtt(mAnza + 1)
        GlAtt(mAnza + 1) = AryIt(AktZa)
    Next AktZa
    CmBrs.Item(3).Visible = True
    GlAtV = True
End If

Set clFil = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDial " & Err.Number
Exit Sub

End Sub

Private Sub FDruk()
On Error GoTo AnErr

Dim DrhDc As Long
Dim TmStr As String
Dim GesZa As Integer
Dim DrKop As Integer
Dim ReDru As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set WeBr1 = Me.WebBrow1
Set TxCoN = Me.TexCont3
Set RbBar = CmBrs.Item(1)
Set CmAcs = CmBrs.Actions
Set RbTab = RbBar.SelectedTab

Set CoDia = frmMain.comDialo

Select Case RbTab.id
Case RibTab_Tex_Dokumt:
        With CoDia
            .CancelError = True
            .DialogStyle = 1
            .DialogTitle = "Drucken"
            .FileName = vbNullString
            ReDru = .ShowPrinter
            DrhDc = .hDC
            DrKop = .Copies
        End With
        WeBr1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
        'WeBr1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0
Case RibTab_Tex_Vorlag:
        S_MaSe
End Select

Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDruk " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

If GlNaT = 1 Then 'Mailtyp (1=View 2=Neu 3=Antwort)
    TeTit = IniGetOpt("Hilfe", 50771)
    TeMai = IniGetOpt("Hilfe", 50772)
    TeInh = IniGetOpt("Hilfe", 50773)
    TeFus = IniGetOpt("Hilfe", 50774)
Else
    TeTit = IniGetOpt("Hilfe", 50781)
    TeMai = IniGetOpt("Hilfe", 50782)
    TeInh = IniGetOpt("Hilfe", 50783)
    TeFus = IniGetOpt("Hilfe", 50784)
End If

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FMenu()
On Error GoTo AnErr

Dim RetWe As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSli As XtremeCommandBars.StatusBarSliderPane
Dim CmPrg As XtremeCommandBars.StatusBarProgressPane

Set FM = frmMaiView
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    .Visible = True
    '----------
    Set CmPan = .AddPane(1)
    CmPan.Width = 140
    CmPan.Alignment = xtpAlignmentCenter
    CmPan.Text = vbNullString
    '----------
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Width = 230
    '----------
    Set CmPan = .AddPane(3)
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    CmPan.Alignment = xtpAlignmentLeft
    '----------
    Set CmPan = .AddPane(4)
    CmPan.Text = vbNullString
    CmPan.Width = 20
    '----------
    Set CmPgs = .AddProgressPane(5)
    CmPgs.Width = 160
    CmPgs.Max = 100
    CmPgs.Min = 0
    '----------
    Set CmPan = .AddPane(6)
    CmPan.Text = vbNullString
    CmPan.Width = 20
    '----------
    Set CmSwi = .AddSwitchPane(7)
    CmSwi.AddSwitch IC16_Arrow_Left, vbNullString
    CmSwi.AddSwitch IC16_Arrow_Up, vbNullString
    CmSwi.AddSwitch IC16_Arrow_Right, vbNullString
    '----------
    Set CmPan = .AddPane(8)
    CmPan.Text = vbNullString
    CmPan.Width = 10
    .Visible = True
End With

CmBrs.RecalcLayout

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FSuFe()
On Error GoTo OrErr
'Suchleiste einblenden oder Suchformular anzeigen

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMaiView
Set TxCoN = FM.TexCont3
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, SY_SuCm1, , True)
            
GlSMv = Not GlSMv

CmBrs.Item(2).Visible = GlSMv
CmAcs(TX_Mail_Suchen).Checked = GlSMv
IniSetVal "Layout", "SuMaVi", GlSMv
            
If GlSMv = True Then
    With CmCom
        .SetFocus
        .Execute
    End With
Else
    TxCoN.SetFocus
End If

Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFe " & Err.Number
Resume Next

End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set WeBr1 = Me.WebBrow1
Set TxCoN = Me.TexCont3
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

 Select Case RbTab.id
Case RibTab_Tex_Dokumt:
Case RibTab_Tex_Vorlag:
End Select

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long, Optional ByVal ColID As Long, Optional ByVal CoTex As String)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: FNeu
Case KY_F5: FSuFe
Case KY_F8:
Case KY_F10: FDruk
Case KY_F11: Unload Me
Case TX_Mail_Clip1: FClip 2
Case TX_Mail_Clip2: FCopy
Case TX_Mail_Clip3: FClip 3
Case AM_Beenden: Unload Me
Case AM_Hilfe: FHilfe
Case TX_Mail_Antworten: SMaAn 3
Case TX_Mail_Weiterleiten: SMaAn 4
Case TX_Mail_PatSuch: GlMaE = 4
                      frmAdrSuch.Show vbModal
Case TX_Mail_PatEdit: MaAdr
Case TX_Mail_Ungelesen: S_MaMa 1
Case TX_Mail_Markieren: S_MaMa 2
Case TX_Mail_Junkmail: S_MaMa 3
Case TX_Mail_Senden: S_MaSe
Case TX_Mail_AttOpen: MaSav 1
Case TX_Mail_AttSave: MaSav 2
Case TX_Mail_AttExpo: MaSav 3
Case TX_Mail_AttImpo: MaSav 9
Case TX_Mail_Rechnun: MaSav 4
Case TX_Mail_AttView: MaViw
Case TX_Mail_Prioritaet: FChek 1
Case TX_Mail_Notific: FChek 2
Case TX_Mail_NoHTML: FChek 3
Case TX_Mail_Suchen: FSuFe
Case TX_Mail_Anhang: FDial
Case TX_Mail_Vorlage: MaVor
Case TX_Mail_Drucken: FDruk
Case TX_Mail_Erneut: SMaAn 5
Case TX_Mail_Loeschen: Unload Me
                       SLoHa
Case SY_SuCm1: M_MaEd 1
Case SY_SuCm2: M_MaEd 2
Case Tex_FntAu4: FTxFA TolId, CoTex
Case Tex_FntGr4: FTxFA TolId, CoTex
Case Tex_AusrLi: FTxFA TolId
Case Tex_AusrRe: FTxFA TolId
Case Tex_AusrZe: FTxFA TolId
Case Tex_AusrBl: FTxFA TolId
Case Tex_EinzLi: FTxFF TolId
Case Tex_EinzRe: FTxFF TolId
Case Tex_Zeiche: FTxCo TolId
Case Tex_Absatz: FTxCo TolId
Case Tex_Aufzah: FTxCo TolId
Case Tex_Numeri: FTxCo TolId
Case Tex_ForFet: FTxFF TolId
Case Tex_ForKur: FTxFF TolId
Case Tex_ForUnt: FTxFF TolId
Case Tex_ForDur: FTxFF TolId
Case Tex_TexMar: FTxFA TolId
Case Tex_FntKle: FTxFF TolId
Case Tex_FntGro: FTxFF TolId
Case Tex_FntHoh: FTxFF TolId
Case Tex_FntTif: FTxFF TolId
Case Tex_KopFus: FTxCo TolId
Case Tex_ForSpl: FTxCo TolId
Case Tex_KopZei: FTxFF TolId
Case Tex_FusZei: FTxFF TolId
Case Tex_TabEin: FTxCo TolId
Case Tex_TabAtr: FTxCo TolId
Case Tex_SpEiRe: FTxCo TolId
Case Tex_SpEiLi: FTxCo TolId
Case Tex_ZeEiUn: FTxCo TolId
Case Tex_ZeEiOb: FTxCo TolId
Case Tex_SpalLo: FTxCo TolId
Case Tex_ZeilLo: FTxCo TolId
Case Tex_TxUndo: FTxCo TolId
Case Tex_TxRedo: FTxCo TolId
Case Tex_ForSty: FTxCo TolId
Case Tex_ForVor: FTxCo TolId
Case Tex_TexRah: FObje TolId
Case Tex_EinTab: FObje TolId
Case Tex_Tabell: 'Tabellenmenü
Case Tex_EinLnk: 'Hyperlink
Case Tex_EinObj: FObje TolId
Case Tex_Suchen: FTxCo TolId
Case Tex_Ersetz: FTxCo TolId
Case Tex_TexCut: FTxCo TolId
Case Tex_TexCop: FTxCo TolId
Case Tex_TexEin: FTxCo TolId
Case Tex_DaFeAd: FTxDa CoTex
Case Tex_DaFeLo: FTxFF TolId
Case Tex_DaFeVe: FObje TolId
Case Tex_KopDat: FTxFF TolId
Case Tex_FusZal: FTxFF TolId
Case Tex_DatLoa: FObje TolId
Case Tex_DatSpe: FObje TolId
Case Tex_DatSpV: FObje TolId
Case Tex_DatSav: FObje TolId
Case Tex_DatLoe: FObje TolId
Case Tex_DatKop: FObje TolId
Case Tex_DocDru: FObje TolId
Case Tex_DocVor: FObje TolId
Case Tex_ZeiAb1: FTxFZ TolId
Case Tex_ZeiAb2: FTxFZ TolId
Case Tex_ZeiAb3: FTxFZ TolId
Case Tex_ZeiAb4: FTxFZ TolId
Case Tex_ZeiAb5: FTxFZ TolId
Case Tex_ZeiAb6: FTxFZ TolId
Case Tex_ZeiAb7: FTxFZ TolId
Case Tex_ClpEin: FObje TolId
Case Tex_ClpInh: FObje TolId
Case Tex_FaVor1: FTxCo TolId, ColID
Case Tex_FaVor2: FTxCo TolId
Case Tex_FaVor3: FObje TolId
Case Tex_FaHin1: FTxCo TolId, ColID
Case Tex_FaHin2: FTxCo TolId
Case Tex_FaHin3: FObje TolId
Case IC16_FarVor: FObje TolId
Case IC16_FarHin: FObje TolId
End Select

GlToo = False

End Sub
Private Sub FObje(ByVal TxFun As Integer)
On Error GoTo PoErr

Dim ObjNr As Long
Dim DrhDc As Long
Dim TxBre As Long
Dim TxHoh As Long
Dim ObjBr As Long
Dim ObjHo As Long
Dim SclBr As Long
Dim SclHo As Long
Dim ColID As Long
Dim ExVer As String
Dim SuFix As String
Dim FiNam As String
Dim SuStr As String
Dim TeStr As String
Dim TeNam As String
Dim DaNam As String
Dim Posit As Integer
Dim LadVa As Integer
Dim GesZa As Integer
Dim Frage As Integer
Dim AktZa As Integer
Dim ReDru As Integer
Dim RetWe As Boolean
Dim TxFnt As New StdFont
Dim Mld2, Tit2 As String

Set FM = frmMaiView
Set TxCoN = FM.TexCont3
Set CoDia = frmMain.comDialo

Set clFil = New clsFile
Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Tit2 = "Dokument Entfernen"
Mld2 = "Möchten Sie das markierte Dokument wirklich entfernen?"
TxBre = TxCoN.Width

Select Case TxFun
Case Tex_DatSpe:
        RetWe = STxSa()
Case Tex_EinTab:
        TxCoN.TabDialog
Case Tex_EinObj:
        ObjNr = TxCoN.ObjectInsert(1, 0, -1, 0, 0, 0, 100, 100, 3, 100, 100, 100, 100)
        TxCoN.ObjectCurrent = ObjNr
        TxCoN.ObjectSizeMode = 3
        TxCoN.ObjectInsertionMode = 2
        TxCoN.ObjectTextflow = 1
        TxCoN.ObjectTransparency = 50
Case IC16_FarVor:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.ForeColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.ForeColor = ColID
Case IC16_FarHin:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.TextBkColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.TextBkColor = ColID
Case Tex_FaVor3:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.ForeColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.ForeColor = ColID
Case Tex_FaHin3:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.TextBkColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.TextBkColor = ColID
Case Tex_TexRah:
    TxCoN.TextFrameInsert -1, 0, 0, 0, 0, 0, 3, 100, 100, 100, 100
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    DoEvents
    ObjNr = TxCoN.TextFrameInsert(-1, 1, 0, 0, 0, 0, 3, 200, 200, 200, 200)
Case Tex_ClpEin:
    With TxCoN
        If .CanPaste = True Then
            .Paste 1
        End If
    End With
Case Tex_ClpInh:
    With TxCoN
        If .CanPaste = True Then
            SuStr = Clipboard.GetText
            Clipboard.Clear
            Clipboard.SetText SuStr
            .Paste 5
        End If
    End With
Case Tex_DaFeVe:
    GlAkt = True
    Screen.MousePointer = vbHourglass
    clFen.FenDsk 2
    
    S_TxEin
    DoEvents 'Laden der Patientendaten in Array GlSer()
    
    STxV2 'Verbinden der Textfelder mit GlSer()
    DoEvents
    GlTSV = True 'Speichern Textverarbeitung

    clFen.FenDsk 3
    Screen.MousePointer = vbNormal
    GlAkt = False
Case Tex_DocVor:
    
Case Tex_DocDru:
    If GlTDa <> vbNullString Then 'Neuer Dateiname
        TxCoN.PrintDialog GlTDa
    Else
        TxCoN.PrintDialog "Druck"
    End If
End Select

Set CoDia = Nothing

Set clFil = Nothing
Set clFen = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FObje " & Err.Number
Resume Next

End Sub
Private Sub FTxFA(ByVal TxFun As Integer, Optional ByVal TxStr As String)
On Error GoTo PoErr

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMaiView
Set TxCoN = FM.TexCont3
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Select Case TxFun
Case Tex_AusrLi: TxCoN.Alignment = 0
Case Tex_AusrRe: TxCoN.Alignment = 1
Case Tex_AusrZe: TxCoN.Alignment = 2
Case Tex_AusrBl: TxCoN.Alignment = 3
Case Tex_FntAu4: TxCoN.FontName = TxStr
Case Tex_FntGr4: TxCoN.FontSize = CLng(TxStr)
Case Tex_TexMar:
    If TxCoN.ControlChars = False Then
        TxCoN.ControlChars = True
        CmAcs(Tex_TexMar).Checked = True
    Else
        TxCoN.ControlChars = False
        CmAcs(Tex_TexMar).Checked = False
    End If
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing

MTxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxFA " & Err.Number
Resume Next

End Sub
Private Sub FTxFF(ByVal TxFun As Integer)
On Error GoTo PoErr

Dim FeIdx As Long
Dim Lange As Long
Dim FeSta As Long
Dim FeEnd As Long
Dim TxGro As Integer
Dim TxIde As Integer
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim TxFnt As New StdFont
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMaiView
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont3
Set CoDia = frmMain.comDialo
Set CmAcs = CmBrs.Actions

Tit1 = "Datenfeld Entfernen"
Mld1 = "Möchten Sie das markierte Datenfeld wirklich entfernen?"

Select Case TxFun
Case Tex_ForFet:
    If TxCoN.FontBold = 0 Then
        TxCoN.FontBold = 1
    Else
        TxCoN.FontBold = 0
    End If
Case Tex_ForKur:
    If TxCoN.FontItalic = 0 Then
        TxCoN.FontItalic = 1
    Else
        TxCoN.FontItalic = 0
    End If
Case Tex_ForUnt:
    If TxCoN.FontUnderline = 0 Then
        TxCoN.FontUnderline = 1
    Else
        TxCoN.FontUnderline = 0
    End If
Case Tex_ForDur:
    If TxCoN.FontStrikethru = 0 Then
        TxCoN.FontStrikethru = 1
    Else
        TxCoN.FontStrikethru = 0
    End If
Case Tex_FntHoh:
    If TxCoN.BaseLine = 0 Then
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = 100
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 3) * 2)
        Else
            TxCoN.FontSize = (CInt(TxGro / 3) * 2) - 1
        End If
    Else
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = 0
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 2) * 3)
        Else
            TxCoN.FontSize = (CInt(TxGro / 2) * 3) - 1
        End If
    End If
Case Tex_FntTif:
    If TxCoN.BaseLine = 0 Then
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = -100
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 3) * 2)
        Else
            TxCoN.FontSize = (CInt(TxGro / 3) * 2) - 1
        End If
    Else
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = 0
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 2) * 3)
        Else
            TxCoN.FontSize = (CInt(TxGro / 2) * 3) - 1
        End If
    End If
Case Tex_FntKle:
    TxGro = TxCoN.FontSize
    TxCoN.FontSize = TxGro - 2
Case Tex_FntGro:
    TxGro = TxCoN.FontSize
    TxCoN.FontSize = TxGro + 2
Case Tex_EinzLi:
    TxCoN.IncreaseIndent
Case Tex_EinzRe:
    TxCoN.DecreaseIndent
Case Tex_KopZei:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstHeader Then
            .HeaderFooterActivate txFirstHeader
            GlHeA = txFirstHeader
        Else
            .HeaderFooterActivate txHeader
            GlHeA = txHeader
        End If
        .HeaderFooterSelect txHeader
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = TxFnt.SIZE
        .SelText = vbNullString
        .HeaderFooterSelect 0
    End With
Case Tex_FusZei:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstFooter Then
            .HeaderFooterActivate txFirstFooter
            GlHeA = txFirstFooter
        Else
            .HeaderFooterActivate txFooter
            GlHeA = txFooter
        End If
        .HeaderFooterSelect txFooter
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = TxFnt.SIZE
        .SelText = vbNullString
        .HeaderFooterSelect 0
    End With
Case Tex_KopDat:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        .HeaderFooterActivate txMainText
        GlHeA = 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstHeader Then
            .HeaderFooterActivate txFirstHeader
            GlHeA = txFirstHeader
        Else
            .HeaderFooterActivate txHeader
            GlHeA = txHeader
        End If
        .HeaderFooterSelect txHeader
        .Alignment = 1
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = TxFnt.SIZE
        .FieldInsert "<" & GlSer(74, 0) & ">" 'Tagesdatum
        FeIdx = .FieldCurrent
        .FieldType(FeIdx) = txFieldStandard
        .FieldData(FeIdx) = GlSer(74, 1)
        .FieldEditAttr(FeIdx) = &H2 + &H10
        FeSta = .FieldStart
        FeEnd = .FieldEnd
        .SelStart = FeEnd
        .SelLength = 1
        .SelText = Space$(1)
        .FieldChangeable = False
        .FieldDeleteable = True
        .HeaderFooterSelect 0
    End With
    CmAcs(Tex_KopDat).Enabled = False
Case Tex_FusZal:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        .HeaderFooterActivate txMainText
        GlHeA = 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstFooter Then
            .HeaderFooterActivate txFirstFooter
            GlHeA = txFirstFooter
        Else
            .HeaderFooterActivate txFooter
            GlHeA = txFooter
        End If
        .HeaderFooterSelect txFooter
        .Alignment = 1
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = 8
        .SelText = "Seite: "
        .SelLength = 0
        .SelStart = 7
        .FieldInsert vbNullString
        FeIdx = .FieldCurrent
        .FieldType(FeIdx) = txFieldPageNumber
        .FieldEditAttr(FeIdx) = &H2 + &H10
        FeSta = .FieldStart
        FeEnd = .FieldEnd
        .SelStart = FeEnd
        .SelLength = 1
        .SelText = Space$(1)
        .FieldChangeable = False
        .FieldDeleteable = True
        .HeaderFooterSelect 0
    End With
    CmAcs(Tex_FusZal).Enabled = False
Case Tex_DaFeLo:
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
        FeIdx = TxCoN.FieldAtInputPos
        If FeIdx > 0 Then
            TxCoN.FieldCurrent = FeIdx
            TxCoN.FieldDelete True
        End If
    End If
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing
Set CoDia = Nothing

MTxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxFF " & Err.Number
Resume Next

End Sub
Private Sub FTxCo(ByVal TxFun As Integer, Optional ByVal ColID As Long)
On Error GoTo PoErr

Dim ImVer As String
Dim AktZa As Integer

Set FM = frmMaiView
Set TxCoN = FM.TexCont3

Select Case TxFun
Case Tex_TexCut: TxCoN.Clip 1
Case Tex_TexCop: TxCoN.Clip 2
Case Tex_TexEin: TxCoN.Paste 5
Case Tex_Suchen: TxCoN.FindReplace 1
Case Tex_Ersetz: TxCoN.FindReplace 2
Case Tex_TxUndo: TxCoN.Undo
Case Tex_TxRedo: TxCoN.Redo
Case Tex_Zeiche: TxCoN.FontDialog
Case Tex_Absatz: TxCoN.ParagraphDialog
Case Tex_Aufzah: TxCoN.ListAttrDialog
Case Tex_Numeri: TxCoN.ListAttrDialog
Case Tex_TabEin: If TxCoN.TableCanInsert = True Then TxCoN.TableInsertDialog
Case Tex_TabAtr: If TxCoN.TableCanChangeAttr = True Then TxCoN.TableAttrDialog
Case Tex_SpEiRe: If TxCoN.TableCanInsertColumn = True Then TxCoN.TableInsertColumn txTableInsertAfter
Case Tex_SpEiLi: If TxCoN.TableCanInsertColumn = True Then TxCoN.TableInsertColumn txTableInsertInFront
Case Tex_ZeEiUn: If TxCoN.TableCanInsertLines = True Then TxCoN.TableInsertLines txTableInsertAfter, 1
Case Tex_ZeEiOb: If TxCoN.TableCanInsertLines = True Then TxCoN.TableInsertLines txTableInsertInFront, 1
Case Tex_SpalLo: If TxCoN.TableCanDeleteColumn = True Then TxCoN.TableDeleteColumn
Case Tex_ZeilLo: If TxCoN.TableCanDeleteLines = True Then TxCoN.TableDeleteLines
Case Tex_FaVor1: TxCoN.ForeColor = ColID
Case Tex_FaVor2: TxCoN.ForeColor = vbBlack
Case Tex_FaHin1: TxCoN.TextBkColor = ColID
Case Tex_FaHin2: TxCoN.TextBkColor = vbWhite
Case Tex_ForVor: TxCoN.SectionFormatDialog 0
Case Tex_KopFus: TxCoN.SectionFormatDialog 1
Case Tex_ForSpl: TxCoN.SectionFormatDialog 2
Case Tex_ForSty: TxCoN.StyleDialog
End Select

MTxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxCO " & Err.Number
Resume Next

End Sub
Private Sub FTxDa(ByVal CoTex As String)
On Error GoTo PoErr
'Datenfeld Einfügen

Dim FlIdx As Long
Dim FeSta As Long
Dim FeEnd As Long
Dim FlDat As String
Dim AktZa As Integer

Set FM = frmMaiView
Set TxCoN = FM.TexCont3

For AktZa = 1 To UBound(GlSer)
    If CoTex = GlSer(AktZa, 0) Then
        FlDat = GlSer(AktZa, 1)
        Exit For
    End If
Next AktZa

With TxCoN
    .FieldInsert "<" & CoTex & ">"
    FlIdx = .FieldCurrent
    If FlIdx > 0 Then
        .FieldType(FlIdx) = txFieldStandard
        .FieldData(FlIdx) = FlDat
        .FieldEditAttr(FlIdx) = &H2 + &H10
        .FieldChangeable = False
        .FieldDeleteable = True
        FeSta = .FieldStart
        FeEnd = .FieldEnd
        .SelStart = FeEnd
        .SelLength = 1
        .SelText = Space$(1)
    End If
End With

MTxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxDa " & Err.Number
Resume Next

End Sub
Private Sub FTxFZ(ByVal TxFun As Integer)
On Error GoTo PoErr
'Zeilenabstand

Set FM = frmMaiView
Set TxCoN = FM.TexCont3

Select Case TxFun
Case Tex_ZeiAb1: TxCoN.LineSpacing = 100
Case Tex_ZeiAb2: TxCoN.LineSpacing = 120
Case Tex_ZeiAb3: TxCoN.LineSpacing = 130
Case Tex_ZeiAb4: TxCoN.LineSpacing = 150
Case Tex_ZeiAb5: TxCoN.LineSpacing = 200
Case Tex_ZeiAb6: TxCoN.LineSpacing = 250
Case Tex_ZeiAb7: TxCoN.LineSpacing = 300
End Select

MTxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxFZ " & Err.Number
Resume Next

End Sub
Private Sub btnEMBCC_Click()
    GlMaE = 3
    frmAdrSuch.Show vbModal
End Sub
Private Sub btnEmCCM_Click()
    GlMaE = 2
    frmAdrSuch.Show vbModal
End Sub
Private Sub btnEmEmp_Click()
    GlMaE = 1
    frmAdrSuch.Show vbModal
End Sub

Private Sub cmbEmBCC_GotFocus()
    Me.cmbEmBCC.SelStart = 0
    Me.cmbEmBCC.SelLength = Len(Me.cmbEmBCC.Text)
End Sub
Private Sub cmbEmEmp_GotFocus()
    Me.cmbEmEmp.SelStart = 0
    Me.cmbEmEmp.SelLength = Len(Me.cmbEmEmp.Text)
End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlAkK = False Then
        If Control.id = XTP_ID_RIBBONCONTROLTAB Then
            FTabu
        Else
            Select Case Control.id
            Case Tex_FaVor1: FTool Control.id, Control.Color
            Case Tex_FaHin1: FTool Control.id, Control.Color
            Case Tex_FntAu4: FTool Control.id, , Control.Text
            Case Tex_FntGr4: FTool Control.id, , Control.Text
            Case Tex_DaFeAd: FTool Control.id, , Control.Text
            Case Else: FTool Control.id
            End Select
        End If
    End If
End Sub
Private Sub Form_Load()

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 20000
    .ClientMaxWidth = 20000
    .ClientMinHeight = 5000
    .ClientMinWidth = 14000
End With

FMenu

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    GlMaE = 0
    FDele
    FClos
    GlMaY = 0 'Emailflyoutfenster Mailindex
    Set frmMaiView = Nothing
End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlAkK = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    MaPos
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub

Private Sub TexCont3_Error(Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean)

Set TxCoN = Me.TexCont3

GlTxE = Scode 'Textcontrol Errorcode

If GlTxE <> 0 Then
    CancelDisplay = True
    TxCoN.Text = "<Es ist nicht möglich, den Text dieser Email darzustellen>"
End If

End Sub

Private Sub TexCont3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Set FM = frmMaiView
Set TxCoN = FM.TexCont3

If Shift = vbCtrlMask Then
    If KeyCode = vbKeyV Then
        KeyCode = 0
        TxCoN.Paste 5
    End If
End If

End Sub

Private Sub TexCont3_PosChange()
    If GlAkK = False Then
        MTxFo
    End If
End Sub
Private Sub txtEmBet_GotFocus()
    Me.txtEmBet.SelStart = 0
    Me.txtEmBet.SelLength = Len(Me.txtEmBet.Text)
End Sub


Private Sub txtEmCCM_GotFocus()
    Me.txtEmCCM.SelStart = 0
    Me.txtEmCCM.SelLength = Len(Me.txtEmCCM.Text)
End Sub


Private Sub CmSta_SwitchPaneClick(ByVal Pane As XtremeCommandBars.StatusBarSwitchPane, ByVal Switch As Long)
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMaiView
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

With CmSta
    Set CmSwi = .FindPane(Pane.id)
    If Pane.id = 7 Then
        MaNav Switch
    End If
End With

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SwitchPaneClick " & Err.Number
Resume Next

End Sub

