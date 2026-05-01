VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmAnzahl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Zahlung Eintragen"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   30
      Top             =   7000
      Width           =   7200
      _Version        =   1048579
      _ExtentX        =   12700
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   5200
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Abbrechen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeite 
         Default         =   -1  'True
         Height          =   400
         Left            =   3800
         TabIndex        =   16
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Weiter"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   2500
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hilfe"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8000
      Width           =   80
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   495
      Left            =   500
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
      _Version        =   1048579
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   7000
      Left            =   300
      TabIndex        =   1
      Top             =   0
      Width           =   6600
      _Version        =   1048579
      _ExtentX        =   11642
      _ExtentY        =   12347
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   345
         Left            =   2820
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1540
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu1"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   345
         Left            =   3090
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   1540
         Width           =   345
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbKonto 
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   3340
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   1500
         TabIndex        =   3
         Top             =   1540
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkGutha 
         Height          =   240
         Left            =   3060
         TabIndex        =   14
         Top             =   6360
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   422
         _StockProps     =   79
         Caption         =   "von Guthaben abziehen"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbGegen 
         Height          =   315
         Left            =   1500
         TabIndex        =   7
         Top             =   2740
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtKomme 
         Height          =   350
         Left            =   1500
         TabIndex        =   10
         Top             =   4540
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   5140
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   1500
         TabIndex        =   12
         Top             =   5740
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtGutha 
         Height          =   350
         Left            =   1500
         TabIndex        =   13
         Top             =   6340
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cmbStKto 
         Height          =   315
         Left            =   1500
         TabIndex        =   9
         Top             =   3940
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbZaTex 
         Height          =   315
         Left            =   1500
         TabIndex        =   6
         Top             =   2140
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtAnzah 
         Height          =   350
         Left            =   1500
         TabIndex        =   2
         Top             =   940
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   500
         Left            =   300
         TabIndex        =   29
         Top             =   100
         Width           =   6200
         _Version        =   1048579
         _ExtentX        =   10936
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   $"frmAnzahl.frx":0000
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         BorderStyle     =   0  'Transparent
         Height          =   650
         Left            =   0
         Top             =   0
         Width           =   6610
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   220
         Left            =   300
         TabIndex        =   28
         Top             =   2190
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zahlungstext :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   225
         Left            =   795
         TabIndex        =   27
         Top             =   1590
         Width           =   600
         _Version        =   1048579
         _ExtentX        =   1058
         _ExtentY        =   386
         _StockProps     =   79
         Caption         =   "Datum :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   220
         Left            =   400
         TabIndex        =   26
         Top             =   990
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Betrag :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   220
         Left            =   300
         TabIndex        =   25
         Top             =   2790
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Geldkonto :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   220
         Left            =   300
         TabIndex        =   24
         Top             =   3390
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Erlöskonto :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   225
         Left            =   300
         TabIndex        =   23
         Top             =   4590
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Kommentar :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   225
         Left            =   300
         TabIndex        =   22
         Top             =   5190
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Mandant :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab09 
         Height          =   225
         Left            =   300
         TabIndex        =   21
         Top             =   5790
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Mitarbeiter :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab10 
         Height          =   225
         Left            =   300
         TabIndex        =   20
         Top             =   6390
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Guthaben :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab11 
         Height          =   220
         Left            =   300
         TabIndex        =   19
         Top             =   3990
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Steuerkonto :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin VB.Shape shpShap1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   650
      Left            =   0
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmAnzahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Lbl01 As XtremeSuiteControls.Label
Private Lbl05 As XtremeSuiteControls.Label
Private Lbl06 As XtremeSuiteControls.Label
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxKom As XtremeSuiteControls.FlatEdit
Private TxGut As XtremeSuiteControls.FlatEdit
Private TxAnz As XtremeSuiteControls.FlatEdit
Private CmTex As XtremeSuiteControls.ComboBox
Private CmKto As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmStu As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private ChGut As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private MoKal As XtremeCalendarControl.DatePicker
Private Kale4 As XtremeCalendarControl.DatePicker
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow

Private mPaNr As Long
Private mReNr As Long
Private mReBe As Double
Private mAnBe As Double
Private mTeBe As Double
Private mSeBe As Double
Private mBeBe As Double
Private mReZi As Integer
Private FoLad As Boolean

Private clFen As clsFenster

Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const KEYEVENTF_KEYUP = &H2
Private Const CB_FINDSTRING = &H14C&
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    TxDa1.Text = NeuDa
    With MoKal
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
    If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    TxDa1.Text = NeuDa
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FEinf()
On Error GoTo SuErr

Dim ReDat As Date
Dim AnzDa As Date
Dim FaDa1 As Date
Dim FaDa2 As Date
Dim TerDa As Date
Dim Frage As Long
Dim ErKId As Long
Dim GeKId As Long
Dim StKID As Long
Dim ErKNr As Long
Dim GeKNr As Long
Dim StKNr As Long
Dim MaNum As Long
Dim MiNum As Long
Dim TerNr As Long
Dim MasNr As Long
Dim RowNr As Long
Dim ManNr As Long
Dim MitNr As Long
Dim BuJah As Long
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim ReTyp As String
Dim ErKBe As String
Dim GeKBe As String
Dim StKBe As String
Dim AnzBe As Double
Dim GutBe As Double
Dim ZaNam As String
Dim GuiID As String
Dim EiTex As String
Dim KoTex As String
Dim ReStr As String
Dim PlStr As String
Dim PaStr As String
Dim ZaZil As Integer
Dim Posit As Integer
Dim AktZa As Integer
Dim Mahnb As Boolean
Dim GuVer As Boolean
Dim Kasse As Boolean
Dim Mld1, Tit1 As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmAnzahl
Set TxAnz = FM.txtAnzah
Set TxKom = FM.txtKomme
Set CmKto = FM.cmbKonto
Set CmGeg = FM.cmbGegen
Set CmMan = FM.cmbManda
Set CmMit = FM.cmbMitar
Set TxDa1 = FM.txtDatu1
Set MoKal = FM.dtpDatu1
Set CmTex = FM.cmbZaTex
Set ChGut = FM.chkGutha
Set PuBu1 = Me.btnWeite
Set Lbl01 = Me.lblLab01
Set RpCo1 = frmMain.repCont1
Set RpCo2 = frmMain.repCont2
Set RpCo4 = frmMain.repCont4
Set RpCo3 = frmMain.repCont3

PuBu1.Enabled = False
Screen.MousePointer = vbHourglass

If IsDate(TxDa1.Text) Then
    AnzDa = TxDa1.Text
    BuJah = Year(TxDa1.Text)
Else
    AnzDa = Date
    BuJah = Year(Date)
End If

If CmTex.Text <> vbNullString Then
    EiTex = CmTex.Text
Else
    EiTex = "Zahlung"
End If

If TxKom.Text <> vbNullString Then
    KoTex = TxKom.Text
Else
    KoTex = vbNullString
End If

If ChGut.Value = xtpChecked Then
    GuVer = True
End If

ErKId = CmKto.ItemData(CmKto.ListIndex)
GeKId = CmGeg.ItemData(CmGeg.ListIndex)
StKID = CmStu.ItemData(CmStu.ListIndex)
If GlKnF = True Then 'Sachkontenformatierung sechsstellig
    ErKNr = Left$(CmKto.Text, 6)
    GeKNr = Left$(CmGeg.Text, 6)
    StKNr = Left$(CmStu.Text, 6)
    ErKBe = Mid$(CmKto.Text, 8, Len(CmKto.Text) - 7)
    GeKBe = Mid$(CmGeg.Text, 8, Len(CmGeg.Text) - 7)
    StKBe = Mid$(CmStu.Text, 8, Len(CmStu.Text) - 7)
Else
    ErKNr = Left$(CmKto.Text, 4)
    GeKNr = Left$(CmGeg.Text, 4)
    StKNr = Left$(CmStu.Text, 4)
    ErKBe = Mid$(CmKto.Text, 6, Len(CmKto.Text) - 5)
    GeKBe = Mid$(CmGeg.Text, 6, Len(CmGeg.Text) - 5)
    StKBe = Mid$(CmStu.Text, 6, Len(CmStu.Text) - 5)
End If

ManNr = CmMan.ItemData(CmMan.ListIndex)
MitNr = CmMit.ItemData(CmMit.ListIndex)
ZaNam = GlZah(mReZi, 1)
Mahnb = CBool(GlZah(mReZi, 3))
GuiID = CreateID("K")
Tit1 = "Zahlung Eintragen"

If TxAnz.Text <> vbNullString Then
    AnzBe = CDbl(Format$(TxAnz.Text, GlWa1))
Else
    AnzBe = 0
End If

If TxAnz.Text = vbNullString Then
    Mld1 = "Es ist kein Zahlbetrag vorhanden!"
    Screen.MousePointer = vbNormal
    SPopu Tit1, Mld1, IC48_Information
    Exit Sub
ElseIf AnzBe <= 0 Then
    Mld1 = "Der Zahlbetrag ist zu gering!"
    Screen.MousePointer = vbNormal
    SPopu Tit1, Mld1, IC48_Information
    Exit Sub
ElseIf GuVer = True And mAnBe > 0 Then
    Screen.MousePointer = vbNormal
    Mld1 = "Wenn ein bestehendes Guthaben verrechnen werden soll, darf keine weitere Zahlung in der Rechnung vorhanden sein."
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
    Exit Sub
ElseIf AnzBe > CDbl(mReBe - mAnBe) Then
    If WindowLoad("frmAdress") = False Then
        Select Case GlBut
        Case RibTab_Adressen:
        Case RibTab_Ter_Listen:
        Case RibTab_Ter_Akont:
        Case Else:
            If CDbl(mReBe - mAnBe) <= 0 Then
                Mld1 = "Der Zahlbetrag darf kein Minuszeichen enthlaten!"
                Screen.MousePointer = vbNormal
                SPopu Tit1, Mld1, IC48_Information
                Exit Sub
            ElseIf (AnzBe - CDbl(mReBe - mAnBe)) > CDbl(0.01) Then
                Mld1 = "Der Zahlbetrag übersteigt den Rechnungsbetrag!"
                Screen.MousePointer = vbNormal
                SPopu Tit1, Mld1, IC48_Information
                Exit Sub
            Else
                Mld1 = "Der überschüssige Zahlbetrag wird dem Patienten gutgeschrieben"
                SPopu Tit1, Mld1, IC48_Information
            End If
        End Select
    End If
ElseIf Mahnb = False Then
    Mld1 = "Dieser Rechnung ist bereits das Zahlungsziel: " & ZaNam & " zugeordnet. Möchten Sie wirklich eine Zahlung einfügen?"
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage <> 6 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
End If

If WindowLoad("frmAdress") = True Then
    Kasse = S_BuZal(AnzDa, AnzBe, GuiID, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, ManNr, MitNr, BuJah, EiTex, KoTex)
    DoEvents
    S_PaGu GlAId, AnzBe
    DoEvents
    GutBe = S_AdIdx(GlAId, "Guthaben")
    frmAdress.txtS1F33.Text = Format$(GutBe, GlWa1)
    GlAdS = False
    DoEvents
    S_KrLa
    Unload Me
    DoEvents
Else
    Unload Me
    DoEvents
    
    Select Case GlBut
    Case RibTab_Adressen:
        Set RpCls = RpCo2.Columns
        Set RpSel = RpCo2.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            mPaNr = AdAry(Adr_ID0, RpRow.Index)
        Else
            mPaNr = 0
        End If
        Kasse = S_BuZal(AnzDa, AnzBe, GuiID, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, ManNr, MitNr, BuJah, EiTex, KoTex)
        DoEvents
        S_PaGu mPaNr, AnzBe
        DoEvents
    Case RibTab_Abrechnung:
        Set RpCls = RpCo3.Columns
        Set RpSel = RpCo3.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            Set RpCol = RpCls.Find(Rec_RechNr)
            ReStr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Rec_Type)
            ReTyp = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Rec_IDKurz)
            PaStr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Rec_Datum)
            ReDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Rec_IDP)
            If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                MaNum = RpRow.Record(RpCol.ItemIndex).Value
            End If
            Set RpCol = RpCls.Find(Rec_IDM)
            If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                MiNum = RpRow.Record(RpCol.ItemIndex).Value
            End If
            If Len(ReStr) > 3 Then
                If GuVer = False Then 'Guthaben verrechnen
                    Kasse = S_BuZal(AnzDa, AnzBe, GuiID, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, ManNr, MitNr, BuJah, EiTex, KoTex)
                    DoEvents
                Else
                    S_BuRe mReNr, ReStr, PaStr
                End If
                S_KrAn AnzDa, AnzBe, GuiID, EiTex, GuVer, KoTex
            Else
                ReStr = S_ReVo(Date, ReTyp, MaNum, MiNum, True)
                DoEvents
                TeTit = "Zahlung eintragen"
                TeMai = "Möchten Sie eine Rechnungsnummer erzeugen? (" & ReStr & ")"
                TeInh = "Diese Rechnung enthält noch keine gültige Rechnungsnummer. Ohne Rechnungsnummer kann keine Erlösbuchung abgeleitet werden."
                TeFus = "Es kann nun automatisch eine gültige Rechnungsnummer (" & ReStr & ") erzeugt und somit die Zahlung eintragen werden."
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
                If GlMes = 33565 Then
                    Posit = InStrRev(ReStr, "-", -1, 1)
                    If Posit > 0 Then
                        PlStr = Mid$(ReStr, Posit + 1, Len(ReStr) - Posit)
                    End If
                    DoEvents
                    DBCmEx2 "qrySimReNu", "@IdStr", "@IdxNr", ReStr, mReNr
                    DoEvents
                    If PlStr <> vbNullString Then
                        DBCmEx2 "qrySimReNp", "@IdStr", "@IdxNr", Format$(PlStr, "00000000"), mReNr
                    End If
                    DoEvents
                    SUpAb RowNr
                    DoEvents
                    If GuVer = False Then 'Guthaben verrechnen
                        Kasse = S_BuZal(AnzDa, AnzBe, GuiID, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, ManNr, MitNr, BuJah, EiTex, KoTex)
                        DoEvents
                    Else
                        S_BuRe mReNr, ReStr, PaStr
                    End If
                    S_KrAn AnzDa, AnzBe, GuiID, EiTex, GuVer, KoTex
                End If
            End If
        End If
    Case RibTab_Ter_Akont:
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Ter_ID2)
                TerNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Ter_ID0)
                mPaNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Ter_VonDat)
                TerDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Ter_TerBet)
                mTeBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Ter_SerBet)
                mSeBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Ter_BezBet)
                mBeBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Ter_MasTer)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    MasNr = RpRow.Record(RpCol.ItemIndex).Value
                Else
                    MasNr = 0
                End If
                Set RpCol = RpCls.Find(Ter_Fallig1)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    FaDa1 = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                End If
                Set RpCol = RpCls.Find(Ter_Fallig2)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    FaDa2 = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                Else
                    For AktZa = 1 To UBound(GlZah) 'Standardzahlungsziel
                        If GlZah(AktZa, 0) = GlStZ Then
                            ZaZil = GlZah(AktZa, 2)
                            Exit For
                        End If
                    Next AktZa
                    FaDa2 = DateAdd("d", ZaZil, TerDa)
                End If
                Ter_Anz MasNr, FaDa2 'Anpassen der Fälligkeit
                DoEvents
                Kasse = S_BuZal(AnzDa, AnzBe, GuiID, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, ManNr, MitNr, BuJah, EiTex, KoTex)
                DoEvents
                S_PaTe AnzBe + mBeBe, TerNr, MasNr
                DoEvents
                S_PaGu mPaNr, AnzBe
            End If
        End If
        DoEvents

        Set RpSel = RpCo1.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpTe RowNr
        End If
    End Select
End If
DoEvents

Screen.MousePointer = vbNormal

If GlTSe > 0 Then 'TSE Aktiviert
    If Kasse = True Then
        GlBu2 = RibTab_Belegmodul
        STaSe ShoCut_Kranken, RibTab_Belegmodul
    End If
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEinf " & Err.Number
Resume Next

End Sub
Private Sub FGeKo()
On Error GoTo SuErr

Dim IdxNr As Integer
Dim AktZa As Integer
Dim IdBnk As Integer

Set CmTex = Me.cmbZaTex
Set CmGeg = Me.cmbGegen

IdxNr = CmTex.ListIndex

IdBnk = GlZTe(IdxNr + 1, 1)

If IdBnk > 0 Then
    For AktZa = 1 To UBound(GlGeK)
        If IdBnk = CInt(GlGeK(AktZa, 0)) Then
            CmGeg.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
Else
    If IdxNr > 4 Then
        For AktZa = 1 To UBound(GlGeK)
            If CBool(GlGeK(AktZa, 5)) = False Then
                CmGeg.ListIndex = AktZa - 1
                Exit For
            End If
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK)
            If CBool(GlGeK(AktZa, 5)) = True Then
                CmGeg.ListIndex = AktZa - 1
                Exit For
            End If
        Next AktZa
    End If
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FGeKo " & Err.Number
Resume Next

End Sub

Private Sub FGuth()
On Error GoTo SuErr

Dim AnzDa As Date
Dim GuBet As Double
Dim TaBet As Double
Dim TmStr As String
Dim Posit As Integer
Dim AktZa As Integer
Dim TeWer As Variant

Set TxAnz = Me.txtAnzah
Set TxDa1 = Me.txtDatu1
Set TxKom = Me.txtKomme
Set CmTex = Me.cmbZaTex
Set ChGut = Me.chkGutha
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen

If IsDate(TxDa1.Text) Then
    AnzDa = TxDa1.Text
Else
    AnzDa = Date
End If

If ChGut.Value = xtpChecked Then
    TeWer = S_AdIdx(mPaNr, "Guthaben")
    If TeWer <> vbNullString Then
        GuBet = CDbl(Format$(TeWer, GlWa1))
        If GuBet > 0 Then
            TxAnz.Text = GlWa2
            For AktZa = 1 To UBound(GlZTe)
                TmStr = LCase(GlZTe(AktZa, 3)) 'Zahlungstexte
                Posit = InStr(1, TmStr, "guthaben", 1)
                If Posit > 0 Then
                    Exit For
                End If
            Next AktZa
            CmTex.ListIndex = AktZa - 1
            CmTex.Enabled = False
            TxAnz.Enabled = False
            CmKto.Enabled = False
            CmGeg.Enabled = False
            If (mReBe - mAnBe) >= GuBet Then
                TxAnz.Text = Format$(GuBet, GlWa1)
            ElseIf (mReBe - mAnBe) < GuBet Then
                TxAnz.Text = Format$((mReBe - mAnBe), GlWa1)
            Else
                TxAnz.Text = Format$(mReBe - mAnBe, GlWa1)
            End If
        End If
    End If
Else
    TxAnz.Text = GlWa2
    For AktZa = 1 To UBound(GlZTe)
        If CBool(GlZTe(AktZa, 4)) = True Then
            Exit For
        End If
    Next AktZa
    CmTex.ListIndex = AktZa - 1
    DoEvents
    CmTex.Enabled = True
    TxAnz.Enabled = True
    CmKto.Enabled = True
    CmGeg.Enabled = True
    TaBet = S_KrTa(mReNr, AnzDa)
    TxAnz.Text = Format$(mReBe - mAnBe, GlWa1)
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FGuth " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim NeuDa As Date
Dim DayFi As Date
Dim DayLa As Date
Dim ManNr As Long
Dim MitNr As Long
Dim TmpNr As Long
Dim StaKt As Long
Dim StaGe As Long
Dim TmStr As String
Dim GuBet As Double
Dim TaBet As Double
Dim AktZa As Integer
Dim AktKo As Integer
Dim StaRa As Integer
Dim Posit As Integer
Dim IdStK As Integer
Dim EiKon As Boolean
Dim TmKo1 As Boolean
Dim TeWer As Variant
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set ImMan = FM.imgManag
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set TxAnz = Me.txtAnzah
Set TxGut = Me.txtGutha
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen
Set CmStu = Me.cmbStKto
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set MoKal = Me.dtpDatu1
Set TxDa1 = Me.txtDatu1
Set TxKom = Me.txtKomme
Set CmTex = Me.cmbZaTex
Set ChGut = Me.chkGutha
Set PuBu1 = Me.btnDatu1
Set TxDa2 = FM.txtDatu1 'Datum Eingabezeile
Set Lbl01 = Me.lblLab01
Set Lbl05 = Me.lblLab05
Set Lbl06 = Me.lblLab06

If WindowLoad("frmAdress") = True Then
    NeuDa = Date
Else
    If GlBut = RibTab_Adressen Then
        NeuDa = Date
    ElseIf GlBut = RibTab_Ter_Listen Then
        NeuDa = Date
    ElseIf GlBut = RibTab_Ter_Akont Then
        NeuDa = Date
    ElseIf GlBut = RibTab_Ter_Warte Then
        NeuDa = Date
    Else
        If IsDate(TxDa2.Text) Then
            NeuDa = TxDa2.Text
        Else
            NeuDa = Date
        End If
    End If
End If

For AktZa = 1 To UBound(GlZTe)
    With CmTex
        .AddItem GlZTe(AktZa, 3)
        .ItemData(AktZa - 1) = GlZTe(AktZa, 0)
    End With
Next AktZa

For AktZa = 1 To UBound(GlZTe)
    If CBool(GlZTe(AktZa, 4)) = True Then
        Exit For
    End If
Next AktZa
CmTex.ListIndex = AktZa - 1

With MoKal
    .AllowNoncontinuousSelection = False
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    If GlSty = 8 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    ElseIf GlSty = 7 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    Else
        .BorderStyle = xtpDatePickerBorderOffice
    End If
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    .MaxSelectionCount = 1
    .RightToLeft = False
    .ShowNoneButton = True
    .ShowNonMonthDays = True
    .ShowTodayButton = True
    .ShowWeekNumbers = False
    .TextNoneButton = "Keine"
    .TextTodayButton = "Heute"
    Select Case GlSty
    Case 8: .VisualTheme = xtpCalendarThemeResource
    Case 7: .VisualTheme = xtpCalendarThemeResource
    Case Else: .VisualTheme = xtpCalendarThemeResource
    End Select
    .PaintManager.ButtonTextColor = -2147483640
    .PaintManager.ControlBackColor = -2147483643
    .PaintManager.DayBackColor = -2147483643
    .PaintManager.DayTextColor = -2147483640
    .PaintManager.DaysOfWeekBackColor = -2147483643
    .PaintManager.DaysOfWeekTextColor = -2147483640
    .PaintManager.ListControlBackColor = -2147483643
    .PaintManager.ListControlTextColor = -2147483640
    .PaintManager.NonMonthDayBackColor = -2147483643
    .PaintManager.NonMonthDayTextColor = -2147483640
    .PaintManager.SelectedDayBackColor = GlFac
    .PaintManager.SelectedDayTextColor = -2147483640
    .PaintManager.WeekNumbersBackColor = -2147483643
    .PaintManager.WeekNumbersTextColor = -2147483640
    .PaintManager.MonthHeaderBackColor = GlMoB
    .MonthDelta = 1
    .YearsTriangle = False
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Behandlungstag"
End With

With CmKto
    If GlMVo = False Then 'mandantenbezogene Vorgaben verwenden
        For AktZa = 1 To UBound(GlErK)
            .AddItem GlErK(AktZa, 1)
            .ItemData(.NewIndex) = GlErK(AktZa, 0) '[IDK]
        Next AktZa
    End If
End With

With CmGeg
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            .AddItem GlGeK(AktZa, 3)
            .ItemData(.NewIndex) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    .AddItem GlSaK(AktKo, 3)
                    .ItemData(.NewIndex) = GlSaK(AktKo, 6) '[IDB]
                End If
            Next AktKo
        Next AktZa
        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(.NewIndex) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
        End If
    End If
End With

With CmStu
    For AktZa = 1 To UBound(GlSaU) 'Sachkonten mit Steuerkontenzuordnung
        .AddItem GlSaU(AktZa, 3)
        .ItemData(AktZa - 1) = GlSaU(AktZa, 6) '[IDI]
    Next AktZa
End With

With CmMan
    For AktZa = 1 To UBound(GlMan)
        .AddItem GlMan(AktZa, 1)
        .ItemData(.NewIndex) = GlMan(AktZa, 2)
    Next AktZa
End With

With CmMit
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        .AddItem GlMiA(AktZa, 1)
        .ItemData(.NewIndex) = GlMiA(AktZa, 2)
    Next AktZa
End With

IdStK = SCmb(CmStu, GlSKo) 'Standardsteuerkonto

If IdStK >= 0 Then
    CmStu.ListIndex = IdStK
Else
    CmStu.ListIndex = 0
End If

If WindowLoad("frmAdress") = True Then
    Set RpCls = RpCo2.Columns
    Set RpSel = RpCo2.SelectedRows
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        mPaNr = AdAry(Adr_ID0, RpRow.Index)
    Else
        mPaNr = 0
    End If
    mReBe = 0
    mReZi = 1
Else
    Select Case GlBut
    Case RibTab_Abrechnung:
            Set RpCls = RpCo3.Columns
            Set RpSel = RpCo3.SelectedRows
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rec_ID1)
                mReNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_ID0)
                mPaNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_IDM)
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_IDP)
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Betrag)
                mReBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Rec_Bezahlt)
                mAnBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Rec_IDZ)
                mReZi = RpRow.Record(RpCol.ItemIndex).Value
            Else
                mReBe = 0
                mReZi = 1
                ManNr = GlMan(GlSMa, 2)
                MitNr = GlMiA(GlSmI, 2)
            End If
    Case RibTab_Rechnungen:
            Set RpCls = RpCo4.Columns
            Set RpSel = RpCo4.SelectedRows
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rec_ID1)
                mReNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_ID0)
                mPaNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_IDM)
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_IDP)
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Betrag)
                mReBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Rec_Bezahlt)
                mAnBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Rec_IDZ)
                mReZi = RpRow.Record(RpCol.ItemIndex).Value
            Else
                mReBe = 0
                mReZi = 1
                ManNr = GlMan(GlSMa, 2)
                MitNr = GlMiA(GlSmI, 2)
            End If
    Case RibTab_Adressen:
            Set RpCls = RpCo2.Columns
            Set RpSel = RpCo2.SelectedRows
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                mPaNr = AdAry(Adr_ID0, RpRow.Index)
                ManNr = AdAry(Adr_IDP, RpRow.Index)
                MitNr = GlMiA(GlSmI, 2)
            Else
                mPaNr = 0
                ManNr = GlMan(GlSMa, 2)
                MitNr = GlMiA(GlSmI, 2)
            End If
            mReBe = 0
            mReZi = 1
    Case RibTab_Ter_Akont:
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Ter_ID0)
                mPaNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Ter_IDP)
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Ter_IDM)
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Ter_TerBet)
                mTeBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Ter_SerBet)
                mSeBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                Set RpCol = RpCls.Find(Ter_BezBet)
                mBeBe = CDbl(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
            Else
                mPaNr = 0
                ManNr = GlMan(GlSMa, 2)
                MitNr = GlMiA(GlSmI, 2)
            End If
        Else
            mPaNr = 0
            ManNr = GlMan(GlSMa, 2)
            MitNr = GlMiA(GlSmI, 2)
        End If
        mReBe = 0
        mReZi = 1
    Case Else:
        mPaNr = GlAdr
        mReBe = 0
        mReZi = 1
        ManNr = GlMan(GlSMa, 2)
        MitNr = GlMiA(GlSmI, 2)
    End Select
End If

If ManNr = 0 Then ManNr = GlMan(GlSMa, 2)
If MitNr = 0 Then MitNr = GlMiA(GlSmI, 2)

TmpNr = SCmX(CmMan, ManNr)
If TmpNr >= 0 Then
    CmMan.ListIndex = TmpNr
Else
    CmMan.ListIndex = GlSMa - 1
End If

TmpNr = SCmX(CmMit, MitNr)
If TmpNr >= 0 Then
    CmMit.ListIndex = TmpNr
Else
    CmMit.ListIndex = GlSmI - 1
End If

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = NeuDa
End With

CmStu.Enabled = GlSpB 'Umsatzsteuer Splittbuchung

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    S_KoMa 1, ManNr
    EiKon = True
    TmKo1 = True
Else
    StaKt = SCmb(CmKto, GlSE1) 'Standarderlöskonto Kasse
    If StaKt >= 0 Then
        CmKto.ListIndex = StaKt
        EiKon = True
    Else
        CmKto.ListIndex = 0
    End If
    StaGe = SCmb(CmGeg, GlGkK) 'Standardgeldkonto Kasse
    If StaGe >= 0 Then
        If CmGeg.ListCount > 0 Then
            CmGeg.ListIndex = StaGe
            TmKo1 = True
        End If
    Else
        CmGeg.ListIndex = 0
    End If
End If

If EiKon = False Then
    CmKto.ListIndex = 1
End If

If CmKto.ListCount > 0 Then
    If CmKto.ListIndex < 0 Then
        CmKto.ListIndex = 0
    End If
End If

If CmGeg.ListIndex < 0 Then
    CmGeg.ListIndex = 0
End If

If TmKo1 = False Then
    If CmGeg.ListCount > 0 Then
        CmGeg.ListIndex = 1
    End If
End If

If WindowLoad("frmAdress") = True Then
    If frmAdress.txtS1F33.Text <> vbNullString Then
        If IsNumeric(frmAdress.txtS1F33.Text) = True Then
            If CDbl(frmAdress.txtS1F33.Text) > 0 Then
                GuBet = CDbl(frmAdress.txtS1F33.Text)
            Else
                GuBet = 0
            End If
        Else
            GuBet = 0
        End If
    Else
        GuBet = 0
    End If
    TxAnz.Text = GlWa2
    TxGut.Text = Format$(GuBet, GlWa1)
Else
    If GlBut = RibTab_Adressen Then
        TxAnz.Text = GlWa2
    ElseIf GlBut = RibTab_Ter_Listen Then
        If mSeBe > 0 Then
            TxAnz.Text = Format$(mSeBe - mBeBe, GlWa1)
        ElseIf mTeBe > 0 Then
            TxAnz.Text = Format$(mTeBe - mBeBe, GlWa1)
        Else
            TxAnz.Text = GlWa2
        End If
    ElseIf GlBut = RibTab_Ter_Akont Then
        If mSeBe > 0 Then
            TxAnz.Text = Format$(mSeBe - mBeBe, GlWa1)
        ElseIf mTeBe > 0 Then
            TxAnz.Text = Format$(mTeBe - mBeBe, GlWa1)
        Else
            TxAnz.Text = GlWa2
        End If
    Else
        TaBet = S_KrTa(mReNr, NeuDa)
        TxAnz.Text = Format$(mReBe - mAnBe, GlWa1)
        TeWer = S_AdIdx(mPaNr, "Guthaben")
        If TeWer <> vbNullString Then
            GuBet = CDbl(Format$(TeWer, GlWa1))
            If GuBet > 0 Then
                TxAnz.Text = vbNullString
                For AktZa = 1 To UBound(GlZTe)
                    TmStr = LCase(GlZTe(AktZa, 3))
                    Posit = InStr(1, TmStr, "guthaben", 1)
                    If Posit > 0 Then
                        Exit For
                    End If
                Next AktZa
                CmTex.ListIndex = AktZa - 1
                CmTex.Enabled = False
                TxAnz.Enabled = False
                CmGeg.Enabled = False
                CmKto.Enabled = False
                ChGut.Value = xtpChecked
                TxGut.Text = Format$(GuBet, GlWa1)
                If (mReBe - mAnBe) >= GuBet Then
                    TxAnz.Text = Format$(GuBet, GlWa1)
                ElseIf (mReBe - mAnBe) < GuBet Then
                    TxAnz.Text = Format$((mReBe - mAnBe), GlWa1)
                Else
                    TxAnz.Text = Format$(mReBe - mAnBe, GlWa1)
                End If
            Else
                TxGut.Text = Format$(0, GlWa1)
            End If
        End If
    End If
End If

If GlBuc = True Then 'Einfache Buchhaltung verwenden
    Lbl05.Caption = "Geldkonto :"
    Lbl06.Caption = "Sachkonto :"
Else
    Lbl05.Caption = "Sollkonto :"
    Lbl06.Caption = "Habenkonto :"
End If

Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
Me.BackColor = GlBak
ChGut.BackColor = GlBak

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub

Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = TxDa1.Top + TxDa1.Height
    .Left = TxDa1.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TxDa1.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FMnd()
On Error GoTo OrErr

Dim ManNr As Long
Dim StaGe As Long

Set CmMan = Me.cmbManda
Set CmGeg = Me.cmbGegen

ManNr = CmMan.ItemData(CmMan.ListIndex)

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    S_KoMa 1, ManNr
Else
    StaGe = SCmb(CmGeg, GlGkK) 'Standardgeldkonto Kasse
    If StaGe >= 0 Then
        CmGeg.ListIndex = StaGe
    Else
        CmGeg.ListIndex = 0
    End If
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMnd " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    FKale
End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50051)
TeMai = IniGetOpt("Hilfe", 50052)
TeInh = IniGetOpt("Hilfe", 50053)
TeFus = IniGetOpt("Hilfe", 50054)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeite_Click()
    FEinf
    Unload Me
End Sub

Private Sub chkGutha_Click()
    If FoLad = False Then
        FGuth
    End If
End Sub


Private Sub cmbManda_Click()
    If FoLad = False Then
        FMnd
    End If
End Sub
Private Sub cmbZaTex_Click()
    If FoLad = False Then
        FGeKo
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
Private Sub dtpDatu1_MonthChanged()

Dim DayFi As Date
Dim DayLa As Date

Set MoKal = Me.dtpDatu1

With MoKal
    DayFi = .FirstDayOfWeek
    DayLa = .LastVisibleDay
End With

S_AbTe DayFi, DayLa

End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub

Private Sub Form_Load()
On Error Resume Next

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

FoLad = True

FInit
FGeKo
AFont Me

FoLad = False

clFen.FenVor

SFrame 1, Me.hwnd

Set clFen = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAnzahl = Nothing
End Sub

Private Sub txtAnzah_GotFocus()
    Me.txtAnzah.SelStart = 0
    Me.txtAnzah.SelLength = Len(Me.txtAnzah.Text)
End Sub

Private Sub txtAnzah_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtAnzah_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtAnzah.SelLength = 0
    Case vbKeyDown: Me.txtDatu1.SetFocus
    End Select
End Sub


Private Sub txtAnzah_LostFocus()
On Error Resume Next

Dim Betra As Double
Dim WeBet As Double

Set FM = frmAnzahl
Set TxAnz = FM.txtAnzah
Set Lbl01 = FM.lblLab01

If TxAnz.Text <> vbNullString Then
    If IsNumeric(TxAnz.Text) = True Then
        Betra = CDbl(TxAnz.Text)
        If Betra < 0 Then
            Betra = Betra * (-1)
        End If
        If Betra > mReBe Then
            If mReBe > 0 Then
                TxAnz.Text = Format$(mReBe, GlWa1)
                WeBet = Betra - (mReBe - mAnBe)
                If WeBet > CDbl(0.01) Then
                    FM.Caption = "Zahlung Eintrag - Wechselgeld: " & Format$(WeBet, GlWa1)
                End If
            Else
                TxAnz.Text = Format$(Betra, GlWa1)
            End If
        Else
            TxAnz.Text = Format$(Betra, GlWa1)
        End If
    Else
        TxAnz.Text = Format$(0, GlWa1)
    End If
End If

End Sub

Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub

Private Sub txtDatu1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtDatu1_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtDatu1.SelLength = 0
    Case vbKeyDown: Me.cmbZaTex.SetFocus
    Case vbKeyUp: Me.txtAnzah.SetFocus
    End Select
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub

Private Sub updCont1_DownClick()
On Error Resume Next

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont1_UpClick()
On Error Resume Next

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub
