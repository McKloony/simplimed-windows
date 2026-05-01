VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmNeuAuf 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Neuer Laborauftrag"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   41
      Top             =   3700
      Width           =   6600
      _Version        =   1048579
      _ExtentX        =   11642
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4600
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schließen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Height          =   400
         Left            =   3200
         TabIndex        =   43
         TabStop         =   0   'False
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
      Begin XtremeSuiteControls.PushButton btnZurück 
         Height          =   400
         Left            =   1800
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zurück"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   500
         TabIndex        =   45
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
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6500
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   6526
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListBox lstList1 
         Height          =   1600
         Left            =   500
         TabIndex        =   2
         Top             =   1800
         Width           =   5400
         _Version        =   1048579
         _ExtentX        =   9525
         _ExtentY        =   2822
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   350
         Left            =   500
         TabIndex        =   1
         Top             =   1200
         Width           =   5400
         _Version        =   1048579
         _ExtentX        =   9525
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin VB.Label Lab01 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNeuAuf.frx":0000
         Height          =   615
         Left            =   500
         TabIndex        =   39
         Top             =   100
         Width           =   5800
      End
      Begin VB.Label Lab02 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Nachname :"
         Height          =   195
         Left            =   510
         TabIndex        =   38
         Top             =   970
         Width           =   3000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3700
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   6500
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   6526
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   310
         Left            =   2100
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1920
         Width           =   310
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   547
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBefun 
         Height          =   225
         Left            =   810
         TabIndex        =   10
         Top             =   3300
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Befundungskosten"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAuftr 
         Height          =   310
         Left            =   800
         TabIndex        =   6
         Top             =   1200
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox cmbAuft 
         Height          =   310
         Left            =   3000
         TabIndex        =   11
         Top             =   1200
         Width           =   2600
         _Version        =   1048579
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMulti 
         Height          =   310
         Left            =   800
         TabIndex        =   9
         Top             =   2720
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.ComboBox cmbTage 
         Height          =   310
         Left            =   3000
         TabIndex        =   12
         Top             =   1920
         Width           =   2600
         _Version        =   1048579
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   310
         Left            =   800
         TabIndex        =   7
         Top             =   1920
         Width           =   1290
         _Version        =   1048579
         _ExtentX        =   2275
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cbmBehan 
         Height          =   315
         Left            =   3000
         TabIndex        =   13
         Top             =   2720
         Width           =   2600
         _Version        =   1048579
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant"
         Height          =   210
         Left            =   3010
         TabIndex        =   40
         Top             =   2500
         Width           =   1200
      End
      Begin VB.Label lab04 
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragsnummer"
         Height          =   210
         Left            =   810
         TabIndex        =   37
         Top             =   970
         Width           =   1500
      End
      Begin VB.Label Lab03 
         BackStyle       =   0  'Transparent
         Caption         =   "Kontrollieren Sie bitte die Auftragsnummer und stellen alle übrigen Auftragsbezogenen Daten ein. Klicken Sie dann auf Weiter."
         Height          =   615
         Left            =   300
         TabIndex        =   36
         Top             =   100
         Width           =   6000
      End
      Begin VB.Label lab05 
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragsdatum"
         Height          =   210
         Left            =   810
         TabIndex        =   35
         Top             =   1700
         Width           =   1500
      End
      Begin VB.Label Lab06 
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragstyp"
         Height          =   210
         Left            =   3010
         TabIndex        =   34
         Top             =   970
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Untersuchungszeit"
         Height          =   210
         Left            =   3010
         TabIndex        =   33
         Top             =   1700
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Steigerungsfaktor"
         Height          =   210
         Left            =   810
         TabIndex        =   32
         Top             =   2500
         Width           =   1500
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   3700
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   6500
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   6526
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtVorna 
         Height          =   310
         Left            =   800
         TabIndex        =   4
         Top             =   1920
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtName 
         Height          =   310
         Left            =   800
         TabIndex        =   3
         Top             =   1200
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtKost 
         Height          =   310
         Left            =   3800
         TabIndex        =   15
         Top             =   1200
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtGebor 
         Height          =   310
         Left            =   800
         TabIndex        =   5
         Top             =   2620
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbAbTyp 
         Height          =   310
         Left            =   3800
         TabIndex        =   16
         Top             =   1920
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGesch 
         Height          =   310
         Left            =   2300
         TabIndex        =   14
         Top             =   2620
         Width           =   1050
         _Version        =   1048579
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbGebüh 
         Height          =   310
         Left            =   3800
         TabIndex        =   17
         Top             =   2620
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.FlatEdit txtPatNu 
         Height          =   315
         Left            =   800
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3200
         Visible         =   0   'False
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin VB.Label Lab10 
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren"
         Height          =   210
         Left            =   810
         TabIndex        =   31
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Lab09 
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname"
         Height          =   210
         Left            =   810
         TabIndex        =   30
         Top             =   1700
         Width           =   1005
      End
      Begin VB.Label Lab08 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   210
         Left            =   810
         TabIndex        =   29
         Top             =   970
         Width           =   1005
      End
      Begin VB.Label Lab11 
         BackStyle       =   0  'Transparent
         Caption         =   "Geschlecht"
         Height          =   210
         Left            =   2310
         TabIndex        =   28
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Lab12 
         BackStyle       =   0  'Transparent
         Caption         =   "Befundungskosten"
         Height          =   210
         Left            =   3800
         TabIndex        =   27
         Top             =   970
         Width           =   1395
      End
      Begin VB.Label Lab13 
         BackStyle       =   0  'Transparent
         Caption         =   "Abrechnungstyp"
         Height          =   210
         Left            =   3800
         TabIndex        =   26
         Top             =   1700
         Width           =   1395
      End
      Begin VB.Label Lab14 
         BackStyle       =   0  'Transparent
         Caption         =   "Gebührentyp"
         Height          =   210
         Left            =   3800
         TabIndex        =   25
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Lab07 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNeuAuf.frx":00D3
         Height          =   615
         Left            =   300
         TabIndex        =   24
         Top             =   100
         Width           =   6000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   3700
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   6500
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   6526
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin VB.Label Lab16 
         BackStyle       =   0  'Transparent
         Caption         =   "Patienten des Auftraggebers"
         Height          =   195
         Left            =   810
         TabIndex        =   23
         Top             =   970
         Width           =   3000
      End
      Begin VB.Label Lab15 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNeuAuf.frx":015B
         Height          =   615
         Left            =   300
         TabIndex        =   22
         Top             =   100
         Width           =   6000
      End
   End
End
Attribute VB_Name = "frmNeuAuf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxNum As XtremeSuiteControls.FlatEdit
Private TxKur As XtremeSuiteControls.FlatEdit
Private TxAuf As XtremeSuiteControls.FlatEdit
Private TxNam As XtremeSuiteControls.FlatEdit
Private TxVor As XtremeSuiteControls.FlatEdit
Private TxGeb As XtremeSuiteControls.FlatEdit
Private TxKos As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private ChBef As XtremeSuiteControls.CheckBox
Private CmAuf As XtremeSuiteControls.ComboBox
Private CmGes As XtremeSuiteControls.ComboBox
Private CmAbr As XtremeSuiteControls.ComboBox
Private CmGeb As XtremeSuiteControls.ComboBox
Private CmTag As XtremeSuiteControls.ComboBox
Private CmMul As XtremeSuiteControls.ComboBox
Private CmThe As XtremeSuiteControls.ComboBox
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private FLis1 As XtremeSuiteControls.ListBox

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    TxDa1.Text = NeuDa
    If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
End If

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

TxDa1.Text = SKaSh(TxDa1.Left, TxDa1.Top + TxDa1.Height, NeuDa, TxDa1.hwnd)

S_LaAu

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo InErr

Dim RetWe As Long
Dim AktZa As Integer

Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set TxDa1 = Me.txtDatu1
Set CmAuf = Me.cmbAuft
Set CmTag = Me.cmbTage
Set CmGes = Me.cmbGesch
Set CmAbr = Me.cmbAbTyp
Set CmGeb = Me.cmbGebüh
Set CmMul = Me.cmbMulti
Set ChBef = Me.chkBefun
Set CmThe = Me.cbmBehan
Set PuBu1 = Me.btnDatu1
Set ImMan = frmMain.imgManag

For AktZa = 1 To UBound(GlMan)
    With CmThe
        .AddItem GlMan(AktZa, 1)
        .ItemData(AktZa - 1) = GlMan(AktZa, 2)
    End With
Next AktZa
CmThe.ListIndex = GlMan(GlSMa, 0) - 1

With CmAuf
    .AddItem "Patientenrechnung"
    .ItemData(0) = 1
    .AddItem "Normalrechnung"
    .ItemData(1) = 2
    .AddItem "Abbucher"
    .ItemData(2) = 3
    .AddItem "Kein Vertrag"
    .ItemData(3) = 4
    .AddItem "Testkunde"
    .ItemData(4) = 5
End With

With CmGes
    .AddItem "M"
    .ItemData(0) = 1
    .AddItem "W"
    .ItemData(1) = 2
    .AddItem "K"
    .ItemData(2) = 3
    .ListIndex = 1
End With

With CmAbr
    .AddItem "Einsender"
    .ItemData(0) = 1
    .AddItem "Kassenpatient"
    .ItemData(1) = 2
    .AddItem "Privatpatient"
    .ItemData(2) = 3
    .AddItem "Sonstige"
    .ItemData(3) = 4
End With

With CmGeb
    .AddItem "BMÄ"
    .ItemData(0) = 1
    .AddItem "EGO"
    .ItemData(1) = 2
    .AddItem "GOÄ 96"
    .ItemData(2) = 3
    .AddItem "BG-Tarif"
    .ItemData(3) = 4
    .AddItem "GOÄ 88"
    .ItemData(4) = 5
End With

With CmTag
    .AddItem "1 Tag"
    .ItemData(0) = 1
    .AddItem "2 Tage"
    .ItemData(1) = 2
    .AddItem "3 Tage"
    .ItemData(2) = 3
    .AddItem "4 Tage"
    .ItemData(3) = 4
    .AddItem "5 Tage"
    .ItemData(4) = 5
    .AddItem "6 Tage"
    .ItemData(5) = 6
    .AddItem "7 Tage"
    .ItemData(6) = 7
    .ListIndex = 0
End With

With CmMul
    .AddItem "1.00"
    .ItemData(0) = 1
    .AddItem "1.15"
    .ItemData(1) = 2
    .AddItem "1.50"
    .ItemData(2) = 3
    .AddItem "2.30"
    .ItemData(3) = 4
    .ListIndex = 0
End With

RetWe = SendMessage(CmAuf.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(CmAbr.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(CmGeb.hwnd, CB_SETCURSEL, 2, ByVal 0&)

TxDa1.Text = Date

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
ChBef.BackColor = GlBak

Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FSuda()
On Error GoTo SuErr

Dim Mld1, Tit1 As String

Set TxKur = Me.txtKurz
Set FLis1 = Me.lstList1

S_LaFin TxKur.Text

If FLis1.ListCount > 0 Then
    FLis1.SetFocus
    FLis1.Selected(0) = True
Else
    Mld1 = "Das von Ihnen eingegebene Suchkriterium brachte leider keine Suchergebnisse"
    Tit1 = "Adressuche"
    WindowMess Mld1, Dial2, Tit1, Me.hwnd
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuda " & Err.Number
Resume Next

End Sub
Private Sub TWeit()
On Error GoTo SuErr

Dim Mld1, Tit1 As String

Set FLis1 = Me.lstList1
Set TxNum = Me.txtPatNu
Set TxAuf = Me.txtAuftr
Set TxNam = Me.txtName
Set TxVor = Me.txtVorna
Set TxGeb = Me.txtGebor
Set CmAuf = Me.cmbAuft
Set TxKos = Me.txtKost
Set CmGes = Me.cmbGesch
Set CmAbr = Me.cmbAbTyp
Set CmGeb = Me.cmbGebüh
Set CmMul = Me.cmbMulti
Set ChBef = Me.chkBefun
Set TxDa1 = Me.txtDatu1
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4

If Rahm1.Visible = True Then
    If FLis1.ListCount > 0 Then
        GlAdr = FLis1.ItemData(FLis1.ListIndex)
        S_LaAu
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
    End If
ElseIf Rahm2.Visible = True Then
    If TxAuf.Text <> vbNullString Then
        If S_LaDo = True Then
            Mld1 = "Die Auftragsnummer " & TxAuf.Text & " existiert bereits"
            Tit1 = "Doppelte Auftragsnummer"
            WindowMess Mld1, Dial3, Tit1, Me.hwnd
        Else
            If ChBef.Value = 1 Then
                TxKos.Text = "18,00"
            Else
                TxKos.Text = GlWa2
            End If
            If S_LaPa = True Then
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = False
                Rahm4.Visible = True
            Else
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = True
                Rahm4.Visible = False
                TxNam.SetFocus
            End If
        End If
    End If
ElseIf Rahm3.Visible = True Then
    If TxAuf.Text = vbNullString Then
        Mld1 = "Bitte tragen Sie zuerst eine gültige Auftragsnummer ein"
        Tit1 = "Keine Auftragsnummer"
        WindowMess Mld1, Dial2, Tit1, Me.hwnd
    ElseIf TxNam.Text = vbNullString Then
        Mld1 = "Bitte tragen Sie zuerst den Nachnamen des Patienten ein"
        Tit1 = "Kein Name"
        WindowMess Mld1, Dial2, Tit1, Me.hwnd
    ElseIf TxVor.Text = vbNullString Then
        Mld1 = "Bitte tragen Sie zuerst den Vornamen des Patienten ein"
        Tit1 = "Kein Name"
        WindowMess Mld1, Dial2, Tit1, Me.hwnd
    ElseIf TxGeb.Text = vbNullString Then
        Mld1 = "Bitte tragen Sie zuerst das Geburtsdatum des Patienten ein"
        Tit1 = "Kein Geburtsdatum"
        WindowMess Mld1, Dial2, Tit1, Me.hwnd
    ElseIf CmGes.Text = vbNullString Then
        Mld1 = "Bitte tragen Sie zuerst das Geschlecht des Patienten ein"
        Tit1 = "Kein Geschlecht"
        WindowMess Mld1, Dial2, Tit1, Me.hwnd
    Else
        S_LaNe
        SUpAu
        SMark
        Unload Me
    End If
ElseIf Rahm4.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TWeit " & Err.Number
Resume Next

End Sub
Private Sub FZur()
On Error Resume Next

Set FLis1 = Me.lstList1
Set TxKur = Me.txtKurz
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4

If Rahm1.Visible = True Then
    TxKur.Text = vbNullString
    FLis1.Clear
    TxKur.SetFocus
ElseIf Rahm2.Visible = True Then
    FLis1.Clear
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = True
ElseIf Rahm3.Visible = True Then
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = True
ElseIf Rahm4.Visible = True Then
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = True
    Rahm1.Visible = False
End If

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

TeTit = ""
TeMai = ""
TeInh = ""
TeFus = ""

'SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    TWeit
End Sub
Private Sub btnZurück_Click()
    FZur
End Sub
Private Sub cmbAbTyp_Click()
    Me.btnWeiter.SetFocus
End Sub
Private Sub cmbAuft_Click()
    Me.btnWeiter.SetFocus
End Sub
Private Sub cmbGebüh_Click()
    Me.btnWeiter.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next

FKonf
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmNeuAuf = Nothing
End Sub
Private Sub lstList1_DblClick()
    TWeit
End Sub

Private Sub lstList1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        TWeit
    End If
End Sub

Private Sub txtAuftr_GotFocus()
    Me.txtAuftr.SelStart = 0
    Me.txtAuftr.SelLength = Len(Me.txtAuftr.Text)
End Sub

Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub
Private Sub txtGebor_Change()
    If Len(Me.txtGebor.Text) = 2 Or Len(Me.txtGebor.Text) = 5 Then
       Me.txtGebor.Text = Me.txtGebor.Text & "."
       Me.txtGebor.SelStart = Len(Me.txtGebor.Text)
    End If
End Sub
Private Sub txtKurz_GotFocus()

Set FLis1 = Me.lstList1

Me.txtKurz.Text = vbNullString

FLis1.Clear

End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub

