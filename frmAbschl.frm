VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmAbschl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Abschluss"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   12
      Top             =   5000
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   4000
         TabIndex        =   15
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Abbrechen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   2600
         TabIndex        =   14
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
         Left            =   1300
         TabIndex        =   13
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
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   5000
      Left            =   200
      TabIndex        =   1
      Top             =   0
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   8819
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkAbKra 
         Height          =   220
         Left            =   1200
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
         _Version        =   1048579
         _ExtentX        =   4471
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Dokumentation festschreiben"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAbRec 
         Height          =   220
         Left            =   1200
         TabIndex        =   3
         Top             =   1400
         Width           =   2775
         _Version        =   1048579
         _ExtentX        =   4895
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungen festschreiben"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAbBuc 
         Height          =   220
         Left            =   1200
         TabIndex        =   2
         Top             =   1000
         Width           =   2775
         _Version        =   1048579
         _ExtentX        =   4895
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Buchungen festschreiben"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit3 
         Height          =   220
         Left            =   1200
         TabIndex        =   9
         Top             =   3500
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Jahr"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit2 
         Height          =   220
         Left            =   1200
         TabIndex        =   7
         Top             =   3000
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Quartal"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit1 
         Height          =   220
         Left            =   1200
         TabIndex        =   5
         Top             =   2500
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Monat"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMonat 
         Height          =   310
         Left            =   2150
         TabIndex        =   6
         Top             =   2480
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbQurta 
         Height          =   310
         Left            =   2150
         TabIndex        =   8
         Top             =   2980
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbJahre 
         Height          =   310
         Left            =   2150
         TabIndex        =   10
         Top             =   3480
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   4300
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   650
         Left            =   150
         TabIndex        =   17
         Top             =   100
         Width           =   5200
         _Version        =   1048579
         _ExtentX        =   9172
         _ExtentY        =   1147
         _StockProps     =   79
         Caption         =   "Festschreibung und Export"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   195
         Left            =   500
         TabIndex        =   16
         Top             =   4330
         Width           =   900
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6500
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
End
Attribute VB_Name = "frmAbschl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private LbLab As XtremeSuiteControls.Label
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private CmMon As XtremeSuiteControls.ComboBox
Private CmQua As XtremeSuiteControls.ComboBox
Private CmJah As XtremeSuiteControls.ComboBox
Private CmThe As XtremeSuiteControls.ComboBox
Private ChBuc As XtremeSuiteControls.CheckBox
Private ChRec As XtremeSuiteControls.CheckBox
Private ChKra As XtremeSuiteControls.CheckBox
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private FrLad As Boolean

Public mExpo As Boolean

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FKonf()
On Error GoTo InErr

Dim RetWe As Long
Dim AkMon As Integer
Dim AkQua As Integer
Dim IdxZa As Integer
Dim AkJah As Integer
Dim AktZa As Integer

Set FM = frmAbschl
Set LbLab = FM.lblLab01
Set Rahm0 = FM.frmRahm0
Set Rahm2 = FM.frmRahm2
Set OpMon = FM.optZeit1
Set OpQua = FM.optZeit2
Set OpJah = FM.optZeit3
Set ChBuc = FM.chkAbBuc
Set ChRec = FM.chkAbRec
Set ChKra = FM.chkAbKra
Set CmMon = FM.cmbMonat
Set CmQua = FM.cmbQurta
Set CmJah = FM.cmbJahre
Set CmThe = FM.cmbBehan

AkMon = Month(Date)

If AkMon <= 3 Then
    AkQua = 1
ElseIf AkMon <= 6 Then
    AkQua = 2
ElseIf AkMon <= 9 Then
    AkQua = 3
ElseIf AkMon <= 12 Then
    AkQua = 4
End If

With CmMon
    .DropDownItemCount = 12
    For IdxZa = 1 To 12
        .AddItem MonthName(IdxZa)
        .ItemData(.NewIndex) = IdxZa
    Next IdxZa
End With

With CmQua
    .AddItem "1. Quartal"
    .ItemData(.NewIndex) = 1
    .AddItem "2. Quartal"
    .ItemData(.NewIndex) = 2
    .AddItem "3. Quartal"
    .ItemData(.NewIndex) = 3
    .AddItem "4. Quartal"
    .ItemData(.NewIndex) = 4
End With

With CmJah
    .DropDownItemCount = 12
    For AkJah = Year(Date) - 15 To Year(Date) + 1
        .AddItem AkJah
        .ItemData(.NewIndex) = IdxZa
        IdxZa = IdxZa + 1
    Next AkJah
    If mExpo = True Then
        .Text = Year(Date)
    Else
        .Text = Year(Date) - 1
    End If
End With

With CmThe
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa - 1) = GlThe(AktZa, 0)
    Next AktZa
    .AddItem "f¸r alle Mandanten"
    .ItemData(AktZa - 1) = 0
    .ListIndex = AktZa - 1
End With

If CmThe.Enabled = False Then
    CmThe.Enabled = True
End If

FM.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm0.BackColor = GlBak
OpJah.BackColor = GlBak
OpMon.BackColor = GlBak
OpQua.BackColor = GlBak
OpJah.BackColor = GlBak
ChBuc.BackColor = GlBak
ChRec.BackColor = GlBak
ChKra.BackColor = GlBak

RetWe = SendMessage(CmMon.hwnd, CB_SETCURSEL, AkMon - 1, ByVal 0&)
RetWe = SendMessage(CmQua.hwnd, CB_SETCURSEL, AkQua - 1, ByVal 0&)

If mExpo = True Then
    FM.Caption = "GoBD Export"
    LbLab.Caption = "Bitte w‰hlen Sie aus, f¸r welchen Zeitraum und Mandanten die Daten exportiert werden sollen. Sollen die Daten f¸r mehrere Jahre exportiert werden, muss dieser Vorgang entsprechend oft wiederholt werden."
    ChBuc.Enabled = False
    ChRec.Enabled = False
    ChKra.Enabled = False
Else
    FM.Caption = "GoBD Festschreibung"
    LbLab.Caption = "Bitte w‰hlen Sie aus, f¸r welchen Zeitraum die Daten festgeschrieben werden sollen. Sollen die Daten f¸r mehrere Jahre festgeschrieben werden, muss dieser Vorgang entsprechend oft wiederholt werden."
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Function FKrit() As String
On Error GoTo InErr

Dim ManNr As Long
Dim VonDa As Date
Dim BisDa As Date
Dim Krit1 As String
Dim Krit2 As String
Dim Krite As String
Dim AkMon As Integer
Dim AkJha As Integer
Dim AkQua As Integer

Set FM = frmAbschl
Set OpMon = FM.optZeit1
Set OpQua = FM.optZeit2
Set OpJah = FM.optZeit3
Set CmMon = FM.cmbMonat
Set CmQua = FM.cmbQurta
Set CmJah = FM.cmbJahre
Set CmThe = FM.cmbBehan

AkJha = CInt(CmJah.Text)
AkMon = CmMon.ItemData(CmMon.ListIndex)
AkQua = CmQua.ItemData(CmQua.ListIndex)

ManNr = CmThe.ItemData(CmThe.ListIndex)

If GlMaV = True Then 'Mandanten vorhanden
    If ManNr > 0 Then
        If GlTyp < 2 Then
            Krit2 = " AND (ManNr = " & ManNr & ")"
        Else
            Krit2 = " AND ([ManNr] = " & ManNr & ")"
        End If
    End If
End If

If OpMon.Value = True Then 'Monatsauswertung

    If GlTyp < 2 Then
        Krit1 = "(((MONTH(Datum))=" & AkMon & ") AND ((YEAR(Datum))=" & AkJha & "))"
    Else
        Krit1 = "(((Month([Datum]))=" & AkMon & ") AND ((Year([Datum]))=" & AkJha & "))"
    End If
    If AkMon = 1 Then
        VonDa = CDate("01." & AkMon - 1 & "." & AkJha)
        BisDa = CDate("01." & AkMon & "." & AkJha)
    Else
        VonDa = CDate("01." & AkMon - 1 & "." & AkJha)
        BisDa = CDate("01." & AkMon & "." & AkJha) - 1
    End If
    
ElseIf OpQua.Value = True Then 'Quartalsauswertung
    
    If GlTyp < 2 Then
        Select Case AkQua
        Case 1: Krit1 = "((Datum >= '01.01." & AkJha & "') AND (Datum <= '31.03." & AkJha & "'))"
        Case 2: Krit1 = "((Datum >= '01.04." & AkJha & "') AND (Datum <= '30.06." & AkJha & "'))"
        Case 3: Krit1 = "((Datum >= '01.07." & AkJha & "') AND (Datum <= '30.09." & AkJha & "'))"
        Case 4: Krit1 = "((Datum >= '01.10." & AkJha & "') AND (Datum <= '31.12." & AkJha & "'))"
        End Select
    Else
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] Between #01/01/" & AkJha & "# AND #03/31/" & AkJha & "#))"
        Case 2: Krit1 = "(([Datum] Between #04/01/" & AkJha & "# AND #06/30/" & AkJha & "#))"
        Case 3: Krit1 = "(([Datum] Between #07/01/" & AkJha & "# AND #09/30/" & AkJha & "#))"
        Case 4: Krit1 = "(([Datum] Between #10/01/" & AkJha & "# AND #12/31/" & AkJha & "#))"
        End Select
    End If
    
    Select Case AkQua
    Case 1: BisDa = CDate("01.01." & AkJha)
    Case 2: BisDa = CDate("01.04." & AkJha) - 1
    Case 3: BisDa = CDate("01.07." & AkJha) - 1
    Case 4: BisDa = CDate("01.10." & AkJha) - 1
    End Select
    
ElseIf OpJah.Value = True Then 'Jahresauswertung
    
    If GlTyp < 2 Then
        Krit1 = "((YEAR(Datum) = " & AkJha & "))"
    Else
        Krit1 = "((Year([Datum]) = " & AkJha & "))"
    End If
    
    BisDa = CDate("01.01." & AkJha)

End If

If Krit1 <> vbNullString Then
    If Krit2 <> vbNullString Then
        Krite = Krit1 & Krit2
    Else
        Krite = Krit1
    End If
Else
    Krite = vbNullString
End If

FKrit = Krite
    
Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKrit " & Err.Number
Resume Next
    
End Function

Private Sub FWeit()
On Error GoTo InErr

Dim Krite As String

Krite = FKrit()

If mExpo = True Then
    Unload Me
    S_AlExT True, True, Krite
Else
    S_AbSch
    Unload Me
End If
    
Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next
    
End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50011)
TeMai = IniGetOpt("Hilfe", 50011)
TeInh = IniGetOpt("Hilfe", 50012)
TeFus = IniGetOpt("Hilfe", 50013)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub
Private Sub cmbJahre_Click()
    Me.optZeit3.Value = True
End Sub
Private Sub cmbJahre_DropDown()
    Me.optZeit3.Value = True
End Sub
Private Sub cmbMonat_Click()
    Me.optZeit1.Value = True
End Sub
Private Sub cmbMonat_DropDown()
    Me.optZeit1.Value = True
End Sub
Private Sub cmbQurta_Click()
    Me.optZeit2.Value = True
End Sub
Private Sub cmbQurta_DropDown()
    Me.optZeit2.Value = True
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FrmEx.TopMost = True

FrLad = True
FKonf
AFont Me
FrLad = False
SFrame 1, Me.hwnd

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAbschl = Nothing
End Sub
