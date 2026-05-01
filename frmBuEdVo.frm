VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmBuEdVo 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Buchungsvorlage"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Width           =   7000
      _Version        =   1048579
      _ExtentX        =   12347
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   5000
         TabIndex        =   14
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
      Begin XtremeSuiteControls.PushButton cmdWeite 
         Height          =   400
         Left            =   3600
         TabIndex        =   13
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Speichern"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   2300
         TabIndex        =   12
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
   Begin XtremeSuiteControls.ComboBox cmbKtoRa 
      Height          =   315
      Left            =   2800
      TabIndex        =   8
      Top             =   3230
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtKonto 
      Height          =   350
      Left            =   1200
      TabIndex        =   1
      Top             =   431
      Width           =   4700
      _Version        =   1048579
      _ExtentX        =   8290
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.CheckBox chkGewEr 
      Height          =   220
      Left            =   2800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3800
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Keine Auswertung bei Erlösermittlung"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtBezei 
      Height          =   200
      Left            =   700
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6800
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6800
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbBuTyp 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1830
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.ComboBox cmbBuStu 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   2530
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox4"
   End
   Begin XtremeSuiteControls.ComboBox cmbManda 
      Height          =   315
      Left            =   2800
      TabIndex        =   6
      Top             =   2530
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox6"
   End
   Begin XtremeSuiteControls.ComboBox cmbBuTex 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1130
      Width           =   4700
      _Version        =   1048579
      _ExtentX        =   8281
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtBuBet 
      Height          =   345
      Left            =   1200
      TabIndex        =   7
      Top             =   3230
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoNr 
      Height          =   200
      Left            =   300
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6800
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.ComboBox cmbGegen 
      Height          =   315
      Left            =   2800
      TabIndex        =   4
      Top             =   1830
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox5"
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoHa 
      Height          =   200
      Left            =   1080
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6800
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtBezHa 
      Height          =   200
      Left            =   1485
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6800
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.ComboBox cmbMitar 
      Height          =   315
      Left            =   2500
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6800
      Width           =   500
      _Version        =   1048579
      _ExtentX        =   873
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtIdxNr 
      Height          =   200
      Left            =   1900
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6800
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtHaben 
      Height          =   350
      Left            =   2800
      TabIndex        =   9
      Top             =   3230
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.Label lblLab10 
      Height          =   210
      Left            =   2810
      TabIndex        =   28
      Top             =   2990
      Width           =   1100
      _Version        =   1048579
      _ExtentX        =   1940
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Kontenrahmen :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab06 
      Height          =   210
      Left            =   2810
      TabIndex        =   26
      Top             =   1590
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Geldkonto :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab05 
      Height          =   210
      Left            =   1205
      TabIndex        =   25
      Top             =   200
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Sachkonto :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Label lblLab09 
      BackStyle       =   0  'Transparent
      Caption         =   "Mandant :"
      Height          =   210
      Left            =   2810
      TabIndex        =   21
      Top             =   2300
      Width           =   900
   End
   Begin VB.Label lblLab07 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungstyp :"
      Height          =   210
      Left            =   1205
      TabIndex        =   20
      Top             =   1590
      Width           =   1100
   End
   Begin VB.Label lblLab04 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungstext :"
      Height          =   210
      Left            =   1205
      TabIndex        =   19
      Top             =   870
      Width           =   1200
   End
   Begin VB.Label lblLab02 
      BackStyle       =   0  'Transparent
      Caption         =   "Betrag :"
      Height          =   210
      Left            =   1205
      TabIndex        =   18
      Top             =   2990
      Width           =   900
   End
   Begin VB.Label lblLab08 
      BackStyle       =   0  'Transparent
      Caption         =   "Steuersatz :"
      Height          =   210
      Left            =   1205
      TabIndex        =   17
      Top             =   2300
      Width           =   1100
   End
End
Attribute VB_Name = "frmBuEdVo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Lbl05 As XtremeSuiteControls.Label
Private Lbl10 As XtremeSuiteControls.Label
Private CmRam As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmBuT As XtremeSuiteControls.ComboBox
Private CmBuS As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmBTe As XtremeSuiteControls.ComboBox
Private CmMta As XtremeSuiteControls.ComboBox
Private CmBar As XtremeCommandBars.CommandBar
Private CmAcs As XtremeCommandBars.CommandBarActions
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private TxKto As XtremeSuiteControls.FlatEdit
Private TxHab As XtremeSuiteControls.FlatEdit
Private PuBu1 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private ChAsw As XtremeSuiteControls.CheckBox
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private FoLad As Boolean
Private KntRa As Integer

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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FMand()
On Error GoTo OrErr

Dim ManNr As Long
Dim StaRa As Integer
Dim AktZa As Integer

Set CmMan = Me.cmbManda
Set CmRam = Me.cmbKtoRa

ManNr = CmMan.ItemData(CmMan.ListIndex)

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            If GlMan(AktZa, 25) <> vbNullString Then
                KntRa = GlMan(AktZa, 25) 'Standardkontenrahmen
            Else
                KntRa = GlKtR 'Standardkontenrahmen
            End If
            Exit For
        End If
    Next AktZa
    CmRam.ListIndex = KntRa - 1
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMand " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim AktZa As Integer
Dim AktKo As Integer
Dim GesZa As Integer
Dim LauZa As Integer

Set FM = frmBuEdVo
Set Rahm0 = FM.frmRahm0
Set Lbl05 = FM.lblLab05
Set Lbl10 = FM.lblLab10
Set CmRam = FM.cmbKtoRa
Set CmGeg = FM.cmbGegen
Set CmBuT = FM.cmbBuTex
Set CmBuS = FM.cmbBuStu
Set CmMan = FM.cmbManda
Set CmMta = FM.cmbMitar
Set TxKto = FM.txtKonto
Set TxHab = FM.txtHaben
Set ChAsw = FM.chkGewEr

With FM.cmbBuTyp
    .AddItem "Ausgabe"
    .ItemData(0) = 1
    .AddItem "Einnahme"
    .ItemData(1) = 2
End With

With CmRam
    For AktZa = 1 To UBound(GlKoR) 'Standardkontenrahmen
        .AddItem GlKoR(AktZa, 0)
        .ItemData(AktZa - 1) = GlKoR(AktZa, 1)
    Next AktZa
End With

With CmGeg
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            .AddItem GlGeK(AktZa, 3)
            .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    .AddItem GlSaK(AktKo, 3)
                    .ItemData(AktZa - 1) = GlSaK(AktKo, 6) '[IDB]
                End If
            Next AktKo
        Next AktZa
        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
        End If
    End If
End With

For AktZa = 1 To UBound(GlBTe)
    CmBuT.AddItem GlBTe(AktZa, 1)
    CmBuT.ItemData(CmBuT.NewIndex) = GlBTe(AktZa, 0)
Next AktZa
CmBuT.AutoComplete = True

For AktZa = 1 To UBound(GlStu)
    CmBuS.AddItem GlStu(AktZa, 2)
    CmBuS.ItemData(CmBuS.NewIndex) = GlStu(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlThe)
    If GlNeB = True Then 'neue Buchung
        If CBool(GlThe(AktZa, 25)) = False Then
            LauZa = LauZa + 1
            CmMan.AddItem GlThe(AktZa, 13)
            CmMan.ItemData(CmMan.NewIndex) = GlThe(AktZa, 0)
        End If
    Else
        LauZa = LauZa + 1
        CmMan.AddItem GlThe(AktZa, 13)
        CmMan.ItemData(CmMan.NewIndex) = GlThe(AktZa, 0)
    End If
Next AktZa

For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
    CmMta.AddItem GlMiK(AktZa, 1)
    CmMta.ItemData(CmMta.NewIndex) = GlMiK(AktZa, 2)
Next AktZa

CmMan.ListIndex = GlMan(GlSMa, 0) - 1
CmMta.ListIndex = GlMiK(GlSmI, 0) - 1

If (GlKtR - 1) <= (CmRam.ListCount) - 1 Then
    CmRam.ListIndex = GlKtR - 1 'Standardkontenrahmen
Else
    CmRam.ListIndex = 0
End If

If GlBuc = True Then 'einfache Buchhaltung verwenden
    TxHab.Visible = False
Else
    CmRam.Visible = False
    TxHab.Visible = True
    Lbl05.Caption = "Soll-Konto :"
    Lbl10.Caption = "Haben-Konto :"
End If

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
ChAsw.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FNeu()
On Error GoTo InErr

Dim GeKto As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmGlk As XtremeCommandBars.CommandBarComboBox

Set FM = frmBuEdVo
Set CmRam = FM.cmbKtoRa
Set CmGeg = FM.cmbGegen
Set CmMan = FM.cmbManda
Set CmMta = FM.cmbMitar
Set CmBrs = frmMain.comBar01

Set CmGlk = CmBrs.FindControl(CmGlk, SY_SuBuh, , True)

GeKto = CmGlk.ItemData(CmGlk.ListIndex)

GlNeB = True 'neue Buchung

FM.txtKonto.Text = vbNullString
FM.txtBezei.Text = vbNullString
FM.txtBuBet.Text = GlWa2
FM.cmbBuTex.Text = vbNullString
FM.cmbBuTyp.ListIndex = 0
FM.cmbBuStu.ListIndex = GlStS - 1

If GeKto = 0 Then
    If CmGeg.ListCount > 0 Then
        CmGeg.ListIndex = 0
    End If
Else
    CmGeg.ListIndex = CmGlk.ListIndex - 1
End If

FM.cmdWeite.Enabled = True

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeu & Err.Number"
Resume Next

End Sub
Private Sub FSave()
On Error GoTo LaErr

If Me.txtBuBet.Text = vbNullString Then
    SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
    Exit Sub
End If

If IsNumeric(Me.txtBuBet.Text) = False Then
    SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
    Exit Sub
End If

If CDbl(Me.txtBuBet.Text) < 0 Then
    SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
    Exit Sub
End If

If Me.cmbBuTex.Text = vbNullString Then
    SPopu "Kein Buchungstext", "Es wurde kein Buchungstext eingegeben.", IC48_Forbidden
    Exit Sub
End If

If Me.txtKtoNr.Text = vbNullString Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If IsNumeric(Me.txtKtoNr.Text) = False Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If CLng(Me.txtKtoNr.Text) <= 0 Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If Me.txtKonto.Text = vbNullString Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If Me.txtBezei.Text = vbNullString Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If GlBuc = False Then 'einfache Buchhaltung verwenden
    If Me.txtKtoHa.Text = vbNullString Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
    
    If IsNumeric(Me.txtKtoHa.Text) = False Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
    
    If CLng(Me.txtKtoHa.Text) <= 0 Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If

    If Me.txtHaben.Text = vbNullString Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If

    If Me.txtBezHa.Text = vbNullString Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
End If

K_BuSa "BuVo"
DoEvents

GlNeB = False 'neue Buchung
Unload Me

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50121)
TeMai = IniGetOpt("Hilfe", 50122)
TeInh = IniGetOpt("Hilfe", 50123)
TeFus = IniGetOpt("Hilfe", 50124)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    GlNeB = False 'neue Buchung
    Unload Me
End Sub

Private Sub cmbBuTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuStu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuStu_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.cmbGegen.SetFocus
    Case vbKeyUp: Me.cmbManda.SetFocus
    End Select
End Sub
Private Sub cmbBuTyp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuTyp_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.cmbManda.SetFocus
    Case vbKeyUp: Me.cmbGegen.SetFocus
    End Select
End Sub
Private Sub cmbGegen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbGegen_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.cmbBuTyp.SetFocus
    Case vbKeyUp: Me.cmbBuTex.SetFocus
    End Select
End Sub

Private Sub cmbKtoRa_Click()

Set CmRam = Me.cmbKtoRa
    
If FoLad = False Then
    KntRa = CmRam.ListIndex + 1
End If

End Sub
Private Sub cmbManda_Click()
    If FoLad = False Then
        FMand
    End If
End Sub

Private Sub cmbManda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbManda_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.cmbBuStu.SetFocus
    Case vbKeyUp: Me.cmbBuTyp.SetFocus
    End Select
End Sub
Private Sub cmbMitar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbMitar_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.txtBuKom.SetFocus
    Case vbKeyUp: Me.cmbGegen.SetFocus
    End Select
End Sub

Private Sub cmdWeite_Click()
    FSave
End Sub

Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde
Set CmRam = Me.cmbKtoRa

FoLad = True

FInit

If GlNeB = True Then 'neue Buchung
    FNeu
Else
    K_BuPo "BuVo"
End If

KntRa = GlKtR

FoLad = False

AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBuEdVo = Nothing
End Sub
Private Sub txtBezei_GotFocus()
    Me.txtBezei.SelStart = 0
    Me.txtBezei.SelLength = Len(Me.txtBezei.Text)
End Sub
Private Sub txtBezei_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBezei_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBezei.SelLength = 0
    Case vbKeyDown:
    Case vbKeyUp: Me.txtBezei.SetFocus
    End Select
End Sub
Private Sub txtBuBet_GotFocus()
    Me.txtBuBet.SelStart = 0
    Me.txtBuBet.SelLength = Len(Me.txtBuBet.Text)
End Sub
Private Sub txtBuBet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBuBet_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBuBet.SelLength = 0
    Case vbKeyDown: Me.cmbMitar.SetFocus
    Case vbKeyUp: Me.cmbManda.SetFocus
    End Select
End Sub
Private Sub txtBuBet_LostFocus()

Dim Betra As Double

If Me.txtBuBet.Text <> vbNullString Then
    If IsNumeric(Me.txtBuBet.Text) Then
        Betra = Me.txtBuBet.Text
        Me.txtBuBet.Text = Format$(Betra, GlWa1)
    End If
End If

End Sub

Private Sub txtHaben_GotFocus()
    Me.txtHaben.SelStart = 0
    Me.txtHaben.SelLength = Len(Me.txtHaben.Text)
End Sub

Private Sub txtHaben_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn: KeyAscii = 0
    Case vbKeyTab: KeyAscii = 0
    End Select
End Sub

Private Sub txtHaben_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode
Case vbKeyF2:
        Me.txtKonto.SelLength = 0
Case vbKeyF8:
        FSave
Case vbKeyDown:
        Me.cmbBuTyp.SetFocus
Case vbKeyUp:
        Me.cmbBuTex.SetFocus
Case vbKeyReturn:
        GlBuF = 6 'Buchungsdialog
        S_KtSu "BaAb", KntRa
End Select

End Sub
Private Sub txtKonto_GotFocus()
    Me.txtKonto.SelStart = 0
    Me.txtKonto.SelLength = Len(Me.txtKonto.Text)
End Sub
Private Sub txtKonto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn: KeyAscii = 0
    Case vbKeyTab: KeyAscii = 0
    End Select
End Sub
Private Sub txtKonto_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
Select Case KeyCode
Case vbKeyF2:
        Me.txtKonto.SelLength = 0
Case vbKeyF8:
        FSave
Case vbKeyDown:
        Me.cmbBuTyp.SetFocus
Case vbKeyUp:
        Me.cmbBuTex.SetFocus
Case vbKeyReturn:
        GlBuF = 2 'Buchungsdialog
        S_KtSu "BaAb", KntRa
End Select

End Sub
