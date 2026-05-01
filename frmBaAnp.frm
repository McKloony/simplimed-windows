VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmBaAnp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kontoausz¸ge Anpassen"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   14
      Top             =   3600
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   4200
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schlieﬂen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Height          =   400
         Left            =   2800
         TabIndex        =   16
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
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   1500
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
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3500
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6100
      _Version        =   1048579
      _ExtentX        =   10760
      _ExtentY        =   6174
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkZaTyp 
         Height          =   225
         Left            =   3800
         TabIndex        =   12
         Top             =   2500
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Zahlungstyp"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKennz 
         Height          =   225
         Left            =   3800
         TabIndex        =   10
         Top             =   1700
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Kennzeichnen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkManda 
         Height          =   225
         Left            =   400
         TabIndex        =   4
         Top             =   1700
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandant"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkGegen 
         Height          =   220
         Left            =   400
         TabIndex        =   2
         Top             =   900
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Geldkonto"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbGegen 
         Height          =   315
         Left            =   400
         TabIndex        =   3
         Top             =   1200
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   400
         TabIndex        =   5
         Top             =   2000
         Width           =   2895
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.CheckBox chkZuord 
         Height          =   225
         Left            =   3800
         TabIndex        =   8
         Top             =   900
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Gebucht"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbZuord 
         Height          =   315
         Left            =   3800
         TabIndex        =   9
         Top             =   1200
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   400
         TabIndex        =   7
         Top             =   2800
         Width           =   2895
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.CheckBox chkMitar 
         Height          =   225
         Left            =   400
         TabIndex        =   6
         Top             =   2500
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mitarbeiter"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbKennz 
         Height          =   315
         Left            =   3800
         TabIndex        =   11
         Top             =   2000
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbZaTyp 
         Height          =   315
         Left            =   3800
         TabIndex        =   13
         Top             =   2800
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBaAnp.frx":0000
         Height          =   585
         Left            =   400
         TabIndex        =   18
         Top             =   100
         Width           =   5500
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   5860
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
End
Attribute VB_Name = "frmBaAnp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmZuo As XtremeSuiteControls.ComboBox
Private CmKen As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private ChGeg As XtremeSuiteControls.CheckBox
Private ChMan As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private ChZuo As XtremeSuiteControls.CheckBox
Private ChKen As XtremeSuiteControls.CheckBox
Private ChTyp As XtremeSuiteControls.CheckBox
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private Sub TAbs()
On Error GoTo OpErr
'ƒndert die Rechnungen

Dim RowNr As Long
Dim AnzPo As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

AnzPo = RpSel.Count

If AnzPo > 0 Then
    Screen.MousePointer = vbHourglass

    S_BaAnp
    DoEvents
    If AnzPo > 1 Then
        SUpBa 0, True
    Else
        Set RpSel = RpCo1.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpBa RowNr
        End If
    End If
    
    Screen.MousePointer = vbNormal
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TAbs " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim AktZa As Integer
Dim AktKo As Integer
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmBaAnp
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set CmGeg = FM.cmbGegen
Set CmMan = FM.cmbManda
Set CmMit = FM.cmbMitar
Set CmZuo = FM.cmbZuord
Set CmKen = FM.cmbKennz
Set CmTyp = FM.cmbZaTyp
Set ChGeg = FM.chkGegen
Set ChMan = FM.chkManda
Set ChMit = FM.chkMitar
Set ChZuo = FM.chkZuord
Set ChKen = FM.chkKennz
Set ChTyp = FM.chkZaTyp

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
        If .ListCount = 0 Then 'f¸ge die Geldkonten aus der einfachen Buchf¸hrung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
        End If
    End If
End With

For AktZa = 1 To UBound(GlMaA)
    CmMan.AddItem GlMaA(AktZa, 1)
    CmMan.ItemData(CmMan.NewIndex) = GlMaA(AktZa, 2)
Next AktZa

For AktZa = 1 To UBound(GlMiA) 'Alle Mitarbeiter
    CmMit.AddItem GlMiA(AktZa, 1)
    CmMit.ItemData(CmMit.NewIndex) = GlMiA(AktZa, 2)
Next AktZa

With CmZuo
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmKen
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmTyp
    .AddItem "Ausgabe"
    .ItemData(0) = 1
    .AddItem "Einnahme"
    .ItemData(1) = 2
End With

If CmGeg.ListCount > 0 Then
    CmGeg.ListIndex = 0
End If
CmMan.ListIndex = GlSMa - 1
CmMit.ListIndex = GlSmI - 1
CmZuo.ListIndex = 1
CmKen.ListIndex = 1
CmTyp.ListIndex = 0

ChMit.Enabled = GlMiV

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
ChGeg.BackColor = GlBak
ChMan.BackColor = GlBak
ChMit.BackColor = GlBak
ChZuo.BackColor = GlBak
ChKen.BackColor = GlBak
ChTyp.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TInit " & Err.Number
Resume Next

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

Private Sub chkGegen_Click()
On Error Resume Next

Set ChGeg = Me.chkGegen
Set CmGeg = Me.cmbGegen

If ChGeg.Value = xtpChecked Then
    CmGeg.Enabled = True
Else
    CmGeg.Enabled = False
End If

End Sub

Private Sub chkKennz_Click()
On Error Resume Next

Set ChKen = Me.chkKennz
Set CmKen = Me.cmbKennz

If ChKen.Value = xtpChecked Then
    CmKen.Enabled = True
Else
    CmKen.Enabled = False
End If

End Sub

Private Sub chkManda_Click()
On Error Resume Next

Set ChMan = Me.chkManda
Set CmMan = Me.cmbManda

If ChMan.Value = xtpChecked Then
    CmMan.Enabled = True
Else
    CmMan.Enabled = False
End If

End Sub
Private Sub chkMitar_Click()
On Error Resume Next

Set ChMit = Me.chkMitar
Set CmMit = Me.cmbMitar

If ChMit.Value = xtpChecked Then
    CmMit.Enabled = True
Else
    CmMit.Enabled = False
End If

End Sub

Private Sub chkZaTyp_Click()
On Error Resume Next

Set ChTyp = Me.chkZaTyp
Set CmTyp = Me.cmbZaTyp

If ChTyp.Value = xtpChecked Then
    CmTyp.Enabled = True
Else
    CmTyp.Enabled = False
End If

End Sub
Private Sub chkZuord_Click()
On Error Resume Next

Set ChZuo = Me.chkZuord
Set ChKen = FM.chkKennz
Set CmZuo = Me.cmbZuord
Set CmKen = FM.cmbKennz

If ChZuo.Value = xtpChecked Then
    CmZuo.Enabled = True
    ChKen.Enabled = False
    ChKen.Value = xtpUnchecked
    CmKen.Enabled = False
    CmKen.ListIndex = 1
Else
    CmZuo.Enabled = False
    ChKen.Enabled = True
End If

End Sub
Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBaAnp = Nothing
End Sub
Private Sub btnWeiter_Click()
    TAbs
    Unload Me
End Sub
