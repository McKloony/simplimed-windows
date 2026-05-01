VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReAnd 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnungen Anpassen"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   33
      Top             =   6500
      Width           =   6500
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4500
         TabIndex        =   36
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
         Left            =   3100
         TabIndex        =   35
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
         Left            =   1800
         TabIndex        =   34
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
      Height          =   6400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6500
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   11289
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkVersa 
         Height          =   225
         Left            =   600
         TabIndex        =   14
         Top             =   5600
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Versandart"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFormu 
         Height          =   225
         Left            =   600
         TabIndex        =   12
         Top             =   4800
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Formular"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkReNum 
         Height          =   255
         Left            =   3900
         TabIndex        =   31
         Top             =   5600
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rechnungsnummer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   360
         Left            =   5480
         TabIndex        =   22
         Top             =   1890
         Width           =   250
         _Version        =   1048579
         _ExtentX        =   441
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         Min             =   1
         Value           =   1
         Max             =   10
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtKopie"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.CheckBox chkManda 
         Height          =   225
         Left            =   600
         TabIndex        =   6
         Top             =   2400
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandant"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKopie 
         Height          =   225
         Left            =   3900
         TabIndex        =   20
         Top             =   1600
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Ausdrucke"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkZahlu 
         Height          =   220
         Left            =   600
         TabIndex        =   4
         Top             =   1600
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zahlungsweise"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKatal 
         Height          =   220
         Left            =   600
         TabIndex        =   2
         Top             =   800
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Gebührensatz"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbVersi 
         Height          =   310
         Left            =   600
         TabIndex        =   3
         Top             =   1100
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         MaxLength       =   10
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbZaZie 
         Height          =   310
         Left            =   600
         TabIndex        =   5
         Top             =   1900
         Width           =   2900
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
      Begin XtremeSuiteControls.FlatEdit txtKopie 
         Height          =   350
         Left            =   3900
         TabIndex        =   21
         Top             =   1900
         Width           =   1560
         _Version        =   1048579
         _ExtentX        =   2752
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "1"
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   600
         TabIndex        =   7
         Top             =   2700
         Width           =   2900
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
      Begin XtremeSuiteControls.CheckBox chkDatum 
         Height          =   225
         Left            =   3900
         TabIndex        =   16
         Top             =   800
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnunsdatum"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   5480
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1100
         Width           =   250
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu1"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   5760
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   1100
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   3900
         TabIndex        =   17
         Top             =   1100
         Width           =   1560
         _Version        =   1048579
         _ExtentX        =   2752
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbReTyp 
         Height          =   310
         Left            =   3900
         TabIndex        =   24
         Top             =   2700
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkReTyp 
         Height          =   225
         Left            =   3900
         TabIndex        =   23
         Top             =   2400
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Belegtyp"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtExtra 
         Height          =   350
         Left            =   3900
         TabIndex        =   30
         Top             =   5100
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkSteue 
         Height          =   225
         Left            =   3900
         TabIndex        =   27
         Top             =   4000
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Steuer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkStorn 
         Height          =   225
         Left            =   3900
         TabIndex        =   25
         Top             =   3200
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Storniert"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbStorn 
         Height          =   315
         Left            =   3900
         TabIndex        =   26
         Top             =   3500
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkPatie 
         Height          =   225
         Left            =   600
         TabIndex        =   8
         Top             =   3200
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Patient"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbPatie 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   3500
         Width           =   2900
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
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   4300
         Width           =   2900
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
         Left            =   600
         TabIndex        =   10
         Top             =   4000
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mitarbeiter"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkExtra 
         Height          =   225
         Left            =   3900
         TabIndex        =   29
         Top             =   4800
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Extragebühr"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbReStu 
         Height          =   315
         Left            =   3900
         TabIndex        =   28
         Top             =   4300
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtReNum 
         Height          =   350
         Left            =   3900
         TabIndex        =   32
         Top             =   5900
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "1"
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbFormu 
         Height          =   315
         Left            =   600
         TabIndex        =   13
         Top             =   5100
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         MaxLength       =   10
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbVersa 
         Height          =   315
         Left            =   600
         TabIndex        =   15
         Top             =   5900
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmReAnd.frx":0000
         Height          =   585
         Left            =   400
         TabIndex        =   38
         Top             =   100
         Width           =   5500
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   7600
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   1000
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   7600
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   0
   End
End
Attribute VB_Name = "frmReAnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmVer As XtremeSuiteControls.ComboBox
Private CmZil As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmPat As XtremeSuiteControls.ComboBox
Private CmSto As XtremeSuiteControls.ComboBox
Private CmReS As XtremeSuiteControls.ComboBox
Private CmFor As XtremeSuiteControls.ComboBox
Private CmVrs As XtremeSuiteControls.ComboBox
Private TxKop As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxExt As XtremeSuiteControls.FlatEdit
Private TxRen As XtremeSuiteControls.FlatEdit
Private ChGeb As XtremeSuiteControls.CheckBox
Private ChVer As XtremeSuiteControls.CheckBox
Private ChZil As XtremeSuiteControls.CheckBox
Private ChKop As XtremeSuiteControls.CheckBox
Private ChMan As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private ChPat As XtremeSuiteControls.CheckBox
Private ChDat As XtremeSuiteControls.CheckBox
Private ChTyp As XtremeSuiteControls.CheckBox
Private ChSto As XtremeSuiteControls.CheckBox
Private ChSte As XtremeSuiteControls.CheckBox
Private ChRen As XtremeSuiteControls.CheckBox
Private ChFor As XtremeSuiteControls.CheckBox
Private ChVrs As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private UpCo1 As XtremeSuiteControls.UpDown
Private UpCo2 As XtremeSuiteControls.UpDown
Private MoKal As XtremeCalendarControl.DatePicker
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private Sub TWeit()
On Error GoTo OpErr
'Ändert die Rechnungen

Dim RowNr As Long
Dim KrRow As Long
Dim AnzPo As Integer
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo4 = FM.repCont4
Set RpCo6 = FM.repCont6
Set RpCo3 = FM.repCont3

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select
AnzPo = RpSel.Count

Screen.MousePointer = vbHourglass
DoEvents

If AnzPo > 0 Then
    S_ReAnp
    DoEvents
    Select Case GlBut
    Case RibTab_Abrechnung:
            If AnzPo > 1 Then
                SUpAb
                SUpRe , True
            Else
                Set RpRow = RpSel(0)
                RowNr = RpRow.Index
                Set RpSel = RpCo6.SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    KrRow = RpRow.Index
                    SUpAb RowNr, KrRow
                Else
                    SUpAb RowNr
                End If
                Set RpSel = RpCo4.SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    RowNr = RpRow.Index
                    SUpRe RowNr
                End If
            End If
    Case RibTab_Rechnungen:
            If AnzPo > 1 Then
                SUpRe , True
            Else
                Set RpSel = RpCo4.SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    RowNr = RpRow.Index
                    SUpRe RowNr
                End If
            End If
    End Select
End If

DoEvents
Screen.MousePointer = vbNormal

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo6 = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "Tweit " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim TmDat As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1
Set Rahm1 = Me.frmRahm1

If IsDate(TxDa1.Text) Then
    NeuDa = CDate(TxDa1.Text)
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = Rahm1.Top + TxDa1.Top + TxDa1.Height
    .Left = Rahm1.Left + TxDa1.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TmDat = .Selection.Blocks(0).DateBegin
            If Year(TmDat) <= Year(Date) Then
                TxDa1.Text = Format$(TmDat, "dd.mm.yyyy")
            End If
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = CDate(MoKal.Selection.Blocks(0).DateBegin)
    If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
        If Year(NeuDa) < Year(Date) Then
            NeuDa = Date
            SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
        End If
    End If
    TxDa1.Text = NeuDa
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim AkJah As String
Dim AkMon As String
Dim AktZa As Integer
Dim FoTyp As Integer
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmReAnd
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set MoKal = FM.dtpDatu1
Set CmTyp = FM.cmbReTyp
Set CmVer = FM.cmbVersi
Set CmZil = FM.cmbZaZie
Set CmSto = FM.cmbStorn
Set CmMan = FM.cmbManda
Set CmMit = FM.cmbMitar
Set CmPat = FM.cmbPatie
Set CmReS = FM.cmbReStu
Set CmFor = FM.cmbFormu
Set CmVrs = FM.cmbVersa
Set ChGeb = FM.chkExtra
Set ChVrs = FM.chkVersa
Set TxDa1 = FM.txtDatu1
Set TxKop = FM.txtKopie
Set TxExt = FM.txtExtra
Set TxRen = FM.txtReNum
Set ChMan = FM.chkManda
Set ChMit = FM.chkMitar
Set ChPat = FM.chkPatie
Set ChVer = FM.chkKatal
Set ChZil = FM.chkZahlu
Set ChKop = FM.chkKopie
Set ChTyp = FM.chkReTyp
Set ChDat = FM.chkDatum
Set ChSto = FM.chkStorn
Set ChSte = FM.chkSteue
Set ChRen = FM.chkReNum
Set ChFor = FM.chkFormu
Set PuBu1 = FM.btnDatu1
Set ImMan = frmMain.imgManag

AkJah = Right$(Year(Date), 2)
AkMon = Format$(Month(Date), "00")

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
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Behandlungstag"
    .MonthDelta = 1
    .YearsTriangle = False
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
End With

For AktZa = 1 To UBound(GlGKa)
    CmVer.AddItem GlGKa(AktZa, 1)
    CmVer.ItemData(AktZa - 1) = GlGKa(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlZah)
    CmZil.AddItem GlZah(AktZa, 1)
    CmZil.ItemData(AktZa - 1) = GlZah(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
    CmMit.AddItem GlMiK(AktZa, 1)
    CmMit.ItemData(AktZa - 1) = GlMiK(AktZa, 2)
Next AktZa

With CmTyp
    .AddItem "R - Standardrechnung"
    .ItemData(0) = 1
    .AddItem "L - Laborrechnung"
    .ItemData(1) = 2
    .AddItem "A - Abrechnungsstelle"
    .ItemData(2) = 3
    .AddItem "G - Gewerberechnung"
    .ItemData(3) = 4
    .AddItem "I - Importrechnung"
    .ItemData(4) = 5
    .ListIndex = 0
End With

With CmVrs
    .AddItem "Postversand"
    .ItemData(0) = 0
    .AddItem "Emailversand"
    .ItemData(1) = 1
    .AddItem "Downloadlink"
    .ItemData(2) = 2
    .ListIndex = GlRVs - 1 'Standard-Rechnungsversandweg
End With

For AktZa = 1 To UBound(GlStu)
    CmReS.AddItem GlStu(AktZa, 2)
    CmReS.ItemData(AktZa - 1) = GlStu(AktZa, 0)
Next AktZa

For AktZa = 0 To 105
    FoTyp = Left$(GlFrm(3, AktZa), 2)
    If FoTyp < 4 Then
        CmFor.AddItem GlFrm(0, AktZa)
        CmFor.ItemData(AktZa) = AktZa + 1
    End If
Next AktZa

With CmSto
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

CmVer.ListIndex = 0
CmZil.ListIndex = 0
CmSto.ListIndex = 1
CmMan.ListIndex = GlSMa - 1
CmMit.ListIndex = GlSmI - 1
CmReS.ListIndex = 0
CmFor.ListIndex = 0

With TxKop
    .Pattern = "\d*"
    .SetMask "0", "_"
    .Text = 1
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxRen
    .Pattern = "\d*"
    .SetMask "000000", "______"
    .Text = "000001"
End With

TxExt.Text = GlWa2

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

If GlRMa = True Then 'getrennter Mandentenrechnungsnummernkreis
    ChMan.Enabled = False
End If

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
ChMan.BackColor = GlBak
ChMit.BackColor = GlBak
ChPat.BackColor = GlBak
ChVer.BackColor = GlBak
ChZil.BackColor = GlBak
ChKop.BackColor = GlBak
ChTyp.BackColor = GlBak
ChGeb.BackColor = GlBak
ChDat.BackColor = GlBak
ChSto.BackColor = GlBak
ChSte.BackColor = GlBak
ChRen.BackColor = GlBak
ChFor.BackColor = GlBak
ChVrs.BackColor = GlBak

Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TInit " & Err.Number
Resume Next

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
        If Year(NeuDa) < Year(Date) Then
            NeuDa = Date
            SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
        End If
    End If
    TxDa1.Text = NeuDa
    With MoKal
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
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

TeTit = IniGetOpt("Hilfe", 50811)
TeMai = IniGetOpt("Hilfe", 50812)
TeInh = IniGetOpt("Hilfe", 50813)
TeFus = IniGetOpt("Hilfe", 50814)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    TWeit
    Unload Me
End Sub

Private Sub chkDatum_Click()
On Error Resume Next

Set ChDat = Me.chkDatum
Set TxDa1 = Me.txtDatu1
Set UpCo1 = Me.updCont1
Set PuBu1 = Me.btnDatu1

If ChDat.Value = xtpChecked Then
    TxDa1.Enabled = True
    UpCo1.Enabled = True
    PuBu1.Enabled = True
Else
    TxDa1.Enabled = False
    UpCo1.Enabled = False
    PuBu1.Enabled = False
End If

End Sub

Private Sub chkExtra_Click()
On Error Resume Next

Set ChGeb = Me.chkExtra
Set TxExt = Me.txtExtra

If ChGeb.Value = xtpChecked Then
    TxExt.Enabled = True
Else
    TxExt.Enabled = False
End If

End Sub

Private Sub chkFormu_Click()
On Error Resume Next

Set ChFor = Me.chkFormu
Set CmFor = Me.cmbFormu

If ChFor.Value = xtpChecked Then
    CmFor.Enabled = True
Else
    CmFor.Enabled = False
End If

End Sub
Private Sub chkKatal_Click()
On Error Resume Next

Set ChVer = Me.chkKatal
Set CmVer = Me.cmbVersi

If ChVer.Value = xtpChecked Then
    CmVer.Enabled = True
Else
    CmVer.Enabled = False
End If

End Sub

Private Sub chkKopie_Click()
On Error Resume Next

Set ChKop = Me.chkKopie
Set TxKop = Me.txtKopie
Set UpCo2 = Me.updCont2

If ChKop.Value = xtpChecked Then
    TxKop.Enabled = True
    UpCo2.Enabled = True
Else
    TxKop.Enabled = False
    UpCo2.Enabled = False
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
Private Sub chkPatie_Click()
On Error Resume Next

Set ChPat = Me.chkPatie
Set CmPat = Me.cmbPatie

If ChPat.Value = xtpChecked Then
    CmPat.Enabled = True
Else
    CmPat.Enabled = False
End If

End Sub
Private Sub chkReNum_Click()
On Error Resume Next

Set FM = frmReAnd
Set ChRen = FM.chkReNum
Set TxRen = FM.txtReNum

If ChRen.Value = xtpChecked Then
    TxRen.Enabled = True
Else
    TxRen.Enabled = False
End If

End Sub

Private Sub chkReTyp_Click()
On Error Resume Next

Set ChTyp = Me.chkReTyp
Set CmTyp = Me.cmbReTyp

If ChTyp.Value = xtpChecked Then
    CmTyp.Enabled = True
Else
    CmTyp.Enabled = False
End If

End Sub

Private Sub chkSteue_Click()
On Error Resume Next

Set ChSte = Me.chkSteue
Set CmReS = Me.cmbReStu

If ChSte.Value = xtpChecked Then
    CmReS.Enabled = True
Else
    CmReS.Enabled = False
End If

End Sub
Private Sub chkStorn_Click()
On Error Resume Next

Set ChSto = Me.chkStorn
Set CmSto = Me.cmbStorn

If ChSto.Value = xtpChecked Then
    CmSto.Enabled = True
Else
    CmSto.Enabled = False
End If

End Sub

Private Sub chkVersa_Click()
On Error Resume Next

Set ChVrs = Me.chkVersa
Set CmVrs = Me.cmbVersa

If ChVrs.Value = xtpChecked Then
    CmVrs.Enabled = True
Else
    CmVrs.Enabled = False
End If

End Sub
Private Sub chkZahlu_Click()
On Error Resume Next

Set ChZil = Me.chkZahlu
Set CmZil = Me.cmbZaZie

If ChZil.Value = xtpChecked Then
    CmZil.Enabled = True
Else
    CmZil.Enabled = False
End If

End Sub
Private Sub cmbManda_Click()
    S_ReAnL
End Sub

Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmReAnd = Nothing
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub

Private Sub updCont1_DownClick()
On Error Resume Next

Dim AltDa As Date
Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text
NeuDa = DateAdd("d", -1, AltDa)

If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
    If Year(NeuDa) < Year(Date) Then
        NeuDa = Date
        SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
    End If
End If

TxDa1.Text = NeuDa

End Sub
Private Sub updCont1_UpClick()
On Error Resume Next

Dim AltDa As Date
Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text
NeuDa = DateAdd("d", 1, AltDa)

If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
    If Year(NeuDa) > Year(Date) Then
        NeuDa = Date
        SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
    End If
End If

TxDa1.Text = NeuDa

End Sub
