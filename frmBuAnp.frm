VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmBuAnp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Buchungen Anpassen"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   13
      Top             =   5800
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
         TabIndex        =   16
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
         Left            =   3600
         TabIndex        =   15
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
         Left            =   2300
         TabIndex        =   14
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
      Height          =   5800
      Left            =   100
      TabIndex        =   1
      Top             =   0
      Width           =   6700
      _Version        =   1048579
      _ExtentX        =   11818
      _ExtentY        =   10231
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkSaKon 
         Height          =   220
         Left            =   400
         TabIndex        =   5
         Top             =   2400
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Sachkonto"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkManda 
         Height          =   225
         Left            =   400
         TabIndex        =   7
         Top             =   3200
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
         TabIndex        =   3
         Top             =   1600
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
         TabIndex        =   4
         Top             =   1900
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
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   400
         TabIndex        =   8
         Top             =   3500
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
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
         Left            =   4400
         TabIndex        =   17
         Top             =   800
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Buchungsdatum"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   350
         Left            =   5730
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1100
         Width           =   255
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
         Left            =   6010
         TabIndex        =   20
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
         Left            =   4400
         TabIndex        =   18
         Top             =   1100
         Width           =   1310
         _Version        =   1048579
         _ExtentX        =   2311
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBuTyp 
         Height          =   315
         Left            =   4400
         TabIndex        =   22
         Top             =   1900
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
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
         Left            =   4400
         TabIndex        =   21
         Top             =   1600
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Buchungstyp"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkStorn 
         Height          =   225
         Left            =   4400
         TabIndex        =   23
         Top             =   2400
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Storniert"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbStorn 
         Height          =   315
         Left            =   4400
         TabIndex        =   24
         Top             =   2700
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbSaKon 
         Height          =   315
         Left            =   400
         TabIndex        =   6
         Top             =   2700
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
      Begin XtremeSuiteControls.CheckBox chkAuswe 
         Height          =   225
         Left            =   4400
         TabIndex        =   25
         Top             =   3200
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Keine Auswertung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbAuswe 
         Height          =   315
         Left            =   4400
         TabIndex        =   26
         Top             =   3500
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
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
         TabIndex        =   10
         Top             =   4300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
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
         TabIndex        =   9
         Top             =   4000
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mitarbeiter"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbReStu 
         Height          =   315
         Left            =   4400
         TabIndex        =   28
         Top             =   4300
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkSteue 
         Height          =   225
         Left            =   4400
         TabIndex        =   27
         Top             =   4000
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Steuer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBeleg 
         Height          =   225
         Left            =   4400
         TabIndex        =   11
         Top             =   4820
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Belegnummernstart"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBeleg 
         Height          =   350
         Left            =   4400
         TabIndex        =   12
         Top             =   5120
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbKtoRa 
         Height          =   315
         Left            =   400
         TabIndex        =   2
         Top             =   1100
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   220
         Left            =   400
         TabIndex        =   31
         Top             =   800
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Kontenrahmen :"
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBuAnp.frx":0000
         Height          =   585
         Left            =   400
         TabIndex        =   29
         Top             =   100
         Width           =   5500
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   7400
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
      Height          =   600
      Left            =   400
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7400
      Visible         =   0   'False
      Width           =   600
      _Version        =   1048579
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   64
      Show3DBorder    =   2
   End
End
Attribute VB_Name = "frmBuAnp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxBel As XtremeSuiteControls.FlatEdit
Private CmRam As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmSaK As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmSto As XtremeSuiteControls.ComboBox
Private CmAus As XtremeSuiteControls.ComboBox
Private CmReS As XtremeSuiteControls.ComboBox
Private ChGeg As XtremeSuiteControls.CheckBox
Private ChDat As XtremeSuiteControls.CheckBox
Private ChMan As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private ChTyp As XtremeSuiteControls.CheckBox
Private ChSto As XtremeSuiteControls.CheckBox
Private ChAus As XtremeSuiteControls.CheckBox
Private ChSaK As XtremeSuiteControls.CheckBox
Private ChSte As XtremeSuiteControls.CheckBox
Private ChBel As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private UpCo2 As XtremeSuiteControls.UpDown
Private MoKal As XtremeCalendarControl.DatePicker
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow

Private FoLad As Boolean
Private Sub TAbs()
On Error GoTo OpErr
'Ändert die Rechnungen

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

    S_BuAnp
    DoEvents
    If AnzPo > 1 Then
        SUpBu 0, True
    Else
        Set RpSel = RpCo1.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpBu RowNr
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

Set FM = frmBuAnp
Set Rahm1 = FM.frmRahm1
Set Rahm0 = FM.frmRahm0
Set MoKal = FM.dtpDatu1
Set TxDa1 = FM.txtDatu1
Set TxBel = FM.txtBeleg
Set CmGeg = FM.cmbGegen
Set CmMan = FM.cmbManda
Set CmMit = FM.cmbMitar
Set CmTyp = FM.cmbBuTyp
Set CmSto = FM.cmbStorn
Set CmAus = FM.cmbAuswe
Set CmSaK = FM.cmbSaKon
Set CmRam = FM.cmbKtoRa
Set ChSte = FM.chkSteue
Set ChGeg = FM.chkGegen
Set ChDat = FM.chkDatum
Set ChMan = FM.chkManda
Set ChMit = FM.chkMitar
Set ChTyp = FM.chkReTyp
Set ChSto = FM.chkStorn
Set ChAus = FM.chkAuswe
Set ChSaK = FM.chkSaKon
Set ChBel = FM.chkBeleg
Set CmReS = FM.cmbReStu

Set PuBu1 = FM.btnDatu1
Set ImMan = frmMain.imgManag

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
        .ListIndex = 0
    End If
End With

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(CmMan.NewIndex) = GlThe(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
    CmMit.AddItem GlMiK(AktZa, 1)
    CmMit.ItemData(CmMit.NewIndex) = GlMiK(AktZa, 2)
Next AktZa

With CmTyp
    .AddItem "Ausgabe"
    .ItemData(0) = 1
    .AddItem "Einnahme"
    .ItemData(1) = 2
End With

With CmSto
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmAus
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

For AktZa = 1 To UBound(GlStu)
    CmReS.AddItem GlStu(AktZa, 2)
    CmReS.ItemData(AktZa - 1) = GlStu(AktZa, 0)
Next AktZa

If (GlKtR - 1) <= (CmRam.ListCount) - 1 Then
    CmRam.ListIndex = GlKtR - 1
Else
    CmRam.ListIndex = 0
End If

CmTyp.ListIndex = 0
CmSto.ListIndex = 1
CmAus.ListIndex = 1
CmMan.ListIndex = GlSMa - 1
CmMit.ListIndex = GlSmI - 1
CmReS.ListIndex = 0

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

If GlBMa = True Then 'Getrennter Mandanten Belegnummernkreis
    ChMan.Enabled = False
ElseIf GlBGe = True Then 'Getrennter Geldkonten Belegnummernkreis
    ChMan.Enabled = False
End If

ChMit.Enabled = GlMiV

With TxBel
    .Pattern = "\d*"
    .SetMask "000000", "______"
    .Text = "000001"
End With

If GlBuc = True Then 'einfache Buchhaltung verwenden
    ChGeg.Caption = "Geldkonto :"
    ChSaK.Caption = "Sachkonto :"
Else
    ChGeg.Caption = "Sollkonto :"
    ChSaK.Caption = "Habenkonto :"
End If

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
ChGeg.BackColor = GlBak
ChDat.BackColor = GlBak
ChMan.BackColor = GlBak
ChMit.BackColor = GlBak
ChTyp.BackColor = GlBak
ChSto.BackColor = GlBak
ChAus.BackColor = GlBak
ChSaK.BackColor = GlBak
ChSte.BackColor = GlBak
ChBel.BackColor = GlBak

Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TInit " & Err.Number
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
Private Sub FMnd()
On Error GoTo LaErr
'Suchten den Kontenrahmen des Mandanten

Dim ManNr As Long
Dim AktZa As Integer
Dim StaRa As Integer

Set CmMan = Me.cmbManda
Set CmRam = Me.cmbKtoRa

ManNr = CmMan.ItemData(CmMan.ListIndex)

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            If GlMan(AktZa, 25) <> vbNullString Then
                StaRa = GlMan(AktZa, 25) 'Standardkontenrahmen
            Else
                StaRa = GlKtR
            End If
            Exit For
        End If
    Next AktZa
    If (StaRa - 1) <= (CmRam.ListCount) - 1 Then
        CmRam.ListIndex = StaRa - 1
    Else
        CmRam.ListIndex = 0
    End If
End If

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMnd " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = CDate(MoKal.Selection.Blocks(0).DateBegin)
    If Year(NeuDa) <= Year(Date) Then
        TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
    Else
        SPopu "Buchungsjahr überschritten", "Das neue Buchungsdatum muss sich im selben Buchungsjahr befinden.", IC48_Information
    End If
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub

Private Sub chkAuswe_Click()
On Error Resume Next

Set ChAus = Me.chkAuswe
Set CmAus = Me.cmbAuswe

If ChAus.Value = xtpChecked Then
    CmAus.Enabled = True
Else
    CmAus.Enabled = False
End If

End Sub

Private Sub chkBeleg_Click()
On Error Resume Next

Set ChBel = Me.chkBeleg
Set TxBel = Me.txtBeleg

If ChBel.Value = xtpChecked Then
    TxBel.Enabled = True
Else
    TxBel.Enabled = False
End If

End Sub
Private Sub chkDatum_Click()
On Error Resume Next

Set ChDat = Me.chkDatum
Set TxDa1 = Me.txtDatu1
Set UpCo2 = Me.updCont2
Set PuBu1 = Me.btnDatu1

If ChDat.Value = xtpChecked Then
    TxDa1.Enabled = True
    UpCo2.Enabled = True
    PuBu1.Enabled = True
Else
    TxDa1.Enabled = False
    UpCo2.Enabled = False
    PuBu1.Enabled = False
End If

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
Private Sub chkReTyp_Click()
On Error Resume Next

Set ChTyp = Me.chkReTyp
Set CmTyp = Me.cmbBuTyp

If ChTyp.Value = xtpChecked Then
    CmTyp.Enabled = True
Else
    CmTyp.Enabled = False
End If

End Sub
Private Sub chkSaKon_Click()
On Error Resume Next

Set ChSaK = Me.chkSaKon
Set CmSaK = Me.cmbSaKon

If ChSaK.Value = xtpChecked Then
    CmSaK.Enabled = True
Else
    CmSaK.Enabled = False
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

Private Sub cmbKtoRa_Click()
    If FoLad = False Then
        S_KtCm
    End If
End Sub
Private Sub cmbManda_Click()
    If FoLad = False Then
        FMnd
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True

FInit
S_KtCm

FoLad = False

AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = CDate(TxDa1.Text)
    If Year(NeuDa) > Year(Date) Then
        NeuDa = "31.12." & Year(Date)
        SPopu "Buchungsjahr überschritten", "Das neue Buchungsdatum muss sich im selben Buchungsjahr befinden.", IC48_Information
    End If
    TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
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

TeTit = "Buchungen Anpassen"
TeMai = "Ändert bestimmte Eigenschaften für die markierten Buchungen."
TeInh = "Diese Funktion wird verwendet, wenn Änderungen an mehreren Buchungen gleichzeitig vorgenommen werden sollen. So zum Beispiel die Änderung eines Sach- oder Geldkontos oder das erneue durchnummerieren der Buchungsnummern."
TeFus = "Um mehrere Buchungen zu markieren ist es gegebenenfalls sinnvoll, die Buchungen vorher zu sortieren, indem auf den jeweiligen Spaltenkopf einer Spalte geklickt wird oder die Buchungen über die gleichnamigen Funktionen links gruppiert werden. Durch das drücken der Shift- oder Strg-Taste können dann mehrere Buchungen auf einmal markiert werden."

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBuAnp = Nothing
End Sub
Private Sub txtBeleg_GotFocus()
    Me.txtBeleg.SelStart = 0
    Me.txtBeleg.SelLength = Len(Me.txtBeleg.Text)
End Sub
Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub
Private Sub updCont2_DownClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub

Private Sub updCont2_UpClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub
Private Sub btnWeiter_Click()
    TAbs
    Unload Me
End Sub
