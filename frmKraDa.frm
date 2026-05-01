VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmKraDa 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Einträge Anpassen"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.CheckBox chkStorn 
      Height          =   255
      Left            =   3400
      TabIndex        =   31
      Top             =   2700
      Width           =   975
      _Version        =   1048579
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Entfernt"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   22
      Top             =   4500
      Width           =   5800
      _Version        =   1048579
      _ExtentX        =   10231
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   3800
         TabIndex        =   25
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
         Default         =   -1  'True
         Height          =   400
         Left            =   2400
         TabIndex        =   24
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
         Left            =   1100
         TabIndex        =   23
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
   Begin XtremeSuiteControls.CheckBox chkPatie 
      Height          =   225
      Left            =   800
      TabIndex        =   18
      Top             =   3600
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Patient"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.UpDown updCont2 
      Height          =   350
      Left            =   1820
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2100
      Width           =   255
      _Version        =   1048579
      _ExtentX        =   450
      _ExtentY        =   600
      _StockProps     =   64
      Enabled         =   0   'False
      Min             =   1
      Value           =   1
      SyncBuddy       =   -1  'True
      AutoBuddy       =   -1  'True
      BuddyControl    =   ""
      BuddyProperty   =   ""
   End
   Begin XtremeSuiteControls.UpDown updCont1 
      Height          =   350
      Left            =   2170
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
      _Version        =   1048579
      _ExtentX        =   450
      _ExtentY        =   617
      _StockProps     =   64
      Enabled         =   0   'False
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtDatu1"
      BuddyProperty   =   ""
   End
   Begin XtremeSuiteControls.CheckBox chkKrTyp 
      Height          =   225
      Left            =   3400
      TabIndex        =   5
      Top             =   900
      Width           =   1605
      _Version        =   1048579
      _ExtentX        =   2831
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Eintragstyp"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   0
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   5800
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkDatum 
      Height          =   225
      Left            =   800
      TabIndex        =   1
      Top             =   900
      Width           =   1605
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Datum"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkAnzal 
      Height          =   225
      Left            =   800
      TabIndex        =   7
      Top             =   1800
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Anzahl"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAnzal 
      Height          =   350
      Left            =   800
      TabIndex        =   8
      Top             =   2100
      Width           =   1000
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      Text            =   "1"
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.PushButton btnDatu1 
      Height          =   350
      Left            =   2460
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Öffnet den Auswahlkalender"
      Top             =   1200
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
      Left            =   800
      TabIndex        =   2
      Top             =   1200
      Width           =   1360
      _Version        =   1048579
      _ExtentX        =   2399
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.ComboBox cmbKrTyp 
      Height          =   315
      Left            =   3400
      TabIndex        =   6
      Top             =   1200
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2805
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkPreis 
      Height          =   225
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   1305
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Einzelpreis"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPreis 
      Height          =   350
      Left            =   2400
      TabIndex        =   11
      Top             =   2100
      Width           =   1095
      _Version        =   1048579
      _ExtentX        =   1931
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      Text            =   "0,00"
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkAbges 
      Height          =   225
      Left            =   3400
      TabIndex        =   26
      Top             =   1800
      Width           =   1305
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Abschließen"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbAbges 
      Height          =   315
      Left            =   3400
      TabIndex        =   27
      Top             =   2100
      Width           =   1100
      _Version        =   1048579
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkMulti 
      Height          =   225
      Left            =   3900
      TabIndex        =   12
      Top             =   1800
      Width           =   1305
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Multiplikator"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtMulti 
      Height          =   350
      Left            =   3900
      TabIndex        =   13
      Top             =   2100
      Width           =   1100
      _Version        =   1048579
      _ExtentX        =   1940
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      Text            =   "1,00"
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkKetKe 
      Height          =   225
      Left            =   800
      TabIndex        =   14
      Top             =   2700
      Width           =   2000
      _Version        =   1048579
      _ExtentX        =   3528
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Kettenkennzeichnung"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbKetKe 
      Height          =   315
      Left            =   800
      TabIndex        =   15
      Top             =   3000
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5133
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbPatie 
      Height          =   315
      Left            =   800
      TabIndex        =   19
      Top             =   3900
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5133
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkAnalo 
      Height          =   225
      Left            =   3900
      TabIndex        =   20
      Top             =   3600
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Analog"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbAnalo 
      Height          =   315
      Left            =   3900
      TabIndex        =   21
      Top             =   3900
      Width           =   1100
      _Version        =   1048579
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtSteue 
      Height          =   350
      Left            =   3900
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
      _Version        =   1048579
      _ExtentX        =   1931
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      Text            =   "0,00"
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkSteue 
      Height          =   225
      Left            =   3900
      TabIndex        =   16
      Top             =   2700
      Width           =   1305
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Steuersatz"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbStorn 
      Height          =   315
      Left            =   3400
      TabIndex        =   30
      Top             =   3000
      Width           =   1095
      _Version        =   1048579
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label lblLabe1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte wählen Sie, welche Änderungen an den markierten Einträgen vorgenommen werden sollen und klicken auf Weiter."
      Height          =   435
      Left            =   800
      TabIndex        =   29
      Top             =   100
      Width           =   4500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   660
      Left            =   0
      Top             =   0
      Width           =   5800
   End
End
Attribute VB_Name = "frmKraDa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private ImMan As XtremeCommandBars.ImageManager
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxAnz As XtremeSuiteControls.FlatEdit
Private TxPre As XtremeSuiteControls.FlatEdit
Private TxMul As XtremeSuiteControls.FlatEdit
Private TxStu As XtremeSuiteControls.FlatEdit
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmAbg As XtremeSuiteControls.ComboBox
Private CmKet As XtremeSuiteControls.ComboBox
Private CmPat As XtremeSuiteControls.ComboBox
Private CmAna As XtremeSuiteControls.ComboBox
Private CmSto As XtremeSuiteControls.ComboBox
Private CheDa As XtremeSuiteControls.CheckBox
Private CheTy As XtremeSuiteControls.CheckBox
Private CheAn As XtremeSuiteControls.CheckBox
Private ChePr As XtremeSuiteControls.CheckBox
Private CheAb As XtremeSuiteControls.CheckBox
Private CheMu As XtremeSuiteControls.CheckBox
Private CheSt As XtremeSuiteControls.CheckBox
Private CheKe As XtremeSuiteControls.CheckBox
Private ChPat As XtremeSuiteControls.CheckBox
Private ChAna As XtremeSuiteControls.CheckBox
Private ChSto As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private UpDo1 As XtremeSuiteControls.UpDown
Private UpDo2 As XtremeSuiteControls.UpDown
Private MoKal As XtremeCalendarControl.DatePicker
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn
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
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1
Set CheDa = Me.chkDatum

CheDa.Value = xtpChecked

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
Private Sub FKonf()
On Error GoTo SuErr

Dim NeuDa As Date
Dim AktZa As Integer
Dim AktPo As Integer
Dim ReAbg As Boolean
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set Rahm0 = Me.frmRahm0
Set UpDo1 = Me.updCont1
Set UpDo2 = Me.updCont2
Set MoKal = Me.dtpDatu1
Set TxDa1 = Me.txtDatu1
Set TxAnz = Me.txtAnzal
Set TxMul = Me.txtMulti
Set TxStu = Me.txtSteue
Set TxPre = Me.txtPreis
Set CmTyp = Me.cmbKrTyp
Set CmAbg = Me.cmbAbges
Set CmKet = Me.cmbKetKe
Set CmAna = Me.cmbAnalo
Set CmPat = Me.cmbPatie
Set CmSto = Me.cmbStorn
Set CheDa = Me.chkDatum
Set CheTy = Me.chkKrTyp
Set CheAn = Me.chkAnzal
Set ChePr = Me.chkPreis
Set CheMu = Me.chkMulti
Set CheSt = Me.chkSteue
Set CheAb = Me.chkAbges
Set CheKe = Me.chkKetKe
Set ChPat = Me.chkPatie
Set ChAna = Me.chkAnalo
Set ChSto = Me.chkStorn
Set PuBu1 = Me.btnDatu1
Set RpCo3 = FM.repCont3
Set RpCo6 = FM.repCont6
Set RpCoK = FM.repContK
Set ImMan = FM.imgManag

Select Case GlBut
Case RibTab_Krankenbla:
        Set RpSel = RpCoK.SelectedRows
        Set RpCls = RpCoK.Columns
Case RibTab_Abrechnung:
        Set RpSel = RpCo3.SelectedRows
        Set RpCls = RpCo3.Columns
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Rec_Selekt)
                ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                ReAbg = True
            End If
        Else
            ReAbg = True
        End If
        Set RpSel = RpCo6.SelectedRows
        Set RpCls = RpCo6.Columns
Case RibTab_LabBericht:
        Set RpSel = RpCo6.SelectedRows
        Set RpCls = RpCo6.Columns
End Select

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

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Kra_Datum)
        If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
            NeuDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
        Else
            NeuDa = Date
        End If
    Else
        NeuDa = Date
    End If
Else
    NeuDa = Date
End If

With CmAbg
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
    .ListIndex = 0
End With

With CmAna
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
    .ListIndex = 0
End With

With CmSto
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
    .ListIndex = 1
End With

With CmKet
    .AddItem "Summation einer Abrechnungsfolge"
    .AddItem "Summation einer analogen Abrechnungsfolge"
    .ListIndex = 0
End With

Select Case GlBut
Case RibTab_Krankenbla:
    With CmTyp
        For AktZa = 1 To UBound(GlKrA)
            If GlKrA(AktZa, 0) > 9 Then
                .AddItem GlKrA(AktZa, 1) & " - " & GlKrA(AktZa, 2)
                .ItemData(AktPo) = GlKrA(AktZa, 0)
                AktPo = AktPo + 1
            End If
        Next AktZa
        .ListIndex = 1
        .AutoComplete = False
        .DropDownWidth = 1900
        .DropDownItemCount = UBound(GlKrA) - 9
    End With
Case RibTab_Abrechnung:
    With CmTyp
        For AktZa = 1 To UBound(GlKrA)
            If GlKrA(AktZa, 0) < 10 Then
                .AddItem GlKrA(AktZa, 1) & " - " & GlKrA(AktZa, 2)
                .ItemData(AktPo) = GlKrA(AktZa, 0)
                AktPo = AktPo + 1
            End If
        Next AktZa
        If GlStS > 1 Then
            .ListIndex = 8
        Else
            .ListIndex = 1
        End If
        .AutoComplete = False
        .DropDownWidth = 1900
        .DropDownItemCount = 9
    End With
Case RibTab_LabBericht:
    With CmTyp
        For AktZa = 1 To UBound(GlKrA)
            If GlKrA(AktZa, 0) > 9 Then
                .AddItem GlKrA(AktZa, 1) & " - " & GlKrA(AktZa, 2)
                .ItemData(AktPo) = GlKrA(AktZa, 0)
                AktPo = AktPo + 1
            End If
        Next AktZa
        .ListIndex = 1
        .AutoComplete = False
        .DropDownWidth = 1900
        .DropDownItemCount = UBound(GlKrA) - 9
    End With
End Select

Select Case GlBut
Case RibTab_Abrechnung:
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        With CmPat
            .AddItem GlMiA(AktZa, 1)
            .ItemData(AktZa - 1) = GlMiA(AktZa, 2)
        End With
    Next AktZa
    CmPat.ListIndex = 0
Case RibTab_LabBericht:
    For AktZa = 1 To UBound(GlLGr) 'Laborgruppen
        With CmPat
            .AddItem GlLGr(AktZa, 1)
            .ItemData(AktZa - 1) = GlLGr(AktZa, 0)
        End With
    Next AktZa
    CmPat.ListIndex = 0
End Select

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = NeuDa
End With

If GlBut = RibTab_Krankenbla Then
    With TxAnz
        .SetMask "00:00", "__:__"
        .Text = Format$(Now, "hh:mm")
    End With
Else
    With TxAnz
        .Pattern = "\d*"
        .SetMask "0", "_"
        .Text = 1
    End With
End If

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

If GlDDa = True Then 'Dauerdiagnosen anpassen
    CheDa.Value = xtpChecked
    CheTy.Visible = False
    CheAn.Visible = False
    ChePr.Visible = False
    CheAb.Visible = False
    CheMu.Visible = False
    CheSt.Visible = False
    CheKe.Visible = False
    ChSto.Visible = False
    CmTyp.Visible = False
    CmAbg.Visible = False
    CmKet.Visible = False
    CmSto.Visible = False
    TxAnz.Visible = False
    TxPre.Visible = False
    TxMul.Visible = False
    UpDo2.Visible = False
End If

Select Case GlBut
Case RibTab_Krankenbla:
    ChePr.Visible = False
    CheMu.Visible = False
    CheSt.Visible = False
    CheKe.Visible = False
    CmKet.Visible = False
    ChAna.Visible = False
    TxStu.Visible = False
    TxPre.Visible = False
    TxMul.Visible = False
    CmAna.Visible = False
    ChPat.Caption = "Patient"
    CheAn.Caption = "Uhrzeit"
Case RibTab_Abrechnung:
    CheAb.Visible = False
    ChSto.Visible = False
    CmAbg.Visible = False
    CmSto.Visible = False
    ChPat.Caption = "Mitarbeiter"
Case RibTab_LabBericht:
    CheDa.Visible = False
    ChSto.Visible = False
    TxDa1.Visible = False
    UpDo1.Visible = False
    PuBu1.Visible = False
    CheAb.Visible = False
    CmAbg.Visible = False
    CheTy.Visible = False
    CmTyp.Visible = False
    CmSto.Visible = False
    CheAn.Visible = False
    TxAnz.Visible = False
    UpDo2.Visible = False
    CheSt.Visible = False
    TxStu.Visible = False
    CheKe.Visible = False
    CmKet.Visible = False
    ChAna.Visible = False
    CmAna.Visible = False
    ChPat.Caption = "Laborgruppe"
End Select

If ReAbg = True Then
    CheAn.Enabled = False
    CheMu.Enabled = False
    ChePr.Enabled = False
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
CheDa.BackColor = GlBak
CheTy.BackColor = GlBak
CheAn.BackColor = GlBak
ChePr.BackColor = GlBak
CheAb.BackColor = GlBak
CheMu.BackColor = GlBak
CheSt.BackColor = GlBak
CheKe.BackColor = GlBak
ChPat.BackColor = GlBak
ChAna.BackColor = GlBak
ChSto.BackColor = GlBak

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo3 = Nothing
Set RpCo6 = Nothing
Set RpCoK = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub TWeit()
On Error GoTo SuErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

If GlDDa = True Then
    Dia_Da NeuDa
Else
    S_KrDa
End If

Set MoKal = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TWeit " & Err.Number
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
    Unload Me
End Sub

Private Sub chkAbges_Click()
On Error Resume Next

Set CheAb = Me.chkAbges
Set CmAbg = Me.cmbAbges

If CheAb.Value = xtpChecked Then
    CmAbg.Enabled = True
Else
    CmAbg.Enabled = False
End If

End Sub
Private Sub chkAnalo_Click()
On Error Resume Next

Set ChAna = Me.chkAnalo
Set CmAna = Me.cmbAnalo

If ChAna.Value = xtpChecked Then
    CmAna.Enabled = True
Else
    CmAna.Enabled = False
End If

End Sub
Private Sub chkAnzal_Click()
On Error Resume Next

Set CheAn = Me.chkAnzal
Set TxAnz = Me.txtAnzal
Set UpDo2 = Me.updCont2

If CheAn.Value = xtpChecked Then
    TxAnz.Enabled = True
    UpDo2.Enabled = True
Else
    TxAnz.Enabled = False
    UpDo2.Enabled = False
End If

End Sub
Private Sub chkDatum_Click()
On Error Resume Next

Set CheDa = Me.chkDatum
Set TxDa1 = Me.txtDatu1
Set UpDo1 = Me.updCont1
Set PuBu1 = Me.btnDatu1

If CheDa.Value = xtpChecked Then
    TxDa1.Enabled = True
    UpDo1.Enabled = True
    PuBu1.Enabled = True
Else
    TxDa1.Enabled = False
    UpDo1.Enabled = False
    PuBu1.Enabled = False
End If

End Sub

Private Sub chkKetKe_Click()
On Error Resume Next

Set CheKe = Me.chkKetKe
Set CmKet = Me.cmbKetKe

If CheKe.Value = xtpChecked Then
    CmKet.Enabled = True
Else
    CmKet.Enabled = False
End If

End Sub
Private Sub chkKrTyp_Click()
On Error Resume Next

Set CheTy = Me.chkKrTyp
Set CmTyp = Me.cmbKrTyp

If CheTy.Value = xtpChecked Then
    CmTyp.Enabled = True
Else
    CmTyp.Enabled = False
End If

End Sub

Private Sub chkMulti_Click()
On Error Resume Next

Set CheMu = Me.chkMulti
Set TxMul = Me.txtMulti

If CheMu.Value = xtpChecked Then
    TxMul.Enabled = True
Else
    TxMul.Enabled = False
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

If GlBut = RibTab_Krankenbla Then
    If CmPat.ListCount = 0 Then
        S_KrPa
    End If
End If

End Sub
Private Sub chkPreis_Click()
On Error Resume Next

Set ChePr = Me.chkPreis
Set TxPre = Me.txtPreis

If ChePr.Enabled = True Then
    TxPre.Enabled = True
Else
    TxPre.Enabled = False
End If

End Sub

Private Sub chkSteue_Click()
On Error Resume Next

Set CheSt = Me.chkSteue
Set TxStu = Me.txtSteue

If CheSt.Value = xtpChecked Then
    TxStu.Enabled = True
Else
    TxStu.Enabled = False
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

Set MoKal = Nothing

End Sub
Private Sub Form_Load()
On Error Resume Next

FKonf
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    GlDDa = False
    Set frmKraDa = Nothing
End Sub
Private Sub txtAnzal_GotFocus()
    Me.txtAnzal.SelStart = 0
    Me.txtAnzal.SelLength = Len(Me.txtAnzal.Text)
End Sub
Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub

Private Sub txtMulti_GotFocus()
    Me.txtMulti.SelStart = 0
    Me.txtMulti.SelLength = Len(Me.txtMulti.Text)
End Sub
Private Sub txtPreis_GotFocus()
    Me.txtPreis.SelStart = 0
    Me.txtPreis.SelLength = Len(Me.txtPreis.Text)
End Sub

Private Sub txtSteue_GotFocus()
    Me.txtSteue.SelStart = 0
    Me.txtSteue.SelLength = Len(Me.txtSteue.Text)
End Sub
Private Sub updCont1_DownClick()
On Error Resume Next

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = CDate(TxDa1.Text)

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont1_UpClick()
On Error Resume Next

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = CDate(TxDa1.Text)

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub

Private Sub updCont2_DownClick()
On Error Resume Next

Dim AlZei As Date
Dim NeZei As Date
Dim AlZal As Integer

Set TxAnz = Me.txtAnzal

If GlBut = RibTab_Krankenbla Then
    AlZei = TimeValue(TxAnz.Text)
    NeZei = DateAdd("n", -1, AlZei)
    TxAnz.Text = Format$(NeZei, "hh:mm")
Else
    AlZal = CInt(TxAnz.Text)
    If (AlZal - 1) > 0 Then
        TxAnz.Text = AlZal - 1
    End If
End If

End Sub
Private Sub updCont2_UpClick()
On Error Resume Next

Dim AlZei As Date
Dim NeZei As Date
Dim AlZal As Integer

Set TxAnz = Me.txtAnzal

If GlBut = RibTab_Krankenbla Then
    AlZei = TimeValue(TxAnz.Text)
    NeZei = DateAdd("n", 1, AlZei)
    TxAnz.Text = Format$(NeZei, "hh:mm")
Else
    AlZal = CInt(TxAnz.Text)
    If (AlZal + 1) < 99 Then
        TxAnz.Text = AlZal + 1
    End If
End If

End Sub
