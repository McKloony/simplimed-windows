VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmRzNeu 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Beleg Hinzufügen"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   6
      Top             =   4200
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
         TabIndex        =   10
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
      Begin XtremeSuiteControls.PushButton btnWeite 
         Default         =   -1  'True
         Height          =   400
         Left            =   3200
         TabIndex        =   9
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
      Begin XtremeSuiteControls.PushButton btnZuruk 
         Height          =   400
         Left            =   1800
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zurück"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   400
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hilfe"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   0
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5500
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4000
      Left            =   700
      TabIndex        =   1
      Top             =   100
      Width           =   5200
      _Version        =   1048579
      _ExtentX        =   9172
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.RadioButton optAnder 
         Height          =   220
         Left            =   1300
         TabIndex        =   2
         Top             =   1300
         Width           =   3800
         _Version        =   1048579
         _ExtentX        =   6703
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "für einen anderen Patienten"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optSelbe 
         Height          =   220
         Left            =   1300
         TabIndex        =   3
         Top             =   900
         Width           =   3800
         _Version        =   1048579
         _ExtentX        =   6703
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "für den ausgewählten Patienten"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   310
         Left            =   1300
         TabIndex        =   4
         Top             =   2670
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkChek1 
         Height          =   220
         Left            =   1320
         TabIndex        =   34
         Top             =   1900
         Width           =   3400
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Aktuelle Diagnose auf Rezept einfügen?"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   310
         Left            =   1300
         TabIndex        =   5
         Top             =   3370
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   200
         Left            =   1320
         TabIndex        =   39
         Top             =   3130
         Width           =   1400
      End
      Begin VB.Label lblLab10 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   200
         Left            =   1320
         TabIndex        =   12
         Top             =   2430
         Width           =   1400
      End
      Begin VB.Label lblLabl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Für welchen Patienten möchten Sie einen neuen Beleg anlegen? Wählen Sie eine Option und klicken auf Weiter."
         Height          =   450
         Left            =   100
         TabIndex        =   11
         Top             =   100
         Width           =   5000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4000
      Left            =   700
      TabIndex        =   22
      Top             =   100
      Visible         =   0   'False
      Width           =   5200
      _Version        =   1048579
      _ExtentX        =   9172
      _ExtentY        =   7056
      _StockProps     =   79
      Caption         =   "GroupBox2"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtBemer 
         Height          =   350
         Left            =   800
         TabIndex        =   23
         Top             =   2220
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtPost 
         Height          =   350
         Left            =   800
         TabIndex        =   24
         Top             =   1440
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   350
         Left            =   800
         TabIndex        =   25
         Top             =   680
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin VB.Label lblLabl2 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Name"
         Height          =   195
         Left            =   810
         TabIndex        =   28
         Top             =   400
         Width           =   3000
      End
      Begin VB.Label lblLabl3 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Postleitzahl"
         Height          =   195
         Left            =   810
         TabIndex        =   27
         Top             =   1160
         Width           =   3000
      End
      Begin VB.Label lblLabl4 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Bemerkung"
         Height          =   195
         Left            =   810
         TabIndex        =   26
         Top             =   1940
         Width           =   3000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4000
      Left            =   700
      TabIndex        =   29
      Top             =   100
      Visible         =   0   'False
      Width           =   5200
      _Version        =   1048579
      _ExtentX        =   9172
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListBox lstList1 
         Height          =   3000
         Left            =   500
         TabIndex        =   30
         Top             =   700
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7056
         _ExtentY        =   5292
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLabl5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie einen der gefundenen Einträge"
         Height          =   200
         Left            =   520
         TabIndex        =   31
         Top             =   250
         Width           =   3600
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   4000
      Left            =   700
      TabIndex        =   13
      Top             =   100
      Visible         =   0   'False
      Width           =   5200
      _Version        =   1048579
      _ExtentX        =   9172
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   2300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
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
         Height          =   350
         Left            =   2580
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   960
         Width           =   310
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbFormu 
         Height          =   310
         Left            =   1000
         TabIndex        =   17
         Top             =   1660
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4763
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   1000
         TabIndex        =   14
         Top             =   960
         Width           =   1280
         _Version        =   1048579
         _ExtentX        =   2258
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtKomme 
         Height          =   350
         Left            =   1000
         TabIndex        =   18
         Top             =   2350
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar :"
         Height          =   220
         Left            =   1010
         TabIndex        =   33
         Top             =   2120
         Width           =   1600
      End
      Begin VB.Label lblLabl9 
         BackStyle       =   0  'Transparent
         Caption         =   "Formularauswahl :"
         Height          =   220
         Left            =   1010
         TabIndex        =   21
         Top             =   1420
         Width           =   1600
      End
      Begin VB.Label lblLabl8 
         BackStyle       =   0  'Transparent
         Caption         =   "Belegdatum :"
         Height          =   220
         Left            =   1010
         TabIndex        =   20
         Top             =   720
         Width           =   1600
      End
      Begin VB.Label lblLabl6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie das gewünschte Formular und Belegdatum aus"
         Height          =   600
         Left            =   100
         TabIndex        =   19
         Top             =   100
         Width           =   5200
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   4000
      Left            =   700
      TabIndex        =   35
      Top             =   100
      Visible         =   0   'False
      Width           =   5200
      _Version        =   1048579
      _ExtentX        =   9172
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView1 
         Height          =   2300
         Left            =   100
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1100
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   4057
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.CheckBox chkRzMar 
         Height          =   220
         Left            =   200
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Einträge automatisch markieren"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie, welche der folgenden Einträge mit auf das neue Rezept übertragen werden sollen und klicken auf Weiter."
         Height          =   450
         Left            =   100
         TabIndex        =   37
         Top             =   100
         Width           =   5000
      End
   End
End
Attribute VB_Name = "frmRzNeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxDaV As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxKom As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private FTex3 As XtremeSuiteControls.FlatEdit
Private FLis1 As XtremeSuiteControls.ListBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private ComFo As XtremeSuiteControls.ComboBox
Private Chek1 As XtremeSuiteControls.CheckBox
Private Chek2 As XtremeSuiteControls.CheckBox
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private LiVw1 As XtremeSuiteControls.ListView
Private LiItm As XtremeSuiteControls.ListViewItem
Private LiIts As XtremeSuiteControls.ListViewItems
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private MoKal As XtremeCalendarControl.DatePicker
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager

Private mPaNr As Long 'Patient
Private mMaNr As Long 'Mandant
Private KatLa As Boolean

Private clFen As clsFenster
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
Set Rahm4 = Me.frmRahm4

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = Rahm4.Top + TxDa1.Top + TxDa1.Height
    .Left = Rahm4.Left + TxDa1.Left
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

Dim IdxNr As Long
Dim ManNr As Long
Dim PaVor As String
Dim mPaKu As String
Dim MaIdx As Integer
Dim AktZa As Integer
Dim KatAn As Integer
Dim LiIdx As Integer
Dim TeWer As Variant
Dim BeVor As Boolean
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set LiVw1 = Me.lstView1
Set TxDa1 = Me.txtDatu1
Set TxKom = Me.txtKomme
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set ComFo = Me.cmbFormu
Set Chek1 = Me.chkChek1
Set Chek2 = Me.chkRzMar
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set PuBu1 = Me.btnDatu1
Set PuBu2 = Me.btnZuruk
Set MoKal = Me.dtpDatu1
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns
Set RpSel = RpCo5.SelectedRows
Set ImMan = frmMain.imgManag

Select Case GlBut
Case RibTab_Rezeptmodul: LiIdx = CLng(Right$(IniGetVal("System", "RzForm"), 2))
Case RibTab_Belegmodul: LiIdx = CLng(Right$(IniGetVal("System", "BlForm"), 2))
End Select

If GlKoZ = True Then 'Rezept kopieren
    Me.Caption = "Beleg Kopieren"
Else
    Me.Caption = "Neuer Beleg"
End If

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

With LiVw1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = True
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = True
    .FlatScrollBar = False
    .Font.SIZE = 10 'GlTFt.size
    .Font.Name = GlTFt.Name
    .ForeColor = vbBlack
    .FullRowSelect = True
    .GridLines = False
    .HideColumnHeaders = False
    .HideSelection = False
    .HotTracking = False
    .HoverSelection = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpListViewLabelManual
    .LabelWrap = True
    .MultiSelect = False
    .OLEDropMode = xtpOLEDropNone
    .View = xtpListViewReport
End With

Select Case GlBut
Case RibTab_Startseite:
    With ComFo
        For AktZa = 1 To UBound(GlRzV)
            .AddItem GlRzV(AktZa)
            .ItemData(AktZa - 1) = AktZa
        Next AktZa
        If LiIdx - 1 > UBound(GlRzV) Then LiIdx = 1
        .ListIndex = LiIdx - 1
    End With
Case RibTab_Rezeptmodul:
    With ComFo
        For AktZa = 1 To UBound(GlRzV)
            .AddItem GlRzV(AktZa)
            .ItemData(AktZa - 1) = AktZa
        Next AktZa
        If LiIdx - 1 > UBound(GlRzV) Then LiIdx = 1
        .ListIndex = LiIdx - 1
    End With
Case RibTab_Belegmodul:
    With ComFo
        For AktZa = 1 To UBound(GlBlV)
            .AddItem GlBlV(AktZa)
            .ItemData(AktZa - 1) = AktZa
        Next AktZa
        If LiIdx - 1 > UBound(GlBlV) Then LiIdx = 1
        .ListIndex = LiIdx - 1
    End With
End Select

KatAn = IniGetVal("System", "KatAut")

If GlBut = RibTab_Belegmodul Then
    Chek1.Visible = False
End If

If GlKoZ = True Then 'Rezept kopieren
    Chek1.Enabled = False
Else
    If KatAn = -1 Then
        Chek1.Value = 1
    Else
        Chek1.Value = 0
    End If
End If

If GlBut <> RibTab_Startseite Then
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        Set RpCol = RpCls.Find(Rzp_ID0)
        mPaNr = RpRow.Record(RpCol.ItemIndex).Value
        Set RpCol = RpCls.Find(Rzp_IDP)
        ManNr = RpRow.Record(RpCol.ItemIndex).Value
        Set RpCol = RpCls.Find(Rzp_Vorname)
        PaVor = RpRow.Record(RpCol.ItemIndex).Value
        Set RpCol = RpCls.Find(Rzp_Name)
        mPaKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
        Opti1.Caption = "für " & PaVor & Chr$(32) & mPaKu
    Else
        mPaNr = GlAdr
        ManNr = S_AdIdx(mPaNr, "IDP")
    End If
Else
    mPaNr = GlAdr
    ManNr = S_AdIdx(mPaNr, "IDP")
End If

For AktZa = 1 To UBound(GlMaA) 'Aktive Mandanten
    CmMan.AddItem GlMaA(AktZa, 1)
    CmMan.ItemData(AktZa - 1) = GlMaA(AktZa, 2)
Next AktZa

For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
    CmMit.AddItem GlMiA(AktZa, 1)
    CmMit.ItemData(AktZa - 1) = GlMiA(AktZa, 2)
Next AktZa

S_AdDe mPaNr 'Adressendetails
With GlADt
    TeWer = .AdMan
    mPaKu = .AdKur
End With
DoEvents

If GlRst = False Then 'Mandantenbezogene Datenbegrenzung
    Select Case GlMaR 'Mandant neue(s) Rechnung/Rezept
    Case "J1": 'Standardmandant aus Optionsdialog
        mMaNr = GlMan(GlSMa, 2)
    Case "J2": 'Mandant aus Adresseneingabemaske
        If TeWer <> vbNullString Then
            mMaNr = CLng(TeWer)
            For AktZa = 1 To UBound(GlMan)
                If mMaNr = GlMan(AktZa, 2) Then
                    BeVor = True
                    Exit For
                End If
            Next AktZa
            If BeVor = True Then
                If CBool(GlMan(AktZa, 5)) = False Then 'Passiv / Aktiv
                    mMaNr = GlThe(AktZa, 0)
                Else
                    mMaNr = GlMan(GlSMa, 2)
                End If
            Else
                mMaNr = GlMan(GlSMa, 2)
            End If
        Else
            mMaNr = GlMan(GlSMa, 2)
        End If
    Case "J3": 'Mandant aus Mitarbeitereingabemaske
        mMaNr = GlMiA(GlSmI, 7)
    End Select
Else
    mMaNr = GlMiA(GlSmI, 7)
End If

MaIdx = SCmb(CmMan, mMaNr)
If MaIdx < 0 Then
    MaIdx = 0
End If
With CmMan
    .ListIndex = MaIdx
    .Enabled = True
End With

CmMit.ListIndex = GlSmI - 1

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

With ComFo
    .AutoComplete = False
    .DropDownItemCount = 20
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Select Case GlBut
Case RibTab_Startseite:
    Opti1.Enabled = False
    Opti2.Value = True
    Me.Caption = "Neues Rezept"
    TxKom.Text = GlRzV(ComFo.ListIndex + 1)
Case RibTab_Rezeptmodul:
    Me.Caption = "Neues Rezept"
    TxKom.Text = GlRzV(ComFo.ListIndex + 1)
Case RibTab_Belegmodul:
    Me.Caption = "Neuer Beleg"
    TxKom.Text = GlBlV(ComFo.ListIndex + 1)
End Select

With LiVw1
    .ColumnHeaders.Add 1, , "PZN", 1360
    .ColumnHeaders.Add 2, , "Heilmitteltext", 3240
    .ColumnHeaders.Add 3, , "IDD", 1
End With

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Opti1.BackColor = GlBak
Opti2.BackColor = GlBak
Chek1.BackColor = GlBak
Chek2.BackColor = GlBak

Set ImMan = Nothing
Set LiVw1 = Nothing
Set RpCo5 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FLiMa(ByVal CoIdx As Integer)
On Error Resume Next

Set LiVw1 = Me.lstView1
Set LiIts = LiVw1.ListItems

If LiIts.Count > 0 Then
    For Each LiItm In LiIts
        Select Case CoIdx
        Case 1: LiItm.Checked = False
        Case 2: LiItm.Checked = True
        End Select
    Next LiItm
End If

End Sub
Private Sub TWeit()
On Error GoTo SuErr

Dim NeuDa As Date
Dim VonDa As Date
Dim TmpNr As Long
Dim RowNr As Long
Dim BeDat As Date
Dim PatNa As String
Dim RzStr As String
Dim BelNr As Integer
Dim FrmNr As Integer
Dim Mld1, Tit1 As String
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set LiVw1 = Me.lstView1
Set TxDa1 = Me.txtDatu1
Set TxKom = Me.txtKomme
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set ComFo = Me.cmbFormu
Set Chek1 = Me.chkChek1
Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtPost
Set FTex3 = Me.txtBemer
Set FLis1 = Me.lstList1
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set PuBu2 = Me.btnZuruk
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set TxDaV = FM.txtDaVon
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns
Set RpSel = RpCo5.SelectedRows
Set LiIts = LiVw1.ListItems

Tit1 = "Beleg erst ausdrucken!"
Mld1 = "Für diesen Patienten existiert noch ein unausgedruckter Beleg. Diese sollte erst ausgedruckt werden."

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

TxDaV.Text = Format$(Date, "dd.mm.yyyy")

VonDa = Date

If Rahm1.Visible = True Then
    If Opti1.Value = True Then
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = True
        Rahm5.Visible = False
        DoEvents
        If S_RzOf(mPaNr) > 0 Then SPopu Tit1, Mld1, IC48_Information
    Else
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        FTex1.SetFocus
    End If
    PuBu2.Enabled = True
ElseIf Rahm2.Visible = True Then
    FSuda
ElseIf Rahm3.Visible = True Then
    mPaNr = FLis1.ItemData(FLis1.ListIndex)
    GlAdr = mPaNr
    GlTDa = vbNullString 'Wichtig für Textverarbeitung
    PatNa = FTex1.Text
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = True
    Rahm5.Visible = False
    DoEvents
    If S_RzOf(mPaNr) > 0 Then SPopu Tit1, Mld1, IC48_Information
    mMaNr = S_AdIdi(mPaNr, "IDP") 'Behandler wird neu ermittelt
ElseIf Rahm4.Visible = True Then
    BeDat = NeuDa
    BelNr = ComFo.ListIndex + 1
    Select Case GlBut
    Case RibTab_Startseite:
        Select Case BelNr
        Case 1: FrmNr = 1
        Case 2: FrmNr = 2
        Case 3: FrmNr = 3
        Case 4: FrmNr = 4
        Case 5: FrmNr = 5
        Case 6: FrmNr = 6
        Case 7: FrmNr = 7
        Case 8: FrmNr = 8
        Case 9: FrmNr = 13
        Case 10: FrmNr = 14
        Case 11: FrmNr = 16
        End Select
    Case RibTab_Rezeptmodul:
        Select Case BelNr
        Case 1: FrmNr = 1
        Case 2: FrmNr = 2
        Case 3: FrmNr = 3
        Case 4: FrmNr = 4
        Case 5: FrmNr = 5
        Case 6: FrmNr = 6
        Case 7: FrmNr = 7
        Case 8: FrmNr = 8
        Case 9: FrmNr = 13
        Case 10: FrmNr = 14
        Case 11: FrmNr = 16
        End Select
    Case RibTab_Belegmodul:
        Select Case BelNr
        Case 1: FrmNr = 9
        Case 2: FrmNr = 10
        Case 3: FrmNr = 11
        Case 4: FrmNr = 12
        Case 5: FrmNr = 15
        Case 6: FrmNr = 17
        Case 7: FrmNr = 18
        End Select
    End Select

    Select Case GlBut
    Case RibTab_Rezeptmodul:
        If LiIts.Count > 0 Then
            If GlKoZ = False Then 'Rezept kopieren
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = False
                Rahm4.Visible = False
                Rahm5.Visible = True
            Else
                With GlNeZ
                    .PatNr = mPaNr
                    .BeDat = BeDat
                    .BelNr = FrmNr
                    .DiaEi = KatLa
                    .TmDat = VonDa
                    .KoTex = TxKom.Text
                    .RzTex = vbNullString
                    .MaNum = CmMan.ItemData(CmMan.ListIndex)
                    .MitNr = CmMit.ItemData(CmMit.ListIndex)
                    If GlKoZ = True Then
                        If RpSel.Count > 0 Then
                            Set RpRow = RpSel(0)
                            If RpRow.GroupRow = False Then
                                RowNr = RpRow.Index
                                Set RpCol = RpCls.Find(Rzp_Rezepttext)
                                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                                    .RzTex = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                                End If
                            End If
                        End If
                    End If
                End With
                S_RzNe
                SUpRz
                If Opti2.Value = True Then FSuPa
                DoEvents
                Unload Me
            End If
        Else
            With GlNeZ
                .PatNr = mPaNr
                .BeDat = BeDat
                .BelNr = FrmNr
                .DiaEi = KatLa
                .TmDat = VonDa
                .KoTex = TxKom.Text
                .RzTex = vbNullString
                .MaNum = CmMan.ItemData(CmMan.ListIndex)
                .MitNr = CmMit.ItemData(CmMit.ListIndex)
                If GlKoZ = True Then
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        If RpRow.GroupRow = False Then
                            RowNr = RpRow.Index
                            Set RpCol = RpCls.Find(Rzp_Rezepttext)
                            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                                .RzTex = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            End If
                        End If
                    End If
                End If
            End With
            S_RzNe
            SUpRz
            If Opti2.Value = True Then FSuPa
            DoEvents
            Unload Me
        End If
    Case RibTab_Belegmodul:
        With GlNeZ
            .PatNr = mPaNr
            .BeDat = BeDat
            .BelNr = FrmNr
            .DiaEi = KatLa
            .TmDat = VonDa
            .KoTex = TxKom.Text
            .RzTex = vbNullString
            .MaNum = CmMan.ItemData(CmMan.ListIndex)
            .MitNr = CmMit.ItemData(CmMit.ListIndex)
            If GlKoZ = True Then 'Rezept kopieren
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        RowNr = RpRow.Index
                        Set RpCol = RpCls.Find(Rzp_Rezepttext)
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            .RzTex = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                        End If
                    End If
                End If
            End If
        End With
        S_RzNe
        SUpRz
        If Opti2.Value = True Then FSuPa
        DoEvents
        Unload Me
    Case RibTab_Startseite:
        If LiIts.Count > 0 Then
            If GlKoZ = False Then
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = False
                Rahm4.Visible = False
                Rahm5.Visible = True
            Else
                With GlNeZ
                    .PatNr = mPaNr
                    .BeDat = BeDat
                    .BelNr = FrmNr
                    .DiaEi = KatLa
                    .TmDat = VonDa
                    .KoTex = TxKom.Text
                    .RzTex = vbNullString
                    .MaNum = CmMan.ItemData(CmMan.ListIndex)
                    .MitNr = CmMit.ItemData(CmMit.ListIndex)
                    If GlKoZ = True Then 'Rezept kopieren
                        If RpSel.Count > 0 Then
                            Set RpRow = RpSel(0)
                            If RpRow.GroupRow = False Then
                                RowNr = RpRow.Index
                                Set RpCol = RpCls.Find(Rzp_Rezepttext)
                                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                                    .RzTex = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                                End If
                            End If
                        End If
                    End If
                End With
                S_RzNe
                SUpRz
                If Opti2.Value = True Then FSuPa
                DoEvents
                Unload Me
            End If
        Else
            With GlNeZ
                .PatNr = mPaNr
                .BeDat = BeDat
                .BelNr = FrmNr
                .DiaEi = KatLa
                .TmDat = VonDa
                .KoTex = TxKom.Text
                .RzTex = vbNullString
                .MaNum = CmMan.ItemData(CmMan.ListIndex)
                .MitNr = CmMit.ItemData(CmMit.ListIndex)
                If GlKoZ = True Then
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        If RpRow.GroupRow = False Then
                            RowNr = RpRow.Index
                            Set RpCol = RpCls.Find(Rzp_Rezepttext)
                            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                                .RzTex = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            End If
                        End If
                    End If
                End If
            End With
            S_RzNe
            SUpRz
            If Opti2.Value = True Then FSuPa
            DoEvents
            Unload Me
        End If
        DoEvents
        SRzZe 0
    End Select
ElseIf Rahm5.Visible = True Then
    BeDat = NeuDa
    BelNr = ComFo.ListIndex + 1
    Select Case GlBut
    Case RibTab_Startseite:
        Select Case BelNr
        Case 1: FrmNr = 1
        Case 2: FrmNr = 2
        Case 3: FrmNr = 3
        Case 4: FrmNr = 4
        Case 5: FrmNr = 5
        Case 6: FrmNr = 6
        Case 7: FrmNr = 7
        Case 8: FrmNr = 8
        Case 9: FrmNr = 13
        Case 10: FrmNr = 14
        Case 11: FrmNr = 16
        End Select
    Case RibTab_Rezeptmodul:
        Select Case BelNr
        Case 1: FrmNr = 1
        Case 2: FrmNr = 2
        Case 3: FrmNr = 3
        Case 4: FrmNr = 4
        Case 5: FrmNr = 5
        Case 6: FrmNr = 6
        Case 7: FrmNr = 7
        Case 8: FrmNr = 8
        Case 9: FrmNr = 13
        Case 10: FrmNr = 14
        Case 11: FrmNr = 16
        End Select
    Case RibTab_Belegmodul:
        Select Case BelNr
        Case 1: FrmNr = 9
        Case 2: FrmNr = 10
        Case 3: FrmNr = 11
        Case 4: FrmNr = 12
        Case 5: FrmNr = 15
        Case 6: FrmNr = 17
        Case 7: FrmNr = 18
        End Select
    End Select
    
    If LiIts.Count > 0 Then
        For Each LiItm In LiIts
            If LiItm.Checked = True Then
                If GlPzn = True Then
                    RzStr = RzStr & LiItm.Text & Chr$(32) & LiItm.SubItems(1) & vbCrLf
                Else
                    RzStr = RzStr & LiItm.SubItems(1) & vbCrLf
                End If
            End If
        Next LiItm
    End If

    With GlNeZ
        .PatNr = mPaNr
        .BeDat = BeDat
        .BelNr = FrmNr
        .DiaEi = KatLa
        .KoTex = TxKom.Text
        .RzTex = RzStr
        .MaNum = CmMan.ItemData(CmMan.ListIndex)
        .MitNr = CmMit.ItemData(CmMit.ListIndex)
    End With
    S_RzNe
    SUpRz
    If Opti2.Value = True Then
        FSuPa
    End If
    DoEvents
    Unload Me
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo5 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TWeit " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    FKale
End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub TRes()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtPost
Set FTex3 = Me.txtBemer

FTex1.Text = vbNullString
FTex2.Text = vbNullString
FTex3.Text = vbNullString

End Sub
Private Sub FZuru()
On Error Resume Next

Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set PuBu2 = Me.btnZuruk

If Rahm2.Visible = True Then
    Rahm5.Visible = False
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = True
    PuBu2.Enabled = False
ElseIf Rahm3.Visible = True Then
    Rahm5.Visible = False
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = True
    Rahm1.Visible = False
ElseIf Rahm4.Visible = True Then
    If Opti1.Value = True Then
        Rahm5.Visible = False
        Rahm4.Visible = False
        Rahm3.Visible = False
        Rahm2.Visible = False
        Rahm1.Visible = True
        PuBu2.Enabled = False
    Else
        Rahm5.Visible = False
        Rahm4.Visible = False
        Rahm3.Visible = True
        Rahm2.Visible = False
        Rahm1.Visible = False
    End If
ElseIf Rahm5.Visible = True Then
    Rahm5.Visible = False
    Rahm4.Visible = True
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = False
End If

End Sub
Private Sub FSuda()
On Error GoTo SuErr

Dim Mld1, Tit1 As String

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtPost
Set FTex3 = Me.txtBemer
Set FLis1 = Me.lstList1
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder

If FTex1.Text <> vbNullString Then
    S_AdFin FTex1.Text, 1, 2
ElseIf FTex2.Text <> vbNullString Then
    S_AdFin FTex2.Text, 2, 2
ElseIf FTex3.Text <> vbNullString Then
    S_AdFin FTex3.Text, 3, 2
End If

If FLis1.ListCount > 0 Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
    Rahm5.Visible = False
    FLis1.SetFocus
    FLis1.Selected(0) = True
Else
    Mld1 = "Das von Ihnen eingegebene Suchkriterium brachte leider keine Suchergebnisse"
    Tit1 = "Adressuche"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuda " & Err.Number
Resume Next

End Sub
Private Sub FSuPa()
On Error GoTo SuErr

Set FM = frmMain

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

GlAkt = True

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

S_List
S_RzLa

clFen.FenDsk 3
Screen.MousePointer = vbNormal

GlAkt = False

Set clFen = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuPa " & Err.Number
Resume Next

End Sub
Private Sub FRzMa()
On Error Resume Next

Set Chek2 = Me.chkRzMar

If Chek2.Value = xtpChecked Then
    GlRzM = True
    IniSetVal "System", "RezMar", -1
    FLiMa 2
Else
    GlRzM = False
    IniSetVal "System", "RezMar", 0
    FLiMa 1
End If

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

Private Sub btnWeite_Click()
    TWeit
End Sub

Private Sub btnZuruk_Click()
    FZuru
End Sub
Private Sub chkRzMar_Click()
    FRzMa
End Sub

Private Sub cmbFormu_Click()
On Error Resume Next

Set TxKom = Me.txtKomme
Set ComFo = Me.cmbFormu

Select Case GlBut
Case RibTab_Rezeptmodul: IniSetVal "System", "RzForm", "O" & Format$(ComFo.ListIndex + 1, "00")
Case RibTab_Belegmodul: IniSetVal "System", "BlForm", "O" & Format$(ComFo.ListIndex + 1, "00")
End Select

Select Case GlBut
Case RibTab_Rezeptmodul:
    TxKom.Text = GlRzV(ComFo.ListIndex + 1)
Case RibTab_Belegmodul:
    TxKom.Text = GlBlV(ComFo.ListIndex + 1)
End Select

End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF11: Unload Me
    End Select
End Sub
Private Sub Form_Load()
On Error Resume Next

FKonf
S_RzKn
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmRzNeu = Nothing
End Sub

Private Sub txtBemer_GotFocus()
    TRes
End Sub
Private Sub txtBemer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub

Private Sub txtKomme_GotFocus()
    Me.txtKomme.SelStart = 0
    Me.txtKomme.SelLength = Len(Me.txtKomme.Text)
End Sub

Private Sub txtKurz_GotFocus()
    TRes
End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtPost_GotFocus()
    TRes
End Sub
Private Sub txtPost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub

Private Sub updCont1_DownClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub
Private Sub FKaLa()
On Error Resume Next

Set Chek1 = Me.chkChek1

If Chek1.Value = xtpChecked Then
    KatLa = True
    IniSetVal "System", "KatAut", -1
Else
    KatLa = False
    IniSetVal "System", "KatAut", 0
End If

End Sub
Private Sub chkChek1_Click()
    FKaLa
End Sub
Private Sub lstView1_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
    FLiMa ColumnHeader.Index
End Sub
