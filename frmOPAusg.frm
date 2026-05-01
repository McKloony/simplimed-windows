VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmOPAusg 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Offene Posten Ausgleichen"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   13
      Top             =   6300
      Width           =   6900
      _Version        =   1048579
      _ExtentX        =   12171
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4900
         TabIndex        =   16
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
         Left            =   3500
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
         Left            =   2200
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
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   1500
      Left            =   550
      TabIndex        =   3
      Top             =   4640
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   2646
      _StockProps     =   79
      Caption         =   "Mandant und Mitarbeiter"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   1900
         TabIndex        =   11
         Top             =   360
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   1900
         TabIndex        =   12
         Top             =   910
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   255
         Left            =   500
         TabIndex        =   20
         Top             =   950
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mitarbeiter :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   255
         Left            =   500
         TabIndex        =   19
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mandant :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   495
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8000
      Visible         =   0   'False
      Width           =   615
      _Version        =   1048579
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   64
      Show3DBorder    =   2
      VisualTheme     =   0
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   2060
      Left            =   550
      TabIndex        =   2
      Top             =   2400
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   3634
      _StockProps     =   79
      Caption         =   "Einnahmebuchung"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cmbKonto 
         Height          =   315
         Left            =   1900
         TabIndex        =   9
         Top             =   910
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbGegen 
         Height          =   315
         Left            =   1900
         TabIndex        =   8
         Top             =   360
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.ComboBox cmbStKto 
         Height          =   315
         Left            =   1900
         TabIndex        =   10
         Top             =   1470
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   255
         Left            =   500
         TabIndex        =   25
         Top             =   1500
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Steuerkonto :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   255
         Left            =   500
         TabIndex        =   22
         Top             =   940
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Erlöskonto :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   255
         Left            =   500
         TabIndex        =   21
         Top             =   390
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Geldkonto :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   1500
      Left            =   550
      TabIndex        =   1
      Top             =   721
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   2646
      _StockProps     =   79
      Caption         =   "Einzahlung"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   3420
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
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
         Left            =   3710
         TabIndex        =   6
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   360
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBezBe 
         Height          =   350
         Left            =   1900
         TabIndex        =   7
         Top             =   910
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   1900
         TabIndex        =   4
         Top             =   360
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   255
         Left            =   500
         TabIndex        =   24
         Top             =   950
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Bezahlbetrag :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   255
         Left            =   500
         TabIndex        =   23
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Bezahlt am :"
         Alignment       =   1
         Transparent     =   -1  'True
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
      Top             =   8000
      Width           =   80
   End
   Begin VB.Label lblLab01 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte geben Sie den bezahlten Betrag und das Datum der Einzahlung ein und wählen Sie ggf. ein anderes Erlös- und Geldkonto."
      Height          =   420
      Left            =   800
      TabIndex        =   17
      Top             =   100
      Width           =   5000
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   650
      Left            =   0
      Top             =   0
      Width           =   6810
   End
End
Attribute VB_Name = "frmOPAusg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private CmKto As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmStu As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private TxBez As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private Lbl02 As XtremeSuiteControls.Label
Private Lbl03 As XtremeSuiteControls.Label
Private Lbl04 As XtremeSuiteControls.Label
Private Lbl05 As XtremeSuiteControls.Label
Private Lbl06 As XtremeSuiteControls.Label
Private Lbl07 As XtremeSuiteControls.Label
Private PuBu1 As XtremeSuiteControls.PushButton
Private UpDo1 As XtremeSuiteControls.UpDown
Private ImMan As XtremeCommandBars.ImageManager
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn
Private MoKal As XtremeCalendarControl.DatePicker
Private RpSel As XtremeReportControl.ReportSelectedRows

Private FoLad As Boolean
Private Const KEYEVENTF_KEYUP = &H2

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    TxDa1.Text = NeuDa
    GlAuD = NeuDa
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

Dim IdEin As Long
Dim AnzPo As Long
Dim ManNr As Long
Dim MitNr As Long
Dim TmpNr As Long
Dim StaGe As Long
Dim StaKt As Long
Dim BuBet As Single
Dim GeBet As Single
Dim OpBe1 As Single
Dim OpBe2 As Single
Dim OpBe3 As Single
Dim OpBe4 As Single
Dim OpBe5 As Single
Dim GeSum As Single
Dim ReBez As String
Dim StaRa As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim IdStK As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set TxBez = Me.txtBezBe
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen
Set CmStu = Me.cmbStKto
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set Lbl02 = Me.lblLab02
Set Lbl03 = Me.lblLab03
Set Lbl04 = Me.lblLab04
Set Lbl05 = Me.lblLab05
Set Lbl06 = Me.lblLab06
Set Lbl07 = Me.lblLab07
Set PuBu1 = Me.btnDatu1
Set UpDo1 = Me.updCont1
Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1
Set RpCo1 = FM.repCont1
Set ImMan = FM.imgManag
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

AnzPo = RpSel.Count

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

If AnzPo = 0 Then

    TxBez.Text = Format$(0, GlWa1)
    TxBez.Enabled = False
    CmGeg.ListIndex = 0
    CmKto.ListIndex = 0
    CmMan.ListIndex = GlSMa - 1
    CmMit.ListIndex = GlSmI - 1
    
ElseIf AnzPo > 1 Then

    If GlBut = RibTab_HomeBanki Then

        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Ban_GeBetrag)
                GeSum = GeSum + CSng(RpRow.Record(RpCol.ItemIndex).Value)
            End If
        Next RpRow
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Ban_IDP)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                ManNr = 0
            End If
            Set RpCol = RpCls.Find(Ban_IDM)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                MitNr = GlMiA(GlSmI, 2)
            End If
            Set RpCol = RpCls.Find(Ban_IDB)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                StaGe = RpRow.Record(RpCol.ItemIndex).Value
            Else
                StaGe = 0
            End If
        End If
        TxDa1.Enabled = False
        PuBu1.Enabled = False
        UpDo1.Enabled = False
        CmGeg.Enabled = False
        CmStu.Enabled = False
        
    Else
    
        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(OPo_OffBetrag)
                GeSum = GeSum + CSng(RpRow.Record(RpCol.ItemIndex).Value)
            End If
        Next RpRow
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(OPo_IDP)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                ManNr = 0
            End If
            Set RpCol = RpCls.Find(OPo_IDT)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                MitNr = GlMiA(GlSmI, 2) 'Standardmitarbeiter
            End If
        End If
    End If

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
    
    TxBez.Text = Format$(GeSum, GlWa1)
    CmMan.Enabled = False
    CmMit.Enabled = False
    TxBez.Enabled = False
    
    If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
        CmGeg.Enabled = False
        CmKto.Enabled = False
        CmStu.Enabled = False
    End If
    
Else

    If GlBut = RibTab_HomeBanki Then
    
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Ban_Bezahlt)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                ReBez = RpRow.Record(RpCol.ItemIndex).Value
            Else
                ReBez = vbNullString
            End If
            Set RpCol = RpCls.Find(Ban_KoBetrag)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                BuBet = RpRow.Record(RpCol.ItemIndex).Value
            Else
                BuBet = 0
            End If
            Set RpCol = RpCls.Find(Ban_GeBetrag)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                GeBet = RpRow.Record(RpCol.ItemIndex).Value
            Else
                GeBet = 0
            End If
            Set RpCol = RpCls.Find(Ban_OPBetrag1)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                OpBe1 = RpRow.Record(RpCol.ItemIndex).Value
            Else
                OpBe1 = 0
            End If
            Set RpCol = RpCls.Find(Ban_OPBetrag2)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                OpBe2 = RpRow.Record(RpCol.ItemIndex).Value
            Else
                OpBe2 = 0
            End If
            Set RpCol = RpCls.Find(Ban_OPBetrag3)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                OpBe3 = RpRow.Record(RpCol.ItemIndex).Value
            Else
                OpBe3 = 0
            End If
            Set RpCol = RpCls.Find(Ban_OPBetrag4)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                OpBe4 = RpRow.Record(RpCol.ItemIndex).Value
            Else
                OpBe4 = 0
            End If
            Set RpCol = RpCls.Find(Ban_OPBetrag5)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                OpBe5 = RpRow.Record(RpCol.ItemIndex).Value
            Else
                OpBe5 = 0
            End If
            Set RpCol = RpCls.Find(Ban_IDB)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                StaGe = RpRow.Record(RpCol.ItemIndex).Value
            Else
                StaGe = 0
            End If
            Set RpCol = RpCls.Find(Ban_IDP)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                ManNr = 0
            End If
            Set RpCol = RpCls.Find(Ban_IDM)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                MitNr = GlMiA(GlSmI, 2)
            End If

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
            
            If ReBez = "Nein" Then
                If GeBet > 0 Then
                    If BuBet > 0 Then
                        TxBez.Text = Format$(BuBet, GlWa1)
                    Else
                        TxBez.Text = Format$(0, GlWa1)
                    End If
                Else
                    TxBez.Text = Format$(0, GlWa1)
                End If
            Else
                TxBez.Text = Format$(0, GlWa1)
            End If
        End If
        TxBez.Enabled = False
        TxDa1.Enabled = False
        PuBu1.Enabled = False
        UpDo1.Enabled = False
        CmGeg.Enabled = False
        CmStu.Enabled = False
        
    Else
    
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then

            Set RpCol = RpCls.Find(OPo_OffBetrag)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                OpBe1 = RpRow.Record(RpCol.ItemIndex).Value
            Else
                OpBe1 = 0
            End If
            
            Set RpCol = RpCls.Find(OPo_IDP)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                ManNr = GlMan(GlSMa, 2) 'Standardmandant
            End If
            
            Set RpCol = RpCls.Find(OPo_IDT)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                MitNr = GlMiA(GlSmI, 2) 'Standardmitarbeiter
            End If
            
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
            
            CmStu.Enabled = GlSpB 'Umsatzsteuer Splittbuchung
            
            TxBez.Text = Format$(OpBe1, GlWa1)
        End If
    End If
End If

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    S_KoMa 2, ManNr, StaGe
Else
    StaKt = SCmb(CmKto, GlSE2) 'Standarderlöskonto Bankkonto
    If StaKt >= 0 Then
        CmKto.ListIndex = StaKt
    Else
        CmKto.ListIndex = 0
    End If
    If StaGe = 0 Then
        StaGe = SCmb(CmGeg, GlGkB) 'Standardgeldkonto Bankkonto
        If StaGe >= 0 Then
            If CmGeg.ListCount > 0 Then
                CmGeg.ListIndex = StaGe
            End If
        Else
            CmGeg.ListIndex = 0
        End If
    Else
        StaGe = SCmb(CmGeg, StaGe)
        If CmGeg.ListCount > 0 Then
            CmGeg.ListIndex = StaGe
        End If
    End If
End If

If CmKto.ListIndex < 0 Then
    CmKto.ListIndex = 0
End If

If CmGeg.ListIndex < 0 Then
    CmGeg.ListIndex = 0
End If

TxDa1.SetMask "00.00.0000", "__.__.____"

If GlBut = RibTab_HomeBanki Then
    If AnzPo > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Ban_Datum)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                    TxDa1.Text = RpRow.Record(RpCol.ItemIndex).Value
                Else
                    TxDa1.Text = Date
                End If
            Else
                TxDa1.Text = Date
            End If
        Else
            TxDa1.Text = Date
        End If
    Else
        TxDa1.Text = Date
    End If
Else
    If IsDate(GlAuD) = True Then
        If Not Year(GlAuD) < Year(Date) - 4 Then
            TxDa1.Text = GlAuD
        Else
            TxDa1.Text = Date
        End If
    Else
        TxDa1.Text = Date
    End If
End If

If GlBuc = True Then 'Einfache Buchhaltung verwenden
    Lbl04.Caption = "Geldkonto :"
    Lbl05.Caption = "Sachkonto :"
Else
    Lbl04.Caption = "Sollkonto :"
    Lbl05.Caption = "Habenkonto :"
End If

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Lbl02.BackColor = GlBak
Lbl03.BackColor = GlBak
Lbl04.BackColor = GlBak
Lbl05.BackColor = GlBak
Lbl06.BackColor = GlBak
Lbl07.BackColor = GlBak

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) = True Then
    NeuDa = CDate(TxDa1.Text)
    TxDa1.Text = NeuDa
    With MoKal
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
    If NeuDa > Date Then
        SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
    End If
End If

Set MoKal = Nothing

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
Set MoKal = Me.dtpDatu1
Set Rahm1 = Me.frmRahm1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
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
Private Sub FMand()
On Error GoTo LdErr

Dim ManNr As Long

Set CmMan = Me.cmbManda

ManNr = CmMan.ItemData(CmMan.ListIndex)

S_KoMa 2, ManNr

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "cmbManda " & Err.Number
Resume Next

End Sub

Private Sub FWeit()
On Error GoTo SuErr

Dim NeuDa As Date
Dim ManNr As Long
Dim MitNr As Long
Dim AnzPo As Long
Dim KtoNr As Long
Dim GeKNr As Long
Dim StKNr As Long
Dim RowNr As Long
Dim KtoID As Long
Dim StKID As Long
Dim BetBz As Double
Dim KtoBe As String
Dim GeKBe As String
Dim StKBe As String
Dim GeKId As Integer
Dim Mld1, Tit1 As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo4 = FM.repCont4
Set RpCo3 = FM.repCont3
Set TxBez = Me.txtBezBe
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen
Set CmStu = Me.cmbStKto
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set TxDa1 = Me.txtDatu1
Set PuBu1 = Me.btnWeite
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

AnzPo = RpSel.Count
KtoID = CmKto.ItemData(CmKto.ListIndex)
GeKId = CmGeg.ItemData(CmGeg.ListIndex)
StKID = CmStu.ItemData(CmStu.ListIndex)

If GlKnF = True Then 'Sachkontenformatierung sechsstellig
    KtoNr = Left$(CmKto.Text, 6)
    GeKNr = Left$(CmGeg.Text, 6)
    If CmStu.Enabled = True Then
        StKNr = Left$(CmStu.Text, 6)
        StKBe = Mid$(CmStu.Text, 8, Len(CmStu.Text) - 7)
    End If
    KtoBe = Mid$(CmKto.Text, 8, Len(CmKto.Text) - 7)
    GeKBe = Mid$(CmGeg.Text, 8, Len(CmGeg.Text) - 7)
Else
    KtoNr = Left$(CmKto.Text, 4)
    GeKNr = Left$(CmGeg.Text, 4)
    If CmStu.Enabled = True Then
        StKNr = Left$(CmStu.Text, 4)
        StKBe = Mid$(CmStu.Text, 6, Len(CmStu.Text) - 5)
    End If
    KtoBe = Mid$(CmKto.Text, 6, Len(CmKto.Text) - 5)
    GeKBe = Mid$(CmGeg.Text, 6, Len(CmGeg.Text) - 5)
End If

PuBu1.Enabled = False
If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    Exit Sub
End If

BetBz = CDbl(TxBez.Text)
BetBz = Round(BetBz, 2)

If AnzPo > 0 Then
    If AnzPo = 1 Then
        ManNr = CmMan.ItemData(CmMan.ListIndex)
        MitNr = CmMit.ItemData(CmMit.ListIndex)
    End If

    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    
    If AnzPo = 1 Then
        If GlBut = RibTab_HomeBanki Then
            Set RpCol = RpCls.Find(Ban_Ausgabe)
            If LCase(RpRow.Record(RpCol.ItemIndex).Value) = "ja" Then
                Mld1 = "Der markierte Umsatz ist eine Ausgabe"
                Tit1 = "Falscher Eintrag markiert"
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
                Unload Me
                Exit Sub
            End If
            Set RpCol = RpCls.Find(Ban_Selekt)
            If RpRow.Record(RpCol.ItemIndex).Value = "Nein" Then
                Mld1 = "Dem markierten Umsatz wurde noch kein offener Posten zugeordnet"
                Tit1 = "Kein Posten zugeordnet"
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
                Unload Me
                Exit Sub
            End If
            Set RpCol = RpCls.Find(Ban_Bezahlt)
            If RpRow.Record(RpCol.ItemIndex).Value = "Ja" Then
                Mld1 = "Für den markierten Umsatz wurde bereits ein Erlös gebucht"
                Tit1 = "Posten bereits ausgeglichen"
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
                Unload Me
                Exit Sub
            End If
        Else
            Set RpCol = RpCls.Find(OPo_Selekt)
            If CBool(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                SPopu "Posten ausgleichen", "Der offene Posten wurde bereits ausgeglichen!", IC48_Information
                Unload Me
                Exit Sub
            End If
        End If
        If TxBez.Text = vbNullString Then
            Mld1 = "Bitte geben Sie erst den bezahlten Betrag ein"
            Tit1 = "Betrag fehlt"
            WindowMess Mld1, Dial2, Tit1, FM.hwnd
            Exit Sub
        ElseIf BetBz = 0 Then
            Mld1 = "Es existiert kein offener Betrag mehr für diesen Posten"
            Tit1 = "Bereits ausgeglichen"
            WindowMess Mld1, Dial2, Tit1, FM.hwnd
            Exit Sub
        ElseIf BetBz <= 0 Then
            Mld1 = "Der bezahlte Betrag ist negativ und somit ungültig"
            Tit1 = "Falscher Betrag"
            WindowMess Mld1, Dial2, Tit1, FM.hwnd
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass

    If GlBut = RibTab_HomeBanki Then
        Unload Me
        DoEvents
        S_KoAu ManNr, MitNr, KtoID, KtoNr, KtoBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe
        DoEvents
        SUpBa RowNr
        DoEvents
    ElseIf GlBut = RibTab_Mahnwesen Then
        Unload Me
        DoEvents
        S_OPAu BetBz, NeuDa, ManNr, MitNr, KtoID, KtoNr, KtoBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe
        DoEvents
        SUpOp RowNr
        DoEvents
    ElseIf GlBut = RibTab_Ter_Akont Then
        Unload Me
        DoEvents
        S_TeAu NeuDa, ManNr, MitNr, KtoID, KtoNr, KtoBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe
        DoEvents
        SUpTe RowNr
        DoEvents
    End If

    If GlBut = RibTab_HomeBanki Then
        Set RpSel = RpCo4.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpRe RowNr, True
        Else
            SUpRe
        End If
    ElseIf GlBut = RibTab_Mahnwesen Then
        Set RpSel = RpCo4.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpRe RowNr, True
        Else
            SUpRe
        End If
    End If
    
    Screen.MousePointer = vbNormal
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
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

Select Case GlBut:
Case RibTab_Mahnwesen:
    TeTit = IniGetOpt("Hilfe", 50651)
    TeMai = IniGetOpt("Hilfe", 50652)
    TeInh = IniGetOpt("Hilfe", 50653)
    TeFus = IniGetOpt("Hilfe", 50654)
Case RibTab_HomeBanki:
    TeTit = IniGetOpt("Hilfe", 50661)
    TeMai = IniGetOpt("Hilfe", 50662)
    TeInh = IniGetOpt("Hilfe", 50663)
    TeFus = IniGetOpt("Hilfe", 50664)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeite_Click()
    FWeit
End Sub
Private Sub cmbGegen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbKonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbManda_Click()
    If FoLad = False Then
        FMand
    End If
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

FoLad = True
FInit
FoLad = False
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmOPAusg = Nothing
End Sub
Private Sub txtBezBe_GotFocus()
    Me.txtBezBe.SelStart = 0
    Me.txtBezBe.SelLength = Len(Me.txtBezBe.Text)
End Sub
Private Sub txtBezBe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBezBe_LostFocus()
On Error Resume Next

Dim Betra As Double

If Me.txtBezBe.Text <> vbNullString Then
    If IsNumeric(Me.txtBezBe.Text) = True Then
        Betra = CDbl(Me.txtBezBe.Text)
        If Betra < 0 Then
            Betra = Betra * (-1)
        End If
        Me.txtBezBe.Text = Format$(Betra, GlWa1)
    End If
End If

End Sub
Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub

Private Sub updCont1_DownClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

If IsDate(TxDa1.Text) = True Then
    AltDa = TxDa1.Text
Else
    AltDa = Date
End If

TxDa1.Text = DateAdd("d", -1, AltDa)

If IsDate(TxDa1.Text) = True Then
    GlAuD = CDate(TxDa1.Text)
Else
    TxDa1.Text = AltDa
End If

End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

If IsDate(TxDa1.Text) = True Then
    AltDa = TxDa1.Text
Else
    AltDa = Date
End If

TxDa1.Text = DateAdd("d", 1, AltDa)

If IsDate(TxDa1.Text) = True Then
    GlAuD = CDate(TxDa1.Text)
Else
    TxDa1.Text = AltDa
End If

End Sub
