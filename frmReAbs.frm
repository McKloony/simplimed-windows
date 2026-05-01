VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReAbs 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnungen Verriegeln"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   17
      Top             =   4400
      Width           =   7800
      _Version        =   1048579
      _ExtentX        =   13758
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   5800
         TabIndex        =   20
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
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   4400
         TabIndex        =   19
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
         Left            =   3100
         TabIndex        =   18
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
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6200
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7700
      _Version        =   1048579
      _ExtentX        =   13582
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   2520
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
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
      Begin XtremeSuiteControls.CheckBox chkGePru 
         Height          =   225
         Left            =   4600
         TabIndex        =   16
         Top             =   2900
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Gebührenziffernanpassung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkDoTer 
         Height          =   225
         Left            =   4600
         TabIndex        =   15
         Top             =   2400
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Doppelte Termine prüfen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkRePru 
         Height          =   225
         Left            =   4600
         TabIndex        =   14
         Top             =   1900
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungsprüfungslauf"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkDatum 
         Height          =   225
         Left            =   1010
         TabIndex        =   2
         Top             =   900
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungsdatum anpassen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkReAbs 
         Height          =   225
         Left            =   4600
         TabIndex        =   12
         Top             =   900
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungen verriegeln"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   2800
         TabIndex        =   5
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
         Left            =   1000
         TabIndex        =   3
         Top             =   1200
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   315
         Left            =   1000
         TabIndex        =   11
         Top             =   3650
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.CheckBox chkManda 
         Height          =   225
         Left            =   1010
         TabIndex        =   10
         Top             =   3350
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandantenzuordnung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbKonto 
         Height          =   315
         Left            =   1000
         TabIndex        =   9
         Top             =   2800
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbGegen 
         Height          =   315
         Left            =   1000
         TabIndex        =   7
         Top             =   2000
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.CheckBox chkSaKon 
         Height          =   225
         Left            =   1010
         TabIndex        =   8
         Top             =   2500
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Erlöskonto"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkGegen 
         Height          =   225
         Left            =   1010
         TabIndex        =   6
         Top             =   1700
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Geldkonto"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkOpGen 
         Height          =   220
         Left            =   4600
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1400
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Posten / Buchung generieren"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmReAbs.frx":0000
         Height          =   580
         Left            =   300
         TabIndex        =   21
         Top             =   100
         Width           =   7100
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6400
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
Attribute VB_Name = "frmReAbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private FTex1 As XtremeSuiteControls.FlatEdit
Private CheRe As XtremeSuiteControls.CheckBox
Private ChePr As XtremeSuiteControls.CheckBox
Private ChDat As XtremeSuiteControls.CheckBox
Private ChGeb As XtremeSuiteControls.CheckBox
Private ChDop As XtremeSuiteControls.CheckBox
Private ChMan As XtremeSuiteControls.CheckBox
Private ChKon As XtremeSuiteControls.CheckBox
Private ChGeg As XtremeSuiteControls.CheckBox
Private ChOpo As XtremeSuiteControls.CheckBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmKto As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private MoKal As XtremeCalendarControl.DatePicker
Private ImMan As XtremeCommandBars.ImageManager
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private UpCo1 As XtremeSuiteControls.UpDown

Private clFen As clsFenster

Private FoLad As Boolean
Private Sub FAbs()
On Error GoTo OpErr
'Erzeugt einen offenen Posten

Dim NeuDa As Date
Dim AltDa As Date
Dim RowNr As Long
Dim KrRow As Long
Dim ManNr As Long
Dim EiKto As Long
Dim DatAn As Boolean
Dim GeAnp As Boolean
Dim OpBuA As Boolean
Dim OpPru As Boolean
Dim ReAbs As Boolean
Dim RePru As Boolean
Dim PrDop As Boolean
Dim RetWe As Boolean
Dim GeKto As Integer
Dim AnzPo As Integer
Dim AbgRe As Integer
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo6 = FM.repCont6
Set TxDa1 = Me.txtDatu1
Set CheRe = Me.chkReAbs 'Rechnung verriegeln
Set ChePr = Me.chkRePru
Set ChDat = Me.chkDatum 'Rechnungsdatum anpassen
Set ChGeb = Me.chkGePru
Set ChDop = Me.chkDoTer
Set ChMan = Me.chkManda
Set ChKon = Me.chkSaKon
Set ChOpo = Me.chkOpGen
Set ChGeg = Me.chkGegen
Set CmMan = Me.cmbBehan
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen

If IsDate(TxDa1.Text) = True Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

If CmMan.ListIndex > -1 Then
    If ChMan.Value = xtpChecked Then
        ManNr = CmMan.ItemData(CmMan.ListIndex)
    Else
        ManNr = 0
    End If
End If

If CmKto.ListIndex > -1 Then
    If ChKon.Value = xtpChecked Then
        EiKto = CmKto.ItemData(CmKto.ListIndex)
    Else
        EiKto = 0
    End If
End If

If CmGeg.ListIndex > -1 Then
    If ChGeg.Value = xtpChecked Then
        GeKto = CmGeg.ItemData(CmGeg.ListIndex)
    Else
        GeKto = 0
    End If
End If

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select
AnzPo = RpSel.Count

If CheRe.Value = xtpChecked Then
    ReAbs = True
Else
    ReAbs = False
End If

If ChePr.Value = 1 Then
    RePru = True
End If

If ChDat.Value = xtpChecked Then
    DatAn = True
Else
    DatAn = False
End If

If ChGeb.Value = xtpChecked Then
    GeAnp = True
Else
    GeAnp = False
End If

If ChOpo.Value = xtpChecked Then
    OpBuA = True
Else
    OpBuA = False
End If

OpPru = OpBuA 'WICHTIG!

If ChDop.Value = 1 Then 'doppelte Termine prüfen
    PrDop = True
End If

Unload Me
DoEvents

If AnzPo > 0 Then
    Set RpRow = RpSel(0)
    Set RpCol = RpCls.Find(Rec_Datum)
    AltDa = RpRow.Record(RpCol.ItemIndex).Value

    If DatAn = True Then 'Rechnungsnummer und Rechnungsdatum generieren
        RetWe = S_ReAn(False, ReAbs, NeuDa)
    Else
        RetWe = S_ReAn(True, ReAbs)
    End If
    If RetWe = True Then
        Exit Sub
    End If
    DoEvents

    If PrDop = True Then 'doppelte Termine prüfen
        S_ReDop
    End If
    DoEvents
    
    If RePru = True Then 'Rechnungspreisprüfung
        S_RePru GeAnp
    End If
    DoEvents

    If ReAbs = True Then
        AbgRe = S_OPAn(NeuDa, DatAn, ManNr, EiKto, GeKto, OpBuA, OpPru)
    End If
    DoEvents

    Select Case GlBut
    Case RibTab_Abrechnung:
            If AnzPo > 1 Then
                SUpAb
                SUpRe , True
            Else
                Set RpSel = RpCo3.SelectedRows
                If RpSel.Count > 0 Then
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
                Else
                    SUpAb
                End If
                Set RpSel = RpCo4.SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    RowNr = RpRow.Index
                    SUpRe RowNr
                Else
                    SUpRe
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
                Else
                    SUpRe
                End If
            End If
    End Select
    DoEvents
    
    SAnza
    DoEvents

    If WindowLoad("frmAufga") = True Then
        If GlWaT = RibTab_Wart_Beha Then
            WaSpl RibTab_Wart_Beha
            S_WaLa RibTab_Wart_Beha
        End If
    End If
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo6 = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAbs " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    TxDa1.Text = FDaPr(NeuDa)
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
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
    TxDa1.Text = FDaPr(NeuDa)
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
Private Function FDaPr(ByVal NeuDa As Date) As Date
On Error GoTo OrErr

Dim ReDat As Date
Dim AnzPo As Integer
Dim ReAbg As Boolean
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select
Set RpRow = RpSel(0)

If RpRow.GroupRow = False Then
    Set RpCol = RpCls.Find(Rec_Datum)
    ReDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
    Select Case GlBut
    Case RibTab_Abrechnung:
        Set RpCol = RpCls.Find(Rec_Selekt)
        ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
    Case RibTab_Rechnungen:
        Set RpCol = RpCls.Find(Rec_Selekt)
        If LCase(RpRow.Record(RpCol.ItemIndex).Value) = "ja" Then
            ReAbg = True
        Else
            ReAbg = False
        End If
    End Select
Else
    ReDat = Date
    ReAbg = False
End If

AnzPo = RpSel.Count

If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
    If AnzPo > 1 Then
        If Year(NeuDa) <> Year(Date) Then
            FDaPr = Date
            SPopu "Ungültiges Rechnungsdatum", "Das Rechnungsdatum muss sich innerhalb des aktuellen Geschäftsjahrs befinden", IC48_Information
        Else
            FDaPr = NeuDa
        End If
    ElseIf AnzPo = 1 Then
        If ReAbg = False Then
            If Year(NeuDa) <> Year(ReDat) Then
                FDaPr = ReDat
                SPopu "Ungültiges Rechnungsdatum", "Das Rechnungsdatum muss sich innerhalb des aktuellen Geschäftsjahrs befinden", IC48_Information
            Else
                FDaPr = NeuDa
            End If
        Else
            FDaPr = NeuDa
        End If
    End If
Else
    FDaPr = NeuDa
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Function

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaPr " & Err.Number
Resume Next

End Function
Private Sub FInit()
On Error GoTo SuErr

Dim ReDat As Date
Dim ManNr As Long
Dim StaKt As Long
Dim StaGe As Long
Dim AnzPo As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim StaRa As Integer
Dim ReAbg As Boolean
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set ImMan = FM.imgManag
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set MoKal = Me.dtpDatu1
Set TxDa1 = Me.txtDatu1
Set PuBu1 = Me.btnDatu1
Set UpCo1 = Me.updCont1
Set CmMan = Me.cmbBehan
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen
Set ChDat = Me.chkDatum
Set CheRe = Me.chkReAbs
Set ChePr = Me.chkRePru
Set ChGeb = Me.chkGePru
Set ChDop = Me.chkDoTer
Set ChMan = Me.chkManda
Set ChKon = Me.chkSaKon
Set ChGeg = Me.chkGegen
Set ChOpo = Me.chkOpGen

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

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select
Set RpRow = RpSel(0)

AnzPo = RpSel.Count

If RpRow.GroupRow = False Then
    Set RpCol = RpCls.Find(Rec_Datum)
    ReDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
    Set RpCol = RpCls.Find(Rec_IDP)
    ManNr = RpRow.Record(RpCol.ItemIndex).Value
    Select Case GlBut
    Case RibTab_Abrechnung:
        Set RpCol = RpCls.Find(Rec_Selekt)
        ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
    Case RibTab_Rechnungen:
        Set RpCol = RpCls.Find(Rec_Selekt)
        If LCase(RpRow.Record(RpCol.ItemIndex).Value) = "ja" Then
            ReAbg = True
        Else
            ReAbg = False
        End If
    End Select
Else
    ReDat = Date
    ReAbg = False
    ManNr = GlMan(GlSMa, 2)
End If

If ReAbg = False Then
    ReDat = Date
End If

If AnzPo > 1 Then
    ReDat = Date
End If

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

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    S_KoMa 4, ManNr
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

With CmMan
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa - 1) = GlThe(AktZa, 0)
    Next AktZa
    .ListIndex = GlSMa - 1
End With

TxDa1.Text = ReDat

If GlBuc = True Then 'Einfache Buchhaltung verwenden
    ChGeg.Caption = "Geldkonto :"
    ChKon.Caption = "Sachkonto :"
Else
    ChGeg.Caption = "Sollkonto :"
    ChKon.Caption = "Habenkonto :"
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
ChDat.BackColor = GlBak
CheRe.BackColor = GlBak
ChePr.BackColor = GlBak
ChGeb.BackColor = GlBak
ChDop.BackColor = GlBak
ChMan.BackColor = GlBak
ChKon.BackColor = GlBak
ChGeg.BackColor = GlBak
ChOpo.BackColor = GlBak

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

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
Private Sub FSet()
On Error GoTo OpErr
'Erzeugt einen offenen Posten

Set ChDat = Me.chkDatum
Set CheRe = Me.chkReAbs
Set ChePr = Me.chkRePru
Set ChGeb = Me.chkGePru
Set ChDop = Me.chkDoTer
Set ChOpo = Me.chkOpGen
Set CmMan = Me.cmbBehan
Set ChMan = Me.chkManda
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen

If CheRe.Value = 1 Then
    ChOpo.Value = xtpChecked
    ChOpo.Enabled = True
    CmKto.Enabled = True
    CmGeg.Enabled = True
Else
    ChDat.Value = xtpUnchecked
    ChOpo.Value = xtpUnchecked
    ChOpo.Enabled = False
    CmKto.Enabled = False
    CmGeg.Enabled = False
End If

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSet " & Err.Number
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

TeTit = IniGetOpt("Hilfe", 50741)
TeMai = IniGetOpt("Hilfe", 50742)
TeInh = IniGetOpt("Hilfe", 50743)
TeFus = IniGetOpt("Hilfe", 50744)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FAbs
End Sub

Private Sub chkDatum_Click()
On Error Resume Next

Dim AltDa As Date

Set ChDat = Me.chkDatum
Set TxDa1 = Me.txtDatu1
Set UpCo1 = Me.updCont1
Set PuBu1 = Me.btnDatu1

AltDa = TxDa1.Text

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

Set CmMan = Me.cmbBehan
Set ChMan = Me.chkManda

If ChMan.Value = xtpChecked Then
    CmMan.Enabled = True
Else
    CmMan.Enabled = False
End If

End Sub

Private Sub chkReAbs_Click()
    FSet
End Sub
Private Sub chkSaKon_Click()
On Error Resume Next

Set ChKon = Me.chkSaKon
Set CmKto = Me.cmbKonto

If ChKon.Value = xtpChecked Then
    CmKto.Enabled = True
Else
    CmKto.Enabled = False
End If

End Sub

Private Sub cmbBehan_Click()
On Error GoTo LdErr

Dim ManNr As Long

Set CmMan = Me.cmbBehan

ManNr = CmMan.ItemData(CmMan.ListIndex)

If FoLad = False Then
    S_KoMa 4, ManNr
End If

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "cmbManda " & Err.Number
Resume Next


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

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

FInit
clFen.FenVor

Set clFen = Nothing

FoLad = False

AFont Me
SFrame 1, Me.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReAbs = Nothing
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

TxDa1.Text = FDaPr(NeuDa)

End Sub
Private Sub updCont1_UpClick()
On Error Resume Next

Dim AltDa As Date
Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text
NeuDa = DateAdd("d", 1, AltDa)

TxDa1.Text = FDaPr(NeuDa)

End Sub
