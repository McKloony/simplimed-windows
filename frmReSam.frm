VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReSam 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnungs¸bersicht"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   18
      Top             =   4600
      Width           =   5200
      _Version        =   1048579
      _ExtentX        =   9172
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   3200
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   400
         Width           =   1140
         _Version        =   1048579
         _ExtentX        =   2011
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Abbrechen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   1700
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   400
         Width           =   1350
         _Version        =   1048579
         _ExtentX        =   2381
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Weiter"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   400
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   400
         Width           =   1140
         _Version        =   1048579
         _ExtentX        =   2011
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hilfe"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3700
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   3800
      _Version        =   1048579
      _ExtentX        =   6703
      _ExtentY        =   6526
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   350
         Left            =   2530
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3260
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   2530
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2760
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit4 
         Height          =   225
         Left            =   300
         TabIndex        =   8
         Top             =   2800
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zeitraum"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit3 
         Height          =   225
         Left            =   300
         TabIndex        =   6
         Top             =   2200
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Jahr"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit2 
         Height          =   225
         Left            =   300
         TabIndex        =   4
         Top             =   1600
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Quartal"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit1 
         Height          =   225
         Left            =   300
         TabIndex        =   2
         Top             =   1000
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Monat"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMonat 
         Height          =   315
         Left            =   1300
         TabIndex        =   3
         Top             =   960
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
         Height          =   315
         Left            =   1300
         TabIndex        =   5
         Top             =   1560
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
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   1300
         TabIndex        =   9
         Top             =   2760
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   350
         Left            =   1300
         TabIndex        =   11
         Top             =   3260
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbJahre 
         Height          =   315
         Left            =   1300
         TabIndex        =   7
         Top             =   2160
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu1 
         Height          =   405
         Left            =   3120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   706
         _StockProps     =   64
         Show3DBorder    =   2
         VisualTheme     =   0
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte legen Sie den Zeitraum fest, f¸r den die Rechnungs¸bersicht erstellt werden soll."
         Height          =   495
         Left            =   300
         TabIndex        =   16
         Top             =   120
         Width           =   3500
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   3300
         Width           =   900
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
   Begin XtremeSuiteControls.ComboBox cmbManda 
      Height          =   315
      Left            =   1500
      TabIndex        =   13
      Top             =   3980
      Width           =   2600
      _Version        =   1048579
      _ExtentX        =   4577
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Mandant :"
      Height          =   195
      Left            =   500
      TabIndex        =   17
      Top             =   4040
      Width           =   900
   End
End
Attribute VB_Name = "frmReSam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxDum As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private CmMon As XtremeSuiteControls.ComboBox
Private CmQua As XtremeSuiteControls.ComboBox
Private CmJah As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private OpZei As XtremeSuiteControls.RadioButton
Private MoKal As XtremeCalendarControl.DatePicker
Private ImMan As XtremeCommandBars.ImageManager
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private KalWa As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub FKale()
On Error GoTo LaErr
'L‰ﬂt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1
Set Rahm1 = Me.frmRahm1

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
    Else
        NeuDa = Date
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
    Else
        NeuDa = Date
    End If
End Select

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    Select Case KalWa
    Case 1: .Top = TxDa1.Top + TxDa1.Height
            .Left = TxDa1.Left + Rahm1.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa1.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left + Rahm1.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    End Select
End With

Datu1 = TxDa1.Text
Datu2 = TxDa2.Text

If Datu2 < Datu1 Then TxDa1.Text = Datu2

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Function FStar() As String
On Error GoTo InErr

Dim DaSta As Date
Dim DaEnd As Date
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim Krit1 As String
Dim Datu1 As String
Dim Datu2 As String
Dim AkMon As Integer
Dim AkJha As Integer
Dim AkQua As Integer
Dim Mld1, Tit1 As String

Set OpMon = Me.optZeit1
Set OpQua = Me.optZeit2
Set OpJah = Me.optZeit3
Set OpZei = Me.optZeit4
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQurta
Set CmJah = Me.cmbJahre
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDum = Me.txtDummy

If IsDate(TxDa1.Text) Then
    DaSta = TxDa1.Text
Else
    DaSta = Date
End If

If IsDate(TxDa2.Text) Then
    DaEnd = TxDa2.Text
Else
    DaEnd = Date
End If

AkJha = CInt(CmJah.Text)
AkMon = CmMon.ItemData(CmMon.ListIndex)
AkQua = CmQua.ItemData(CmQua.ListIndex)

Datu1 = DatePart("m", DaSta) & "/" & DatePart("d", DaSta) & "/" & DatePart("yyyy", DaSta)
Datu2 = DatePart("m", DaEnd) & "/" & DatePart("d", DaEnd) & "/" & DatePart("yyyy", DaEnd)
Mld1 = "Sie haben keinen Auswertungszeitraum gew‰hlt"
Tit1 = "Rechnungs¸bersicht"

If OpMon.Value = True Then
    If GlTyp < 2 Then
        Krit1 = "(((MONTH(Datum))=" & AkMon & ") AND ((YEAR(Datum))=" & AkJha & "))"
    Else
        Krit1 = "(((Month([Datum]))=" & AkMon & ") AND ((Year([Datum]))=" & AkJha & "))"
    End If
    TxDum.Text = CmMon.Text & " / " & CmJah.Text
ElseIf OpQua.Value = True Then
    Select Case GlTyp
    Case 0:
        Select Case AkQua
        Case 1: Krit1 = "((Datum >= '01.01." & AkJha & "') AND (Datum <= '31.03." & AkJha & "'))"
        Case 2: Krit1 = "((Datum >= '01.04." & AkJha & "') AND (Datum <= '30.06." & AkJha & "'))"
        Case 3: Krit1 = "((Datum >= '01.07." & AkJha & "') AND (Datum <= '30.09." & AkJha & "'))"
        Case 4: Krit1 = "((Datum >= '01.10." & AkJha & "') AND (Datum <= '31.12." & AkJha & "'))"
        End Select
    Case 1:
        Select Case AkQua
        Case 1: Krit1 = "((Datum >= '01.01." & AkJha & "') AND (Datum <= '31.03." & AkJha & "'))"
        Case 2: Krit1 = "((Datum >= '01.04." & AkJha & "') AND (Datum <= '30.06." & AkJha & "'))"
        Case 3: Krit1 = "((Datum >= '01.07." & AkJha & "') AND (Datum <= '30.09." & AkJha & "'))"
        Case 4: Krit1 = "((Datum >= '01.10." & AkJha & "') AND (Datum <= '31.12." & AkJha & "'))"
        End Select
    Case 2:
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] Between #01/01/" & AkJha & "# AND #03/31/" & AkJha & "#))"
        Case 2: Krit1 = "(([Datum] Between #04/01/" & AkJha & "# AND #06/30/" & AkJha & "#))"
        Case 3: Krit1 = "(([Datum] Between #07/01/" & AkJha & "# AND #09/30/" & AkJha & "#))"
        Case 4: Krit1 = "(([Datum] Between #10/01/" & AkJha & "# AND #12/31/" & AkJha & "#))"
        End Select
    Case 3:
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] Between #01/01/" & AkJha & "# AND #03/31/" & AkJha & "#))"
        Case 2: Krit1 = "(([Datum] Between #04/01/" & AkJha & "# AND #06/30/" & AkJha & "#))"
        Case 3: Krit1 = "(([Datum] Between #07/01/" & AkJha & "# AND #09/30/" & AkJha & "#))"
        Case 4: Krit1 = "(([Datum] Between #10/01/" & AkJha & "# AND #12/31/" & AkJha & "#))"
        End Select
    End Select
    TxDum.Text = CmQua.Text & " / " & CmJah.Text
ElseIf OpJah.Value = True Then
    If GlTyp < 2 Then
        Krit1 = "((YEAR(Datum) = " & AkJha & "))"
    Else
        Krit1 = "((Year([Datum]) = " & AkJha & "))"
    End If
    TxDum.Text = "Jahr: " & CmJah.Text
ElseIf OpZei.Value = True Then
    Select Case GlTyp
    Case 0: Krit1 = "((Datum >= '" & DaSta & "') AND (Datum <= '" & DaEnd & "'))"
    Case 1: Krit1 = "((Datum >= '" & DaSta & "') AND (Datum <= '" & DaEnd & "'))"
    Case 2: Krit1 = "(([Datum] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
    Case 3: Krit1 = "(([Datum] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
    End Select
    TxDum.Text = DaSta & " - " & DaEnd
Else
    WindowMess Mld1, Dial2, Tit1, Me.hwnd
End If

If Krit1 <> vbNullString Then
    FStar = Krit1
Else
    FStar = vbNullString
End If

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStar " & Err.Number
Resume Next

End Function
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
        TxDa1.Text = NeuDa
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
        TxDa2.Text = NeuDa
    End If
End Select

With MoKal
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo LaErr

Dim RetWe As Long
Dim AktZa As Integer
Dim IdxZa As Integer
Dim AkMon As Integer
Dim AkQua As Integer
Dim BuJah As Integer

Set FM = frmReSam
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set CmMan = Me.cmbManda
Set OpMon = FM.optZeit1
Set OpQua = FM.optZeit2
Set OpJah = FM.optZeit3
Set OpZei = FM.optZeit4
Set CmMon = FM.cmbMonat
Set CmQua = FM.cmbQurta
Set CmJah = FM.cmbJahre
Set MoKal = FM.dtpDatu1
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnDatu2
Set ImMan = frmMain.imgManag

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
    .ToolTipText = "Markieren Sie bitte hier den gw¸nschten Rechnungstag"
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
    For BuJah = Year(Date) - 15 To Year(Date) + 1
        .AddItem BuJah
        .ItemData(.NewIndex) = IdxZa
        IdxZa = IdxZa + 1
    Next BuJah
    .Text = Year(Date)
End With

With CmMan
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa - 1) = GlThe(AktZa, 0)
    Next AktZa
    .AddItem "Alle Mandanten"
    .ItemData(AktZa - 1) = 0
    .ListIndex = UBound(GlThe)
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date) - 1
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date) + 1
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

RetWe = SendMessage(CmMon.hwnd, CB_SETCURSEL, AkMon - 1, ByVal 0&)
RetWe = SendMessage(CmQua.hwnd, CB_SETCURSEL, AkQua - 1, ByVal 0&)

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

FM.BackColor = GlBak
OpMon.BackColor = GlBak
OpQua.BackColor = GlBak
OpJah.BackColor = GlBak
OpZei.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak

Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    Select Case KalWa
    Case 1: TxDa1.Text = NeuDa
            TxDa1.SetFocus
    Case 2: TxDa2.Text = NeuDa
            TxDa2.SetFocus
    End Select
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo InErr

Dim ManNr As Long
Dim KopTe As String
Dim Krit1 As String
Dim Krit2 As String
Dim ZeRau As Integer

Set CmMan = Me.cmbManda
Set OpMon = Me.optZeit1
Set OpQua = Me.optZeit2
Set OpJah = Me.optZeit3
Set OpZei = Me.optZeit4

Krit1 = FStar
Krit2 = " AND ([PatNr]=" & GlAdr & ")"

KopTe = TxDum.Text

ManNr = CmMan.ItemData(CmMan.ListIndex)

If ManNr > 0 Then
    Krit2 = Krit2 & " AND ([IDP] = " & ManNr & ")"
End If

If OpMon.Value = True Then
    ZeRau = 1
ElseIf OpJah.Value = True Then
    ZeRau = 2
ElseIf OpQua.Value = True Then
    ZeRau = 3
Else
    ZeRau = 4
End If

If Krit1 <> vbNullString Then
    With GlBuD
        .Krit1 = Krit1 & Krit2
        .ZeRau = ZeRau
        .ManNr = ManNr
    End With
    SDruck "ReUber", True, KopTe, False
Else
    Unload Me
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub
Private Sub btnDatu2_Click()
    KalWa = 2
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
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub

Private Sub cmbJahre_Click()
    Me.optZeit3.Value = True
End Sub
Private Sub cmbMonat_Click()
    Me.optZeit1.Value = True
End Sub
Private Sub cmbQurta_Click()
    Me.optZeit2.Value = True
End Sub
Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub

Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub

