VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmWieder 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Wiedervorlage"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   13
      Top             =   4900
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
         TabIndex        =   17
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
         Left            =   4400
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "Sp&eichern"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnKommen 
         Height          =   400
         Left            =   3000
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Kommentar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   1700
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
   Begin XtremeSuiteControls.FlatEdit txtKomme 
      Height          =   1900
      Left            =   800
      TabIndex        =   12
      Tag             =   "0Kommentar"
      Top             =   2840
      Width           =   6100
      _Version        =   1048579
      _ExtentX        =   10760
      _ExtentY        =   3351
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   2700
      Left            =   800
      TabIndex        =   11
      Top             =   40
      Width           =   6080
      _Version        =   1048579
      _ExtentX        =   10724
      _ExtentY        =   4762
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox txtBetre 
         Height          =   315
         Left            =   1000
         TabIndex        =   2
         Tag             =   "0IDKurz"
         Top             =   750
         Width           =   4500
         _Version        =   1048579
         _ExtentX        =   7938
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   350
         Left            =   4010
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtVonZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   2250
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatum"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.CheckBox chkErled 
         Height          =   240
         Left            =   4600
         TabIndex        =   8
         Tag             =   "0Erledigt"
         Top             =   1220
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Erledigt"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatum 
         Height          =   350
         Left            =   1000
         TabIndex        =   3
         Tag             =   "0VonDat"
         Top             =   1200
         Width           =   1230
         _Version        =   1048579
         _ExtentX        =   2170
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtVonZe 
         Height          =   350
         Left            =   3200
         TabIndex        =   6
         Tag             =   "0ZeiVon"
         Top             =   1200
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   310
         Left            =   1000
         TabIndex        =   9
         Tag             =   "0IDP"
         Top             =   1650
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtPatie 
         Height          =   350
         Left            =   1000
         TabIndex        =   1
         Tag             =   "0Patient"
         Top             =   300
         Width           =   4500
         _Version        =   1048579
         _ExtentX        =   7937
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   2520
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   1200
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   310
         Left            =   1000
         TabIndex        =   10
         Tag             =   "0IDM"
         Top             =   2100
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2140
         Width           =   840
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   350
         Width           =   840
         _Version        =   1048579
         _ExtentX        =   1482
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Patient :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Lab08 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1690
         Width           =   840
      End
      Begin VB.Label Lab01 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Betreff :"
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   790
         Width           =   840
      End
      Begin VB.Label Lab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Datum :"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1240
         Width           =   840
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtID2 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Tag             =   "0ID2"
      Top             =   6100
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   5
   End
   Begin XtremeSuiteControls.FlatEdit txtID0 
      Height          =   200
      Left            =   300
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "0ID0"
      Top             =   6100
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   5
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   405
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6500
      Visible         =   0   'False
      Width           =   405
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
End
Attribute VB_Name = "frmWieder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxAdr As XtremeSuiteControls.FlatEdit
Private TxKom As XtremeSuiteControls.FlatEdit
Private FoID0 As XtremeSuiteControls.FlatEdit
Private VoZei As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private CmBet As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private ChErl As XtremeSuiteControls.CheckBox
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private MoKal As XtremeCalendarControl.DatePicker

Private Const CB_SHOWDROPDOWN = &H14F
Private Const KEYEVENTF_KEYUP = &H2

Private TagWe As String

Private clFen As clsFenster
Private clFil As clsFile
Private clWor As clsWord

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub FKaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatum
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    TxDa1.Text = NeuDa
End If

With MoKal
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaKo " & Err.Number
Resume Next

End Sub

Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu1 As Date

Set TxDa1 = Me.txtDatum
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
    .Top = (TxDa1.Top + Rahm1.Top) + TxDa1.Height
    .Left = TxDa1.Left + Rahm1.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TxDa1.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Datu1 = TxDa1.Text

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatum

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    TxDa1.Text = NeuDa
End If

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
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
Private Sub btnKommen_Click()
On Error Resume Next

Dim AltDa As Date
Dim TmVon As Date
Dim KoStr As String
Dim TagWe As String

Set VoZei = Me.txtVonZe
Set TxDa1 = Me.txtDatum
Set TxKom = Me.txtKomme

AltDa = TxDa1.Text
TmVon = TimeValue(VoZei.Text)

If TxKom.Text <> vbNullString Then
    KoStr = TxKom.Text & vbCrLf & AltDa & " - " & Format$(TimeValue(Now), "hh:mm") & ": "
Else
    KoStr = AltDa & " - " & Format$(TimeValue(Now), "hh:mm") & ": "
End If

TxKom.Text = KoStr
TagWe = Mid$(TxKom.Tag, 2, Len(TxKom.Tag) - 1)
TxKom.Tag = 1 & TagWe
TxKom.SetFocus
TxKom.SelLength = 0
TxKom.SelStart = Len(TxKom.Text)

GlKoS = True

End Sub
Private Sub btnSchließ_Click()
    FClos
    Unload Me
End Sub

Private Sub btnWeiter_Click()
    FKSav
    Unload Me
End Sub

Private Sub chkErled_Click()

TagWe = Mid$(Me.chkErled.Tag, 2, Len(Me.chkErled.Tag) - 1)

If GlKoL = False Then
    Me.chkErled.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub cmbBehan_Click()

TagWe = Mid$(Me.cmbBehan.Tag, 2, Len(Me.cmbBehan.Tag) - 1)

If GlKoL = False Then
    Me.cmbBehan.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub cmbMitar_Click()

TagWe = Mid$(Me.cmbMitar.Tag, 2, Len(Me.cmbMitar.Tag) - 1)

If GlKoL = False Then
    Me.cmbMitar.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub txtBetre_Change()
    
TagWe = Mid$(Me.txtBetre.Tag, 2, Len(Me.txtBetre.Tag) - 1)

If GlKoL = False Then 'Formular wird geladen
    Me.txtBetre.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub txtBetre_Click()

TagWe = Mid$(Me.txtBetre.Tag, 2, Len(Me.txtBetre.Tag) - 1)

If GlKoL = False Then 'Formular wird geladen
    Me.txtBetre.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtBetre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtDatum_Change()

TagWe = Mid$(Me.txtDatum.Tag, 2, Len(Me.txtDatum.Tag) - 1)

If GlKoL = False Then
    Me.txtDatum.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtDatum_GotFocus()
    Me.txtDatum.SelStart = 0
    Me.txtDatum.SelLength = Len(Me.txtDatum.Text)
    GlKoL = False
End Sub
Private Sub txtDatum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtDatum_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtDatum.SelLength = 0
    End If
End Sub
Private Sub txtDatum_LostFocus()
    FDaKo
End Sub

Private Sub txtPatie_GotFocus()
    If GlTeF = False Then
        Me.txtPatie.SelStart = 0
        Me.txtPatie.SelLength = Len(Me.txtPatie.Text)
    End If
End Sub

Private Sub txtPatie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtPatie_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtPatie.SelLength = 0
    End If
End Sub

Private Sub txtPatie_LostFocus()
On Error Resume Next

Dim GesZa As Long
Dim SuStr As String
Dim FLis1 As XtremeSuiteControls.ListBox
    
Set FoID0 = Me.txtID0
Set TxAdr = Me.txtPatie

SuStr = TxAdr.Text

If SuStr <> vbNullString Then
    GesZa = Wie_Adr(SuStr)
    If GesZa > 1 Then
        Load frmWiedAnh
        Set FM = frmWiedAnh
        Set FLis1 = FM.lstList1
        FM.Show vbModal
    End If
    GlKoS = True
End If

End Sub
Private Sub txtVonZe_Click()
On Error Resume Next

TagWe = Mid$(Me.txtVonZe.Tag, 2, Len(Me.txtVonZe.Tag) - 1)

If GlKoL = False Then
    Me.txtVonZe.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtVonZe_GotFocus()
    Me.txtVonZe.SelStart = 0
    Me.txtVonZe.SelLength = Len(Me.txtVonZe.Text)
    GlKoS = True
    GlKoL = False
End Sub

Private Sub txtVonZe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtVonZe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtVonZe.SelLength = 0
End Sub

Private Sub FClos()
On Error GoTo SaErr

GlKoS = False

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub

Private Sub FKonf()
On Error Resume Next

Dim NeuDa As Date
Dim AktZa As Integer
Dim mAnza As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmDat As XtremeCommandBars.CommandBarEdit

Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set VoZei = Me.txtVonZe
Set TxDa1 = Me.txtDatum
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set CmBet = Me.txtBetre
Set TxKom = Me.txtKomme
Set MoKal = Me.dtpDatu1

If WindowLoad("frmAufga") = True Then
    Set CmBrs = frmAufga.comBar02
    Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)
    NeuDa = CDate(CmDat.Text)
Else
    NeuDa = Date
End If

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

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

With VoZei
    .SetMask "00:00", "__:__"
    .Text = "12:00"
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(NeuDa, "dd.mm.yyyy")
End With

If GlMaV = True Then 'Mandanten vorhanden
    For AktZa = 1 To UBound(GlMaA) 'Aktive Mitarbeiter
        If GlKoN = True Then
            CmMan.AddItem GlMaA(AktZa, 1)
            CmMan.ItemData(AktZa - 1) = GlMaA(AktZa, 2)
        Else
            CmMan.AddItem GlMaA(AktZa, 1)
            CmMan.ItemData(AktZa - 1) = GlMaA(AktZa, 2)
        End If
    Next AktZa
End If

If GlMiV = True Then 'Mitarbeiter vorhanden
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        If GlKoN = True Then
            CmMit.AddItem GlMiA(AktZa, 1)
            CmMit.ItemData(AktZa - 1) = GlMiA(AktZa, 2)
        Else
            CmMit.AddItem GlMiA(AktZa, 1)
            CmMit.ItemData(AktZa - 1) = GlMiA(AktZa, 2)
        End If
    Next AktZa
End If

For AktZa = 1 To UBound(GlBtr)
    With CmBet
        .AddItem GlBtr(AktZa, 1)
        .ItemData(AktZa - 1) = GlBtr(AktZa, 0)
    End With
Next AktZa

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

TxKom.Font.Name = GlTFt.Name
TxKom.Font.SIZE = GlTFt.SIZE

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Me.chkErled.BackColor = GlBak

clFen.FenVor

Set clFen = Nothing

End Sub
Private Sub FKSav()
On Error GoTo SaErr

If GlKoS = True Then
    If GlKoN = True Then
        Wie_San
    Else
        Wie_Sav
    End If
    If WindowLoad("frmAufga") = True Then
        S_WaLa RibTab_Wart_Wied
    Else
        S_StSt3
    End If
    GlKoS = False
End If

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "Kon_Sav " & Err.Number
Resume Next

End Sub
Private Sub Form_Load()
On Error Resume Next

GlKoL = True

FKonf
AFont Me
SFrame 1, Me.hwnd

GlKoL = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmWieder = Nothing
End Sub

Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub txtKomme_Change()

TagWe = Mid$(Me.txtKomme.Tag, 2, Len(Me.txtKomme.Tag) - 1)

If GlKoL = False Then
    Me.txtKomme.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub txtKomme_GotFocus()
    GlKoL = False
End Sub

Private Sub txtKomme_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtKomme.SelLength = 0
    End If
End Sub

Private Sub txtVonZe_LostFocus()
On Error Resume Next

Dim TmVon As Date
Dim TmBis As Date

Set VoZei = Me.txtVonZe

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
Else
    TmVon = Now
    VoZei.Text = Format$(TimeValue(Now), "hh:mm")
End If

If GlKoL = False Then
    GlKoS = True
End If

End Sub

Private Sub updCont1_DownClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatum

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatum

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub

Private Sub updCont2_DownClick()
On Error Resume Next

Dim AlDa1 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If GlMiT(MiIdx, 8) > 0 Then
        ZeiRa = GlMiT(MiIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If GlMaT(MaIdx, 8) > 0 Then
        ZeiRa = GlMaT(MaIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If VoZei.Text <> vbNullString Then
    AlDa1 = TimeValue(VoZei.Text)
    TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
    
    TmVon = DateAdd("n", -MiDif, AlDa1)
    VoZei.Text = Format$(TmVon, "hh:mm")
    VoZei.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub updCont2_UpClick()
On Error Resume Next

Dim AlDa1 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If GlMiT(MiIdx, 8) > 0 Then
        ZeiRa = GlMiT(MiIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If GlMaT(MaIdx, 8) > 0 Then
        ZeiRa = GlMaT(MaIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If VoZei.Text <> vbNullString Then
    AlDa1 = TimeValue(VoZei.Text)
    TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
    
    TmVon = DateAdd("n", MiDif, AlDa1)
    VoZei.Text = Format$(TmVon, "hh:mm")
    VoZei.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub btnDatu1_Click()
    FKale
End Sub

Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)

Dim AktTa As Long
Dim AktZa As Integer

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTeV = True Then 'Termine vorhanden
    If GlTpV = True Then 'Kalendermarker vorhanden
        For AktTa = 0 To GlKMa - 1 'Anzahl Kalendermatker
            If Day = Left$(GlTEr(0, AktTa), 10) Then
                For AktZa = 1 To UBound(GlTep) 'Kalendermarker
                    If GlTep(AktZa, 0) = GlTEr(1, AktTa) Then
                        Metrics.BackColor = GlTep(AktZa, 2)
                        Exit For
                    End If
                Next AktZa
            End If
        Next AktTa
    End If
End If

End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub

Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatum
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

