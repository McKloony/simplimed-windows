VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmBuAnf 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Anfangsbestände & Saldenvortrag"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1860
      Left            =   100
      TabIndex        =   9
      Top             =   4200
      Width           =   9300
      _Version        =   1048579
      _ExtentX        =   16404
      _ExtentY        =   3281
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   10
      Top             =   6300
      Width           =   9500
      _Version        =   1048579
      _ExtentX        =   16757
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   3400
         TabIndex        =   11
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "Hilfe"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   7500
         TabIndex        =   14
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
         Left            =   6100
         TabIndex        =   13
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
      Begin XtremeSuiteControls.PushButton btnDelet 
         Height          =   400
         Left            =   4700
         TabIndex        =   12
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Löschen"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtKonto 
      Height          =   350
      Left            =   2400
      TabIndex        =   5
      Top             =   2220
      Width           =   4500
      _Version        =   1048579
      _ExtentX        =   7937
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtBetra 
      Height          =   350
      Left            =   2400
      TabIndex        =   3
      Top             =   1620
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cmbBuJah 
      Height          =   310
      Left            =   2400
      TabIndex        =   1
      Top             =   1020
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2302
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbGegen 
      Height          =   310
      Left            =   2400
      TabIndex        =   6
      Top             =   2820
      Width           =   4500
      _Version        =   1048579
      _ExtentX        =   7938
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.ComboBox cmbBehan 
      Height          =   310
      Left            =   2400
      TabIndex        =   8
      Top             =   3420
      Width           =   4500
      _Version        =   1048579
      _ExtentX        =   7938
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox3"
   End
   Begin XtremeSuiteControls.ComboBox cmbWarun 
      Height          =   310
      Left            =   3900
      TabIndex        =   4
      Top             =   1620
      Width           =   2000
      _Version        =   1048579
      _ExtentX        =   3519
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox4"
   End
   Begin XtremeSuiteControls.FlatEdit txtBezei 
      Height          =   200
      Left            =   800
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoNr 
      Height          =   200
      Left            =   400
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.CheckBox chkGewEr 
      Height          =   220
      Left            =   1800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   388
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbBuTex 
      Height          =   315
      Left            =   1200
      TabIndex        =   22
      Top             =   7995
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   714
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtHaben 
      Height          =   310
      Left            =   2400
      TabIndex        =   7
      Top             =   2820
      Width           =   4500
      _Version        =   1048579
      _ExtentX        =   7937
      _ExtentY        =   547
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoHa 
      Height          =   200
      Left            =   2200
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8000
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
      Left            =   2600
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtIdxNr 
      Height          =   200
      Left            =   3000
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.ComboBox cmbKtoRa 
      Height          =   315
      Left            =   3900
      TabIndex        =   2
      Top             =   1020
      Width           =   2000
      _Version        =   1048579
      _ExtentX        =   3519
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   500
      Left            =   400
      TabIndex        =   26
      Top             =   100
      Width           =   8600
      _Version        =   1048579
      _ExtentX        =   15169
      _ExtentY        =   882
      _StockProps     =   79
      Caption         =   $"frmBuAnf.frx":0000
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab05 
      Height          =   210
      Left            =   1000
      TabIndex        =   23
      Top             =   2840
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Geldkonto :"
      Alignment       =   5
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
   Begin XtremeSuiteControls.Label lblLab04 
      Height          =   210
      Left            =   1000
      TabIndex        =   18
      Top             =   2240
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Sachkonto :"
      Alignment       =   5
      Transparent     =   -1  'True
   End
   Begin VB.Label lblLab06 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Mandant :"
      Height          =   200
      Left            =   1000
      TabIndex        =   17
      Top             =   3460
      Width           =   1300
   End
   Begin VB.Label lblLab03 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Betrag :"
      Height          =   210
      Left            =   1000
      TabIndex        =   16
      Top             =   1640
      Width           =   1300
   End
   Begin VB.Label lblLab02 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungsjahr :"
      Height          =   210
      Left            =   1000
      TabIndex        =   15
      Top             =   1040
      Width           =   1300
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   860
      Left            =   0
      Top             =   0
      Width           =   9450
   End
End
Attribute VB_Name = "frmBuAnf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Lbl04 As XtremeSuiteControls.Label
Private Lbl05 As XtremeSuiteControls.Label
Private TxDum As XtremeSuiteControls.FlatEdit
Private TxBet As XtremeSuiteControls.FlatEdit
Private TxHab As XtremeSuiteControls.FlatEdit
Private CmRam As XtremeSuiteControls.ComboBox
Private CmEiK As XtremeSuiteControls.ComboBox
Private ComBu As XtremeSuiteControls.ComboBox
Private ComWa As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private ChRow As XtremeReportControl.ReportRow
Private ChRws As XtremeReportControl.ReportRows
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private FoLad As Boolean
Private KntRa As Integer

Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub FHilfe()
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
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FInit()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim AktZa As Integer
Dim AktKo As Integer
Dim BuJah As Integer
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmBuAnf
Set Rahm0 = FM.frmRahm0
Set CmRam = FM.cmbKtoRa
Set ComBu = FM.cmbBuJah
Set ComWa = FM.cmbWarun
Set CmGeg = FM.cmbGegen
Set CmMan = FM.cmbBehan
Set TxBet = FM.txtBetra
Set TxHab = FM.txtHaben
Set Lbl04 = FM.lblLab04
Set Lbl05 = FM.lblLab05
Set RpCon = FM.repCont1
Set ImMan = frmMain.imgManag

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderFlat
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = True
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = GlAnZ
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.FixedRowHeight = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With CmMan
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa - 1) = GlThe(AktZa, 0)
    Next AktZa
    .AddItem "Alle Mandanten"
    .ItemData(AktZa - 1) = 0
    .ListIndex = 0
    .Enabled = True
End With

AktZa = 1
With ComBu
    For BuJah = Year(Date) - 15 To Year(Date) + 1
        .AddItem BuJah
        .ItemData(AktZa - 1) = AktZa
        AktZa = AktZa + 1
    Next BuJah
    .Text = Year(Date)
End With

For AktZa = 1 To UBound(GlWar)
    ComWa.AddItem GlWar(AktZa, 1)
    ComWa.ItemData(ComWa.NewIndex) = GlWar(AktZa, 0)
Next AktZa
ComWa.ListIndex = GlStW - 1

With CmGeg
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            .AddItem GlGeK(AktZa, 3)
            .ItemData(.NewIndex) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    .AddItem GlSaK(AktKo, 3)
                    .ItemData(.NewIndex) = GlSaK(AktKo, 6) '[IDB]
                End If
            Next AktKo
        Next AktZa
        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(.NewIndex) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
        End If
    End If
End With

For AktZa = 1 To UBound(GlGeK)
    If GlGeK(AktZa, 0) = GlGkK Then
        CmGeg.ListIndex = AktZa - 1
        Exit For
    End If
Next AktZa

With CmRam
    For AktZa = 1 To UBound(GlKoR) 'Standardkontenrahmen
        .AddItem GlKoR(AktZa, 0)
        .ItemData(AktZa - 1) = GlKoR(AktZa, 1)
    Next AktZa
End With

If (GlKtR - 1) <= (CmRam.ListCount) - 1 Then
    CmRam.ListIndex = GlKtR - 1 'Standardkontenrahmen
Else
    CmRam.ListIndex = 0
End If

If GlBuc = True Then 'einfache Buchhaltung verwenden
    TxHab.Visible = False
    CmGeg.Visible = True
    Lbl04.Caption = "Sachkonto :"
    Lbl05.Caption = "Geldkonto :"
Else
    TxHab.Visible = True
    CmGeg.Visible = False
    Lbl04.Caption = "Soll-Konto :"
    Lbl05.Caption = "Haben-Konto :"
End If

TxBet.Text = GlWa2

FM.BackColor = GlBak
Rahm0.BackColor = GlBak

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FKont()
On Error GoTo InErr

Dim ManNr As Long
Dim KtoSo As Long
Dim GeKoB As Long
Dim KtoSt As String
Dim AktZa As Integer

ManNr = GlMan(GlSMa, 2) 'Standardmandant

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then 'Mandantennummer
            If GlMan(AktZa, 28) <> vbNullString Then
                GeKoB = GlMan(AktZa, 28) 'Standardgeldkonto Bankkonto
                Exit For
            End If
        End If
    Next AktZa
Else
    GeKoB = GlGkB 'Standardgeldkonto Bankkonto
End If

For AktZa = 1 To UBound(GlGeK) 'Geldkonten
    If GlGeK(AktZa, 0) = GeKoB Then
        KtoSo = GlGeK(AktZa, 2)
        Exit For
    End If
Next AktZa

KtoSt = SBuFo(KtoSo) 'Sachkontenformatierung

If GlBuc = True Then 'einfache Buchhaltung verwenden
    Me.txtKonto.Text = "9000"
    GlBuF = 9 'Buchungsdialog
    S_KtSu "BuAn", 0, True
Else
    Me.txtHaben.Text = "9000"
    GlBuF = 10 'Buchungsdialog
    S_KtSu "BuAn", 0, True
    Me.txtKonto.Text = KtoSt
    GlBuF = 9
    S_KtSu "BuAn", 0, True
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKont " & Err.Number
Resume Next

End Sub
Private Sub FMand()
On Error GoTo OrErr

Dim ManNr As Long
Dim StaRa As Integer
Dim AktZa As Integer

Set CmMan = Me.cmbBehan
Set CmRam = Me.cmbKtoRa

ManNr = CmMan.ItemData(CmMan.ListIndex)

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            If GlMan(AktZa, 25) <> vbNullString Then
                KntRa = GlMan(AktZa, 25) 'Standardkontenrahmen
            Else
                KntRa = GlKtR
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
Private Sub FSpla()
On Error GoTo InErr

Dim AktZa As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmBuAnf
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns

With RpCls
    Set RpCol = .Add(Anf_ID0, "ID0", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        Set RpCol = .Add(Anf_Konto, "Sachkonto", 80, False)
    Else
        Set RpCol = .Add(Anf_Konto, "Sollkonto", 80, False)
    End If
    With RpCol
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        Set RpCol = .Add(Anf_Geldkonto, "Geldkonto", 80, False)
    Else
        Set RpCol = .Add(Anf_Geldkonto, "Habenkonto", 80, False)
    End If
    With RpCol
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Anf_Datum, "Datum", 80, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Anf_Ausgabe, "Soll", 75, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Anf_Einnahme, "Haben", 75, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Anf_Steuer, "Steuer", 0, False)
    With RpCol
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If GlBuc = True Then 'Einfache Buchhaltung verwenden
        Set RpCol = .Add(Anf_Jahr, "Geldkonto", 90, False)
    Else
        Set RpCol = .Add(Anf_Jahr, "Sachkonto", 90, False)
    End If
    With RpCol
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Anf_Mandant, "Mandant", 50, True)
    With RpCol
        .Editable = False
        .Groupable = False
        .Sortable = False
        .AutoSize = True
        .EditOptions.AllowEdit = False
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        For AktZa = 1 To UBound(GlMan)
            .EditOptions.Constraints.Add GlMan(AktZa, 1), GlMan(AktZa, 2)
        Next AktZa
    End With
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpla " & Err.Number
Resume Next

End Sub

Private Sub btnDelet_Click()
    S_BuLo
    S_BuAnf
End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50101)
TeMai = IniGetOpt("Hilfe", 50102)
TeInh = IniGetOpt("Hilfe", 50103)
TeFus = IniGetOpt("Hilfe", 50104)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    S_BuSm
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    S_BuAnS
    S_BuAnf
End Sub
Private Sub cmbBehan_Click()
    If FoLad = False Then
        S_BuAnf
        FMand
    End If
End Sub
Private Sub cmbBuJah_Click()
    If FoLad = False Then
        S_BuAnf
    End If
End Sub

Private Sub cmbKtoRa_Click()

Set CmRam = Me.cmbKtoRa
    
If FoLad = False Then
    KntRa = CmRam.ListIndex + 1
End If

End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde
Set CmRam = Me.cmbKtoRa

FoLad = True
FInit
FMand
FSpla
S_BuAnf

KntRa = GlKtR

FoLad = False

AFont Me
SFrame 1, Me.hwnd

FKont
DoEvents

FrmEx.TopMost = True

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBuAnf = Nothing
End Sub

Private Sub txtBetra_GotFocus()
    Me.txtBetra.SelStart = 0
    Me.txtBetra.SelLength = Len(Me.txtBetra.Text)
End Sub
Private Sub txtBetra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBetra_LostFocus()
On Error Resume Next

Dim Betra As Double

If Me.txtBetra.Text <> vbNullString Then
    If IsNumeric(Me.txtBetra.Text) = True Then
        Betra = CDbl(Me.txtBetra.Text)
        Me.txtBetra.Text = Format$(Betra, GlWa1)
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
    Select Case KeyCode
    Case vbKeyF2:
            Me.txtHaben.SelLength = 0
    Case vbKeyDown:
            Me.cmbBehan.SetFocus
    Case vbKeyUp:
            Me.txtKonto.SetFocus
    Case vbKeyReturn:
            GlBuF = 8 'Buchungsdialog
            S_KtSu "BuAn", KntRa
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
    Select Case KeyCode
    Case vbKeyF2:
            Me.txtKonto.SelLength = 0
    Case vbKeyDown:
            If GlBuc = True Then 'Einfache Buchhaltung verwenden
                Me.cmbGegen.SetFocus
            Else
                Me.txtHaben.SetFocus
            End If
    Case vbKeyUp:
                Me.cmbWarun.SetFocus
    Case vbKeyReturn:
            GlBuF = 4 'Buchungsdialog
            S_KtSu "BuAn", KntRa
    End Select
End Sub
