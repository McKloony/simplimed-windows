VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmAdrBLZ 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Bankleitzahlen"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   2540
      Left            =   30
      TabIndex        =   0
      Top             =   900
      Width           =   7860
      _Version        =   1048579
      _ExtentX        =   13864
      _ExtentY        =   4480
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   1100
      Left            =   0
      TabIndex        =   8
      Top             =   3900
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   6000
         TabIndex        =   5
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
         Height          =   400
         Left            =   4600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Einf¸gen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtSuOrt 
      Height          =   350
      Left            =   4360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   300
      Width           =   2000
      _Version        =   1048579
      _ExtentX        =   3528
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtSuPLZ 
      Height          =   350
      Left            =   1660
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   1000
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.CheckBox chkTeOrt 
      Height          =   220
      Left            =   400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2200
      _Version        =   1048579
      _ExtentX        =   3881
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Kurzbezeichnung einf¸gen"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   225
      Left            =   300
      TabIndex        =   7
      Top             =   330
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Suche nach BLZ :"
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label lblLab02 
      Height          =   225
      Left            =   2900
      TabIndex        =   6
      Top             =   330
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Suche nach Bank :"
      Alignment       =   1
   End
End
Attribute VB_Name = "frmAdrBLZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ChOrt As XtremeSuiteControls.CheckBox
Private TxBLZ As XtremeSuiteControls.FlatEdit
Private TxBnk As XtremeSuiteControls.FlatEdit
Private Lbl01 As XtremeSuiteControls.Label
Private Lbl02 As XtremeSuiteControls.Label
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private clFen As clsFenster
Private Sub FKonf()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmAdrBLZ
Set Rahm0 = FM.frmRahm0
Set RpCon = FM.repCont1
Set ChOrt = FM.chkTeOrt
Set TxBLZ = FM.txtSuPLZ
Set TxBnk = FM.txtSuOrt
Set Lbl01 = FM.lblLab01
Set Lbl02 = FM.lblLab02
Set RpCls = RpCon.Columns
Set ImMan = frmMain.imgManag

ChOrt.Value = CBool(IniGetVal("System", "KurBnk"))

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
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
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
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenkˆpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind keine Leistungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Leistungen vorhanden"
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
    .PaintManager.FixedRowHeight = False
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

With RpCls
    Set RpCol = .Add(0, "BLZ", 90, True)
    Set RpCol = .Add(1, "Kreditinstitut", 200, True)
    RpCol.AutoSize = True
    RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
    RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
    Set RpCol = .Add(2, "Ortsangabe", 150, True)
    Set RpCol = .Add(3, "BIC", 130, True)
    Set RpCol = .Add(4, "Kurzname", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Groupable = True
    RpCol.Sortable = False
Next RpCol

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
ChOrt.BackColor = GlBak
Lbl01.BackColor = GlBak
Lbl02.BackColor = GlBak

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FKurz()
On Error GoTo InErr

Set FM = frmAdrBLZ
Set ChOrt = FM.chkTeOrt

If ChOrt.Value = xtpChecked Then
    IniSetVal "System", "KurBnk", -1
Else
    IniSetVal "System", "KurBnk", 0
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FRes()
On Error Resume Next

Set FM = frmAdrBLZ
Set TxBLZ = FM.txtSuPLZ
Set TxBnk = FM.txtSuOrt

TxBLZ.Text = vbNullString
TxBnk.Text = vbNullString

End Sub
Private Sub FSett()
On Error GoTo SeErr
'L‰dt die ausgew‰hlte Adresse in das Adressformular

Dim StBLZ As String
Dim StBnk As String
Dim StOrt As String
Dim StBIC As String
Dim FeBLZ As XtremeSuiteControls.FlatEdit
Dim FeBnk As XtremeSuiteControls.FlatEdit
Dim FeKto As XtremeSuiteControls.FlatEdit
Dim FeBIC As XtremeSuiteControls.FlatEdit
Dim RpCon As XtremeReportControl.ReportControl

If WindowLoad("frmMandant") = True Then
    Set FM = frmMandant
    Set FeBIC = FM.txtS2F34
Else
    Set FM = frmAdress
    Set FeBIC = FM.txtS2F35
End If

Set FeBLZ = FM.txtS2F04
Set FeBnk = FM.txtS2F03
Set FeKto = FM.txtS2F05
Set RpCon = Me.repCont1
Set ChOrt = Me.chkTeOrt
Set RpSel = RpCon.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        StBLZ = Trim$(RpRow.Record(0).Value)
        If RpRow.Record(2).Value <> vbNullString Then StOrt = Trim$(RpRow.Record(2).Value)
        If RpRow.Record(3).Value <> vbNullString Then StBIC = Trim$(RpRow.Record(3).Value)
        If ChOrt.Value = xtpChecked Then
            StBnk = Trim$(RpRow.Record(4).Value)
        Else
            StBnk = Trim$(RpRow.Record(1).Value) & ", " & StOrt
        End If
        FeBLZ.Text = StBLZ
        FeBnk.Text = StBnk
        FeBIC.Text = StBIC
        Unload Me
        FeKto.SetFocus
    End If
End If

Set RpSel = Nothing
Set RpCon = Nothing

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub btnWeiter_Click()
    FSett
End Sub
Private Sub chkTeOrt_Click()
    FKurz
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FKonf
AFont Me

FrmEx.TopMost = True

SFrame 1, Me.hwnd

Set FrmEx = Nothing

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAdrBLZ = Nothing
End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(1).Value = Row.Record(2).Value Then Metrics.Font.Bold = True
End Sub

Private Sub repCont1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FSett
    End If
End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FSett
End Sub

Private Sub txtSuOrt_GotFocus()
    FRes
End Sub

Private Sub txtSuOrt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.txtSuOrt.Text <> vbNullString Then
            Opt_PLZ Me.txtSuOrt.Text, 2
        End If
    End If
End Sub


Private Sub txtSuPLZ_GotFocus()
    FRes
End Sub

Private Sub txtSuPLZ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.txtSuPLZ.Text <> vbNullString Then
            Opt_PLZ Me.txtSuPLZ.Text, 1
        End If
    End If
End Sub
