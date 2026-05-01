VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatEM 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   2775
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      _Version        =   1048579
      _ExtentX        =   7223
      _ExtentY        =   4895
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
End
Attribute VB_Name = "frmKatEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form

Private AktCo As VB.Control
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCls As XtremeReportControl.ReportColumns
Private RpRow As XtremeReportControl.ReportRow
Private Sub FKonf()
On Error GoTo SuErr

Dim RetWe As Long
Dim ZeiUm As Boolean
Dim LiLin As Boolean
Dim AnzZe As Integer
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set RpCon = Me.repCont7
Set ImMan = frmMain.imgManag

AnzZe = 14
ZeiUm = CBool(IniGetVal("Katalog", "KatZei"))
LiLin = CBool(IniGetVal("Katalog", "KatGrl"))

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
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
    .HeaderRowsAllowAccess = False
    .HeaderRowsAllowEdit = False
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .MultiSelectionMode = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind keine Termine vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Termine vorhanden"
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
    .PaintManager.MaxPreviewLines = AnzZe
    .PaintManager.ThemedInplaceButtons = True
    If LiLin = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    If ZeiUm = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.HeaderRowsDividerStyle = xtpReportFixedRowsDividerOutlook
    .ShowGroupBox = False
    .ShowHeaderRows = True
    .ShowHeader = True
    .PreviewMode = True
    .SortedDragDrop = True
    .UnrestrictedDragDrop = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo OpErr

Dim RpCon As XtremeReportControl.ReportControl

Dim FenBr As Long
Dim FenHo As Long

Set RpCon = Me.repCont7

FenBr = Me.ScaleWidth
FenHo = Me.ScaleHeight

RpCon.Move 0, 0, FenBr, FenHo

Set RpCon = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Public Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(0, "IDA", 0, False)
    Set RpCol = .Add(1, "Patientendetails", 400, False)
    If RpCo7.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(2, "ID0", 0, False)
    Set RpCol = .Add(3, "Gelesen", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Resizable = False
    RpCol.Sortable = True
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(1).AutoSize = True

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpal " & Err.Number
Resume Next

End Sub
Private Sub FTerm()
On Error GoTo SuErr

Dim MaiNr As Long
Dim PatNr As Long
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo7 = Me.repCont7
Set RpRws = RpCo7.Rows
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(0)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            MaiNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            MaiNr = 0
        End If
        Set RpCol = RpCls.Find(3)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            PatNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            PatNr = 0
        End If
        If MaiNr > 0 Then
            GlMaY = MaiNr 'Emailflyoutfenster Mailindex
            GlNaT = 1 'Mailtyp (1=View 2=Neu 3=Antwort)
            MaMain MaiNr
        End If
    End If
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpRws = Nothing
Set RpCo7 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTerm " & Err.Number
Resume Next

End Sub
Private Sub Form_Load()
    FKonf
    FSpal
End Sub
Private Sub Form_Resize()
    If GlDcP = False Then
        FPosi
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatEM = Nothing
End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FTerm
End Sub

