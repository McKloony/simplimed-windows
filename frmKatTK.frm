VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatTK 
   BorderStyle     =   0  'Kein
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   3201
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
End
Attribute VB_Name = "frmKatTK"
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
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set RpCon = Me.repCont7
Set ImMan = frmMain.imgManag

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
    .PaintManager.MaxPreviewLines = GlAnZ
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
    .PreviewMode = False
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
Private Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(0, "ID2", 0, False)
    Set RpCol = .Add(1, "Patientendetails", 400, False)
    If RpCo7.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(2, "ID0", 0, False)
    Set RpCol = .Add(3, "Farbe", 0, False)
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

Dim TerNr As Long
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
            TerNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            TerNr = 0
        End If
        Set RpCol = RpCls.Find(2)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            PatNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            PatNr = 0
        End If
        If TerNr > 0 Then
            TeMain TerNr
        ElseIf PatNr > 0 Then
            AMain PatNr
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
    Set frmKatTK = Nothing
End Sub
Private Sub repCont7_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

If Row.Record(3).Value <> vbNullString Then
    If IsNumeric(Row.Record(3).Value) Then
        FrbZa = Row.Record(3).Value
        If FrbZa > 1 And FrbZa <= 20 Then
            Metrics.BackColor = GlTmF(FrbZa, 1)
        End If
    End If
End If

If Row.Record(2).Value <> vbNullString Then
    If Row.Record(2).Value = 1 Then
        Metrics.Font.Strikethrough = True
        Metrics.ForeColor = 8421504
    End If
End If

End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FTerm
End Sub
