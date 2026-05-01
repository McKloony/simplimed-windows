VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatDL 
   BorderStyle     =   0  'Kein
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   2205
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1965
      _Version        =   1048579
      _ExtentX        =   3466
      _ExtentY        =   3889
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
End
Attribute VB_Name = "frmKatDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form

Private AktCo As VB.Control
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Public Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(Med_ID2, "ID2", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Med_ID0, "ID0", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Med_IDR, "IDR", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Med_GOID, "PZN", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(Med_IDKurz, "Heilmitteltext", 10, True)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
    End With
    If RpCo7.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Med_IDD, "Diagnose", 10, True)
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.AllowEdit = True
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleReadOnly
    End With
End With

For Each RpCol In RpCls
    RpCol.Groupable = False
    RpCol.Sortable = False
    RpCol.AutoSize = False
    RpCol.Resizable = False
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(Med_IDKurz).AutoSize = True
RpCls(Med_IDD).AutoSize = True

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpal " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo OpErr

Dim FenBr As Long
Dim FenHo As Long
Dim RpCon As XtremeReportControl.ReportControl

Set RpCon = Me.repCont7

FenBr = Me.ScaleWidth
FenHo = Me.ScaleHeight

If Me.WindowState <> vbMinimized Then
    RpCon.Move 0, 0, FenBr, FenHo
End If

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo SuErr

Dim RetWe As Long
Dim ZeiUm As Boolean
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set RpCon = Me.repCont7
Set ImMan = frmMain.imgManag

ZeiUm = False

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
    .AllowColumnSort = False
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FocusSubItems = True
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Zuordnungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Zuordnungen vorhanden"
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
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    .OLEDropMode = xtpOLEDropNone
End With

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub

Private Sub Form_Resize()
    FPosi
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmKatDL = Nothing
End Sub
Private Sub Form_Load()
    FKonf
    FSpal
End Sub
Private Sub repCont7_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim IdxNr As Long
Dim DiaNr As Long
Dim PZNst As String
Dim RowNr As Integer

If Item.Record(Med_ID2).Value <> vbNullString Then
    If Item.Record(Med_IDD).Value <> vbNullString Then
        If Item.Record(Med_GOID).Value <> vbNullString Then
            RowNr = Item.Index
            IdxNr = Val(Item.Record(Med_ID2).Value)
            DiaNr = Val(Item.Record(Med_IDD).Value)
            PZNst = Item.Record(Med_GOID).Value
            S_RzDi DiaNr, IdxNr, PZNst, RowNr
        End If
    End If
End If

End Sub
Private Sub repCont7_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim AktZa As Integer
Dim GesZa As Integer
Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo7 = Me.repCont7
Set RpRws = RpCo7.Rows
Set HiTes = RpCo7.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

Select Case HiTes.ht
Case xtpHitTestGroupBox:
Case xtpHitTestHeader:
Case xtpHitTestReportArea:
        If Button = vbRightButton Then
            SMePo 2
        End If
Case xtpHitTestUnknown:
End Select

Set RpRws = Nothing
Set RpCo7 = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

