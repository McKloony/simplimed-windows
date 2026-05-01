VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmBaAbl 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Buchungsableitung"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   5200
      Left            =   220
      TabIndex        =   1
      Top             =   200
      Width           =   6500
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   9172
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   6
      Top             =   5900
      Width           =   7000
      _Version        =   1048579
      _ExtentX        =   12347
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   5000
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
      Begin XtremeSuiteControls.PushButton cmdWeite 
         Height          =   400
         Left            =   3600
         TabIndex        =   4
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
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   2300
         TabIndex        =   3
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
   Begin XtremeSuiteControls.CheckBox chkBuSpl 
      Height          =   220
      Left            =   400
      TabIndex        =   2
      Top             =   5600
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Gesplittete Buchung"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   7500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBaAbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private ChSpl As XtremeSuiteControls.CheckBox
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows

Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Sub FEinf()
On Error Resume Next
    
Dim BuSpl As Boolean 'Splittbuchung
    
Set ChSpl = FM.chkBuSpl

If ChSpl.Value = xtpChecked Then
    BuSpl = True
End If
    
GlBuV = True 'Buchungsvorlage einf¸gen
K_BuVs BuSpl

End Sub
Private Sub FInit()
On Error GoTo LaErr
'Reportcontrol Settings

Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmBaAbl
Set Rahm0 = FM.frmRahm0
Set ChSpl = FM.chkBuSpl
Set RpCon = FM.repCont7
Set ImMan = frmMain.imgManag

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
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .MultiSelectionMode = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenkˆpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Buchungsvorlagen vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Buchungsvorlagen vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = -2147483641
    .PaintManager.MaxPreviewLines = 5
    .PaintManager.ThemedInplaceButtons = True
    If GlKLi = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    If GlKaU = True Then 'Zeilenumbruch der Kataloge
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
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlKaZ
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .SortedDragDrop = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    .UnrestrictedDragDrop = True
End With

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
ChSpl.BackColor = GlBak

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " FInit " & Err.Number
Resume Next

End Sub
Private Sub FSpal()
On Error GoTo SpErr
'Formratieren der Spalten

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(Buh_ID0, "ID0", 0, False)
    Set RpCol = .Add(Buh_Datum, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Buh_Buchtext, "Buchungstext", 0, True)
    Set RpCol = .Add(Buh_Einnahme, "Betrag", 0, True)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Buh_Ausgabe, "Brutto", 0, True)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    RpCol.Alignment = xtpAlignmentRight
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        Set RpCol = .Add(Buh_Sachkonto, "Sachkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Geldkonto", 0, True)
    Else
        Set RpCol = .Add(Buh_Sachkonto, "Sachkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Geldkonto", 0, True)
    End If
    Set RpCol = .Add(Buh_RechNr, "Belegzeichen", 0, True)
    Set RpCol = .Add(Buh_IDR, "IDR", 0, False)
    Set RpCol = .Add(Buh_Beleg, "Nummer", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Buh_Sachkontenbez, "Sachkontenbezeichnung", 0, True)
    Set RpCol = .Add(Buh_Geldkontenbez, "Geldkontenbezeichnung", 0, True)
    Set RpCol = .Add(Buh_Steuer, "Steuer", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Buh_W, "W", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Buh_Privat, "Privat", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_Abziehbar, "Abziehbar", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDB, "IDB", 0, False)
    Set RpCol = .Add(Buh_IDA, "IDA", 0, False)
    Set RpCol = .Add(Buh_W‰hrung, "W‰hrung", 0, False)
    Set RpCol = .Add(Buh_Ermittlung, "KE", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Dokument, "DK", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_Paperclip
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_IDP, "IDP", 0, False)
    Set RpCol = .Add(Buh_IDArt, "IDArt", 0, False)
    Set RpCol = .Add(Buh_IDBank, "IDBank", 0, False)
    Set RpCol = .Add(Buh_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Buh_IDT, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Buh_Berichtdatum, "Berichtdatum", 0, True)
    Set RpCol = .Add(Buh_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Buh_Monat, "Monat", 0, False)
    Set RpCol = .Add(Buh_Storniert, "Storniert", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Buh_Zuordnung, "ZU", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_User_Norm
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Lock, "Lock", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconLeft
        .Icon = IC16_Lock
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Datei, "Datei", 0, False)
    Set RpCol = .Add(Buh_Doppelt, "DO", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
End With

For Each RpCol In RpCls
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = True
        .Sortable = True
        .AutoSize = False
        .AutoSortWhenGrouped = False
    End With
Next RpCol

If GlTFt.SIZE > 10 Then
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 0
    RpCls(Buh_Buchtext).Width = 250
    RpCls(Buh_Einnahme).Width = 100
    RpCls(Buh_Ausgabe).Width = 0
Else
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 0
    RpCls(Buh_Buchtext).Width = 220
    RpCls(Buh_Einnahme).Width = 80
    RpCls(Buh_Ausgabe).Width = 0
End If
RpCls(Buh_Sachkonto).Width = 80
RpCls(Buh_Gegenkonto).Width = 0
RpCls(Buh_RechNr).Width = 0
RpCls(Buh_IDR).Width = 0
RpCls(Buh_Beleg).Width = 0
RpCls(Buh_Sachkontenbez).Width = 180
RpCls(Buh_Geldkontenbez).Width = 0
RpCls(Buh_Steuer).Width = 75
RpCls(Buh_W).Width = 40
RpCls(Buh_Privat).Width = 0
RpCls(Buh_Abziehbar).Width = 0
RpCls(Buh_IDB).Width = 0
RpCls(Buh_IDA).Width = 0
RpCls(Buh_W‰hrung).Width = 0
RpCls(Buh_Ermittlung).Width = 25
RpCls(Buh_Dokument).Width = 0
RpCls(Buh_IDP).Width = 0
RpCls(Buh_IDArt).Width = 0
RpCls(Buh_IDBank).Width = 0
RpCls(Buh_Kommentar).Width = 0
RpCls(Buh_IDT).Width = 150
RpCls(Buh_Berichtdatum).Width = 0
RpCls(Buh_GuiID).Width = 0
RpCls(Buh_Monat).Width = 0
RpCls(Buh_Storniert).Width = 0
RpCls(Buh_IDM).Width = 150
RpCls(Buh_Zuordnung).Width = 18
RpCls(Buh_Lock).Width = 18
RpCls(Buh_Datei).Width = 0
RpCls(Buh_Doppelt).Width = 0

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpal " & Err.Number
Resume Next

End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50061)
TeMai = IniGetOpt("Hilfe", 50062)
TeInh = IniGetOpt("Hilfe", 50063)
TeFus = IniGetOpt("Hilfe", 50064)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload frmBaAbl
End Sub
Private Sub cmdWeite_Click()
    FEinf
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FInit
FSpal
K_BuVpl "BaAb"
P_List "BaAb", 0, 1

AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

Set FrmEx = Nothing
    
End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FEinf
End Sub
Private Sub repCont7_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim GeKto As Long

If Row.Record(Buh_IDB).Value <> vbNullString Then
    If Row.Record(Buh_IDB).Value > 0 Then
        GeKto = Row.Record(Buh_IDB).Value
    Else
        GeKto = 0
    End If
Else
    GeKto = 0
End If

If GeKto > 0 Then
    If GeKto <= UBound(GlGeK) Then
        If CBool(GlGeK(Row.Record(Buh_IDB).Value, 5)) = True Then
            Metrics.ForeColor = 16711680
        End If
    End If
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBaAbl = Nothing
End Sub
