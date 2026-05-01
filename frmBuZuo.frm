VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmBuZuo 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Buchung Zuordnen"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   18
      Top             =   4200
      Width           =   6400
      _Version        =   1048579
      _ExtentX        =   11289
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4400
         TabIndex        =   19
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
         Default         =   -1  'True
         Height          =   400
         Left            =   3000
         TabIndex        =   20
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
      Begin XtremeSuiteControls.PushButton btnZuruk 
         Height          =   400
         Left            =   1600
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zurück"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   300
         TabIndex        =   22
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
      Left            =   100
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4000
      Left            =   300
      TabIndex        =   1
      Top             =   100
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.RadioButton optAufhe 
         Height          =   220
         Left            =   1700
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1900
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zuordnung aufheben"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optRechn 
         Height          =   220
         Left            =   1700
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1500
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "einer Rechnung zuordnen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optPatie 
         Height          =   220
         Left            =   1700
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1100
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "einem Patienten zuordnen"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   400
         Left            =   400
         TabIndex        =   17
         Top             =   2300
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   706
         _StockProps     =   79
         ForeColor       =   192
         Alignment       =   4
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Soll die Buchung einem Patienten oder einer Rechnung zugeordnet werden? Bitte wählen Sie eine Option und klicken auf Weiter."
         Height          =   400
         Left            =   400
         TabIndex        =   12
         Top             =   100
         Width           =   5000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4000
      Left            =   300
      TabIndex        =   2
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   7056
      _StockProps     =   79
      Caption         =   "GroupBox2"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtPost 
         Height          =   350
         Left            =   1000
         TabIndex        =   9
         Top             =   2220
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtNumm 
         Height          =   350
         Left            =   1000
         TabIndex        =   8
         Top             =   1440
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   350
         Left            =   1000
         TabIndex        =   7
         Top             =   680
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   200
         Left            =   1010
         TabIndex        =   16
         Top             =   1940
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Suche nach Postleitzahl :"
         Alignment       =   4
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   200
         Left            =   1010
         TabIndex        =   15
         Top             =   1160
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Suche nach Patientennummer :"
         Alignment       =   4
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   200
         Left            =   1010
         TabIndex        =   14
         Top             =   400
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Suche nach Patientenname :"
         Alignment       =   4
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4000
      Left            =   300
      TabIndex        =   3
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   3240
         Left            =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   700
         Width           =   5650
         _Version        =   1048579
         _ExtentX        =   9984
         _ExtentY        =   5715
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin VB.Label lblLab05 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie dem gewünschten Eintrag, um die Buchung zuzuordnen und klicken auf Weiter."
         Height          =   400
         Left            =   400
         TabIndex        =   13
         Top             =   100
         Width           =   5000
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
End
Attribute VB_Name = "frmBuZuo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Lab02 As XtremeSuiteControls.Label
Private Lab03 As XtremeSuiteControls.Label
Private lab04 As XtremeSuiteControls.Label
Private Lab06 As XtremeSuiteControls.Label
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private FTex3 As XtremeSuiteControls.FlatEdit
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private Opti3 As XtremeSuiteControls.RadioButton
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ImMan As XtremeCommandBars.ImageManager
Private PuBu3 As XtremeSuiteControls.PushButton

Private SuTyp As Integer
Private Function FBuZu() As Boolean
On Error GoTo OpErr
'Buchung zuordnen

Dim PatNr As Long
Dim RecNr As Long
Dim BuLok As Boolean
Dim BuTyp As Integer
Dim Mld1, Mld2 As String
Dim Mld3, Mld4 As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows
Set Lab06 = Me.lblLab06

Mld1 = "Die markierte Buchung wurde bereits festgeschrieben und kann daher nicht mehr zugeordnet werden."
Mld2 = "Die markierte Buchung wurde bereits einem Patienten zugeweisen."
Mld3 = "Die markierte Buchung wurde bereits einer Rechnung zugeweisen."
Mld4 = "Die markierte Buchung ist eine Ausgabe und keine Einnahme."

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Buh_IDP)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            PatNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            PatNr = 0
        End If
        Set RpCol = RpCls.Find(Buh_IDR)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            RecNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            RecNr = 0
        End If
        Set RpCol = RpCls.Find(Buh_IDA)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            BuTyp = RpRow.Record(RpCol.ItemIndex).Value
        Else
            BuTyp = 2
        End If
        Set RpCol = RpCls.Find(Buh_Lock)
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            BuLok = True
        Else
            BuLok = False
        End If
    End If
End If

If BuLok = True Then
    Lab06.Caption = Mld1
    FBuZu = True
Else
    If SuTyp = 1 Then
        If PatNr > 0 Then
            Lab06.Caption = Mld2
            FBuZu = True
        End If
    Else
        If RecNr > 0 Then
            Lab06.Caption = Mld3
            FBuZu = True
        End If
        If BuTyp = 1 Then
            Lab06.Caption = Mld4
            FBuZu = True
        End If
    End If
End If

Exit Function

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FBuZu " & Err.Number
Resume Next

End Function
Private Sub FKonf()
On Error GoTo InErr

Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCon = Me.repCont1
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Opti1 = Me.optPatie
Set Opti2 = Me.optRechn
Set Opti3 = Me.optAufhe
Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumm
Set FTex3 = Me.txtPost
Set Lab02 = Me.lblLab02
Set Lab03 = Me.lblLab03
Set lab04 = Me.lblLab04
Set Lab06 = Me.lblLab06
Set PuBu3 = Me.btnZuruk
Set ImMan = FM.imgManag
Set RpCls = RpCon.Columns
Set RpRcs = RpCon.Records

SuTyp = 1

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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Einträge vorhanden"
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
    If GlAnZ = True Then
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
    .PreviewMode = False
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Opti1.BackColor = GlBak
Opti2.BackColor = GlBak
Opti3.BackColor = GlBak
Lab02.BackColor = GlBak
Lab03.BackColor = GlBak
lab04.BackColor = GlBak
Lab06.BackColor = GlBak

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FRes()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumm
Set FTex3 = Me.txtPost

FTex1.Text = vbNullString
FTex2.Text = vbNullString
FTex3.Text = vbNullString

End Sub
Private Sub FSpla()
On Error GoTo InErr

Dim AktZa As Integer
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCon = Me.repCont1
Set RpCls = RpCon.Columns
Set RpRcs = RpCon.Records

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

If SuTyp = 1 Then
    With RpCls
        Set RpCol = .Add(Adr_ID0, "ID0", 0, False)
        Set RpCol = .Add(Adr_ID3, "ID3", 0, False)
        Set RpCol = .Add(Adr_IDKurz, "Suchbegriff", 0, True)
        Set RpCol = .Add(Adr_Geboren, "Geboren", 0, True)
        Set RpCol = .Add(Adr_Name, "Name", 0, True)
        Set RpCol = .Add(Adr_Vorname, "Vorname", 0, True)
        Set RpCol = .Add(Adr_Straße, "Straße", 0, True)
        Set RpCol = .Add(Adr_PLZ, "PLZ", 0, True)
        Set RpCol = .Add(Adr_Ort, "Ort", 0, True)
        Set RpCol = .Add(Adr_Firma1, "Firma", 0, True)
        Set RpCol = .Add(Adr_Telefon1, "Privat", 0, True)
        Set RpCol = .Add(Adr_Telefon2, "Büro", 0, True)
        Set RpCol = .Add(Adr_Telefon3, "Telefax", 0, True)
        Set RpCol = .Add(Adr_Telefon4, "Mobil", 0, True)
        Set RpCol = .Add(Adr_Telefon5, "Email", 0, True)
        Set RpCol = .Add(Adr_Geschlecht, "Geschlecht", 0, True)
        Set RpCol = .Add(Adr_Datum, "Datun", 0, False)
        Set RpCol = .Add(Adr_Briefanrede, "Briefanrede", 0, False)
        Set RpCol = .Add(Adr_Anschrift, "Anschrift", 0, False)
        Set RpCol = .Add(Adr_TreKey, "TreKey", 0, False)
        Set RpCol = .Add(Adr_Grafik, "Grafik", 0, False)
        Set RpCol = .Add(Adr_GuiID, "GuiID", 0, False)
        Set RpCol = .Add(Adr_Objekt, "Objekt", 0, False)
        Set RpCol = .Add(Adr_IDP, "Mandant", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Adr_Mandant, "Nr.", 0, True)
        Set RpCol = .Add(Adr_VIP, "VIP", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Adr_Titel, "Titel", 0, False)
        Set RpCol = .Add(Adr_Land, "Land", 0, False)
        Set RpCol = .Add(Adr_Behindert, "Behindert", 0, False)
        Set RpCol = .Add(Adr_Passiv, "Passiv", 0, False)
        Set RpCol = .Add(Adr_Gruppen, "Gruppen", 0, True)
        Set RpCol = .Add(Adr_Versand, "V", 0, True)
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
        
    RpCls(Adr_ID0).Width = 0
    RpCls(Adr_ID3).Width = 0
    RpCls(Adr_IDKurz).Width = 220
    If GlTFt.SIZE > 10 Then
        RpCls(Adr_Geboren).Width = 110
    Else
        RpCls(Adr_Geboren).Width = 80
    End If
    RpCls(Adr_Name).Width = 100
    RpCls(Adr_Vorname).Width = 100
    RpCls(Adr_Straße).Width = 120
    RpCls(Adr_PLZ).Width = 60
    RpCls(Adr_Ort).Width = 100
    RpCls(Adr_Firma1).Width = 150
    RpCls(Adr_Telefon1).Width = 90
    RpCls(Adr_Telefon2).Width = 90
    RpCls(Adr_Telefon3).Width = 90
    RpCls(Adr_Telefon4).Width = 90
    RpCls(Adr_Telefon5).Width = 120
    RpCls(Adr_Geschlecht).Width = 80
    RpCls(Adr_Datum).Width = 0
    RpCls(Adr_Briefanrede).Width = 0
    RpCls(Adr_Anschrift).Width = 0
    RpCls(Adr_TreKey).Width = 0
    RpCls(Adr_Grafik).Width = 0
    RpCls(Adr_GuiID).Width = 0
    RpCls(Adr_Objekt).Width = 0
    RpCls(Adr_IDP).Width = 0
    RpCls(Adr_Mandant).Width = 50
    RpCls(Adr_VIP).Width = 0
    RpCls(Adr_Titel).Width = 0
    RpCls(Adr_Land).Width = 0
    RpCls(Adr_Behindert).Width = 0
    RpCls(Adr_Passiv).Width = 0
    RpCls(Adr_Gruppen).Width = 150
    RpCls(Adr_Versand).Width = 20
Else
    With RpCls
        Set RpCol = .Add(Rec_ID1, "ID1", 0, False)
        Set RpCol = .Add(Rec_ID0, "ID0", 0, False)
        Set RpCol = .Add(Rec_RechNr, "Rechnung", 0, True)
        Set RpCol = .Add(Rec_Datum, "Datum", 0, True)
        RpCol.Groupable = False
        Set RpCol = .Add(Rec_Selekt, "Abgeschlossen", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Type, "T", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Versand, "V", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Betrag, "Betrag", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Bezahlt, "Bezahlt", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Differe, "Offen", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_IDKurz, "Patient", 0, True)
        Set RpCol = .Add(Rec_Offen, "B", 0, False)
        With RpCol
            .Alignment = xtpAlignmentIconCenter
            .HeaderAlignment = xtpAlignmentCenter
            .Tag = 1
        End With
        Set RpCol = .Add(Rec_Extrageb, "Extrageb.", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Rec_Fallig, "Fälligkeit", 0, True)
        Set RpCol = .Add(Rec_Wahrung, "Währung", 0, False)
        Set RpCol = .Add(Rec_IDR, "Zähler", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_ID3, "ID3", 0, False)
        Set RpCol = .Add(Rec_IDZ, "IDZ", 0, False)
        Set RpCol = .Add(Rec_Versicherer, "Katalog", 0, True)
        Set RpCol = .Add(Rec_Zahlziel, "Zahlungsziel", 0, True)
        Set RpCol = .Add(Rec_Drucken, "Drucken", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_IDW, "IDW", 0, False)
        Set RpCol = .Add(Rec_Symbol, "W", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Faktor, "Faktor", 0, False)
        Set RpCol = .Add(Rec_Ziel, "Ziel", 0, False)
        Set RpCol = .Add(Rec_Kommentar, "Kommentar", 0, False)
        Set RpCol = .Add(Rec_IDP, "Mandant", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Rec_Druckdatum, "Gedruckt", 0, True)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Kopie, "Kopie", 0, False)
        Set RpCol = .Add(Rec_Steuer, "Steuer", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
        End With
        Set RpCol = .Add(Rec_Monat, "Monat", 0, True)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            .EditOptions.Constraints.Add "Januar", 1
            .EditOptions.Constraints.Add "Februar", 2
            .EditOptions.Constraints.Add "März", 3
            .EditOptions.Constraints.Add "April", 4
            .EditOptions.Constraints.Add "Mai", 5
            .EditOptions.Constraints.Add "Juni", 6
            .EditOptions.Constraints.Add "Juli", 7
            .EditOptions.Constraints.Add "August", 8
            .EditOptions.Constraints.Add "September", 9
            .EditOptions.Constraints.Add "Oktober", 10
            .EditOptions.Constraints.Add "November", 11
            .EditOptions.Constraints.Add "Dezember", 12
        End With
        Set RpCol = .Add(Rec_Termin, "Termins.", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Storniert, "Storniert", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_PKU, "PKU", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Gruppe, "G", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Beendet, "E", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Rabatt, "Rabatt", 0, False)
        Set RpCol = .Add(Rec_IDM, "Mitarbeiter", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Rec_GuStr, "Gutschrift", 0, False)
        Set RpCol = .Add(Rec_GutNr, "GutNr", 0, False)
        Set RpCol = .Add(Rec_GuiID, "GuiID", 0, False)
        Set RpCol = .Add(Rec_AufNr, "AufNr", 0, False)
        Set RpCol = .Add(Rec_AuStr, "Auftrag", 0, False)
        Set RpCol = .Add(Rec_Formu, "Formular", 0, False)
        Set RpCol = .Add(Rec_OPLoe, "OPL", 0, False)
        RpCol.Alignment = xtpAlignmentIconLeft
        RpCol.Icon = IC16_Pin_Green
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Lock, "Lock", 0, False)
        RpCol.Alignment = xtpAlignmentIconLeft
        RpCol.Icon = IC16_Lock
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_IDO, "IDO", 0, False)
        Set RpCol = .Add(Rec_RzDat, "RzDat", 0, False)
        Set RpCol = .Add(Rec_RzNum, "RzNum", 0, False)
        Set RpCol = .Add(Rec_RzTex, "RzTex", 0, False)
        Set RpCol = .Add(Rec_Grund, "Grund", 0, False)
        Set RpCol = .Add(Rec_ForID, "FID", 0, False)
    End With
    
    For Each RpCol In RpCls
        With RpCol
            .Editable = False
            .Groupable = True
            .Sortable = True
            .AutoSize = False
            .AutoSortWhenGrouped = False
        End With
    Next RpCol
    
    RpCls(Rec_ID1).Width = 0
    RpCls(Rec_ID0).Width = 0
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_RechNr).Width = 140
        RpCls(Rec_Datum).Width = 110
    Else
        RpCls(Rec_RechNr).Width = 110
        RpCls(Rec_Datum).Width = 80
    End If
    RpCls(Rec_Selekt).Width = 0
    RpCls(Rec_Type).Width = 20
    RpCls(Rec_Versand).Width = 20
    RpCls(Rec_Betrag).Width = 75
    RpCls(Rec_Bezahlt).Width = 75
    RpCls(Rec_Differe).Width = 75
    RpCls(Rec_IDKurz).Width = 220
    RpCls(Rec_Offen).Width = 0
    RpCls(Rec_Extrageb).Width = 75
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_Fallig).Width = 110
    Else
        RpCls(Rec_Fallig).Width = 80
    End If
    RpCls(Rec_Wahrung).Width = 0
    RpCls(Rec_IDR).Width = 60
    RpCls(Rec_ID3).Width = 0
    RpCls(Rec_IDZ).Width = 0
    RpCls(Rec_Versicherer).Width = 140
    RpCls(Rec_Zahlziel).Width = 140
    RpCls(Rec_Drucken).Width = 0
    RpCls(Rec_IDW).Width = 0
    RpCls(Rec_Symbol).Width = 30
    RpCls(Rec_Faktor).Width = 0
    RpCls(Rec_Ziel).Width = 0
    RpCls(Rec_Kommentar).Width = 0
    RpCls(Rec_IDP).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_Druckdatum).Width = 110
    Else
        RpCls(Rec_Druckdatum).Width = 80
    End If
    RpCls(Rec_Kopie).Width = 0
    RpCls(Rec_Steuer).Width = 60
    RpCls(Rec_Monat).Width = 0
    RpCls(Rec_Termin).Width = 75
    RpCls(Rec_Storniert).Width = 0
    RpCls(Rec_PKU).Width = 50
    RpCls(Rec_Beendet).Width = 0
    RpCls(Rec_Rabatt).Width = 0
    RpCls(Rec_IDM).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_GuStr).Width = 110
    Else
        RpCls(Rec_GuStr).Width = 80
    End If
    RpCls(Rec_GutNr).Width = 0
    RpCls(Rec_AufNr).Width = 0
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_AuStr).Width = 110
    Else
        RpCls(Rec_AuStr).Width = 80
    End If
    RpCls(Rec_Formu).Width = 120
    RpCls(Rec_OPLoe).Width = 18
    RpCls(Rec_Lock).Width = 18
    DoEvents
    
    Set RpCls = RpCon.Columns
    Set RpCol = RpCls.Find(Rec_IDP)
    For AktZa = 1 To UBound(GlMan)
        RpCol.EditOptions.Constraints.Add GlMan(AktZa, 1), GlMan(AktZa, 2)
    Next AktZa
    Set RpCol = RpCls.Find(Rec_IDM)
    For AktZa = 1 To UBound(GlMiK)
        RpCol.EditOptions.Constraints.Add GlMiK(AktZa, 1), GlMiK(AktZa, 2)
    Next AktZa
    RpCon.Redraw
End If

Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpla " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo SuErr

Dim GesZa As Long
Dim Mld1, Tit1 As String
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set RpCon = Me.repCont1
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set RpCon = Me.repCont1
Set Opti1 = Me.optPatie
Set Opti2 = Me.optRechn
Set Opti3 = Me.optAufhe
Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumm
Set FTex3 = Me.txtPost
Set Lab02 = Me.lblLab02
Set Lab03 = Me.lblLab03
Set lab04 = Me.lblLab04
Set Lab06 = Me.lblLab06
Set PuBu3 = Me.btnZuruk
Set ImMan = FM.imgManag
Set RpCls = RpCon.Columns
Set RpRcs = RpCon.Records

Tit1 = "Eintrag nicht gefunden"
Mld1 = "Der gewünschte Eintrag kann nicht gefunden werden"

If Rahm1.Visible = True Then
    If Opti1.Value = True Then
        Lab03.Caption = "Suche nach Patientennummer :"
        lab04.Caption = "Suche nach Postleitzahl :"
    ElseIf Opti2.Value = True Then
        Lab03.Caption = "Suche nach Rechnungsnummer :"
        lab04.Caption = "Suche nach Rechnungsdatum :"
    Else
        S_BuZu5
        DoEvents
        Unload Me
        Exit Sub
    End If
    If FBuZu = False Then
        PuBu3.Enabled = True
        Lab06.Caption = vbNullString
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        DoEvents
        FTex1.SetFocus
    End If
ElseIf Rahm2.Visible = True Then
    If Opti1.Value = True Then
        SuTyp = 1
    ElseIf Opti2.Value = True Then
        SuTyp = 2
    ElseIf Opti3.Value = True Then
        SuTyp = 3
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    FSpla
    DoEvents
    If FTex1.Text <> vbNullString Then
        GesZa = S_BuZu1(FTex1.Text, 1, SuTyp)
    ElseIf FTex2.Text <> vbNullString Then
        GesZa = S_BuZu1(FTex2.Text, 2, SuTyp)
    ElseIf FTex3.Text <> vbNullString Then
        GesZa = S_BuZu1(FTex3.Text, 3, SuTyp)
    End If
    If GesZa > 0 Then
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = True
        DoEvents
        Set RpRws = RpCon.Rows
        RpRws.Row(0).Selected = True
        RpCon.SetFocus
        DoEvents
        GlAkt = False 'WICHTIG!
    Else
        FRes
        SPopu Tit1, Mld1, IC48_Information
    End If
    Screen.MousePointer = vbNormal
ElseIf Rahm3.Visible = True Then
    If GlAkt = False Then
        S_BuZu4 SuTyp
        DoEvents
        Unload Me
        Exit Sub
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub FZuru()
On Error Resume Next

Set Opti1 = Me.optPatie
Set Opti2 = Me.optRechn
Set Opti3 = Me.optAufhe
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumm
Set FTex3 = Me.txtPost
Set Lab02 = Me.lblLab02
Set Lab03 = Me.lblLab03
Set lab04 = Me.lblLab04
Set Lab06 = Me.lblLab06
Set PuBu3 = Me.btnZuruk

FTex1.Text = vbNullString
FTex2.Text = vbNullString
FTex3.Text = vbNullString
Lab06.Caption = vbNullString

If Rahm2.Visible = True Then
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = True
    PuBu3.Enabled = False
ElseIf Rahm3.Visible = True Then
    Rahm3.Visible = False
    Rahm2.Visible = True
    Rahm1.Visible = False
    PuBu3.Enabled = True
End If

End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50161)
TeMai = IniGetOpt("Hilfe", 50162)
TeInh = IniGetOpt("Hilfe", 50163)
TeFus = IniGetOpt("Hilfe", 50164)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub

Private Sub btnZuruk_Click()
    FZuru
End Sub
Private Sub Form_Load()
On Error Resume Next

FKonf
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub optAufhe_Click()
    SuTyp = 3
End Sub
Private Sub optPatie_Click()
    SuTyp = 1
End Sub
Private Sub optRechn_Click()
    SuTyp = 2
End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If SuTyp = 2 Then
    If Row.GroupRow = False Then
        Select Case Row.Record(Rec_Type).Value
        Case "M": Metrics.ForeColor = 16744448
        Case "L": Metrics.ForeColor = 33023
        Case "V": Metrics.ForeColor = 8421631
        Case "I": Metrics.ForeColor = 13138080
        Case "U": Metrics.ForeColor = 6604830
        Case Else:
            If CBool(Row.Record(Rec_Selekt).Value) = False Then
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
        End Select
        If CBool(Row.Record(Rec_Selekt).Value) = False Then
            Metrics.Font.Bold = True
        End If
        If Row.Record(Rec_Storniert).Value = True Then
            Metrics.Font.Strikethrough = True
            Metrics.ForeColor = 8421504
        End If
    End If
End If

End Sub
Private Sub repCont1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim RpCon As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCon = Me.repCont1
Set RpRws = RpCon.Rows

If GlAkt = False Then
    If RpRws.Count > 0 Then
        If KeyCode = vbKeyReturn Then
            S_BuZu4 SuTyp
            Unload Me
        End If
    End If
End If

Set RpRws = Nothing
Set RpCon = Nothing

End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim RpCon As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCon = Me.repCont1
Set RpRws = RpCon.Rows

If GlAkt = False Then
    If RpRws.Count > 0 Then
        If Row.GroupRow = False Then
            S_BuZu4 SuTyp
            Unload Me
        End If
    End If
End If

Set RpRws = Nothing
Set RpCon = Nothing

End Sub
Private Sub txtKurz_GotFocus()
    FRes
End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FWeit
    End If
End Sub
Private Sub txtNumm_GotFocus()
    FRes
End Sub

Private Sub txtNumm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FWeit
    End If
End Sub
Private Sub txtPost_GotFocus()
    FRes
End Sub

Private Sub txtPost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FWeit
    End If
End Sub
