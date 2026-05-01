VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmAdrPost 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Postleitzahlen"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   2540
      Left            =   30
      TabIndex        =   0
      Top             =   900
      Width           =   6860
      _Version        =   1048579
      _ExtentX        =   12100
      _ExtentY        =   4480
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   6
      Top             =   3900
      Width           =   7000
      _Version        =   1048579
      _ExtentX        =   12347
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
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
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Height          =   400
         Left            =   3600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Einfügen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
   End
   Begin XtremeSuiteControls.CheckBox chkTeOrt 
      Height          =   220
      Left            =   400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2000
      _Version        =   1048579
      _ExtentX        =   3528
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Ortsteil mit einfügen"
      UseVisualStyle  =   -1  'True
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
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   300
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6000
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   1
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
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   225
      Left            =   300
      TabIndex        =   8
      Top             =   330
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Suche nach PLZ :"
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label lblLab02 
      Height          =   225
      Left            =   3000
      TabIndex        =   7
      Top             =   330
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Suche nach Ort :"
      Alignment       =   1
   End
End
Attribute VB_Name = "frmAdrPost"
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
Private TxPLZ As XtremeSuiteControls.FlatEdit
Private TxOrt As XtremeSuiteControls.FlatEdit
Private Lbl01 As XtremeSuiteControls.Label
Private Lbl02 As XtremeSuiteControls.Label
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private clFen As clsFenster

Public FeTyp As Integer
Private Sub FKonf()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmAdrPost
Set Rahm0 = FM.frmRahm0
Set RpCon = FM.repCont1
Set ChOrt = FM.chkTeOrt
Set TxPLZ = FM.txtSuPLZ
Set TxOrt = FM.txtSuOrt
Set Lbl01 = FM.lblLab01
Set Lbl02 = FM.lblLab02
Set RpCls = RpCon.Columns
Set ImMan = frmMain.imgManag

ChOrt.Value = CBool(IniGetVal("System", "TeiOrt"))

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
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
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

With RpCls
    Set RpCol = .Add(0, "PLZ", 100, True)
    Set RpCol = .Add(1, "Hauptort", 200, True)
    Set RpCol = .Add(2, "Ortsteil", 200, True)
    RpCol.AutoSize = True
    Set RpCol = .Add(3, "Vorwahl", 100, True)
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
Private Sub FOrt()
On Error GoTo InErr

Set FM = frmAdrPost
Set ChOrt = FM.chkTeOrt

If ChOrt.Value = xtpChecked Then
    IniSetVal "System", "TeiOrt", -1
Else
    IniSetVal "System", "TeiOrt", 0
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FRes()
On Error Resume Next

Set FM = frmAdrPost
Set TxPLZ = FM.txtSuPLZ
Set TxOrt = FM.txtSuOrt

TxPLZ.Text = vbNullString
TxOrt.Text = vbNullString

End Sub

Private Sub FSett()
On Error GoTo SeErr
'Lädt die ausgewählte Adresse in das Adressformular

Dim StPLz As String
Dim StHor As String
Dim StOrt As String
Dim StVor As String
Dim FePos As XtremeSuiteControls.FlatEdit
Dim FeOrt As XtremeSuiteControls.FlatEdit
Dim FeTe1 As XtremeSuiteControls.FlatEdit
Dim FeTe2 As XtremeSuiteControls.FlatEdit
Dim FeTe3 As XtremeSuiteControls.FlatEdit
Dim FeBer As XtremeSuiteControls.FlatEdit
Dim FeLa2 As XtremeSuiteControls.FlatEdit
Dim FeLa1 As XtremeSuiteControls.ComboBox
Dim FeLa3 As XtremeSuiteControls.ComboBox
Dim RpCon As XtremeReportControl.ReportControl

Set RpCon = Me.repCont1
Set ChOrt = Me.chkTeOrt
Set RpSel = RpCon.SelectedRows

Select Case FeTyp
Case 1:
    Set FM = frmAdress
    Set FePos = FM.txtS1F08
    Set FeOrt = FM.txtS1F09
    Set FeLa1 = FM.txtS1F12
    Set FeLa2 = FM.txtS2F22
    Set FeTe1 = FM.txtS1F15
    Set FeTe2 = FM.txtS1F16
    Set FeTe3 = FM.txtS1F17
Case 2:
    Set FM = frmAdress
    Set FePos = FM.txtS2F18
    Set FeOrt = FM.txtS2F19
    Set FeLa1 = FM.txtS1F12
    Set FeLa2 = FM.txtS2F22
    Set FeTe1 = FM.txtS1F15
    Set FeTe2 = FM.txtS1F16
    Set FeTe3 = FM.txtS1F17
Case 3:
    Set FM = frmAdress
    Set FePos = FM.txtS4F08
    Set FeOrt = FM.txtS4F09
    Set FeLa3 = FM.cmbS4F12
    Set FeTe1 = FM.txtS4F15
Case 4:
    Set FM = frmMandant
    Set FePos = FM.txtS1F08
    Set FeOrt = FM.txtS1F09
    Set FeBer = FM.txtS2F24
    Set FeLa1 = FM.txtS1F12
    Set FeTe2 = FM.txtS1F16
    Set FeTe3 = FM.txtS1F17
Case 5:
    Set FM = frmTermin
    Set FePos = FM.txtS4F08
    Set FeOrt = FM.txtS4F09
    Set FeLa3 = FM.cmbS4F12
    Set FeTe1 = FM.txtS4F15
End Select

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        StPLz = Trim$(RpRow.Record(0).Value)
        StHor = Trim$(RpRow.Record(1).Value)
        StVor = Trim$(RpRow.Record(3).Value)
        If ChOrt.Value = xtpChecked Then
            If RpRow.Record(1).Value <> RpRow.Record(2).Value Then
                StOrt = Trim$(RpRow.Record(1).Value) & ", " & Trim$(RpRow.Record(2).Value)
            Else
                StOrt = Trim$(RpRow.Record(1).Value)
            End If
        Else
            StOrt = Trim$(RpRow.Record(1).Value)
        End If
        FePos.Text = StPLz
        FeOrt.Text = StOrt

        Select Case FeTyp
        Case 1:
            If GlInt = True Then 'Internationale Telefonnummer
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = "+49 " & StVor & Chr$(32)
                If Len(FeTe2.Text) < 7 Then FeTe2.Text = "+49 " & StVor & Chr$(32)
                If Len(FeTe3.Text) < 7 Then FeTe3.Text = "+49 " & StVor & Chr$(32)
            Else
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = StVor & Chr$(32)
                If Len(FeTe2.Text) < 7 Then FeTe2.Text = StVor & Chr$(32)
                If Len(FeTe3.Text) < 7 Then FeTe3.Text = StVor & Chr$(32)
            End If
        Case 2:
            If GlInt = True Then 'Internationale Telefonnummer
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = "+49 " & StVor & Chr$(32)
                If Len(FeTe2.Text) < 7 Then FeTe2.Text = "+49 " & StVor & Chr$(32)
                If Len(FeTe3.Text) < 7 Then FeTe3.Text = "+49 " & StVor & Chr$(32)
            Else
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = StVor & Chr$(32)
                If Len(FeTe2.Text) < 7 Then FeTe2.Text = StVor & Chr$(32)
                If Len(FeTe3.Text) < 7 Then FeTe3.Text = StVor & Chr$(32)
            End If
        Case 3:
            If GlInt = True Then 'Internationale Telefonnummer
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = "+49 " & StVor & Chr$(32)
            Else
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = StVor & Chr$(32)
            End If
        Case 4:
            If GlInt = True Then 'Internationale Telefonnummer
                If Len(FeTe2.Text) < 7 Then FeTe2.Text = "+49 " & StVor & Chr$(32)
                If Len(FeTe3.Text) < 7 Then FeTe3.Text = "+49 " & StVor & Chr$(32)
            Else
                If Len(FeTe2.Text) < 7 Then FeTe2.Text = StVor & Chr$(32)
                If Len(FeTe3.Text) < 7 Then FeTe3.Text = StVor & Chr$(32)
            End If
        Case 5:
            If GlInt = True Then 'Internationale Telefonnummer
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = "+49 " & StVor & Chr$(32)
            Else
                If Len(FeTe1.Text) < 7 Then FeTe1.Text = StVor & Chr$(32)
            End If
        End Select

        DoEvents
        Unload Me

        Select Case FeTyp
        Case 1:
            AKopi
            FeLa1.SetFocus
        Case 2:
            AKopi
            FeLa2.SetFocus
        Case 3:
            FeLa3.SetFocus
        Case 4:
            MKopi
            FeBer.SetFocus
        Case 5:
            FeLa3.SetFocus
        End Select

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
    FOrt
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FKonf
AFont Me
FrmEx.TopMost = True
SFrame 1, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAdrPost = Nothing
End Sub
Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(1).Value = Row.Record(2).Value Then Metrics.Font.Bold = True
End Sub
Private Sub repCont1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FSett
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
