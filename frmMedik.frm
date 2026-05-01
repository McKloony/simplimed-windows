VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Object = "{621DDB00-A516-11E8-A658-0013D350667C}#3.2#0"; "tx4ole26.ocx"
Begin VB.Form frmMedik 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ABDATA® Heilmitteldatenbank"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   34
      Top             =   6600
      Width           =   12000
      _Version        =   1048579
      _ExtentX        =   21167
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnBut04 
         Height          =   400
         Left            =   10000
         TabIndex        =   35
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
      Begin XtremeSuiteControls.PushButton btnBut03 
         Height          =   400
         Left            =   8600
         TabIndex        =   36
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
      Begin XtremeSuiteControls.PushButton btnBut02 
         Height          =   400
         Left            =   7200
         TabIndex        =   37
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
      Begin XtremeSuiteControls.PushButton btnBut01 
         Height          =   400
         Left            =   5900
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Detailinfo"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   8000
      Width           =   80
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   6500
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11900
      _Version        =   1048579
      _ExtentX        =   20990
      _ExtentY        =   11465
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkOpt02 
         Height          =   220
         Left            =   3010
         TabIndex        =   7
         Top             =   5800
         Visible         =   0   'False
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7056
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Als Krankenblatt Aufnahme-Medikamente einfügen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkOpt01 
         Height          =   220
         Left            =   3010
         TabIndex        =   6
         Top             =   5400
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7056
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Ausgewählte Produkte erst in Sammelliste einfügen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch4 
         Height          =   350
         Left            =   3000
         TabIndex        =   4
         Top             =   3900
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch3 
         Height          =   350
         Left            =   3000
         TabIndex        =   3
         Top             =   3100
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch2 
         Height          =   350
         Left            =   3000
         TabIndex        =   2
         Top             =   2300
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch1 
         Height          =   350
         Left            =   3000
         TabIndex        =   1
         Top             =   1500
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch5 
         Height          =   350
         Left            =   3000
         TabIndex        =   5
         Top             =   4700
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab12 
         Height          =   220
         Left            =   3005
         TabIndex        =   22
         Top             =   2850
         Width           =   3195
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach Hersteller / Anbieter :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab13 
         Height          =   220
         Left            =   3005
         TabIndex        =   21
         Top             =   3650
         Width           =   3195
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach Wirkstoff :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab14 
         Height          =   220
         Left            =   3005
         TabIndex        =   20
         Top             =   4450
         Width           =   3195
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach Indikation :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab11 
         Height          =   220
         Left            =   3005
         TabIndex        =   19
         Top             =   2050
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach PZN (Pharmazentralnummer) :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab10 
         Height          =   220
         Left            =   3005
         TabIndex        =   18
         Top             =   1250
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach Produktbezeichnung :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   400
         Left            =   300
         TabIndex        =   15
         Top             =   200
         Width           =   11300
         _Version        =   1048579
         _ExtentX        =   19932
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   $"frmMedik.frx":0000
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H80000005&
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   11900
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   6500
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   11900
      _Version        =   1048579
      _ExtentX        =   20990
      _ExtentY        =   11465
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont2 
         Height          =   5680
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   800
         Width           =   11880
         _Version        =   1048579
         _ExtentX        =   20955
         _ExtentY        =   10019
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   400
         Left            =   300
         TabIndex        =   16
         Top             =   200
         Width           =   11300
         _Version        =   1048579
         _ExtentX        =   19932
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   $"frmMedik.frx":0098
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H80000005&
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   11900
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   6500
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   11900
      _Version        =   1048579
      _ExtentX        =   20990
      _ExtentY        =   11465
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   5680
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   800
         Width           =   11880
         _Version        =   1048579
         _ExtentX        =   20955
         _ExtentY        =   10019
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   400
         Left            =   300
         TabIndex        =   17
         Top             =   200
         Width           =   11300
         _Version        =   1048579
         _ExtentX        =   19932
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   $"frmMedik.frx":0134
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H80000005&
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   11900
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   6500
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   11900
      _Version        =   1048579
      _ExtentX        =   20990
      _ExtentY        =   11465
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont3 
         Height          =   5680
         Left            =   0
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   800
         Width           =   11880
         _Version        =   1048579
         _ExtentX        =   20955
         _ExtentY        =   10019
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   400
         Left            =   300
         TabIndex        =   24
         Top             =   200
         Width           =   11300
         _Version        =   1048579
         _ExtentX        =   19932
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   $"frmMedik.frx":0239
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H80000005&
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   11900
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   6500
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   11900
      _Version        =   1048579
      _ExtentX        =   20990
      _ExtentY        =   11465
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin Tx4oleLib.TXTextControl TexCont4 
         Height          =   5680
         Left            =   10
         TabIndex        =   33
         Top             =   800
         Width           =   11920
         _Version        =   196610
         _ExtentX        =   21026
         _ExtentY        =   10019
         _StockProps     =   73
         BackColor       =   -2147483643
         Language        =   49
         BorderStyle     =   0
         BackStyle       =   1
         ControlChars    =   0   'False
         EditMode        =   0
         HideSelection   =   -1  'True
         InsertionMode   =   -1  'True
         MousePointer    =   0
         ZoomFactor      =   100
         ViewMode        =   3
         ClipChildren    =   0   'False
         ClipSiblings    =   -1  'True
         SizeMode        =   0
         TabKey          =   -1  'True
         FormatSelection =   0   'False
         VTSpellDictionary=   "C:\PROGRA~1\TEXTCO~1\TXTEXT~1.0AC\Bin\AMERICAN.VTD"
         ScrollBars      =   3
         PageWidth       =   12240
         PageHeight      =   15840
         PageMarginL     =   1440
         PageMarginT     =   1440
         PageMarginR     =   1440
         PageMarginB     =   1440
         PrintZoom       =   100
         PrintOffset     =   0   'False
         PrintColors     =   -1  'True
         FontName        =   "Arial"
         FontSize        =   12
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Baseline        =   0
         TextBkColor     =   -2147483643
         Alignment       =   0
         LineSpacing     =   100
         LineSpacingT    =   0
         FrameStyle      =   32
         FrameDistance   =   0
         FrameLineWidth  =   20
         IndentL         =   0
         IndentR         =   0
         IndentFL        =   0
         IndentT         =   0
         IndentB         =   0
         Text            =   ""
         WordWrapMode    =   1
         AllowUndo       =   -1  'True
         TextFrameMarkerLines=   -1  'True
         FieldLinkTargetMarkers=   0   'False
         PageOrientation =   0
         PageViewStyle   =   1
         FontSettings    =   0
         AllowDrag       =   0   'False
         AllowDrop       =   0   'False
         SelectionViewMode=   1
         SectionRestartPageNumbering=   0
         PermanentControlChars=   16
         RightToLeft     =   0   'False
         TextDirection   =   2
         Locale          =   1031
         Justification   =   1
         FrameColor      =   16777215
         FrameLineColor  =   0
         DocumentPermissions=   31
         SelectObjects   =   -1  'True
         IsTrackChangesEnabled=   0   'False
         IsFormulaCalculationEnabled=   -1  'True
         FormulaReferenceStyle=   0
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   400
         Left            =   300
         TabIndex        =   26
         Top             =   200
         Width           =   11300
         _Version        =   1048579
         _ExtentX        =   19932
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   $"frmMedik.frx":0320
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H80000005&
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   11900
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm6 
      Height          =   6500
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   11900
      _Version        =   1048579
      _ExtentX        =   20990
      _ExtentY        =   11465
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ProgressBar prbStat2 
         Height          =   360
         Left            =   1500
         TabIndex        =   30
         Top             =   3500
         Width           =   8000
         _Version        =   1048579
         _ExtentX        =   14111
         _ExtentY        =   635
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.ProgressBar prbStat1 
         Height          =   360
         Left            =   1500
         TabIndex        =   29
         Top             =   2500
         Width           =   8000
         _Version        =   1048579
         _ExtentX        =   14111
         _ExtentY        =   635
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.Label lblLab09 
         Height          =   220
         Left            =   1520
         TabIndex        =   32
         Top             =   3250
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7056
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Produkt :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   220
         Left            =   1520
         TabIndex        =   31
         Top             =   2250
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7056
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Katalog :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   400
         Left            =   300
         TabIndex        =   28
         Top             =   200
         Width           =   11300
         _Version        =   1048579
         _ExtentX        =   19932
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   $"frmMedik.frx":03FD
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H80000005&
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   11900
      End
   End
   Begin XtremeSuiteControls.Label lblLab04 
      Height          =   220
      Left            =   3100
      TabIndex        =   23
      Top             =   6940
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   388
      _StockProps     =   79
      Alignment       =   4
   End
End
Attribute VB_Name = "frmMedik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Lbl04 As XtremeSuiteControls.Label
Private TxSu1 As XtremeSuiteControls.FlatEdit
Private TxSu2 As XtremeSuiteControls.FlatEdit
Private TxSu3 As XtremeSuiteControls.FlatEdit
Private TxSu4 As XtremeSuiteControls.FlatEdit
Private TxSu5 As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Chk01 As XtremeSuiteControls.CheckBox
Private Chk02 As XtremeSuiteControls.CheckBox
Private But01 As XtremeSuiteControls.PushButton
Private But02 As XtremeSuiteControls.PushButton
Private But03 As XtremeSuiteControls.PushButton
Private But04 As XtremeSuiteControls.PushButton
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
Private TxCoN As Tx4oleLib.TXTextControl

Private SaLis As Boolean
Private KraMe As Boolean
Private FoLad As Boolean
Private SuFel As Integer
Private Sub FDet()
On Error GoTo InErr

Dim RowNr As Long
Dim SuVal As String
Dim RpCo1 As XtremeReportControl.ReportControl

Set FM = frmMedik
Set TxCoN = FM.TexCont4
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set But01 = FM.btnBut01
Set But02 = FM.btnBut02
Set But03 = FM.btnBut03
Set But04 = FM.btnBut04
Set RpCo1 = FM.repCont1
Set RpSel = RpCo1.SelectedRows

If Rahm3.Visible = True Then
    TxCoN.ResetContents
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    Rahm6.Visible = False
    But03.Enabled = False
    But01.Caption = "&PDF-Dokument"
    If SaLis = True Then
        But04.Caption = "&Schließen"
    End If
    DoEvents
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            SuVal = AbAry(ABDA_PZN, RpRow.Index)
            ABD_Det SuVal
        End If
    End If
ElseIf Rahm5.Visible = True Then
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            SuVal = AbAry(ABDA_PZN, RpRow.Index)
            ABD_Dok SuVal
        End If
    End If
End If

Set RpCo1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDet " & Err.Number
Resume Next

End Sub
Private Sub FEinf()
On Error GoTo InErr

Dim RowNr As Long
Dim AktPo As Long
Dim GesPo As Long
Dim AkSpa As Integer
Dim RpCo1 As XtremeReportControl.ReportControl

Set FM = frmMedik
Set RpCo1 = FM.repCont1
Set RpSel = RpCo1.SelectedRows

GesPo = RpSel.Count

Screen.MousePointer = vbHourglass

If GesPo > 0 Then
    ReDim GlClp(GesPo, 17) 'WICHTIG!
    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            For AkSpa = 0 To 16 'WICHTIG!
                If AbAry(AkSpa, RpRow.Index) <> vbNullString Then
                    GlClp(AktPo, AkSpa) = AbAry(AkSpa, RpRow.Index)
                Else
                    GlClp(AktPo, AkSpa) = vbNullString
                End If
            Next AkSpa
            AktPo = AktPo + 1
        End If
    Next RpRow
End If

Set RpCo1 = Nothing

Select Case GlBut
Case RibTab_Krankenbla: ABD_Kra GesPo, KraMe
Case RibTab_Abrechnung: ABD_Abr GesPo
Case RibTab_Rezeptmodul: ABD_Rez GesPo
Case RibTab_Belegmodul: ABD_Rez GesPo
Case RibTab_Tex_Rezept: ABD_Trz GesPo
Case RibTab_Kat_Eintrg: ABD_Kat GesPo
End Select

Screen.MousePointer = vbNormal

Unload FM

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEinf " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim RetWe As Long
Dim LiTip As Boolean
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl

Set FM = frmMedik
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set But01 = FM.btnBut01
Set But02 = FM.btnBut02
Set But03 = FM.btnBut03
Set But04 = FM.btnBut04
Set Chk01 = FM.chkOpt01
Set Chk02 = FM.chkOpt02
Set Lbl04 = FM.lblLab04
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set TxCoN = FM.TexCont4
Set ImMan = frmMain.imgManag

SaLis = CBool(IniGetVal("System", "AbdSam"))
KraMe = CBool(IniGetVal("System", "AbdKra"))
LiTip = CBool(IniGetVal("Layout", "GrdTip"))

With TxCoN
    .ViewMode = 2
    .Alignment = 0
    .AllowDrop = True
    .AllowUndo = True
    .Enabled = True
    .DataTextFormat = 0
    .AutoExpand = False
    .ClipChildren = False
    .ClipSiblings = False
    .ControlChars = False
    .ColumnLineColor = 0
    .Columns = 2 'WICHTIG!
    .BackColor = -2147483643 '16777215
    .BackStyle = 1
    .BaseLine = 2
    .BorderStyle = 0
    .EditMode = 0
    .FontBold = GlXFt.Bold
    .FontItalic = GlXFt.Italic
    .FontUnderline = GlXFt.Underline
    .FontStrikethru = GlXFt.Strikethrough
    If GlViW <> 2 Then .FontName = GlXFt.Name
    .FontSize = 10
    .ForeColor = vbBlack
    .FormatSelection = True
    .HeaderFooterStyle = txDividingLine + txMouseClick
    .HideSelection = False
    .InsertionMode = True
    .Language = 49
    .LineSpacing = 110
    .PageViewStyle = txGradientColors
    .PageHeight = (500 / 10) * 567
    .PageWidth = (150 / 10) * 567
    .PageMarginL = (10 / 10) * 567
    .PageMarginR = (10 / 10) * 567
    .PageMarginT = (5 / 10) * 567
    .PageMarginB = (5 / 10) * 567
    .PageOrientation = 0
    .PrintColors = True
    .ScrollBars = 3
    .SizeMode = 0
    .SelectionViewMode = 1
    .TabKey = True
    .TextBkColor = 16777215
    .TextFrameMarkerLines = True
    .TableGridLines = True
    .EnableHyperlinks = True
    .ZoomFactor = 100
    .ViewMode = 2
    .WordWrapMode = 1
    .DisplayColor(txDesktopColor) = GlBkk
End With

With RpCo1
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
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = False
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    '.OverrideThemeMetrics = True  no matter what this function do !?
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Produkte vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Produkte vorhanden"
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
    .PaintManager.MaxPreviewLines = 0
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = True
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.ColumnWidthWYSIWYG = False
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .MultiSelectionMode = False
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = GlGKo
    .SelectionEnable = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo2
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
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = False
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    '.OverrideThemeMetrics = True  no matter what this function do !?
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Produkte vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Produkte vorhanden"
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
    .PaintManager.MaxPreviewLines = 0
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = True
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.ColumnWidthWYSIWYG = False
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .MultiSelectionMode = False
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = GlGKo
    .SelectionEnable = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo3
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = False
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .EnableMarkup = False 'XAML
    .FocusSubItems = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FooterRowsAllowAccess = False
    .FooterRowsAllowEdit = False
    .Icons = ImMan.Icons
    .MultipleSelection = True
    .LockExpand = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.AllowMergeCells = True
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
    .PaintManager.ShadeGroupHeadings = True
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = 0
    .PaintManager.ShowNonActiveInPlaceButton = True
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = True
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.ColumnWidthWYSIWYG = False
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.FooterRowsDividerStyle = xtpReportFixedRowsDividerOutlook
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .MultiSelectionMode = False
    .ShowGroupBox = False
    .ShowIconWhenEditing False
    .PreviewMode = GlVoA
    .ShowHeader = GlGKo
    .ShowFooter = False
    .ShowFooterRows = False
    .SortedDragDrop = True
    .SelectionEnable = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

If SaLis = True Then
    Chk01.Value = xtpChecked
End If

If GlBut = RibTab_Krankenbla Then
    If KraMe = True Then
        Chk02.Value = xtpChecked
    End If
Else
    KraMe = False
End If

Select Case GlBut
Case RibTab_Krankenbla:
    Chk02.Caption = "Als Krankenblatt Aufnahme-Arzneimittel einfügen"
    Chk02.Visible = True
Case RibTab_Kat_Eintrg:
    Chk02.Caption = "Arzneikatalog-Abgleich durchführen"
    Chk02.Visible = True
End Select

Lbl04.BackColor = GlBak
Chk01.BackColor = GlBak
Chk02.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Rahm6.BackColor = GlBak
FM.BackColor = GlBak

Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FRes()
On Error Resume Next

Set TxSu1 = Me.txtSuch1
Set TxSu2 = Me.txtSuch2
Set TxSu3 = Me.txtSuch3
Set TxSu4 = Me.txtSuch4
Set TxSu5 = Me.txtSuch5

TxSu1.Text = vbNullString
TxSu2.Text = vbNullString
TxSu3.Text = vbNullString
TxSu4.Text = vbNullString
TxSu5.Text = vbNullString

End Sub
Private Sub FSam()
On Error GoTo InErr

Dim AktPo As Integer
Dim AkSpa As Integer
Dim GesPo As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl

Set FM = frmMedik
Set RpCo1 = FM.repCont1
Set RpCo3 = FM.repCont3
Set RpSel = RpCo1.SelectedRows
Set RpRcs = RpCo3.Records

GesPo = RpSel.Count

Screen.MousePointer = vbHourglass

If GesPo > 0 Then
    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then
            Set RpRec = RpRcs.Add()
            For AkSpa = 0 To 16 'WICHTIG!
                If AbAry(AkSpa, RpRow.Index) <> vbNullString Then
                    Set RpItm = RpRec.AddItem(AbAry(AkSpa, RpRow.Index))
                Else
                    Set RpItm = RpRec.AddItem(vbNullString)
                End If
                RpItm.Focusable = False
            Next AkSpa
            AktPo = AktPo + 1
            For AkSpa = 17 To 22 'WICHTIG!
                Set RpItm = RpRec.AddItem(vbNullString)
            Next AkSpa
            Set RpItm = RpRec.AddItem(vbNullString)
            With RpItm
                .HasCheckbox = True
                .Checked = True
                .Alignment = xtpAlignmentIconCenter
            End With
        End If
    Next RpRow
End If

RpCo3.Populate

Screen.MousePointer = vbNormal

Set RpCo1 = Nothing
Set RpCo3 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSam " & Err.Number
Resume Next

End Sub
Private Sub FSpl1()
On Error GoTo InErr

Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMedik
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCls
    Set RpCol = .Add(ABDA_Sort, "Sorter", 0, False)
    Set RpCol = .Add(ABDA_PZN, "PZN", 90, False)
    Set RpCol = .Add(ABDA_IDLang, "Bezeichnung", 300, True)
    Set RpCol = .Add(ABDA_IDKurz, "Bezeichnung", 0, False)
    Set RpCol = .Add(ABDA_Darreichung, "Darreichung", 150, False)
    Set RpCol = .Add(ABDA_Darreichung_Kurz, "Darreichung_Kurz", 0, False)
    Set RpCol = .Add(ABDA_ApoPflicht, "AP", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_ApoPreis, "AVP", 50, False)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Packungskomponente, "Packung", 0, False)
    Set RpCol = .Add(ABDA_Einstufung, "NP", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Typ, "Typ", 0, False)
    Set RpCol = .Add(ABDA_Menge, "Menge", 60, False)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Einheit, "Einheit", 50, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Wirkstoff, "Wirkstoff", 200, True)
    Set RpCol = .Add(ABDA_Warengruppe, "Warengruppe", 200, True)
    Set RpCol = .Add(ABDA_Vertriebsstatus, "VS", 50, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Firmenname, "Anbietername", 200, True)
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpl1 " & Err.Number
Resume Next

End Sub
Private Sub FSpl2(ByVal TaTyp As Integer)
On Error GoTo InErr

Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMedik
Set RpCo2 = FM.repCont2
Set RpCls = RpCo2.Columns

With RpCo2
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Select Case TaTyp
Case 3:
    With RpCls
        Set RpCol = .Add(ABAN_ID3, "ID3", 0, False)
        Set RpCol = .Add(ABAN_Sorter, "Sorter", 0, False)
        Set RpCol = .Add(ABAN_Firma1, "Firma", 250, True)
        Set RpCol = .Add(ABAN_Firma2, "Zusatz", 140, True)
        Set RpCol = .Add(ABAN_Strasse, "Straße", 150, True)
        Set RpCol = .Add(ABAN_Hausnr_von, "Hausnr_von", 0, False)
        Set RpCol = .Add(ABAN_Hausnr_bis, "Hausnr_bis", 0, False)
        Set RpCol = .Add(ABAN_PLZ, "PLZ", 50, True)
        Set RpCol = .Add(ABAN_Ort, "Ort", 140, True)
        Set RpCol = .Add(ABAN_Land, "Land", 30, True)
    End With
Case 4:
    With RpCls
        Set RpCol = .Add(ABAN_ID3, "ID3", 0, False)
        Set RpCol = .Add(ABAN_Sorter, "Wirkstoff", 750, True)
    End With
Case 5:
    With RpCls
        Set RpCol = .Add(ABAN_ID3, "IDI", 0, False)
        Set RpCol = .Add(ABAN_Sorter, "Indikation", 750, True)
    End With
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpl2 " & Err.Number
Resume Next

End Sub
Private Sub FSpl3()
On Error GoTo InErr

Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMedik
Set RpCo3 = FM.repCont3
Set RpCls = RpCo3.Columns

With RpCls
    Set RpCol = .Add(ABDA_Sort, "Sorter", 0, False)
    Set RpCol = .Add(ABDA_PZN, "PZN", 90, False)
    Set RpCol = .Add(ABDA_IDLang, "Bezeichnung", 300, True)
    Set RpCol = .Add(ABDA_IDKurz, "Bezeichnung", 0, False)
    Set RpCol = .Add(ABDA_Darreichung, "Darreichung", 0, False)
    Set RpCol = .Add(ABDA_Darreichung_Kurz, "Dar.", 60, False)
    Set RpCol = .Add(ABDA_ApoPflicht, "AP", 0, False)
    Set RpCol = .Add(ABDA_ApoPreis, "AVP", 0, False)
    Set RpCol = .Add(ABDA_Packungskomponente, "Packung", 0, False)
    Set RpCol = .Add(ABDA_Einstufung, "NP", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Typ, "Typ", 0, False)
    Set RpCol = .Add(ABDA_Menge, "Menge", 0, False)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Einheit, "Einheit", 50, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Wirkstoff, "Wirkstoff", 0, True)
    Set RpCol = .Add(ABDA_Warengruppe, "Warengruppe", 0, True)
    Set RpCol = .Add(ABDA_Vertriebsstatus, "VS", 50, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Firmenname, "Anbietername", 0, True)
    Set RpCol = .Add(ABDA_MO, "MO", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_VM, "VM", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_MI, "MI", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_NM, "NM", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_AB, "AB", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_NA, "NA", 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(ABDA_Check, vbNullString, 30, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    RpCol.Icon = IC16_Check
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo3 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpl3 " & Err.Number
Resume Next

End Sub

Private Sub FUber()
On Error GoTo InErr

Dim RowNr As Long
Dim AktPo As Long
Dim GesPo As Long
Dim AkSpa As Integer
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMedik
Set RpCo3 = FM.repCont3
Set RpCls = RpCo3.Columns
Set RpRws = RpCo3.Rows

GesPo = RpRws.Count

Screen.MousePointer = vbHourglass

If GesPo > 0 Then
    ReDim GlClp(GesPo, 24) 'WICHTIG!
    For Each RpRow In RpRws
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            If RpRow.Record(23).Checked = True Then
                For AkSpa = 0 To 23 'WICHTIG!
                    If AkSpa < 17 Then
                        If RpRow.Record(AkSpa).Value <> vbNullString Then
                            GlClp(AktPo, AkSpa) = RpRow.Record(AkSpa).Value
                        Else
                            GlClp(AktPo, AkSpa) = vbNullString
                        End If
                    Else
                        If RpRow.Record(AkSpa).Value <> vbNullString Then
                            GlClp(AktPo, AkSpa) = RpRow.Record(AkSpa).Value
                        Else
                            GlClp(AktPo, AkSpa) = 0
                        End If
                    End If
                Next AkSpa
                AktPo = AktPo + 1
            End If
        End If
    Next RpRow
End If

Set RpCo3 = Nothing

Select Case GlBut
Case RibTab_Krankenbla: ABD_Kra AktPo, KraMe
Case RibTab_Abrechnung: ABD_Abr AktPo
Case RibTab_Rezeptmodul: ABD_Rez AktPo, True
Case RibTab_Belegmodul: ABD_Rez AktPo, True
Case RibTab_Tex_Rezept: ABD_Trz AktPo, True
Case RibTab_Kat_Eintrg: ABD_Kat AktPo
End Select

Screen.MousePointer = vbNormal

Unload FM

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FUber " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo InErr

Dim SuVal As Long
Dim SuStr As String
Dim Frage As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim Mld1, Tit1 As String

Set FM = frmMedik
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set TxSu1 = FM.txtSuch1
Set TxSu2 = FM.txtSuch2
Set TxSu3 = FM.txtSuch3
Set TxSu4 = FM.txtSuch4
Set TxSu5 = FM.txtSuch5
Set But01 = FM.btnBut01
Set But02 = FM.btnBut02
Set But03 = FM.btnBut03
Set But04 = FM.btnBut04
Set Lbl04 = FM.lblLab04
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3

Tit1 = "Arzneikatalog-Abgleich"
Mld1 = "Möchten Sie jetzt wirklich einen Arzneikatalog-Abgleich durchführen? Dieser Vorgang kann einige Minuten in Anspruch nehmen."

If Rahm1.Visible = True Then
    If GlBut = RibTab_Kat_Eintrg Then
        If KraMe = True Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = False
                Rahm4.Visible = False
                Rahm5.Visible = False
                Rahm6.Visible = True
                But03.Enabled = False
                DoEvents
                ABD_Ab1
                DoEvents
                Unload FM
                Exit Sub
            End If
        End If
    End If

    If TxSu1.Text <> vbNullString Then 'Arzneimittelname
        SuFel = 1
        SuStr = TxSu1.Text
        ABD_Filt SuStr, SuFel
        DoEvents
        If GlAbP > 0 Then
            Lbl04.Caption = "Gefunden : " & GlAbP
            Rahm1.Visible = False
            Rahm2.Visible = False
            Rahm3.Visible = True
            Rahm4.Visible = False
            Rahm5.Visible = False
            Rahm6.Visible = False
            But01.Enabled = True
            But02.Enabled = True
            If SaLis = True Then
                But03.Caption = "&Sammenln"
                But04.Caption = "&Sammelliste"
            Else
                But03.Caption = "&Einfügen"
            End If
            RpCo1.SetFocus
        End If
    ElseIf TxSu2.Text <> vbNullString Then 'PZN Suche
        SuFel = 2
        SuStr = TxSu2.Text
        SuVal = Val(SuStr)
        SuStr = Format$(SuVal, "00000000")
        ABD_Filt SuStr, SuFel
        DoEvents
        If GlAbP > 0 Then
            Lbl04.Caption = "Gefunden : " & GlAbP
            Rahm1.Visible = False
            Rahm2.Visible = False
            Rahm3.Visible = True
            Rahm4.Visible = False
            Rahm5.Visible = False
            Rahm6.Visible = False
            But01.Enabled = True
            But02.Enabled = True
            If SaLis = True Then
                But03.Caption = "&Sammenln"
                But04.Caption = "&Sammelliste"
            Else
                But03.Caption = "&Einfügen"
            End If
            RpCo1.SetFocus
        End If
    ElseIf TxSu3.Text <> vbNullString Then 'Herstellersuche
        SuFel = 3
        SuStr = TxSu3.Text
        FSpl2 SuFel
        DoEvents
        ABD_Anb SuStr, SuFel
        DoEvents
        If GlAbA > 0 Then
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
            Rahm5.Visible = False
            Rahm6.Visible = False
            But02.Enabled = True
            RpCo2.SetFocus
        End If
    ElseIf TxSu4.Text <> vbNullString Then 'Wirkstoffsuche
        SuFel = 4
        SuStr = TxSu4.Text
        FSpl2 SuFel
        DoEvents
        ABD_Anb SuStr, SuFel
        DoEvents
        If GlAbA > 0 Then
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
            Rahm5.Visible = False
            Rahm6.Visible = False
            But02.Enabled = True
            RpCo2.SetFocus
        End If
    ElseIf TxSu5.Text <> vbNullString Then 'Indikation
        SuFel = 5
        SuStr = TxSu5.Text
        FSpl2 SuFel
        DoEvents
        ABD_Anb SuStr, SuFel
        DoEvents
        If GlAbA > 0 Then
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
            Rahm5.Visible = False
            Rahm6.Visible = False
            But02.Enabled = True
            RpCo2.SetFocus
        End If
    End If
ElseIf Rahm2.Visible = True Then
    Set RpSel = RpCo2.SelectedRows
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        SuVal = AnAry(ABAN_ID3, RpRow.Index)
        ABD_Filt vbNullString, SuFel, SuVal
        DoEvents
        If GlAbP > 0 Then
            Lbl04.Caption = "Gefunden : " & GlAbP
            Rahm1.Visible = False
            Rahm2.Visible = False
            Rahm3.Visible = True
            Rahm4.Visible = False
            Rahm5.Visible = False
            Rahm6.Visible = False
            But01.Enabled = True
            But02.Enabled = True
            If SaLis = True Then
                But03.Caption = "&Sammenln"
                But04.Caption = "&Sammelliste"
            Else
                But03.Caption = "&Einfügen"
            End If
            RpCo1.SetFocus
        End If
    End If
ElseIf Rahm3.Visible = True Then
        If SaLis = True Then
            Set RpSel = RpCo3.SelectedRows
            If RpSel.Count > 10 Then
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = False
                Rahm4.Visible = False
                Rahm5.Visible = False
                Rahm6.Visible = True
                But03.Enabled = False
            End If
            DoEvents
            FSam
        Else
            Set RpSel = RpCo1.SelectedRows
            If RpSel.Count > 10 Then
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = False
                Rahm4.Visible = False
                Rahm5.Visible = False
                Rahm6.Visible = True
                But03.Enabled = False
            End If
            DoEvents
            FEinf
        End If
ElseIf Rahm4.Visible = True Then
        FUber
End If

Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub

Private Sub FZuru()
On Error GoTo InErr

Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmMedik
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set TxSu1 = FM.txtSuch1
Set TxSu2 = FM.txtSuch2
Set TxSu3 = FM.txtSuch3
Set TxSu4 = FM.txtSuch4
Set TxSu5 = FM.txtSuch5
Set But01 = FM.btnBut01
Set But02 = FM.btnBut02
Set But03 = FM.btnBut03
Set But04 = FM.btnBut04
Set Lbl04 = FM.lblLab04
Set RpCo2 = FM.repCont2

If Rahm2.Visible = True Then
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    TxSu1.SetFocus
ElseIf Rahm3.Visible = True Then
    If SuFel > 2 Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        RpCo2.SetFocus
    Else
        Rahm1.Visible = True
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        TxSu1.SetFocus
    End If
    But01.Enabled = False
    But03.Caption = "&Weiter"
    But04.Caption = "&Schließen"
    Lbl04.Caption = vbNullString
ElseIf Rahm4.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    But03.Caption = "&Sammenln"
    But04.Caption = "&Sammelliste"
    But01.Enabled = True
ElseIf Rahm5.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    But03.Enabled = True
    But01.Caption = "&Detailinfo"
    If SaLis = True Then
        But04.Caption = "&Sammelliste"
    End If
End If

Set RpCo2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZuru " & Err.Number
Resume Next

End Sub

Private Sub btnBut01_Click()
    FDet
End Sub
Private Sub btnBut02_Click()
    FRes
    FZuru
End Sub
Private Sub btnBut03_Click()
    FWeit
End Sub

Private Sub btnBut04_Click()
On Error Resume Next

Set FM = frmMedik
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set But01 = FM.btnBut01
Set But02 = FM.btnBut02
Set But03 = FM.btnBut03
Set But04 = FM.btnBut04

If SaLis = True Then
    If Rahm3.Visible = True Then
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = True
        Rahm5.Visible = False
        Rahm6.Visible = False
        But01.Enabled = False
        But03.Caption = "&Einfügen"
        But04.Caption = "&Schließen"
    Else
        Unload Me
    End If
Else
    Unload Me
End If

End Sub
Private Sub chkOpt01_Click()
On Error Resume Next

Set Chk01 = Me.chkOpt01

If FoLad = False Then
    If Chk01.Value = xtpChecked Then
        SaLis = True
    Else
        SaLis = False
    End If
    
    IniSetVal "System", "AbdSam", SaLis
End If

End Sub

Private Sub chkOpt02_Click()
On Error Resume Next

Set Chk01 = Me.chkOpt01
Set Chk02 = Me.chkOpt02
Set TxSu1 = FM.txtSuch1
Set TxSu2 = FM.txtSuch2
Set TxSu3 = FM.txtSuch3
Set TxSu4 = FM.txtSuch4
Set TxSu5 = FM.txtSuch5

If FoLad = False Then
    If Chk02.Value = xtpChecked Then
        KraMe = True
    Else
        KraMe = False
    End If

    If GlBut = RibTab_Kat_Eintrg Then
        Chk01.Enabled = Not KraMe
        TxSu1.Enabled = Not KraMe
        TxSu2.Enabled = Not KraMe
        TxSu3.Enabled = Not KraMe
        TxSu4.Enabled = Not KraMe
        TxSu5.Enabled = Not KraMe
    Else
        IniSetVal "System", "AbdKra", KraMe
    End If
End If

End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True

FKonf
FSpl1
FSpl3

FoLad = False
AFont Me
SFrame 1, Me.hwnd

SDBStb True

End Sub
Private Sub Form_Unload(Cancel As Integer)

SDBStb False
DoEvents
Set frmMedik = Nothing
    
End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If GlAbP > 0 Then
    If GlAbP > Row.Index Then
        If AbAry(Item.Index, Row.Index) <> vbNullString Then
            Metrics.Text = AbAry(Item.Index, Row.Index)
            If Item.Index = ABDA_PZN Then
                Metrics.ItemIcon = IC16_Doc_Norm
            End If
            If AbAry(ABDA_Vertriebsstatus, Row.Index) = "aV" Then
                Metrics.Font.Strikethrough = True
            End If
            If AbAry(ABDA_Vertriebsstatus, Row.Index) = "Zu" Then
                Metrics.Font.Strikethrough = True
                Metrics.ForeColor = vbRed
            End If
        End If
    End If
End If

End Sub
Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If GlKal = False Then
    If KeyCode = vbKeyReturn Then
        FWeit
    End If
End If

End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

If GlKal = False Then
    If Row.GroupRow = False Then
        FWeit
    End If
End If

End Sub
Private Sub repCont2_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If GlAbA > 0 Then
    If GlAbA > Row.Index Then
        If AnAry(Item.Index, Row.Index) <> vbNullString Then
            Metrics.Text = AnAry(Item.Index, Row.Index)
            If SuFel = 3 Then
                If Item.Index = ABAN_Firma1 Then
                    Metrics.ItemIcon = IC16_Doc_Norm
                End If
            Else
                If Item.Index = ABAN_Sorter Then
                    Metrics.ItemIcon = IC16_Doc_Norm
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub repCont2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If GlKal = False Then
    If KeyCode = vbKeyReturn Then
        FWeit
    End If
End If

End Sub
Private Sub repCont2_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

If GlKal = False Then
    If Row.GroupRow = False Then
        FWeit
    End If
End If

End Sub


Private Sub txtSuch1_GotFocus()
    FRes
End Sub

Private Sub txtSuch1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FWeit
    End If
End Sub
Private Sub txtSuch2_GotFocus()
    FRes
End Sub

Private Sub txtSuch2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FWeit
    End If
End Sub
Private Sub txtSuch3_GotFocus()
    FRes
End Sub

Private Sub txtSuch3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FWeit
    End If
End Sub
Private Sub txtSuch4_GotFocus()
    FRes
End Sub

Private Sub txtSuch4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FWeit
    End If
End Sub
Private Sub txtSuch5_GotFocus()
    FRes
End Sub

Private Sub txtSuch5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FWeit
    End If
End Sub
