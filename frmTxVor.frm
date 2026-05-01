VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{0EDB0C00-C493-11E7-A629-0013D350667C}#3.1#0"; "tx4ole25.ocx"
Begin VB.Form frmTxVor 
   Caption         =   "Dokumentenanzeige"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   ControlBox      =   0   'False
   Icon            =   "frmTxVor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7245
   Begin Tx4oleLib.TXTextControl TexCont3 
      Height          =   1995
      Left            =   4320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2205
      Width           =   1995
      _Version        =   196609
      _ExtentX        =   3519
      _ExtentY        =   3528
      _StockProps     =   73
      BackColor       =   16777215
      Language        =   49
      BorderStyle     =   1
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
      VTSpellDictionary=   "C:\PROGRA~1\TEXTCO~1\TXTEXT~2.0AC\Bin\AMERICAN.VTD"
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
      TextBkColor     =   16777215
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
   End
   Begin Tx4oleLib.TXRuler TexRule2 
      Height          =   555
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   405
      _Version        =   196609
      _ExtentX        =   714
      _ExtentY        =   979
      _StockProps     =   96
      Language        =   49
      ScaleUnits      =   0
      Appearance      =   3
      Direction       =   1
      EnablePageMargins=   -1  'True
      RightToLeft     =   0   'False
      ReadOnly        =   0   'False
   End
   Begin Tx4oleLib.TXRuler TexRule1 
      Height          =   405
      Left            =   3960
      TabIndex        =   2
      Top             =   1560
      Width           =   1005
      _Version        =   196609
      _ExtentX        =   1764
      _ExtentY        =   714
      _StockProps     =   96
      Language        =   49
      ScaleUnits      =   0
      Appearance      =   3
      Direction       =   0
      EnablePageMargins=   -1  'True
      RightToLeft     =   0   'False
      ReadOnly        =   0   'False
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTxVor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAcs As XtremeCommandBars.CommandBarActions
Private TxCoN As Tx4oleLib.TXTextControl
Private TxRu1 As Tx4oleLib.TXRuler
Private TxRu2 As Tx4oleLib.TXRuler

Private WithEvents CmSta As XtremeCommandBars.StatusBar
Attribute CmSta.VB_VarHelpID = -1

Public DaNam As String
Private Sub FClos()
On Error GoTo InErr

Set FM = frmTxVor
Set TxCoN = FM.TexCont3

IniSetVal "GUI", "ZomWor", GlZoW

TxCoN.ResetContents

Set TxCoN = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
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

Private Sub FMenu()
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSli As XtremeCommandBars.StatusBarSliderPane
Dim CmSwi As XtremeCommandBars.StatusBarSwitchPane
Dim CmPrg As XtremeCommandBars.StatusBarProgressPane

Set FM = frmTxVor
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

With CmSta
    .removeAll
    .EnableCustomization False
    .Font.Name = GlTFt.Name
    .Font.SIZE = 8
    .DrawDisabledText = True
    .EnableMarkup = False
    .IdleText = vbNullString
    .ShowSizeGripper = True
    Set CmPan = .AddPane(Tex_Pa_Plac1)
    With CmPan
        .Enabled = False
        .BeginGroup = False
        .Width = 40
    End With
    Set CmPan = .AddPane(Tex_Pa_Labl1)
    With CmPan
        .Style = SBPS_STRETCH
        .Text = DaNam
        .BeginGroup = False
        .Alignment = xtpAlignmentLeft
        .Style = SBPS_NOBORDERS
    End With
    Set CmPan = .AddPane(Tex_Pa_Plac5)
    With CmPan
        .Style = SBPS_STRETCH
        .Text = vbNullString
    End With
    Set CmPan = .AddPane(Tex_Pa_Seite)
    With CmPan
        .Alignment = xtpAlignmentLeft
        .BeginGroup = True
        .Text = " Seiten:"
        .Width = 100
    End With
    Set CmSwi = .AddSwitchPane(Tex_Pa_Layou)
    CmSwi.AddSwitch IC16_AnsNor, "Normalansicht"
    CmSwi.AddSwitch IC16_AnsBre, "Seitenansicht"
    CmSwi.AddSwitch IC16_AnsFli, "Fließtextansicht"
    Select Case GlViW 'ViewMode Textvorschau
    Case 0: CmSwi.Checked = IC16_AnsNor
    Case 2: CmSwi.Checked = IC16_AnsBre
    Case 3: CmSwi.Checked = IC16_AnsFli
    End Select
    Set CmPan = .AddPane(Tex_Pa_Linia)
    With CmPan
        .Alignment = xtpAlignmentRight
        .BeginGroup = True
        .Text = "Liniale:"
    End With
    Set CmSwi = .AddSwitchPane(Tex_Linial)
    CmSwi.AddSwitch IC16_Ruler, "Liniale"
    If GlLiW = True Then 'Lineal Textvorschau
        CmSwi.Checked = IC16_Ruler
    End If
    Set CmPan = .AddPane(Tex_Pa_ZoPan)
    With CmPan
        .SetPadding 15, 0, 15, 0
        .Text = "Zoom: " & GlZoW & "%"
        .ToolTip = "Zeigt die Zoomgröße des Dokuments"
    End With
    Set CmPan = .AddPane(Tex_Pa_Plac2)
    With CmPan
        .Enabled = False
        .Style = SBPS_NOBORDERS
        .BeginGroup = False
        .Width = 15
    End With
    Set CmSli = .AddSliderPane(Tex_Pa_ZoSli)
    With CmSli
        .BeginGroup = False
        .Min = 10
        .Max = 200
        .SetTicks 100
        .SetTooltipPart XTP_SB_LINELEFT, "Zoom verringern"
        .SetTooltipPart XTP_SB_LINERIGHT, "Zoom vergrößern"
        .ToolTip = "Zoomgröße"
        .Value = GlZoW
        .Width = 120
    End With
    Set CmPan = .AddPane(Tex_Pa_Plac3)
    With CmPan
        .Enabled = False
        .Style = SBPS_NOBORDERS
        .BeginGroup = False
        .Width = 50
    End With
    .Visible = True
End With

CmBrs.PaintManager.RefreshMetrics
DoEvents
CmBrs.RecalcLayout
DoEvents

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FToSe(ByVal TolId As Integer)
On Error GoTo InErr

Dim SuStr As String
Dim DatNa As String
Dim Posit As Integer
Dim AnIdx As Integer
Dim Frage As Integer
Dim SeiPo As Variant
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCo1 As XtremeCommandBars.CommandBarComboBox

Set FM = frmTxVor
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont3
Set TxRu1 = FM.TexRule1
Set TxRu2 = FM.TexRule2
Set CmAcs = CmBrs.Actions

Set CmCo1 = CmBrs.FindControl(CmCo1, SY_OP_Ansicht, , True)
Set CmEdi = CmBrs.FindControl(CmEdi, SY_OP_SubDe1, , True)

AnIdx = CmCo1.ListIndex
SuStr = CmEdi.Text

Select Case TolId:
Case SY_OP_Ansicht:
        With TxCoN
            .HeaderFooterActivate txMainText
            .SetFocus
            .PageSelect AnIdx
        End With
Case SY_OP_Kopieren:
        TxCoN.Clip 2
Case SY_OP_Speichern:
        Posit = InStrRev(DaNam, ".", Len(DaNam), 1)
        If Posit > 0 Then
            DatNa = Mid$(DaNam, 1, Len(DaNam) - (Len(DaNam) - Posit)) & "pdf"
        Else
            DatNa = vbNullString
        End If
        SExFo 9, 1, 0, DatNa
Case SY_OP_Drucken:
        TxCoN.PrintDialog vbNullString
Case SY_OP_SubDe1:
        TxCoN.Find SuStr, 1
Case SY_OP_SubDe2:
        TxCoN.Find SuStr, -1, 1
Case SY_OP_SubDe3:
        TxCoN.Find SuStr, -1
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing

Set TxCoN = Nothing
Set TxRu1 = Nothing
Set TxRu2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FToSe " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

Select Case TolId
Case KY_F1: FHilfe
Case KY_F10: FToSe SY_OP_Drucken
Case KY_F11: Unload Me
Case SY_OP_Ansicht: FToSe TolId
Case SY_OP_Kopieren: FToSe TolId
Case SY_OP_Speichern: FToSe TolId
Case SY_OP_Drucken: FToSe TolId
Case SY_OP_SubDe1: FToSe TolId
Case SY_OP_SubDe2: FToSe TolId
Case SY_OP_SubDe3: FToSe TolId
Case SY_OP_Abbruch: Unload Me
End Select

End Sub
Private Sub Form_Load()
    FMenu
    GlKeL = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmTxVor = Nothing
End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub

Private Sub comBar02_Resize()
    If GlKeL = False Then
        VoTxPo
    End If
End Sub

Private Sub CmSta_SliderPaneClick(ByVal Pane As XtremeCommandBars.StatusBarSliderPane, ByVal Command As XtremeCommandBars.XTPSliderCommand, ByVal Pos As Long)
On Error Resume Next

Dim FrZom As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTxVor
Set TxCoN = FM.TexCont3
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

FrZom = GlZoW

If GlKeL = True Then Exit Sub

Select Case Command
Case XTP_SB_LEFT: FrZom = 0
Case XTP_SB_RIGHT: FrZom = 200
Case XTP_SB_LINELEFT: FrZom = WinMax((Int(FrZom / 10) - 1) * 10, 0)
Case XTP_SB_LINERIGHT: FrZom = WinMin((Int(FrZom / 10) + 1) * 10, 200)
Case XTP_SB_THUMBTRACK: FrZom = Pos
Case XTP_SB_PAGELEFT: FrZom = WinMax(FrZom - 20, 0)
Case XTP_SB_PAGERIGHT: FrZom = WinMin(FrZom + 20, 200)
End Select

If (FrZom = GlZoW) Then Exit Sub

If FrZom > 10 Then
    GlZoW = FrZom
    Pane.Value = FrZom
    TxCoN.ZoomFactor = GlZoW
    CmSta.FindPane(Tex_Pa_ZoPan).Text = "Zoom: " & Format$(FrZom, "000") & "%"
End If

End Sub
Private Sub CmSta_SwitchPaneClick(ByVal Pane As XtremeCommandBars.StatusBarSwitchPane, ByVal Switch As Long)
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSwi As XtremeCommandBars.StatusBarSwitchPane

Set FM = frmTxVor
Set TxCoN = FM.TexCont3
Set TxRu1 = FM.TexRule1
Set TxRu2 = FM.TexRule2
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

If GlKeL = True Then Exit Sub

With CmSta
    Set CmSwi = .FindPane(Pane.id)
    Select Case Pane.id
    Case Tex_Pa_Layou:
        Select Case Switch
        Case IC16_AnsNor:
            GlViW = 0
            CmSwi.Checked = IC16_AnsNor
        Case IC16_AnsBre:
            GlViW = 2
            CmSwi.Checked = IC16_AnsBre
        Case IC16_AnsFli:
            GlViW = 3
            CmSwi.Checked = IC16_AnsFli
        End Select
        TxCoN.ViewMode = GlViW 'ViewMode Textvorschau
        IniSetVal "System", "ViewWo", "C" & GlViW
    Case Tex_Linial:
        GlLiW = Not GlLiW 'Lineal Textvorschau
        If GlLiW = True Then
            CmSwi.Checked = IC16_Ruler
        Else
            CmSwi.Checked = 0
        End If
        TxRu1.Visible = GlLiW
        TxRu2.Visible = GlLiW
        IniSetVal "Layout", "LinWor", GlLiW
        VoTxPo
    End Select
End With

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SwitchPaneClick " & Err.Number
Resume Next

End Sub
Private Sub TexCont3_Change()
On Error Resume Next

Set TxCoN = Me.TexCont3

End Sub
Private Sub TexCont3_Error(Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next

Set TxCoN = Me.TexCont3

If Number = 321 Then
    Select Case LCase(GlTxU)
    Case "doc": TxCoN.Load GlTxF, , 13 'Filname für Textcontrol Error
    Case "docx": TxCoN.Load GlTxF, , 9
    End Select
End If

CancelDisplay = True

End Sub
