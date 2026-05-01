VERSION 5.00
Object = "{28EBE202-E7C7-11D0-B183-0040E994B58D}#1.0#0"; "cmll22v.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Begin VB.Form frmDruVo 
   Caption         =   "Druckvorschau"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13245
   ControlBox      =   0   'False
   Icon            =   "frmDruVo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   13245
   Begin CMLL22VLibCtl.LlViewCtrl LLDruVo 
      Height          =   5415
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   11295
      _cx             =   19923
      _cy             =   9551
      _cx             =   19923
      _cy             =   9551
      ToolbarEnabled  =   -1  'True
      SaveAsFilePath  =   ""
      FileURL         =   ""
      BackColor       =   13160660
      Enabled         =   -1  'True
      AsyncDownload   =   -1  'True
      ShowExitButton  =   -1  'True
      ShowThumbnails  =   -1  'True
      Language        =   -1
      SlideshowMode   =   0   'False
      ShowUnprintableArea=   0   'False
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
Attribute VB_Name = "frmDruVo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PrtVo As LlViewCtrl
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAcs As XtremeCommandBars.CommandBarActions

Private WithEvents CmSta As XtremeCommandBars.StatusBar
Attribute CmSta.VB_VarHelpID = -1

Public DaNam As String

Private FoZom As Long

Private clFil As clsFile

Private Sub FClos()
On Error GoTo InErr

Dim LoNam As String

Set FM = frmDruVo
Set PrtVo = FM.LLDruVo

Set clFil = New clsFile

PrtVo.FileURL = vbNullString
DoEvents

With clFil
    LoNam = GlTmp & "*.LL"
    .DaLoe = LoNam & vbNullChar
    .FilLoe
End With

If FoZom <> GlZoD Then
    IniSetVal "GUI", "ZomDru", GlZoD
End If

Set clFil = Nothing

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

Set FM = frmDruVo
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
    Set CmPan = .AddPane(Tex_Pa_Linia)
    With CmPan
        .Alignment = xtpAlignmentRight
        .BeginGroup = True
        .Text = "Thumbnails:"
    End With
    Set CmSwi = .AddSwitchPane(Tex_Linial)
    CmSwi.AddSwitch IC16_AnsFli, "Thumbnails"
    If GlLiD = True Then 'Lineal Druckvorschau
        CmSwi.Checked = IC16_AnsFli
    End If
    Set CmPan = .AddPane(Tex_Pa_ZoPan)
    With CmPan
        .SetPadding 15, 0, 15, 0
        .Text = "Zoom: " & GlZoD & "%"
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
        .Value = GlZoD
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
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

Select Case TolId
Case KY_F1: FHilfe
Case KY_F10: FToSe SY_OP_Drucken
Case KY_F11: Unload Me
Case SY_OP_Nav_Zuru: FToSe TolId
Case SY_OP_Nav_Vor: FToSe TolId
Case SY_OP_Ansicht: FToSe TolId
Case SY_OP_SubDe1: FToSe TolId
Case SY_OP_SubDe2: FToSe TolId
Case SY_OP_SubDe3: FToSe TolId
Case SY_OP_Speichern: FToSe TolId
Case SY_OP_Drucken: FToSe TolId
Case SY_OP_UeberEinz: FToSe TolId
Case SY_OP_UeberNetz: FToSe TolId
Case SY_OP_Nav_Erst: FToSe TolId
Case SY_OP_SubDe1: FToSe TolId
Case SY_OP_SubDe2: FToSe TolId
Case SY_OP_SubDe3: FToSe TolId
Case SY_OP_Abbruch: Unload Me
End Select

End Sub
Private Sub FToSe(ByVal TolId As Integer)
On Error GoTo InErr

Dim SuStr As String
Dim AkZom As Integer
Dim AktZa As Integer
Dim AkSei As Integer
Dim GeSei As Integer
Dim AnIdx As Integer
Dim Frage As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCo1 As XtremeCommandBars.CommandBarComboBox

Set FM = frmDruVo
Set CmBrs = FM.comBar02
Set PrtVo = FM.LLDruVo
Set CmAcs = CmBrs.Actions

Set CmCo1 = CmBrs.FindControl(CmCo1, SY_OP_Ansicht, , True)
Set CmEdi = CmBrs.FindControl(CmEdi, SY_OP_SubDe1, , True)

AnIdx = CmCo1.ListIndex

SuStr = CmEdi.Text

AkZom = PrtVo.GetZoom
GeSei = PrtVo.PaGes

Select Case TolId:
Case SY_OP_Nav_Zuru:
        PrtVo.GotoPrev
        AkSei = PrtVo.CurrentPage
        CmCo1.ListIndex = AkSei
Case SY_OP_Nav_Vor:
        PrtVo.GotoNext
        AkSei = PrtVo.CurrentPage
        CmCo1.ListIndex = AkSei
Case SY_OP_Ansicht:
        PrtVo.GotoFirst
        For AktZa = 1 To GeSei
            If AktZa < AnIdx Then
                PrtVo.GotoNext
            Else
                Exit For
            End If
        Next AktZa
Case SY_OP_SubDe1:
        PrtVo.SearchFirst SuStr, False
Case SY_OP_SubDe2:
        PrtVo.SearchFirst vbNullString, False
Case SY_OP_SubDe3:
        PrtVo.SearchNext
Case SY_OP_Speichern:
        If GlDru.GoBDk = True Then
            If GlDru.ReAbs = True Then 'WICHTIG GoBD
                PrtVo.SaveAs
            End If
        Else
            PrtVo.SaveAs
        End If
Case SY_OP_Drucken:
        If GlDru.GoBDk = True Then
            If GlDru.ReAbs = True Then 'WICHTIG GoBD
                PrtVo.PrintAllPages False
            End If
        Else
            PrtVo.PrintAllPages False
        End If
Case SY_OP_UeberEinz:
        PrtVo.PrintAllPages True
Case SY_OP_UeberNetz:
        PrtVo.PrintCurrentPage
Case SY_OP_Nav_Erst:
        PrtVo.PrintPage 1
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing
Set PrtVo = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FToSe " & Err.Number
Resume Next

End Sub
Private Sub Form_Load()
    FMenu
    FoZom = GlZoD
    GlKeL = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmDruVo = Nothing
End Sub
Private Function LLDruVo_BtnPress(ByVal nID As Long) As Boolean
On Error Resume Next

Select Case nID
Case 114: Unload Me
End Select

End Function
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
Private Sub comBar02_Resize()
    If GlKeL = False Then
        VoDrPo
    End If
End Sub
Private Sub CmSta_SliderPaneClick(ByVal Pane As XtremeCommandBars.StatusBarSliderPane, ByVal Command As XtremeCommandBars.XTPSliderCommand, ByVal Pos As Long)
On Error Resume Next

Dim FrZom As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmDruVo
Set PrtVo = FM.LLDruVo
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

FrZom = GlZoD

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

If (FrZom = GlZoD) Then
    Exit Sub
End If

If FrZom > 10 Then
    GlZoD = FrZom
    Pane.Value = FrZom
    PrtVo.SetZoom GlZoD
    CmSta.FindPane(Tex_Pa_ZoPan).Text = "Zoom: " & Format$(FrZom, "000") & "%"
End If

End Sub

Private Sub CmSta_SwitchPaneClick(ByVal Pane As XtremeCommandBars.StatusBarSwitchPane, ByVal Switch As Long)
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSwi As XtremeCommandBars.StatusBarSwitchPane

Set FM = frmDruVo
Set PrtVo = FM.LLDruVo
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

If GlKeL = True Then Exit Sub

With CmSta
    Set CmSwi = .FindPane(Pane.id)
    Select Case Pane.id
    Case Tex_Linial:
        GlLiD = Not GlLiD 'Lineal Druckvorschau
        If GlLiD = True Then
            CmSwi.Checked = IC16_AnsFli
        Else
            CmSwi.Checked = 0
        End If
        PrtVo.ShowThumbnails = GlLiD
        IniSetVal "Layout", "LinDru", GlLiD
    End Select
End With

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SwitchPaneClick " & Err.Number
Resume Next

End Sub
