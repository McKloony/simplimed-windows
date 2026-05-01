VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{621DDB00-A516-11E8-A658-0013D350667C}#3.2#0"; "tx4ole26.ocx"
Begin VB.Form frmTxStat 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Seriendruck"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   Icon            =   "frmTxStat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4000
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6400
      _Version        =   1048579
      _ExtentX        =   11289
      _ExtentY        =   7056
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtEmEmp 
         Height          =   350
         Left            =   500
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2260
         Width           =   4900
         _Version        =   1048579
         _ExtentX        =   8643
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmBet 
         Height          =   350
         Left            =   500
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   4900
         _Version        =   1048579
         _ExtentX        =   8643
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.ProgressBar prbStat1 
         Height          =   350
         Left            =   500
         TabIndex        =   3
         Top             =   1000
         Width           =   4900
         _Version        =   1048579
         _ExtentX        =   8643
         _ExtentY        =   617
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   4400
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3100
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schlieﬂen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnPuBu2 
         Height          =   400
         Left            =   3000
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3100
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Drucken"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnPuBu1 
         Height          =   400
         Left            =   1600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3100
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Vorschau"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   300
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3100
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hilfe"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   240
         Left            =   495
         TabIndex        =   9
         Top             =   1460
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   423
         _StockProps     =   79
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   600
         Left            =   495
         TabIndex        =   8
         Top             =   135
         Width           =   5200
         _Version        =   1048579
         _ExtentX        =   9172
         _ExtentY        =   1058
         _StockProps     =   79
         Caption         =   $"frmTxStat.frx":6852
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin Tx4oleLib.TXTextControl TexCont3 
      Height          =   5300
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2800
      Visible         =   0   'False
      Width           =   5960
      _Version        =   196610
      _ExtentX        =   10513
      _ExtentY        =   9349
      _StockProps     =   73
      BackColor       =   16777215
      Language        =   49
      BorderStyle     =   0
      BackStyle       =   1
      ControlChars    =   0   'False
      EditMode        =   0
      HideSelection   =   -1  'True
      InsertionMode   =   -1  'True
      MousePointer    =   0
      ZoomFactor      =   45
      ViewMode        =   2
      ClipChildren    =   0   'False
      ClipSiblings    =   -1  'True
      SizeMode        =   0
      TabKey          =   -1  'True
      FormatSelection =   0   'False
      VTSpellDictionary=   "C:\PROGRA~1\THEIMA~1\TXTEXT~1.0\Bin\AMERICAN.VTD"
      ScrollBars      =   0
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
      IsTrackChangesEnabled=   0   'False
      IsFormulaCalculationEnabled=   -1  'True
      FormulaReferenceStyle=   0
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Text            =   "A"
      Top             =   12000
      Width           =   80
   End
End
Attribute VB_Name = "frmTxStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Lbl01 As XtremeSuiteControls.Label
Private Lbl02 As XtremeSuiteControls.Label
Private TxBet As XtremeSuiteControls.FlatEdit
Private TxEmp As XtremeSuiteControls.FlatEdit
Private CmBar As XtremeCommandBars.CommandBar
Private CmAcs As XtremeCommandBars.CommandBarActions
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private CoDia As XtremeSuiteControls.CommonDialog
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private Rahm1 As XtremeSuiteControls.GroupBox
Private TxCoM As Tx4oleLib.TXTextControl
Private TxCoN As Tx4oleLib.TXTextControl

Private clFil As clsFile

Public EmSen As Boolean
Public EmTes As Boolean
Private Sub FDruk()
On Error GoTo InErr

Dim DrhDc As Long
Dim PatNr As Long
Dim GesZa As Long
Dim AktZa As Long
Dim PaStr As String
Dim TmDat() As Byte
Dim SeiZa As Integer
Dim DrKop As Integer
Dim ReDru As Integer
Dim TxDum As VB.TextBox
Dim RpCo9 As XtremeReportControl.ReportControl

Set FM = frmTxStat
Set TxCoM = FM.TexCont3
Set PrBr1 = FM.prbStat1
Set TxDum = FM.txtDummy
Set Lbl02 = FM.lblLab02

Set CoDia = frmMain.comDialo
Set TxCoN = frmMain.TexCont1
Set RpCo9 = frmMain.repCont9
Set RpRcs = RpCo9.Records

GesZa = RpRcs.Count

If GesZa > 0 Then
    SeiZa = TxCoN.CurrentPages
    TmDat = TxCoN.SaveToMemory(3, False)

    PrBr1.Min = 0
    PrBr1.Max = GesZa

    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = GlTDa 'Textverarbeitung Dateiname
        .FileName = vbNullString
        ReDru = .ShowPrinter
        DrhDc = .hDC
        DrKop = .Copies
    End With
    TxCoM.PrintDevice = DrhDc

    For AktZa = 0 To GesZa - 1
        If CBool(SeAry(4, AktZa)) = False Then
            PatNr = SeAry(0, AktZa)
            PaStr = SeAry(1, AktZa)

            Lbl02.Caption = PaStr
            S_TxEin PatNr 'Laden der Patientendaten in Array GlSer()
            DoEvents
            
            TxCoM.LoadFromMemory TmDat, 3, False 'Laden des Dokumentes
            DoEvents
            
            STxV3 'Verbinden der Textfelder mit GlSer()
            DoEvents
            If ReDru > 0 Then
                TxCoM.PrintDoc GlTDa, 1, SeiZa, DrKop
            End If
        End If
        DoEvents
        If TxDum.Text = "B" Then Exit For 'Abbrechen
        PrBr1.Value = AktZa + 1
    Next AktZa
End If

Set CoDia = Nothing
Set RpRcs = Nothing
Set RpCo9 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDruk " & Err.Number
Resume Next

End Sub
Private Sub FLoad()
On Error GoTo InErr

Dim TxFnt As New StdFont

Set FM = frmTxStat
Set TxCoN = FM.TexCont3
Set TxBet = FM.txtEmBet
Set TxEmp = FM.txtEmEmp
Set PrBr1 = FM.prbStat1
Set PuBu1 = FM.btnPuBu1
Set PuBu2 = FM.btnPuBu2
Set Rahm1 = FM.frmRahm1
Set Lbl01 = FM.lblLab01
Set Lbl02 = FM.lblLab02

TxFnt.Name = GlXFt.Name
TxFnt.SIZE = GlXFt.SIZE

With PrBr1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .FlatStyle = False
    .Scrolling = xtpProgressBarStandard
    .UseVisualStyle = False
End With

With TxCoN
    .ViewMode = 2
    .Alignment = 0
    .AllowUndo = False 'WICHTIG!
    .Enabled = True
    .DataTextFormat = 0
    .AutoExpand = False
    .ClipChildren = False
    .ClipSiblings = False
    .ControlChars = False
    .ColumnLineColor = 0
    .BackColor = -2147483643 '16777215
    .BackStyle = 1
    .BaseLine = 0
    .BorderStyle = 0
    .EditMode = 2
    .FontBold = TxFnt.Bold
    .FontItalic = TxFnt.Italic
    .FontUnderline = TxFnt.Underline
    .FontStrikethru = TxFnt.Strikethrough
    .FontName = TxFnt.Name
    .FontSize = TxFnt.SIZE
    .FormatSelection = True
    .HeaderFooterStyle = txUnframed
    .HideSelection = False
    .InsertionMode = True
    .Language = 49
    .PageViewStyle = txGradientColors
    .PageOrientation = 0
    .PrintColors = True
    .ScrollBars = 3
    .SizeMode = 0
    .SelectionViewMode = 1
    .TabKey = True
    .TextBkColor = 16777215
    .TextFrameMarkerLines = False
    .TableGridLines = True
    .EnableHyperlinks = True
    .ZoomFactor = 45
    .WordWrapMode = 1
End With

Me.BackColor = GlBak
Rahm1.BackColor = GlBak

If EmSen = True Then
    FM.Caption = "Newsletterversand"
    TxBet.Enabled = True
    TxBet.Text = "Newsletterbetreff... "
    PuBu1.Enabled = False
    PuBu2.Caption = "Senden"
    If EmTes = True Then
        TxEmp.Enabled = True
        Lbl01.Caption = "Bitte klicken Sie auf Senden, um den Newsletter Testversand zu starten."
        If GlMkt(1, 13) <> vbNullString Then
            TxEmp.Text = GlMkt(1, 13)
        Else
            If GlThe(GlSMa, 16) <> vbNullString Then
                TxEmp.Text = GlThe(GlSMa, 16)
            Else
                TxEmp.Text = "ihre@emailadresse.com"
            End If
        End If
    Else
        Lbl01.Caption = "Bitte klicken Sie auf Senden, um den Newsletter-Versand zu starten. Dieser Vorgang kann eine l‰ngere Zeit in Anspruch nehmen."
    End If
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FVors()
On Error GoTo InErr
'Serienbrief Exportieren in eine Datei

Dim RetWe As Long
Dim DrhDc As Long
Dim PatNr As Long
Dim Lange As Long
Dim AktZa As Long
Dim GesZa As Long
Dim PaStr As String
Dim TmDat() As Byte
Dim TmPuf() As Byte
Dim SeiZa As Integer
Dim AnzDa As Integer
Dim StaNa As Integer
Dim SeiPo As Variant
Dim TxFnt As New StdFont
Dim TxDum As VB.TextBox
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo9 As XtremeReportControl.ReportControl

Set FM = frmMain
Set TxCoN = FM.TexCont1
Set RpCo9 = FM.repCont9
Set RpRcs = RpCo9.Records

Set TxCoM = Me.TexCont3
Set Lbl02 = Me.lblLab02
Set TxDum = Me.txtDummy
Set PrBr1 = Me.prbStat1

GesZa = RpRcs.Count

PrBr1.Max = GesZa

If GesZa > 0 Then
    GlTxS = True
    With TxCoN
        TmDat = .SaveToMemory(3, False)
        Lange = Len(.Text)
    End With

    For AktZa = 0 To GesZa - 1
        If CBool(SeAry(4, AktZa)) = False Then
            PatNr = SeAry(0, AktZa)
            PaStr = SeAry(1, AktZa)
                                    
            Lbl02.Caption = PaStr
            DoEvents

            S_TxEin PatNr 'Laden der Patientendaten in Array GlSer()
            DoEvents
            
            TxCoM.LoadFromMemory TmDat, 3, False 'Laden des Dokumentes
            DoEvents
                        
            STxV3 'Verbinden der Textfelder mit GlSer()
            DoEvents
            
            TmPuf = TxCoM.SaveToMemory(3, False)
            DoEvents
            
            TxCoN.SelStart = Len(TxCoN.Text)
            DoEvents
            
            TxCoN.SelText = Chr$(12)
            DoEvents
            
            If AktZa = 0 Then
                TxCoN.LoadFromMemory TmPuf, 3, False
            Else
                TxCoN.LoadFromMemory TmPuf, 3, True
            End If
        End If
        DoEvents
        If TxDum.Text = "B" Then Exit For 'Abbrechen
        PrBr1.Value = AktZa + 1
    Next AktZa
    
    DoEvents
    Set CmBrs = frmMain.comBar01
    Set CmAcs = CmBrs.Actions
    CmAcs(Tex_DocVor).Enabled = True
    CmAcs(Tex_DocExp).Enabled = True
    CmAcs(Tex_DocMa1).Enabled = True
    CmAcs(Tex_DocMa2).Enabled = True
    CmAcs(Tex_DocSe1).Enabled = True
    CmAcs(Tex_DocSe2).Enabled = True
    CmAcs(Tex_DocSe3).Enabled = True
    CmAcs(Tex_EtiDru).Enabled = True
    CmAcs(Tex_Eigens).Enabled = True
End If

Set RpRcs = Nothing
Set RpCo9 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FVors " & Err.Number
Resume Next

End Sub
Private Sub FSend()
On Error GoTo InErr

Dim PatNr As Long
Dim PaStr As String
Dim EmEmp As String
Dim TeEmp As String
Dim EmBet As String
Dim TmHtm As String
Dim MaTex As String
Dim ASCSt As String
Dim GesZa As Integer
Dim AktZa As Integer
Dim EmVer As Boolean
Dim TmDat() As Byte
Dim SeiZa As Integer
Dim TxDum As VB.TextBox
Dim RpCo9 As XtremeReportControl.ReportControl

Set FM = frmTxStat
Set TxCoM = FM.TexCont3
Set TxBet = FM.txtEmBet
Set TxEmp = FM.txtEmEmp
Set PrBr1 = FM.prbStat1
Set TxDum = FM.txtDummy
Set Lbl02 = FM.lblLab02

Set TxCoN = frmMain.TexCont1
Set RpCo9 = frmMain.repCont9
Set RpRcs = RpCo9.Records

GesZa = RpRcs.Count

If TxBet.Text <> vbNullString Then
    EmBet = TxBet.Text
Else
    EmBet = GlMan(GlSMa, 1)
End If

If TxEmp.Text <> vbNullString Then
    TeEmp = TxEmp.Text
Else
    TeEmp = vbNullString
End If

If EmTes = True Then
    SeiZa = TxCoN.CurrentPages
    TmDat = TxCoN.SaveToMemory(3, False)
        
    PatNr = SeAry(0, 0)
    PaStr = SeAry(1, 0)
    
    Lbl02.Caption = PaStr
    S_TxEin PatNr 'Laden der Patientendaten in Array GlSer()
    DoEvents
    
    TxCoM.LoadFromMemory TmDat, 3, False 'Laden des Dokumentes
    DoEvents
    
    STxV3 'Verbinden der Textfelder mit GlSer()
    DoEvents
    
    TmHtm = TxCoM.SaveToMemoryBuffer(TmHtm, 4, 0) 'HTML Umwandlung
    MaTex = TxCoM.Text

    If TeEmp <> vbNullString Then
        EmVer = SEmSe(TeEmp, EmBet, MaTex, , TmHtm, , False)
    End If
Else
    If GesZa > 0 Then
        SeiZa = TxCoN.CurrentPages
        TmDat = TxCoN.SaveToMemory(3, False)

        PrBr1.Min = 0
        PrBr1.Max = GesZa

        For AktZa = 0 To GesZa - 1
            If CBool(SeAry(4, AktZa)) = False Then
                PatNr = SeAry(0, AktZa)
                PaStr = SeAry(1, AktZa)
                If SeAry(9, AktZa) <> vbNullString Then
                    EmEmp = SeAry(9, AktZa)
                ElseIf SeAry(10, AktZa) <> vbNullString Then
                    EmEmp = SeAry(10, AktZa)
                Else
                    EmEmp = vbNullString
                End If
    
                Lbl02.Caption = PaStr
                S_TxEin PatNr 'Laden der Patientendaten in Array GlSer()
                DoEvents
                
                TxCoM.LoadFromMemory TmDat, 3, False 'Laden des Dokumentes
                DoEvents
                
                STxV3 'Verbinden der Textfelder mit GlSer()
                DoEvents
                
                TmHtm = TxCoM.SaveToMemoryBuffer(TmHtm, 4, 0) 'HTML Umwandlung
                MaTex = TxCoM.Text

                If EmEmp <> vbNullString Then
                    EmVer = SEmSe(EmEmp, EmBet, MaTex, , TmHtm, , False)
                    DoEvents
                End If
            End If
            DoEvents
            If TxDum.Text = "B" Then Exit For 'Abbrechen
            PrBr1.Value = AktZa + 1
        Next AktZa
    End If
End If

Set RpRcs = Nothing
Set RpCo9 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSend " & Err.Number
Resume Next

End Sub

Private Sub btnHilfe_Click()
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
Private Sub btnPuBu1_Click()
    FVors
    Unload Me
End Sub
Private Sub btnPuBu2_Click()
    If EmSen = True Then
        FSend
    Else
        FDruk
    End If
    Unload Me
End Sub
Private Sub btnSchlieﬂ_Click()
    Me.txtDummy.Text = "B"
    Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

AFont Me

FLoad

SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTxStat = Nothing
End Sub

Private Sub txtEmBet_GotFocus()
    Me.txtEmBet.SelStart = 0
    Me.txtEmBet.SelLength = Len(Me.txtEmBet.Text)
End Sub

Private Sub txtEmEmp_GotFocus()
    Me.txtEmEmp.SelStart = 0
    Me.txtEmEmp.SelLength = Len(Me.txtEmEmp.Text)
End Sub
