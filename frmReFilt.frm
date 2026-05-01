VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReFilt 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnungsfilter"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   17
      Top             =   5600
      Width           =   9700
      _Version        =   1048579
      _ExtentX        =   17110
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   7700
         TabIndex        =   16
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
         Default         =   -1  'True
         Height          =   400
         Left            =   6300
         TabIndex        =   15
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
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   5000
         TabIndex        =   14
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
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4700
      Left            =   600
      TabIndex        =   1
      Top             =   700
      Width           =   4000
      _Version        =   1048579
      _ExtentX        =   7056
      _ExtentY        =   8290
      _StockProps     =   79
      Caption         =   "Adressengruppen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   2700
         Left            =   340
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   4762
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         HideSelection   =   0   'False
         BackColor       =   16777215
         ForeColor       =   4473924
      End
      Begin XtremeSuiteControls.CheckBox chkMarki 
         Height          =   225
         Left            =   350
         TabIndex        =   4
         Top             =   3500
         Width           =   3400
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "gefundene Rechnungen markieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   315
         Left            =   340
         TabIndex        =   5
         Top             =   4100
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab09 
         Height          =   220
         Left            =   350
         TabIndex        =   27
         Top             =   3860
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandant :"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4700
      Left            =   5000
      TabIndex        =   2
      Top             =   700
      Width           =   4000
      _Version        =   1048579
      _ExtentX        =   7064
      _ExtentY        =   8290
      _StockProps     =   79
      Caption         =   "Rechnungsoptionen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cmbReTyp 
         Height          =   310
         Left            =   340
         TabIndex        =   6
         Top             =   600
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbDiagn 
         Height          =   310
         Left            =   340
         TabIndex        =   7
         Top             =   1300
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbZahlu 
         Height          =   310
         Left            =   340
         TabIndex        =   10
         Top             =   2700
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.FlatEdit txtReGre 
         Height          =   350
         Left            =   2200
         TabIndex        =   13
         Top             =   4100
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cmbVersa 
         Height          =   310
         Left            =   340
         TabIndex        =   8
         Top             =   2000
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2805
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbStatu 
         Height          =   315
         Left            =   340
         TabIndex        =   11
         Top             =   3400
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbJahre 
         Height          =   315
         Left            =   340
         TabIndex        =   12
         Top             =   4100
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2805
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbBezah 
         Height          =   310
         Left            =   2200
         TabIndex        =   9
         Top             =   2000
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   220
         Left            =   2210
         TabIndex        =   26
         Top             =   1760
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Offen :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   220
         Left            =   350
         TabIndex        =   25
         Top             =   3860
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungsdatum :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   220
         Left            =   350
         TabIndex        =   24
         Top             =   3160
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Abgeschlossen :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   220
         Left            =   350
         TabIndex        =   23
         Top             =   1760
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungsversand :"
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Belegtyp :"
         Height          =   220
         Left            =   350
         TabIndex        =   22
         Top             =   360
         Width           =   1400
      End
      Begin VB.Label lblLab02 
         BackStyle       =   0  'Transparent
         Caption         =   "Untergrenze :"
         Height          =   220
         Left            =   2210
         TabIndex        =   21
         Top             =   3860
         Width           =   1200
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosen :"
         Height          =   220
         Left            =   350
         TabIndex        =   20
         Top             =   1060
         Width           =   1400
      End
      Begin VB.Label lblLab05 
         BackStyle       =   0  'Transparent
         Caption         =   "Zahlungsweise :"
         Height          =   220
         Left            =   350
         TabIndex        =   19
         Top             =   2460
         Width           =   1400
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   16000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin VB.Label Lab01 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte stellen Sie die gewünschten Kriterien ein und wählen, ob die Rechnungen markiert werden sollen."
      Height          =   440
      Left            =   700
      TabIndex        =   18
      Top             =   150
      Width           =   8000
   End
End
Attribute VB_Name = "frmReFilt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Bild1 As PictureBox
Private TreKy As XtremeSuiteControls.FlatEdit
Private TxGre As XtremeSuiteControls.FlatEdit
Private ChMaR As XtremeSuiteControls.CheckBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmDia As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmVer As XtremeSuiteControls.ComboBox
Private CmSta As XtremeSuiteControls.ComboBox
Private CmJah As XtremeSuiteControls.ComboBox
Private CmBez As XtremeSuiteControls.ComboBox
Private CmZah As XtremeSuiteControls.ComboBox
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private RpRow As XtremeReportControl.ReportRow
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox

Private KatGe As Boolean
Private ReMar As Boolean

Private clFil As clsFile
Private clFen As clsFenster
Private Sub GFilt()
On Error GoTo AnErr
'Filtert die Gruppen

Dim AktZa As Long
Dim ManNr As Long
Dim GruKy As String
Dim GrIdx As String
Dim ReTyp As String
Dim ReUnG As String
Dim Gedru As Integer
Dim ReDia As Integer
Dim Versa As Integer
Dim BuJah As Integer
Dim Bezah As Integer
Dim Zahlu As Integer
Dim RetWe As Boolean
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set TrLi1 = Me.trvList1
Set TxGre = Me.txtReGre
Set ChMaR = Me.chkMarki
Set CmTyp = Me.cmbReTyp
Set CmDia = Me.cmbDiagn
Set CmMan = Me.cmbBehan
Set CmVer = Me.cmbVersa
Set CmSta = Me.cmbStatu
Set CmJah = Me.cmbJahre
Set CmBez = Me.cmbBezah
Set CmZah = Me.cmbZahlu
Set RpCo4 = FM.repCont4

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

GlFil = GlFiX

If TxGre.Text <> vbNullString Then
    ReUnG = CLng(TxGre.Text)
Else
    ReUnG = 0
End If

If CmJah.ItemData(CmJah.ListIndex) < 18 Then
    BuJah = CmJah.Text
Else
    BuJah = 0
End If

If CmTyp.Text <> vbNullString Then
    Select Case CmTyp.ListIndex
    Case 0: ReTyp = "R"
    Case 1: ReTyp = "K"
    Case 2: ReTyp = "L"
    Case 3: ReTyp = "A"
    Case 4: ReTyp = "U"
    Case 5: ReTyp = "M"
    Case 6: ReTyp = "G"
    Case 7: ReTyp = "I"
    Case 8: ReTyp = "X"
    End Select
End If

For Each Knote In TrLi1.Nodes
    If Knote.Checked = True Then
        If Knote.Key <> "P801" Then
            GrIdx = Mid$(Knote.Key, 2, Len(Knote.Key) - 1)
            If Len(GruKy) = 0 Then
                GruKy = "o" & GrIdx & "o"
            Else
                GruKy = GruKy & GrIdx & "o"
            End If
        End If
    End If
Next Knote

If CmMan.ListIndex > -1 Then
    ManNr = CmMan.ItemData(CmMan.ListIndex)
End If

If CmVer.ListIndex > -1 Then
    Versa = CmVer.ItemData(CmVer.ListIndex)
End If

If CmBez.ListIndex > -1 Then
    Bezah = CmBez.ItemData(CmBez.ListIndex)
End If

If CmDia.ListIndex > -1 Then
    ReDia = CmDia.ItemData(CmDia.ListIndex)
End If

If CmSta.ListIndex > -1 Then
    Gedru = CmSta.ItemData(CmSta.ListIndex)
End If

If CmZah.ListIndex > -1 Then
    Zahlu = CmZah.ItemData(CmZah.ListIndex)
End If

With GlFil
    .Gedru = Gedru
    .ReDia = ReDia
    .ReTyp = ReTyp
    .ReUnG = ReUnG
    .MaNum = ManNr
    .Versa = Versa
    .ReJah = BuJah
    .Bezah = Bezah
    .Zahlu = Zahlu
    If GruKy <> vbNullString Then
        .GruKy = GruKy
    End If
End With
S_ReFi
DoEvents

If ChMaR.Value = 1 Then
    ReMar = True
End If

If KatGe = True Then
    IniSetVal "System", "ReFiMa", ReMar
End If

If ReMar = True Then
    Set RpRws = RpCo4.Rows
    For Each RpRow In RpRws
        RpRow.Selected = True
    Next RpRow
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RpRws = Nothing
Set RpCo4 = Nothing

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "GFilt " & Err.Number
Resume Next

End Sub
Private Sub GInit()
On Error GoTo InErr

Dim KatAn As Boolean
Dim AktZa As Integer
Dim BuJah As Integer
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmReFilt
Set TrLi1 = FM.trvList1
Set TxGre = FM.txtReGre
Set ChMaR = FM.chkMarki
Set CmTyp = FM.cmbReTyp
Set CmDia = FM.cmbDiagn
Set CmSta = Me.cmbStatu
Set CmMan = FM.cmbBehan
Set CmVer = Me.cmbVersa
Set CmJah = Me.cmbJahre
Set CmBez = Me.cmbBezah
Set CmZah = Me.cmbZahlu
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set ImMan = frmMain.imgManag

With TrLi1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .Checkboxes = True
    .Font.SIZE = GlTFt.SIZE
    .Font.Name = GlTFt.Name
    .ForeColor = -2147483641
    .FullRowSelect = False
    .HideSelection = False
    .HotTracking = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpTreeViewLabelManual
    .Scroll = True
    .ShowLines = xtpTreeViewShowLines
    .ShowPlusMinus = True
    .SingleSel = False
End With

With CmJah
    .DropDownItemCount = 12
    For BuJah = Year(Date) - 15 To Year(Date) + 1
        .AddItem BuJah
        .ItemData(.NewIndex) = AktZa
        AktZa = AktZa + 1
    Next BuJah
    .Text = Year(Date)
    .AddItem "Alle Jahre"
    .ItemData(.NewIndex) = AktZa + 1
End With

With CmTyp
    .AddItem "R - Standardrechnung"
    .ItemData(.NewIndex) = 1
    .AddItem "V - Kostenvoranschlag"
    .ItemData(.NewIndex) = 2
    .AddItem "L - Laborrechnung"
    .ItemData(.NewIndex) = 3
    .AddItem "A - Abrechnungsstelle"
    .ItemData(.NewIndex) = 4
    .AddItem "U - Gutschrift"
    .ItemData(.NewIndex) = 5
    .AddItem "M - Rechnungsauftrag"
    .ItemData(.NewIndex) = 6
    .AddItem "G - Gewerberechnung"
    .ItemData(.NewIndex) = 7
    .AddItem "I - Importrechnung"
    .ItemData(.NewIndex) = 8
    .AddItem "Alle Belegtypen"
    .ItemData(.NewIndex) = 9
    .ListIndex = 8
End With

With CmDia
    .AddItem "Rechnungen mit Diagnosen"
    .ItemData(.NewIndex) = 1
    .AddItem "Rechnungen ohne Diagnosen"
    .ItemData(.NewIndex) = 2
    .AddItem "Rechnungen mit und ohne Diagnosen"
    .ItemData(.NewIndex) = 3
    .ListIndex = 2
End With

With CmVer
    .AddItem "Postversand"
    .ItemData(0) = 0
    .AddItem "Emailversand"
    .ItemData(1) = 1
    .AddItem "Downloadlink"
    .ItemData(2) = 2
    .AddItem "Alle Versandarten"
    .ItemData(3) = 3
    .ListIndex = 3
End With

With CmBez
    .AddItem "Ja"
    .ItemData(.NewIndex) = 1
    .AddItem "Nein"
    .ItemData(.NewIndex) = 2
    .AddItem "Alle"
    .ItemData(.NewIndex) = 3
    .ListIndex = 2
End With

With CmSta
    .AddItem "Abgeschlossene Rechnungen"
    .ItemData(.NewIndex) = 1
    .AddItem "Unabgeschlossene Rechnungen"
    .ItemData(.NewIndex) = 2
    .AddItem "Un- und abgeschlossene Rechnungen"
    .ItemData(.NewIndex) = 3
    .ListIndex = 1
End With

With CmMan
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(.NewIndex) = GlThe(AktZa, 0)
    Next AktZa
End With

With CmMan
    .AddItem "Alle Mandanten"
    .ItemData(.NewIndex) = 0
    .ListIndex = AktZa - 1
End With

With CmZah
    For AktZa = 1 To UBound(GlZah)
        .AddItem GlZah(AktZa, 1)
        .ItemData(.NewIndex) = GlZah(AktZa, 0)
    Next AktZa
End With

With CmZah
    .AddItem "Alle Zahlungsweisen"
    .ItemData(.NewIndex) = 0
    .ListIndex = AktZa - 1
End With

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

Set Knote = TrLi1.Nodes.Add(, , "P801", "Adressen", IC16_Folder_View)
With Knote
    .Bold = True
    .Checked = False
    .Expanded = True
End With

KatAn = CBool(IniGetVal("System", "ReFiMa"))
TxGre.Text = Format$(IniGetVal("System", "ReUnGr"), GlWa1)

If KatAn = True Then
    ChMaR.Value = 1
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
ChMaR.BackColor = GlBak

Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "GInit " & Err.Number
Resume Next

End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50831)
TeMai = IniGetOpt("Hilfe", 50832)
TeInh = IniGetOpt("Hilfe", 50833)
TeFus = IniGetOpt("Hilfe", 50834)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()

Set ChMaR = Me.chkMarki

If ChMaR.Value = 1 Then
    ReMar = True
End If

If KatGe = True Then
    IniSetVal "System", "ReFiMa", ReMar
End If

Unload Me

End Sub
Private Sub btnWeiter_Click()
    GFilt
    Unload Me
End Sub
Private Sub chkMarki_Click()
    KatGe = True
End Sub

Private Sub Form_Load()
On Error Resume Next

GInit
AFont Me
AdGru 2
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmReFilt = Nothing
End Sub

Private Sub trvList1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo LaErr

Set TrLi1 = Me.trvList1

If Button = vbRightButton Then
    Set TrLi1.SelectedItem = TrLi1.HitTest(x, y)
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub trvList1_NodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Dim TreKy As String

Set TrLi1 = Me.trvList1
    
For Each Knote In TrLi1.Nodes
    Knote.Image = IC16_Folder_Close
Next Knote

Node.Image = IC16_Folder_Open
TrLi1.Nodes(1).Image = IC16_Folder_View

If Node.Key = "P801" Then
    For Each Knote In TrLi1.Nodes
        Knote.Checked = Node.Checked
    Next Knote
End If

End Sub
Private Sub trvList1_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Dim TreKy As String

Set TrLi1 = Me.trvList1
    
For Each Knote In TrLi1.Nodes
    Knote.Image = IC16_Folder_Close
Next Knote

Node.Image = IC16_Folder_Open
TrLi1.Nodes(1).Image = IC16_Folder_View

End Sub

Private Sub txtReGre_GotFocus()
    Me.txtReGre.SelStart = 0
    Me.txtReGre.SelLength = Len(Me.txtReGre.Text)
End Sub


