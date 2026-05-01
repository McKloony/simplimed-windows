VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Importieren"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   9
      Top             =   4800
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         Left            =   1700
         TabIndex        =   10
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
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4100
      Left            =   300
      TabIndex        =   2
      Top             =   700
      Width           =   5740
      _Version        =   1048579
      _ExtentX        =   10125
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbForma 
         Height          =   310
         Left            =   1300
         TabIndex        =   3
         Top             =   440
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
         DropDownItemCount=   0
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   310
         Left            =   1300
         TabIndex        =   4
         Top             =   1160
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
      Begin XtremeSuiteControls.CheckBox chkReAbs 
         Height          =   225
         Left            =   1300
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3200
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Rechnungen abschließen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbReTyp 
         Height          =   315
         Left            =   1300
         TabIndex        =   6
         Top             =   2600
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
      Begin XtremeSuiteControls.CheckBox chkOrReN 
         Height          =   225
         Left            =   1300
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Rechnungsnummer Importieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   1300
         TabIndex        =   5
         Top             =   1880
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
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   220
         Left            =   1320
         TabIndex        =   18
         Top             =   1640
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mitarbeiter :"
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   220
         Left            =   1300
         TabIndex        =   17
         Top             =   200
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Importformat :"
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   220
         Left            =   1300
         TabIndex        =   16
         Top             =   920
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandant :"
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   220
         Left            =   1320
         TabIndex        =   15
         Top             =   2340
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungstyp :"
      End
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6100
      Width           =   80
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4100
      Left            =   300
      TabIndex        =   13
      Top             =   700
      Visible         =   0   'False
      Width           =   5440
      _Version        =   1048579
      _ExtentX        =   9596
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   2980
         Left            =   700
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   4100
         _Version        =   1048579
         _ExtentX        =   7232
         _ExtentY        =   5256
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         HideSelection   =   0   'False
      End
   End
   Begin VB.Label lblLabe1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte wählen Sie, aus welchem Format die  Einträge importiert werden sollen und klicken dann auf Weiter."
      Height          =   435
      Left            =   1000
      TabIndex        =   1
      Top             =   135
      Width           =   4500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   700
      Left            =   0
      Top             =   0
      Width           =   6400
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Lbl01 As XtremeSuiteControls.Label
Private Lbl02 As XtremeSuiteControls.Label
Private Lbl03 As XtremeSuiteControls.Label
Private Lbl04 As XtremeSuiteControls.Label
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private CheAb As XtremeSuiteControls.CheckBox
Private CheNu As XtremeSuiteControls.CheckBox
Private CmFma As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows

Public FiNam As String
Private FoLad As Boolean
Private Sub FLoad()
On Error GoTo LdErr

Dim ManNr As Long
Dim StaGe As Long
Dim ThIdx As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim LiIdx As Integer
Dim OrReN As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmThe As XtremeCommandBars.CommandBarComboBox
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmMain
Set CmBrs = FM.comBar01
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Lbl01 = Me.lblLab01
Set Lbl02 = Me.lblLab02
Set Lbl03 = Me.lblLab03
Set Lbl04 = Me.lblLab04
Set TrLi1 = Me.trvList1
Set CmFma = Me.cmbForma
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set CmTyp = Me.cmbReTyp
Set CheAb = Me.chkReAbs
Set CheNu = Me.chkOrReN
Set ImMan = FM.imgManag

OrReN = CBool(IniGetVal("System", "ReNuIm"))

Set CmThe = CmBrs.FindControl(CmThe, SY_SuMan, , True)
ThIdx = CmThe.ListIndex

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

For AktZa = 1 To UBound(GlThe)
    With CmMan
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa - 1) = GlThe(AktZa, 0)
    End With
Next AktZa
If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    ManNr = GlMiA(GlSmI, 7)
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then 'Mandantennummer
            CmMan.ListIndex = AktZa - 1
        End If
    Next AktZa
Else
    If CmThe.Visible = True Then
        CmMan.ListIndex = GlMan(ThIdx, 0) - 1
    Else
        CmMan.ListIndex = GlSMa - 1
    End If
End If

For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
    With CmMit
        .AddItem GlMiA(AktZa, 1)
        .ItemData(AktZa - 1) = GlMiA(AktZa, 2)
    End With
Next AktZa
CmMit.ListIndex = GlSmI - 1

Select Case GlBut
Case RibTab_Rechnungen:
    Lbl03.Caption = "Belegtyp :"
    With CmTyp
        .AddItem "R - Standardrechnung"
        .ItemData(0) = 1
        .AddItem "V - Kostenvoranschlag"
        .ItemData(1) = 2
        .AddItem "L - Laborrechnung"
        .ItemData(2) = 3
        .AddItem "A - Abrechnungsstelle"
        .ItemData(3) = 4
        .AddItem "U - Gutschrift"
        .ItemData(4) = 5
        .AddItem "M - Rechnungsauftrag"
        .ItemData(5) = 6
        .AddItem "G - Gewerberechnung"
        .ItemData(6) = 7
        .AddItem "I - Importrechnung"
        .ItemData(7) = 8
        .ListIndex = 7
    End With
Case RibTab_HomeBanki:
    Lbl03.Caption = "Geldkonto :"
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            CmTyp.AddItem GlGeK(AktZa, 3)
            CmTyp.ItemData(AktZa - 1) = GlGeK(AktZa, 0)
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    CmTyp.AddItem GlSaK(AktKo, 3)
                    CmTyp.ItemData(AktZa - 1) = GlSaK(AktKo, 0) '[IDB]
                    Exit For
                End If
            Next AktKo
        Next AktZa
        If CmTyp.ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                CmTyp.AddItem GlGeK(AktZa, 3)
                CmTyp.ItemData(AktZa - 1) = GlGeK(AktZa, 0)
            Next AktZa
        End If
    End If
    If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
        For AktZa = 1 To UBound(GlMan)
            If ManNr = GlMan(AktZa, 2) Then
                If GlMan(AktZa, 28) <> vbNullString Then
                    If GlMan(AktZa, 28) > 0 Then
                        StaGe = GlMan(AktZa, 28) 'Standardgeldkonto Bankkonto
                    Else
                        StaGe = GlGkB
                    End If
                Else
                    StaGe = CInt(GlSet(2, 26))
                End If
                StaGe = SCmb(CmTyp, StaGe)
                CmTyp.ListIndex = StaGe
                Exit For
            End If
        Next AktZa
    Else
        StaGe = SCmb(CmTyp, GlGkB) 'Standardgeldkonto Bankkonto
        If StaGe >= 0 Then
            CmTyp.ListIndex = StaGe
        Else
            CmTyp.ListIndex = 0
        End If
    End If
    CheAb.Visible = False
    CheNu.Visible = False
Case Else:
    Lbl03.Visible = False
    CmTyp.Visible = False
    CheAb.Visible = False
    CheNu.Visible = False
End Select

Select Case GlBut
Case RibTab_Adressen:
    With CmFma
        .AddItem "Acsii Textdatei (*.txt)"
        .ItemData(.NewIndex) = 1
        .AddItem "Outlook Kontakte (*.psx)"
        .ItemData(.NewIndex) = 2
        .AddItem "XML-Datendatei (*.xml)"
        .ItemData(.NewIndex) = 3
        .AddItem "BDT Textdatei (*.bdt)"
        .ItemData(.NewIndex) = 4
        .AddItem "GDT Textdatei (*.gdt)"
        .ItemData(.NewIndex) = 5
        .AddItem "THEDEX-Dateien (*.thx)"
        .ItemData(.NewIndex) = 6
        .AddItem "xBDT Textdatei (*.bdt)"
        .ItemData(.NewIndex) = 7
        .AddItem "SMA Stammdaten (*.sma)"
        .ItemData(.NewIndex) = 8
        .ListIndex = 0
    End With
Case RibTab_Mandanten:
    With CmFma
        .AddItem "Acsii Textdatei (*.txt)"
        .ItemData(.NewIndex) = 1
        .AddItem "Outlook Kontakte (*.psx)"
        .ItemData(.NewIndex) = 2
        .AddItem "XML-Datendatei (*.xml)"
        .ItemData(.NewIndex) = 3
        .AddItem "BDT Textdatei (*.bdt)"
        .ItemData(.NewIndex) = 4
        .AddItem "GDT Textdatei (*.gdt)"
        .ItemData(.NewIndex) = 5
        .AddItem "THEDEX-Dateien (*.thx)"
        .ItemData(.NewIndex) = 6
        .ListIndex = 0
    End With
Case RibTab_Verordner:
    With CmFma
        .AddItem "Acsii Textdatei (*.txt)"
        .ItemData(.NewIndex) = 1
        .AddItem "Outlook Kontakte (*.psx)"
        .ItemData(.NewIndex) = 2
        .AddItem "XML-Datendatei (*.xml)"
        .ItemData(.NewIndex) = 3
        .AddItem "BDT Textdatei (*.bdt)"
        .ItemData(.NewIndex) = 4
        .AddItem "GDT Textdatei (*.gdt)"
        .ItemData(.NewIndex) = 5
        .AddItem "THEDEX-Dateien (*.thx)"
        .ItemData(.NewIndex) = 6
        .ListIndex = 0
    End With
Case RibTab_Mitarbeit:
    With CmFma
        .AddItem "Acsii Textdatei (*.txt)"
        .ItemData(.NewIndex) = 1
        .AddItem "Outlook Kontakte (*.psx)"
        .ItemData(.NewIndex) = 2
        .AddItem "XML-Datendatei (*.xml)"
        .ItemData(.NewIndex) = 3
        .AddItem "BDT Textdatei (*.bdt)"
        .ItemData(.NewIndex) = 4
        .AddItem "GDT Textdatei (*.gdt)"
        .ItemData(.NewIndex) = 5
        .AddItem "THEDEX-Dateien (*.thx)"
        .ItemData(.NewIndex) = 6
        .ListIndex = 0
    End With
Case RibTab_HomeBanki:
    LiIdx = IniGetVal("System", "BaImFo")
    With CmFma
        .AddItem "Comma-Separated Values (*.csv)"
        .ItemData(.NewIndex) = 0
        .AddItem "Textdatei (*.txt)"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
Case RibTab_Ter_Listen:
    With CmFma
        .AddItem "iCalendar Datei (*.ics)"
        .ItemData(.NewIndex) = 1
        .AddItem "Outlook Kontakte (*.psx)"
        .ItemData(.NewIndex) = 2
        .AddItem "SimpliMed Termine (*.smt)"
        .ItemData(.NewIndex) = 3
        .ListIndex = 0
    End With
Case RibTab_Ter_Akont:
    With CmFma
        .AddItem "iCalendar Datei (*.ics)"
        .ItemData(.NewIndex) = 1
        .AddItem "Outlook Kontakte (*.psx)"
        .ItemData(.NewIndex) = 2
        .AddItem "SimpliMed Termine (*.smt)"
        .ItemData(.NewIndex) = 3
        .ListIndex = 0
    End With
Case Else:
    With CmFma
        .AddItem "SMP-Datendatei (*.smp)"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    CheNu.Value = OrReN
End Select

Set Knote = TrLi1.Nodes.Add(, , "P801", "Adressen", IC16_Folder_View)
With Knote
    .Bold = True
    .Checked = False
    .Expanded = True
End With

With CmFma
    .DropDownItemCount = 28 'WICHIG!
    .MaxLength = 12
End With

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
CheAb.BackColor = GlBak
CheNu.BackColor = GlBak
Lbl01.BackColor = GlBak
Lbl02.BackColor = GlBak
Lbl03.BackColor = GlBak
Lbl04.BackColor = GlBak

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FImpo()
On Error GoTo LdErr
'Importieren

Dim ManNr As Long
Dim MitNr As Long
Dim GruKy As String
Dim GrIdx As String
Dim TyStr As String
Dim ImFmt As String
Dim ImpFo As Integer
Dim ImTyp As Integer
Dim IdBnk As Integer
Dim ReAbs As Boolean
Dim OrReN As Boolean

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set CmFma = Me.cmbForma
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set CmTyp = Me.cmbReTyp
Set CheAb = Me.chkReAbs
Set CheNu = Me.chkOrReN

Select Case GlBut
Case RibTab_Adressen:
    If Rahm1.Visible = True Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        Exit Sub
    End If
Case RibTab_Mandanten:
    If Rahm1.Visible = True Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        Exit Sub
    End If
Case RibTab_Verordner:
    If Rahm1.Visible = True Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        Exit Sub
    End If
Case RibTab_Mitarbeit:
    If Rahm1.Visible = True Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        Exit Sub
    End If
End Select

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

ManNr = CmMan.ItemData(CmMan.ListIndex)
MitNr = CmMit.ItemData(CmMit.ListIndex)

ImpFo = CmFma.ListIndex
ImTyp = CmTyp.ListIndex

If CheAb.Value = xtpChecked Then ReAbs = True
If CheNu.Value = xtpChecked Then OrReN = True

Select Case GlBut
Case RibTab_Rechnungen:
        Select Case ImTyp
        Case 0: TyStr = "R"
        Case 1: TyStr = "V"
        Case 2: TyStr = "L"
        Case 3: TyStr = "A"
        Case 4: TyStr = "U"
        Case 5: TyStr = "M"
        Case 6: TyStr = "G"
        Case 7: TyStr = "I"
        End Select
Case RibTab_HomeBanki:
        IdBnk = CmTyp.ItemData(ImTyp)
End Select

If ManNr = 0 Then ManNr = GlMan(GlSMa, 2)
If MitNr = 0 Then MitNr = GlMiA(GlSmI, 2)

Unload Me
DoEvents

Select Case GlBut
Case RibTab_Adressen:
    Select Case ImpFo
    Case 0: S_AdIm "txt", ManNr, GruKy
    Case 1: S_AdIm "psx", ManNr
    Case 2: S_AdIm "xml", ManNr
    Case 3: SBDT
    Case 4: S_GDT
    Case 5: SAdTx
    Case 6: SBDT True
    Case 7: S_PaImT ManNr, MitNr
    End Select
Case RibTab_Mandanten:
    Select Case ImpFo
    Case 0: S_AdIm "txt", ManNr, GruKy
    Case 1: S_AdIm "psx", ManNr
    Case 2: S_AdIm "xml", ManNr
    Case 3: SBDT
    Case 4: S_GDT
    Case 5: SAdTx
    Case 6: SBDT True
    End Select
Case RibTab_Verordner:
    Select Case ImpFo
    Case 0: S_AdIm "txt", ManNr, GruKy, True
    Case 1: S_AdIm "psx", ManNr
    Case 2: S_AdIm "xml", ManNr
    Case 3: SBDT
    Case 4: S_GDT
    Case 5: SAdTx
    Case 6: SBDT True
    End Select
Case RibTab_Mitarbeit:
    Select Case ImpFo
    Case 0: S_AdIm "txt", ManNr, GruKy
    Case 1: S_AdIm "psx", ManNr
    Case 2: S_AdIm "xml", ManNr
    Case 3: SBDT
    Case 4: S_GDT
    Case 5: SAdTx
    Case 6: SBDT True
    End Select
Case RibTab_Ter_Listen:
    Select Case ImpFo
    Case 0: S_TeImI ManNr, MitNr
    Case 1: S_TeImO ManNr, MitNr
    Case 2: S_TeImT ManNr, MitNr
    End Select
Case RibTab_Ter_Akont:
    Select Case ImpFo
    Case 0: S_TeImI ManNr, MitNr
    Case 1: S_TeImO ManNr, MitNr
    Case 2: S_TeImT ManNr, MitNr
    End Select
Case RibTab_Ter_Warte:

Case RibTab_HomeBanki:
    Imp01 IdBnk, ManNr, MitNr
Case Else:
    Select Case ImpFo
    Case 0: S_ReImT ManNr, MitNr, FiNam, TyStr, ReAbs, OrReN
    End Select
End Select

Exit Sub

LdErr:
If GlDbg = True Then
    MsgBox Err.Description, 48, "FImpo " & Err.Number
    SErLog Err.Description & " FImpo " & Err.Number & " - " & Err.Source
End If
Resume Next

End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case GlBut
Case RibTab_Adressen:
    TeTit = IniGetOpt("Hilfe", 50251)
    TeMai = IniGetOpt("Hilfe", 50252)
    TeInh = IniGetOpt("Hilfe", 50253)
    TeFus = IniGetOpt("Hilfe", 50254)
Case RibTab_Abrechnung:
    TeTit = IniGetOpt("Hilfe", 50261)
    TeMai = IniGetOpt("Hilfe", 50262)
    TeInh = IniGetOpt("Hilfe", 50263)
    TeFus = IniGetOpt("Hilfe", 50264)
Case RibTab_Rechnungen:
    TeTit = IniGetOpt("Hilfe", 50271)
    TeMai = IniGetOpt("Hilfe", 50272)
    TeInh = IniGetOpt("Hilfe", 50273)
    TeFus = IniGetOpt("Hilfe", 50274)
Case RibTab_HomeBanki:
    TeTit = IniGetOpt("Hilfe", 50281)
    TeMai = IniGetOpt("Hilfe", 50282)
    TeInh = IniGetOpt("Hilfe", 50283)
    TeFus = IniGetOpt("Hilfe", 50284)
Case RibTab_Ter_Listen:
    TeTit = IniGetOpt("Hilfe", 50291)
    TeMai = IniGetOpt("Hilfe", 50292)
    TeInh = IniGetOpt("Hilfe", 50293)
    TeFus = IniGetOpt("Hilfe", 50294)
Case Else:
    TeTit = IniGetOpt("Hilfe", 50251)
    TeMai = IniGetOpt("Hilfe", 50252)
    TeInh = IniGetOpt("Hilfe", 50253)
    TeFus = IniGetOpt("Hilfe", 50254)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FImpo
End Sub

Private Sub chkOrReN_Click()
On Error Resume Next

Set CheNu = Me.chkOrReN

If CheNu.Value = xtpChecked Then
    IniSetVal "System", "ReNuIm", -1
Else
    IniSetVal "System", "ReNuIm", 0
End If

End Sub

Private Sub cmbForma_Click()
On Error Resume Next

Dim LiIdx As Integer

Set CmFma = Me.cmbForma

If FoLad = False Then
    If GlBut = RibTab_HomeBanki Then
        LiIdx = CmFma.ListIndex
        IniSetVal "System", "BaImFo", LiIdx
    End If
End If

End Sub

Private Sub cmbManda_Click()
On Error GoTo LdErr

Dim StaGe As Long
Dim ManNr As Long
Dim AktZa As Integer

Set CmMan = Me.cmbManda
Set CmTyp = Me.cmbReTyp

ManNr = CmMan.ItemData(CmMan.ListIndex)

If FoLad = False Then
    If GlBut <> RibTab_Ter_Listen Then
        If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
            For AktZa = 1 To UBound(GlMan)
                If ManNr = GlMan(AktZa, 2) Then
                    If GlMan(AktZa, 28) <> vbNullString Then
                        If GlMan(AktZa, 28) > 0 Then
                            StaGe = GlMan(AktZa, 28) 'Standardgeldkonto Bankkonto
                        ElseIf GlMan(AktZa, 26) > 0 Then
                            StaGe = CInt(GlSet(2, 26))
                        Else
                            StaGe = GlGkB
                        End If
                    ElseIf GlMan(AktZa, 26) <> vbNullString Then
                        If GlMan(AktZa, 26) > 0 Then
                            StaGe = CInt(GlSet(2, 26))
                        Else
                            StaGe = GlGkB
                        End If
                    End If
                    StaGe = SCmb(CmTyp, StaGe)
                    CmTyp.ListIndex = StaGe
                    Exit For
                End If
            Next AktZa
        Else
            If CmTyp.ListCount > 0 Then
                StaGe = SCmb(CmTyp, GlGkB) 'Standardgeldkonto Bankkonto
                If StaGe > 0 Then
                    CmTyp.ListIndex = StaGe
                Else
                    CmTyp.ListIndex = 0
                End If
            End If
        End If
    End If
End If

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "cmbManda " & Err.Number
Resume Next

End Sub

Private Sub Form_Load()
On Error Resume Next

FoLad = True
FLoad
AdGru 7
FoLad = False
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmImport = Nothing
End Sub
Private Sub trvList1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo LaErr

Set TrLi1 = Me.trvList1

If Button = vbRightButton Then
    Set TrLi1.SelectedItem = TrLi1.HitTest(x, Y)
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

