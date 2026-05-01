VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmKaKop 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kopieren"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   2
      Top             =   4400
      Width           =   5600
      _Version        =   1048579
      _ExtentX        =   9878
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   405
         Left            =   3600
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "&Abbrechen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   2200
         TabIndex        =   4
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
         Left            =   900
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
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5600
      _Version        =   1048579
      _ExtentX        =   9878
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   3400
         Left            =   300
         TabIndex        =   1
         Top             =   700
         Width           =   4900
         _Version        =   1048579
         _ExtentX        =   8643
         _ExtentY        =   5997
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         HideSelection   =   0   'False
         BackColor       =   16777215
         ForeColor       =   4473924
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte markieren Sie die Kataloge, in die Sie die markieren Einträge kopieren möchten und klicken auf Weiter"
         Height          =   580
         Left            =   400
         TabIndex        =   6
         Top             =   100
         Width           =   4600
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   7
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
Attribute VB_Name = "frmKaKop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private TrLi1 As XtremeSuiteControls.TreeView
Private TrLi2 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private Sub FOpen()
On Error GoTo PeErr
'Stellt Baumstruktur im Katalogexplorer zusammen

Dim TreKy As String
Dim DatKy As String
Dim KnoKy As String
Dim VorKy As String
Dim TmpKy As String
Dim AkPos As Integer
Dim AltPo As Integer
Dim AktZa As Integer
Dim LauZa As Integer
Dim GrGef As Boolean

TreKy = Left$(GlNod, 1)

Set TrLi2 = Me.trvList1

If GlBut = RibTab_Tex_Email Then
    With TrLi2
        Set Knote = .Nodes.Add(, xtpTreeViewFirst, "P800", "Emailgruppen", IC16_Mailbox)
        Knote.Bold = True
        Set Knote = .Nodes.Add("P800", 4, "P801", "Emailhauptordner", IC16_Folder_View)
        Set Knote = .Nodes.Add("P800", 4, "P802", "Mitarbeiter", IC16_Folder_Edit)
        Set Knote = .Nodes.Add("P800", 4, "P803", "Rechnungen", IC16_Folder_Check)
        Set Knote = .Nodes.Add("P800", 4, "P804", "Telefaxe", IC16_Folder_Paper)
        .Nodes("P800").Expanded = True
    End With
    DoEvents
    
    AktZa = TrLi2.Nodes.Count + 1

    For LauZa = 1 To UBound(GlEmG) 'Emailgruppen
        If GlEmG(LauZa, 2) <> vbNullString Then
            DatKy = GlEmG(LauZa, 2)
        Else
            DatKy = 0
        End If
        KnoKy = "G" & GlEmG(LauZa, 0)
        AkPos = InStrRev(DatKy, ".", Len(DatKy), 1)
        If AkPos > 0 Then
            If AkPos > AltPo Then
                VorKy = TrLi2.Nodes(AktZa - 1).Key
            ElseIf AkPos < AltPo Then
                VorKy = TrLi2.Nodes(AktZa - 1).Parent.Parent.Key
            End If
        Else
            VorKy = "P801"
        End If
        If GlEmG(LauZa, 1) <> vbNullString Then
            Set Knote = TrLi2.Nodes.Add(VorKy, 4, KnoKy, GlEmG(LauZa, 1), IC16_Folder_Close)
        Else
            Set Knote = TrLi2.Nodes.Add(VorKy, 4, KnoKy, "-", IC16_Folder_Close)
        End If
        If Knote.Key = TmpKy Then
            Knote.Selected = True
            GrGef = True
        End If
        If GrGef = False Then TrLi2.Nodes("P801").Selected = True
        AltPo = AkPos
        AktZa = AktZa + 1
    Next LauZa
    
    For AktZa = 1 To UBound(GlMiK) 'Mitarbeiter
        Set Knote = TrLi2.Nodes.Add("P802", 4, "G" & Format$(AktZa + 900, "000"), GlMiK(AktZa, 1), IC16_Folder_Close)
    Next AktZa
    
    With TrLi2
        .Nodes("P800").Expanded = True
        .Nodes("P801").Expanded = True
        .Nodes("P802").Expanded = GlMOr
    End With

    If GlExp = True Then
        For Each Knote In TrLi2.Nodes
            If Left$(Knote.Key, 1) = "G" Then
                Knote.Expanded = True
            End If
        Next Knote
    End If
Else
    Select Case TreKy
    Case "A": 'Gebührenkatalog
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "A00", "Gebührenkataloge", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlGKa)
            Set Knote = TrLi2.Nodes.Add("A00", 4, "A" & GlGKa(AktZa, 0), GlGKa(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "C": 'Diagnosen
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "C00", "Diagnosekataloge", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlDia)
            Set Knote = TrLi2.Nodes.Add("C00", 4, "C" & GlDia(AktZa, 0), GlDia(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "D": 'Gebührenketten
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "D00", "Gebührenketten", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlGKa)
            Set Knote = TrLi2.Nodes.Add("D00", 4, "D" & GlGKa(AktZa, 0), GlGKa(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "F": 'Diagnoseketten
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "F00", "Diagnoseketten", IC16_Folder_Close)
        Set Knote = TrLi2.Nodes.Add("F00", 4, "F1", "Alle Diagnosegruppen", IC16_Folder_Close)
    Case "G": 'Laborparameter
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "G00", "Laborparameter", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlLab)
            Set Knote = TrLi2.Nodes.Add("G00", 4, "G" & GlLab(AktZa, 0), GlLab(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "H": 'Laborketten
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "H00", "Laborparameter", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlLab)
            Set Knote = TrLi2.Nodes.Add("H00", 4, "H" & GlLab(AktZa, 0), GlLab(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "I": 'Arzneimittel
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "I00", "Arzneimittel", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlMed) '
            Set Knote = TrLi2.Nodes.Add("I00", 4, "I" & GlMed(AktZa, 0), GlMed(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "J": 'Arzneiketten
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "J00", "Arzneiketten", IC16_Folder_Close)
        Set Knote = TrLi2.Nodes.Add("J00", 4, "J1", "Alle Arzneigruppen", IC16_Folder_Close)
    Case "K": 'Begründungen
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "K00", "Begründungen", IC16_Folder_Close)
        Set Knote = TrLi2.Nodes.Add("K00", 4, "K3", "Alle Begründungen", IC16_Folder_Close)
    Case "L": 'Anamnesetexte
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "L00", "Anamnesetexte", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlAnG)
            Set Knote = TrLi2.Nodes.Add("L00", 4, "L" & GlAnG(AktZa, 0), GlAnG(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "M": 'Terminbetreffs
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "M00", "Terminbetreffs", IC16_Folder_Close)
        Set Knote = TrLi2.Nodes.Add("K00", 4, "M5", "Alle Terminbetreffs", IC16_Folder_Close)
    Case "N": 'Fragebogen
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "N00", "Anamnebogen", IC16_Folder_Close)
        If GlBoV > 0 Then 'Fragebogen vorhanden
            For AktZa = 1 To GlBoV
                Set Knote = TrLi2.Nodes.Add("N00", 4, "N" & GlFrB(AktZa, 0), GlFrB(AktZa, 1), IC16_Folder_Close)
            Next AktZa
        End If
    Case "O": 'Textphrasenkatalog
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "O00", "Textphrasenkatalog", IC16_Folder_Close)
        Set Knote = TrLi2.Nodes.Add("O00", 4, "O9", "Alle Textphrasen", IC16_Folder_Close)
    Case "P": 'Artikelliste
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "P00", "Artikelliste", IC16_Folder_Close)
        For AktZa = 1 To UBound(GlArt) '
            Set Knote = TrLi2.Nodes.Add("P00", 4, "P" & GlArt(AktZa, 0), GlArt(AktZa, 1), IC16_Folder_Close)
        Next AktZa
    Case "R": 'Terminketten
        Set Knote = TrLi2.Nodes.Add("Z00", 4, "R00", "Terminketten", IC16_Folder_Close)
        Set Knote = TrLi2.Nodes.Add("J00", 4, "R1", "Alle Terminketten", IC16_Folder_Close)
    Case "Q": 'Artikelketten
        Set Knote = TrLi2.Nodes.Add("Z00", 5, "Q00", "Artikelketten", IC16_Folder_Close)
        Set Knote = TrLi2.Nodes.Add("J00", 5, "Q1", "Alle Artikelketten", IC16_Folder_Close)
    End Select
    
    With TrLi2
        .Nodes(GlNod).Expanded = True
        .Nodes(GlNod).Selected = True
        .Nodes(GlNod).Checked = True
        .Nodes(GlNod).EnsureVisible
    End With
End If

Exit Sub

PeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpen " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmKaKop
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set TrLi1 = FM.trvList1
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
    .SingleSel = True
End With

Set Knote = TrLi1.Nodes.Add(, , "Z00", "Kataloge", IC16_Folder_View)
With Knote
    .Bold = True
    .Checked = False
End With

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak

Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TInit " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo SuErr

Dim RowNr As Long
Dim KeySt As String

Set FM = frmKaKop
Set TrLi1 = FM.trvList1

If GlBut = RibTab_Tex_Email Then
    For Each Knote In TrLi1.Nodes
        If Knote.Checked = True Then
            KeySt = Knote.Key
            Exit For
        End If
    Next Knote
    If KeySt <> vbNullString Then
        GlSuI = GlSuX 'Suchkriterien Emails
        RowNr = GrMa_Ord(KeySt)
        DoEvents
        SUpMa RowNr
    End If
Else
    KeySt = Left$(TrLi1.SelectedItem.Key, 1)
    Select Case KeySt
    Case "A": K_Kop1        'Gebührenkatalog
    Case "C": K_Kop1        'Diagnosekatalog
    Case "D": E_KeKo KeySt  'Gebührenketten
    Case "F": E_KeKo KeySt  'Diagnoseketten
    Case "G": K_Kop1        'Laborparameter
    Case "H": E_KeKo KeySt  'Laborprofile
    Case "I": K_Kop1        'Arzneikatalog
    Case "J": E_KeKo KeySt  'Arzneiketten
    Case "K": K_Kop1        'Begründungen
    Case "L": K_Kop1        'Anamnesetexte
    Case "M": K_Kop1        'Terminbetreffs
    Case "N": K_Kop1        'Fragebogen
    Case "O": K_Kop1        'Textphrasen
    Case "P": K_Kop1        'Artikelkatalog
    Case "R": E_KeKo KeySt  'Terminketten
    Case "Q": E_KeKo KeySt  'Artikelketten
    End Select
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
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
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
    Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
FOpen
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKaKop = Nothing
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

If Right$(Node.Key, 2) = "00" Then
    For Each Knote In TrLi1.Nodes
        If Right$(Knote.Key, 2) <> "00" Then
            Knote.Checked = Node.Checked
        End If
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
