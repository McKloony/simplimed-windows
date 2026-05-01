VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmChipcard 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Chipkarte Einlesen"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   7
      Top             =   2600
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
         TabIndex        =   11
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
         TabIndex        =   10
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
      Begin XtremeSuiteControls.PushButton btnZurück 
         Height          =   400
         Left            =   1600
         TabIndex        =   9
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Neue Adresse"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.TextBox txoDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   80
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   2500
      Left            =   200
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   4410
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnPictu 
         Height          =   2100
         Left            =   200
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   200
         Width           =   2100
         _Version        =   1048579
         _ExtentX        =   3704
         _ExtentY        =   3704
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   6
         DrawFocusRect   =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cmbChipd 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   480
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4763
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   12
      End
      Begin XtremeSuiteControls.FlatEdit txtMeldu 
         Height          =   1320
         Left            =   2700
         TabIndex        =   6
         Top             =   920
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   2328
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         MultiLine       =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   240
         Left            =   2700
         TabIndex        =   5
         Top             =   200
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Chipkartenleser :"
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   2500
      Left            =   200
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   4410
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView3 
         Height          =   2100
         Left            =   300
         TabIndex        =   12
         Top             =   200
         Width           =   5300
         _Version        =   1048579
         _ExtentX        =   9349
         _ExtentY        =   3704
         _StockProps     =   77
         BackColor       =   -2147483643
         View            =   2
      End
   End
End
Attribute VB_Name = "frmChipcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Pict1 As VB.PictureBox
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private Lbl01 As XtremeSuiteControls.Label
Private CmChp As XtremeSuiteControls.ComboBox
Private TxMel As XtremeSuiteControls.FlatEdit
Private CmBar As XtremeCommandBars.CommandBar
Private CmAcs As XtremeCommandBars.CommandBarActions
Private ImMan As XtremeCommandBars.ImageManager
Private LiVw3 As XtremeSuiteControls.ListView
Private LiItm As XtremeSuiteControls.ListViewItem
Private LiIts As XtremeSuiteControls.ListViewItems

Private clFen As clsFenster
Private clChe As clsChipcard
Private Sub FChip()
On Error GoTo CrErr
'Liest die Chipkarte ein

Dim RetBy As Byte
Dim RetWe As Long
Dim GesZa As Long
Dim PatGe As String
Dim RetSt As String
Dim KaGul As String
Dim StrFe As String
Dim KasNa As String
Dim KasNr As String
Dim KarNr As String
Dim VerNr As String
Dim KaSta As String
Dim StaEr As String
Dim PaTit As String
Dim PaVor As String
Dim PaZun As String
Dim PaNam As String
Dim PaGeb As String
Dim PaGes As String
Dim PaStr As String
Dim PaLKZ As String
Dim PaPLZ As String
Dim PaOrt As String
Dim KaDat As String
Dim SuStr As String
Dim RetAb As Integer

Set ImMan = FM.imgManag
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set CmChp = Me.cmbChipd
Set TxMel = Me.txtMeldu
Set Lbl01 = Me.lblLab01
Set LiVw3 = Me.lstView3

Set LiIts = LiVw3.ListItems

Set clChe = New clsChipcard

GesZa = 0
StrFe = vbNullString

Screen.MousePointer = vbHourglass

TxMel.Text = vbNullString

If GlChp = 0 Then
    TxMel.Text = TxMel.Text & vbCrLf & "Es wurde kein Chipkartenlesegerät ausgewählt!" & vbCrLf
Else
    KasNa = 0
    StrFe = vbNullString
    KasNa = vbNullString
    KasNr = vbNullString
    KarNr = vbNullString
    VerNr = vbNullString
    KaSta = vbNullString
    StaEr = vbNullString
    PaTit = vbNullString
    PaVor = vbNullString
    PaZun = vbNullString
    PaNam = vbNullString
    PaGeb = vbNullString
    PaStr = vbNullString
    PaLKZ = vbNullString
    PaPLZ = vbNullString
    PaOrt = vbNullString
    KaGul = vbNullString
    KaDat = vbNullString

    With clChe
        .SmartReader = GlKvk(GlChp, 1)
        .TerminalPort = GlChR

        RetWe = .CarRead()
        StrFe = .Returnstring
        
        If RetWe = 0 Then
            PaTit = .Patient_Titel
            PaVor = .Patient_Vorname
            PaZun = .Patient_Zusatz
            PaNam = .Patient_Name
            PaGeb = .Patient_Geboren
            If .IsteGK = True Then
                PaStr = .Patient_Strasse & Chr$(32) & .Patient_Hausnummer
                PaGes = LCase(.Patient_Geschlecht)
            Else
                PaStr = .Patient_Strasse
                PaGes = "w"
            End If
            PaLKZ = .Patient_Landeskennzeichen
            PaPLZ = .Patient_PLZ
            PaOrt = .Patient_Ort
            KasNa = .Kostentraegername2
            KasNr = .Kostentraegerkennung
            KarNr = .Kartennummer
            VerNr = .Versicherten_ID
            KaSta = .Kartenstatus
            StaEr = .KVKStatus
            KaGul = .Kartengueltigkeit
            KaDat = .Kartendatum
            
            PatGe = Left$(PaGeb, 2) & "." & Mid$(PaGeb, 3, 2) & "." & Right$(PaGeb, 4)
            SuStr = PaNam & ", " & PaVor & "(" & PatGe & ")"
            
            GesZa = Adr_Fil(4, SuStr, 2)
            If GesZa > 0 Then
                Rahm1.Visible = False
                Rahm2.Visible = True
                LiVw3.SetFocus
                LiIts(1).Selected = True
                Me.btnZurück.Enabled = True
            Else
                If IsDate(PatGe) = True Then
                    GesZa = Adr_Fil(4, vbNullString, 5, PatGe)
                    If GesZa > 0 Then
                        Rahm1.Visible = False
                        Rahm2.Visible = True
                        LiVw3.SetFocus
                        LiIts(1).Selected = True
                        Me.btnZurück.Enabled = True
                    Else
                        TxMel.Text = TxMel.Text & vbCrLf & "Der Patient konnte nicht gefunden werden." & vbCrLf
                    End If
                End If
            End If
            
        End If
    End With
    TxMel.Text = TxMel.Text & vbCrLf & StrFe & vbCrLf
End If

Screen.MousePointer = vbNormal

Set clChe = Nothing

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FChip " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error Resume Next

Dim AktZa As Integer

Set FM = frmMain
Set ImMan = FM.imgManag
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set TxMel = Me.txtMeldu
Set CmChp = Me.cmbChipd
Set Lbl01 = Me.lblLab01
Set LiVw3 = Me.lstView3

Set PuBu1 = Me.btnPictu

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

With LiVw3
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = True
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = False
    .FlatScrollBar = False
    .Font.SIZE = 10 'GlTFt.size
    .Font.Name = GlTFt.Name
    .ForeColor = vbBlack
    .FullRowSelect = True
    .GridLines = False
    .HideColumnHeaders = False
    .HideSelection = False
    .HotTracking = False
    .HoverSelection = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpListViewLabelManual
    .LabelWrap = True
    .MultiSelect = False
    .OLEDropMode = xtpOLEDropNone
    .View = xtpListViewReport
End With

For AktZa = 0 To UBound(GlKvk) - 1
    CmChp.AddItem GlKvk(AktZa, 0)
Next AktZa
CmChp.ListIndex = GlChp

Select Case GlChp
Case 0: PuBu1.Icon = ImMan.Icons.GetImage(IC128_SmartCard, 128)
Case 1: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Zemo_GK2, 128)
Case 2: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1504, 128)
Case 3: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1504, 128)
Case 4: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1503, 128)
Case 5: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st2052, 128)
Case 6: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1530, 128)
Case 7: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Orga930M, 128)
Case 8: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Orga920MPlus, 128)
Case 9: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Orga6041, 128)
Case 10: PuBu1.Icon = ImMan.Icons.GetImage(IC128_SCM_2700R, 128)
Case 11: PuBu1.Icon = ImMan.Icons.GetImage(IC128_REINER_SCT, 128)
End Select

Me.BackColor = GlBak
PuBu1.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Lbl01.BackColor = GlBak

With LiVw3
    .ColumnHeaders.Add 1, , "Adresse", 3000
    .ColumnHeaders.Add 2, , "Mandant", 1900
    .ColumnHeaders.Add 3, , "Email", 0
    .ColumnHeaders.Add 4, , "Briefanrede", 0
End With

clFen.FenVor

Set clFen = Nothing

Set ImMan = Nothing

End Sub
Private Sub FSett()
On Error GoTo SeErr
'Lädt die ausgewählte Adresse in das Adressformular

Dim GesZa As Long
Dim IdxNr As Long
Dim BerNr As Long
Dim IdStr As String
Dim EmStr As String
Dim EmaNr As String
Dim PaStr As String
Dim TmStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set LiVw3 = Me.lstView3
Set LiIts = LiVw3.ListItems

GesZa = LiVw3.ListItems.Count

If GesZa > 0 Then
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
            IdStr = LiItm.Text
            Exit For
        End If
    Next LiItm
    
    Select Case GlBut
    Case RibTab_Startseite:
                Select Case GlAdO
                Case 0: SReZe GlAId
                Case 1: SKrZe GlAId
                Case 2: SKrZe GlAId
                End Select
    Case RibTab_Fragebogen:
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
    Case RibTab_Tagesproto:
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
    Case RibTab_Krankenbla:
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
    Case RibTab_Abrechnung:
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
    Case RibTab_Rezeptmodul:
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
    Case RibTab_Belegmodul:
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
    Case RibTab_Bildmodul:
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
    End Select
    DoEvents
    GlTDa = vbNullString 'Wichtig für Textverarbeitung
    
    Unload Me
End If

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub FZur()
On Error Resume Next

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2

If Rahm2.Visible = True Then
    Rahm2.Visible = False
    Rahm1.Visible = True
End If

End Sub
Private Sub btnHilfe_Click()
    Unload Me
    If WindowLoad("frmAdress") = True Then
        ASper True
        ANeue True
        Kon_Lis
    Else
        SAdre 1
        DoEvents
        AChip
    End If
End Sub
Private Sub btnPictu_Click()
    FChip
End Sub
Private Sub btnWeiter_Click()
On Error Resume Next

Set Rahm1 = Me.frmRahm1

If Rahm1.Visible = True Then
    FChip
Else
    FSett
End If

Me.btnZurück.Enabled = True
    
End Sub
Private Sub cmbChipd_Click()
On Error Resume Next

Set FM = frmMain
Set ImMan = FM.imgManag
Set PuBu1 = Me.btnPictu
Set CmChp = Me.cmbChipd

GlChp = CmChp.ListIndex

Select Case GlChp
Case 0: PuBu1.Icon = ImMan.Icons.GetImage(IC128_SmartCard, 128)
Case 1: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Zemo_GK2, 128)
Case 2: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1504, 128)
Case 3: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1504, 128)
Case 4: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1503, 128)
Case 5: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st2052, 128)
Case 6: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Cherry_st1530, 128)
Case 7: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Orga930M, 128)
Case 8: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Orga920MPlus, 128)
Case 9: PuBu1.Icon = ImMan.Icons.GetImage(IC128_Orga6041, 128)
Case 10: PuBu1.Icon = ImMan.Icons.GetImage(IC128_SCM_2700R, 128)
Case 11: PuBu1.Icon = ImMan.Icons.GetImage(IC128_REINER_SCT, 128)
End Select

End Sub
Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmChipcard = Nothing
End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub lstView3_DblClick()
    FSett
End Sub
Private Sub lstView3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub
Private Sub btnZurück_Click()
    FZur
    Me.btnZurück.Enabled = True
End Sub
