VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmAnaEdi 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Fragebogen"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   11
      Top             =   3700
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   3700
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schlieﬂen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Height          =   400
         Left            =   2300
         TabIndex        =   9
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
         Left            =   1000
         TabIndex        =   8
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
      Left            =   4000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   500
      _Version        =   1048579
      _ExtentX        =   882
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   5200
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3700
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   6526
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbBogen 
         Height          =   315
         Left            =   1005
         TabIndex        =   5
         Top             =   1635
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   2340
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   930
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu1"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   2610
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "÷ffnet den Auswahlkalender"
         Top             =   930
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   1000
         TabIndex        =   1
         Top             =   930
         Width           =   1320
         _Version        =   1048579
         _ExtentX        =   2328
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBezei 
         Height          =   350
         Left            =   1000
         TabIndex        =   6
         Top             =   2330
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   1000
         TabIndex        =   7
         Top             =   3030
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin VB.Label lblLab59 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   210
         Left            =   1000
         TabIndex        =   17
         Top             =   2800
         Width           =   1395
      End
      Begin VB.Label lblLab36 
         BackStyle       =   0  'Transparent
         Caption         =   "Anamnesebogen :"
         Height          =   210
         Left            =   1000
         TabIndex        =   16
         Top             =   1400
         Width           =   1395
      End
      Begin VB.Label lblLab30 
         BackStyle       =   0  'Transparent
         Caption         =   "Datum :"
         Height          =   210
         Left            =   1000
         TabIndex        =   15
         Top             =   700
         Width           =   1395
      End
      Begin VB.Label lblLab29 
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar :"
         Height          =   210
         Left            =   1000
         TabIndex        =   14
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte w‰hlen Sie den gew¸nschten Anamnesebogen und erg‰nzen alle weiteren Datenfelder. Klicken Sie auf Weiter."
         Height          =   440
         Left            =   600
         TabIndex        =   13
         Top             =   100
         Width           =   4400
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   3705
      Left            =   5880
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   6526
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView3 
         Height          =   2450
         Left            =   200
         TabIndex        =   19
         Top             =   930
         Width           =   5300
         _Version        =   1048579
         _ExtentX        =   9349
         _ExtentY        =   4322
         _StockProps     =   77
         BackColor       =   -2147483643
         View            =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Folgende Eintr‰ge wurden gefunden. Bitte w‰hlen Sie den gew¸nschten Patienten und klicken auf Weiter."
         Height          =   440
         Left            =   600
         TabIndex        =   33
         Top             =   100
         Width           =   4400
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte w‰hlen Sie einen der gefundenen Eintr‰ge :"
         Height          =   200
         Left            =   210
         TabIndex        =   20
         Top             =   700
         Width           =   3600
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3705
      Left            =   5880
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   6526
      _StockProps     =   79
      Caption         =   "GroupBox1"
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtBemer 
         Height          =   300
         Left            =   1000
         TabIndex        =   22
         Top             =   3030
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   300
         Left            =   1000
         TabIndex        =   23
         Top             =   930
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNumme 
         Height          =   300
         Left            =   3100
         TabIndex        =   24
         Top             =   1630
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGebor 
         Height          =   300
         Left            =   1000
         TabIndex        =   25
         Top             =   1630
         Width           =   1900
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPost 
         Height          =   300
         Left            =   1000
         TabIndex        =   26
         Top             =   2330
         Width           =   1900
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "F¸r welchen Patienten soll ein Fragebogen angelegt werden? Bitte geben Sie das Suchkriterium ein."
         Height          =   440
         Left            =   600
         TabIndex        =   32
         Top             =   100
         Width           =   4400
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Geburtsdatum"
         Height          =   200
         Left            =   1000
         TabIndex        =   31
         Top             =   1400
         Width           =   2000
      End
      Begin VB.Label Lab03 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Bemerkung"
         Height          =   200
         Left            =   1000
         TabIndex        =   30
         Top             =   2800
         Width           =   3000
      End
      Begin VB.Label Lab02 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Postleitzahl"
         Height          =   200
         Left            =   1000
         TabIndex        =   29
         Top             =   2100
         Width           =   3000
      End
      Begin VB.Label Lab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Patientenname"
         Height          =   200
         Left            =   1000
         TabIndex        =   28
         Top             =   700
         Width           =   3000
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Nummer"
         Height          =   200
         Left            =   3100
         TabIndex        =   27
         Top             =   1400
         Width           =   1600
      End
   End
End
Attribute VB_Name = "frmAnaEdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private FTex3 As XtremeSuiteControls.FlatEdit
Private FTex4 As XtremeSuiteControls.FlatEdit
Private FTex5 As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private CmBog As XtremeSuiteControls.ComboBox
Private CmThe As XtremeSuiteControls.ComboBox
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxBez As XtremeSuiteControls.FlatEdit
Private LiVw3 As XtremeSuiteControls.ListView
Private LiItm As XtremeSuiteControls.ListViewItem
Private LiIts As XtremeSuiteControls.ListViewItems
Private PuBu1 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private MoKal As XtremeCalendarControl.DatePicker

Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const KEYEVENTF_KEYUP = &H2
Private Const CB_FINDSTRING = &H14C&
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private PatNr As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    TxDa1.Text = NeuDa
    With MoKal
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
    If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    TxDa1.Text = NeuDa
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'L‰ﬂt den Kalender aufklappen

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = TxDa1.Top + TxDa1.Height
    .Left = TxDa1.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TxDa1.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FSett()
On Error GoTo SeErr
'L‰dt die ausgew‰hlte Adresse in das Adressformular

Dim GesZa As Long

Set FTex1 = Me.txtKurz
Set LiVw3 = Me.lstView3
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set LiIts = LiVw3.ListItems

GesZa = LiVw3.ListItems.Count

If GesZa > 0 Then
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            PatNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
            Exit For
        End If
    Next LiItm
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
End If

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo SuErr

Dim NeuDa As Date
Dim ManNr As Long
Dim BogNr As Long
Dim TmStr As String

Set FM = frmAnaEdi
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set CmBog = FM.cmbBogen
Set CmThe = FM.cmbManda
Set PuBu1 = FM.btnDatu1
Set TxDa1 = FM.txtDatu1
Set TxBez = FM.txtBezei

If Rahm1.Visible = True Then
    If PatNr > 0 Then
        GlAdr = PatNr
        With GlSuV
            .SuIdx = 1
            .SuNum = PatNr
        End With
        With GlSuA
            .SuIdx = 1
            .SuNum = PatNr
        End With
        With GlSuP
            .SuIdx = 1
            .SuNum = PatNr
        End With
        S_KrLa
        DoEvents
        If GlBut = RibTab_Startseite Then
            GlBu2 = RibTab_Fragebogen
            STaSe ShoCut_Kranken, RibTab_Fragebogen
        End If
    End If

    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
    Else
        NeuDa = Date
    End If

    ManNr = CmThe.ItemData(CmThe.ListIndex)
    
    If GlBoV > 0 Then 'Fragebogen vorhanden
        BogNr = CmBog.ItemData(CmBog.ListIndex)
    Else
        BogNr = GlFrB(1, 0)
    End If
    
    If TxBez.Text <> vbNullString Then
        TmStr = TxBez.Text
    Else
        TmStr = CmBog.Text
    End If
    
    With GlBoX 'neuer Fragebogen
        .PatNr = GlAdr
        .BoDat = NeuDa
        .BoMan = ManNr
        .BoNum = BogNr
        .KoTex = TmStr
    End With
    
    If GlAnN = True Then 'Neuer Fragebogen
        S_AnBoN
    Else
        S_Save
    End If
    DoEvents
    
    SUpAn
    DoEvents

    Unload Me
ElseIf Rahm2.Visible = True Then
    FSuda
ElseIf Rahm3.Visible = True Then
    FSett
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub TRes()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtPost
Set FTex3 = Me.txtNumme
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor

FTex1.Text = vbNullString
FTex2.Text = vbNullString
FTex3.Text = vbNullString
FTex4.Text = vbNullString
FTex5.Text = vbNullString

End Sub
Private Sub FSuda()
On Error GoTo SeErr

Dim GesZa As Long
Dim Mld1, Tit1 As String

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumme
Set FTex3 = Me.txtPost
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set LiVw3 = Me.lstView3
Set LiIts = LiVw3.ListItems

If FTex1.Text <> vbNullString Then
    GesZa = Adr_Fil(3, FTex1.Text, 1)
    If GesZa = 0 Then
        GesZa = Adr_Fil(3, SUmw(FTex1.Text), 1)
    End If
ElseIf FTex2.Text <> vbNullString Then
    GesZa = Adr_Fil(3, FTex2.Text, 2)
ElseIf FTex3.Text <> vbNullString Then
    GesZa = Adr_Fil(3, FTex3.Text, 3)
ElseIf FTex4.Text <> vbNullString Then
    GesZa = Adr_Fil(3, FTex4.Text, 4)
ElseIf FTex5.Text <> vbNullString Then
    GesZa = Adr_Fil(3, vbNullString, 5, FTex5.Text)
End If

If GesZa > 0 Then
    Rahm2.Visible = False
    Rahm3.Visible = True
    LiVw3.SetFocus
    LiIts(1).Selected = True
Else
    If FTex1.Text <> vbNullString Then
        FTex1.SelStart = 0
        FTex1.SelLength = Len(FTex1.Text)
    ElseIf FTex2.Text <> vbNullString Then
        FTex2.SelStart = 0
        FTex2.SelLength = Len(FTex2.Text)
    ElseIf FTex3.Text <> vbNullString Then
        FTex3.SelStart = 0
        FTex3.SelLength = Len(FTex3.Text)
    ElseIf FTex4.Text <> vbNullString Then
        FTex4.SelStart = 0
        FTex4.SelLength = Len(FTex4.Text)
    ElseIf FTex5.Text <> vbNullString Then
        FTex5.SelStart = 0
        FTex5.SelLength = Len(FTex5.Text)
    End If
    SPopu "Patient nicht gefunden", "Der von Ihnen gesuchte Patient, konnte nicht gefunden werden", IC48_Forbidden
End If

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuda " & Err.Number
Resume Next

End Sub
Private Sub FZur()
On Error Resume Next

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3

If Rahm3.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = True
    Rahm3.Visible = False
ElseIf Rahm2.Visible = True Then
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
End If

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim IdxNr As Long
Dim MaNum As Long
Dim AktZa As Integer
Dim GesZa As Integer
Dim LauZa As Integer
Dim TeWer As Variant
Dim BeVor As Boolean

Set FM = frmAnaEdi
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumme
Set FTex3 = Me.txtPost
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor
Set LiVw3 = Me.lstView3
Set CmBog = FM.cmbBogen
Set CmThe = FM.cmbManda
Set PuBu1 = FM.btnDatu1
Set TxDa1 = FM.txtDatu1
Set TxBez = FM.txtBezei
Set MoKal = FM.dtpDatu1
Set ImMan = frmMain.imgManag

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

FTex2.Pattern = "\d*"
FTex3.Pattern = "\d*"
FTex5.SetMask "00.00.0000", "__.__.____"

With MoKal
    .AllowNoncontinuousSelection = False
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    If GlSty = 8 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    ElseIf GlSty = 7 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    Else
        .BorderStyle = xtpDatePickerBorderOffice
    End If
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    .MaxSelectionCount = 1
    .RightToLeft = False
    .ShowNoneButton = True
    .ShowNonMonthDays = True
    .ShowTodayButton = True
    .ShowWeekNumbers = False
    .TextNoneButton = "Keine"
    .TextTodayButton = "Heute"
    .ToolTipText = "Markieren Sie bitte hier das gw¸nschte Buchungsdatum"
    .MonthDelta = 1
    .YearsTriangle = False
    Select Case GlSty
    Case 8: .VisualTheme = xtpCalendarThemeResource
    Case 7: .VisualTheme = xtpCalendarThemeResource
    Case Else: .VisualTheme = xtpCalendarThemeResource
    End Select
    .PaintManager.ButtonTextColor = -2147483640
    .PaintManager.ControlBackColor = -2147483643
    .PaintManager.DayBackColor = -2147483643
    .PaintManager.DayTextColor = -2147483640
    .PaintManager.DaysOfWeekBackColor = -2147483643
    .PaintManager.DaysOfWeekTextColor = -2147483640
    .PaintManager.ListControlBackColor = -2147483643
    .PaintManager.ListControlTextColor = -2147483640
    .PaintManager.NonMonthDayBackColor = -2147483643
    .PaintManager.NonMonthDayTextColor = -2147483640
    .PaintManager.SelectedDayBackColor = GlFac
    .PaintManager.SelectedDayTextColor = -2147483640
    .PaintManager.WeekNumbersBackColor = -2147483643
    .PaintManager.WeekNumbersTextColor = -2147483640
    .PaintManager.MonthHeaderBackColor = GlMoB
End With

If GlBoV > 0 Then 'Fragebogen vorhanden
    For AktZa = 1 To GlBoV
        CmBog.AddItem GlFrB(AktZa, 1)
        CmBog.ItemData(AktZa - 1) = GlFrB(AktZa, 0)
    Next AktZa
End If

For AktZa = 1 To UBound(GlThe)
    If GlAnN = True Then 'Neuer Fragebogen
        If CBool(GlThe(AktZa, 25)) = False Then
            LauZa = LauZa + 1
            CmThe.AddItem GlThe(AktZa, 13)
            CmThe.ItemData(LauZa - 1) = GlThe(AktZa, 0)
        End If
    Else
        LauZa = LauZa + 1
        CmThe.AddItem GlThe(AktZa, 13)
        CmThe.ItemData(LauZa - 1) = GlThe(AktZa, 0)
    End If
Next AktZa

TeWer = S_AdIdi(GlAdr, "IDP")
If TeWer > 0 Then
    MaNum = CLng(TeWer)
    For AktZa = 1 To UBound(GlMan)
        If MaNum = GlMan(AktZa, 2) Then
            BeVor = True
            Exit For
        End If
    Next AktZa
    If BeVor = True Then
        If CBool(GlMan(AktZa, 5)) = False Then 'Passiv / Aktiv
            MaNum = GlThe(AktZa, 0)
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
    Else
        MaNum = GlMan(GlSMa, 2)
    End If
Else
    MaNum = GlMan(GlSMa, 2)
End If

IdxNr = SCmb(CmThe, MaNum)
CmThe.ListIndex = IdxNr

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

If CmThe.Enabled = False Then
    CmThe.Enabled = True
End If

If GlAnN = True Then 'Neuer Fragebogen
    GlBoX = GlBoY
    CmBog.ListIndex = 0
Else
    S_Posi
End If

With LiVw3
    .ColumnHeaders.Add 1, , "Adresse", 3000
    .ColumnHeaders.Add 2, , "Mandant", 1900
End With

Rahm1.Move 0, 0
Rahm2.Move 0, 0
Rahm3.Move 0, 0

Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
FM.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    FKale
End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50041)
TeMai = IniGetOpt("Hilfe", 50042)
TeInh = IniGetOpt("Hilfe", 50043)
TeFus = IniGetOpt("Hilfe", 50044)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub

Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub

Private Sub Form_Activate()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3

If GlBut = RibTab_Startseite Then
    Rahm1.Visible = False
    Rahm2.Visible = True
    Rahm3.Visible = False
    FTex1.SetFocus
Else
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
End If

End Sub

Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAnaEdi = Nothing
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub

Private Sub updCont1_DownClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub
Private Sub lstView3_DblClick()
    FSett
End Sub
Private Sub lstView3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub
Private Sub txtBemer_GotFocus()
    TRes
End Sub
Private Sub txtBemer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub

Private Sub txtGebor_GotFocus()
    TRes
End Sub
Private Sub txtGebor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub

Private Sub txtKurz_GotFocus()
    TRes
End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtNumme_GotFocus()
    TRes
End Sub
Private Sub txtNumme_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then FSuda
End Sub

Private Sub txtNumme_Validate(Cancel As Boolean)
    If (Not txtNumme.isValid) Then Cancel = True
End Sub

Private Sub txtPost_GotFocus()
    TRes
End Sub
Private Sub txtPost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then FSuda
End Sub

Private Sub txtPost_Validate(Cancel As Boolean)
    If (Not txtPost.isValid) Then Cancel = True
End Sub


