VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmKontakt 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Notiz"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   ControlBox      =   0   'False
   Icon            =   "frmKontakt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
      Height          =   400
      Left            =   4600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4000
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
      Left            =   3200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4000
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "Sp&eichern"
      UseVisualStyle  =   -1  'True
      PushButtonStyle =   2
   End
   Begin XtremeSuiteControls.PushButton btnZur¸ck 
      Height          =   400
      Left            =   1800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4000
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Lˆschen"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtKomme 
      Height          =   1780
      Left            =   300
      TabIndex        =   11
      Tag             =   "0Kommentar"
      Top             =   1920
      Width           =   6100
      _Version        =   1048579
      _ExtentX        =   10760
      _ExtentY        =   3140
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   1800
      Left            =   300
      TabIndex        =   1
      Top             =   40
      Width           =   6080
      _Version        =   1048579
      _ExtentX        =   10724
      _ExtentY        =   3175
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.UpDown updCont3 
         Height          =   350
         Left            =   4020
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1280
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBisZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   350
         Left            =   4020
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   830
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtVonZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   2260
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   830
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatum"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtBehan 
         Height          =   350
         Left            =   1000
         TabIndex        =   5
         Tag             =   "0Behandler"
         Top             =   1280
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.CheckBox chkErled 
         Height          =   240
         Left            =   4600
         TabIndex        =   10
         Tag             =   "0Erledigt"
         Top             =   1300
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Erledigt"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAnlaﬂ 
         Height          =   350
         Left            =   1000
         TabIndex        =   2
         Tag             =   "0IDKurz"
         Top             =   300
         Width           =   4800
         _Version        =   1048579
         _ExtentX        =   8467
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtDatum 
         Height          =   350
         Left            =   1000
         TabIndex        =   3
         Tag             =   "0VonDat"
         Top             =   830
         Width           =   1245
         _Version        =   1048579
         _ExtentX        =   2196
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtVonZe 
         Height          =   350
         Left            =   3200
         TabIndex        =   6
         Tag             =   "0ZeiVon"
         Top             =   830
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBisZe 
         Height          =   350
         Left            =   3200
         TabIndex        =   8
         Tag             =   "0ZeiBis"
         Top             =   1280
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin VB.Label Lab04 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2760
         TabIndex        =   19
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Lab03 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von :"
         Height          =   240
         Left            =   2760
         TabIndex        =   18
         Top             =   880
         Width           =   405
      End
      Begin VB.Label Lab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Datum :"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   880
         Width           =   840
      End
      Begin VB.Label Lab01 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anlass :"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   340
         Width           =   840
      End
      Begin VB.Label Lab08 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bearbeitet :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   840
      End
   End
   Begin VB.TextBox txtID2 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Tag             =   "0ID2"
      Top             =   5000
      Width           =   80
   End
End
Attribute VB_Name = "frmKontakt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxKom As XtremeSuiteControls.FlatEdit
Private VoZei As XtremeSuiteControls.FlatEdit
Private BiZei As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private Rahm1 As XtremeSuiteControls.GroupBox
Private ChErl As XtremeSuiteControls.CheckBox
Private CoDia As XtremeSuiteControls.CommonDialog

Private Const CB_SHOWDROPDOWN = &H14F
Private Const KEYEVENTF_KEYUP = &H2

Private TagWe As String

Private clFen As clsFenster
Private clFil As clsFile
Private clWor As clsWord

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatum

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    TxDa1.Text = NeuDa
    If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
End If

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub btnSchlieﬂ_Click()
    FClos
    Unload Me
End Sub

Private Sub btnWeiter_Click()
    FKSav
    Unload Me
End Sub

Private Sub btnZur¸ck_Click()
    Kon_Loe 1
    Kon_Lis
    Unload Me
End Sub

Private Sub chkErled_Click()

TagWe = Mid$(Me.chkErled.Tag, 2, Len(Me.chkErled.Tag) - 1)

If GlKoL = False Then
    Me.chkErled.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtBisZe_Click()
On Error Resume Next

TagWe = Mid$(Me.txtBisZe.Tag, 2, Len(Me.txtBisZe.Tag) - 1)

If GlKoL = False Then
    Me.txtBisZe.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtBisZe_GotFocus()
    Me.txtBisZe.SelStart = 0
    Me.txtBisZe.SelLength = Len(Me.txtBisZe.Text)
    GlKoS = True
    GlKoL = False
End Sub

Private Sub txtBisZe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBisZe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtBisZe.SelLength = 0
End Sub

Private Sub txtBisZe_LostFocus()
On Error Resume Next

Dim TmVon As Date
Dim TmBis As Date

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
Else
    VoZei.Text = Format$(TimeValue(Now), "hh:mm")
    TmVon = Now
End If

If BiZei.Text <> vbNullString Then
    TmBis = TimeValue(BiZei.Text)
Else
    BiZei.Text = Format$(TimeValue(Now), "hh:mm")
    TmVon = Now
End If

If TmVon > TmBis Then
    BiZei.Text = VoZei.Text
End If

TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)

If GlKoL = False Then
    BiZei.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub txtDatum_Change()

TagWe = Mid$(Me.txtDatum.Tag, 2, Len(Me.txtDatum.Tag) - 1)

If GlKoL = False Then
    Me.txtDatum.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtDatum_GotFocus()
    Me.txtDatum.SelStart = 0
    Me.txtDatum.SelLength = Len(Me.txtDatum.Text)
    GlKoL = False
End Sub
Private Sub txtDatum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtDatum_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtDatum.SelLength = 0
End Sub
Private Sub txtDatum_LostFocus()
    FDaKo
End Sub
Private Sub txtVonZe_Click()
On Error Resume Next

TagWe = Mid$(Me.txtVonZe.Tag, 2, Len(Me.txtVonZe.Tag) - 1)

If GlKoL = False Then
    Me.txtVonZe.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtVonZe_GotFocus()
    Me.txtVonZe.SelStart = 0
    Me.txtVonZe.SelLength = Len(Me.txtVonZe.Text)
    GlKoS = True
    GlKoL = False
End Sub

Private Sub txtVonZe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtVonZe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtVonZe.SelLength = 0
End Sub

Private Sub FClos()
On Error GoTo SaErr

Dim Mld1, Tit1 As String
Dim Frage As Integer
Dim FoID2 As VB.TextBox

Set FM = frmAdress
Set FoID2 = Me.txtID2

Tit1 = "Datensatz Speichern"
Mld1 = "Soll der Datensatz gespeichert werden?"

If GlKoS = True Then
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
        If GlKoN = True Then
            Kon_San
        Else
            GlKoG = FoID2.Text
            Kon_Sav
        End If
        Kon_Lis
        GlKoS = False
    End If
End If

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error Resume Next

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set TxDa1 = Me.txtDatum
Set TxKom = Me.txtKomme

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

With VoZei
    .SetMask "00:00", "__:__"
    .Text = Format$(TimeValue(Now), "hh:mm")
End With

With BiZei
    .SetMask "00:00", "__:__"
    .Text = Format$(DateAdd("n", 10, TimeValue(Now)), "hh:mm")
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

TxKom.Font.Name = GlTFt.Name
TxKom.Font.SIZE = GlTFt.SIZE

Me.BackColor = GlBak
Me.frmRahm1.BackColor = GlBak
Me.chkErled.BackColor = GlBak

clFen.FenVor

Set clFen = Nothing

End Sub
Private Sub FKSav()
On Error GoTo SaErr

Dim FoID2 As VB.TextBox

Set FoID2 = Me.txtID2

If GlKoS = True Then
    If GlKoN = True Then
        Kon_San
    Else
        GlKoG = FoID2.Text
        Kon_Sav
    End If
    Kon_Lis
    GlKoS = False
End If

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "Kon_Sav " & Err.Number
Resume Next

End Sub

Private Sub Form_Activate()
    GlKoL = True
    Me.txtAnlaﬂ.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next

FKonf
AFont Me
SFrame 1, Me.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmKontakt = Nothing
End Sub

Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub txtAnlaﬂ_Change()

TagWe = Mid$(Me.txtAnlaﬂ.Tag, 2, Len(Me.txtAnlaﬂ.Tag) - 1)

If GlKoL = False Then
    Me.txtAnlaﬂ.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub
Private Sub txtAnlaﬂ_GotFocus()
    Me.txtAnlaﬂ.SelStart = 0
    Me.txtAnlaﬂ.SelLength = Len(Me.txtAnlaﬂ.Text)
    GlKoL = False
End Sub

Private Sub txtAnlaﬂ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
    GlKoL = False
End Sub

Private Sub txtAnlaﬂ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtAnlaﬂ.SelLength = 0
End Sub

Private Sub txtBehan_Change()

TagWe = Mid$(Me.txtBehan.Tag, 2, Len(Me.txtBehan.Tag) - 1)

If GlKoL = False Then
    Me.txtBehan.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub txtBehan_GotFocus()
    Me.txtBehan.SelStart = 0
    Me.txtBehan.SelLength = Len(Me.txtBehan.Text)
    GlKoL = False
End Sub

Private Sub txtBehan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
    GlKoL = False
End Sub

Private Sub txtBehan_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtBehan.SelLength = 0
End Sub

Private Sub txtKomme_Change()

TagWe = Mid$(Me.txtKomme.Tag, 2, Len(Me.txtKomme.Tag) - 1)

If GlKoL = False Then
    Me.txtKomme.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub txtKomme_GotFocus()
    Me.txtKomme.SelStart = 0
    Me.txtKomme.SelLength = Len(Me.txtKomme.Text)
    GlKoL = False
End Sub

Private Sub txtKomme_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtKomme.SelLength = 0
End Sub

Private Sub txtVonZe_LostFocus()
On Error Resume Next

Dim TmVon As Date
Dim TmBis As Date

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
Else
    TmVon = Now
    VoZei.Text = Format$(TimeValue(Now), "hh:mm")
End If

If BiZei.Text <> vbNullString Then
    TmBis = TimeValue(BiZei.Text)
    BiZei.Text = Format$(TimeValue(Now), "hh:mm")
Else
    TmVon = Now
End If

If TmVon > TmBis Then
    BiZei.Text = VoZei.Text
End If

TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)

If GlKoL = False Then
    VoZei.Tag = 1 & TagWe
    GlKoS = True
End If

End Sub

Private Sub updCont1_DownClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatum

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatum

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub

Private Sub updCont2_DownClick()

Dim AlDa1 As Date
Dim AlDa2 As Date

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

AlDa1 = TimeValue(VoZei.Text)
AlDa2 = TimeValue(BiZei.Text)
TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)

VoZei.Text = Format$(DateAdd("n", -1, AlDa1), "hh:mm")
VoZei.Tag = 1 & TagWe
GlKoS = True

End Sub
Private Sub updCont2_UpClick()

Dim AlDa1 As Date
Dim AlDa2 As Date

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

AlDa1 = TimeValue(VoZei.Text)
AlDa2 = TimeValue(BiZei.Text)
TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)

VoZei.Text = Format$(DateAdd("n", 1, AlDa1), "hh:mm")
If AlDa1 >= AlDa2 Then
    BiZei.Text = Format$(DateAdd("n", 1, AlDa2), "hh:mm")
End If
VoZei.Tag = 1 & TagWe
GlKoS = True

End Sub
Private Sub updCont3_DownClick()

Dim AlDa1 As Date
Dim AlDa2 As Date

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

AlDa1 = TimeValue(VoZei.Text)
AlDa2 = TimeValue(BiZei.Text)
TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)

BiZei.Text = Format$(DateAdd("n", -1, AlDa2), "hh:mm")
If AlDa1 >= AlDa2 Then
    VoZei.Text = Format$(DateAdd("n", -1, AlDa1), "hh:mm")
End If
BiZei.Tag = 1 & TagWe
GlKoS = True

End Sub
Private Sub updCont3_UpClick()

Dim AlDa1 As Date
Dim AlDa2 As Date

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

AlDa1 = TimeValue(VoZei.Text)
AlDa2 = TimeValue(BiZei.Text)
TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)

BiZei.Text = Format$(DateAdd("n", 1, AlDa2), "hh:mm")
BiZei.Tag = 1 & TagWe
GlKoS = True

End Sub
