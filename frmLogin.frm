VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.ShortcutBar.v16.3.1.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Benutzeranmeldung"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   5
      Top             =   3100
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Cancel          =   -1  'True
         Height          =   400
         Left            =   4000
         TabIndex        =   4
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
         Left            =   2600
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&OK"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.CheckBox chkStaMi 
      Height          =   255
      Left            =   2120
      TabIndex        =   2
      Top             =   2640
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Standardmitarbeiter festlegen"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPassw 
      Height          =   350
      Left            =   2120
      TabIndex        =   1
      Top             =   2100
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      PasswordChar    =   "*"
   End
   Begin XtremeSuiteControls.ComboBox cmbMitar 
      Height          =   315
      Left            =   2120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1500
      Width           =   2800
      _Version        =   1048579
      _ExtentX        =   4948
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin VB.Label lblLab03 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Mitarbeiter :"
      Height          =   220
      Left            =   1060
      TabIndex        =   9
      Top             =   1540
      Width           =   1000
   End
   Begin XtremeSuiteControls.Label lblLab02 
      Height          =   240
      Left            =   100
      TabIndex        =   8
      Top             =   900
      Width           =   5800
      _Version        =   1048579
      _ExtentX        =   10231
      _ExtentY        =   423
      _StockProps     =   79
      ForeColor       =   192
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label lblLab04 
      Height          =   220
      Left            =   1060
      TabIndex        =   7
      Top             =   2140
      Width           =   1000
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Passwort :"
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption schCapt1 
      Height          =   800
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   1411
      _StockProps     =   14
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   4
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Lbl02 As XtremeSuiteControls.Label
Private CmMit As XtremeSuiteControls.ComboBox
Private TxPas As XtremeSuiteControls.FlatEdit
Private ChMit As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private ShCap As XtremeShortcutBar.ShortcutCaption
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmSta As XtremeCommandBars.StatusBar

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub FKont(Optional ByVal PrSta As Boolean = False)
On Error GoTo MnErr

Dim ManNr As Long
Dim ErKId As Long
Dim PasWo As String
Dim Recht As String
Dim MiNam As String
Dim MaNam As String
Dim AktZa As Integer
Dim AkKat As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmPa3 As XtremeCommandBars.StatusBarPane

Set FM = frmMain
Set Lbl02 = Me.lblLab02
Set CmMit = Me.cmbMitar
Set ChMit = Me.chkStaMi
Set TxPas = Me.txtPassw
Set PuBu1 = Me.btnSchließ
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmPa3 = CmSta.FindPane(Tex_Pa_Labl3)

GlSmI = CmMit.ListIndex + 1

If GlSMo > UBound(GlMiT) Then 'Standardmitarbeiter Online-Terminbuchungs Sytem
    GlSMo = 1
End If

MiNam = GlMiA(GlSmI, 1)
ManNr = GlMiA(GlSmI, 7)
PasWo = GlMiA(GlSmI, 18)
Recht = GlMiA(GlSmI, 19)

If Recht = vbNullString Then
    Recht = GlStR 'Rechtestring
End If

If IsNumeric(Recht) = False Then
    Recht = GlStR 'Rechtestring
End If

If Len(Recht) <> GlZaR Then 'Rechteanzahl
    Recht = GlStR 'Rechtestring
End If

GlReM = ManNr 'Standardmandant Mandantenrestriktionen

For AktZa = 1 To UBound(GlMaA)
    If ManNr = GlMaA(AktZa, 2) Then
        GlFri = GlMaA(AktZa, 35) 'Fachrichtung
        If GlMaR = "J3" Then 'Mandant aus Mitarbeitereingabemaske
            GlSMa = AktZa 'neuer Standardmandant!
            Exit For
        ElseIf GlRst = True Then 'Mandantenbezogene Datenbegrenzung
            GlSMa = AktZa 'neuer Standardmandant!
            Exit For
        End If
    End If
Next AktZa

If GlFri > 9 Then
    GlFri = 2
End If

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If GlOTS = True Then 'Online-Terminbuchungs Sytem
        If GlSMo <= UBound(GlMiA) Then 'Standardmitarbeiter Online-Terminbuchungs Sytem
            GlTBx = GlMiA(GlSMo, 0) - 1 'Termin Behandlerindex
            GlTBn = GlMiA(GlSMo, 2) 'Termin Behandlernummer
        Else
            GlTBx = GlMiA(GlSmI, 0) - 1
            GlTBn = GlMiA(GlSmI, 2)
        End If
    Else
        GlTBx = GlMiA(GlSmI, 0) - 1
        GlTBn = GlMiA(GlSmI, 2)
    End If
Else
    GlTBx = GlMaA(GlSMa, 0) - 1
    GlTBn = GlMaA(GlSMa, 2)
End If

If PasWo = vbNullString Then
    If PuBu1.Enabled = True Then
        PuBu1.Enabled = False
    End If
    Lbl02.Caption = "Für diesen Benutzer wurde noch kein Passwort festgelegt. Bitte OK klicken!"
Else
    If GlRst = True Then 'Mandantenbezogene Datenbegrenzung
        For AktZa = 1 To UBound(GlMaA)
            If ManNr = GlMaA(AktZa, 2) Then 'Mandantennummer
                MaNam = GlMaA(AktZa, 1)
                Lbl02.Caption = "Mandant: " & MaNam
            End If
        Next AktZa
    Else
        Lbl02.Caption = vbNullString
    End If
    If PuBu1.Enabled = False Then
        PuBu1.Enabled = True
    End If
End If

TxPas.Text = vbNullString

For AktZa = 0 To GlZaR - 1 'Rechteanzahl
    If Mid$(Recht, AktZa + 1, 1) = "1" Then
        GlRch(0, AktZa) = 1
    Else
        GlRch(0, AktZa) = 0
    End If
Next AktZa

If GlRst = True Then 'Mandantenbezogene Datenbegrenzung
    CmPa3.Text = "Mitarbeiter: " & MiNam & " (Mandantenrestriktionen)"
Else
    CmPa3.Text = "Mitarbeiter: " & MiNam
End If

DoEvents
If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMaA)
        If ManNr = GlMaA(AktZa, 2) Then 'Mandantennummer
            If GlMaA(AktZa, 22) <> vbNullString Then 'Standardgebührenkatalog
                For AkKat = 1 To UBound(GlGKa)
                    If GlGKa(AkKat, 0) = GlMaA(AktZa, 22) Then
                        GlStK = AkKat
                        Exit For
                    End If
                Next AkKat
            End If
            If GlMaA(AktZa, 23) <> vbNullString Then GlKe1 = GlMaA(AktZa, 23) 'Standardgebührenkette 1
            If GlMaA(AktZa, 38) <> vbNullString Then GlKe2 = GlMaA(AktZa, 38) 'Standardgebührenkette 2
            If GlMaA(AktZa, 24) <> vbNullString Then GlStS = GlMaA(AktZa, 24) 'Standardsteuersatz
            If GlMaA(AktZa, 25) <> vbNullString Then GlKtR = GlMaA(AktZa, 25) 'Standardkontenrahmen
            If GlMaA(AktZa, 26) <> vbNullString Then ErKId = GlMaA(AktZa, 26) 'Standarderlöskonto Bank
            S_KoDe ErKId
            GlSE2 = GlKto.KtoNr 'Standarderlöskonto Bankkonto
            If GlMaA(AktZa, 27) <> vbNullString Then ErKId = GlMaA(AktZa, 27) 'Standarderlöskonto Kasse
            S_KoDe ErKId
            GlSE1 = GlKto.KtoNr
            If GlMaA(AktZa, 28) <> vbNullString Then GlGkB = GlMaA(AktZa, 28) 'Standardgeldkonto Bankkonto
            If GlMaA(AktZa, 29) <> vbNullString Then GlGkK = GlMaA(AktZa, 29) 'Standardgeldkonto Kasse
            If GlMaA(AktZa, 30) <> vbNullString Then GlReT = GlMaA(AktZa, 30) 'Standardbelegtyp
            If GlMaA(AktZa, 35) <> vbNullString Then
                If GlMaA(AktZa, 35) > UBound(GlFch) Then
                    GlFri = 2 'Heilpraktiker (GebüH)
                Else
                    GlFri = GlMaA(AktZa, 35) 'Fachrichtung
                End If
            End If
            If GlMaR = "J4" Then 'Mandant neue(s) Rechnung/Rezept
                GlSMa = AktZa 'neuer Standardmandant!
            End If
            If GlMPl = False Then 'Mitarbeiterplan anstelle von Mandantenplan
                GlTBx = GlMaA(GlSMa, 0) - 1
                GlTBn = GlMaA(GlSMa, 2)
            End If
            Exit For
        End If
    Next AktZa
End If

If PrSta = False Then
    If ChMit.Enabled = False Then
        ChMit.Enabled = True
    End If
End If

DoEvents
SRecht

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKont " & Err.Number
Resume Next

End Sub
Private Sub FLoad()
On Error GoTo MnErr

Dim RetWe As Long
Dim PasWo As String
Dim AktZa As Integer
Dim ImMan As XtremeCommandBars.ImageManager
Dim ImIcs As XtremeCommandBars.ImageManagerIcons

Set Rahm0 = Me.frmRahm0
Set CmMit = Me.cmbMitar
Set ChMit = Me.chkStaMi
Set Lbl02 = Me.lblLab02
Set TxPas = Me.txtPassw
Set ShCap = Me.schCapt1
Set PuBu1 = Me.btnSchließ
Set ImMan = frmMain.imgManag
Set ImIcs = ImMan.Icons

With ShCap
    Select Case GlSty
    Case 8: .VisualTheme = xtpShortcutThemeOffice2013
    Case 7: .VisualTheme = xtpShortcutThemeOffice2013
    Case Else: .VisualTheme = xtpShortcutThemeResource
    End Select
    .Font.Name = GlTFt.Name
    .Font.SIZE = 14
    .Font.Bold = True
    .Alignment = xtpAlignmentLeft
    .Caption = "Benutzeranmeldung"
    .SubItemCaption = False
    .Icon = ImMan.Icons.GetImage(IC48_Lock, 48)
    If GlSty = 8 Then 'Office 2013
        .GradientColorDark = GlMoB
        .GradientColorLight = GlMoB
        .ForeColor = -2147483641
    ElseIf GlSty = 7 Then 'Office 2013
        .GradientColorDark = GlMoB
        .GradientColorLight = GlMoB
        .ForeColor = -2147483641
    End If
End With

If CBool(IniGetVal("Vorgabe", "StMaVo")) = True Then  'Standardmandant vorhanden
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        CmMit.AddItem GlMiA(AktZa, 1)
        CmMit.ItemData(AktZa - 1) = GlMiA(AktZa, 2)
    Next AktZa
    If CmMit.ListCount > 1 Then
        RetWe = SendMessage(CmMit.hwnd, CB_SETCURSEL, GlSmI - 1, ByVal 0&)
    Else
        RetWe = SendMessage(CmMit.hwnd, CB_SETCURSEL, 0, ByVal 0&)
    End If
    PasWo = GlMiA(GlSmI, 18)
End If

If PasWo = vbNullString Then
    If PuBu1.Enabled = True Then
        PuBu1.Enabled = False
    End If
End If

ChMit.BackColor = GlBak
Rahm0.BackColor = GlBak

DoEvents
FKont True

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FPass()
On Error GoTo MnErr

Dim PasWo As String
Dim DaPas As String
Dim MiIdx As Integer
Dim Kurat As Integer
Dim Mld1, Tit1 As String

Set CmMit = Me.cmbMitar
Set ChMit = Me.chkStaMi
Set TxPas = Me.txtPassw
Set ShCap = Me.schCapt1
Set Lbl02 = Me.lblLab02

MiIdx = CmMit.ListIndex + 1

Tit1 = "Falsches Passwort"
Mld1 = "Das von Ihnen eingegebene Passwort ist nicht richtig!"

DaPas = IniGetVal("System", "DatPas") 'Datenbankpasswort Access
DaPas = SCrypt(DaPas, False)

If GlMiV = True Then 'Mitarbeiter vorhanden
    PasWo = LCase(GlMiA(MiIdx, 18)) 'aus Mitarbeiter-Arrey
    Kurat = CInt(GlMiA(MiIdx, 38))
End If

If ChMit.Enabled = True Then
    If ChMit.Value = xtpChecked Then
        IniSetVal "Vorgabe", "StaMit", MiIdx
    End If
End If

If Len(PasWo) = 0 Then
    Unload Me
    DoEvents
ElseIf LCase(TxPas.Text) = PasWo Then
    Unload Me
    DoEvents
ElseIf TxPas.Text = DaPas Then
    If GlNoM = False Then 'kein Meisterp...
        If Kurat <> 100 Then
            GlMVo = False 'mandantenbezogene Vorgaben verwenden
            Unload Me
            DoEvents
        Else
            TxPas.Text = vbNullString
            Lbl02.Caption = Mld1
        End If
    Else
        TxPas.Text = vbNullString
        Lbl02.Caption = Mld1
    End If
Else
    TxPas.Text = vbNullString
    Lbl02.Caption = Mld1
End If

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPass " & Err.Number
Resume Next

End Sub
Private Sub FUnloa()
On Error GoTo MnErr

Dim AktZa As Integer

For AktZa = 0 To GlZaR - 1 'Rechteanzahl
    GlRch(0, AktZa) = 0
Next AktZa
DoEvents

Select Case GlSHt
Case ShoCut_Start:
        GlBut = GlBu0
Case ShoCut_Adresse:
        GlBut = GlBu1
Case ShoCut_Kranken:
        GlBut = GlBu2
Case ShoCut_Finanz:
        GlBut = GlBu3
Case ShoCut_Termin:
        GlBut = GlBu4
Case ShoCut_Labor:
        GlBut = GlBu5
Case ShoCut_Texte:
        GlBut = GlBu6
Case ShoCut_Katalog:
        GlBut = GlBu7
Case ShoCut_Abrechn:
        GlBut = GlBu8
End Select

GlRes = True 'Reset der Einstellungen
STaSe ShoCut_Start, 0
DoEvents
SRecht
DoEvents
Unload Me

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FUnloa " & Err.Number
Resume Next

End Sub
Private Sub btnSchließ_Click()
    FUnloa
End Sub
Private Sub btnWeiter_Click()
    FPass
End Sub
Private Sub cmbMitar_Click()
    FKont
    S_MaGrp
    GrMa_Lad True
End Sub
Private Sub Form_Load()
On Error Resume Next

FLoad
AFont Me
Me.BackColor = GlBak
Me.lblLab02.BackColor = GlBak
Me.lblLab03.BackColor = GlBak
Me.lblLab04.BackColor = GlBak
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub txtPassw_GotFocus()
    Me.txtPassw.SelStart = 0
    Me.txtPassw.SelLength = Len(Me.txtPassw.Text)
End Sub
