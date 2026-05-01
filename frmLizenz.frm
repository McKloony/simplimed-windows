VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmLizenz 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Lizenzierung"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   9
      Top             =   4600
      Width           =   6045
      _Version        =   1048579
      _ExtentX        =   10663
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAbbru 
         Height          =   400
         Left            =   4400
         TabIndex        =   13
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
         Left            =   3000
         TabIndex        =   12
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
      Begin XtremeSuiteControls.PushButton btnZuruck 
         Height          =   400
         Left            =   300
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Kopieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnLosche 
         Height          =   400
         Left            =   1600
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Löschen"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   2900
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   5115
      _StockProps     =   79
      Caption         =   "GroupBox1"
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtSchlu 
         Height          =   350
         Left            =   1200
         TabIndex        =   3
         Top             =   1070
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   600
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtKey01 
         Height          =   350
         Left            =   1200
         TabIndex        =   5
         Top             =   2140
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         MaxLength       =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtKey02 
         Height          =   350
         Left            =   2150
         TabIndex        =   6
         Top             =   2140
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         MaxLength       =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtKey04 
         Height          =   350
         Left            =   4030
         TabIndex        =   8
         Top             =   2140
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         MaxLength       =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtKey03 
         Height          =   350
         Left            =   3090
         TabIndex        =   7
         Top             =   2140
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         MaxLength       =   4
      End
      Begin XtremeSuiteControls.PushButton btnEmail 
         Height          =   400
         Left            =   4550
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Sendet den Lizenzierungsschlüssel an die angezeigte Emailadrese"
         Top             =   1040
         Width           =   1140
         _Version        =   1048579
         _ExtentX        =   2011
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Senden"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblLabe3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tragen Sie hier Ihre Seriennummer ein:"
         Height          =   220
         Left            =   1220
         TabIndex        =   21
         Top             =   1890
         Width           =   2800
      End
      Begin VB.Label lblLabe2 
         BackStyle       =   0  'Transparent
         Caption         =   "Übermitteln Sie uns diesen Schlüssel:"
         Height          =   220
         Left            =   1220
         TabIndex        =   20
         Top             =   820
         Width           =   2800
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000006&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   880
         Left            =   0
         Top             =   1820
         Width           =   6000
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000006&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Ausgefüllt
         Height          =   880
         Left            =   0
         Top             =   760
         Width           =   6000
      End
      Begin VB.Label lblLabe1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLizenz.frx":0000
         Height          =   620
         Left            =   300
         TabIndex        =   19
         Top             =   60
         Width           =   5300
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   2700
      Left            =   0
      TabIndex        =   2
      Top             =   1700
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   4762
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtKey05 
         Height          =   400
         Left            =   1200
         TabIndex        =   14
         Top             =   2140
         Width           =   3750
         _Version        =   1048579
         _ExtentX        =   6615
         _ExtentY        =   706
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Geben Sie hier den Produktvarianten-Schlüssel ein:"
         Height          =   220
         Left            =   1220
         TabIndex        =   18
         Top             =   1890
         Width           =   3700
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000006&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   880
         Left            =   0
         Top             =   1820
         Width           =   6000
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Geben Sie hier den Produktvarianten-Schlüssel ein:"
         Height          =   220
         Left            =   1220
         TabIndex        =   17
         Top             =   1850
         Width           =   3700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Geben Sie bitte jetzt noch in das gelb umrandete Feld den Schlüssel für die Produktvariante ein und klicken auf Weiter."
         Height          =   620
         Left            =   300
         TabIndex        =   16
         Top             =   0
         Width           =   5300
      End
   End
   Begin VB.PictureBox picPict1 
      BorderStyle     =   0  'Kein
      Height          =   1500
      Left            =   0
      ScaleHeight     =   1500
      ScaleWidth      =   6000
      TabIndex        =   15
      Top             =   0
      Width           =   6000
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6000
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6000
      Y1              =   1510
      Y2              =   1510
   End
End
Attribute VB_Name = "frmLizenz"
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
Private TxKey As XtremeSuiteControls.FlatEdit
Private TxSe1 As XtremeSuiteControls.FlatEdit
Private TxSe2 As XtremeSuiteControls.FlatEdit
Private TxSe3 As XtremeSuiteControls.FlatEdit
Private TxSe4 As XtremeSuiteControls.FlatEdit
Private TxSe5 As XtremeSuiteControls.FlatEdit
Private Butt1 As XtremeSuiteControls.PushButton
Private Butt2 As XtremeSuiteControls.PushButton

Private Const SeKey = "Seriennummer"
Private Const VaKey = "Variante"
Private Const LiKey = "Lizenz"
Private Const MaKey = "M33A1114S88M3338"

Private clLiz As clsLizenz
Private clFil As clsFile
Private Sub ALoad()
On Error Resume Next

Dim FiNam As String
Dim RegNr As String
Dim SerNr As String
Dim VarNr As String
Dim LizNr As String
Dim Mld1, Mld2, Mld3, Mld4, Tit1 As String

Set TxKey = Me.txtSchlu
Set TxSe1 = Me.txtKey01
Set TxSe2 = Me.txtKey02
Set TxSe3 = Me.txtKey03
Set TxSe4 = Me.txtKey04
Set TxSe5 = Me.txtKey05
Set Butt1 = Me.btnWeiter
Set Butt2 = Me.btnZuruck
Set Pict1 = Me.picPict1
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2

Set clLiz = New clsLizenz
Set clFil = New clsFile

Tit1 = "Programmlizenzierung"
FiNam = gAnPfa & gIniNa
Mld1 = "Die in der Lizenzdatei gespeicherte Seriennummer zu dieser Software hat keine 16 Zeichen"
Mld2 = "Die in der Lizenzdatei gespeicherte Seriennummer zu dieser Software ist falsch. " & _
"Dieses kann mehrere Ursachen haben. Entweder wurde die Lizenzdatei bzw. deren Inhalt ausgetauscht" & _
"oder es wurde eine Hardwarekomponente wie CPU, Netzwerkkarte, Festplatte oder Grafikkarte ausgewechselt bzw. diese Software auf einem neuen PC installiert."
Mld3 = "Der in der Lizenzdatei gespeicherte Produktvarianten-Schlüssel hat keine 6 Zeichen"
Mld4 = "Der in der Lizenzdatei gespeicherte Produktvarianten-Schlüssel ist falsch."

If clFil.FilVor(gFrmGr) = True Then
    Pict1.Picture = LoadPicture(gFrmGr)
End If

With clLiz
    .AnwNa = gAnwKy
    .KyHar = gHrdKy
    .KySer = gSerKy
    .LiHdSer
    .LiGeSer
    TxKey.Text = .HdSer & "SM" & App.Major & App.Minor & App.Revision
    SerNr = .SerNu
End With

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak

If clFil.FilVor(FiNam) = True Then
    RegNr = IniGetFil(FiNam, gIniAn, SeKey)
    If Len(RegNr) = 16 Then
        TxSe1.Text = Mid$(RegNr, 1, 4)
        TxSe2.Text = Mid$(RegNr, 5, 4)
        TxSe3.Text = Mid$(RegNr, 9, 4)
        TxSe4.Text = Mid$(RegNr, 13, 4)
        If SerNr = RegNr Then
            Butt1.Enabled = False
            TxSe1.Enabled = False
            TxSe2.Enabled = False
            TxSe3.Enabled = False
            TxSe4.Enabled = False
            If gSoVar = True Then
                VarNr = IniGetFil(FiNam, gIniAn, VaKey)
                If Len(VarNr) = 6 Then
                    TxSe5.Enabled = False
                ElseIf Len(VarNr) > 0 Then
                    MsgBox Mld3, vbExclamation, Tit1
                End If
            End If
        Else
            LizNr = IniGetFil(FiNam, gIniAn, LiKey)
            If LizNr = "OK" Then
                Butt1.Enabled = False
                TxSe1.Enabled = False
                TxSe2.Enabled = False
                TxSe3.Enabled = False
                TxSe4.Enabled = False
                TxSe5.Enabled = False
            Else
                MsgBox Mld2, vbExclamation, Tit1
            End If
        End If
    ElseIf Len(RegNr) > 0 Then
        MsgBox Mld1, vbExclamation, Tit1
    End If
End If

Set clFil = Nothing
Set clLiz = Nothing

End Sub
Private Sub ASave()
On Error Resume Next

Dim FiNam As String
Dim RegNr As String
Dim SerNr As String
Dim VarNr As String
Dim Mld1, Mld2, Mld3  As String
Dim Mld4, Mld5, Tit1 As String

Set TxKey = Me.txtSchlu
Set TxSe1 = Me.txtKey01
Set TxSe2 = Me.txtKey02
Set TxSe3 = Me.txtKey03
Set TxSe4 = Me.txtKey04
Set TxSe5 = Me.txtKey05
Set Butt1 = Me.btnWeiter
Set Butt2 = Me.btnZuruck

Set clLiz = New clsLizenz

VarNr = UCase(TxSe5.Text)

Tit1 = "Programmlizenzierung"
FiNam = gAnPfa & gIniNa
Mld1 = "Der Lizenzierungsschlüssel hat keine 8 Zeichen"
Mld2 = "Der von Ihnen eingegebene Seriennummer hat keine 16 Zeichen"
Mld3 = "Die von Ihnen eingegebene Seriennummer ist falsch. Bitte prüfen Sie diese und versuchen es erneut."
Mld4 = "Die Lizenzierung wurde erfolgreich abgeschlossen. Bitte starten Sie SimpliMed erneut, damit die Lizenzierung wirksam werden kann." & vbCrLf & vbCrLf & _
"HINWEIS! Die Lizenzinformationen werden in der Datei Lizenzierung.ini abgespeichert. Diese befindet sich im Ordner: " & gAnPfa & _
" Bitte kopieren Sie diese auf eine Diskette und bewahren Sie diese gut auf."
Mld5 = "Der von Ihnen eingegebene Produktvarianten-Schlüssel hat keine 6 Zeichen"

If Len(TxKey.Text) >= 8 Then
    RegNr = UCase(Right("0000" & TxSe1.Text, 4) & Right("0000" & TxSe2.Text, 4) & Right("0000" & TxSe3.Text, 4) & Right("0000" & TxSe4.Text, 4))
    If Len(RegNr) = 16 Then
        With clLiz
            .AnwNa = gAnwKy
            .KyHar = gHrdKy
            .KySer = gSerKy
            .LiGeSer
            SerNr = .SerNu
        End With
        If RegNr = SerNr Then
            If gSoVar = True Then
                If Len(VarNr) = 6 Then
                    IniSetFil FiNam, gIniAn, SeKey, RegNr
                    IniSetFil FiNam, gIniAn, VaKey, VarNr
                    MsgBox Mld4, vbInformation, Tit1
                    Unload Me
                Else
                    MsgBox Mld5, vbExclamation, Tit1
                End If
            Else
                IniSetFil FiNam, gIniAn, SeKey, RegNr
                MsgBox Mld4, vbInformation, Tit1
                Unload Me
            End If
        Else
            If RegNr = MaKey Then
                IniSetFil FiNam, gIniAn, SeKey, RegNr
                IniSetFil FiNam, gIniAn, VaKey, VarNr
                IniSetFil FiNam, gIniAn, LiKey, "OK"
                MsgBox Mld4, vbInformation, Tit1
                Unload Me
            Else
                MsgBox Mld3, vbCritical, Tit1
            End If
        End If
    Else
        MsgBox Mld2, vbExclamation, Tit1
    End If
Else
    MsgBox Mld1, vbExclamation, Tit1
End If

Set clLiz = Nothing

End Sub
Private Sub ADele()
On Error Resume Next

Dim FiNam As String
Dim Frage, Dialo As Integer
Dim Mld1, Tit1 As String

Set TxKey = Me.txtSchlu
Set TxSe1 = Me.txtKey01
Set TxSe2 = Me.txtKey02
Set TxSe3 = Me.txtKey03
Set TxSe4 = Me.txtKey04
Set TxSe5 = Me.txtKey05
Set Butt1 = Me.btnWeiter
Set Butt2 = Me.btnZuruck
Set Pict1 = Me.picPict1

Set clFil = New clsFile

FiNam = gAnPfa & gIniNa
Dialo = vbYesNo + vbQuestion

Tit1 = "Programmlizenzierung"
Mld1 = "Möchten Sie den Lizenzierungscode wirklich löschen?" & vbCrLf & vbCrLf & "ACHTUNG! Beim löschen des Lizenzierungscodes wird Ihre Lizenz gelöscht und Sie müssen diese erneut eingeben."

Frage = MsgBox(Mld1, Dialo, Tit1)
If Frage = 6 Then
    If clFil.FilVor(FiNam) = True Then
        IniDelSek FiNam, gIniAn
        TxSe1.Text = vbNullString
        TxSe2.Text = vbNullString
        TxSe3.Text = vbNullString
        TxSe4.Text = vbNullString
        TxSe5.Text = vbNullString
        Butt1.Enabled = True
        TxSe1.Enabled = True
        TxSe2.Enabled = True
        TxSe3.Enabled = True
        TxSe4.Enabled = True
        TxSe5.Enabled = True
    End If
End If

Set clFil = Nothing

End Sub
Private Sub AFnts(ByVal FoNam As Form)
On Error GoTo ReErr

Set FM = FoNam

For Each AktCo In FM.Controls
    Select Case TypeName(AktCo)
    Case "GroupBox":
            With AktCo
                .Font.Name = GlTFt.Name
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "PushButton":
            With AktCo
                .Font.Name = GlTFt.Name
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "Label":
            With AktCo
                .Font.Name = GlTFt.Name
            End With
    Case "RadioButton":
            With AktCo
                .Font.Name = GlTFt.Name
                .Font.SIZE = 8
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "CheckBox":
            With AktCo
                .Font.Name = GlTFt.Name
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "FlatEdit":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
                .Font.Name = GlTFt.Name
                .ForeColor = -2147483641
            End With
    Case "ComboBox":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
                .Font.Name = GlTFt.Name
                .ForeColor = -2147483641
            End With
    Case "ListBox":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
                .Font.Name = GlTFt.Name
                .ForeColor = -2147483641
            End With
    Case "UpDown":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    End Select
Next AktCo

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AFnts " & Err.Number
Resume Next

End Sub

Private Function ACryp(ByVal SuStr As String, ByVal Flag As Boolean) As String
On Error Resume Next
'Ver/- und Entschlüsselt einen String

Dim AktZa As Long
Dim Posit As Long
Dim CptZa As Long
Dim OrgZa As Long
Dim KeyZa As Long
Dim CptSt As String
Dim KeOut As String
  
For AktZa = 1 To Len(SuStr)
    Posit = Posit + 1
    If Posit > Len(gVerKy) Then Posit = 1
    KeyZa = Asc(Mid(gVerKy, Posit, 1))
    If Flag = True Then ' Verschlüsseln
        OrgZa = Asc(Mid(SuStr, AktZa, 1))
        CptZa = OrgZa Xor KeyZa
        CptSt = Hex(CptZa)
        If Len(CptSt) < 2 Then CptSt = "0" & CptSt
        KeOut = KeOut & CptSt
    Else 'Entschlüsseln
        If AktZa > Len(SuStr) \ 2 Then Exit For
        CptZa = CByte("&H" & Mid$(SuStr, AktZa * 2 - 1, 2))
        OrgZa = CptZa Xor KeyZa
        KeOut = KeOut & Chr$(OrgZa)
    End If
Next AktZa
 
ACryp = KeOut

End Function
Private Sub AWeit()
On Error Resume Next

Dim RegNr As String

Set TxSe1 = Me.txtKey01
Set TxSe2 = Me.txtKey02
Set TxSe3 = Me.txtKey03
Set TxSe4 = Me.txtKey04
Set TxSe5 = Me.txtKey05
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Butt1 = Me.btnWeiter
Set Butt2 = Me.btnZuruck

If Len(TxSe1.Text) <> 4 Then Exit Sub
If Len(TxSe2.Text) <> 4 Then Exit Sub
If Len(TxSe3.Text) <> 4 Then Exit Sub
If Len(TxSe1.Text) <> 4 Then Exit Sub

RegNr = TxSe1.Text & TxSe2.Text & TxSe3.Text & TxSe4.Text

If gSoVar = True Then
    If Rahm2.Visible = True Then
        ASave
    Else
        If Len(RegNr) = 16 Then
            Rahm1.Visible = False
            Rahm2.Visible = True
            Butt2.Enabled = True
            TxSe5.SetFocus
        End If
    End If
Else
    ASave
End If

End Sub
Private Sub btnAbbru_Click()
    Unload Me
End Sub

Private Sub btnEmail_Click()
On Error Resume Next

Dim EmAdr As String
Dim MaTex As String
Dim LzStr As String
Dim EmVer As Boolean
Dim Frage, Dialo As Integer

Set TxKey = Me.txtSchlu

LzStr = TxKey.Text

If GlMaV = False Then
    WindowMess "Es wurden noch keine Benutzerdaten bzw. Mandantendaten eingetargen", Dial2, "Keine Mandanten vorhanden", Me.hwnd
    Exit Sub
End If

If GlThe(GlSMa, 16) = vbNullString Then
    WindowMess "Bei den Benutzer- bzw. Mandantendaten wurde noch keine Emailadresse eingetargen", Dial2, "Keine Emailadresse vorhanden", Me.hwnd
    Exit Sub
End If

If GlThe(GlSMa, 2) = "Mustermann" Then
    WindowMess "Die Benutzer- bzw. Mandantendaten wurde noch nicht an Ihre Praxisdaten angepasst", Dial2, "Mandantendaten nicht angepasst", Me.hwnd
    Exit Sub
End If

If GlThe(GlSMa, 16) = "info@praxis-mustermann.de" Then
    WindowMess "Bei den Benutzer- bzw. Mandantendaten wurde die Emailadresse noch nicht angepasst", Dial2, "Emailadresse nicht angepasst", Me.hwnd
    Exit Sub
End If

MaTex = MaTex & vbCrLf & GlThe(GlSMa, 13)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 19) 'Praxisname
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 1)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 2)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 3)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 4)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 5)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 6)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 7)
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 16) 'Email
MaTex = MaTex & vbCrLf & GlThe(GlSMa, 12) 'Beruf
MaTex = MaTex & vbCrLf & vbCrLf & LzStr

Frage = MsgBox("Sollen der Lizenzierungsschlüssel: " & LzStr & " nun an die SimpliMed GmbH gesendet werden?", Dial1, "Lizenzierung per E-Mail")
If Frage = 6 Then
    EmAdr = "shop@simplimed.de"
    EmVer = SEmSe(EmAdr, "SimpliMed Lizenzierung", MaTex, , , , False)
End If

End Sub
Private Sub btnLosche_Click()
    ADele
End Sub
Private Sub btnWeiter_Click()
    AWeit
End Sub
Private Sub btnZuruck_Click()

Set TxKey = Me.txtSchlu

Clipboard.Clear
Clipboard.SetText TxKey.Text

End Sub
Private Sub Form_Load()
On Error Resume Next

ALoad
AFnts Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmLizenz = Nothing
End Sub
Private Sub txtKey01_GotFocus()
    Me.txtKey01.SelStart = 0
    Me.txtKey01.SelLength = Len(Me.txtKey01.Text)
End Sub
Private Sub txtKey01_KeyUp(KeyCode As Integer, Shift As Integer)

Set TxSe1 = Me.txtKey01
Set TxSe2 = Me.txtKey02

TxSe1.Text = UCase(TxSe1.Text)
TxSe1.SelStart = Len(TxSe1.Text)

If Len(TxSe1.Text) >= 4 Then TxSe2.SetFocus

End Sub
Private Sub txtKey02_GotFocus()
    Me.txtKey02.SelStart = 0
    Me.txtKey02.SelLength = Len(Me.txtKey02.Text)
End Sub
Private Sub txtKey02_KeyUp(KeyCode As Integer, Shift As Integer)

Set TxSe2 = Me.txtKey02
Set TxSe3 = Me.txtKey03

TxSe2.Text = UCase(TxSe2.Text)
TxSe2.SelStart = Len(TxSe2.Text)

If Len(TxSe2.Text) >= 4 Then TxSe3.SetFocus

End Sub
Private Sub txtKey03_GotFocus()
    Me.txtKey03.SelStart = 0
    Me.txtKey03.SelLength = Len(Me.txtKey03.Text)
End Sub
Private Sub txtKey03_KeyUp(KeyCode As Integer, Shift As Integer)

Set TxSe3 = Me.txtKey03
Set TxSe4 = Me.txtKey04

TxSe3.Text = UCase(TxSe3.Text)
TxSe3.SelStart = Len(TxSe3.Text)

If Len(TxSe3.Text) >= 4 Then TxSe4.SetFocus

End Sub
Private Sub txtKey04_GotFocus()
    Me.txtKey04.SelStart = 0
    Me.txtKey04.SelLength = Len(Me.txtKey04.Text)
End Sub
Private Sub txtKey04_KeyUp(KeyCode As Integer, Shift As Integer)

Set TxSe4 = Me.txtKey04

TxSe4.Text = UCase(TxSe4.Text)
TxSe4.SelStart = Len(TxSe4.Text)

If Len(TxSe4.Text) >= 4 Then txtDummy.SetFocus

End Sub

