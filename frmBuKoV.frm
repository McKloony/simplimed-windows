VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmBuKoV 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Sachkontenauswahl"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   0
      Top             =   3200
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4200
         TabIndex        =   1
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
         Height          =   400
         Left            =   2800
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Einfügen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   1500
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
   Begin XtremeSuiteControls.ListBox lstList1 
      Height          =   2520
      Left            =   280
      TabIndex        =   4
      Top             =   480
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   4445
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte wählen Sie einen Eintrag und drücken die Tab-Taste :"
      Height          =   200
      Left            =   280
      TabIndex        =   5
      Top             =   200
      Width           =   5500
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBuKoV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private FLis1 As XtremeSuiteControls.ListBox
Private CmAcs As XtremeCommandBars.CommandBarActions
Private TxOrt As XtremeSuiteControls.FlatEdit
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private TagWe As String
Public PaFrm As String

Private clFen As clsFenster
Private Sub FLoad()
On Error GoTo SeErr
'Lädt die ausgewählte Adresse in das Adressformular

Set FLis1 = Me.lstList1
Set Rahm0 = Me.frmRahm0

With FLis1
    .Font.Name = GlTFt.Name
    .Font.SIZE = GlTFt.SIZE
End With

Rahm0.BackColor = GlBak

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FSett()
On Error GoTo SeErr
'Lädt die ausgewählte Adresse in das Adressformular

Dim KoIDI As Long
Dim GesZa As Integer

Select Case PaFrm
Case "BuAn": Set FM = frmBuAnf
Case "BuVo": Set FM = frmBuEdit
Case "BuSe": Set FM = frmBuSer
Case "BaAb": Set FM = frmBuEdVo
Case "BaVo": Set FM = frmBaEdVo
Case "BaRe": Set FM = frmBaEdRe
Case "BuSe": Set FM = frmBuSer
End Select

Set FLis1 = Me.lstList1

GesZa = FLis1.ListCount

If GesZa > 0 Then
    KoIDI = FLis1.ItemData(FLis1.ListIndex)
    S_KtSe PaFrm, KoIDI
    DoEvents
    
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        If GlBuF < 9 Then 'Buchungsdialog
            FM.cmbBuTex.SetFocus
        End If
    Else
        If GlBuF <= 4 Then
            FM.txtHaben.SetFocus
        ElseIf GlBuF <= 7 Then
            FM.cmbBuTyp.SetFocus
        ElseIf GlBuF = 8 Then
            FM.cmbBehan.SetFocus
        End If
    End If
End If

Unload Me

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50141)
TeMai = IniGetOpt("Hilfe", 50142)
TeInh = IniGetOpt("Hilfe", 50143)
TeFus = IniGetOpt("Hilfe", 50144)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
On Error Resume Next

Dim TmStr As String
Dim FoNam As String
Dim Lerze As Integer

Select Case PaFrm
Case "BuAn":
    Set FM = frmBuAnf
    FoNam = "frmBuAnf"
Case "BuVo":
    Set FM = frmBuEdit
    FoNam = "frmBuEdit"
Case "BaVo":
    Set FM = frmBaEdit
    FoNam = "frmBaEdit"
Case "BaAb":
    Set FM = frmBuEdVo
    FoNam = "frmBuEdVo"
Case "BaRe":
    Set FM = frmBaEdRe
    FoNam = "frmBaEdRe"
Case "BuSe":
    Set FM = frmBuSer
    FoNam = "frmBuSer"
End Select

If WindowLoad(FoNam) = True Then
    If GlBuF <= 4 Then 'Buchungsdialog
        TmStr = FM.txtKonto.Text
        Lerze = InStrRev(TmStr, Chr$(32), -1, 1)
        If Lerze = 0 Then
            FM.txtKonto.Text = vbNullString
            FM.txtKonto.SetFocus
        End If
    Else
        TmStr = FM.txtHaben.Text
        Lerze = InStrRev(TmStr, Chr$(32), -1, 1)
        If Lerze = 0 Then
            FM.txtHaben.Text = vbNullString
            FM.txtHaben.SetFocus
        End If
    End If
End If

Unload Me

End Sub
Private Sub btnWeiter_Click()
    FSett
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Me.BackColor = GlBak

AFont Me

clFen.FenVor

Set clFen = Nothing

SFrame 1, Me.hwnd

FLoad

FrmEx.TopMost = True

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBuKoV = Nothing
End Sub
Private Sub lstList1_DblClick()
    FSett
End Sub

Private Sub lstList1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyTab Or KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub


