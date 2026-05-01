VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmMailAdr 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Gefundene Adressen"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   3
      Top             =   5100
      Width           =   5100
      _Version        =   1048579
      _ExtentX        =   8996
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   3100
         TabIndex        =   6
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
         Left            =   1700
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zuordnen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   400
         TabIndex        =   4
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
   Begin XtremeSuiteControls.ListBox lstList1 
      Height          =   2520
      Left            =   140
      TabIndex        =   0
      Top             =   480
      Width           =   4700
      _Version        =   1048579
      _ExtentX        =   8290
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
   Begin XtremeSuiteControls.FlatEdit txMaiAdr 
      Height          =   1800
      Left            =   140
      TabIndex        =   2
      Top             =   3140
      Width           =   4700
      _Version        =   1048579
      _ExtentX        =   8290
      _ExtentY        =   3175
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      MultiLine       =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Es wurden die folgenden Eintr‰ge gefunden:"
      Height          =   200
      Left            =   160
      TabIndex        =   1
      Top             =   200
      Width           =   4500
   End
End
Attribute VB_Name = "frmMailAdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private FLis1 As XtremeSuiteControls.ListBox
Private Rahm0 As XtremeSuiteControls.GroupBox

Public PaPrx As String
Public PaVor As String
Public PaNam As String
Public PaStr As String
Public PaPLZ As String
Public PaOrt As String
Public PaLan As String
Public PaTel As String
Public PaEma As String

Private clFen As clsFenster
Private Sub FLoad()
On Error GoTo SeErr
'L‰dt die ausgew‰hlte Adresse in das Adressformular

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
Private Sub FNeAd()
On Error GoTo SeErr

Dim TagWe As String
Dim FeAnr As XtremeSuiteControls.ComboBox
Dim FeLa1 As XtremeSuiteControls.ComboBox
Dim FeTit As XtremeSuiteControls.FlatEdit
Dim FeFir As XtremeSuiteControls.FlatEdit
Dim FeVor As XtremeSuiteControls.FlatEdit
Dim FeNam As XtremeSuiteControls.FlatEdit
Dim FeStr As XtremeSuiteControls.FlatEdit
Dim FePLZ As XtremeSuiteControls.FlatEdit
Dim FeOrt As XtremeSuiteControls.FlatEdit
Dim FeTel As XtremeSuiteControls.FlatEdit
Dim FeEm1 As XtremeSuiteControls.FlatEdit

Set FM = frmAdress
Set FeAnr = FM.txtS1F02
Set FeTit = FM.txtS1F03
Set FeVor = FM.txtS1F04
Set FeNam = FM.txtS1F05
Set FeStr = FM.txtS1F06
Set FePLZ = FM.txtS1F08
Set FeOrt = FM.txtS1F09
Set FeLa1 = FM.txtS1F12
Set FeFir = FM.txtS2F11
Set FeTel = FM.txtS1F16
Set FeEm1 = FM.txtS1F19

SAdre 1
DoEvents

If PaPrx <> vbNullString Then
    FeFir.Text = PaPrx
    TagWe = Mid$(FeFir.Tag, 2, Len(FeFir.Tag) - 1)
    FeFir.Tag = "1" & TagWe
    GlAdS = True
End If
If PaVor <> vbNullString Then
    FeVor.Text = PaVor
    TagWe = Mid$(FeVor.Tag, 2, Len(FeVor.Tag) - 1)
    FeVor.Tag = "1" & TagWe
    GlAdS = True
End If
If PaNam <> vbNullString Then
    FeNam.Text = PaNam
    TagWe = Mid$(FeNam.Tag, 2, Len(FeNam.Tag) - 1)
    FeNam.Tag = "1" & TagWe
    GlAdS = True
    FeAnr.ListIndex = 1
    TagWe = Mid$(FeAnr.Tag, 2, Len(FeAnr.Tag) - 1)
    FeAnr.Tag = "1" & TagWe
End If
If PaStr <> vbNullString Then
    FeStr.Text = PaStr
    TagWe = Mid$(FeStr.Tag, 2, Len(FeStr.Tag) - 1)
    FeStr.Tag = "1" & TagWe
    GlAdS = True
End If
If PaPLZ <> vbNullString Then
    FePLZ.Text = PaPLZ
    TagWe = Mid$(FePLZ.Tag, 2, Len(FePLZ.Tag) - 1)
    FePLZ.Tag = "1" & TagWe
    GlAdS = True
End If
If PaOrt <> vbNullString Then
    FeOrt.Text = PaOrt
    TagWe = Mid$(FeOrt.Tag, 2, Len(FeOrt.Tag) - 1)
    FeOrt.Tag = "1" & TagWe
    GlAdS = True
End If
If PaLan <> vbNullString Then
    FeLa1.Text = PaLan
    TagWe = Mid$(FeLa1.Tag, 2, Len(FeLa1.Tag) - 1)
    FeLa1.Tag = "1" & TagWe
    GlAdS = True
End If
If PaTel <> vbNullString Then
    FeTel.Text = PaTel
    TagWe = Mid$(FeTel.Tag, 2, Len(FeTel.Tag) - 1)
    FeTel.Tag = "1" & TagWe
    GlAdS = True
End If
If PaEma <> vbNullString Then
    FeEm1.Text = PaEma
    TagWe = Mid$(FeEm1.Tag, 2, Len(FeEm1.Tag) - 1)
    FeEm1.Tag = "1" & TagWe
    GlAdS = True
End If

Unload Me

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeAd " & Err.Number
Resume Next

End Sub
Private Sub FSett()
On Error GoTo SeErr
    
Dim IdxNr As Long
Dim IdStr As String
Dim GesZa As Integer
    
Set FLis1 = Me.lstList1

GesZa = FLis1.ListCount

If GesZa > 0 Then
    IdxNr = FLis1.ItemData(FLis1.ListIndex)
    IdStr = FLis1.Text
    S_MaMa 5, IdxNr, IdStr
End If

Unload Me
    
Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub btnHilfe_Click()
    FNeAd
End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FSett
End Sub
Private Sub Form_Load()
On Error Resume Next

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Me.BackColor = GlBak

AFont Me

clFen.FenVor

Set clFen = Nothing

SFrame 1, Me.hwnd

FLoad

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmMailAdr = Nothing
End Sub
Private Sub lstList1_DblClick()
    FSett
End Sub
Private Sub lstList1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub

