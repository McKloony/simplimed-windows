VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmWiedAnh 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Gefundene Adressen"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   1
      Top             =   3100
      Width           =   5400
      _Version        =   1048579
      _ExtentX        =   9525
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   3400
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
         Left            =   2000
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Einf¸gen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   700
         TabIndex        =   2
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
      Left            =   300
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte w‰hlen Sie einen der gefundenen Eintr‰ge :"
      Height          =   200
      Left            =   400
      TabIndex        =   5
      Top             =   200
      Width           =   3600
   End
End
Attribute VB_Name = "frmWiedAnh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private FLis1 As XtremeSuiteControls.ListBox
Private CmAcs As XtremeCommandBars.CommandBarActions
Private Rahm0 As XtremeSuiteControls.GroupBox

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

Private Sub FSett()
On Error GoTo SeErr
'L‰dt die ausgew‰hlte Adresse in das Adressformular

Dim IdxNr As Long
Dim LiIdx As Long
Dim MaNum As Long
Dim IdStr As String
Dim TagWe As String
Dim AktZa As Integer
Dim GesZa As Integer

Set FM = frmWieder
Set FLis1 = Me.lstList1

GesZa = FLis1.ListCount

If GesZa > 0 Then
    IdxNr = FLis1.ItemData(FLis1.ListIndex)
    IdStr = FLis1.Text

    If IdxNr > 0 Then
        FM.txtID0.Text = IdxNr
        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
        FM.txtID0.Tag = 1 & TagWe
    End If
    If IdStr <> vbNullString Then
        FM.txtPatie.Text = IdStr
        TagWe = Mid$(FM.txtPatie.Tag, 2, Len(FM.txtPatie.Tag) - 1)
        FM.txtPatie.Tag = 1 & TagWe
    End If

    Unload Me
End If

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

TeTit = ""
TeMai = ""
TeInh = ""
TeFus = ""

'SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

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
    Set frmWiedAnh = Nothing
End Sub
Private Sub lstList1_DblClick()
    FSett
End Sub
Private Sub lstList1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub
