VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmTSEInit 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TSE Initialisierung"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.PushButton btnFormat 
      Height          =   300
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5000
      Width           =   800
      _Version        =   1048579
      _ExtentX        =   1411
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "&Format"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
      Height          =   400
      Left            =   3400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3740
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Schlieﬂen"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   4600
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtTSEIn 
      Height          =   3400
      Left            =   130
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   100
      Width           =   7690
      _Version        =   1048579
      _ExtentX        =   13564
      _ExtentY        =   5997
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
End
Attribute VB_Name = "frmTSEInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control

Private clFen As clsFenster
Private Sub FForm()
On Error Resume Next

Dim RetWe As Long
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

TeTit = "TSE Einrichtung"
TeMai = "Mˆchten Sie den TSE Stick jetzt einrichten?"
TeInh = "Der TSE Stick wird mit den Kannennahmen : " & GlTSN & " eingerichtet."
TeFus = "Der TSE Stick hat den Key : " & GlTSK

RetWe = DirectReadDriveNT(GLTSL & "TSE_COMM.DAT", 0, 0, RuAry(), 512)

If RetWe = 0 Then
    TxTSE.Text = TxTSE.Text & vbCrLf & "Der TSE Stick unter " & GLTSL & " kann nicht gefunden werden!"
    Exit Sub
End If

SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
If GlMes = 33565 Then
    Screen.MousePointer = vbHourglass
    DoEvents
    
    TSE_Neue
    
    DoEvents
    Screen.MousePointer = vbNormal
End If

End Sub

Private Sub FSett()
On Error GoTo OrErr

Dim SuStr As String
Dim TmStr As String
Dim Posi1 As Integer
Dim Posi2 As Integer
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

SuStr = TxTSE.Text

If GlTSE > 0 Then
    Select Case GlTSE
    Case 1:
        Posi1 = InStr(1, SuStr, "einrichten", 1)
        Posi2 = InStr(1, SuStr, "formatieren", 1)
        If Posi1 > 0 Then
            TmStr = Replace(SuStr, "einrichten", vbNullString, 1)
            TxTSE.Text = TmStr
            FForm
        End If
        If Posi2 > 0 Then
            TmStr = Replace(SuStr, "formatieren", vbNullString, 1)
            TxTSE.Text = TmStr
            FForm
        End If
    Case 2:
        Posi1 = InStr(1, SuStr, "deaktivieren", 1)
        If Posi1 > 0 Then
            TmStr = Replace(SuStr, "deaktivieren", vbNullString, 1)
            TxTSE.Text = TmStr
            TSEDisa
        End If
    End Select
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
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

End Sub

Private Sub txtTSEIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub
