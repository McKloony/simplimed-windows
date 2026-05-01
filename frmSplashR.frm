VERSION 5.00
Begin VB.Form frmSplashR 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'Kein
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   8508.708
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picPict1 
      BorderStyle     =   0  'Kein
      Height          =   2250
      Left            =   0
      ScaleHeight     =   2250
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   9380
      TabIndex        =   0
      Top             =   0
      Width           =   9380
   End
End
Attribute VB_Name = "frmSplashR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Pict1 As VB.PictureBox

Private clFen As clsFenster
Private Sub Form_Load()
On Error GoTo WiErr

Dim FiNam As String

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

clFen.FenVor

Set Pict1 = Me.picPict1

FiNam = App.Path & "\Skins\RDPSplash.skn"

If Dir$(FiNam, vbNormal) <> vbNullString Then
    Pict1.Picture = LoadPicture(FiNam)
End If

Set clFen = Nothing

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " frmSplashR " & Err.Number
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
    TimEnde 4
    TimEnde 5
    Set frmSplashR = Nothing
End Sub
