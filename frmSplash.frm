VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'Kein
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PopupControl popCont4 
      Left            =   600
      Top             =   240
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clFen As clsFenster
Private Sub Form_Load()
On Error GoTo WiErr

Dim ScreX As Long
Dim ScreY As Long
Dim FiNam As String
Dim Popu4 As XtremeSuiteControls.PopupControl

Set Popu4 = Me.popCont4

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

FiNam = App.Path & "\Skins\Splash.skn"

With clFen
    ScreX = .FenGro(3)
    ScreY = .FenGro(4)
    .FeLin = 0
    .FeObn = 0
    .FeBre = 0
    .FeHoh = 0
    .FenMov
End With

With Popu4
    .RemoveAllItems
    .Icons.LoadBitmap FiNam, 100, xtpImageNormal
    .BackgroundBitmap = 100
    .AllowMove = True
    .AnimateDelay = 200
    .ShowDelay = 9000
    .Animation = xtpPopupAnimationNone
    .Bottom = (ScreY / 2) + 75
    .Right = (ScreX / 2) + 300
    .Height = 150
    .Width = 600
    .Show
    .Animation = xtpPopupAnimationFade
End With

Set Popu4 = Nothing
Set clFen = Nothing

TimInit 4, 30 'Max. Länge des Splashscreens

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " frmSplash " & Err.Number
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
    TimEnde 4
    TimEnde 5
    Set frmSplash = Nothing
End Sub

