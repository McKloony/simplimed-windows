VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Begin VB.Form frmPrintPrev 
   Caption         =   "Druckvorschau"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin XtremeCommandBars.PrintPreview prtPrev1 
      Height          =   5175
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      _Version        =   1048579
      _ExtentX        =   14631
      _ExtentY        =   9128
      _StockProps     =   1
      BackColor       =   -2147483636
      Title           =   "PrintPreview1"
      VisualTheme     =   12
      Orientation     =   2
   End
End
Attribute VB_Name = "frmPrintPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PrtPr As XtremeCommandBars.PrintPreview
Attribute PrtPr.VB_VarHelpID = -1

Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Sub FInit()
On Error GoTo PoErr

Set PrtPr = Me.prtPrev1

With PrtPr
    Select Case GlSty
    Case 7: .VisualTheme = xtpThemeOffice2013
    Case 8: .VisualTheme = xtpThemeOffice2013
    Case Else: .VisualTheme = xtpThemeRibbon
    End Select
End With

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " FInit " & Err.Number
Resume Next

End Sub
Private Sub Form_Activate()
    PrtRez
End Sub
Private Sub Form_Load()
    GlKeL = True
    FInit
End Sub
Private Sub Form_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlKeL = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    PrtRez
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmPrintPrev = Nothing
End Sub
Private Sub prtPrev1_CloseClick()
    Unload Me
End Sub
Private Sub prtPrev1_PrintClick()

Set PrtPr = Me.prtPrev1

PrtPr.ShowPrintDialog

Unload Me
    
End Sub
