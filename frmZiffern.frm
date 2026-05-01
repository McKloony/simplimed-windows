VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmZiffern 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Eintrag W‰hlen"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
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
      TabIndex        =   1
      Top             =   3100
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
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
         Caption         =   "&Einf¸gen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   1300
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
   Begin XtremeSuiteControls.ListView lstView1 
      Height          =   2300
      Left            =   200
      TabIndex        =   0
      Top             =   660
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   4057
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Der von Ihnen gesuchte Eintrag wurde mehrfach gefunden. Bitte w‰hlen Sie den gew¸nschten Eintrag und klicken auf Einf¸gen."
      Height          =   500
      Left            =   200
      TabIndex        =   5
      Top             =   160
      Width           =   5500
   End
End
Attribute VB_Name = "frmZiffern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private LiVw1 As XtremeSuiteControls.ListView
Private LiItm As XtremeSuiteControls.ListViewItem
Private LiIts As XtremeSuiteControls.ListViewItems

Private TagWe As String

Private clFen As clsFenster
Private Sub FLoad()
On Error GoTo SeErr
'L‰dt die ausgew‰hlte Adresse in das Adressformular

Dim ImMan As XtremeCommandBars.ImageManager

Set ImMan = frmMain.imgManag
Set LiVw1 = Me.lstView1
Set Rahm0 = Me.frmRahm0

With LiVw1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = True
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = False
    .FlatScrollBar = False
    .Font.SIZE = 10 'GlTFt.size
    .Font.Name = GlTFt.Name
    .ForeColor = vbBlack
    .FullRowSelect = True
    .GridLines = False
    .HideColumnHeaders = False
    .HideSelection = False
    .HotTracking = False
    .HoverSelection = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpListViewLabelManual
    .LabelWrap = True
    .MultiSelect = False
    .OLEDropMode = xtpOLEDropNone
    .View = xtpListViewReport
End With

With LiVw1
    .ColumnHeaders.Add 1, , "Nummer", 1200
    .ColumnHeaders.Add 2, , "Bezeichnungstext", 3700
    .ColumnHeaders.Add 3, , "ID0", 1
End With

Rahm0.BackColor = GlBak

Set LiVw1 = Nothing

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FSett()
On Error GoTo SeErr

Dim GesZa As Long
Dim IdxNr As Long
Dim CoStr As String

Set LiVw1 = Me.lstView1
Set LiIts = LiVw1.ListItems

GesZa = LiIts.Count

If GesZa > 0 Then
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            CoStr = LiItm.Text
            IdxNr = CLng(LiItm.SubItems(2))
            Exit For
        End If
    Next LiItm
    Unload Me
    DoEvents
    K_Eing 4, , IdxNr
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
    Set frmZiffern = Nothing
End Sub
Private Sub lstView1_DblClick()
    FSett
End Sub
Private Sub lstView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub

