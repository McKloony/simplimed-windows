VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.ShortcutBar.v16.3.1.ocx"
Begin VB.Form frmPasswort 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Passworteingabe"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.FlatEdit txtPass1 
      Height          =   350
      Left            =   2120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1500
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      PasswordChar    =   "*"
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   4
      Top             =   2900
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
         TabIndex        =   3
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
         TabIndex        =   2
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
   Begin XtremeSuiteControls.FlatEdit txtPass2 
      Height          =   350
      Left            =   2120
      TabIndex        =   1
      Top             =   2100
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BackColor       =   16777215
      PasswordChar    =   "*"
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeShortcutBar.ShortcutCaption schCapt1 
      Height          =   800
      Left            =   0
      TabIndex        =   8
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
   Begin XtremeSuiteControls.Label lblLab04 
      Height          =   220
      Left            =   880
      TabIndex        =   7
      Top             =   2140
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Wiederholung :"
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label lblLab02 
      Height          =   240
      Left            =   100
      TabIndex        =   6
      Top             =   900
      Width           =   5800
      _Version        =   1048579
      _ExtentX        =   10231
      _ExtentY        =   423
      _StockProps     =   79
      ForeColor       =   192
      Alignment       =   2
   End
   Begin VB.Label lblLab03 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Passwort :"
      Height          =   220
      Left            =   880
      TabIndex        =   5
      Top             =   1540
      Width           =   1200
   End
End
Attribute VB_Name = "frmPasswort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private FS As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Lbl02 As XtremeSuiteControls.Label
Private TxPa1 As XtremeSuiteControls.FlatEdit
Private TxPa2 As XtremeSuiteControls.FlatEdit
Private PuBu1 As XtremeSuiteControls.PushButton
Private ShCap As XtremeShortcutBar.ShortcutCaption
Private CmAcs As XtremeCommandBars.CommandBarActions
Private PrGr1 As XtremePropertyGrid.PropertyGrid
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Public PaStr As String
Private Sub FClos()
On Error GoTo MnErr

Dim FeNam As String
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrIts As XtremePropertyGrid.PropertyGridItems
    
If WindowLoad("frmMandant") = True Then
        Set FS = frmMandant
        FS.txtS1F37.Text = vbNullString
ElseIf WindowLoad("frmMailKont") = True Then
        Set FS = frmMailKont
        FS.txtUsPas.Text = vbNullString
ElseIf WindowLoad("frmAdress") = True Then
        Set FS = frmAdress
        Set PrGr1 = FS.prpGrid1
        Set PrIts = PrGr1.Categories
        For Each PrKat In PrIts
            For Each PrItm In PrKat.Childs
                FeNam = Right$(PrItm.Tag, Len(PrItm.Tag) - 1)
                If FeNam = "Em_Pass" Then
                    PrItm.Value = vbNullString
                    Exit For
                End If
            Next PrItm
        Next PrKat
        Set FS = Nothing
End If
    
Unload Me

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FLoad()
On Error GoTo MnErr

Dim RetWe As Long
Dim PasWo As String
Dim AktZa As Integer
Dim ImMan As XtremeCommandBars.ImageManager
Dim ImIcs As XtremeCommandBars.ImageManagerIcons

Set FM = frmPasswort
Set Rahm0 = FM.frmRahm0
Set Lbl02 = FM.lblLab02
Set TxPa1 = FM.txtPass1
Set TxPa2 = FM.txtPass2
Set ShCap = FM.schCapt1
Set PuBu1 = FM.btnSchließ
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
    .Caption = "Passwortbestätigung"
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

TxPa1.Text = PaStr

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
FM.lblLab02.BackColor = GlBak
FM.lblLab03.BackColor = GlBak
FM.lblLab04.BackColor = GlBak

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo MnErr

Dim FeNam As String
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrIts As XtremePropertyGrid.PropertyGridItems

Set Lbl02 = Me.lblLab02
Set TxPa1 = Me.txtPass1
Set TxPa2 = Me.txtPass2

If TxPa1.Text <> TxPa2.Text Then
    Lbl02.Caption = "Die Passwörter stimmen nicht überein"
Else
    PaStr = TxPa2.Text
    If WindowLoad("frmMandant") = True Then
        Set FS = frmMandant
        FS.txtS1F37.Text = PaStr
    ElseIf WindowLoad("frmMailKont") = True Then
        Set FS = frmMailKont
        FS.txtUsPas.Text = PaStr
    ElseIf WindowLoad("frmAdress") = True Then
        Set FS = frmAdress
        Set PrGr1 = FS.prpGrid1
        Set PrIts = PrGr1.Categories
        For Each PrKat In PrIts
            For Each PrItm In PrKat.Childs
                FeNam = Right$(PrItm.Tag, Len(PrItm.Tag) - 1)
                If FeNam = "Em_Pass" Then
                    PrItm.Value = PaStr
                    Exit For
                End If
            Next PrItm
        Next PrKat
        Set FS = Nothing
    End If
    Unload Me
End If

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub btnSchließ_Click()
    FClos
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FLoad
AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

Set FrmEx = Nothing

End Sub

Private Sub txtPass1_GotFocus()
    Me.txtPass1.SelStart = 0
    Me.txtPass1.SelLength = Len(Me.txtPass1.Text)
End Sub

Private Sub txtPass2_GotFocus()
    Me.txtPass2.SelStart = 0
    Me.txtPass2.SelLength = Len(Me.txtPass2.Text)
End Sub
