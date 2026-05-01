VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmPublizieren 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Fragebogen Publizieren"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4100
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6800
      _Version        =   1048579
      _ExtentX        =   11994
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkReduz 
         Height          =   220
         Left            =   2200
         TabIndex        =   8
         Top             =   2500
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Stammdaten-Pflichtfelder Reduzierung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkDokum 
         Height          =   220
         Left            =   2200
         TabIndex        =   9
         Top             =   2900
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Dokument(e) für Digitalunterschrift"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optFrag2 
         Height          =   220
         Left            =   2200
         TabIndex        =   7
         Top             =   1800
         Width           =   2300
         _Version        =   1048579
         _ExtentX        =   4057
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Neuanmeldeformular"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optFrag1 
         Height          =   220
         Left            =   2200
         TabIndex        =   6
         Top             =   1400
         Width           =   2300
         _Version        =   1048579
         _ExtentX        =   4057
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Patientenfragebogen"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkDoUpl 
         Height          =   220
         Left            =   2200
         TabIndex        =   10
         Top             =   3300
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Dokumentenupload für Patient"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   500
         Left            =   500
         TabIndex        =   20
         Top             =   200
         Width           =   5600
         _Version        =   1048579
         _ExtentX        =   9878
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   $"frmPublizieren.frx":0000
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   29
      Top             =   4200
      Width           =   6800
      _Version        =   1048579
      _ExtentX        =   11994
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4800
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schließen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeite 
         Height          =   400
         Left            =   3400
         TabIndex        =   31
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
      Begin XtremeSuiteControls.PushButton btnZuruk 
         Height          =   400
         Left            =   2000
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zurück"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   700
         TabIndex        =   33
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
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4100
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   6800
      _Version        =   1048579
      _ExtentX        =   11994
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView1 
         Height          =   3080
         Left            =   500
         TabIndex        =   18
         Top             =   800
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10231
         _ExtentY        =   5433
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   500
         Left            =   500
         TabIndex        =   21
         Top             =   200
         Width           =   5600
         _Version        =   1048579
         _ExtentX        =   9878
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   $"frmPublizieren.frx":0092
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4100
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6800
      _Version        =   1048579
      _ExtentX        =   11994
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkTerBu 
         Height          =   220
         Left            =   2200
         TabIndex        =   13
         Top             =   2500
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Terminbuchungsanfrage für Patient"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optFrag3 
         Height          =   220
         Left            =   2200
         TabIndex        =   11
         Top             =   1400
         Width           =   2300
         _Version        =   1048579
         _ExtentX        =   4057
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Ungruppierter Fragebogen"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optFrag4 
         Height          =   220
         Left            =   2200
         TabIndex        =   12
         Top             =   1800
         Width           =   2300
         _Version        =   1048579
         _ExtentX        =   4057
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Gruppierter Fragebogen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   500
         Left            =   800
         TabIndex        =   22
         Top             =   200
         Width           =   5600
         _Version        =   1048579
         _ExtentX        =   9878
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   "Bitte wählen Sie, ob der Fragebogen gruppiert oder ungruppiert dargestellt werden soll und klicken auf Weiter."
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   4100
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   6800
      _Version        =   1048579
      _ExtentX        =   11994
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtPraxi 
         Height          =   350
         Left            =   900
         TabIndex        =   16
         Top             =   2740
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFraBo 
         Height          =   350
         Left            =   900
         TabIndex        =   15
         Top             =   1940
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   350
         Left            =   900
         TabIndex        =   14
         Top             =   1140
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDocNa 
         Height          =   350
         Left            =   900
         TabIndex        =   17
         Top             =   3540
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab09 
         Height          =   220
         Left            =   900
         TabIndex        =   28
         Top             =   3300
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Dokumentenname:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   220
         Left            =   900
         TabIndex        =   27
         Top             =   2500
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Praxisdarstellung :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   220
         Left            =   900
         TabIndex        =   26
         Top             =   1700
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Fragebogenname :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   220
         Left            =   900
         TabIndex        =   25
         Top             =   900
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Emailadresse :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   500
         Left            =   500
         TabIndex        =   23
         Top             =   200
         Width           =   5600
         _Version        =   1048579
         _ExtentX        =   9878
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   $"frmPublizieren.frx":012D
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   4100
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   6800
      _Version        =   1048579
      _ExtentX        =   11994
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtLinke 
         Height          =   3080
         Left            =   500
         TabIndex        =   19
         Top             =   800
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10231
         _ExtentY        =   5433
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   500
         Left            =   500
         TabIndex        =   5
         Top             =   200
         Width           =   5600
         _Version        =   1048579
         _ExtentX        =   9878
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   $"frmPublizieren.frx":01C0
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   9500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
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
Attribute VB_Name = "frmPublizieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control

Private FTeEm As XtremeSuiteControls.FlatEdit
Private FTeFr As XtremeSuiteControls.FlatEdit
Private FTePr As XtremeSuiteControls.FlatEdit
Private FTeDo As XtremeSuiteControls.FlatEdit
Private FTeLi As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private ChDoc As XtremeSuiteControls.CheckBox
Private ChTer As XtremeSuiteControls.CheckBox
Private ChReD As XtremeSuiteControls.CheckBox
Private ChUpl As XtremeSuiteControls.CheckBox
Private OpFr1 As XtremeSuiteControls.RadioButton
Private OpFr2 As XtremeSuiteControls.RadioButton
Private OpFr3 As XtremeSuiteControls.RadioButton
Private OpFr4 As XtremeSuiteControls.RadioButton
Private TrLi2 As XtremeSuiteControls.TreeView
Private LiVw1 As XtremeSuiteControls.ListView
Private Knote As XtremeSuiteControls.TreeViewNode
Private LiItm As XtremeSuiteControls.ListViewItem
Private LiIts As XtremeSuiteControls.ListViewItems
Private ImMan As XtremeCommandBars.ImageManager
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private TxCoN As Tx4oleLib.TXTextControl

Public mNeAu As Boolean

Private mDaNa() As String
Private mAkZa As Integer
Private mAnDa As Integer

Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private clFil As clsFile
Private Sub FKonf()
On Error GoTo InErr

Dim BogNr As Long
Dim DaNam As String
Dim AktZa As Integer
Dim AnzDa As Integer
Dim DiNam() As String

Set FM = frmMain
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set LiVw1 = Me.lstView1
Set FTeEm = Me.txtEmail
Set FTeFr = Me.txtFraBo
Set FTePr = Me.txtPraxi
Set FTeDo = Me.txtDocNa
Set FTeLi = Me.txtLinke
Set OpFr1 = Me.optFrag1
Set OpFr2 = Me.optFrag2
Set OpFr3 = Me.optFrag3
Set OpFr4 = Me.optFrag4
Set ChDoc = Me.chkDokum
Set ChReD = Me.chkReduz
Set ChUpl = Me.chkDoUpl
Set ChTer = Me.chkTerBu
Set PuBu1 = Me.btnWeite
Set PuBu2 = Me.btnZuruk
Set ImMan = FM.imgManag
Set LiIts = LiVw1.ListItems

BogNr = Mid$(GlNod, 2, Len(GlNod) - 1)

Set clFil = New clsFile

With LiVw1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = False
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = True
    .FlatScrollBar = False
    .Font.SIZE = GlTFt.SIZE
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
    .LockRedraw = False
    .MultiSelect = True
    .OLEDropMode = xtpOLEDropNone
    .View = xtpListViewList
End With

If clFil.FilVor(GlVor & "*.txm") = True Then
    AnzDa = clFil.FilLis(GlVor, "*.txm", DiNam)
    If AnzDa > 0 Then
        For AktZa = 1 To AnzDa
            DaNam = DiNam(AktZa)
            Set LiItm = LiIts.Add(, , DaNam, IC16_Doc_Norm)
        Next AktZa
    End If
End If

If GlFDo = True Then
    ChDoc.Value = xtpChecked
End If

If GlFrS = True Then
    ChReD.Value = xtpChecked
End If

If GlFUp = True Then
    ChUpl.Value = xtpChecked
End If

If BogNr = 0 Then
    OpFr1.Enabled = False
    OpFr2.Value = True
End If

If mNeAu = True Then
    OpFr1.Enabled = False
    OpFr2.Value = True
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
OpFr1.BackColor = GlBak
OpFr2.BackColor = GlBak
OpFr3.BackColor = GlBak
OpFr4.BackColor = GlBak
ChDoc.BackColor = GlBak
ChReD.BackColor = GlBak
ChTer.BackColor = GlBak
ChUpl.BackColor = GlBak

Set LiVw1 = Nothing

Set clFil = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FPrax()
On Error GoTo InErr
'Sammelt die Praxisangaben

Dim BogNr As Long
Dim MitNr As Long
Dim ManNr As Long
Dim MaEma As String
Dim MaNam As String
Dim MaBrf As String
Dim DocNa As String
Dim DoStr As String
Dim BogNa As String
Dim AktZa As Integer

Set FM = frmMain
Set ChDoc = Me.chkDokum
Set OpFr1 = Me.optFrag1
Set OpFr2 = Me.optFrag2
Set OpFr3 = Me.optFrag3
Set OpFr4 = Me.optFrag4
Set FTeEm = Me.txtEmail
Set FTeFr = Me.txtFraBo
Set FTePr = Me.txtPraxi
Set FTeDo = Me.txtDocNa
Set LiVw1 = Me.lstView1
Set TrLi2 = FM.trvList2
Set LiIts = LiVw1.ListItems

Set Knote = TrLi2.Nodes(GlNod)

MitNr = GlMiA(GlSmI, 2)

If GlEKV = True Then 'Emailkonten vorhanden
    For AktZa = 1 To UBound(GlMkt)
        If CLng(GlMkt(AktZa, 1)) = MitNr Then
            If CBool(GlMkt(AktZa, 20)) = True Then 'Standardemailkonto
                If GlMkt(AktZa, 13) <> vbNullString Then
                    MaEma = GlMkt(AktZa, 13)
                    Exit For
                End If
            End If
        End If
    Next AktZa
End If

For AktZa = 1 To UBound(GlMiA)
    If MitNr = GlMiA(AktZa, 2) Then
        ManNr = GlMiA(AktZa, 7)
        If MaEma = vbNullString Then
            If GlMiA(AktZa, 22) <> vbNullString Then
                MaEma = GlMiA(AktZa, 22)
            End If
        End If
    End If
Next AktZa

For AktZa = 1 To UBound(GlThe) 'Mandanten
    If ManNr = GlThe(AktZa, 0) Then
        MaNam = GlThe(AktZa, 13)
        MaBrf = GlThe(AktZa, 36)
        If MaEma = vbNullString Then
            MaEma = GlThe(AktZa, 16)
        End If
        Exit For
    End If
Next AktZa

If OpFr1.Value = True Then
    BogNr = Mid$(GlNod, 2, Len(GlNod) - 1)
    BogNa = S_AnBoD(BogNr, "IDKurz")
    FTeFr.Text = BogNa
Else
    FTeFr.Text = "Neuaufnahme"
End If

If ChDoc.Value = xtpChecked Then
    If LiIts.Count > 0 Then
        For Each LiItm In LiIts
            If LiItm.Checked = True Then
                DocNa = Left$(LiItm.Text, Len(LiItm.Text) - 4)
                If Left$(DocNa, 1) = "_" Then
                    DocNa = Right$(DocNa, Len(DocNa) - 1)
                End If
                If DoStr = vbNullString Then
                    DoStr = DocNa
                Else
                    DoStr = DoStr & Chr$(59) & DocNa
                End If
            End If
        Next LiItm
        FTeDo.Text = DoStr
    End If
End If

FTeEm.Text = MaEma
FTePr.Text = MaBrf

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPrax " & Err.Number
Resume Next

End Sub
Private Sub FWart(ByVal IdSek As Integer)
On Error GoTo InErr

IdSek = IdSek * 1000

WaitForSingleObject -1, IdSek

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPrax " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo InErr

Dim BogNa As String
Dim DaNam As String
Dim DaNaO As String
Dim NeuNa As String
Dim MaEma As String
Dim MaNam As String
Dim MaBrf As String
Dim FiNam As String
Dim DocPf As String
Dim NeuPf As String
Dim FrTit As String
Dim ReStr As String
Dim FrSet As Integer
Dim FrTer As Integer
Dim FrRed As Integer
Dim FrUpl As Integer
Dim AktZa As Integer
Dim DoPru As Boolean

Set FM = frmMain
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set LiVw1 = Me.lstView1
Set FTeEm = Me.txtEmail
Set FTeFr = Me.txtFraBo
Set FTePr = Me.txtPraxi
Set FTeDo = Me.txtDocNa
Set FTeLi = Me.txtLinke
Set OpFr1 = Me.optFrag1
Set OpFr2 = Me.optFrag2
Set OpFr3 = Me.optFrag3
Set OpFr4 = Me.optFrag4
Set ChDoc = Me.chkDokum
Set ChReD = Me.chkReduz
Set ChUpl = Me.chkDoUpl
Set ChTer = Me.chkTerBu
Set PuBu1 = Me.btnWeite
Set PuBu2 = Me.btnZuruk
Set ImMan = FM.imgManag
Set TxCoN = FM.TexCont1
Set LiIts = LiVw1.ListItems

If Rahm1.Visible = True Then

    Rahm1.Visible = False
    If ChDoc.Value = xtpChecked Then
        Rahm2.Visible = True
    Else
        If OpFr1.Value = True Then
            Rahm3.Visible = True
        Else
            FPrax 'Sammelt die Praxisangaben
            If ChDoc.Value = xtpUnchecked Then
                FTeDo.Enabled = False
            End If
            Rahm4.Visible = True
        End If
    End If
    PuBu2.Enabled = True
    
ElseIf Rahm2.Visible = True Then

    Set clFil = New clsFile
    If LiIts.Count > 0 Then
        For Each LiItm In LiIts
            If LiItm.Checked = True Then
                DaNam = LiItm.Text
                If DaNam <> vbNullString Then
                    FiNam = GlVor & DaNam
                    If clFil.FilVor(FiNam) = True Then
                        mAkZa = mAkZa + 1
                        ReDim Preserve mDaNa(mAkZa)
                        mDaNa(mAkZa) = DaNam
                        STxFr
                        DoEvents
                        STxNe
                        DoEvents
                        TxCoN.Load FiNam
                        DoEvents
                        If STxPl(True, True) = False Then 'Prüft die Platzhalter im Text
                            TxCoN.Text = vbNullString
                            FTeLi.Text = "Im Dokument: " & DaNam & " fehlen die notwendigen Platzhalter." & vbCrLf & vbCrLf & "Soll ein Dokument digital unterschrieben werden, ist es erforderlich zu kennzeichnen, an welcher Stelle das Dokument unterzeichnet werden soll. Dazu fügen Sie an die gewünschte Position die Zeichenkette $$$ ein." & vbCrLf & vbCrLf & "Sollen die Adressdaten des Patienten in das Dokument eingefügt werden, ist es erforderlich zu kennzeichnen, an welcher Stelle das Dokuments diese dargestellt werden soll. Dazu fügen Sie an die gewünschte Position die Zeichenkette &&&1 oder &&&2 und gegebenenfalls auch &&&3 für die Briefanrede ein."
                            Rahm2.Visible = False
                            Rahm5.Visible = True
                            DoPru = True
                            Exit For
                        End If
                    Else
                        DoPru = True
                        Exit For
                    End If
                Else
                    DoPru = True
                    Exit For
                End If
            End If
        Next LiItm
        TxCoN.Text = vbNullString
        mAnDa = mAkZa
    End If
    Set clFil = Nothing
    DoEvents
    
    If mAnDa > 1 Then
        FTeDo.Enabled = False
    End If
    
    If DoPru = False Then
        Rahm2.Visible = False
        If OpFr1.Value = True Then
            Rahm3.Visible = True
        Else
            FPrax 'Sammelt die Praxisangaben
            Rahm4.Visible = True
        End If
    End If

ElseIf Rahm3.Visible = True Then

    Rahm3.Visible = False
    FPrax 'Sammelt die Praxisangaben
    If ChDoc.Value = xtpUnchecked Then
        FTeDo.Enabled = False
    End If
    Rahm4.Visible = True
    
ElseIf Rahm4.Visible = True Then

    If FTeEm.Text = vbNullString Then
        SPopu "Keine Emailadresse", "Es wurde keine Emailadresse angegeben", IC48_Forbidden
        Exit Sub
    End If
    If FTePr.Text = vbNullString Then
        SPopu "Keine Praxisangaben", "Es wurde keine Praxisdaten angegeben", IC48_Forbidden
        Exit Sub
    End If
    If FTeFr.Text = vbNullString Then
        SPopu "Kein Fragebogenname", "Es wurde kein Fragebogenname angegeben", IC48_Forbidden
        Exit Sub
    End If
    If ChDoc.Value = xtpChecked Then
        If FTeDo.Text = vbNullString Then
            SPopu "Kein Dokumentenname", "Es wurde keine Dokumentenname angegeben", IC48_Forbidden
            Exit Sub
        End If
    End If
    If OpFr3.Value = True Then
        FrSet = 0
    Else
        FrSet = 1
    End If
    If ChTer.Value = xtpChecked Then
        FrTer = 1
    Else
        FrTer = 0
    End If
    If ChReD.Value = xtpChecked Then
        FrRed = 1
    Else
        FrRed = 0
    End If
    If ChUpl.Value = xtpChecked Then
        FrUpl = 1
    Else
        FrUpl = 0
    End If

    MaEma = FTeEm.Text
    BogNa = FTeFr.Text
    MaBrf = FTePr.Text
    FrTit = FTeDo.Text

    If ChDoc.Value = xtpChecked Then
        Set clFil = New clsFile
        For AktZa = 1 To mAnDa
            DaNam = mDaNa(AktZa)
            FiNam = GlVor & DaNam
            DaNaO = Left$(DaNam, Len(DaNam) - 4)
            If clFil.FilVor(FiNam) = True Then
                Screen.MousePointer = vbHourglass
                DoEvents
                
                DaNaO = SNaFi(DaNaO, True, True, True)
                If Left$(DaNaO, 1) = "_" Then
                    DaNaO = Right$(DaNaO, Len(DaNaO) - 1)
                End If
                NeuNa = CreateID("D") & "_" & DaNaO & ".pdf"
                NeuPf = GlEPf & NeuNa
                If clFil.FilVor(NeuPf) = True Then
                    clFil.DaLoe = NeuPf & vbNullChar
                    clFil.FilLoe
                    DoEvents
                End If
                STxFr 'Formatanpassungen Textcontrol
                DoEvents
                STxNe 'Legt eine neues leeres Textdokument an
                DoEvents
                TxCoN.Load FiNam
                DoEvents
                S_TxSer 'Einlesen der Daten für die Datenfelder
                DoEvents
                STxV4 'Verbindet die Datenfelder mit den Variabler GlSer
                DoEvents
                TxCoN.Save NeuPf, 0, 12 'PDF
                DoEvents
                If GlWZe > 0 Then 'Wartezeit in Sekunden
                    FWart GlWZe
                End If
                If clFil.FilVor(NeuPf) = False Then
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    TxCoN.Save NeuPf, 0, 12 'PDF
                    DoEvents
                End If
                TxCoN.Text = vbNullString
                If DocPf = vbNullString Then
                    DocPf = NeuPf
                Else
                    DocPf = DocPf & ";" & NeuPf
                End If
                
                DoEvents
                Screen.MousePointer = vbNormal
            End If
            DoEvents
        Next AktZa
        Set clFil = Nothing
    End If
    
    Rahm4.Visible = False
    Rahm5.Visible = True
    PuBu1.Enabled = False
    DoEvents

    If OpFr1.Value = True Then
        FTeLi.Text = vbCrLf & "Der Fragebogen wird publiziert..." & vbCrLf
    Else
        FTeLi.Text = vbCrLf & "Das Neuaufnahmeformular wird publiziert..." & vbCrLf
    End If
    If OpFr1.Value = True Then
        ReStr = S_AnBoV(FrSet, BogNa, MaEma, MaBrf, DocPf, FrTit, FrTer, FrRed, FrUpl, 0)
    Else
        ReStr = S_AnBoY(BogNa, MaEma, MaBrf, DocPf, FrTit, FrTer, FrRed, FrUpl, 0)
    End If
    DoEvents
    If OpFr1.Value = True Then
        FTeLi.Text = FTeLi.Text & vbCrLf & "Die URL zu Ihrem Fragebogen lautet:" & vbCrLf & vbCrLf & ReStr & vbCrLf & vbCrLf & "Dieser Weblink befindet sich jetzt in Ihrer Zwischenablage, so dass dieser an anderer Stelle wieder eingefügt und verwendet werden kann. Dieser Weblink kann jederzeit über das Kontextmenü des Fragebogens erneut in die Zwischenablage kopiert werden." & vbCrLf & vbCrLf & "Die Darstellung des Fragebogens, sowohl als Inlineframe auf Ihrer Webiste als auch Standallone, kann durch unterschliedliche CCS Ergänzungen des Frames verändert werden (siehe Dokumentation)."
    Else
        FTeLi.Text = FTeLi.Text & vbCrLf & "Die URL zu Ihrem Neuaufnahmeformular lautet:" & vbCrLf & vbCrLf & ReStr & vbCrLf & vbCrLf & "Dieser Weblink befindet sich jetzt in Ihrer Zwischenablage, so dass dieser an anderer Stelle wieder eingefügt und verwendet werden kann." & vbCrLf & vbCrLf & "Die Darstellung des Neuaufnahmeformulars, sowohl als Inlineframe auf Ihrer Webiste als auch Standallone, kann durch unterschliedliche CCS Ergänzungen des Frames verändert werden (siehe Dokumentation)."
    End If
    FTeLi.SelStart = Len(FTeLi.Text)
    
ElseIf Rahm5.Visible = True Then

    If mDaNa(mAkZa) <> vbNullString Then
        DaNam = mDaNa(mAkZa)
        FiNam = GlVor & DaNam
        GlBu6 = RibTab_Tex_Dokumt
        STaSe ShoCut_Texte, RibTab_Tex_Dokumt
        STxFr 'Formatanpassungen Textcontrol
        DoEvents
        STxNe 'Legt eine neues leeres Textdokument an
        DoEvents
        TxCoN.Load FiNam
        DoEvents
        Unload Me
        Exit Sub
    End If
    
End If

Set LiVw1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub FZuru()
On Error GoTo InErr

Dim DaNam As String
Dim AktZa As Integer
Dim DiNam() As String

Set FTeEm = Me.txtEmail
Set FTeFr = Me.txtFraBo
Set FTePr = Me.txtPraxi
Set FTeDo = Me.txtDocNa
Set ChDoc = Me.chkDokum
Set OpFr1 = Me.optFrag1
Set OpFr2 = Me.optFrag2
Set OpFr3 = Me.optFrag3
Set OpFr4 = Me.optFrag4
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set PuBu1 = Me.btnWeite
Set PuBu2 = Me.btnZuruk

If FTeDo.Enabled = False Then FTeDo.Enabled = True
If FTeFr.Text <> vbNullString Then FTeFr.Text = vbNullString
If FTeDo.Text <> vbNullString Then FTeDo.Text = vbNullString

If Rahm2.Visible = True Then
    Rahm2.Visible = False
    Rahm1.Visible = True
    PuBu2.Enabled = False
ElseIf Rahm3.Visible = True Then
    Rahm3.Visible = False
    If ChDoc.Value = xtpChecked Then
        Rahm2.Visible = True
    Else
        Rahm1.Visible = True
        PuBu2.Enabled = False
    End If
ElseIf Rahm4.Visible = True Then
    Rahm4.Visible = False
    If OpFr1.Value = True Then
        Rahm3.Visible = True
    Else
        If ChDoc.Value = xtpChecked Then
            Rahm2.Visible = True
        Else
            Rahm1.Visible = True
        End If
    End If
ElseIf Rahm5.Visible = True Then
    Rahm5.Visible = False
    Rahm4.Visible = True
    PuBu1.Enabled = True
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50841)
TeMai = IniGetOpt("Hilfe", 50842)
TeInh = IniGetOpt("Hilfe", 50843)
TeFus = IniGetOpt("Hilfe", 50844)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub

Private Sub btnWeite_Click()
    FWeit
End Sub
Private Sub btnZuruk_Click()
    FZuru
End Sub
Private Sub chkDokum_Click()
On Error GoTo InErr

Set ChDoc = Me.chkDokum

If ChDoc.Value = xtpChecked Then
    IniSetVal "System", "FraDoc", -1
    GlFDo = True
Else
    IniSetVal "System", "FraDoc", 0
    GlFDo = False
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDocu " & Err.Number
Resume Next

End Sub

Private Sub chkDoUpl_Click()
On Error GoTo InErr

Set ChUpl = Me.chkDoUpl

If ChUpl.Value = xtpChecked Then
    IniSetVal "System", "FraUpl", -1
    GlFUp = True
Else
    IniSetVal "System", "FraUpl", 0
    GlFUp = False
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDocu " & Err.Number
Resume Next

End Sub
Private Sub chkReduz_Click()
On Error GoTo InErr

Set ChReD = Me.chkReduz

If ChReD.Value = xtpChecked Then
    IniSetVal "System", "FraRed", -1
    GlFrS = True
Else
    IniSetVal "System", "FraRed", 0
    GlFrS = False
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRedu " & Err.Number
Resume Next

End Sub

Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FKonf
AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmPublizieren = Nothing
End Sub

Private Sub lstView1_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error Resume Next

Set LiVw1 = Me.lstView1
Set LiIts = LiVw1.ListItems

If LiIts.Count > 0 Then
    For Each LiItm In LiIts
        If LiItm.Checked = True Then
            LiItm.Selected = True
        Else
            LiItm.Selected = False
        End If
    Next LiItm
End If

End Sub
Private Sub lstView1_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error Resume Next

Set LiVw1 = Me.lstView1
Set LiIts = LiVw1.ListItems

If LiIts.Count > 0 Then
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            LiItm.Checked = True
        Else
            LiItm.Checked = False
        End If
    Next LiItm
End If

End Sub
Private Sub txtDocNa_GotFocus()
    Me.txtDocNa.SelStart = 0
    Me.txtDocNa.SelLength = Len(Me.txtDocNa.Text)
End Sub
Private Sub txtEmail_GotFocus()
    Me.txtEmail.SelStart = 0
    Me.txtEmail.SelLength = Len(Me.txtEmail.Text)
End Sub

Private Sub txtFraBo_GotFocus()
    Me.txtFraBo.SelStart = 0
    Me.txtFraBo.SelLength = Len(Me.txtFraBo.Text)
End Sub

Private Sub txtPraxi_GotFocus()
    Me.txtPraxi.SelStart = 0
    Me.txtPraxi.SelLength = Len(Me.txtPraxi.Text)
End Sub
