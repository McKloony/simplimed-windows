VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmMailKont 
   Caption         =   "Emailkonten"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows-Standard
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1455
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      _Version        =   1048579
      _ExtentX        =   5106
      _ExtentY        =   2566
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   2500
      Left            =   400
      TabIndex        =   2
      Top             =   1800
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   4410
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtEmAbs 
         Height          =   350
         Left            =   1800
         TabIndex        =   26
         Top             =   1100
         Width           =   2895
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtEmBet 
         Height          =   350
         Left            =   1800
         TabIndex        =   25
         Top             =   700
         Width           =   2895
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cbmEmGrp 
         Height          =   315
         Left            =   1800
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1500
         Width           =   2895
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtFiNam 
         Height          =   350
         Left            =   1800
         TabIndex        =   24
         Top             =   300
         Width           =   2895
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbMita2 
         Height          =   315
         Left            =   1800
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1900
         Width           =   2895
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin VB.Label lblLab07 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   255
         Left            =   420
         TabIndex        =   40
         Top             =   1960
         Width           =   1300
      End
      Begin VB.Label lblLab06 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Filtername :"
         Height          =   210
         Left            =   420
         TabIndex        =   39
         Top             =   340
         Width           =   1300
      End
      Begin VB.Label lblLab05 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Adresse :"
         Height          =   210
         Left            =   420
         TabIndex        =   38
         Top             =   1160
         Width           =   1300
      End
      Begin VB.Label lblLab04 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Betreff :"
         Height          =   210
         Left            =   420
         TabIndex        =   37
         Top             =   740
         Width           =   1300
      End
      Begin VB.Label lblLab03 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Gruppe :"
         Height          =   255
         Left            =   420
         TabIndex        =   36
         Top             =   1560
         Width           =   1300
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4100
      Left            =   400
      TabIndex        =   1
      Top             =   4700
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkAnony 
         Height          =   250
         Left            =   4920
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "Service E-Mail-Adresse"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkEmMas 
         Height          =   250
         Left            =   4920
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1600
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "E-Mail-Standardkonto"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkAdChk 
         Height          =   250
         Left            =   4920
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2800
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "Als gelesen kennzeichnen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAlter 
         Height          =   250
         Left            =   4920
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3200
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "Nur neue E-Mails abrufen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtESMTP 
         Height          =   350
         Left            =   1800
         TabIndex        =   8
         Top             =   1100
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5115
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtEMPOP 
         Height          =   350
         Left            =   1800
         TabIndex        =   5
         Top             =   700
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5115
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtEmAdr 
         Height          =   350
         Left            =   1800
         TabIndex        =   11
         Top             =   1500
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5115
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtAnAdr 
         Height          =   350
         Left            =   1800
         TabIndex        =   12
         Top             =   1900
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5115
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtName 
         Height          =   350
         Left            =   1800
         TabIndex        =   13
         Top             =   2300
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5115
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtUsNam 
         Height          =   350
         Left            =   1800
         TabIndex        =   14
         Top             =   2700
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5115
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtUsPas 
         Height          =   350
         Left            =   1800
         TabIndex        =   15
         Top             =   3090
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5115
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         PasswordChar    =   "*"
      End
      Begin XtremeSuiteControls.CheckBox chkEmAut 
         Height          =   250
         Left            =   4920
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2000
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "Erfordert Authentifizierung"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtPort1 
         Height          =   350
         Left            =   6400
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   700
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtPort2 
         Height          =   350
         Left            =   6400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1100
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   310
         Left            =   6400
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1217
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.CheckBox chkKopie 
         Height          =   250
         Left            =   4920
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2400
         _Version        =   1048579
         _ExtentX        =   4233
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "E-Mails auf Server belassen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbProvi 
         Height          =   310
         Left            =   1800
         TabIndex        =   3
         Top             =   300
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
         DropDownItemCount=   15
      End
      Begin XtremeSuiteControls.ComboBox cmbPostf 
         Height          =   310
         Left            =   4920
         TabIndex        =   4
         Top             =   300
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbVer01 
         Height          =   310
         Left            =   4920
         TabIndex        =   6
         Top             =   700
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbVer02 
         Height          =   310
         Left            =   4920
         TabIndex        =   9
         Top             =   1100
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbAlter 
         Height          =   310
         Left            =   1800
         TabIndex        =   16
         Top             =   3500
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab30 
         Height          =   210
         Left            =   420
         TabIndex        =   42
         Top             =   760
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "IMAP-Server :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   210
         Left            =   420
         TabIndex        =   41
         Top             =   360
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Provider :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Abrufalter :"
         Height          =   255
         Left            =   420
         TabIndex        =   35
         Top             =   3560
         Width           =   1300
      End
      Begin VB.Label lblLab01 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Passwort :"
         Height          =   255
         Left            =   420
         TabIndex        =   34
         Top             =   3160
         Width           =   1300
      End
      Begin VB.Label lblLab32 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Adresse :"
         Height          =   210
         Left            =   420
         TabIndex        =   33
         Top             =   1560
         Width           =   1300
      End
      Begin VB.Label lblLab34 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Absendename :"
         Height          =   210
         Left            =   420
         TabIndex        =   32
         Top             =   2360
         Width           =   1300
      End
      Begin VB.Label lblLab31 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP-Server :"
         Height          =   210
         Left            =   420
         TabIndex        =   31
         Top             =   1160
         Width           =   1300
      End
      Begin VB.Label lblLab33 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Antwortadresse :"
         Height          =   210
         Left            =   420
         TabIndex        =   30
         Top             =   1960
         Width           =   1300
      End
      Begin VB.Label lblLab35 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Benutzername :"
         Height          =   255
         Left            =   420
         TabIndex        =   29
         Top             =   2760
         Width           =   1300
      End
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   600
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMailKont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private FS As Form
Private AktCo As VB.Control
Private Lab30 As XtremeSuiteControls.Label
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmMi2 As XtremeSuiteControls.ComboBox
Private CmGrp As XtremeSuiteControls.ComboBox
Private CmPrv As XtremeSuiteControls.ComboBox
Private CmPsf As XtremeSuiteControls.ComboBox
Private CmVe1 As XtremeSuiteControls.ComboBox
Private CmVe2 As XtremeSuiteControls.ComboBox
Private CmAlt As XtremeSuiteControls.ComboBox
Private ChAut As XtremeSuiteControls.CheckBox
Private ChKop As XtremeSuiteControls.CheckBox
Private ChAlt As XtremeSuiteControls.CheckBox
Private ChChk As XtremeSuiteControls.CheckBox
Private ChMas As XtremeSuiteControls.CheckBox
Private ChAno As XtremeSuiteControls.CheckBox
Private TxPo1 As XtremeSuiteControls.FlatEdit
Private TxPo2 As XtremeSuiteControls.FlatEdit
Private TxSv1 As XtremeSuiteControls.FlatEdit
Private TxSv2 As XtremeSuiteControls.FlatEdit
Private TxUsN As XtremeSuiteControls.FlatEdit
Private TxUsP As XtremeSuiteControls.FlatEdit
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpGrw As XtremeReportControl.ReportGroupRow
Private ChRow As XtremeReportControl.ReportRow
Private ChRws As XtremeReportControl.ReportRows
Private WithEvents TbBar As XtremeCommandBars.TabToolBar
Attribute TbBar.VB_VarHelpID = -1
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const KEYEVENTF_KEYUP = &H2
Private Const CB_FINDSTRING = &H14C&
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157

Private LiTyp As Integer
Private KtNeu As Boolean
Private KtKop As Boolean

Private clFen As clsFenster

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Function FChek() As Boolean
On Error GoTo InErr

Dim MitNr As Long
Dim AktZa As Integer
Dim GesZa As Integer
Dim MaVor As Boolean

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMailKont
Set CmBrs = FM.comBar02

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

If GlMiV = True Then
    MitNr = CmMit.ItemData(CmMit.ListIndex)
Else
    MitNr = 0
End If

If GlEKV = False Then 'Emailkonten vorhanden
    For AktZa = 1 To GesZa
        If CLng(GlMkt(AktZa, 1)) = MitNr Then
            If CBool(GlMkt(AktZa, 20)) = True Then 'Standardemailkonto
                MaVor = True
                Exit For
            End If
        End If
    Next AktZa
End If

FChek = MaVor

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FChek " & Err.Number
Resume Next

End Function

Private Sub FClip(ByVal KoKop As Boolean)
On Error GoTo InErr

Dim MiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMailKont
Set CmBrs = FM.comBar02
Set CmPrv = FM.cmbProvi
Set CmPsf = FM.cmbPostf
Set CmVe1 = FM.cmbVer01
Set CmVe2 = FM.cmbVer02
Set CmAlt = FM.cmbAlter
Set TxSv1 = FM.txtEMPOP
Set TxSv2 = FM.txtESMTP
Set TxPo1 = FM.txtPort1
Set TxPo2 = FM.txtPort2
Set TxUsN = FM.txtUsNam
Set TxUsP = FM.txtUsPas
Set ChAut = FM.chkEmAut
Set ChMas = FM.chkEmMas
Set ChKop = FM.chkKopie
Set ChAno = FM.chkAnony
Set CmMit = FM.cmbMitar
Set CmMi2 = FM.cmbMita2
Set ChAlt = FM.chkAlter
Set ChChk = FM.chkAdChk
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpRws = RpCon.Rows
Set RpSel = RpCon.SelectedRows

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

MiIdx = CmCom.ListIndex

If KoKop = True Then
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            ReDim GlClp(1, 19)
            GlClp(1, 0) = CmPrv.ListIndex
            GlClp(1, 1) = CmAlt.ListIndex
            GlClp(1, 2) = CmPsf.ListIndex
            GlClp(1, 3) = CmVe1.ListIndex
            GlClp(1, 4) = CmVe2.ListIndex
            GlClp(1, 5) = TxSv1.Text
            GlClp(1, 6) = TxSv2.Text
            GlClp(1, 7) = TxPo1.Text
            GlClp(1, 8) = TxPo2.Text
            GlClp(1, 9) = TxUsN.Text
            GlClp(1, 10) = TxUsP.Text
            GlClp(1, 11) = FM.txtEmAdr.Text
            GlClp(1, 12) = FM.txtAnAdr.Text
            GlClp(1, 13) = FM.txtName.Text
            GlClp(1, 14) = ChAut.Value
            GlClp(1, 15) = ChKop.Value
            GlClp(1, 16) = ChAlt.Value
            GlClp(1, 17) = ChChk.Value
            GlClp(1, 18) = ChMas.Value
            GlClp(1, 19) = ChAno.Value
            KtKop = True
        End If
    End If
Else
    If KtKop = True Then
        If UBound(GlClp) > 0 Then
            CmPrv.ListIndex = GlClp(1, 0)
            CmAlt.ListIndex = GlClp(1, 1)
            CmPsf.ListIndex = GlClp(1, 2)
            CmVe1.ListIndex = GlClp(1, 3)
            CmVe2.ListIndex = GlClp(1, 4)
            TxSv1.Text = GlClp(1, 5)
            TxSv2.Text = GlClp(1, 6)
            TxPo1.Text = GlClp(1, 7)
            TxPo2.Text = GlClp(1, 8)
            TxUsN.Text = GlClp(1, 9)
            If GlNoM = False Then 'kein Meisterp...
                TxUsP.Text = GlClp(1, 10)
            End If
            FM.txtEmAdr.Text = GlClp(1, 11)
            FM.txtAnAdr.Text = GlClp(1, 12)
            FM.txtName.Text = GlClp(1, 13)
            ChAut.Value = GlClp(1, 14)
            ChKop.Value = GlClp(1, 15)
            ChAlt.Value = GlClp(1, 16)
            ChChk.Value = GlClp(1, 17)
            ChAno.Value = GlClp(1, 19)
            If RpRws.Count = 0 Then
                ChMas.Value = xtpChecked
            Else
                ChMas.Value = xtpUnchecked
            End If
            CmMit.ListIndex = MiIdx - 1
            Erase GlClp
            ClKop = False
            DoEvents
            S_MaKSa True, LiTyp
            S_MaKLa 0, LiTyp
            S_MaKPo LiTyp
            DoEvents
            S_Ary2
        End If
        KtKop = False
    End If
DoEvents
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClip " & Err.Number
Resume Next

End Sub

Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50931)
TeMai = IniGetOpt("Hilfe", 50932)
TeInh = IniGetOpt("Hilfe", 50933)
TeFus = IniGetOpt("Hilfe", 50934)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FKonf()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim AktZa As Integer
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMailKont
Set CmMit = FM.cmbMitar
Set CmMi2 = FM.cmbMita2
Set CmGrp = FM.cbmEmGrp
Set CmPrv = FM.cmbProvi
Set CmPsf = FM.cmbPostf
Set CmVe1 = FM.cmbVer01
Set CmVe2 = FM.cmbVer02
Set CmAlt = FM.cmbAlter
Set TxPo1 = FM.txtPort1
Set TxPo2 = FM.txtPort2
Set TxSv1 = FM.txtEMPOP
Set TxSv2 = FM.txtESMTP
Set ChAut = FM.chkEmAut
Set ChKop = FM.chkKopie
Set ChAlt = FM.chkAlter
Set ChChk = FM.chkAdChk
Set ChMas = FM.chkEmMas
Set ChAno = FM.chkAnony
Set RpCon = FM.repCont1
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set ImMan = frmMain.imgManag

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = True
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = vbBlack
    .PaintManager.MaxPreviewLines = 5
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.FixedRowHeight = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With CmMit
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        .AddItem GlMiA(AktZa, 1)
        .ItemData(AktZa - 1) = GlMiA(AktZa, 2)
    Next AktZa
    .ListIndex = GlSmI - 1
End With

With CmMi2
    For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
        .AddItem GlMiK(AktZa, 1)
        .ItemData(AktZa - 1) = GlMiK(AktZa, 2)
    Next AktZa
    .ListIndex = GlSmI - 1
End With

With CmPrv
    For AktZa = 1 To UBound(GlPro) 'Emailprovider
        .AddItem GlPro(AktZa, 0)
        .ItemData(AktZa - 1) = AktZa
    Next AktZa
    .ListIndex = 14
End With

With CmPsf
    .AddItem "IMAP4"
    .ItemData(0) = 1
    .AddItem "POP3"
    .ItemData(1) = 2
    .ListIndex = 0
End With

With CmVe1
    .AddItem "SSLAuto"
    .ItemData(0) = 0
    .AddItem "SSL"
    .ItemData(1) = 1
    .AddItem "TLS"
    .ItemData(2) = 2
End With

With CmVe2
    .AddItem "DirectSSL"
    .ItemData(0) = 0
    .AddItem "STARTTLS"
    .ItemData(1) = 1
    .AddItem "TryTLS"
    .ItemData(2) = 2
End With

With CmAlt
    .AddItem "Unbeschränkt"
    .ItemData(0) = 0
    .AddItem "07 Tage"
    .AddItem "30 Tage"
    .ItemData(1) = 1
    .AddItem "60 Tage"
    .ItemData(2) = 2
    .AddItem "90 Tage"
    .ItemData(3) = 3
    .ListIndex = 0
End With

With CmGrp
    For AktZa = 1 To UBound(GlEmG)
        .AddItem GlEmG(AktZa, 1)
        .ItemData(AktZa - 1) = GlEmG(AktZa, 0)
    Next AktZa
    .AddItem "Rechnungen"
    .ItemData(AktZa - 1) = 802
    .AddItem "Mitarbeiter"
    .ItemData(AktZa) = 803
    .AddItem "Telefaxe"
    .ItemData(AktZa + 1) = 804
    .AddItem "Junkmail"
    .ItemData(AktZa + 2) = 805
    .ListIndex = 0
End With

With TxPo1
    .MaxLength = 3
    .Pattern = "\d*"
    .SetMask "000", "___"
End With

With TxPo2
    .MaxLength = 3
    .Pattern = "\d*"
    .SetMask "000", "___"
End With

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
ChAut.BackColor = GlBak
ChKop.BackColor = GlBak
ChAlt.BackColor = GlBak
ChChk.BackColor = GlBak
ChMas.BackColor = GlBak
ChAno.BackColor = GlBak

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FMenu()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim RetWe As Long
Dim KeyNa As String
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim ToTab As XtremeCommandBars.TabControlItem

Set FM = frmMailKont
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmAcs = CmBrs.Actions
Set CmOpt = CmBrs.Options
Set ImMan = frmMain.imgManag

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmSta
    .Font.SIZE = 8
    Set CmPan = .AddPane(1)
    CmPan.Text = Date
    CmPan.Width = 100
    CmPan.Alignment = xtpAlignmentCenter
    Set CmPan = .AddPane(2)
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    .Visible = True
End With

With CmAcs
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_OP_Uebernahme, vbNullString, vbNullString, vbNullString, vbNullString)
End With

Set TbBar = CmBrs.AddTabToolBar("TabBar")

'______________________________________________________________________

Set ToTab = TbBar.InsertCategory(RibTab_Wart_Wied, "Emailkonten")
With ToTab
    .ToolTip = "Bearbeiten der einzelnen Emailkonten"
    .Visible = True
    .Selected = True
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Emailkonten"
        .ToolTipText = "Legt ein neues Emailkonto an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Emailkonten"
        .ToolTipText = "Speichert die Änderungen"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Emailkonten"
        .ToolTipText = "Löscht die markierten Emailkonten"
        .BeginGroup = True
        .IconId = IC24_Doc_Del
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Kopieren, "Kopieren")
    With CmCon
        .Category = "Emailkonten"
        .ToolTipText = "Kopiert das markierte Emailkonto in die Zwischenablage"
        .BeginGroup = True
        .IconId = IC24_Copy
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Einfuegen, "Einfügen")
    With CmCon
        .Category = "Emailkonten"
        .ToolTipText = "Fügt das kopierte Emailkonto aus der Zwischenablage ein"
        .BeginGroup = True
        .IconId = IC24_Paste
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Emailkonten"
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
    End With
End With

'______________________________________________________________________

Set ToTab = TbBar.InsertCategory(RibTab_Wart_Beha, "Emailfilter")
With ToTab
    .ToolTip = "Bearbeiten von Emailfiltern und Spammeil"
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Emailfilter"
        .ToolTipText = "Legt einen neuen Filter an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Emailfilter"
        .ToolTipText = "Speichert die Änderungen"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Emailfilter"
        .ToolTipText = "Löscht die markierten Emailfilter"
        .BeginGroup = True
        .IconId = IC24_Doc_Del
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Emailfilter"
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
    End With
End With

'______________________________________________________________________

For Each CmBar In CmBrs
    If CmBar.Type = xtpBarTypeNormal Then
        Set CmCoS = CmBar.Controls
        For Each CmCon In CmCoS
            CmCon.Style = xtpButtonIconAndCaption
        Next CmCon
    End If
Next CmBar

'______________________________________________________________________

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Cap01, "Mitarbeiter :")
    With CmCon
        .ToolTipText = "Wählen Sie den gewünschten Mitarbeiter"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With
    Set CmCom = .Add(xtpControlComboBox, KA_SuCo1, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Wählen Sie den gewünschten Mitarbeiter"
        .ThemedItems = True
        .Width = 140
        For AktZa = 1 To UBound(GlMiA)
            .AddItem GlMiA(AktZa, 1)
            .ItemData(AktZa) = GlMiA(AktZa, 2)
        Next AktZa
        .ListIndex = GlSmI
    End With
    Set CmCon = .Add(xtpControlButton, KA_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Style = xtpButtonIconAndCaption
        .flags = xtpFlagRightAlign
        .IconId = IC16_Sign_Help
        .ShortcutText = "F1"
    End With
End With

'______________________________________________________________________

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7: .VisualTheme = xtpThemeOffice2013
    Case 8: .VisualTheme = xtpThemeOffice2013
    Case Else: .VisualTheme = xtpThemeRibbon
    End Select
    If GlSty = 8 Then 'Office 2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    ElseIf GlSty = 7 Then 'Office 2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Else
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End If
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = False
    .ToolTipContext.ShowOfficeBorder = True
    .ToolTipContext.ShowShadow = True
    .ToolTipContext.ShowTitleAndDescription False, xtpToolTipIconNone
    .ToolTipContext.Style = xtpToolTipResource
    .PaintManager.AutoResizeIcons = False
    .PaintManager.ClearTypeTextQuality = GlCle
    .PaintManager.EnableAnimation = GlMeA
    .PaintManager.FlatMenuBar = False
    .PaintManager.FlatToolBar = False
    .PaintManager.SelectImageInPopupBar = True
    .PaintManager.ShowShadow = True
    .PaintManager.ThemedCheckBox = True
    .PaintManager.ThemedStatusBar = True
    .PaintManager.ThickCheckMark = False
    .KeyBindings.Add 0, VK_F1, KY_F1
    .KeyBindings.Add 0, VK_F2, KY_F2
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
End With

With CmOpt
    .AltDragCustomization = False
    .AlwaysShowFullMenus = True
    .AutoHideUnusedPopups = False
    .ExpandDelay = 100
    .ExpandHoverDelay = 100
    .FloatToolbarsByDoubleClick = False
    .IconsWithShadow = False
    .KeyboardCuesShow = xtpKeyboardCuesShowAlways
    .KeyboardCuesUse = xtpKeyboardCuesUseMenuOnly
    .LargeIcons = False
    .LunaColors = GlLun
    .MaxPopupWidth = 0.5
    .OfficeStyleDisabledIcons = True
    .SetIconSize True, 24, 24
    .ShowExpandButtonAlways = False
    .ShowFullAfterDelay = True
    .ShowPopupBarToolTips = False
    .ShowTextBelowIcons = False
    .ShowKeyboardTips = True
    .SyncFloatingToolbars = True
    .ToolBarAccelTips = True
    .ToolBarScreenTips = True
    .UpdatePeriod = 100
    .UseAltNumPadKeys = False
    .UseDisabledIcons = True
    .UseFadedIcons = False
    .UseSharedImageList = False
    .UseSystemSaveBitsStyle = False
    .Animation = xtpAnimateWindowsDefault
    .Font.SIZE = 8
    .ComboBoxFont.SIZE = 8
End With

With TbBar
    .AllowReorder = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableAnimation = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 24, 24
    .ShowExpandButton = False
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .SetIconSize 24, 24
    Select Case GlSty
    Case 8:
        .TabPaintManager.Appearance = xtpTabThemeOffice2013
        .TabPaintManager.Color = xtpTabColorOffice2013
    Case 7:
        .TabPaintManager.Appearance = xtpTabThemeOffice2013
        .TabPaintManager.Color = xtpTabColorOffice2013
    Case Else:
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2007
        .TabPaintManager.Color = xtpTabColorResource
    End Select
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ButtonMargin.Top = 6
    .TabPaintManager.FixedTabWidth = 110
    .TabPaintManager.ButtonMargin.Bottom = 0
    .TabPaintManager.ButtonMargin.Left = 0
    .TabPaintManager.ButtonMargin.Right = 0
    .TabPaintManager.ClientFrame = xtpTabFrameSingleLine
    .TabPaintManager.ClientMargin.Bottom = 0
    .TabPaintManager.ClientMargin.Top = 0
    .TabPaintManager.ClientMargin.Left = 0
    .TabPaintManager.ClientMargin.Right = 0
    .TabPaintManager.ControlMargin.Top = 0
    .TabPaintManager.ControlMargin.Bottom = 0
    .TabPaintManager.ControlMargin.Left = 0
    .TabPaintManager.ControlMargin.Right = 0
    .TabPaintManager.HeaderMargin.Top = 0
    .TabPaintManager.HeaderMargin.Bottom = 0
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.HeaderMargin.Right = 0
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = True
    .TabPaintManager.HotTracking = True
    .TabPaintManager.Layout = xtpTabLayoutFixed
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.Font.SIZE = 8
End With

DoEvents
CmBrs.RecalcLayout
DoEvents
CmBrs.PaintManager.RefreshMetrics

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FNeLa()
On Error GoTo AnErr

Dim MiIdx As Integer
Dim IdPsf As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMailKont
Set CmBrs = FM.comBar02
Set CmPsf = FM.cmbPostf
Set CmAlt = FM.cmbAlter
Set CmMit = FM.cmbMitar
Set CmMi2 = FM.cmbMita2
Set ChAlt = FM.chkAlter
Set ChMas = FM.chkEmMas

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

MiIdx = CmCom.ListIndex

IdPsf = CmPsf.ItemData(CmPsf.ListIndex)

ChMas.Value = xtpUnchecked

Select Case LiTyp
Case 1: FSpla1
        S_MaKLa 0, LiTyp
        S_MaKPo LiTyp
        CmMit.ListIndex = MiIdx - 1
Case 2: FSpla2
        S_MaKLa 0, LiTyp
        S_MaKPo LiTyp
        CmMi2.ListIndex = MiIdx - 1
End Select

If LiTyp = 1 Then
    Select Case IdPsf
    Case 1: ChAlt.Enabled = True
            If ChAlt.Value = xtpChecked Then
                CmAlt.Enabled = True
            Else
                CmAlt.Enabled = False
            End If
    Case 2: CmAlt.Enabled = False
            ChAlt.Enabled = False
    End Select
End If

DoEvents
FPosi

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeLa " & Err.Number
Resume Next

End Sub
Private Sub FPrt1()
On Error GoTo InErr

Dim IdPrv As Integer
Dim IdPsf As Integer
Dim IdPo1 As Integer

Set FM = frmMailKont
Set Lab30 = FM.lblLab30
Set CmPrv = FM.cmbProvi
Set CmPsf = FM.cmbPostf
Set CmVe1 = FM.cmbVer01
Set CmAlt = FM.cmbAlter
Set TxPo1 = FM.txtPort1
Set TxSv1 = FM.txtEMPOP
Set ChAlt = FM.chkAlter
Set ChChk = FM.chkAdChk

IdPrv = CmPrv.ListIndex + 1
IdPsf = CmPsf.ItemData(CmPsf.ListIndex)
IdPo1 = CmVe1.ItemData(CmVe1.ListIndex)

Select Case IdPsf
Case 1: Lab30.Caption = "IMAP-Server :"
        If GlPro(IdPrv, 2) <> vbNullString Then
            TxSv1.Text = GlPro(IdPrv, 2)
        End If
        Select Case IdPo1
        Case 0: TxPo1.Text = 143
        Case 1: TxPo1.Text = 993
        Case 2: TxPo1.Text = 143
        End Select
        ChAlt.Enabled = True
        ChChk.Enabled = True
        If ChAlt.Value = xtpChecked Then
            CmAlt.Enabled = True
        Else
            CmAlt.Enabled = False
        End If
Case 2: Lab30.Caption = "POP3-Server :"
        If GlPro(IdPrv, 2) <> vbNullString Then
            TxSv1.Text = GlPro(IdPrv, 1)
        End If
        Select Case IdPo1
        Case 0: TxPo1.Text = 110
        Case 1: TxPo1.Text = 995
        Case 2: TxPo1.Text = 110
        End Select
        CmAlt.Enabled = False
        ChAlt.Enabled = False
        ChChk.Enabled = False
        ChChk.Value = xtpUnchecked
End Select

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPrt1 " & Err.Number
Resume Next

End Sub
Private Sub FPrt2()
On Error GoTo InErr

Dim IdPrv As Integer
Dim IdPo2 As Integer

Set FM = frmMailKont
Set CmPrv = FM.cmbProvi
Set CmVe2 = FM.cmbVer02
Set TxPo2 = FM.txtPort2
Set TxSv2 = FM.txtESMTP

IdPrv = CmPrv.ListIndex + 1
IdPo2 = CmVe2.ItemData(CmVe2.ListIndex)

If GlPro(IdPrv, 3) <> vbNullString Then
    TxSv2.Text = GlPro(IdPrv, 3)
End If

Select Case IdPo2
Case 0: TxPo2.Text = 465
Case 1: TxPo2.Text = 465
Case 2: TxPo2.Text = 587
Case 3: TxPo2.Text = 25
End Select

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPrt2 " & Err.Number
Resume Next

End Sub
Private Sub FSave()
On Error GoTo InErr

Dim TmStr As String
Dim RowNr As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMailKont
Set TxUsN = FM.txtUsNam
Set TxUsP = FM.txtUsPas
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        RowNr = RpRow.Index
    End If
End If

If TxUsN.Text <> vbNullString Then
    TmStr = TxUsN.Text
    If InStr(1, TmStr, ";", 1) > 0 Then
        TxUsN.Text = Replace(TmStr, ";", vbNullString, 1)
        SPopu "Ungültiges Sonderzeichen", "Der Benurtzername darf kein Semikolon enthlaten.", IC48_Forbidden
    End If
End If

If TxUsP.Text <> vbNullString Then
    TmStr = TxUsP.Text
    If InStr(1, TmStr, ";", 1) > 0 Then
        TxUsP.Text = Replace(TmStr, ";", vbNullString, 1)
        SPopu "Ungültiges Sonderzeichen", "Das Passwort darf kein Semikolon enthlaten.", IC48_Forbidden
    End If
End If
DoEvents

S_MaKSa KtNeu, LiTyp
DoEvents

KtNeu = False

If KtNeu = True Then
    S_MaKLa 0, LiTyp
Else
    S_MaKLa RowNr, LiTyp
End If
DoEvents

S_MaKPo LiTyp
DoEvents

S_Ary2

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSett()
On Error GoTo InErr

Dim IdPrv As Integer
Dim IdPsf As Integer
Dim IdPo1 As Integer
Dim IdPo2 As Integer

Set FM = frmMailKont
Set Lab30 = FM.lblLab30
Set CmPrv = FM.cmbProvi
Set CmPsf = FM.cmbPostf
Set CmVe1 = FM.cmbVer01
Set CmVe2 = FM.cmbVer02
Set ChAut = FM.chkEmAut
Set ChKop = FM.chkKopie
Set ChAno = FM.chkAnony
Set TxPo1 = FM.txtPort1
Set TxPo2 = FM.txtPort2
Set TxSv1 = FM.txtEMPOP
Set TxSv2 = FM.txtESMTP

IdPrv = CmPrv.ListIndex + 1
CmVe1.ListIndex = GlPro(IdPrv, 4)
CmVe2.ListIndex = GlPro(IdPrv, 5)

IdPsf = CmPsf.ItemData(CmPsf.ListIndex)
IdPo1 = CmVe1.ItemData(CmVe1.ListIndex)
IdPo2 = CmVe2.ItemData(CmVe2.ListIndex)

Select Case IdPsf
Case 1: Lab30.Caption = "IMAP-Server :"
        If GlPro(IdPrv, 2) <> vbNullString Then
            TxSv1.Text = GlPro(IdPrv, 2)
        Else
            TxSv1.Text = vbNullString
        End If
        Select Case IdPo1
        Case 0: TxPo1.Text = 143
        Case 1: TxPo1.Text = 993
        Case 2: TxPo1.Text = 143
        End Select
Case 2: Lab30.Caption = "POP3-Server :"
        If GlPro(IdPrv, 1) <> vbNullString Then
            TxSv1.Text = GlPro(IdPrv, 1)
        Else
            TxSv1.Text = vbNullString
        End If
        Select Case IdPo1
        Case 0: TxPo1.Text = 110
        Case 1: TxPo1.Text = 995
        Case 2: TxPo1.Text = 110
        End Select
End Select

If GlPro(IdPrv, 3) <> vbNullString Then
    TxSv2.Text = GlPro(IdPrv, 3)
Else
    TxSv2.Text = vbNullString
End If

Select Case IdPo2
Case 0: TxPo2.Text = 465
Case 1: TxPo2.Text = 465
Case 2: TxPo2.Text = 587
Case 3: TxPo2.Text = 25
End Select

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMailKont
Set RpCon = FM.repCont1
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set CmBrs = FM.comBar02

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    Select Case LiTyp
    Case 1:
        RpCon.Move 0, ClObn, ClBre, ClHoh - 5380
        Rahm1.Move 60, ClHoh - 4160, ClBre - 120
    Case 2:
        RpCon.Move 0, ClObn, ClBre, ClHoh - 3800
        Rahm2.Move 60, ClHoh - 2560, ClBre - 120
    End Select
End If

Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FNeu()
On Error GoTo InErr

Dim MiIdx As Integer
Dim IdPsf As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMailKont
Set CmBrs = FM.comBar02
Set CmPrv = FM.cmbProvi
Set CmPsf = FM.cmbPostf
Set CmAlt = FM.cmbAlter
Set TxSv1 = FM.txtEMPOP
Set TxSv2 = FM.txtESMTP
Set TxUsN = FM.txtUsNam
Set TxUsP = FM.txtUsPas
Set ChAut = FM.chkEmAut
Set ChMas = FM.chkEmMas
Set ChAno = FM.chkAnony
Set CmMit = FM.cmbMitar
Set CmMi2 = FM.cmbMita2
Set ChAlt = FM.chkAlter
Set RpCon = FM.repCont1
Set RpRws = RpCon.Rows

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

MiIdx = CmCom.ListIndex

IdPsf = CmPsf.ItemData(CmPsf.ListIndex)

Select Case LiTyp
Case 1:
    CmPrv.ListIndex = 14 'WICHTIG!
    CmAlt.ListIndex = 0
    FSett
    TxSv1.Text = vbNullString
    TxSv2.Text = vbNullString
    TxUsN.Text = vbNullString
    TxUsP.Text = vbNullString
    FM.txtEmAdr.Text = vbNullString
    FM.txtAnAdr.Text = vbNullString
    FM.txtName.Text = vbNullString
    ChAut.Value = xtpChecked
    CmMit.ListIndex = MiIdx - 1
Case 2:
    FM.txtFiNam.Text = "Neuer Emailfilter"
    FM.txtEmBet.Text = vbNullString
    FM.txtEmAbs.Text = vbNullString
    FM.cbmEmGrp.ListIndex = 0
    CmMi2.ListIndex = MiIdx - 1
End Select

Select Case IdPsf
Case 1: ChAlt.Enabled = True
        If ChAlt.Value = xtpChecked Then
            CmAlt.Enabled = True
        Else
            CmAlt.Enabled = False
        End If
Case 2: CmAlt.Enabled = False
        ChAlt.Enabled = False
End Select

If RpRws.Count = 0 Then
    ChMas.Value = xtpChecked
Else
    ChMas.Value = xtpUnchecked
End If

Set RpCon = Nothing
Set RpRws = Nothing

KtNeu = True

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeu " & Err.Number
Resume Next

End Sub
Private Sub FOpn()
On Error GoTo InErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMailKont
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

With clFen
    If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
        .FeLin = (GlxGr - GlFeB) / 2
        .FeObn = (GlyGr - GlFeH) / 2
        .FeBre = 572
        .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
    Else
        .FeLin = 80
        .FeObn = 10
        .FeBre = 572
        .FeHoh = GlyGr - 40
    End If
    .FenMov
End With

S_MaKSt 1
DoEvents
S_MaKPo 1
DoEvents

Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpn " & Err.Number
Resume Next

End Sub
Private Sub FSpla1()
On Error GoTo InErr

Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMailKont
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(EMK_IDK, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_IDM, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_POP, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_SMTP, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_SeNam, "Emailkonto", 10, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = True
        .AutoSize = True
    End With
    If RpCon.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(EMK_Em_Adresse, vbNullString, 300, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_Repl, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_User, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_Pass, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_Port1, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_Port2, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Em_Aut, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(EMK_Selekt, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpla1 " & Err.Number
Resume Next

End Sub
Private Sub FSpla2()
On Error GoTo InErr

Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMailKont
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(SPM_IDA, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(SPM_ID1, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(SPM_Subject, "Emailbetreff", 200, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(SPM_SenderMail, "Emailadresse", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(SPM_Selekt, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(SPM_TreKey, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(SPM_IDKurz, "Filtername", 10, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = True
        .AutoSize = True
    End With
    If RpCon.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(SPM_IDM, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpla2 " & Err.Number
Resume Next

End Sub
Private Sub FTabu(ByVal TaIdx As Long)
On Error GoTo AnErr

Dim MiIdx As Integer
Dim IdPsf As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMailKont
Set CmBrs = FM.comBar02
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set CmPsf = FM.cmbPostf
Set CmMit = FM.cmbMitar
Set CmMi2 = FM.cmbMita2

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

MiIdx = CmCom.ListIndex

IdPsf = CmPsf.ItemData(CmPsf.ListIndex)
    
LiTyp = TaIdx + 1

Select Case TaIdx
Case 0: FSpla1
        S_MaKLa 0, LiTyp
        S_MaKPo LiTyp
        Rahm2.Visible = False
        Rahm1.Visible = True
        CmMit.ListIndex = MiIdx - 1
Case 1: FSpla2
        S_MaKLa 0, LiTyp
        S_MaKPo LiTyp
        Rahm2.Visible = True
        Rahm1.Visible = False
        CmMi2.ListIndex = MiIdx - 1
End Select

DoEvents
FPosi

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

GlToo = True

Set FM = frmMailKont

Select Case TolId
Case KA_SuCo1: FNeLa
Case KA_SuCo2: FNeLa
Case KA_Hilfe: FHilfe
Case KY_F1: FHilfe
Case KY_F3: FNeu
Case KY_F8: FSave
Case KY_F11: Unload FM
Case SY_OP_Hinzufuegen: FNeu
Case SY_OP_Speichern: FSave
Case SY_OP_Loeschen: S_MaKLo LiTyp
Case SY_OP_Kopieren: FClip True
Case SY_OP_Einfuegen: FClip False
Case SY_OP_Abbruch: Unload FM
End Select

GlToo = False

End Sub

Private Sub chkAlter_Click()
On Error Resume Next

Set FM = frmMailKont
Set ChAlt = FM.chkAlter
Set CmAlt = FM.cmbAlter

If GlAkK = False Then
    If ChAlt.Value = xtpChecked Then
        CmAlt.Enabled = True
    Else
        CmAlt.Enabled = False
    End If
End If

End Sub
Private Sub chkEmAut_Click()
On Error Resume Next

Set FM = frmMailKont
Set ChAut = FM.chkEmAut
Set TxUsN = FM.txtUsNam
Set TxUsP = FM.txtUsPas

If GlAkK = False Then
    If ChAut.Value = xtpChecked Then
        TxUsN.Enabled = True
        TxUsP.Enabled = True
    Else
        TxUsN.Enabled = False
        TxUsP.Enabled = False
    End If
End If

End Sub

Private Sub chkEmMas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Set FM = frmMailKont
Set ChMas = FM.chkEmMas

If ChMas.Value = xtpChecked Then
    If FChek = True Then
        ChMas.Value = xtpUnchecked
        SPopu "E-Mail-Standardkonto", "Es darf nur ein E-Mail-Standardkonto geben.", IC48_Forbidden
    End If
End If

End Sub
Private Sub chkKopie_Click()
On Error Resume Next

Set FM = frmMailKont
Set ChKop = FM.chkKopie
Set ChChk = FM.chkAdChk

If GlAkK = False Then
    If ChKop.Value = xtpChecked Then 'Emails auf Server belassen
        ChChk.Enabled = True
    Else
        ChChk.Enabled = False 'Als gelesen kennzeichnen
        ChChk.Value = xtpUnchecked
    End If
End If

End Sub

Private Sub cmbPostf_Click()
    If GlAkK = False Then
        FPrt1
    End If
End Sub
Private Sub cmbProvi_Click()
    If GlAkK = False Then
        FSett
    End If
End Sub

Private Sub cmbVer01_Click()
    If GlAkK = False Then
        FPrt1
    End If
End Sub
Private Sub cmbVer02_Click()
    If GlAkK = False Then
        FPrt2
    End If
End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
Private Sub comBar02_Resize()
    If GlAkK = False Then
        FPosi
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

GlAkK = True
Screen.MousePointer = vbHourglass
DoEvents

With FrmEx
    .ClientMaxHeight = 14000
    .ClientMaxWidth = 11000
    .ClientMinHeight = 8000
    .ClientMinWidth = 8100
End With

LiTyp = 1

AFont Me
FMenu
FSpla1
FPosi
FKonf
FOpn

If GlRah = True Then
    SFrame 1, Me.hwnd
End If

Set FrmEx = Nothing

DoEvents
Screen.MousePointer = vbNormal
GlAkK = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmMailKont = Nothing
End Sub

Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        S_MaKPo LiTyp
    End If
End Sub

Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    S_MaKPo LiTyp
End Sub
Private Sub repCont1_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    S_MaKPo LiTyp
End Sub

Private Sub txtAnAdr_GotFocus()
    Me.txtAnAdr.SelStart = 0
    Me.txtAnAdr.SelLength = Len(Me.txtAnAdr.Text)
End Sub
Private Sub txtEmAdr_GotFocus()
    Me.txtEmAdr.SelStart = 0
    Me.txtEmAdr.SelLength = Len(Me.txtEmAdr.Text)
End Sub
Private Sub txtEMPOP_GotFocus()
    Me.txtEMPOP.SelStart = 0
    Me.txtEMPOP.SelLength = Len(Me.txtEMPOP.Text)
End Sub


Private Sub txtESMTP_GotFocus()
    Me.txtESMTP.SelStart = 0
    Me.txtESMTP.SelLength = Len(Me.txtESMTP.Text)
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0
    Me.txtName.SelLength = Len(Me.txtName.Text)
End Sub

Private Sub txtPort1_GotFocus()
    Me.txtPort1.SelStart = 0
    Me.txtPort1.SelLength = Len(Me.txtPort1.Text)
End Sub

Private Sub txtPort2_GotFocus()
    Me.txtPort2.SelStart = 0
    Me.txtPort2.SelLength = Len(Me.txtPort2.Text)
End Sub
Private Sub txtUsNam_GotFocus()
    Me.txtUsNam.SelStart = 0
    Me.txtUsNam.SelLength = Len(Me.txtUsNam.Text)
End Sub

Private Sub txtUsNam_LostFocus()
On Error Resume Next

Dim TmStr As String

Set TxUsN = Me.txtUsNam

If TxUsN.Text <> vbNullString Then
    TmStr = TxUsN.Text
    If InStr(1, TmStr, ";", 1) > 0 Then
        TxUsN.Text = Replace(TmStr, ";", vbNullString, 1)
        SPopu "Ungültiges Sonderzeichen", "Der Benutzername darf kein Semikolon enthlaten.", IC48_Forbidden
    End If
End If

End Sub


Private Sub txtUsPas_GotFocus()
    Me.txtUsPas.SelStart = 0
    Me.txtUsPas.SelLength = Len(Me.txtUsPas.Text)
End Sub
Private Sub TbBar_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    If GlAkK = False Then
        FTabu Item.Index
    End If
End Sub

Private Sub txtUsPas_LostFocus()
On Error Resume Next

Dim TmStr As String

Set TxUsP = Me.txtUsPas

If TxUsP.Text <> vbNullString Then
    TmStr = TxUsP.Text
    If InStr(1, TmStr, ";", 1) > 0 Then
        SPopu "Ungültiges Sonderzeichen", "Das Passwort darf kein ';' enthlaten.", IC48_Forbidden
        TxUsP.Text = vbNullString
    ElseIf InStr(1, TmStr, "#", 1) > 0 Then
        SPopu "Ungültiges Sonderzeichen", "Das Passwort darf kein '#' enthlaten.", IC48_Forbidden
        TxUsP.Text = vbNullString
    End If
End If

End Sub
