VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmOutlook 
   Caption         =   "Outlookabgleich"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   ControlBox      =   0   'False
   Icon            =   "frmOutlook.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   10110
   StartUpPosition =   2  'Bildschirmmitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   34
      Top             =   4500
      Width           =   10230
      _Version        =   1048579
      _ExtentX        =   18045
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   8000
         TabIndex        =   38
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
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Height          =   400
         Left            =   6600
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Weiter >"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnZurück 
         Height          =   400
         Left            =   5200
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "< &Zurück"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   3900
         TabIndex        =   35
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
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   4395
      Left            =   6200
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4005
      _Version        =   1048579
      _ExtentX        =   7056
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   2500
         Left            =   200
         TabIndex        =   31
         Top             =   1000
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   4410
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.RadioButton optOpti2 
         Height          =   225
         Left            =   600
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4000
         Width           =   2055
         _Version        =   1048579
         _ExtentX        =   3625
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "ODER-Verknüpfung"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optOpti1 
         Height          =   225
         Left            =   600
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3680
         Width           =   2055
         _Version        =   1048579
         _ExtentX        =   3625
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "UND-Verknüpfung"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4400
      Left            =   6200
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   4000
      _Version        =   1048579
      _ExtentX        =   7056
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   315
         Left            =   2420
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3460
         Width           =   315
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   547
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   315
         Left            =   2420
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2960
         Width           =   315
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   547
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton optZeit4 
         Height          =   225
         Left            =   200
         TabIndex        =   21
         Top             =   3000
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zeitraum"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton optZeit3 
         Height          =   225
         Left            =   200
         TabIndex        =   22
         Top             =   2400
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Jahr"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton optZeit2 
         Height          =   225
         Left            =   200
         TabIndex        =   23
         Top             =   1800
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Quartal"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton optZeit1 
         Height          =   225
         Left            =   200
         TabIndex        =   24
         Top             =   1200
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Monat"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox cmbMonat 
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   1160
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbQurta 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   1760
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   2960
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   315
         Left            =   1200
         TabIndex        =   28
         Top             =   3460
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cmbJahre 
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   2360
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis"
         Height          =   195
         Left            =   200
         TabIndex        =   30
         Top             =   3500
         Width           =   900
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkLoWei 
         Height          =   220
         Left            =   440
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2250
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Einträge zusätzlich auf Löschungen prüfen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkVergl 
         Height          =   220
         Left            =   440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1850
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Einträge zusätzlich auf Änderungen prüfen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu1 
         Height          =   405
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4000
         Visible         =   0   'False
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   706
         _StockProps     =   64
         Show3DBorder    =   2
         VisualTheme     =   0
      End
      Begin XtremeSuiteControls.ComboBox cbmAbgle 
         Height          =   315
         Left            =   400
         TabIndex        =   10
         Top             =   1300
         Width           =   4500
         _Version        =   1048579
         _ExtentX        =   7938
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   315
         Left            =   400
         TabIndex        =   14
         Top             =   3560
         Width           =   4500
         _Version        =   1048579
         _ExtentX        =   7938
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.CheckBox chkMapFo 
         Height          =   225
         Left            =   440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2800
         Width           =   4995
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Outlook Standard MAPI-Ordner wählen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   225
         Left            =   440
         TabIndex        =   16
         Top             =   3320
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   406
         _StockProps     =   79
         Caption         =   "für Mandant :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   225
         Left            =   440
         TabIndex        =   15
         Top             =   1060
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   406
         _StockProps     =   79
         Caption         =   "Datenabgleich :"
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOutlook.frx":6852
         Height          =   500
         Left            =   400
         TabIndex        =   2
         Top             =   200
         Width           =   5800
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   7000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4400
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   10200
      _Version        =   1048579
      _ExtentX        =   17992
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   2800
         Left            =   40
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1500
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   4939
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeReportControl.ReportControl repCont2 
         Height          =   2800
         Left            =   5100
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1500
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   4939
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkSync1 
         Height          =   240
         Left            =   400
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1160
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Alle Outlook Einträge markieren"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkSync2 
         Height          =   240
         Left            =   5400
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1160
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Alle Eigenen Einträge markieren"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFilt1 
         Height          =   240
         Left            =   400
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   820
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Nur markierte Einträge zeigen"
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFilt2 
         Height          =   240
         Left            =   5400
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   820
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Nur markierte Einträge zeigen"
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOutlook.frx":68EB
         Height          =   500
         Left            =   400
         TabIndex        =   9
         Top             =   200
         Width           =   8300
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   4400
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   10200
      _Version        =   1048579
      _ExtentX        =   17992
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ProgressBar prbStat2 
         Height          =   345
         Left            =   1000
         TabIndex        =   40
         Top             =   1400
         Width           =   8000
         _Version        =   1048579
         _ExtentX        =   14111
         _ExtentY        =   609
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ProgressBar prbStat1 
         Height          =   345
         Left            =   1000
         TabIndex        =   41
         Top             =   2100
         Width           =   8000
         _Version        =   1048579
         _ExtentX        =   14111
         _ExtentY        =   609
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAbbru 
         Height          =   345
         Left            =   1040
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1155
         _Version        =   1048579
         _ExtentX        =   2046
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "&Abbrechen"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   225
         Left            =   5800
         TabIndex        =   44
         Top             =   2600
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "..."
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   225
         Left            =   1040
         TabIndex        =   8
         Top             =   1100
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Synchronisation bitte warten ..."
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxDum As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private CmMon As XtremeSuiteControls.ComboBox
Private CmQua As XtremeSuiteControls.ComboBox
Private CmJah As XtremeSuiteControls.ComboBox
Private CmAbg As XtremeSuiteControls.ComboBox
Private CmBeh As XtremeSuiteControls.ComboBox
Private ChSy1 As XtremeSuiteControls.CheckBox
Private ChSy2 As XtremeSuiteControls.CheckBox
Private ChFi1 As XtremeSuiteControls.CheckBox
Private ChFi2 As XtremeSuiteControls.CheckBox
Private ChMap As XtremeSuiteControls.CheckBox
Private ChVgl As XtremeSuiteControls.CheckBox
Private ChLoe As XtremeSuiteControls.CheckBox
Private ChOp1 As XtremeSuiteControls.RadioButton
Private ChOp2 As XtremeSuiteControls.RadioButton
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private OpZei As XtremeSuiteControls.RadioButton
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private PuBu4 As XtremeSuiteControls.PushButton
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ImMan As XtremeCommandBars.ImageManager
Private MoKal As XtremeCalendarControl.DatePicker

Private GeZa1 As Long
Private GeZa2 As Long
Private FiAk1 As Boolean
Private FiAk2 As Boolean
Private FiAdr As Variant 'Outlookabgleich Eigene Adressen
Private FiKon As Variant 'Outlookabgleich Outlook Kontakte
Private FiEvt As Variant 'Outlookabgleich Eigene Events
Private FiTer As Variant 'Outlookabgleich Outlook Termine

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private KalWa As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FChek(ByVal TaTyp As Integer)
On Error GoTo InErr

Dim GesZa As Long
Dim AktZa As Long
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim OuAbg As Integer
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmOutlook
Set ChSy1 = FM.chkSync1
Set ChSy2 = FM.chkSync2
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set CmAbg = FM.cbmAbgle

OuAbg = CmAbg.ListIndex

Select Case TaTyp
Case 1:
    TeTit = "Alle Outlook Einträge markieren"
    TeMai = "Möchten Sie jetzt wirklich alle Outlook Einträge markieren?"
    TeInh = "Wenn Sie mit Ja bestätigen, werden alle Outlook Einträge zum erneuten Import frei gegeben. Dieses kann dazu führen, dass diese Einträge in Ihrer Datenbank nach dem Import mehrfach vorkommen können."
    TeFus = "Sie können in der Terminübersicht bzw. in der Adressenverwaltung alle Einträge markieren und mit der rechten Maustaste auch als nicht synchronisiert kennzeichnen."
Case 2:
    TeTit = "Alle eigenen Einträge markieren"
    TeMai = "Möchten Sie jetzt wirklich alle eigenen Einträge markieren?"
    TeInh = "Wenn Sie mit Ja bestätigen, werden alle eigenen Einträge zum erneuten Export frei gegeben. Dieses kann dazu führen, dass diese Einträge in Outlook nach dem Export  mehrfach vorkommen können."
    TeFus = "Sie können in der Terminübersicht bzw. in der Adressenverwaltung alle Einträge markieren und mit der rechten Maustaste auch als nicht synchronisiert kennzeichnen."
End Select

Select Case TaTyp
Case 1:
    Set RpCls = RpCo1.Columns
    Set RpRcs = RpCo1.Records
    If ChSy1.Value <> xtpChecked Then Exit Sub
Case 2:
    Set RpCls = RpCo2.Columns
    Set RpRcs = RpCo2.Records
    If ChSy2.Value <> xtpChecked Then Exit Sub
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, False, FM.hwnd

GesZa = RpRcs.Count - 1

If GesZa > 0 Then
    If GlMes = 33565 Then
        Select Case TaTyp
        Case 1:
            If OuAbg = 0 Or OuAbg = 2 Then
                For AktZa = 0 To GesZa
                    OuKon(26, AktZa) = False
                Next AktZa
            End If
            If OuAbg = 1 Or OuAbg = 3 Then
                For AktZa = 0 To GesZa
                    OuTer(15, AktZa) = False
                Next AktZa
            End If
            DoEvents
            RpCo1.Redraw
            ChSy1.Enabled = False
        Case 2:
            If OuAbg = 0 Or OuAbg = 2 Then
                For AktZa = 0 To GesZa
                    OuAdr(139, AktZa) = False 'Replicated
                    OuAdr(51, AktZa) = Null 'LastModification
                Next AktZa
            End If
            If OuAbg = 1 Or OuAbg = 3 Then
                For AktZa = 0 To GesZa
                    OuEvt(36, AktZa) = False 'Replicated
                    OuEvt(17, AktZa) = Null 'LastModification
                Next AktZa
            End If
            DoEvents
            RpCo2.Redraw
            ChSy2.Enabled = False
        End Select
    Else
        Select Case TaTyp
        Case 1: ChSy1.Value = xtpUnchecked
        Case 2: ChSy2.Value = xtpUnchecked
        End Select
    End If
End If

Set RpCol = Nothing
Set RpRcs = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FChek " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
    Else
        NeuDa = Date
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
    Else
        NeuDa = Date
    End If
End Select

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    Select Case KalWa
    Case 1: .Top = TxDa1.Top + TxDa1.Height
            .Left = TxDa1.Left + Rahm3.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa1.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left + Rahm3.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    End Select
End With

Datu1 = TxDa1.Text
Datu2 = TxDa2.Text

If Datu2 < Datu1 Then TxDa1.Text = Datu2

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    Select Case KalWa
    Case 1: TxDa1.Text = NeuDa
            TxDa1.SetFocus
    Case 2: TxDa2.Text = NeuDa
            TxDa2.SetFocus
    End Select
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FFilt(ByVal TaTyp As Integer)
On Error GoTo InErr

Dim AktZa As Long
Dim TmpZa As Long
Dim DatZa As Long
Dim FelZa As Long
Dim OuAbg As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpRc1 As XtremeReportControl.ReportRecords
Dim RpRc2 As XtremeReportControl.ReportRecords

Set ChFi1 = Me.chkFilt1
Set ChFi2 = Me.chkFilt2
Set ChSy1 = Me.chkSync1
Set ChSy2 = Me.chkSync2
Set RpCo1 = Me.repCont1
Set RpCo2 = Me.repCont2
Set PuBu3 = Me.btnWeiter
Set RpRc1 = RpCo1.Records
Set RpRc2 = RpCo2.Records

OuAbg = CmAbg.ListIndex

Select Case TaTyp
Case 1:
    If ChFi1.Value = xtpChecked Then
        PuBu3.Enabled = False
        FiAk1 = True
        ChSy1.Enabled = False
        If OuAbg = 0 Or OuAbg = 2 Then
            For AktZa = 0 To GeZa1 - 1
                If CBool(OuKon(26, AktZa)) = False Then 'Replaicatded?
                    TmpZa = TmpZa + 1
                End If
            Next AktZa
            DoEvents
            If TmpZa > 0 Then
                ReDim FiKon(28, TmpZa - 1)
                For AktZa = 0 To GeZa1 - 1
                    If CBool(OuKon(26, AktZa)) = False Then 'Replaicatded?
                        For FelZa = 0 To 28
                            FiKon(FelZa, DatZa) = OuKon(FelZa, AktZa)
                        Next FelZa
                        DatZa = DatZa + 1
                    End If
                Next AktZa
                With RpCo1
                     If .Records.Count > 0 Then .Records.DeleteAll
                    .SetVirtualMode DatZa
                    .Populate
                    .SetCustomDraw xtpCustomBeforeDrawRow
                End With
            End If
        Else
            For AktZa = 0 To GeZa1 - 1
                If CBool(OuTer(15, AktZa)) = False Then 'Replaicatded?
                    TmpZa = TmpZa + 1
                End If
            Next AktZa
            DoEvents
            If TmpZa > 0 Then
                ReDim FiTer(17, TmpZa - 1)
                For AktZa = 0 To GeZa1 - 1
                    If CBool(OuTer(15, AktZa)) = False Then 'Replaicatded?
                        For FelZa = 0 To 17
                            FiTer(FelZa, DatZa) = OuTer(FelZa, AktZa)
                        Next FelZa
                        DatZa = DatZa + 1
                    End If
                Next AktZa
                With RpCo1
                     If .Records.Count > 0 Then .Records.DeleteAll
                    .SetVirtualMode DatZa
                    .Populate
                    .SetCustomDraw xtpCustomBeforeDrawRow
                End With
            End If
        End If
    Else
        PuBu3.Enabled = True
        FiAk1 = False
        ChSy1.Enabled = True
        With RpCo1
             If .Records.Count > 0 Then .Records.DeleteAll
            .SetVirtualMode GeZa2
            .Populate
            .SetCustomDraw xtpCustomBeforeDrawRow
        End With
    End If
Case 2:
    If ChFi2.Value = xtpChecked Then
        PuBu3.Enabled = False
        FiAk2 = True
        ChSy2.Enabled = False
        If OuAbg = 0 Or OuAbg = 2 Then
            For AktZa = 0 To GeZa2 - 1
                If CBool(OuAdr(139, AktZa)) = False Then 'Replaicatded?
                    TmpZa = TmpZa + 1
                End If
            Next AktZa
            DoEvents
            If TmpZa > 0 Then
                ReDim FiAdr(139, TmpZa - 1)
                For AktZa = 0 To GeZa2 - 1
                    If CBool(OuAdr(139, AktZa)) = False Then 'Replaicatded?
                        For FelZa = 0 To 139
                            FiAdr(FelZa, DatZa) = OuAdr(FelZa, AktZa)
                        Next FelZa
                        DatZa = DatZa + 1
                    End If
                Next AktZa
                With RpCo2
                     If .Records.Count > 0 Then .Records.DeleteAll
                    .SetVirtualMode DatZa
                    .Populate
                    .SetCustomDraw xtpCustomBeforeDrawRow
                End With
            End If
        Else
            For AktZa = 0 To GeZa2 - 1
                If CBool(OuEvt(36, AktZa)) = False Then 'Replaicatded?
                    TmpZa = TmpZa + 1
                End If
            Next AktZa
            DoEvents
            If TmpZa > 0 Then
                ReDim FiEvt(42, TmpZa - 1)
                For AktZa = 0 To GeZa2 - 1
                    If CBool(OuEvt(36, AktZa)) = False Then 'Replaicatded?
                        For FelZa = 0 To 42
                            FiEvt(FelZa, DatZa) = OuEvt(FelZa, AktZa)
                        Next FelZa
                        DatZa = DatZa + 1
                    End If
                Next AktZa
                With RpCo2
                     If .Records.Count > 0 Then .Records.DeleteAll
                    .SetVirtualMode DatZa
                    .Populate
                    .SetCustomDraw xtpCustomBeforeDrawRow
                End With
            End If
        End If
    Else
        PuBu3.Enabled = True
        FiAk2 = False
        ChSy2.Enabled = True
        With RpCo2
             If .Records.Count > 0 Then .Records.DeleteAll
            .SetVirtualMode GeZa2
            .Populate
            .SetCustomDraw xtpCustomBeforeDrawRow
        End With
    End If
End Select

Set RpRc1 = Nothing
Set RpRc2 = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FFilt " & Err.Number
Resume Next

End Sub
Private Function FStar(Optional ByVal RepEi As Boolean) As String
On Error GoTo InErr

Dim DaSta As Date
Dim DaEnd As Date
Dim ManNr As Long
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim Krit1 As String
Dim Krit2 As String
Dim Krit3 As String
Dim Datu1 As String
Dim Datu2 As String
Dim GruKy As String
Dim GrIdx As String
Dim AkMon As Integer
Dim AkJha As Integer
Dim AkQua As Integer
Dim OuAbg As Integer
Dim AktZa As Integer
Dim Mld1, Tit1 As String

Set CmBeh = Me.cmbBehan
Set CmAbg = Me.cbmAbgle
Set OpMon = Me.optZeit1
Set OpQua = Me.optZeit2
Set OpJah = Me.optZeit3
Set OpZei = Me.optZeit4
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQurta
Set CmJah = Me.cmbJahre
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set ChOp1 = Me.optOpti1
Set ChOp2 = Me.optOpti2
Set TrLi1 = Me.trvList1

AktZa = 1

OuAbg = CmAbg.ListIndex
ManNr = CmBeh.ItemData(CmBeh.ListIndex)

If IsDate(TxDa1.Text) Then
    DaSta = TxDa1.Text
Else
    DaSta = Date
End If

If IsDate(TxDa2.Text) Then
    DaEnd = TxDa2.Text
Else
    DaEnd = Date
End If

AkJha = CInt(CmJah.Text)
AkMon = CmMon.ItemData(CmMon.ListIndex)
AkQua = CmQua.ItemData(CmQua.ListIndex)

Datu1 = DatePart("m", DaSta) & "/" & DatePart("d", DaSta) & "/" & DatePart("yyyy", DaSta)
Datu2 = DatePart("m", DaEnd) & "/" & DatePart("d", DaEnd) & "/" & DatePart("yyyy", DaEnd)
Mld1 = "Sie haben keinen Auswertungszeitraum gewählt"
Tit1 = "Einzelbriefübergabe"

Select Case OuAbg
Case 0:
    If TrLi1.Nodes("P801").Checked = True Then
        Select Case GlTyp
        Case 0: Krit1 = "(ID0 > 0)"
        Case 1: Krit1 = "(ID0 > 0)"
        Case 2: Krit1 = "([ID0] > 0)"
        Case 3: Krit1 = "([ID0] > 0)"
        End Select
    Else
        If TrLi1.Nodes("P802").Checked = True Then
            Select Case GlTyp
            Case 0: SQL1 = "(Mailing = 1)"
            Case 1: SQL1 = "(Mailing = 1)"
            Case 2: SQL1 = "([Mailing] = -1)"
            Case 3: SQL1 = "([Mailing] = -1)"
            End Select
        End If

        If TrLi1.Nodes("P803").Checked = True Then
            Select Case GlTyp
            Case 0: SQL2 = "(Edit = 1)"
            Case 1: SQL2 = "(Edit = 1)"
            Case 2: SQL2 = "([Edit] = -1)"
            Case 3: SQL2 = "([Edit] = -1)"
            End Select
        End If
    
        For Each Knote In TrLi1.Nodes
            If Knote.Checked = True Then
                If Knote.Key <> "P801" Then
                    If Knote.Key <> "P802" Then
                        If Knote.Key <> "P803" Then
                            GrIdx = Mid$(Knote.Key, 2, Len(Knote.Key) - 1)
                            GruKy = "o" & GrIdx & "o"
                            If ChOp1.Value = True Then
                                If AktZa > 1 Then
                                    SQL3 = SQL3 & " AND [TreKey] Like '%" & GruKy & "%'"
                                Else
                                    SQL3 = SQL3 & "[TreKey] Like '%" & GruKy & "%'"
                                End If
                            Else
                                If AktZa > 1 Then
                                    SQL3 = SQL3 & " OR [TreKey] Like '%" & GruKy & "%'"
                                Else
                                    SQL3 = SQL3 & "[TreKey] Like '%" & GruKy & "%'"
                                End If
                            End If
                            AktZa = AktZa + 1
                        End If
                    End If
                End If
            End If
        Next Knote

        If SQL1 <> vbNullString Then
            If SQL2 <> vbNullString Then
                If SQL3 <> vbNullString Then
                    If ChOp1.Value = True Then
                        Krit1 = "(" & SQL1 & " AND " & SQL2 & " AND " & SQL3 & ")"
                    Else
                        Krit1 = "(" & SQL1 & " OR " & SQL2 & " OR " & SQL3 & ")"
                    End If
                Else
                    If ChOp1.Value = True Then
                        Krit1 = "(" & SQL1 & " AND " & SQL2 & ")"
                    Else
                        Krit1 = "(" & SQL1 & " OR " & SQL2 & ")"
                    End If
                End If
            Else
                Krit1 = "(" & SQL1 & ")"
            End If
        ElseIf SQL2 <> vbNullString Then
            If SQL3 <> vbNullString Then
                If ChOp1.Value = True Then
                    Krit1 = "(" & SQL2 & " AND " & SQL3 & ")"
                Else
                    Krit1 = "(" & SQL2 & " OR " & SQL3 & ")"
                End If
            Else
                Krit1 = "(" & SQL2 & ")"
            End If
        ElseIf SQL3 <> vbNullString Then
                Krit1 = "(" & SQL3 & ")"
        End If
    End If
Case 1:
    If OpMon.Value = True Then
        Krit1 = "(((Month([VonDat]))=" & AkMon & ") AND ((Year([VonDat]))=" & AkJha & "))"
    ElseIf OpQua.Value = True Then
        Select Case GlTyp
        Case 0:
            Select Case AkQua
            Case 1: Krit1 = "((VonDat >= '01.01." & AkJha & "') AND (VonDat <= '31.03." & AkJha & "'))"
            Case 2: Krit1 = "((VonDat >= '01.04." & AkJha & "') AND (VonDat <= '30.06." & AkJha & "'))"
            Case 3: Krit1 = "((VonDat >= '01.07." & AkJha & "') AND (VonDat <= '30.09." & AkJha & "'))"
            Case 4: Krit1 = "((VonDat >= '01.10." & AkJha & "') AND (VonDat <= '31.12." & AkJha & "'))"
            End Select
        Case 1:
            Select Case AkQua
            Case 1: Krit1 = "((VonDat >= '01.01." & AkJha & "') AND (VonDat <= '31.03." & AkJha & "'))"
            Case 2: Krit1 = "((VonDat >= '01.04." & AkJha & "') AND (VonDat <= '30.06." & AkJha & "'))"
            Case 3: Krit1 = "((VonDat >= '01.07." & AkJha & "') AND (VonDat <= '30.09." & AkJha & "'))"
            Case 4: Krit1 = "((VonDat >= '01.10." & AkJha & "') AND (VonDat <= '31.12." & AkJha & "'))"
            End Select
        Case 2:
            Select Case AkQua
            Case 1: Krit1 = "(([VonDat] Between #01/01/" & AkJha & "# AND #03/31/" & AkJha & "#))"
            Case 2: Krit1 = "(([VonDat] Between #04/01/" & AkJha & "# AND #06/30/" & AkJha & "#))"
            Case 3: Krit1 = "(([VonDat] Between #07/01/" & AkJha & "# AND #09/30/" & AkJha & "#))"
            Case 4: Krit1 = "(([VonDat] Between #10/01/" & AkJha & "# AND #12/31/" & AkJha & "#))"
            End Select
        Case 3:
            Select Case AkQua
            Case 1: Krit1 = "(([VonDat] Between #01/01/" & AkJha & "# AND #03/31/" & AkJha & "#))"
            Case 2: Krit1 = "(([VonDat] Between #04/01/" & AkJha & "# AND #06/30/" & AkJha & "#))"
            Case 3: Krit1 = "(([VonDat] Between #07/01/" & AkJha & "# AND #09/30/" & AkJha & "#))"
            Case 4: Krit1 = "(([VonDat] Between #10/01/" & AkJha & "# AND #12/31/" & AkJha & "#))"
            End Select
        End Select
    ElseIf OpJah.Value = True Then
        Krit1 = "((Year([VonDat])=" & AkJha & "))"
    ElseIf OpZei.Value = True Then
        Select Case GlTyp
        Case 0: Krit1 = "(([VonDat] >= '" & DaSta & "') AND ([VonDat] <= '" & DaEnd & "'))"
        Case 1: Krit1 = "(([VonDat] >= '" & DaSta & "') AND ([VonDat] <= '" & DaEnd & "'))"
        Case 2: Krit1 = "(([VonDat] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
        Case 3: Krit1 = "(([VonDat] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
        End Select
    Else
        WindowMess Mld1, Dial2, Tit1, Me.hwnd
    End If
End Select

Krit3 = " AND (([Replicated]= 0))"
Krit2 = " AND (([IDP]=" & ManNr & "))"

If RepEi = True Then
    If Krit1 <> vbNullString Then
        FStar = Krit1 & Krit2 & Krit3
    Else
        FStar = vbNullString
    End If
Else
    If Krit1 <> vbNullString Then
        FStar = Krit1 & Krit2
    Else
        FStar = vbNullString
    End If
End If

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStar " & Err.Number
Resume Next

End Function
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
        TxDa1.Text = NeuDa
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
        TxDa2.Text = NeuDa
    End If
End Select

With MoKal
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo InErr

Dim NeuDa As Date
Dim FiNam As String
Dim Krite As String
Dim KmStr As String
Dim OuAbg As Integer
Dim DaVer As Boolean
Dim DaLoe As Boolean
Dim PrTer As Boolean
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpRc1 As XtremeReportControl.ReportRecords
Dim RpRc2 As XtremeReportControl.ReportRecords

Set ChFi1 = Me.chkFilt1
Set ChFi2 = Me.chkFilt2
Set TxDum = Me.txtDummy
Set ChSy1 = Me.chkSync1
Set ChSy2 = Me.chkSync2
Set ChMap = Me.chkMapFo
Set ChVgl = Me.chkVergl
Set ChLoe = Me.chkLoWei
Set RpCo1 = Me.repCont1
Set RpCo2 = Me.repCont2
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set CmAbg = Me.cbmAbgle
Set PuBu3 = Me.btnWeiter
Set PuBu4 = Me.btnZurück

Set RpRc1 = RpCo1.Records
Set RpRc2 = RpCo2.Records

OuAbg = CmAbg.ListIndex

If ChVgl.Value = xtpChecked Then DaVer = True
If ChLoe.Value = xtpChecked Then DaLoe = True

If Rahm1.Visible = True Then
    PrTer = CBool(IniGetVal("TerSys", "OutPri"))
    
    Krite = FStar()
    DoEvents
    
    IniSetVal "System", "OutAbg", OuAbg
    
    If ChMap.Value = xtpChecked Then
        IniSetVal "System", "OutStM", -1
    Else
        IniSetVal "System", "OutStM", 0
    End If
    If ChVgl.Value = xtpChecked Then
        IniSetVal "System", "OutVgl", -1
    Else
        IniSetVal "System", "OutVgl", 0
    End If
    If ChLoe.Value = xtpChecked Then
        IniSetVal "System", "OutLoe", -1
    Else
        IniSetVal "System", "OutLoe", 0
    End If
    
    With RpCo1
        .EditItem Nothing, Nothing
        If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
        If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
        If .Records.Count > 0 Then .Records.DeleteAll
        If .Columns.Count > 0 Then .Columns.DeleteAll
        .Populate
    End With
    
    With RpCo2
        .EditItem Nothing, Nothing
        If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
        If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
        If .Records.Count > 0 Then .Records.DeleteAll
        If .Columns.Count > 0 Then .Columns.DeleteAll
        .Populate
    End With
    
    If OuAbg = 0 Or OuAbg = 2 Then
        Set RpCls = RpCo1.Columns
        Set RpRcs = RpCo1.Records
        With RpCls
            Set RpCol = .Add(0, "ID0", 0, False)
            Set RpCol = .Add(1, "Mandant", 0, False)
            Set RpCol = .Add(2, "Suchbegriff", 200, True)
            RpCol.AutoSize = True
            Set RpCol = .Add(3, "Titel", 0, False)
            Set RpCol = .Add(4, "R_Firma1", 0, False)
            Set RpCol = .Add(5, "Vorname", 0, False)
            Set RpCol = .Add(6, "Name", 0, False)
            Set RpCol = .Add(7, "Bemerkung", 0, False)
            Set RpCol = .Add(8, "R_Straße", 0, False)
            Set RpCol = .Add(9, "R_Ort", 0, False)
            Set RpCol = .Add(10, "R_PLZ", 0, False)
            Set RpCol = .Add(11, "R_Land", 0, False)
            Set RpCol = .Add(12, "Telefon1", 0, False)
            Set RpCol = .Add(13, "Telefon2", 0, False)
            Set RpCol = .Add(14, "Telefon3", 0, False)
            Set RpCol = .Add(15, "Telefon4", 0, False)
            Set RpCol = .Add(16, "Telefon5", 0, False)
            Set RpCol = .Add(17, "Telefon6", 0, False)
            Set RpCol = .Add(18, "Internet", 0, False)
            Set RpCol = .Add(19, "Straße", 0, False)
            Set RpCol = .Add(20, "Ort", 0, False)
            Set RpCol = .Add(21, "Land", 0, False)
            Set RpCol = .Add(22, "PLZ", 0, False)
            Set RpCol = .Add(23, "Beruf", 0, False)
            Set RpCol = .Add(24, "Geboren", 0, False)
            Set RpCol = .Add(25, "Synchronisiert", 80, False)
            Set RpCol = .Add(26, "", 22, False)
            RpCol.HeaderAlignment = xtpAlignmentCenter
            If OuAbg = 2 Then
                RpCol.Icon = IC16_Garbage
            Else
                RpCol.Icon = IC16_Disk_Save
            End If
        End With
        
        Set RpCls = RpCo2.Columns
        Set RpRcs = RpCo2.Records
        With RpCls
            Set RpCol = .Add(0, "ID0", 0, False)
            Set RpCol = .Add(1, "Mandant", 0, False)
            Set RpCol = .Add(2, "Suchbegriff", 200, True)
            RpCol.AutoSize = True
            Set RpCol = .Add(3, "Titel", 0, False)
            Set RpCol = .Add(4, "R_Firma1", 0, False)
            Set RpCol = .Add(5, "Vorname", 0, False)
            Set RpCol = .Add(6, "Name", 0, False)
            Set RpCol = .Add(7, "Bemerkung", 0, False)
            Set RpCol = .Add(8, "R_Straße", 0, False)
            Set RpCol = .Add(9, "R_Ort", 0, False)
            Set RpCol = .Add(10, "R_PLZ", 0, False)
            Set RpCol = .Add(11, "R_Land", 0, False)
            Set RpCol = .Add(12, "Telefon1", 0, False)
            Set RpCol = .Add(13, "Telefon2", 0, False)
            Set RpCol = .Add(14, "Telefon3", 0, False)
            Set RpCol = .Add(15, "Telefon4", 0, False)
            Set RpCol = .Add(16, "Telefon5", 0, False)
            Set RpCol = .Add(17, "Telefon6", 0, False)
            Set RpCol = .Add(18, "Internet", 0, False)
            Set RpCol = .Add(19, "Straße", 0, False)
            Set RpCol = .Add(20, "Ort", 0, False)
            Set RpCol = .Add(21, "Land", 0, False)
            Set RpCol = .Add(22, "PLZ", 0, False)
            Set RpCol = .Add(23, "Beruf", 0, False)
            Set RpCol = .Add(24, "Geboren", 0, False)
            Set RpCol = .Add(25, "Synchronisiert", 80, False)
            Set RpCol = .Add(26, "", 22, False)
            RpCol.HeaderAlignment = xtpAlignmentCenter
            If OuAbg = 2 Then
                RpCol.Icon = IC16_Garbage
            Else
                RpCol.Icon = IC16_Disk_Save
            End If
        End With
    End If
    
    If OuAbg = 1 Or OuAbg = 3 Then
        Set RpCls = RpCo1.Columns
        Set RpRcs = RpCo1.Records
        With RpCls
            Set RpCol = .Add(0, "ID2", 0, False)
            Set RpCol = .Add(1, "Datum", 100, False)
            Set RpCol = .Add(2, "Betreff", 200, True)
            RpCol.AutoSize = True
            Set RpCol = .Add(3, "Enddatum", 0, False)
            Set RpCol = .Add(4, "Von", 0, False)
            Set RpCol = .Add(5, "Bis", 0, False)
            Set RpCol = .Add(6, "Farbtyp", 0, False)
            Set RpCol = .Add(7, "Selekt", 0, False)
            Set RpCol = .Add(8, "Erinnerung", 0, False)
            Set RpCol = .Add(9, "Kommentar", 0, False)
            Set RpCol = .Add(10, "Raum", 0, False)
            Set RpCol = .Add(11, "Priorität", 0, False)
            Set RpCol = .Add(12, "Vorwarn", 0, False)
            Set RpCol = .Add(13, "Synchronisiert", 80, False)
            Set RpCol = .Add(14, "Patient", 0, False)
            Set RpCol = .Add(15, "", 22, False)
            RpCol.HeaderAlignment = xtpAlignmentCenter
            If OuAbg = 3 Then
                RpCol.Icon = IC16_Garbage
            Else
                RpCol.Icon = IC16_Disk_Save
            End If
        End With
        
        Set RpCls = RpCo2.Columns
        Set RpRcs = RpCo2.Records
        With RpCls
            Set RpCol = .Add(0, "ID2", 0, False)
            Set RpCol = .Add(1, "Datum", 100, False)
            Set RpCol = .Add(2, "Betreff", 200, True)
            RpCol.AutoSize = True
            Set RpCol = .Add(3, "Enddatum", 0, False)
            Set RpCol = .Add(4, "Von", 0, False)
            Set RpCol = .Add(5, "Bis", 0, False)
            Set RpCol = .Add(6, "Farbtyp", 0, False)
            Set RpCol = .Add(7, "Selekt", 0, False)
            Set RpCol = .Add(8, "Erinnerung", 0, False)
            Set RpCol = .Add(9, "Kommentar", 0, False)
            Set RpCol = .Add(10, "Raum", 0, False)
            Set RpCol = .Add(11, "Priorität", 0, False)
            Set RpCol = .Add(12, "Vorwarn", 0, False)
            Set RpCol = .Add(13, "Synchronisiert", 80, False)
            Set RpCol = .Add(14, "Patient", 0, False)
            Set RpCol = .Add(15, "", 22, False)
            RpCol.HeaderAlignment = xtpAlignmentCenter
            If OuAbg = 3 Then
                RpCol.Icon = IC16_Garbage
            Else
                RpCol.Icon = IC16_Disk_Save
            End If
        End With
    End If
    
    Set RpCls = RpCo1.Columns
    For Each RpCol In RpCls
        RpCol.Editable = False
        RpCol.Resizable = False
        RpCol.Sortable = False
        RpCol.AutoSortWhenGrouped = False
    Next RpCol
    
    Set RpCls = RpCo2.Columns
    For Each RpCol In RpCls
        RpCol.Editable = False
        RpCol.Resizable = False
        RpCol.Sortable = False
        RpCol.AutoSortWhenGrouped = False
    Next RpCol
    
    DoEvents
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    DoEvents
    
    Select Case OuAbg
    Case 0:
        Out_Lad OuAbg, Krite
        Out_Imp OuAbg
    Case 1:
        Out_Lad OuAbg, Krite
        Out_Imp OuAbg, PrTer
    Case 2:
        Out_Imp OuAbg
    Case 3:
        Out_Imp OuAbg
    End Select
        
    GeZa1 = RpRc1.Count
    GeZa2 = RpRc2.Count
        
    DoEvents
    Rahm1.Visible = False
    Rahm2.Visible = True
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    DoEvents
Else
    DoEvents
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    DoEvents
    
    Select Case OuAbg
    Case 0:
        Out_Sav OuAbg, DaVer, DaLoe
        SUpAd
    Case 1:
        Out_Sav OuAbg, DaVer, DaLoe
        STeAk
        SUpTe
    Case 2:
        Out_Del OuAbg
    Case 3:
        Out_Del OuAbg
    End Select
    
    DoEvents
    S_StSt1
    DoEvents
    S_StSt2
    DoEvents
    S_StSt3
    DoEvents
    
    Unload Me
End If

Set RpCol = Nothing
Set RpRcs = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub FZuru()
On Error GoTo InErr

Set CmAbg = Me.cbmAbgle
Set ChSy1 = Me.chkSync1
Set ChSy2 = Me.chkSync2
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set PuBu3 = Me.btnWeiter
Set PuBu4 = Me.btnZurück

If Rahm2.Visible = True Then
    With ChSy1
        .Enabled = True
        .Value = xtpUnchecked
    End With
    With ChSy2
        .Enabled = True
        .Value = xtpUnchecked
    End With
    Rahm1.Visible = True
    Rahm2.Visible = False
    Select Case CmAbg.ListIndex
    Case 0: Rahm3.Visible = False
            Rahm4.Visible = True
            ChSy2.Enabled = True
    Case 1: Rahm3.Visible = True
            Rahm4.Visible = False
            ChSy2.Enabled = True
    Case 2: Rahm3.Visible = False
            Rahm4.Visible = True
            ChSy2.Enabled = False
            ChSy2.Value = xtpUnchecked
    Case 3: Rahm3.Visible = True
            Rahm4.Visible = False
            ChSy2.Enabled = True
    End Select
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZuru " & Err.Number
Resume Next

End Sub
Private Sub btnAbbru_Click()
    Me.txtDummy.Text = "B"
End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub
Private Sub btnDatu2_Click()
    KalWa = 2
    FKale
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
Private Sub btnZurück_Click()
    FZuru
End Sub

Private Sub cbmAbgle_Click()
On Error Resume Next

Set CmAbg = Me.cbmAbgle
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set ChSy1 = Me.chkSync1
Set ChSy2 = Me.chkSync2
Set ChVgl = Me.chkVergl
Set ChLoe = Me.chkLoWei
Set ChMap = Me.chkMapFo

Select Case CmAbg.ListIndex
Case 0: Rahm3.Visible = False
        Rahm4.Visible = True
        ChSy2.Enabled = True
        ChVgl.Enabled = True
        ChLoe.Enabled = True
        ChMap.Enabled = True
Case 1: Rahm3.Visible = True
        Rahm4.Visible = False
        ChSy2.Enabled = True
        ChVgl.Enabled = True
        ChLoe.Enabled = True
        ChMap.Enabled = True
Case 2: Rahm3.Visible = False
        Rahm4.Visible = True
        ChSy2.Enabled = False
        ChSy2.Value = xtpUnchecked
        ChVgl.Enabled = False
        ChLoe.Enabled = False
        ChMap.Enabled = False
Case 3: Rahm3.Visible = True
        Rahm4.Visible = False
        ChSy2.Enabled = True
        ChVgl.Enabled = False
        ChLoe.Enabled = False
        ChMap.Enabled = False
End Select

ReDim OuAdr(0, 0)
ReDim OuKon(0, 0)
ReDim OuEvt(0, 0)
ReDim OuTer(0, 0)

End Sub
Private Sub chkFilt1_Click()
    FFilt 1
End Sub
Private Sub chkFilt2_Click()
    FFilt 2
End Sub

Private Sub chkMapFo_Click()
On Error Resume Next

Set ChMap = Me.chkMapFo

If ChMap.Value = xtpChecked Then
    GlMaF = True
Else
    GlMaF = False
End If

End Sub
Private Sub chkSync1_Click()
    FChek 1
End Sub
Private Sub chkSync2_Click()
    FChek 2
End Sub
Private Sub cmbJahre_Click()
    Me.optZeit3.Value = True
End Sub
Private Sub cmbJahre_DropDown()
    Me.optZeit3.Value = True
End Sub
Private Sub cmbMonat_Click()
    Me.optZeit1.Value = True
End Sub

Private Sub cmbMonat_DropDown()
    Me.optZeit1.Value = True
End Sub
Private Sub cmbQurta_Click()
    Me.optZeit2.Value = True
End Sub
Private Sub cmbQurta_DropDown()
    Me.optZeit2.Value = True
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub

Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub
Private Sub Form_Load()
    SFrame 1, Me.hwnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmOutlook = Nothing
End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim OuAbg As Integer

Set CmAbg = Me.cbmAbgle

OuAbg = CmAbg.ListIndex

If OuAbg = 0 Or OuAbg = 2 Then
    If FiAk1 = True Then
        Select Case Item.Index
        Case 2:
            Metrics.Text = FiKon(Item.Index, Row.Index)
            If FiKon(26, Row.Index) <> vbNullString Then
                If CBool(FiKon(26, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Pin_Norm
                End If
            End If
        Case 25:
            If FiKon(Item.Index, Row.Index) <> vbNullString Then
                Metrics.Text = DateValue(FiKon(Item.Index, Row.Index))
            End If
        Case 26:
            If FiKon(Item.Index, Row.Index) <> vbNullString Then 'Replicated
                If CBool(FiKon(Item.Index, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Check
                Else
                    Metrics.ItemIcon = 0
                End If
            End If
        Case Else:
            Metrics.Text = FiKon(Item.Index, Row.Index)
        End Select
        If FiKon(26, Row.Index) <> vbNullString Then 'Replicated
            If CBool(FiKon(26, Row.Index)) = False Then
                Metrics.Font.Bold = True
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
        End If
    Else
        Select Case Item.Index
        Case 2:
            Metrics.Text = OuKon(Item.Index, Row.Index)
            If OuKon(26, Row.Index) <> vbNullString Then
                If CBool(OuKon(26, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Pin_Norm
                End If
            End If
        Case 25:
            If OuKon(Item.Index, Row.Index) <> vbNullString Then
                Metrics.Text = DateValue(OuKon(Item.Index, Row.Index))
            End If
        Case 26:
            If OuKon(Item.Index, Row.Index) <> vbNullString Then 'Replicated
                If CBool(OuKon(Item.Index, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Check
                Else
                    Metrics.ItemIcon = 0
                End If
            End If
        Case Else:
            Metrics.Text = OuKon(Item.Index, Row.Index)
        End Select
        If OuKon(26, Row.Index) <> vbNullString Then 'Replicated
            If CBool(OuKon(26, Row.Index)) = False Then
                Metrics.Font.Bold = True
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
        End If
    End If
End If

If OuAbg = 1 Or OuAbg = 3 Then
    If FiAk1 = True Then
        Select Case Item.Index
        Case 1:
            Metrics.Text = FiTer(1, Row.Index)
            If FiTer(15, Row.Index) <> vbNullString Then
                If CBool(FiTer(15, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Pin_Norm
                End If
            End If
        Case 13:
            If FiTer(Item.Index, Row.Index) <> vbNullString Then
                Metrics.Text = DateValue(FiTer(Item.Index, Row.Index))
            End If
        Case 15:
            If FiTer(Item.Index, Row.Index) <> vbNullString Then 'Replicated
                If CBool(FiTer(Item.Index, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Check
                Else
                    Metrics.ItemIcon = 0
                End If
            End If
        Case Else:
            Metrics.Text = FiTer(Item.Index, Row.Index)
        End Select
        If FiTer(15, Row.Index) <> vbNullString Then 'Replicated
            If CBool(FiTer(15, Row.Index)) = False Then
                Metrics.Font.Bold = True
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
        End If
    Else
        Select Case Item.Index
        Case 1:
            Metrics.Text = OuTer(1, Row.Index)
            If OuTer(15, Row.Index) <> vbNullString Then
                If CBool(OuTer(15, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Pin_Norm
                End If
            End If
        Case 13:
            If OuTer(Item.Index, Row.Index) <> vbNullString Then
                Metrics.Text = DateValue(OuTer(Item.Index, Row.Index))
            End If
        Case 15:
            If OuTer(Item.Index, Row.Index) <> vbNullString Then 'Replicated
                If CBool(OuTer(Item.Index, Row.Index)) = False Then
                    Metrics.ItemIcon = IC16_Check
                Else
                    Metrics.ItemIcon = 0
                End If
            End If
        Case Else:
            Metrics.Text = OuTer(Item.Index, Row.Index)
        End Select
        If OuTer(15, Row.Index) <> vbNullString Then 'Replicated
            If CBool(OuTer(15, Row.Index)) = False Then
                Metrics.Font.Bold = True
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
        End If
    End If
End If

End Sub
Private Sub repCont1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim OuAbg As Integer
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl

Set RpCo1 = Me.repCont1
Set CmAbg = Me.cbmAbgle
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

OuAbg = CmAbg.ListIndex

If RpSel.Count > 0 Then
    Set RpRow = RpCo1.HitTest(x, y).Row
    If RpCo1.HitTest(x, y).ht = xtpHitTestReportArea Then
        If RpRow.GroupRow = False Then
            If OuAbg = 0 Or OuAbg = 2 Then
                If CBool(OuKon(26, RpRow.Index)) = True Then
                    OuKon(26, RpRow.Index) = False
                Else
                    OuKon(26, RpRow.Index) = True
                End If
            End If
            If OuAbg = 1 Or OuAbg = 3 Then
                If CBool(OuTer(15, RpRow.Index)) = True Then
                    OuTer(15, RpRow.Index) = False
                Else
                    OuTer(15, RpRow.Index) = True
                End If
            End If
        End If
    End If
End If

Set RpCo1 = Nothing

End Sub
Private Sub repCont1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Set ChSy1 = Me.chkSync1
Set ChSy2 = Me.chkSync2

With ChSy1
    .Enabled = True
    .Value = xtpUnchecked
End With

End Sub

Private Sub repCont2_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim LiIdx As Integer

Set CmAbg = Me.cbmAbgle

LiIdx = CmAbg.ListIndex

Select Case LiIdx
Case 0:
    If FiAk2 = True Then
        Select Case Item.Index
        Case 0: Metrics.Text = FiAdr(0, Row.Index)
        Case 1: Metrics.Text = FiAdr(60, Row.Index)
        Case 2: Metrics.Text = FiAdr(7, Row.Index) 'IDKurz
            If CBool(FiAdr(139, Row.Index)) = False Then 'Replicated?
                Metrics.ItemIcon = IC16_Pin_Norm
            End If
        Case 3: Metrics.Text = FiAdr(10, Row.Index)
        Case 4: Metrics.Text = FiAdr(61, Row.Index)
        Case 5: Metrics.Text = FiAdr(13, Row.Index)
        Case 6: Metrics.Text = FiAdr(12, Row.Index)
        Case 7: Metrics.Text = FiAdr(80, Row.Index)
        Case 8: Metrics.Text = FiAdr(66, Row.Index)
        Case 9: Metrics.Text = FiAdr(69, Row.Index)
        Case 10: Metrics.Text = FiAdr(68, Row.Index)
        Case 11: Metrics.Text = FiAdr(71, Row.Index)
        Case 12: Metrics.Text = FiAdr(23, Row.Index)
        Case 13: Metrics.Text = FiAdr(22, Row.Index)
        Case 14: Metrics.Text = FiAdr(23, Row.Index)
        Case 15: Metrics.Text = FiAdr(24, Row.Index)
        Case 16: Metrics.Text = FiAdr(25, Row.Index)
        Case 17: Metrics.Text = FiAdr(26, Row.Index)
        Case 17: Metrics.Text = FiAdr(27, Row.Index)
        Case 18: Metrics.Text = FiAdr(59, Row.Index)
        Case 19: Metrics.Text = FiAdr(14, Row.Index)
        Case 20: Metrics.Text = FiAdr(17, Row.Index)
        Case 21: Metrics.Text = FiAdr(19, Row.Index)
        Case 22: Metrics.Text = FiAdr(16, Row.Index)
        Case 23: Metrics.Text = FiAdr(73, Row.Index)
        Case 24: If FiAdr(47, Row.Index) <> vbNullString Then Metrics.Text = DateValue(FiAdr(47, Row.Index))
        Case 25: If FiAdr(51, Row.Index) <> vbNullString Then Metrics.Text = DateValue(FiAdr(51, Row.Index))
        Case 26:
                If CBool(FiAdr(139, Row.Index)) = False Then 'Replaicatded?
                    If IsNull(FiAdr(51, Row.Index)) Then  'LastModification?
                        Metrics.ItemIcon = IC16_Check
                    Else
                        Metrics.ItemIcon = 0
                    End If
                Else
                    Metrics.ItemIcon = 0
                End If
        End Select
        If CBool(FiAdr(139, Row.Index)) = False Then 'Replaicatded?
            Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            If IsNull(FiAdr(51, Row.Index)) Then  'LastModification?
                Metrics.Font.Bold = True
            End If
        End If
    Else
        Select Case Item.Index
        Case 0: Metrics.Text = OuAdr(0, Row.Index)
        Case 1: Metrics.Text = OuAdr(60, Row.Index)
        Case 2: Metrics.Text = OuAdr(7, Row.Index) 'IDKurz
            If OuAdr(139, Row.Index) <> vbNullString Then
                If CBool(OuAdr(139, Row.Index)) = False Then 'Replicated?
                    Metrics.ItemIcon = IC16_Pin_Norm
                End If
            Else
                Metrics.ItemIcon = IC16_Pin_Norm
            End If
        Case 3: Metrics.Text = OuAdr(10, Row.Index)
        Case 4: Metrics.Text = OuAdr(61, Row.Index)
        Case 5: Metrics.Text = OuAdr(13, Row.Index)
        Case 6: Metrics.Text = OuAdr(12, Row.Index)
        Case 7: Metrics.Text = OuAdr(80, Row.Index)
        Case 8: Metrics.Text = OuAdr(66, Row.Index)
        Case 9: Metrics.Text = OuAdr(69, Row.Index)
        Case 10: Metrics.Text = OuAdr(68, Row.Index)
        Case 11: Metrics.Text = OuAdr(71, Row.Index)
        Case 12: Metrics.Text = OuAdr(23, Row.Index)
        Case 13: Metrics.Text = OuAdr(22, Row.Index)
        Case 14: Metrics.Text = OuAdr(23, Row.Index)
        Case 15: Metrics.Text = OuAdr(24, Row.Index)
        Case 16: Metrics.Text = OuAdr(25, Row.Index)
        Case 17: Metrics.Text = OuAdr(26, Row.Index)
        Case 17: Metrics.Text = OuAdr(27, Row.Index)
        Case 18: Metrics.Text = OuAdr(59, Row.Index)
        Case 19: Metrics.Text = OuAdr(14, Row.Index)
        Case 20: Metrics.Text = OuAdr(17, Row.Index)
        Case 21: Metrics.Text = OuAdr(19, Row.Index)
        Case 22: Metrics.Text = OuAdr(16, Row.Index)
        Case 23: Metrics.Text = OuAdr(73, Row.Index)
        Case 24: If OuAdr(47, Row.Index) <> vbNullString Then Metrics.Text = DateValue(OuAdr(47, Row.Index))
        Case 25: If OuAdr(51, Row.Index) <> vbNullString Then Metrics.Text = DateValue(OuAdr(51, Row.Index))
        Case 26:
                If OuAdr(139, Row.Index) <> vbNullString Then
                    If CBool(OuAdr(139, Row.Index)) = False Then 'Replaicatded?
                        If IsNull(OuAdr(51, Row.Index)) Then  'LastModification?
                            Metrics.ItemIcon = IC16_Check
                        Else
                            Metrics.ItemIcon = 0
                        End If
                    Else
                        Metrics.ItemIcon = 0
                    End If
                Else
                    Metrics.ItemIcon = 0
                End If
        End Select
        If OuAdr(139, Row.Index) <> vbNullString Then
            If CBool(OuAdr(139, Row.Index)) = False Then 'Replaicatded?
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
                If IsNull(OuAdr(51, Row.Index)) Then  'LastModification?
                    Metrics.Font.Bold = True
                End If
            End If
        End If
    End If
Case 1:
    If FiAk2 = True Then
        Select Case Item.Index
        Case 0: Metrics.Text = FiEvt(1, Row.Index)
        Case 1: Metrics.Text = FiEvt(5, Row.Index)
            If CBool(FiEvt(34, Row.Index)) = False Then 'Replicated?
                Metrics.ItemIcon = IC16_Pin_Norm
            End If
        Case 2:
            If FiEvt(4, Row.Index) <> vbNullString Then
                If FiEvt(14, Row.Index) <> vbNullString Then
                    Metrics.Text = FiEvt(4, Row.Index) & Chr$(32) & FiEvt(14, Row.Index)
                Else
                    Metrics.Text = FiEvt(4, Row.Index)
                End If
            Else
                If FiEvt(14, Row.Index) <> vbNullString Then
                    Metrics.Text = FiEvt(14, Row.Index)
                Else
                    Metrics.Text = "Termin"
                End If
            End If
        Case 3: Metrics.Text = FiEvt(6, Row.Index)
        Case 4: Metrics.Text = FiEvt(7, Row.Index)
        Case 5: Metrics.Text = FiEvt(8, Row.Index)
        Case 6: Metrics.Text = FiEvt(18, Row.Index)
        Case 7: Metrics.Text = FiEvt(31, Row.Index)
        Case 8: Metrics.Text = FiEvt(28, Row.Index)
        Case 9: Metrics.Text = FiEvt(26, Row.Index)
        Case 10: Metrics.Text = FiEvt(23, Row.Index)
        Case 11: Metrics.Text = FiEvt(10, Row.Index)
        Case 12: Metrics.Text = FiEvt(11, Row.Index)
        Case 13: If Not IsNull(FiEvt(17, Row.Index)) Then Metrics.Text = DateValue(FiEvt(17, Row.Index))
        Case 14: Metrics.Text = FiEvt(14, Row.Index)
        Case 15:
                If CBool(FiEvt(34, Row.Index)) = False Then 'Replaicatded?
                    If IsNull(FiEvt(17, Row.Index)) Then 'LastModification?
                        Metrics.ItemIcon = IC16_Check
                    Else
                        Metrics.ItemIcon = 0
                    End If
                Else
                    Metrics.ItemIcon = 0
                End If
        End Select
        If CBool(FiEvt(34, Row.Index)) = False Then 'Replaicatded?
            Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            If IsNull(FiEvt(17, Row.Index)) Then 'LastModification?
                Metrics.Font.Bold = True
            End If
        End If
    Else
        Select Case Item.Index
        Case 0: Metrics.Text = OuEvt(1, Row.Index)
        Case 1: Metrics.Text = OuEvt(5, Row.Index)
            If CBool(OuEvt(36, Row.Index)) = False Then 'Replicated?
                Metrics.ItemIcon = IC16_Pin_Norm
            End If
        Case 2:
            If OuEvt(4, Row.Index) <> vbNullString Then
                If OuEvt(14, Row.Index) <> vbNullString Then
                    Metrics.Text = OuEvt(4, Row.Index) & Chr$(32) & OuEvt(14, Row.Index)
                Else
                    Metrics.Text = OuEvt(4, Row.Index)
                End If
            Else
                If OuEvt(14, Row.Index) <> vbNullString Then
                    Metrics.Text = OuEvt(14, Row.Index)
                Else
                    Metrics.Text = "Termin"
                End If
            End If
        Case 3: Metrics.Text = OuEvt(6, Row.Index)
        Case 4: Metrics.Text = OuEvt(7, Row.Index)
        Case 5: Metrics.Text = OuEvt(8, Row.Index)
        Case 6: Metrics.Text = OuEvt(18, Row.Index)
        Case 7: Metrics.Text = OuEvt(31, Row.Index)
        Case 8: Metrics.Text = OuEvt(28, Row.Index)
        Case 9: Metrics.Text = OuEvt(26, Row.Index)
        Case 10: Metrics.Text = OuEvt(23, Row.Index)
        Case 11: Metrics.Text = OuEvt(10, Row.Index)
        Case 12: Metrics.Text = OuEvt(11, Row.Index)
        Case 13: If Not IsNull(OuEvt(17, Row.Index)) Then Metrics.Text = DateValue(OuEvt(17, Row.Index))
        Case 14: Metrics.Text = OuEvt(14, Row.Index)
        Case 15:
                If CBool(OuEvt(36, Row.Index)) = False Then 'Replaicatded?
                    If IsNull(OuEvt(17, Row.Index)) Then 'LastModification?
                        Metrics.ItemIcon = IC16_Check
                    Else
                        Metrics.ItemIcon = 0
                    End If
                Else
                    Metrics.ItemIcon = 0
                End If
        End Select
        If CBool(OuEvt(36, Row.Index)) = False Then 'Replaicatded?
            Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            If IsNull(OuEvt(17, Row.Index)) Then 'LastModification?
                Metrics.Font.Bold = True
            End If
        End If
    End If
End Select

End Sub
Private Sub repCont2_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo2 As XtremeReportControl.ReportControl

Set RpCo2 = Me.repCont2
Set CmAbg = Me.cbmAbgle
Set RpCls = RpCo2.Columns
Set RpSel = RpCo2.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpCo2.HitTest(x, y).Row
    If RpCo2.HitTest(x, y).ht = xtpHitTestReportArea Then
        If RpRow.GroupRow = False Then
            Select Case CmAbg.ListIndex
            Case 0:
                If CBool(OuAdr(139, RpRow.Index)) = True Then
                    OuAdr(139, RpRow.Index) = False
                Else
                    OuAdr(139, RpRow.Index) = True
                End If
            Case 1:
                If CBool(OuEvt(36, RpRow.Index)) = True Then
                    OuEvt(36, RpRow.Index) = False
                Else
                    OuEvt(36, RpRow.Index) = True
                End If
            End Select
        End If
    End If
End If

Set RpCo2 = Nothing

End Sub
Private Sub repCont2_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)

Set ChSy1 = Me.chkSync1
Set ChSy2 = Me.chkSync2

With ChSy2
    .Enabled = True
    .Value = xtpUnchecked
End With

End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub
Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
End Sub


