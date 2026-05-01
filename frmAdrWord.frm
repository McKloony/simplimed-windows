VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmAdrWord 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Berichtassistent"
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
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   52
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
         TabIndex        =   53
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
         Left            =   3400
         TabIndex        =   54
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
      Begin XtremeSuiteControls.PushButton btnZurück 
         Height          =   400
         Left            =   2000
         TabIndex        =   55
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
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Fertigstellen"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4095
      Left            =   400
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7223
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView4 
         Height          =   3080
         Left            =   100
         TabIndex        =   9
         Top             =   800
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10231
         _ExtentY        =   5433
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblLab07 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie die Dokumentenvorlage, auf dessen Basis der Bericht erstellt werden soll un klicken auf Weiter."
         Height          =   500
         Left            =   200
         TabIndex        =   38
         Top             =   200
         Width           =   5600
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4095
      Left            =   400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7223
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   3080
         Left            =   100
         TabIndex        =   10
         Top             =   800
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10231
         _ExtentY        =   5433
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrWord.frx":0000
         Height          =   500
         Left            =   200
         TabIndex        =   39
         Top             =   200
         Width           =   5600
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4095
      Left            =   400
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7223
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   310
         Left            =   3620
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3260
         Width           =   310
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   547
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   310
         Left            =   3620
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2760
         Width           =   310
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   547
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit4 
         Height          =   220
         Left            =   1400
         TabIndex        =   17
         Top             =   2800
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zeitraum"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit3 
         Height          =   220
         Left            =   1400
         TabIndex        =   15
         Top             =   2200
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Jahr"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit2 
         Height          =   220
         Left            =   1400
         TabIndex        =   13
         Top             =   1600
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Quartal"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit1 
         Height          =   220
         Left            =   1400
         TabIndex        =   11
         Top             =   1000
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Monat"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMonat 
         Height          =   310
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbQurta 
         Height          =   310
         Left            =   2400
         TabIndex        =   14
         Top             =   1560
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   310
         Left            =   2400
         TabIndex        =   18
         Top             =   2760
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   310
         Left            =   2400
         TabIndex        =   20
         Top             =   3260
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbJahre 
         Height          =   310
         Left            =   2400
         TabIndex        =   16
         Top             =   2160
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrWord.frx":008E
         Height          =   500
         Left            =   200
         TabIndex        =   37
         Top             =   200
         Width           =   5600
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis"
         Height          =   200
         Left            =   1400
         TabIndex        =   36
         Top             =   3300
         Width           =   900
      End
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   405
      Left            =   0
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5200
      Visible         =   0   'False
      Width           =   405
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   9600
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   4095
      Left            =   400
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7223
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView1 
         Height          =   1440
         Left            =   105
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   800
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10239
         _ExtentY        =   2540
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ListView lstView2 
         Height          =   1440
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2450
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10231
         _ExtentY        =   2540
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblLab09 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie aus, welche Fragebögen und welche Laborbefunde des Patienten zur Dokumentenrrstellung herangezogen werden sollen."
         Height          =   500
         Left            =   200
         TabIndex        =   41
         Top             =   200
         Width           =   5600
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm6 
      Height          =   4095
      Left            =   400
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7223
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtKomme 
         Height          =   2700
         Left            =   100
         TabIndex        =   28
         Top             =   1200
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10231
         _ExtentY        =   4762
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   14000
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   340
         Left            =   1300
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   780
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu3"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.PushButton btnDatu3 
         Height          =   315
         Left            =   1570
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   800
         Width           =   315
         _Version        =   1048579
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu3 
         Height          =   310
         Left            =   100
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   800
         Width           =   1190
         _Version        =   1048579
         _ExtentX        =   2099
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin VB.Label lblLab08 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrWord.frx":0116
         Height          =   500
         Left            =   200
         TabIndex        =   40
         Top             =   200
         Width           =   5600
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   4100
      Left            =   400
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7232
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont2 
         Height          =   3080
         Left            =   100
         TabIndex        =   24
         Top             =   800
         Width           =   5800
         _Version        =   1048579
         _ExtentX        =   10231
         _ExtentY        =   5433
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin VB.Label lblLab10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie, aus welchen Rechnungen Informationen für den zu erstellenden Bericht zusammengefasst werden sollen."
         Height          =   500
         Left            =   200
         TabIndex        =   42
         Top             =   200
         Width           =   5600
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm7 
      Height          =   4100
      Left            =   400
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7232
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtBemer 
         Height          =   300
         Left            =   1000
         TabIndex        =   33
         Top             =   3230
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   300
         Left            =   1000
         TabIndex        =   29
         Top             =   1130
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtNumme 
         Height          =   300
         Left            =   3100
         TabIndex        =   31
         Top             =   1830
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtGebor 
         Height          =   300
         Left            =   1000
         TabIndex        =   30
         Top             =   1830
         Width           =   1900
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPost 
         Height          =   300
         Left            =   1000
         TabIndex        =   32
         Top             =   2530
         Width           =   1900
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.PushButton btnPictu 
         Height          =   555
         Left            =   600
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   8040
         Width           =   555
         _Version        =   1048579
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   6
         DrawFocusRect   =   0   'False
      End
      Begin VB.Label lblLab14 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrWord.frx":01A4
         Height          =   500
         Left            =   200
         TabIndex        =   50
         Top             =   200
         Width           =   5600
      End
      Begin VB.Label lblLab12 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Geburtsdatum"
         Height          =   200
         Left            =   1000
         TabIndex        =   48
         Top             =   1600
         Width           =   2000
      End
      Begin VB.Label lblLab06 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Bemerkung"
         Height          =   200
         Left            =   1000
         TabIndex        =   47
         Top             =   3000
         Width           =   3000
      End
      Begin VB.Label lblLab05 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Postleitzahl"
         Height          =   200
         Left            =   1000
         TabIndex        =   46
         Top             =   2300
         Width           =   3000
      End
      Begin VB.Label lblLab04 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Patientenname"
         Height          =   200
         Left            =   1000
         TabIndex        =   45
         Top             =   900
         Width           =   3000
      End
      Begin VB.Label lblLab11 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Nummer"
         Height          =   200
         Left            =   3100
         TabIndex        =   44
         Top             =   1600
         Width           =   1600
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm8 
      Height          =   4100
      Left            =   400
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7232
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView3 
         Height          =   2450
         Left            =   200
         TabIndex        =   34
         Top             =   1030
         Width           =   5300
         _Version        =   1048579
         _ExtentX        =   9349
         _ExtentY        =   4322
         _StockProps     =   77
         BackColor       =   -2147483643
         View            =   2
      End
      Begin VB.Label lblLab15 
         BackStyle       =   0  'Transparent
         Caption         =   "Folgende Einträge wurden gefunden. Bitte wählen Sie den gewünschten Patienten und bestätigen mit der ENTER-Taste."
         Height          =   500
         Left            =   200
         TabIndex        =   51
         Top             =   200
         Width           =   5600
      End
      Begin VB.Label lblLab13 
         BackStyle       =   0  'Transparent
         Caption         =   "Gefundene Einträge:"
         Height          =   200
         Left            =   210
         TabIndex        =   49
         Top             =   800
         Width           =   3600
      End
   End
End
Attribute VB_Name = "frmAdrWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control

Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private FTex3 As XtremeSuiteControls.FlatEdit
Private FTex4 As XtremeSuiteControls.FlatEdit
Private FTex5 As XtremeSuiteControls.FlatEdit
Private FTex6 As XtremeSuiteControls.FlatEdit
Private TxDum As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Rahm7 As XtremeSuiteControls.GroupBox
Private Rahm8 As XtremeSuiteControls.GroupBox
Private CmMon As XtremeSuiteControls.ComboBox
Private CmQua As XtremeSuiteControls.ComboBox
Private CmJah As XtremeSuiteControls.ComboBox
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private OpZei As XtremeSuiteControls.RadioButton
Private MoKal As XtremeCalendarControl.DatePicker
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private PuBu4 As XtremeSuiteControls.PushButton
Private LiVw1 As XtremeSuiteControls.ListView
Private LiVw2 As XtremeSuiteControls.ListView
Private LiVw3 As XtremeSuiteControls.ListView
Private LiVw4 As XtremeSuiteControls.ListView
Private LiItm As XtremeSuiteControls.ListViewItem
Private LiIts As XtremeSuiteControls.ListViewItems
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ImMan As XtremeCommandBars.ImageManager

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private clFil As clsFile

Private KalWa As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date
Dim Datu3 As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
Set MoKal = Me.dtpDatu1

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
Case 3:
    If IsDate(TxDa3.Text) Then
        NeuDa = TxDa3.Text
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
            .Left = TxDa1.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa1.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 3: .Top = TxDa3.Top + TxDa3.Height
            .Left = TxDa3.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa3.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    End Select
End With

Datu1 = TxDa1.Text
Datu2 = TxDa2.Text
Datu3 = TxDa3.Text

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
Set TxDa3 = Me.txtDatu3
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    Select Case KalWa
    Case 1: TxDa1.Text = NeuDa
            TxDa1.SetFocus
    Case 2: TxDa2.Text = NeuDa
            TxDa2.SetFocus
    Case 3: TxDa3.Text = NeuDa
            TxDa3.SetFocus
    End Select
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo InErr

Dim RetWe As Long
Dim DaNam As String
Dim AkMon As Integer
Dim AkQua As Integer
Dim IdxZa As Integer
Dim BuJah As Integer
Dim AktZa As Integer
Dim AnzDa As Integer
Dim DiNam() As String
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmMain
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set Rahm7 = Me.frmRahm7
Set Rahm8 = Me.frmRahm8
Set LiVw1 = Me.lstView1
Set LiVw2 = Me.lstView2
Set LiVw3 = Me.lstView3
Set LiVw4 = Me.lstView4
Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumme
Set FTex3 = Me.txtPost
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor
Set FTex6 = Me.txtKomme
Set RpCo1 = Me.repCont1
Set RpCo2 = Me.repCont2
Set OpMon = Me.optZeit1
Set OpQua = Me.optZeit2
Set OpJah = Me.optZeit3
Set OpZei = Me.optZeit4
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQurta
Set CmJah = Me.cmbJahre
Set MoKal = Me.dtpDatu1
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
Set PuBu1 = Me.btnDatu1
Set PuBu2 = Me.btnDatu2
Set PuBu3 = Me.btnDatu3
Set ImMan = FM.imgManag
Set LiIts = LiVw4.ListItems

AkMon = Month(Date)

If AkMon <= 3 Then
    AkQua = 1
ElseIf AkMon <= 6 Then
    AkQua = 2
ElseIf AkMon <= 9 Then
    AkQua = 3
ElseIf AkMon <= 12 Then
    AkQua = 4
End If

With LiVw1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = True
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = True
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

With LiVw2
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = True
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = True
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

With LiVw3
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

With LiVw4
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = False
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = False
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
    .MultiSelect = False
    .OLEDropMode = xtpOLEDropNone
    .View = xtpListViewList
End With

With MoKal
    .AllowNoncontinuousSelection = False
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    If GlSty = 8 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    ElseIf GlSty = 7 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    Else
        .BorderStyle = xtpDatePickerBorderOffice
    End If
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    .MaxSelectionCount = 1
    .RightToLeft = False
    .ShowNoneButton = True
    .ShowNonMonthDays = True
    .ShowTodayButton = True
    .ShowWeekNumbers = False
    .TextNoneButton = "Keine"
    .TextTodayButton = "Heute"
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Rechnungstag"
    .MonthDelta = 1
    .YearsTriangle = False
    Select Case GlSty
    Case 8: .VisualTheme = xtpCalendarThemeResource
    Case 7: .VisualTheme = xtpCalendarThemeResource
    Case Else: .VisualTheme = xtpCalendarThemeResource
    End Select
    .PaintManager.ButtonTextColor = -2147483640
    .PaintManager.ControlBackColor = -2147483643
    .PaintManager.DayBackColor = -2147483643
    .PaintManager.DayTextColor = -2147483640
    .PaintManager.DaysOfWeekBackColor = -2147483643
    .PaintManager.DaysOfWeekTextColor = -2147483640
    .PaintManager.ListControlBackColor = -2147483643
    .PaintManager.ListControlTextColor = -2147483640
    .PaintManager.NonMonthDayBackColor = -2147483643
    .PaintManager.NonMonthDayTextColor = -2147483640
    .PaintManager.SelectedDayBackColor = GlFac
    .PaintManager.SelectedDayTextColor = -2147483640
    .PaintManager.WeekNumbersBackColor = -2147483643
    .PaintManager.WeekNumbersTextColor = -2147483640
    .PaintManager.MonthHeaderBackColor = GlMoB
End With

With RpCo1
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = True
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = True
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False 'Gruppierungen
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = GlAnZ
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.ThemedInplaceButtons = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.FixedRowHeight = True
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo2
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = True
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = True
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False 'Gruppierungen
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = GlAnZ
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.ThemedInplaceButtons = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.FixedRowHeight = True
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With CmMon
    .DropDownItemCount = 12
    For IdxZa = 1 To 12
        .AddItem MonthName(IdxZa)
        .ItemData(.NewIndex) = IdxZa
    Next IdxZa
End With

With CmQua
    .AddItem "1. Quartal"
    .ItemData(.NewIndex) = 1
    .AddItem "2. Quartal"
    .ItemData(.NewIndex) = 2
    .AddItem "3. Quartal"
    .ItemData(.NewIndex) = 3
    .AddItem "4. Quartal"
    .ItemData(.NewIndex) = 4
End With

With CmJah
    .DropDownItemCount = 12
    For BuJah = Year(Date) - 15 To Year(Date) + 1
        .AddItem BuJah
        .ItemData(.NewIndex) = IdxZa
        IdxZa = IdxZa + 1
    Next BuJah
    .Text = Year(Date)
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxDa3
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

FTex6.Font.Name = GlTFt.Name
FTex6.Font.SIZE = GlTFt.SIZE

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu3.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

RetWe = SendMessage(CmMon.hwnd, CB_SETCURSEL, AkMon - 1, ByVal 0&)
RetWe = SendMessage(CmQua.hwnd, CB_SETCURSEL, AkQua - 1, ByVal 0&)

Set RpCls = RpCo1.Columns
With RpCls
    Set RpCol = .Add(0, "", 0, False)
    Set RpCol = .Add(1, "Kürzel", 50, False)
    Set RpCol = .Add(2, "Bezeichnung", 200, True)
    RpCol.AutoSize = True
    Set RpCol = .Add(3, "", 0, False)
    RpCol.Visible = False
    Set RpCol = .Add(4, "Farbe", 40, False)
    RpCol.EditOptions.AllowEdit = True
    RpCol.EditOptions.AddExpandButton
    Set RpCol = .Add(5, "D", 25, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
End With

For Each RpCol In RpCls
    RpCol.Editable = True
    RpCol.Groupable = False
    RpCol.Sortable = False
Next RpCol

Set RpCls = RpCo2.Columns
With RpCls
    Set RpCol = .Add(Rec_ID1, "ID1", 0, False)
    Set RpCol = .Add(Rec_ID0, "ID0", 0, False)
    Set RpCol = .Add(Rec_RechNr, "Rechnung", 0, True)
    Set RpCol = .Add(Rec_Datum, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Rec_Selekt, "Abgeschlossen", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rec_Type, "T", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Versand, "V", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Betrag, "Betrag", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Bezahlt, "Bezahlt", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Differe, "Offen", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_IDKurz, "Patient", 0, True)
    Set RpCol = .Add(Rec_Offen, "B", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Set RpCol = .Add(Rec_Extrageb, "Extrageb.", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Rec_Fallig, "Fälligkeit", 0, True)
    Set RpCol = .Add(Rec_Wahrung, "Währung", 0, False)
    Set RpCol = .Add(Rec_IDR, "Zähler", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_ID3, "ID3", 0, False)
    Set RpCol = .Add(Rec_IDZ, "IDZ", 0, False)
    Set RpCol = .Add(Rec_Versicherer, "Katalog", 0, True)
    Set RpCol = .Add(Rec_Zahlziel, "Zahlungsziel", 0, True)
    Set RpCol = .Add(Rec_Drucken, "Drucken", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rec_IDW, "IDW", 0, False)
    Set RpCol = .Add(Rec_Symbol, "W", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Faktor, "Faktor", 0, False)
    Set RpCol = .Add(Rec_Ziel, "Ziel", 0, False)
    Set RpCol = .Add(Rec_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Rec_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Rec_Druckdatum, "Gedruckt", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Kopie, "Kopie", 0, False)
    Set RpCol = .Add(Rec_Steuer, "Steuer", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Rec_Monat, "Monat", 0, True)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        .EditOptions.Constraints.Add "Januar", 1
        .EditOptions.Constraints.Add "Februar", 2
        .EditOptions.Constraints.Add "März", 3
        .EditOptions.Constraints.Add "April", 4
        .EditOptions.Constraints.Add "Mai", 5
        .EditOptions.Constraints.Add "Juni", 6
        .EditOptions.Constraints.Add "Juli", 7
        .EditOptions.Constraints.Add "August", 8
        .EditOptions.Constraints.Add "September", 9
        .EditOptions.Constraints.Add "Oktober", 10
        .EditOptions.Constraints.Add "November", 11
        .EditOptions.Constraints.Add "Dezember", 12
    End With
    Set RpCol = .Add(Rec_Termin, "Termins.", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Storniert, "Storniert", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rec_PKU, "PKU", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Rec_Gruppe, "G", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rec_Beendet, "E", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rec_Rabatt, "Rabatt", 0, False)
    Set RpCol = .Add(Rec_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Rec_GuStr, "Gutschrift", 0, False)
    Set RpCol = .Add(Rec_GutNr, "GutNr", 0, False)
    Set RpCol = .Add(Rec_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Rec_AufNr, "AufNr", 0, False)
    Set RpCol = .Add(Rec_AuStr, "Auftrag", 0, False)
    Set RpCol = .Add(Rec_Formu, "Formular", 0, False)
    Set RpCol = .Add(Rec_OPLoe, "OPL", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.Icon = IC16_Pin_Green
    RpCol.Tag = 1
    Set RpCol = .Add(Rec_Lock, "Lock", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.Icon = IC16_Lock
    RpCol.Tag = 1
    Set RpCol = .Add(Rec_IDO, "IDO", 0, False)
    Set RpCol = .Add(Rec_RzDat, "RzDat", 0, False)
    Set RpCol = .Add(Rec_RzNum, "RzNum", 0, False)
    Set RpCol = .Add(Rec_RzTex, "RzTex", 0, False)
    Set RpCol = .Add(Rec_Grund, "Grund", 0, False)
    Set RpCol = .Add(Rec_ForID, "FID", 0, False)
End With

For Each RpCol In RpCls
    With RpCol
        .Editable = True
        .Groupable = False
        .Sortable = False
        .AutoSize = False
        .AutoSortWhenGrouped = False
    End With
Next RpCol

RpCls(Rec_ID1).Width = 0
RpCls(Rec_ID0).Width = 0
If GlTFt.SIZE > 10 Then
    RpCls(Rec_RechNr).Width = 140
    RpCls(Rec_Datum).Width = 110
Else
    RpCls(Rec_RechNr).Width = 110
    RpCls(Rec_Datum).Width = 80
End If
RpCls(Rec_Selekt).Width = 0
RpCls(Rec_Type).Width = 20
RpCls(Rec_Betrag).Width = 75
RpCls(Rec_Bezahlt).Width = 75
RpCls(Rec_Differe).Width = 75
RpCls(Rec_IDKurz).Width = 220
RpCls(Rec_Offen).Width = 0
RpCls(Rec_Extrageb).Width = 75
If GlTFt.SIZE > 10 Then
    RpCls(Rec_Fallig).Width = 110
Else
    RpCls(Rec_Fallig).Width = 80
End If
RpCls(Rec_Wahrung).Width = 0
RpCls(Rec_IDR).Width = 60
RpCls(Rec_ID3).Width = 0
RpCls(Rec_IDZ).Width = 0
RpCls(Rec_Versicherer).Width = 140
RpCls(Rec_Zahlziel).Width = 140
RpCls(Rec_Drucken).Width = 0
RpCls(Rec_IDW).Width = 0
RpCls(Rec_Symbol).Width = 30
RpCls(Rec_Faktor).Width = 0
RpCls(Rec_Ziel).Width = 0
RpCls(Rec_Kommentar).Width = 0
RpCls(Rec_IDP).Width = 0
If GlTFt.SIZE > 10 Then
    RpCls(Rec_Druckdatum).Width = 110
Else
    RpCls(Rec_Druckdatum).Width = 80
End If
RpCls(Rec_Kopie).Width = 0
RpCls(Rec_Steuer).Width = 60
RpCls(Rec_Monat).Width = 0
RpCls(Rec_Termin).Width = 75
RpCls(Rec_Storniert).Width = 0
RpCls(Rec_PKU).Width = 50
RpCls(Rec_Versand).Width = 20
RpCls(Rec_Beendet).Width = 0
RpCls(Rec_Rabatt).Width = 0
RpCls(Rec_IDM).Width = 0
If GlTFt.SIZE > 10 Then
    RpCls(Rec_GuStr).Width = 110
Else
    RpCls(Rec_GuStr).Width = 80
End If
RpCls(Rec_GutNr).Width = 0
RpCls(Rec_AufNr).Width = 0
If GlTFt.SIZE > 10 Then
    RpCls(Rec_AuStr).Width = 110
Else
    RpCls(Rec_AuStr).Width = 80
End If
RpCls(Rec_Formu).Width = 120
RpCls(Rec_OPLoe).Width = 18
RpCls(Rec_Lock).Width = 18

With LiVw1
    .ColumnHeaders.Add 1, , "Fragebogen", 3700
    .ColumnHeaders.Add 2, , "Datum", 1300
End With

With LiVw2
    .ColumnHeaders.Add 1, , "Laborauftrag", 1900
    .ColumnHeaders.Add 2, , "Labornummer", 1800
    .ColumnHeaders.Add 3, , "Datum", 1300
End With

With LiVw3
    .ColumnHeaders.Add 1, , "Adresse", 3000
    .ColumnHeaders.Add 2, , "Mandant", 1900
End With

FTex2.Pattern = "\d*"
FTex3.Pattern = "\d*"
FTex5.SetMask "00.00.0000", "__.__.____"

Set clFil = New clsFile

If clFil.FilVor(GlVor & "*.txm") = True Then
    AnzDa = clFil.FilLis(GlVor, "*.txm", DiNam)
    If AnzDa > 0 Then
        For AktZa = 1 To AnzDa
            DaNam = DiNam(AktZa)
            Set LiItm = LiIts.Add(, , DaNam, IC16_Doc_Norm)
        Next AktZa
        LiIts(1).Selected = True
    End If
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Rahm6.BackColor = GlBak
Rahm7.BackColor = GlBak
Rahm8.BackColor = GlBak
OpJah.BackColor = GlBak
OpMon.BackColor = GlBak
OpQua.BackColor = GlBak
OpJah.BackColor = GlBak
OpZei.BackColor = GlBak

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set ImMan = Nothing
Set LiVw1 = Nothing
Set LiVw2 = Nothing
Set LiVw3 = Nothing
Set LiVw4 = Nothing

Set clFil = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Function FStar() As String
On Error GoTo InErr

Dim DaSta As Date
Dim DaEnd As Date
Dim Krit1 As String
Dim Datu1 As String
Dim Datu2 As String
Dim AkMon As Integer
Dim AkJha As Integer
Dim AkQua As Integer
Dim Mld1, Tit1 As String

Set OpMon = Me.optZeit1
Set OpQua = Me.optZeit2
Set OpJah = Me.optZeit3
Set OpZei = Me.optZeit4
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQurta
Set CmJah = Me.cmbJahre
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3

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

If OpMon.Value = True Then
    If GlTyp < 2 Then
        Krit1 = "(((MONTH(Datum))=" & AkMon & ") AND ((YEAR(Datum))=" & AkJha & "))"
    Else
        Krit1 = "(((Month([Datum]))=" & AkMon & ") AND ((Year([Datum]))=" & AkJha & "))"
    End If
ElseIf OpQua.Value = True Then
    Select Case GlTyp
    Case 0:
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] >= '01.01." & AkJha & "') And ([Datum] <= '31.03." & AkJha & "'))"
        Case 2: Krit1 = "(([Datum] >= '01.04." & AkJha & "') And ([Datum] <= '30.06." & AkJha & "'))"
        Case 3: Krit1 = "(([Datum] >= '01.07." & AkJha & "') And ([Datum] <= '30.09." & AkJha & "'))"
        Case 4: Krit1 = "(([Datum] >= '01.10." & AkJha & "') And ([Datum] <= '31.12." & AkJha & "'))"
        End Select
    Case 1:
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] >= '01.01." & AkJha & "') And ([Datum] <= '31.03." & AkJha & "'))"
        Case 2: Krit1 = "(([Datum] >= '01.04." & AkJha & "') And ([Datum] <= '30.06." & AkJha & "'))"
        Case 3: Krit1 = "(([Datum] >= '01.07." & AkJha & "') And ([Datum] <= '30.09." & AkJha & "'))"
        Case 4: Krit1 = "(([Datum] >= '01.10." & AkJha & "') And ([Datum] <= '31.12." & AkJha & "'))"
        End Select
    Case 2:
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] Between #01/01/" & AkJha & "# And #03/31/" & AkJha & "#))"
        Case 2: Krit1 = "(([Datum] Between #04/01/" & AkJha & "# And #06/30/" & AkJha & "#))"
        Case 3: Krit1 = "(([Datum] Between #07/01/" & AkJha & "# And #09/30/" & AkJha & "#))"
        Case 4: Krit1 = "(([Datum] Between #10/01/" & AkJha & "# And #12/31/" & AkJha & "#))"
        End Select
    Case 3:
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] Between #01/01/" & AkJha & "# And #03/31/" & AkJha & "#))"
        Case 2: Krit1 = "(([Datum] Between #04/01/" & AkJha & "# And #06/30/" & AkJha & "#))"
        Case 3: Krit1 = "(([Datum] Between #07/01/" & AkJha & "# And #09/30/" & AkJha & "#))"
        Case 4: Krit1 = "(([Datum] Between #10/01/" & AkJha & "# And #12/31/" & AkJha & "#))"
        End Select
    End Select
ElseIf OpJah.Value = True Then
    If GlTyp < 2 Then
        Krit1 = "((YEAR(Datum) = " & AkJha & "))"
    Else
        Krit1 = "((Year([Datum]) = " & AkJha & "))"
    End If
ElseIf OpZei.Value = True Then
    Select Case GlTyp
    Case 0: Krit1 = "((Datum >= '" & DaSta & "') AND (Datum <= '" & DaEnd & "'))"
    Case 1: Krit1 = "((Datum >= '" & DaSta & "') AND (Datum <= '" & DaEnd & "'))"
    Case 2: Krit1 = "(([Datum] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
    Case 3: Krit1 = "(([Datum] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
    End Select
Else
    WindowMess Mld1, Dial2, Tit1, Me.hwnd
End If

If Krit1 <> vbNullString Then
    FStar = Krit1
Else
    FStar = vbNullString
End If

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStar " & Err.Number
Resume Next

End Function
Private Sub FSett()
On Error GoTo SeErr
'Lädt die ausgewählte Adresse in das Adressformular

Dim GesZa As Long
Dim IdxNr As Long

Set FTex1 = Me.txtKurz
Set LiVw3 = Me.lstView3
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set Rahm7 = Me.frmRahm7
Set Rahm8 = Me.frmRahm8
Set LiIts = LiVw3.ListItems

GesZa = LiVw3.ListItems.Count

If GesZa > 0 Then
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
            Exit For
        End If
    Next LiItm
    With GlSuV
        .SuIdx = 1
        .SuNum = GlAdr
    End With
    With GlSuA
        .SuIdx = 1
        .SuNum = GlAdr
    End With
    With GlSuP
        .SuIdx = 1
        .SuNum = GlAdr
    End With
    S_KrLa
    DoEvents
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
End If

GlTDa = vbNullString 'Wichtig für Textverarbeitung

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub FSuda()
On Error GoTo SeErr

Dim GesZa As Long
Dim Mld1, Tit1 As String

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumme
Set FTex3 = Me.txtPost
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor
Set Rahm7 = Me.frmRahm7
Set Rahm8 = Me.frmRahm8
Set PuBu4 = Me.btnZurück
Set LiVw3 = Me.lstView3
Set LiIts = LiVw3.ListItems

If FTex1.Text <> vbNullString Then
    GesZa = Adr_Fil(2, FTex1.Text, 1)
    If GesZa = 0 Then
        GesZa = Adr_Fil(2, SUmw(FTex1.Text), 1)
    End If
ElseIf FTex2.Text <> vbNullString Then
    GesZa = Adr_Fil(2, FTex2.Text, 2)
ElseIf FTex3.Text <> vbNullString Then
    GesZa = Adr_Fil(2, FTex3.Text, 3)
ElseIf FTex4.Text <> vbNullString Then
    GesZa = Adr_Fil(2, FTex4.Text, 4)
ElseIf FTex5.Text <> vbNullString Then
    GesZa = Adr_Fil(2, vbNullString, 5, FTex5.Text)
End If

If GesZa > 0 Then
    Rahm7.Visible = False
    Rahm8.Visible = True
    LiVw3.SetFocus
    LiIts(1).Selected = True
    PuBu4.Enabled = True
Else
    If FTex1.Text <> vbNullString Then
        FTex1.SelStart = 0
        FTex1.SelLength = Len(FTex1.Text)
    ElseIf FTex2.Text <> vbNullString Then
        FTex2.SelStart = 0
        FTex2.SelLength = Len(FTex2.Text)
    ElseIf FTex3.Text <> vbNullString Then
        FTex3.SelStart = 0
        FTex3.SelLength = Len(FTex3.Text)
    ElseIf FTex4.Text <> vbNullString Then
        FTex4.SelStart = 0
        FTex4.SelLength = Len(FTex4.Text)
    ElseIf FTex5.Text <> vbNullString Then
        FTex5.SelStart = 0
        FTex5.SelLength = Len(FTex5.Text)
    End If
    SPopu "Patient nicht gefunden", "Der von Ihnen gesuchte Patient, konnte nicht gefunden werden", IC48_Forbidden
End If

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuda " & Err.Number
Resume Next

End Sub
Private Sub TRes()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtPost
Set FTex3 = Me.txtNumme
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor

FTex1.Text = vbNullString
FTex2.Text = vbNullString
FTex3.Text = vbNullString
FTex4.Text = vbNullString
FTex5.Text = vbNullString

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
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
Case 3:
    If IsDate(TxDa3.Text) Then
        NeuDa = TxDa3.Text
        TxDa3.Text = NeuDa
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
Dim FrIdx As Long
Dim ReNum As Long
Dim FiNam As String
Dim Krite As String
Dim KmStr As String
Dim AnStr As String
Dim DiStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set LiVw1 = Me.lstView1
Set LiVw2 = Me.lstView2
Set LiVw3 = FM.lstView3
Set LiVw4 = Me.lstView4
Set TxDum = Me.txtDummy
Set TxDa3 = Me.txtDatu3
Set FTex6 = Me.txtKomme
Set RpCo1 = Me.repCont1
Set RpCo2 = Me.repCont2
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set Rahm7 = Me.frmRahm7
Set Rahm8 = Me.frmRahm8
Set PuBu4 = Me.btnZurück
Set LiIts = LiVw4.ListItems
Set RpCls = RpCo2.Columns
Set RpRcs = RpCo2.Records
Set RpSel = RpCo2.SelectedRows

If Rahm1.Visible = True Then
    If LiIts.Count > 0 Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = False
        Rahm8.Visible = False
        PuBu4.Enabled = True
        DoEvents
        RpCo1.Redraw
    End If
ElseIf Rahm2.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm3.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = True
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm4.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm5.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = True
    Rahm7.Visible = False
    Rahm8.Visible = False
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            FTex6.Text = Left$(LiItm.Text, Len(LiItm.Text) - 4)
            Exit For
        End If
    Next LiItm
ElseIf Rahm6.Visible = True Then
    FTex6.SetFocus

    If IsDate(TxDa3.Text) Then
        NeuDa = CDate(TxDa3.Text)
    Else
        NeuDa = Date
    End If
    
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            FiNam = GlVor & LiItm.Text
            Exit For
        End If
    Next LiItm
    
    Krite = FStar()
    
    Set LiIts = LiVw1.ListItems
    For Each LiItm In LiIts
        If LiItm.Checked = True Then
            FrIdx = Right$(LiItm.Key, Len(LiItm.Key) - 1)
            AnStr = AnStr & S_AnTex(FrIdx) & vbCrLf & vbCrLf
        End If
    Next LiItm

    If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Rec_IDR)
            ReNum = RpRow.Record(RpCol.ItemIndex).Value
            DiStr = S_DrDi(ReNum)
        End If
    End If
    
    KmStr = FTex6.Text

    With GlTxD
        .PatNr = GlAdr
        .FiNam = FiNam
        .Krite = Krite
        .KmStr = KmStr
        .NeuDa = NeuDa
        .FrStr = AnStr
        .DiStr = DiStr
    End With
    
    S_TxRe
    DoEvents
    S_TxLa
    DoEvents
    S_TxKr
    DoEvents
    
    Unload Me
    DoEvents

    STxOp
ElseIf Rahm7.Visible = True Then
    FSuda
ElseIf Rahm8.Visible = True Then
    FSett
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub FZuru()
On Error GoTo InErr

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set Rahm7 = Me.frmRahm7
Set Rahm8 = Me.frmRahm8
Set PuBu4 = Me.btnZurück

If Rahm2.Visible = True Then
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
    PuBu4.Enabled = False
ElseIf Rahm3.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = True
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm4.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm5.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = True
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm6.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm7.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = True
    Rahm7.Visible = False
    Rahm8.Visible = False
ElseIf Rahm8.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = True
    Rahm8.Visible = False
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZuru " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub
Private Sub btnDatu2_Click()
    KalWa = 2
    FKale
End Sub

Private Sub btnDatu3_Click()
    KalWa = 3
    FKale
End Sub

Private Sub btnZurück_Click()
    FZuru
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

Private Sub Form_Activate()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set Rahm7 = Me.frmRahm7
Set Rahm8 = Me.frmRahm8

If GlBut = RibTab_Startseite Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = True
    Rahm8.Visible = False
    FTex1.SetFocus
Else
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
End If

End Sub

Private Sub Form_Load()
On Error Resume Next

FKonf
S_WoLa
Adr_Ana
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAdrWord = Nothing
End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Item.Index = 4 Then
        Item.BackColor = Row.Record.Item(3).Value
        Item.ForeColor = Row.Record.Item(3).Value
    End If
End Sub

Private Sub repCont1_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

Dim TmTag As String

TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

Item.Tag = "@" & TmTag

S_WoSa

Item.Tag = TmTag

End Sub
Private Sub repCont1_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)

Dim TmTag As String

TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

Item.Tag = "@" & TmTag

S_WoSa

Item.Tag = TmTag

End Sub
Private Sub repCont2_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If GlSta = False Then
    If Row.GroupRow = False Then
        If Row.Record(Rec_Selekt).Value = 0 Then
            Metrics.Font.Bold = True
        End If
        Select Case Row.Record(Rec_Type).Value
        Case "M": Metrics.ForeColor = 16744448
        Case "L": Metrics.ForeColor = 33023
        Case "V": Metrics.ForeColor = 8421631
        Case "U": Metrics.ForeColor = 6604830
        Case Else:
            If Row.Record(Rec_Selekt).Value = 0 Then
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
        If Row.Record(Rec_Storniert).Value = True Then
            Metrics.Font.Strikethrough = True
            Metrics.ForeColor = 8421504
        End If
        End Select
    End If
End If

End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub
Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
End Sub
Private Sub btnHilfe_Click()
On Error GoTo InErr

Dim NeuDa As Date
Dim FrIdx As Long
Dim ReNum As Long
Dim FiNam As String
Dim Krite As String
Dim KmStr As String
Dim AnStr As String
Dim DiStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set LiVw1 = Me.lstView1
Set LiVw2 = Me.lstView2
Set LiVw3 = Me.lstView3
Set LiVw4 = Me.lstView4
Set TxDum = Me.txtDummy
Set TxDa3 = Me.txtDatu3
Set FTex6 = Me.txtKomme
Set RpCo1 = Me.repCont1
Set RpCo2 = Me.repCont2
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set Rahm7 = Me.frmRahm7
Set Rahm8 = Me.frmRahm8
Set PuBu4 = Me.btnZurück
Set LiIts = LiVw4.ListItems

If IsDate(TxDa3.Text) Then
    NeuDa = TxDa3.Text
Else
    NeuDa = Date
End If

If LiIts.Count > 0 Then
    For Each LiItm In LiIts
        If LiItm.Selected = True Then
            FiNam = GlVor & LiItm.Text
            Exit For
        End If
    Next LiItm
    
    Krite = FStar()
    
    Set LiIts = LiVw1.ListItems
    For Each LiItm In LiIts
        If LiItm.Checked = True Then
            FrIdx = Right$(LiItm.Key, Len(LiItm.Key) - 1)
            AnStr = AnStr & S_AnTex(FrIdx) & vbCrLf & vbCrLf
        End If
    Next LiItm
    
    Set RpRcs = RpCo2.Records
    For Each RpRec In RpRcs
        If RpRec.Item(1).Checked = True Then
            ReNum = RpRec.Item(0).Value
            DiStr = S_DrDi(ReNum)
            Exit For
        End If
    Next RpRec
    
    KmStr = FTex6.Text
    
    With GlTxD
        .PatNr = GlAdr
        .FiNam = FiNam
        .Krite = Krite
        .KmStr = KmStr
        .NeuDa = NeuDa
        .FrStr = AnStr
        .DiStr = DiStr
    End With
    
    S_TxRe
    DoEvents
    S_TxLa
    DoEvents
    S_TxKr
    DoEvents
    
    Unload Me
    DoEvents

    STxOp
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "Click " & Err.Number
Resume Next

End Sub

Private Sub txtDatu3_LostFocus()
    KalWa = 3
    FDaKo
End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

Set TxDa3 = Me.txtDatu3

AltDa = TxDa1.Text

TxDa3.Text = DateAdd("d", 1, AltDa)

End Sub
Private Sub updCont1_DownClick()

Dim AltDa As Date

Set TxDa3 = Me.txtDatu3

AltDa = TxDa3.Text

TxDa3.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub lstView3_DblClick()
    FSett
End Sub
Private Sub lstView3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub
Private Sub txtBemer_GotFocus()
    TRes
End Sub
Private Sub txtBemer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub

Private Sub txtGebor_GotFocus()
    TRes
End Sub
Private Sub txtGebor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtKurz_GotFocus()
    TRes
End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtNumme_GotFocus()
    TRes
End Sub
Private Sub txtNumme_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtNumme_Validate(Cancel As Boolean)
    If (Not txtNumme.isValid) Then Cancel = True
End Sub
Private Sub txtPost_GotFocus()
    TRes
End Sub
Private Sub txtPost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtPost_Validate(Cancel As Boolean)
    If (Not txtPost.isValid) Then Cancel = True
End Sub
