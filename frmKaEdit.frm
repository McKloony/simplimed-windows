VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKaEdit 
   Caption         =   "Bearbeiten"
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11100
   ScaleWidth      =   19875
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   151
      Top             =   11500
      Width           =   80
   End
   Begin VB.TextBox txtGrupe 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   240
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   11500
      Width           =   80
   End
   Begin VB.TextBox txtIdxNr 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   11500
      Width           =   80
   End
   Begin XtremeSuiteControls.GroupBox frmRahm8 
      Height          =   3500
      Left            =   6500
      TabIndex        =   0
      Top             =   7400
      Visible         =   0   'False
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6174
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ColorPicker colPick1 
         Height          =   360
         Left            =   1400
         TabIndex        =   1
         Top             =   2900
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1676
         _ExtentY        =   635
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         SelectedColor   =   16777215
         ShowAutomaticColor=   0   'False
      End
      Begin XtremeSuiteControls.UpDown UpDown2 
         Height          =   350
         Left            =   2370
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1900
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Max             =   9999
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtMnute"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtBetre 
         Height          =   350
         Left            =   1400
         TabIndex        =   3
         Top             =   400
         Width           =   4400
         _Version        =   1048579
         _ExtentX        =   7761
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtMnute 
         Height          =   350
         Left            =   1400
         TabIndex        =   4
         Top             =   1900
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1676
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "0"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   310
         Left            =   1400
         TabIndex        =   5
         Top             =   900
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbRaum1 
         Height          =   310
         Left            =   1400
         TabIndex        =   6
         Top             =   1400
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtNachl 
         Height          =   350
         Left            =   1400
         TabIndex        =   7
         Top             =   2400
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1676
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Text            =   "0"
         Alignment       =   2
      End
      Begin XtremeSuiteControls.UpDown UpDown3 
         Height          =   350
         Left            =   2370
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2400
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Max             =   9999
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtNachl"
         BuddyProperty   =   ""
      End
      Begin VB.Label lblLab22 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Terminbetreff :"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   450
         Width           =   1200
      End
      Begin VB.Label lblLab23 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   940
         Width           =   1200
      End
      Begin VB.Label lblLab25 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Terminlänge :"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1950
         Width           =   1200
      End
      Begin VB.Label lblLab24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Raumplan :"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblLab26 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Farbe :"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   2940
         Width           =   1200
      End
      Begin XtremeSuiteControls.Label lblLab27 
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   2440
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Nachlauf :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab28 
         Height          =   240
         Left            =   2700
         TabIndex        =   10
         Top             =   1950
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Min."
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab29 
         Height          =   240
         Left            =   2700
         TabIndex        =   9
         Top             =   2440
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Min."
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm7 
      Height          =   3500
      Left            =   120
      TabIndex        =   18
      Top             =   7400
      Visible         =   0   'False
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6174
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtDaBel 
         Height          =   350
         Left            =   4480
         TabIndex        =   19
         Top             =   900
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDaBes 
         Height          =   350
         Left            =   4480
         TabIndex        =   20
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDaEin 
         Height          =   350
         Left            =   4480
         TabIndex        =   21
         Top             =   1400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtLaOrt 
         Height          =   350
         Left            =   1400
         TabIndex        =   22
         Top             =   2900
         Width           =   4400
         _Version        =   1048579
         _ExtentX        =   7761
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   250
      End
      Begin XtremeSuiteControls.FlatEdit txtBeMax 
         Height          =   350
         Left            =   1400
         TabIndex        =   23
         Top             =   1400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtBeMel 
         Height          =   350
         Left            =   1400
         TabIndex        =   24
         Top             =   1900
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtBeMin 
         Height          =   350
         Left            =   1400
         TabIndex        =   25
         Top             =   900
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtBeSol 
         Height          =   350
         Left            =   1400
         TabIndex        =   26
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtBeIst 
         Height          =   350
         Left            =   1400
         TabIndex        =   27
         Top             =   2400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtMeBes 
         Height          =   350
         Left            =   4480
         TabIndex        =   28
         Top             =   1900
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtMeEin 
         Height          =   350
         Left            =   4480
         TabIndex        =   29
         Top             =   2400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin VB.Label lblLab16 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Lagerort :"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2940
         Width           =   1200
      End
      Begin VB.Label lblLab15 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Inventur-Bestand :"
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   2440
         Width           =   1300
      End
      Begin VB.Label lblLab13 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Max.-Bestand :"
         Height          =   240
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblLab11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Soll-Bestand :"
         Height          =   240
         Left            =   120
         TabIndex        =   37
         Top             =   450
         Width           =   1200
      End
      Begin VB.Label lblLab12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Min.-Bestand :"
         Height          =   240
         Left            =   120
         TabIndex        =   36
         Top             =   940
         Width           =   1200
      End
      Begin VB.Label lblLab14 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Melde-Bestand :"
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   1940
         Width           =   1200
      End
      Begin VB.Label lblLab20 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bestell-Menge :"
         Height          =   240
         Left            =   3100
         TabIndex        =   34
         Top             =   1950
         Width           =   1300
      End
      Begin VB.Label lblLab18 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bestell-Datum :"
         Height          =   240
         Left            =   3100
         TabIndex        =   33
         Top             =   940
         Width           =   1300
      End
      Begin VB.Label lblLab17 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bestand-Datum :"
         Height          =   240
         Left            =   3100
         TabIndex        =   32
         Top             =   450
         Width           =   1300
      End
      Begin VB.Label lblLab19 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Eingang- Datum :"
         Height          =   240
         Left            =   3100
         TabIndex        =   31
         Top             =   1440
         Width           =   1300
      End
      Begin VB.Label lblLab21 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Eingang-Menge :"
         Height          =   255
         Left            =   3100
         TabIndex        =   30
         Top             =   2440
         Width           =   1300
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm6 
      Height          =   3500
      Left            =   12800
      TabIndex        =   41
      Top             =   3800
      Visible         =   0   'False
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6174
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont2 
         Height          =   3000
         Left            =   100
         TabIndex        =   42
         Top             =   200
         Width           =   5500
         _Version        =   1048579
         _ExtentX        =   9701
         _ExtentY        =   5292
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   3500
      Left            =   12800
      TabIndex        =   43
      Top             =   200
      Visible         =   0   'False
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6174
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit GrK6b 
         Height          =   310
         Left            =   4605
         TabIndex        =   44
         Top             =   2650
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK6a 
         Height          =   310
         Left            =   4605
         TabIndex        =   45
         Top             =   2300
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF6b 
         Height          =   310
         Left            =   4605
         TabIndex        =   46
         Top             =   1850
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF6a 
         Height          =   310
         Left            =   4605
         TabIndex        =   47
         Top             =   1500
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM6b 
         Height          =   310
         Left            =   4605
         TabIndex        =   48
         Top             =   1050
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM6a 
         Height          =   310
         Left            =   4605
         TabIndex        =   49
         Top             =   700
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK5b 
         Height          =   310
         Left            =   3795
         TabIndex        =   50
         Top             =   2650
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK5a 
         Height          =   310
         Left            =   3795
         TabIndex        =   51
         Top             =   2300
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF5b 
         Height          =   310
         Left            =   3795
         TabIndex        =   52
         Top             =   1850
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF5a 
         Height          =   310
         Left            =   3795
         TabIndex        =   53
         Top             =   1500
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM5b 
         Height          =   310
         Left            =   3795
         TabIndex        =   54
         Top             =   1050
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM5a 
         Height          =   310
         Left            =   3795
         TabIndex        =   55
         Top             =   700
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK4b 
         Height          =   310
         Left            =   3000
         TabIndex        =   56
         Top             =   2650
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK4a 
         Height          =   310
         Left            =   3000
         TabIndex        =   57
         Top             =   2300
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF4b 
         Height          =   310
         Left            =   3000
         TabIndex        =   58
         Top             =   1850
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF4a 
         Height          =   310
         Left            =   3000
         TabIndex        =   59
         Top             =   1500
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM4b 
         Height          =   310
         Left            =   3000
         TabIndex        =   60
         Top             =   1050
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM4a 
         Height          =   310
         Left            =   3000
         TabIndex        =   61
         Top             =   700
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK3b 
         Height          =   310
         Left            =   2205
         TabIndex        =   62
         Top             =   2650
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK3a 
         Height          =   310
         Left            =   2205
         TabIndex        =   63
         Top             =   2300
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF3b 
         Height          =   310
         Left            =   2205
         TabIndex        =   64
         Top             =   1850
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF3a 
         Height          =   310
         Left            =   2205
         TabIndex        =   65
         Top             =   1500
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM3b 
         Height          =   310
         Left            =   2205
         TabIndex        =   66
         Top             =   1050
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM3a 
         Height          =   310
         Left            =   2205
         TabIndex        =   67
         Top             =   700
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK2b 
         Height          =   310
         Left            =   1395
         TabIndex        =   68
         Top             =   2650
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK2a 
         Height          =   310
         Left            =   1395
         TabIndex        =   69
         Top             =   2300
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF2b 
         Height          =   310
         Left            =   1395
         TabIndex        =   70
         Top             =   1850
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF2a 
         Height          =   310
         Left            =   1395
         TabIndex        =   71
         Top             =   1500
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM2b 
         Height          =   310
         Left            =   1395
         TabIndex        =   72
         Top             =   1050
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM2a 
         Height          =   310
         Left            =   1395
         TabIndex        =   73
         Top             =   700
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK1b 
         Height          =   310
         Left            =   600
         TabIndex        =   74
         Top             =   2650
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrK1a 
         Height          =   310
         Left            =   600
         TabIndex        =   75
         Top             =   2300
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF1b 
         Height          =   310
         Left            =   600
         TabIndex        =   76
         Top             =   1850
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrF1a 
         Height          =   310
         Left            =   600
         TabIndex        =   77
         Top             =   1500
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM1b 
         Height          =   310
         Left            =   600
         TabIndex        =   78
         Top             =   1050
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit GrM1a 
         Height          =   310
         Left            =   600
         TabIndex        =   79
         Top             =   700
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin VB.Label lblLab01 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "(--)"
         Height          =   255
         Left            =   600
         TabIndex        =   85
         Top             =   320
         Width           =   705
      End
      Begin VB.Label lblLab02 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         Height          =   255
         Left            =   1395
         TabIndex        =   84
         Top             =   320
         Width           =   705
      End
      Begin VB.Label lblLab04 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "!(-)"
         Height          =   255
         Left            =   2205
         TabIndex        =   83
         Top             =   320
         Width           =   705
      End
      Begin VB.Label lblLab05 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "(+)"
         Height          =   255
         Left            =   3795
         TabIndex        =   82
         Top             =   320
         Width           =   705
      End
      Begin VB.Label lblLab06 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "(++)"
         Height          =   255
         Left            =   4605
         TabIndex        =   81
         Top             =   320
         Width           =   705
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "(+)!"
         Height          =   255
         Left            =   3000
         TabIndex        =   80
         Top             =   320
         Width           =   705
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   3500
      Left            =   120
      TabIndex        =   86
      Top             =   3800
      Visible         =   0   'False
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6174
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   1400
         Left            =   100
         TabIndex        =   87
         Top             =   1900
         Width           =   5500
         _Version        =   1048579
         _ExtentX        =   9701
         _ExtentY        =   2469
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRest2 
         Height          =   350
         Left            =   1980
         TabIndex        =   88
         Top             =   570
         Width           =   500
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRest1 
         Height          =   350
         Left            =   1980
         TabIndex        =   89
         Top             =   190
         Width           =   500
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.RadioButton optRest3 
         Height          =   240
         Left            =   1000
         TabIndex        =   90
         Top             =   940
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Darf nur als alleinige Leistung am Tag abgerechnet werden"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optRest2 
         Height          =   240
         Left            =   1000
         TabIndex        =   91
         Top             =   600
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Darf max."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optRest1 
         Height          =   240
         Left            =   1000
         TabIndex        =   92
         Top             =   240
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Darf max."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbZiffe 
         Height          =   315
         Left            =   140
         TabIndex        =   93
         Top             =   1440
         Width           =   1140
         _Version        =   1048579
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         DropDownItemCount=   20
      End
      Begin XtremeSuiteControls.ComboBox cmbBezei 
         Height          =   315
         Left            =   1340
         TabIndex        =   94
         Top             =   1440
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7064
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         AutoComplete    =   -1  'True
         DropDownItemCount=   20
      End
      Begin XtremeSuiteControls.Label Label19 
         Height          =   240
         Left            =   2540
         TabIndex        =   96
         Top             =   240
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "mal pro Behandlungstag abgerechnet werden"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label21 
         Height          =   240
         Left            =   2540
         TabIndex        =   95
         Top             =   600
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "mal pro Rechnung abgerechnet werden"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3500
      Left            =   6500
      TabIndex        =   97
      Top             =   200
      Visible         =   0   'False
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6165
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtZusat 
         Height          =   1100
         Left            =   100
         TabIndex        =   98
         Top             =   1900
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   1940
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtKomme 
         Height          =   1000
         Left            =   100
         TabIndex        =   99
         Top             =   500
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   1764
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar :"
         Height          =   240
         Left            =   120
         TabIndex        =   101
         Top             =   260
         Width           =   1995
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Infotext / Anweisung :"
         Height          =   240
         Left            =   120
         TabIndex        =   100
         Top             =   1660
         Width           =   1995
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   3500
      Left            =   6500
      TabIndex        =   102
      Top             =   3800
      Visible         =   0   'False
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6174
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkProfi 
         Height          =   240
         Left            =   4300
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   2850
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Profilwert"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkHochs 
         Height          =   240
         Left            =   4300
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   2350
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Höchstwert"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKette 
         Height          =   240
         Left            =   4300
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   1850
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Kettenwert"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtNorKi 
         Height          =   350
         Left            =   1275
         TabIndex        =   106
         Top             =   2800
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtNorFr 
         Height          =   350
         Left            =   1275
         TabIndex        =   107
         Top             =   2300
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtNorMa 
         Height          =   350
         Left            =   1275
         TabIndex        =   108
         Top             =   1800
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtEinhe 
         Height          =   350
         Left            =   4300
         TabIndex        =   109
         Top             =   1300
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtTesCo 
         Height          =   350
         Left            =   4300
         TabIndex        =   110
         Top             =   800
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtTesId 
         Height          =   350
         Left            =   4300
         TabIndex        =   111
         Top             =   300
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbStatu 
         Height          =   310
         Left            =   1275
         TabIndex        =   112
         Top             =   300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGrupp 
         Height          =   310
         Left            =   1275
         TabIndex        =   113
         Top             =   800
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbProbe 
         Height          =   310
         Left            =   1275
         TabIndex        =   114
         Top             =   1300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin VB.Label lblLab08 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Testident :"
         Height          =   240
         Left            =   3360
         TabIndex        =   123
         Top             =   350
         Width           =   900
      End
      Begin VB.Label lblLab09 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Einheit :"
         Height          =   240
         Left            =   3360
         TabIndex        =   122
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblLab10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Testcode :"
         Height          =   240
         Left            =   3360
         TabIndex        =   121
         Top             =   850
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   240
         Left            =   140
         TabIndex        =   120
         Top             =   350
         Width           =   1100
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Gruppierung :"
         Height          =   240
         Left            =   140
         TabIndex        =   119
         Top             =   850
         Width           =   1100
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Probenident :"
         Height          =   240
         Left            =   140
         TabIndex        =   118
         Top             =   1350
         Width           =   1100
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Norm-Mann :"
         Height          =   240
         Left            =   140
         TabIndex        =   117
         Top             =   1850
         Width           =   1100
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Norm-Frau :"
         Height          =   240
         Left            =   140
         TabIndex        =   116
         Top             =   2350
         Width           =   1100
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Norm-Kind :"
         Height          =   240
         Left            =   140
         TabIndex        =   115
         Top             =   2850
         Width           =   1100
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3500
      Left            =   120
      TabIndex        =   124
      Top             =   200
      Width           =   6200
      _Version        =   1048579
      _ExtentX        =   10936
      _ExtentY        =   6174
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkFakPr 
         Height          =   255
         Left            =   2240
         TabIndex        =   125
         Top             =   1940
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Faktorpreis"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAnalo 
         Height          =   210
         Left            =   1200
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   2950
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Analog"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown UpDown1 
         Height          =   350
         Left            =   2160
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   1400
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Max             =   9999
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtMinut"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.CheckBox chkFavor 
         Height          =   240
         Left            =   2500
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   2950
         Width           =   850
         _Version        =   1048579
         _ExtentX        =   1499
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Favorit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtZiff1 
         Height          =   350
         Left            =   1200
         TabIndex        =   129
         Top             =   900
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtBezei 
         Height          =   350
         Left            =   1200
         TabIndex        =   130
         Top             =   400
         Width           =   4600
         _Version        =   1048579
         _ExtentX        =   8114
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbGrupe 
         Height          =   310
         Left            =   1200
         TabIndex        =   131
         Top             =   2400
         Width           =   4580
         _Version        =   1048579
         _ExtentX        =   8070
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtPrei1 
         Height          =   350
         Left            =   4480
         TabIndex        =   132
         Top             =   1400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtMinut 
         Height          =   350
         Left            =   1200
         TabIndex        =   133
         Top             =   1400
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1676
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "0"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMulti 
         Height          =   350
         Left            =   1200
         TabIndex        =   134
         Top             =   1900
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1676
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "1,0"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPrei2 
         Height          =   350
         Left            =   4480
         TabIndex        =   135
         Top             =   1900
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtSorte 
         Height          =   350
         Left            =   4480
         TabIndex        =   136
         Top             =   900
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "1"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtSteue 
         Height          =   350
         Left            =   1200
         TabIndex        =   137
         Top             =   2900
         Width           =   945
         _Version        =   1048579
         _ExtentX        =   1667
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "0,0"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmpEiTyp 
         Height          =   315
         Left            =   4480
         TabIndex        =   138
         Top             =   2900
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2275
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLabl7 
         Height          =   240
         Left            =   3400
         TabIndex        =   149
         Top             =   2940
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Eintragstyp :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLabl3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Faktor :"
         Height          =   240
         Left            =   120
         TabIndex        =   148
         Top             =   1950
         Width           =   1000
      End
      Begin VB.Label lblLabl1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Ziffer :"
         Height          =   240
         Left            =   120
         TabIndex        =   147
         Top             =   940
         Width           =   1000
      End
      Begin VB.Label lblLab03 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Text :"
         Height          =   240
         Left            =   120
         TabIndex        =   146
         Top             =   450
         Width           =   1000
      End
      Begin VB.Label lblLabl5 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Dauer :"
         Height          =   240
         Left            =   120
         TabIndex        =   145
         Top             =   1450
         Width           =   1000
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Sortierung :"
         Height          =   240
         Left            =   3400
         TabIndex        =   144
         Top             =   940
         Width           =   1000
      End
      Begin VB.Label lblLabl6 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Gruppe :"
         Height          =   255
         Left            =   120
         TabIndex        =   143
         Top             =   2440
         Width           =   1000
      End
      Begin VB.Label lblLabl8 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Steuer :"
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   2940
         Width           =   1000
      End
      Begin XtremeSuiteControls.Label lblLab30 
         Height          =   240
         Left            =   2500
         TabIndex        =   141
         Top             =   1450
         Width           =   400
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Min."
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl2 
         Height          =   240
         Left            =   3300
         TabIndex        =   140
         Top             =   1450
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Preis :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl4 
         Height          =   240
         Left            =   3500
         TabIndex        =   139
         Top             =   1950
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Einzelpreis :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   480
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private ColMa As XtremeCommandBars.ColorManager
Private LbPr2 As XtremeSuiteControls.Label
Private LbPr4 As XtremeSuiteControls.Label
Private TxZei As XtremeSuiteControls.FlatEdit
Private TxRs1 As XtremeSuiteControls.FlatEdit
Private TxRs2 As XtremeSuiteControls.FlatEdit
Private TxZus As XtremeSuiteControls.FlatEdit
Private TxPr1 As XtremeSuiteControls.FlatEdit
Private TxPr2 As XtremeSuiteControls.FlatEdit
Private TxFak As XtremeSuiteControls.FlatEdit
Private FakPr As XtremeSuiteControls.CheckBox
Private CmZif As XtremeSuiteControls.ComboBox
Private CmBez As XtremeSuiteControls.ComboBox
Private CmRmu As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Rahm7 As XtremeSuiteControls.GroupBox
Private Rahm8 As XtremeSuiteControls.GroupBox
Private OpRS1 As XtremeSuiteControls.RadioButton
Private OpRS2 As XtremeSuiteControls.RadioButton
Private OpRS3 As XtremeSuiteControls.RadioButton
Private CoPic As XtremeSuiteControls.ColorPicker
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
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private FaPre As Double
Private GeDia As String
Private clFen As clsFenster

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Sub FClos()
On Error GoTo LiErr

Dim RetWe As Long

Set FM = frmKaEdit

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

If GlRes = False Then 'Reset der Einstellungen
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "KatBear", "FenLin", clFen.FeLin
        IniSetVal "KatBear", "FenObe", clFen.FeObn
        IniSetVal "KatBear", "FenBre", clFen.FeBre
        IniSetVal "KatBear", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
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
Private Sub FInit()
On Error GoTo InErr

Dim RetWe As Long
Dim KeyNa As String
Dim TreKy As String
Dim AktZa As Integer
Dim TmFnt As New StdFont
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim ToTab As XtremeCommandBars.TabControlItem

Set FM = frmKaEdit
Set CmZif = FM.cmbZiffe
Set CmBez = FM.cmbBezei
Set CmMit = FM.cmbMitar
Set CmRmu = FM.cmbRaum1
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set Rahm8 = FM.frmRahm8
Set TxZei = FM.txtMinut
Set TxRs1 = FM.txtRest1
Set TxRs2 = FM.txtRest2
Set TxZus = FM.txtZusat
Set OpRS1 = FM.optRest1
Set OpRS2 = FM.optRest2
Set OpRS3 = FM.optRest3
Set CoPic = FM.colPick1
Set CmBrs = FM.comBar02
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag

KeyNa = "ToolTips"
TreKy = Left$(GlNod, 1)

TmFnt.Name = GlTFt.Name
TmFnt.SIZE = GlTFt.SIZE

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmSta
    .Font.SIZE = 8
    Set CmPan = .AddPane(1)
    CmPan.Text = vbNullString
    CmPan.Width = 100
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Style = SBPS_STRETCH
    .Visible = True
End With

With CmAcs
    Set CmAct = .Add(SY_OP_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
End With

Set TbBar = CmBrs.AddTabToolBar("TabBar")

'------------------------------------------------------------------------------------------

Set ToTab = TbBar.InsertCategory(RibTab_Opti1, "Hauptdaten")
With ToTab
    .ToolTip = "Die Hauptdaten des gewählten Eintrags"
    .Selected = True
    Select Case TreKy
    Case "A": .Visible = True   'Gebührenkataloge
    Case "C": .Visible = True   'Diagnosekatalog
    Case "G": .Visible = True   'Laborparameter
    Case "I": .Visible = True   'Arzneikatalog
    Case "K": .Visible = True   'Begründungen
    Case "L": .Visible = True   'Anamnesetexte
    Case "M": .Visible = True   'Terminbetreffs
    Case "N": .Visible = True   'Fragenkatalog
    Case "P": .Visible = True   'Artikelkatalog
    End Select
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neuer Eintrag")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Legt einen neuen Eintrag an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Speichert den aktuellen Eintrag"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Löscht den aktuellen Eintrag"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
        .Enabled = Not GlKaN
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Hauptdaten"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

'------------------------------------------------------------------------------------------

Set ToTab = TbBar.InsertCategory(RibTab_Opti2, "Kommentar")
With ToTab
    .ToolTip = "Erweiterte Textinformationen, die den Eintrag ergänzen"
    .Selected = False
    Select Case TreKy
    Case "A": .Visible = True   'Gebührenkataloge
    Case "C": .Visible = True   'Diagnosekatalog
    Case "G": .Visible = True   'Laborparameter
    Case "I": .Visible = True   'Arzneikatalog
    Case "K": .Visible = False  'Begründungen
    Case "L": .Visible = False  'Anamnesetexte
    Case "M": .Visible = True  'Terminbetreffs
    Case "N": .Visible = True   'Fragenkatalog
    Case "O": .Visible = True   'Textphrasenkatalog
    Case "P": .Visible = True   'Artikelkatalog
    End Select
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neuer Eintrag")
    With CmCon
        .Category = "Kommentar"
        .ToolTipText = "Legt einen neuen Eintrag an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Kommentar"
        .ToolTipText = "Speichert den aktuellen Eintrag"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Kommentar"
        .ToolTipText = "Löscht den aktuellen Eintrag"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
        .Enabled = Not GlKaN
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Kommentar"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Kommentar"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

'------------------------------------------------------------------------------------------

Set ToTab = TbBar.InsertCategory(RibTab_Opti3, "Regelprüfung")
With ToTab
    .ToolTip = "Diesem Eintrag angehängte Einträge"
    .Selected = False
    Select Case TreKy
    Case "A": .Visible = True   'Gebührenkataloge
    Case "C": .Visible = False  'Diagnosekatalog
    Case "G": .Visible = False  'Laborparameter
    Case "I": .Visible = False  'Arzneikatalog
    Case "K": .Visible = False  'Begründungen
    Case "L": .Visible = False  'Anamnesetexte
    Case "M": .Visible = False  'Terminbetreffs
    Case "N": .Visible = False  'Fragenkatalog
    Case "O": .Visible = False  'Textphrasenkatalog
    Case "P": .Visible = False  'Artikelkatalog
    End Select
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neuer Eintrag")
    With CmCon
        .Category = "Regelprüfung"
        .ToolTipText = "Legt einen neuen Eintrag an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Regelprüfung"
        .ToolTipText = "Speichert den aktuellen Eintrag"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlSplitButtonPopup, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Regelprüfung"
        .ToolTipText = "Löscht den aktuellen Eintrag"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
        .Enabled = Not GlKaN
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_Loeschen, "Eintrag Entfernen")
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_SubDe1, "Regeleintrag Entfernen")
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Regelprüfung"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Regelprüfung"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

'------------------------------------------------------------------------------------------

Set ToTab = TbBar.InsertCategory(RibTab_Opti4, "Labordaten")
With ToTab
    .ToolTip = "Spezielle Daten für Laborparameter"
    .Selected = False
    Select Case TreKy
    Case "A": .Visible = False  'Gebührenkataloge
    Case "C": .Visible = False  'Diagnosekatalog
    Case "G": .Visible = True   'Laborparameter
    Case "I": .Visible = False  'Arzneikatalog
    Case "K": .Visible = False  'Begründungen
    Case "L": .Visible = False  'Anamnesetexte
    Case "M": .Visible = False  'Terminbetreffs
    Case "N": .Visible = False  'Fragenkatalog
    Case "O": .Visible = False  'Textphrasenkatalog
    Case "P": .Visible = False  'Artikelkatalog
    End Select
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neuer Eintrag")
    With CmCon
        .Category = "Labordaten"
        .ToolTipText = "Legt einen neuen Eintrag an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Labordaten"
        .ToolTipText = "Speichert den aktuellen Eintrag"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Labordaten"
        .ToolTipText = "Löscht den aktuellen Eintrag"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
        .Enabled = Not GlKaN
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Labordaten"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Labordaten"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

'------------------------------------------------------------------------------------------

Set ToTab = TbBar.InsertCategory(RibTab_Opti5, "Grenzwerte")
With ToTab
    .ToolTip = "Grenzwerte für Laborparameter"
    .Selected = False
    Select Case TreKy
    Case "A": .Visible = False  'Gebührenkataloge
    Case "C": .Visible = False  'Diagnosekatalog
    Case "G": .Visible = True   'Laborparameter
    Case "I": .Visible = False  'Arzneikatalog
    Case "K": .Visible = False  'Begründungen
    Case "L": .Visible = False  'Anamnesetexte
    Case "M": .Visible = False  'Terminbetreffs
    Case "N": .Visible = False  'Fragenkatalog
    Case "O": .Visible = False  'Textphrasenkatalog
    Case "P": .Visible = False  'Artikelkatalog
    End Select
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neuer Eintrag")
    With CmCon
        .Category = "Grenzwerte"
        .ToolTipText = "Legt einen neuen Eintrag an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Grenzwerte"
        .ToolTipText = "Speichert den aktuellen Eintrag"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Grenzwerte"
        .ToolTipText = "Löscht den aktuellen Eintrag"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
        .Enabled = Not GlKaN
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Grenzwerte"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Grenzwerte"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

'------------------------------------------------------------------------------------------

Set ToTab = TbBar.InsertCategory(RibTab_Opti6, "Zuordungen")
With ToTab
    .ToolTip = "Diesem Eintrag angehängte Einträge"
    .Selected = False
    Select Case TreKy
    Case "A": .Visible = True   'Gebührenkataloge
    Case "C": .Visible = False  'Diagnosekatalog
    Case "G": .Visible = False  'Laborparameter
    Case "I": .Visible = False  'Arzneikatalog
    Case "K": .Visible = False  'Begründungen
    Case "L": .Visible = False  'Anamnesetexte
    Case "M": .Visible = False  'Terminbetreffs
    Case "N": .Visible = False  'Fragenkatalog
    Case "O": .Visible = False  'Textphrasenkatalog
    Case "P": .Visible = False  'Artikelkatalog
    End Select
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neuer Eintrag")
    With CmCon
        .Category = "Zuordungen"
        .ToolTipText = "Legt einen neuen Eintrag an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Zuordungen"
        .ToolTipText = "Speichert den aktuellen Eintrag"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlSplitButtonPopup, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Zuordungen"
        .ToolTipText = "Löscht den aktuellen Eintrag"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
        .Enabled = Not GlKaN
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_Loeschen, "Eintrag Entfernen")
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_SubDe2, "Zuordnung Eentfernen")
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Zuordungen"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Zuordungen"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

'------------------------------------------------------------------------------------------

Set ToTab = TbBar.InsertCategory(RibTab_Opti7, "Lagerbestände")
With ToTab
    .ToolTip = "Mengen des Warenlagers bearbeiten"
    .Selected = False
    Select Case TreKy
    Case "A": .Visible = False   'Gebührenkataloge
    Case "C": .Visible = False  'Diagnosekatalog
    Case "G": .Visible = False  'Laborparameter
    Case "I": .Visible = False  'Arzneikatalog
    Case "K": .Visible = False  'Begründungen
    Case "L": .Visible = False  'Anamnesetexte
    Case "M": .Visible = False  'Terminbetreffs
    Case "N": .Visible = False  'Fragenkatalog
    Case "O": .Visible = False  'Textphrasenkatalog
    Case "P": .Visible = True   'Artikelkatalog
    End Select
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neuer Eintrag")
    With CmCon
        .Category = "Lagerbestände"
        .ToolTipText = "Legt einen neuen Eintrag an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Lagerbestände"
        .ToolTipText = "Speichert den aktuellen Eintrag"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Lagerbestände"
        .ToolTipText = "Löscht den aktuellen Eintrag"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
        .Enabled = Not GlKaN
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Lagerbestände"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Lagerbestände"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

'------------------------------------------------------------------------------------------

For Each CmBar In CmBrs
    If CmBar.Type = xtpBarTypeNormal Then
        Set CmCoS = CmBar.Controls
        For Each CmCon In CmCoS
            CmCon.Style = xtpButtonIconAndCaption
        Next CmCon
    End If
Next CmBar

'------------------------------------------------------------------------------------------

With RpCo1
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = False
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
    .PaintManager.ThemedInplaceButtons = True
    .PaintManager.HorizontalGridStyle = xtpGridNoLines
    .PaintManager.VerticalGridStyle = xtpGridNoLines
    If GlGZe = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
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
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = False
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
    .PaintManager.ThemedInplaceButtons = True
    .PaintManager.HorizontalGridStyle = xtpGridNoLines
    .PaintManager.VerticalGridStyle = xtpGridNoLines
    If GlGZe = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
    .PaintManager.GridLineColor = GlGrC
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

With CmZif
    .AutoComplete = False
    .DropDownItemCount = 20
    .DropDownWidth = 4400
End With

With CmBez
    .AutoComplete = True
    .DropDownItemCount = 20
End With

TxZei.Pattern = "\d*"
With TxRs1
    .Pattern = "\d*"
    .SetMask "0", "_"
End With
With TxRs2
    .Pattern = "\d*"
    .SetMask "0", "_"
End With

If TreKy = "I" Then
    TxZus.Font.Name = GlRFt.Name
    TxZus.Font.SIZE = GlRFt.SIZE
End If

'------------------------------------------------------------------------------------------

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
    .TabPaintManager.FixedTabWidth = 100
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

With CoPic
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceOffice2007
    End Select
    .ControlToolTip = "Terminfarbe auswählen"
    .DefaultColor = vbWhite
    .ShowAutomaticColor = False
    .ShowMoreColors = True
End With

DoEvents
CmBrs.RecalcLayout
DoEvents
CmBrs.PaintManager.RefreshMetrics

FM.txtDaBes.SetMask "00.00.0000", "__.__.____"
FM.txtDaBel.SetMask "00.00.0000", "__.__.____"
FM.txtDaEin.SetMask "00.00.0000", "__.__.____"

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Rahm6.BackColor = GlBak
Rahm7.BackColor = GlBak
Rahm8.BackColor = GlBak
OpRS1.BackColor = GlBak
OpRS2.BackColor = GlBak
OpRS3.BackColor = GlBak
FM.chkFavor.BackColor = GlBak
FM.chkFakPr.BackColor = GlBak
FM.chkAnalo.BackColor = GlBak
FM.chkKette.BackColor = GlBak
FM.chkHochs.BackColor = GlBak
FM.chkProfi.BackColor = GlBak

If TreKy = "M" Then
    Rahm8.Visible = True
    Rahm1.Visible = False
End If

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FTabu(ByVal TaIdx As Long)
On Error GoTo AnErr

Dim IdxNr As Long
Dim TreKy As String
Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmKaEdit
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set RpCo2 = FM.repCont2

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

TreKy = Left$(GlNod, 1)

If FM.txtIdxNr.Text <> vbNullString Then
    If IsNumeric(FM.txtIdxNr.Text) = True Then
        If Val(FM.txtIdxNr.Text) > 0 Then
            IdxNr = CLng(FM.txtIdxNr.Text)
        Else
            IdxNr = 0
        End If
    Else
        IdxNr = 0
    End If
Else
    IdxNr = 0
End If

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If TreKy = "M" Then
    Select Case TaIdx
    Case 0:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = False
        Rahm8.Visible = True
    Case 1:
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = False
        Rahm8.Visible = False
    End Select
Else
    Select Case TaIdx
    Case 0:
        Rahm1.Visible = True
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = False
    Case 1:
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = False
    Case 2:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = True
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = False
    Case 3:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = True
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = False
    Case 4:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = True
        Rahm6.Visible = False
        Rahm7.Visible = False
    Case 5:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = True
        Rahm7.Visible = False
        If RpCo2.Rows.Count = 0 Then
            K_AnVo IdxNr
        End If
    Case 6:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        Rahm6.Visible = False
        Rahm7.Visible = True
    End Select
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub

Private Sub chkFakPr_Click()
On Error Resume Next

Set FakPr = Me.chkFakPr
Set TxPr1 = Me.txtPrei1
Set TxPr2 = Me.txtPrei2
Set LbPr2 = Me.lblLabl2
Set TxFak = Me.txtMulti
Set LbPr4 = Me.lblLabl4

If GlKal = False Then
    If FakPr.Value = xtpChecked Then
        LbPr2.Caption = "Einzelpreis :"
        LbPr4.Caption = "Faktorpreis :"
        If TxPr2.Text <> vbNullString Then
            If IsNumeric(TxPr2.Text) = True Then
                If CDbl(TxPr2.Text) < 0 Then
                    TxPr2.Text = Abs(CDbl(TxPr2.Text))
                End If
            Else
                GeDia = TxPr2.Text
                If FaPre > 0 Then
                    TxPr2.Text = Format$(FaPre, GlWa1)
                Else
                    If TxPr1.Text <> vbNullString Then
                        If IsNumeric(TxPr1.Text) = True Then
                            If TxFak.Text <> vbNullString Then
                                If IsNumeric(TxFak.Text) = True Then
                                    TxPr2.Text = Format$(WinRound(TxPr1.Text * TxFak.Text, 2), GlWa1)
                                Else
                                    TxPr2.Text = Format$(0, GlWa1)
                                End If
                            Else
                                TxPr2.Text = Format$(0, GlWa1)
                            End If
                        Else
                            TxPr2.Text = Format$(0, GlWa1)
                        End If
                    Else
                        TxPr2.Text = Format$(0, GlWa1)
                    End If
                End If
            End If
        Else
            If FaPre > 0 Then
                TxPr2.Text = Format$(FaPre, GlWa1)
            Else
                If TxPr1.Text <> vbNullString Then
                    If IsNumeric(TxPr1.Text) = True Then
                        If TxFak.Text <> vbNullString Then
                            If IsNumeric(TxFak.Text) = True Then
                                TxPr2.Text = Format$(WinRound(TxPr1.Text * TxFak.Text, 2), GlWa1)
                            Else
                                TxPr2.Text = Format$(0, GlWa1)
                            End If
                        Else
                            TxPr2.Text = Format$(0, GlWa1)
                        End If
                    Else
                        TxPr2.Text = Format$(0, GlWa1)
                    End If
                Else
                    TxPr2.Text = Format$(0, GlWa1)
                End If
            End If
        End If
    Else
        If IsNull(TxPr2.Text) = False Then
            If IsNumeric(TxPr2.Text) = True Then
                FaPre = TxPr2.Text
            End If
        End If
        LbPr2.Caption = "Preis :"
        LbPr4.Caption = "Diagnose :"
        If GeDia <> vbNullString Then
            TxPr2.Text = GeDia
        Else
            TxPr2.Text = vbNullString
        End If
    End If
End If

End Sub

Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlKal = False Then FTool Control.id
End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlKal = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    KaPos
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: KaRes
Case KY_F8: K_Save
Case KY_F11: Unload Me
Case SY_OP_Hilfe: FHilfe
Case SY_OP_Hinzufuegen: KaRes
Case SY_OP_Speichern: K_Save
Case SY_OP_Loeschen: K_Loe
                     Unload Me
Case SY_OP_Abbruch: Unload Me
Case SY_OP_SubDe1: K_AnLo
Case SY_OP_SubDe2: K_AnLe
End Select

GlToo = False

End Sub
Private Sub Form_Activate()
    KaPos
End Sub
Private Sub Form_Load()
    
Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 6000
    .ClientMaxWidth = 8200
    .ClientMinHeight = 4800
    .ClientMinWidth = 6700
    .TopMost = True
End With

Set FrmEx = Nothing

AFont Me
FInit

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmKaEdit = Nothing
End Sub

Private Sub optRest1_Click()
On Error Resume Next

Set TxRs1 = Me.txtRest1
Set TxRs2 = Me.txtRest2

TxRs1.Enabled = True
TxRs2.Enabled = False

End Sub

Private Sub optRest2_Click()
On Error Resume Next

Set TxRs1 = Me.txtRest1
Set TxRs2 = Me.txtRest2

TxRs1.Enabled = False
TxRs2.Enabled = True

End Sub

Private Sub optRest3_Click()
On Error Resume Next

Set TxRs1 = Me.txtRest1
Set TxRs2 = Me.txtRest2

TxRs1.Enabled = False
TxRs2.Enabled = False

End Sub

Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyDelete Then
            K_AnLo
        End If
    End If
End Sub

Private Sub repCont2_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim AktZa As Integer

If GlADi > 0 Then
    If GlDiy(Item.Index, Row.Index) <> vbNullString Then
        Metrics.Text = GlDiy(Item.Index, Row.Index)
    End If
End If

End Sub
Private Sub TbBar_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    FTabu Item.Index
End Sub

Private Sub txtBeIst_GotFocus()
    Me.txtBeIst.SelStart = 0
    Me.txtBeIst.SelLength = Len(Me.txtBeIst.Text)
End Sub
Private Sub txtBeMax_GotFocus()
    Me.txtBeMax.SelStart = 0
    Me.txtBeMax.SelLength = Len(Me.txtBeMax.Text)
End Sub

Private Sub txtBeMel_GotFocus()
    Me.txtBeMel.SelStart = 0
    Me.txtBeMel.SelLength = Len(Me.txtBeMel.Text)
End Sub
Private Sub txtBeMin_GotFocus()
    Me.txtBeMin.SelStart = 0
    Me.txtBeMin.SelLength = Len(Me.txtBeMin.Text)
End Sub
Private Sub txtBeSol_GotFocus()
    Me.txtBeSol.SelStart = 0
    Me.txtBeSol.SelLength = Len(Me.txtBeSol.Text)
End Sub

Private Sub txtBetre_GotFocus()
    Me.txtBetre.SelStart = 0
    Me.txtBetre.SelLength = Len(Me.txtBetre.Text)
End Sub
Private Sub txtBezei_GotFocus()
    Me.txtBezei.SelStart = 0
    Me.txtBezei.SelLength = Len(Me.txtBezei.Text)
End Sub

Private Sub txtBezei_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtDaBel_GotFocus()
    Me.txtDaBel.SelStart = 0
    Me.txtDaBel.SelLength = Len(Me.txtDaBel.Text)
End Sub
Private Sub txtDaBes_GotFocus()
    Me.txtDaBes.SelStart = 0
    Me.txtDaBes.SelLength = Len(Me.txtDaBes.Text)
End Sub

Private Sub txtDaEin_GotFocus()
    Me.txtDaEin.SelStart = 0
    Me.txtDaEin.SelLength = Len(Me.txtDaEin.Text)
End Sub

Private Sub txtKomme_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtLaOrt_GotFocus()
    Me.txtLaOrt.SelStart = 0
    Me.txtLaOrt.SelLength = Len(Me.txtLaOrt.Text)
End Sub
Private Sub txtMeBes_GotFocus()
    Me.txtMeBes.SelStart = 0
    Me.txtMeBes.SelLength = Len(Me.txtMeBes.Text)
End Sub

Private Sub txtMeEin_GotFocus()
    Me.txtMeEin.SelStart = 0
    Me.txtMeEin.SelLength = Len(Me.txtMeEin.Text)
End Sub
Private Sub txtMinut_GotFocus()
    Me.txtMinut.SelStart = 0
    Me.txtMinut.SelLength = Len(Me.txtMinut.Text)
End Sub

Private Sub txtMnute_GotFocus()
    Me.txtMnute.SelStart = 0
    Me.txtMnute.SelLength = Len(Me.txtMnute.Text)
End Sub
Private Sub txtMulti_GotFocus()
    Me.txtMulti.SelStart = 0
    Me.txtMulti.SelLength = Len(Me.txtMulti.Text)
End Sub

Private Sub txtNachl_GotFocus()
    Me.txtNachl.SelStart = 0
    Me.txtNachl.SelLength = Len(Me.txtNachl.Text)
End Sub
Private Sub txtPrei1_GotFocus()
    Me.txtPrei1.SelStart = 0
    Me.txtPrei1.SelLength = Len(Me.txtPrei1.Text)
End Sub
Private Sub txtPrei1_LostFocus()
On Error Resume Next

Set TxPr1 = Me.txtPrei1

If IsNull(TxPr1.Text) = False Then
    If IsNumeric(TxPr1.Text) = True Then
        If CDbl(TxPr1.Text) < 0 Then
            TxPr1.Text = Abs(CDbl(TxPr1.Text))
        End If
    Else
        TxPr1.Text = Format$(0, GlWa1)
    End If
End If

End Sub
Private Sub txtPrei2_GotFocus()
    Me.txtPrei2.SelStart = 0
    Me.txtPrei2.SelLength = Len(Me.txtPrei2.Text)
End Sub

Private Sub txtPrei2_LostFocus()
On Error Resume Next

Set FakPr = Me.chkFakPr
Set TxPr2 = Me.txtPrei2

If FakPr.Value = xtpChecked Then
    If IsNull(TxPr2.Text) = False Then
        If IsNumeric(TxPr2.Text) = True Then
            If CDbl(TxPr2.Text) < 0 Then
                TxPr2.Text = Abs(CDbl(TxPr2.Text))
            End If
        Else
            TxPr2.Text = Format$(0, GlWa1)
        End If
    End If
End If

End Sub

Private Sub txtRest1_GotFocus()
    Me.txtRest1.SelStart = 0
    Me.txtRest1.SelLength = Len(Me.txtRest1.Text)
End Sub
Private Sub txtSorte_GotFocus()
    Me.txtSorte.SelStart = 0
    Me.txtSorte.SelLength = Len(Me.txtSorte.Text)
End Sub

Private Sub txtSteue_GotFocus()
    Me.txtSteue.SelStart = 0
    Me.txtSteue.SelLength = Len(Me.txtSteue.Text)
End Sub
Private Sub txtZiff1_GotFocus()
    Me.txtZiff1.SelStart = 0
    Me.txtZiff1.SelLength = Len(Me.txtZiff1.Text)
End Sub
Private Sub txtRest2_GotFocus()
    Me.txtRest2.SelStart = 0
    Me.txtRest2.SelLength = Len(Me.txtRest2.Text)
End Sub
Private Sub cmbZiffe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        K_AnEi 1
    End If
End Sub
Private Sub cmbBezei_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        K_AnEi 2
    End If
End Sub

Private Sub txtZusat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub


