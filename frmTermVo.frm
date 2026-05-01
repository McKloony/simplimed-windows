VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmTermVo 
   Caption         =   "Terminvorschlag"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   ControlBox      =   0   'False
   Icon            =   "frmTermVo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   14280
   Begin XtremeReportControl.ReportControl repCont6 
      Height          =   1500
      Left            =   3240
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   2646
      _StockProps     =   64
      BorderStyle     =   3
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1500
      Left            =   1320
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   2646
      _StockProps     =   64
      BorderStyle     =   3
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu2 
      Height          =   400
      Left            =   1080
      TabIndex        =   97
      Top             =   100
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm9 
      Height          =   4700
      Left            =   100
      TabIndex        =   1
      Top             =   960
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   8290
      _StockProps     =   79
      Caption         =   "Termindaten"
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown updCont3 
         Height          =   350
         Left            =   6510
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1300
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBisZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   350
         Left            =   5120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1300
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtVonZe"
         BuddyProperty   =   ""
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu1 
         Height          =   400
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   400
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   706
         _StockProps     =   64
         Show3DBorder    =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRaum1 
         Height          =   195
         Left            =   2800
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "0Raum"
         Top             =   1800
         Visible         =   0   'False
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAdres 
         Height          =   350
         Left            =   1100
         TabIndex        =   2
         Tag             =   "0Patient"
         Top             =   300
         Width           =   6000
         _Version        =   1048579
         _ExtentX        =   10583
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   2790
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   1300
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox txtBetre 
         Height          =   310
         Left            =   1100
         TabIndex        =   3
         Tag             =   "0IDKurz"
         Top             =   800
         Width           =   6000
         _Version        =   1048579
         _ExtentX        =   10583
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         MaxLength       =   250
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbStatu 
         Height          =   310
         Left            =   1100
         TabIndex        =   13
         Tag             =   "0Farbtyp"
         Top             =   2300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3545
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox cmbRaum1 
         Height          =   315
         Left            =   4200
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "0IDR"
         Top             =   1800
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox6"
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   310
         Left            =   4200
         TabIndex        =   16
         Tag             =   "0IDP"
         Top             =   2800
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox7"
      End
      Begin XtremeSuiteControls.ComboBox cmbRemin 
         Height          =   310
         Left            =   1100
         TabIndex        =   15
         Tag             =   "0Vorwarn"
         Top             =   2800
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3545
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox8"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   1100
         TabIndex        =   4
         Tag             =   "0VonDat"
         Top             =   1300
         Width           =   1660
         _Version        =   1048579
         _ExtentX        =   2928
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtVonZe 
         Height          =   350
         Left            =   4200
         TabIndex        =   6
         Tag             =   "0ZeiVon"
         Top             =   1300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBisZe 
         Height          =   350
         Left            =   5600
         TabIndex        =   8
         Tag             =   "0ZeiBis"
         Top             =   1300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu5 
         Height          =   350
         Left            =   1100
         TabIndex        =   10
         Top             =   1800
         Width           =   1660
         _Version        =   1048579
         _ExtentX        =   2928
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbTeTyp 
         Height          =   315
         Left            =   4200
         TabIndex        =   14
         Tag             =   "0TerTyp"
         Top             =   2300
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   4200
         TabIndex        =   18
         Tag             =   "0IDM"
         Top             =   3300
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox7"
      End
      Begin XtremeSuiteControls.ComboBox cmbGanzt 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   3800
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtKomme 
         Height          =   350
         Left            =   1100
         TabIndex        =   21
         Tag             =   "0Kommentar"
         Top             =   4300
         Width           =   6000
         _Version        =   1048579
         _ExtentX        =   10583
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   250
      End
      Begin XtremeSuiteControls.ComboBox cmbNotVa 
         Height          =   315
         Left            =   1100
         TabIndex        =   17
         Tag             =   "0NotifyValue"
         Top             =   3300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3545
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbAbger 
         Height          =   315
         Left            =   4200
         TabIndex        =   20
         Tag             =   "0Aufgabe"
         Top             =   3800
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.Label lblLab23 
         Height          =   240
         Left            =   3200
         TabIndex        =   110
         Top             =   1350
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Terminzeit :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab22 
         Height          =   240
         Left            =   3200
         TabIndex        =   109
         Top             =   3850
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Leistungen :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab21 
         Height          =   240
         Left            =   140
         TabIndex        =   108
         Top             =   3350
         Width           =   920
         _Version        =   1048579
         _ExtentX        =   1614
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Emailerinn.: "
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   240
         Left            =   3200
         TabIndex        =   107
         Top             =   1850
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Raumplan :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab18 
         Height          =   240
         Left            =   3200
         TabIndex        =   99
         Top             =   2850
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Mandant :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab04 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Faktisch :"
         Height          =   240
         Left            =   140
         TabIndex        =   95
         Top             =   1850
         Width           =   920
      End
      Begin VB.Label lblLab14 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Marker :"
         Height          =   240
         Left            =   3200
         TabIndex        =   87
         Top             =   2350
         Width           =   930
      End
      Begin VB.Label lblLab12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar :"
         Height          =   240
         Left            =   140
         TabIndex        =   86
         Top             =   4350
         Width           =   920
      End
      Begin VB.Label lblLab15 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   240
         Left            =   140
         TabIndex        =   85
         Top             =   2350
         Width           =   920
      End
      Begin VB.Label lblLab10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   240
         Left            =   3200
         TabIndex        =   84
         Top             =   3350
         Width           =   930
      End
      Begin VB.Label lblLab09 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Ganztägig :"
         Height          =   240
         Left            =   140
         TabIndex        =   83
         Top             =   3850
         Width           =   920
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Betreff :"
         Height          =   240
         Left            =   140
         TabIndex        =   82
         Top             =   850
         Width           =   920
      End
      Begin VB.Label lblLab01 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Patient :"
         Height          =   240
         Left            =   140
         TabIndex        =   81
         Top             =   340
         Width           =   920
      End
      Begin VB.Label lblLab03 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Terminstart :"
         Height          =   240
         Left            =   140
         TabIndex        =   80
         Top             =   1350
         Width           =   920
      End
      Begin VB.Label lblLab13 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Erinnerung :"
         Height          =   240
         Left            =   140
         TabIndex        =   79
         Top             =   2850
         Width           =   920
      End
   End
   Begin XtremeSuiteControls.FlatEdit txoDummy 
      Height          =   200
      Left            =   200
      TabIndex        =   0
      Top             =   13000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtID2 
      Height          =   195
      Left            =   1200
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "0ID2"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtFarbe 
      Height          =   195
      Left            =   600
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "0Farbe"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtID0 
      Height          =   195
      Left            =   1600
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "0ID0"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtIdSer 
      Height          =   195
      Left            =   2000
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "0IDSer"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtMaTer 
      Height          =   195
      Left            =   2400
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "0MasTer"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtSeTyp 
      Height          =   195
      Left            =   2760
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "0SerTyp"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm8 
      Height          =   4700
      Left            =   8000
      TabIndex        =   23
      Top             =   960
      Width           =   7200
      _Version        =   1048579
      _ExtentX        =   12700
      _ExtentY        =   8290
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox frmRahm6 
         Height          =   840
         Left            =   100
         TabIndex        =   32
         Top             =   3740
         Width           =   7000
         _Version        =   1048579
         _ExtentX        =   12347
         _ExtentY        =   1482
         _StockProps     =   79
         Caption         =   "Seriendauer"
         Appearance      =   12
         Begin XtremeCalendarControl.DatePicker dtpDatu3 
            Height          =   405
            Left            =   6500
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   200
            Visible         =   0   'False
            Width           =   405
            _Version        =   1048579
            _ExtentX        =   706
            _ExtentY        =   706
            _StockProps     =   64
            Show3DBorder    =   2
         End
         Begin XtremeSuiteControls.PushButton btnDatu4 
            Height          =   350
            Left            =   6120
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "Öffnet den Auswahlkalender"
            Top             =   300
            Width           =   350
            _Version        =   1048579
            _ExtentX        =   617
            _ExtentY        =   617
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZyEn3 
            Height          =   220
            Left            =   3900
            TabIndex        =   70
            Top             =   350
            Width           =   900
            _Version        =   1048579
            _ExtentX        =   1587
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "bis zum :"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZyEn2 
            Height          =   220
            Left            =   400
            TabIndex        =   68
            Top             =   350
            Width           =   900
            _Version        =   1048579
            _ExtentX        =   1587
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Anzahl :"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDatu4 
            Height          =   350
            Left            =   4900
            TabIndex        =   71
            Top             =   300
            Width           =   1200
            _Version        =   1048579
            _ExtentX        =   2117
            _ExtentY        =   617
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            BackColor       =   16777215
            Alignment       =   2
         End
         Begin XtremeSuiteControls.ComboBox cmbZyEn1 
            Height          =   310
            Left            =   1300
            TabIndex        =   69
            Top             =   300
            Width           =   1800
            _Version        =   1048579
            _ExtentX        =   3175
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Style           =   2
            Text            =   "ComboBox1"
         End
      End
      Begin XtremeSuiteControls.GroupBox frmRahm5 
         Height          =   1710
         Left            =   100
         TabIndex        =   31
         Top             =   1940
         Width           =   7000
         _Version        =   1048579
         _ExtentX        =   12347
         _ExtentY        =   3016
         _StockProps     =   79
         Caption         =   "Serienmuster"
         Appearance      =   12
         Begin XtremeSuiteControls.GroupBox frmRahm1 
            Height          =   1400
            Left            =   1900
            TabIndex        =   33
            Top             =   120
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2469
            _StockProps     =   79
            Appearance      =   6
            BorderStyle     =   2
            Begin XtremeSuiteControls.FlatEdit txoTage1 
               Height          =   310
               Left            =   870
               TabIndex        =   50
               Top             =   150
               Width           =   500
               _Version        =   1048579
               _ExtentX        =   873
               _ExtentY        =   547
               _StockProps     =   77
               ForeColor       =   4473924
               BackColor       =   16777215
               Text            =   "1"
               BackColor       =   16777215
               Alignment       =   2
            End
            Begin XtremeSuiteControls.RadioButton optZyTa2 
               Height          =   220
               Left            =   150
               TabIndex        =   51
               Top             =   560
               Width           =   1600
               _Version        =   1048579
               _ExtentX        =   2822
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Jeden Arbeitstag"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton optZyTa1 
               Height          =   220
               Left            =   150
               TabIndex        =   49
               Top             =   200
               Width           =   600
               _Version        =   1048579
               _ExtentX        =   1058
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Alle"
               UseVisualStyle  =   -1  'True
            End
            Begin VB.Label lblLab11 
               BackStyle       =   0  'Transparent
               Caption         =   "Tag(e)"
               Height          =   220
               Left            =   1460
               TabIndex        =   90
               Top             =   210
               Width           =   600
            End
         End
         Begin XtremeSuiteControls.GroupBox frmRahm2 
            Height          =   1400
            Left            =   1900
            TabIndex        =   34
            Top             =   120
            Visible         =   0   'False
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2469
            _StockProps     =   79
            Appearance      =   6
            BorderStyle     =   2
            Begin XtremeSuiteControls.CheckBox choTaSam 
               Height          =   220
               Left            =   2600
               TabIndex        =   58
               Top             =   940
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Samstag"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox choTaFre 
               Height          =   220
               Left            =   1400
               TabIndex        =   57
               Top             =   940
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Freitag"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox choTaDon 
               Height          =   220
               Left            =   160
               TabIndex        =   56
               Top             =   940
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Donnerstag"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox choTaSon 
               Height          =   220
               Left            =   3800
               TabIndex        =   89
               Top             =   940
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Sonntag"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox choTaMit 
               Height          =   220
               Left            =   2600
               TabIndex        =   55
               Top             =   600
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Mittwoch"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox choTaDin 
               Height          =   220
               Left            =   1400
               TabIndex        =   54
               Top             =   600
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Dienstag"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox choTaMon 
               Height          =   220
               Left            =   160
               TabIndex        =   53
               Top             =   600
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Montag"
               UseVisualStyle  =   -1  'True
               Value           =   1
            End
            Begin XtremeSuiteControls.ComboBox cmbWoche 
               Height          =   310
               Left            =   160
               TabIndex        =   52
               Top             =   150
               Width           =   1900
               _Version        =   1048579
               _ExtentX        =   3334
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   4473924
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
         End
         Begin XtremeSuiteControls.GroupBox frmRahm3 
            Height          =   1400
            Left            =   1900
            TabIndex        =   35
            Top             =   120
            Visible         =   0   'False
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2469
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            BorderStyle     =   2
            Begin XtremeSuiteControls.RadioButton optZyMo2 
               Height          =   220
               Left            =   150
               TabIndex        =   75
               Top             =   750
               Width           =   600
               _Version        =   1048579
               _ExtentX        =   1058
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Am"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton optZyMo1 
               Height          =   220
               Left            =   150
               TabIndex        =   59
               Top             =   200
               Width           =   600
               _Version        =   1048579
               _ExtentX        =   1058
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Am"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.ComboBox cmoMona1 
               Height          =   310
               Left            =   800
               TabIndex        =   76
               Top             =   700
               Width           =   1000
               _Version        =   1048579
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   4473924
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox5"
            End
            Begin XtremeSuiteControls.ComboBox cmoMona2 
               Height          =   310
               Left            =   1860
               TabIndex        =   77
               Top             =   700
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   4473924
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox6"
            End
            Begin XtremeSuiteControls.ComboBox cmbMonat 
               Height          =   310
               Left            =   1860
               TabIndex        =   74
               Top             =   150
               Width           =   1800
               _Version        =   1048579
               _ExtentX        =   3175
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   4473924
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cmbMona3 
               Height          =   315
               Left            =   3120
               TabIndex        =   78
               Top             =   700
               Width           =   1800
               _Version        =   1048579
               _ExtentX        =   3175
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   4473924
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cmbMona1 
               Height          =   310
               Left            =   800
               TabIndex        =   60
               Top             =   150
               Width           =   1000
               _Version        =   1048579
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   4473924
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
         End
         Begin XtremeSuiteControls.RadioButton optZykl4 
            Height          =   220
            Left            =   400
            TabIndex        =   48
            Top             =   1320
            Width           =   1300
            _Version        =   1048579
            _ExtentX        =   2293
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Jährlich"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZykl3 
            Height          =   220
            Left            =   400
            TabIndex        =   47
            Top             =   980
            Width           =   1300
            _Version        =   1048579
            _ExtentX        =   2293
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Monatlich"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZykl2 
            Height          =   220
            Left            =   400
            TabIndex        =   46
            Top             =   630
            Width           =   1300
            _Version        =   1048579
            _ExtentX        =   2293
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Wöchentlich"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZykl1 
            Height          =   220
            Left            =   400
            TabIndex        =   45
            Top             =   290
            Width           =   1300
            _Version        =   1048579
            _ExtentX        =   2293
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Täglich"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox frmRahm4 
            Height          =   1400
            Left            =   1900
            TabIndex        =   36
            Top             =   120
            Visible         =   0   'False
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2469
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            BorderStyle     =   2
            Begin XtremeSuiteControls.RadioButton optZyJa2 
               Height          =   220
               Left            =   150
               TabIndex        =   64
               Top             =   750
               Width           =   740
               _Version        =   1048579
               _ExtentX        =   1305
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Am"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton optZyJa1 
               Height          =   220
               Left            =   150
               TabIndex        =   61
               Top             =   200
               Width           =   740
               _Version        =   1048579
               _ExtentX        =   1305
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Jeden"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.ComboBox cmoJahr1 
               Height          =   310
               Left            =   2100
               TabIndex        =   63
               Top             =   150
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
            Begin XtremeSuiteControls.ComboBox cmoJahr2 
               Height          =   310
               Left            =   1000
               TabIndex        =   65
               Top             =   700
               Width           =   1000
               _Version        =   1048579
               _ExtentX        =   1746
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox2"
            End
            Begin XtremeSuiteControls.ComboBox cmoJahr3 
               Height          =   310
               Left            =   2100
               TabIndex        =   66
               Top             =   700
               Width           =   1300
               _Version        =   1048579
               _ExtentX        =   2302
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox3"
            End
            Begin XtremeSuiteControls.ComboBox cmoJahr4 
               Height          =   310
               Left            =   3735
               TabIndex        =   67
               Top             =   700
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox4"
            End
            Begin XtremeSuiteControls.ComboBox cmbJahr1 
               Height          =   310
               Left            =   1000
               TabIndex        =   62
               Top             =   150
               Width           =   1000
               _Version        =   1048579
               _ExtentX        =   1746
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin VB.Label lblLab20 
               BackStyle       =   0  'Transparent
               Caption         =   "im"
               Height          =   220
               Left            =   3500
               TabIndex        =   91
               Top             =   720
               Width           =   200
            End
         End
      End
      Begin XtremeSuiteControls.GroupBox frmRahm7 
         Height          =   1860
         Left            =   100
         TabIndex        =   30
         Top             =   0
         Width           =   7000
         _Version        =   1048579
         _ExtentX        =   12347
         _ExtentY        =   3281
         _StockProps     =   79
         Caption         =   "Serienberechnung"
         Appearance      =   12
         Begin XtremeSuiteControls.CheckBox chkTeSpl 
            Height          =   220
            Left            =   400
            TabIndex        =   40
            Top             =   1160
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Jeden Termin in zwei Teile aufsplitten"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkDopTe 
            Height          =   220
            Left            =   400
            TabIndex        =   41
            Top             =   1440
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Zwei Termine pro Tag berechnen"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox txoZwZei 
            Height          =   315
            Left            =   5400
            TabIndex        =   44
            Top             =   1400
            Width           =   1000
            _Version        =   1048579
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Enabled         =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkFreTe 
            Height          =   220
            Left            =   400
            TabIndex        =   37
            Top             =   320
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Belegte Termine berücksichtigen"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkRauZu 
            Height          =   220
            Left            =   400
            TabIndex        =   39
            Top             =   890
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Raumzuordnungen berücksichtigen"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox txoSplAn 
            Height          =   315
            Left            =   5400
            TabIndex        =   42
            Top             =   280
            Width           =   1000
            _Version        =   1048579
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Enabled         =   0   'False
            Style           =   2
         End
         Begin XtremeSuiteControls.ComboBox txoSplPa 
            Height          =   315
            Left            =   5400
            TabIndex        =   43
            Top             =   840
            Width           =   1000
            _Version        =   1048579
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Enabled         =   0   'False
            Style           =   2
         End
         Begin XtremeSuiteControls.CheckBox chkSprZe 
            Height          =   220
            Left            =   400
            TabIndex        =   38
            Top             =   600
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Sprechzeiten berücksichtigen"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblLab17 
            Height          =   240
            Left            =   3600
            TabIndex        =   94
            Top             =   320
            Width           =   1700
            _Version        =   1048579
            _ExtentX        =   2999
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Anzahl der Teiltermine :"
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblLab16 
            Height          =   240
            Left            =   3600
            TabIndex        =   93
            Top             =   890
            Width           =   1700
            _Version        =   1048579
            _ExtentX        =   2999
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Pause für Teiltermine :"
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblLab19 
            Height          =   240
            Left            =   3600
            TabIndex        =   92
            Top             =   1440
            Width           =   1700
            _Version        =   1048579
            _ExtentX        =   2999
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Beginn Zweittermin :"
            Alignment       =   1
            Transparent     =   -1  'True
         End
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtPaTel 
      Height          =   195
      Left            =   3200
      TabIndex        =   98
      TabStop         =   0   'False
      Tag             =   "0Datei"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkWiede 
      CausesValidation=   0   'False
      Height          =   220
      Left            =   5400
      TabIndex        =   100
      TabStop         =   0   'False
      Tag             =   "0Wiederholung"
      Top             =   13000
      Width           =   220
      _Version        =   1048579
      _ExtentX        =   388
      _ExtentY        =   388
      _StockProps     =   79
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtFall1 
      Height          =   195
      Left            =   3600
      TabIndex        =   101
      TabStop         =   0   'False
      Tag             =   "0Fallig1"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtFall2 
      Height          =   195
      Left            =   4000
      TabIndex        =   102
      TabStop         =   0   'False
      Tag             =   "0Fallig2"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtBetra 
      Height          =   195
      Left            =   4400
      TabIndex        =   103
      TabStop         =   0   'False
      Tag             =   "0GesBetrag"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtRefNr 
      Height          =   195
      Left            =   4800
      TabIndex        =   104
      TabStop         =   0   'False
      Tag             =   "0MasTer"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtBehin 
      Height          =   195
      Left            =   6000
      TabIndex        =   105
      TabStop         =   0   'False
      Tag             =   "0Behindert"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbGesch 
      Height          =   315
      Left            =   6600
      TabIndex        =   106
      TabStop         =   0   'False
      Tag             =   "0Geschlecht"
      Top             =   13000
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   609
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtS4F20 
      Height          =   195
      Left            =   5100
      TabIndex        =   111
      TabStop         =   0   'False
      Tag             =   "0Behindert"
      Top             =   13000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   480
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
Attribute VB_Name = "frmTermVo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Rahm7 As XtremeSuiteControls.GroupBox
Private Rahm8 As XtremeSuiteControls.GroupBox
Private Rahm9 As XtremeSuiteControls.GroupBox
Private TxID0 As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa4 As XtremeSuiteControls.FlatEdit
Private TxKom As XtremeSuiteControls.FlatEdit
Private TxID2 As XtremeSuiteControls.FlatEdit
Private TxFar As XtremeSuiteControls.FlatEdit
Private TxAdr As XtremeSuiteControls.FlatEdit
Private TxMas As XtremeSuiteControls.FlatEdit
Private TxSTy As XtremeSuiteControls.FlatEdit
Private VoZei As XtremeSuiteControls.FlatEdit
Private BiZei As XtremeSuiteControls.FlatEdit
Private ZyTag As XtremeSuiteControls.FlatEdit
Private TxBeG As XtremeSuiteControls.FlatEdit
Private TxNoD As XtremeSuiteControls.FlatEdit
Private TxNoZ As XtremeSuiteControls.FlatEdit
Private ChBsp As XtremeSuiteControls.CheckBox
Private ChGnz As XtremeSuiteControls.CheckBox
Private ChPrv As XtremeSuiteControls.CheckBox
Private ChTer As XtremeSuiteControls.CheckBox
Private ChRau As XtremeSuiteControls.CheckBox
Private ChSpr As XtremeSuiteControls.CheckBox
Private ChMon As XtremeSuiteControls.CheckBox
Private ChDin As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private ChDon As XtremeSuiteControls.CheckBox
Private ChFre As XtremeSuiteControls.CheckBox
Private ChSam As XtremeSuiteControls.CheckBox
Private ChSon As XtremeSuiteControls.CheckBox
Private ChRmu As XtremeSuiteControls.CheckBox
Private ChDop As XtremeSuiteControls.CheckBox
Private ChSpl As XtremeSuiteControls.CheckBox
Private CmRem As XtremeSuiteControls.ComboBox
Private CmBet As XtremeSuiteControls.ComboBox
Private CmRmu As XtremeSuiteControls.ComboBox
Private CmNot As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmPri As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private ZwZei As XtremeSuiteControls.ComboBox
Private ZyWoh As XtremeSuiteControls.ComboBox
Private ZyMo1 As XtremeSuiteControls.ComboBox
Private ZyMo2 As XtremeSuiteControls.ComboBox
Private ZyMo3 As XtremeSuiteControls.ComboBox
Private ZyMo4 As XtremeSuiteControls.ComboBox
Private ZyMoT As XtremeSuiteControls.ComboBox
Private ZyJa1 As XtremeSuiteControls.ComboBox
Private ZyJa2 As XtremeSuiteControls.ComboBox
Private ZyJa3 As XtremeSuiteControls.ComboBox
Private ZyJa4 As XtremeSuiteControls.ComboBox
Private ZyJaT As XtremeSuiteControls.ComboBox
Private ZyEnT As XtremeSuiteControls.ComboBox
Private ZyWho As XtremeSuiteControls.ComboBox
Private ZyMe1 As XtremeSuiteControls.ComboBox
Private ZyMe2 As XtremeSuiteControls.ComboBox
Private ZyMe3 As XtremeSuiteControls.ComboBox
Private ZyJe1 As XtremeSuiteControls.ComboBox
Private ZyTer As XtremeSuiteControls.ComboBox
Private TxSp1 As XtremeSuiteControls.ComboBox
Private TxSp2 As XtremeSuiteControls.ComboBox
Private OpZy1 As XtremeSuiteControls.RadioButton
Private OpZy2 As XtremeSuiteControls.RadioButton
Private OpZy3 As XtremeSuiteControls.RadioButton
Private OpZy4 As XtremeSuiteControls.RadioButton
Private FoZy1 As XtremeSuiteControls.RadioButton
Private FoZy2 As XtremeSuiteControls.RadioButton
Private FoZy3 As XtremeSuiteControls.RadioButton
Private FoZy4 As XtremeSuiteControls.RadioButton
Private ZyEn2 As XtremeSuiteControls.RadioButton
Private ZyEn3 As XtremeSuiteControls.RadioButton
Private TaZy1 As XtremeSuiteControls.RadioButton
Private TaZy2 As XtremeSuiteControls.RadioButton
Private MoZy1 As XtremeSuiteControls.RadioButton
Private MoZy2 As XtremeSuiteControls.RadioButton
Private JaZy1 As XtremeSuiteControls.RadioButton
Private JaZy2 As XtremeSuiteControls.RadioButton

Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions

Private TabCo As XtremeSuiteControls.TabControl
Private TabIt As XtremeSuiteControls.TabControlItem
Private CaCol As XtremeCalendarControl.CalendarControl
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private ChRow As XtremeReportControl.ReportRow
Private ChRws As XtremeReportControl.ReportRows
Private MoKa1 As XtremeCalendarControl.DatePicker
Private MoKa2 As XtremeCalendarControl.DatePicker
Private MoKa3 As XtremeCalendarControl.DatePicker
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private RetWe As Long
Private TagWe As String
Private KalWa As Integer

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

Private clFil As clsFile
Private clWor As clsWord
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
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date
Dim AnzVo As Integer

Set ZyEnT = Me.cmbZyEn1
Set TxDa1 = Me.txtDatu1
Set TxDa4 = Me.txtDatu4
Set MoKa1 = Me.dtpDatu1
Set MoKa3 = Me.dtpDatu3

AnzVo = ZyEnT.ItemData(ZyEnT.ListIndex)

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
        TxDa1.Text = NeuDa
        TxDa4.Text = DateAdd("d", AnzVo, NeuDa)
    End If
    With MoKa1
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
Case 4:
    If IsDate(TxDa4.Text) Then
        NeuDa = TxDa4.Text
        TxDa4.Text = NeuDa
    End If
    With MoKa3
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
    Datu1 = CDate(TxDa1.Text)
    Datu2 = CDate(TxDa4.Text)
    If Datu2 <= Datu1 Then
        TxDa4.Text = DateAdd("d", AnzVo, NeuDa)
    End If
End Select

FTeLo

Set MoKa1 = Nothing
Set MoKa3 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date
Dim AnzVo As Integer

Set ZyEnT = Me.cmbZyEn1
Set TxDa1 = Me.txtDatu1
Set TxDa4 = Me.txtDatu4
Set MoKa1 = Me.dtpDatu1
Set MoKa3 = Me.dtpDatu3

AnzVo = ZyEnT.ItemData(ZyEnT.ListIndex)

Select Case KalWa
Case 1: NeuDa = MoKa1.Selection.Blocks(0).DateBegin
        TxDa1.Text = NeuDa
        TxDa4.Text = DateAdd("d", AnzVo, NeuDa)
        TxDa1.SetFocus
Case 4: NeuDa = MoKa3.Selection.Blocks(0).DateBegin
        TxDa4.Text = NeuDa
        Datu1 = CDate(TxDa1.Text)
        Datu2 = CDate(TxDa4.Text)
        If Datu2 <= Datu1 Then
            TxDa4.Text = DateAdd("d", AnzVo, NeuDa)
        End If
        TxDa4.SetFocus
End Select

FTeLo

Set MoKa1 = Nothing
Set MoKa3 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FDoTe()
On Error GoTo OrErr

Set ZwZei = Me.txoZwZei
Set ChDop = Me.chkDopTe

If ChDop.Value = xtpChecked Then
    ZwZei.Enabled = True
Else
    ZwZei.Enabled = False
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDoTe " & Err.Number
Resume Next

End Sub
Public Sub FDel()
On Error GoTo LiErr
'Löscht einen Gebühreneitrag

Dim RowNr As Integer
Dim KrRow As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpFm1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermVo
Set RpCo1 = FM.repCont1
Set RpRcs = RpCo1.Records

Set RpSel = RpCo1.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    KrRow = RpRow.Index
    Ter_Del
    DoEvents
    TVoUp KrRow
End If

Set RpFm1 = frmMain.repCont1
Set RpCls = RpFm1.Columns
Set RpSel = RpFm1.SelectedRows

Ter_VoL
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpTe RowNr
End If

Set RpRcs = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing
Set RpFm1 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDel " & Err.Number
Resume Next

End Sub
Private Sub FDrop()
On Error GoTo OrErr

Dim RowNr As Integer
Dim KrRow As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpFm1 As XtremeReportControl.ReportControl
Dim RpCo1 As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpFm1 = FM.repCont1
Set RpCls = RpFm1.Columns
Set RpSel = RpFm1.SelectedRows
Set RpCo1 = Me.repCont1
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Tr_VoEi
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpTe RowNr
End If

Set RpSel = RpCo1.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    KrRow = RpRow.Index
Else
    KrRow = 1
End If

TVoUp KrRow

Set RpCo1 = Nothing
Set RpFm1 = Nothing
Set RpCls = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDrop " & Err.Number
Resume Next

End Sub
Private Sub FDruck()
On Error GoTo LiErr

Dim IdxNr As Long
Dim Mld1, Tit1 As String

Set TxID0 = Me.txtID0

Mld1 = "Es wurde noch kein Patient zugeordnet"
Tit1 = "Kei Patint"

If TxID0.Text <> vbNullString Then
    IdxNr = TxID0.Text
    Unload Me
    STeDr "TerPat", False, IdxNr
Else
    WindowMess Mld1, Dial3, Tit1, Me.hwnd
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDruck " & Err.Number
Resume Next

End Sub
Private Sub FEiKe()
On Error GoTo OrErr

Dim TmVon As Date
Dim TmBis As Date
Dim RowNr As Long
Dim KrRow As Integer
Dim AdMin As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

Set RpCon = Me.repCont1
Set CmBrs = Me.comBar02
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmAcs = CmBrs.Actions

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
    TmBis = TimeValue(BiZei.Text)
    AdMin = DateDiff("n", TmVon, TmBis)
End If

Tr_VoEi 1
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpTe RowNr
End If

Set RpSel = RpCo1.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    KrRow = RpRow.Index
Else
    KrRow = 1
End If

TVoUp KrRow

Set RpCon = Nothing
Set RpCo1 = Nothing
Set RpCls = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEiKe " & Err.Number
Resume Next

End Sub

Private Sub FAdre()
On Error GoTo LiErr

Dim Mld1, Tit1 As String

Set TxID0 = Me.txtID0

Mld1 = "Es wurde noch kein Patient zugeordnet"
Tit1 = "Kei Patint"

If TxID0.Text <> vbNullString Then
    AMain CLng(TxID0.Text)
Else
    WindowMess Mld1, Dial3, Tit1, Me.hwnd
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAdre " & Err.Number
Resume Next

End Sub
Private Sub FBetr()
On Error GoTo LiErr

Dim FarWe As Long
Dim TmVon As Date
Dim TmBis As Date
Dim RmuNr As Long
Dim IdxNr As Long
Dim MitNr As Long
Dim ManNr As Long
Dim RmIdx As Integer
Dim MiIdx As Integer
Dim MaIdx As Integer
Dim AktZa As Integer
Dim ZeiVo As Integer
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim DaPro As XtremeCalendarControl.CalendarDataProvider
Dim CaLbs As XtremeCalendarControl.CalendarEventLabels
Dim CaLbl As XtremeCalendarControl.CalendarEventLabel

Set TxFar = Me.txtFarbe
Set CmBet = Me.txtBetre
Set CmRmu = Me.cmbRaum1
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CaCol = frmMain.calCont1
Set DaPro = CaCol.DataProvider
Set CaLbs = DaPro.LabelList

RmIdx = 0
MiIdx = 0
IdxNr = CmBet.ListIndex + 1

If IdxNr > 0 Then
    RmuNr = GlBtr(IdxNr, 2)
    ZeiVo = GlBtr(IdxNr, 3)
    
    If GlBtr(IdxNr, 4) <> vbNullString Then
        If GlBtr(IdxNr, 4) > 0 Then
            FarWe = GlBtr(IdxNr, 4)
        Else
            FarWe = vbWhite
        End If
    Else
        FarWe = vbWhite
    End If
    
    If GlBtr(IdxNr, 5) <> vbNullString Then
        MitNr = GlBtr(IdxNr, 5)
    Else
        MitNr = 0
    End If
    
    If RmuNr > 0 Then
        For AktZa = 1 To UBound(GlRmu)
            If RmuNr = GlRmu(AktZa, 0) Then
                RmIdx = AktZa - 1
                CmRmu.ListIndex = RmIdx
                Exit For
            End If
        Next AktZa
    End If
    
    If FarWe > 0 Then 'Farbwert aus Terminbetreffs
        For Each CaLbl In CaLbs
            If CaLbl.Color = FarWe Then
                TxFar.Text = CaLbl.LabelID
                TagWe = Mid$(TxFar.Tag, 2, Len(TxFar.Tag) - 1)
                TxFar.Tag = 1 & TagWe
                CmBet.BackColor = FarWe
                Exit For
            End If
        Next CaLbl
    End If
    
    If MitNr > 0 Then
        For AktZa = 1 To UBound(GlMiK)
            If MitNr = GlMiK(AktZa, 2) Then
                ManNr = GlMiK(AktZa, 7)
                MiIdx = AktZa - 1
                CmMit.ListIndex = MiIdx
                Exit For
            End If
        Next AktZa

        For AktZa = 1 To UBound(GlMan)
            If ManNr = GlMan(AktZa, 2) Then
                MaIdx = AktZa - 1
                CmMan.ListIndex = MaIdx
                Exit For
            End If
        Next AktZa
    End If
End If

If ZeiVo > 0 Then
    If VoZei.Text <> vbNullString Then
        TmVon = TimeValue(VoZei.Text)
        TmBis = DateAdd("n", ZeiVo, TmVon)
        BiZei.Text = Format$(TmBis, "hh:mm")
    End If
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FBetr " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case GlBut
Case RibTab_Rechnungen:
    TeTit = IniGetOpt("Hilfe", 50881)
    TeMai = IniGetOpt("Hilfe", 50882)
    TeInh = IniGetOpt("Hilfe", 50883)
    TeFus = IniGetOpt("Hilfe", 50884)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
Case Else:
    TeTit = IniGetOpt("Hilfe", 50901)
    TeMai = IniGetOpt("Hilfe", 50902)
    TeInh = IniGetOpt("Hilfe", 50903)
    TeFus = IniGetOpt("Hilfe", 50904)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
End Select

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date
Dim AnzVo As Integer

Set ZyEnT = Me.cmbZyEn1
Set TxDa1 = Me.txtDatu1
Set TxDa4 = Me.txtDatu4
Set MoKa1 = Me.dtpDatu1
Set MoKa3 = Me.dtpDatu3

AnzVo = ZyEnT.ItemData(ZyEnT.ListIndex)

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
    Else
        NeuDa = Date
    End If
    With MoKa1
        .EnsureVisible NeuDa
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
        .Top = TxDa1.Top + TxDa1.Height
        .Left = TxDa1.Left
        If .ShowModal(1, 1) Then
            If .Selection.BlocksCount > 0 Then
                NeuDa = .Selection.Blocks(0).DateBegin
                TxDa1.Text = NeuDa
                TxDa4.Text = DateAdd("d", AnzVo, NeuDa)
            End If
        End If
    End With
Case 4:
    If IsDate(TxDa4.Text) Then
        NeuDa = TxDa4.Text
    Else
        NeuDa = Date
    End If
    With MoKa3
        .EnsureVisible NeuDa
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
        .Top = TxDa4.Top + TxDa4.Height
        .Left = TxDa4.Left
        If .ShowModal(1, 1) Then
            If .Selection.BlocksCount > 0 Then
                NeuDa = .Selection.Blocks(0).DateBegin
                TxDa4.Text = NeuDa
                Datu1 = CDate(TxDa1.Text)
                Datu2 = CDate(TxDa4.Text)
                If Datu2 <= Datu1 Then
                    TxDa4.Text = DateAdd("d", AnzVo, NeuDa)
                End If
            End If
        End If
    End With
End Select

FTeLo

Set MoKa1 = Nothing
Set MoKa3 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKaRo(ByVal RpBut As XtremeReportControl.IReportInplaceButton)
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim ItmLi As Long
Dim ItmOb As Long
Dim ItmRe As Long
Dim ItmHo As Long
Dim ItmBr As Long
Dim ItmTo As Long
Dim RmuNr As Long
Dim AltDa As Date
Dim NeuDa As Date
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermVo
Set ChRau = FM.chkRauZu
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns
Set RpSel = RpCo6.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If RpRow.Record(2).Value <> vbNullString Then
            AltDa = CDate(RpRow.Record(2).Value)
        Else
            AltDa = Date
        End If
        RpBut.GetRect ItmLi, ItmOb, ItmRe, ItmHo
        ItmBr = ItmRe
        ItmTo = ItmHo + 1
        NeuDa = FKaSh(ItmBr, ItmTo, AltDa, RpCo6.hwnd, True)
        If IsDate(NeuDa) Then
            RpCo6.EditItem Nothing, Nothing
            RpRow.Record(1).Value = Format$(NeuDa, "dddd")
            RpRow.Record(2).Value = CDate(NeuDa)
            RpCo6.Populate
            If ChRau.Value = xtpChecked Then
                Ter_Spa RmuNr
            Else
                Ter_Spa
            End If
        End If
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo6 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaRo " & Err.Number
Resume Next

End Sub
Private Function FKaSh(ByVal KalLi As Long, ByVal KalOb As Long, ByVal NeuDa As Date, ByVal mHwnd As Long, Optional ByVal Flag As Boolean = False) As Date
On Error GoTo LaErr

Dim RpCo6 As XtremeReportControl.ReportControl

Dim Datu1 As Date
Dim DayFi As Date
Dim DayLa As Date
Dim KaBre As Long
Dim KaHoh As Long
Dim RetWe As Boolean

Set FM = frmTermVo
Set MoKa2 = FM.dtpDatu2

DayFi = NeuDa - 30
DayLa = NeuDa + 30

With MoKa2
    .RedrawControl
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .GetMinReqRect KaBre, KaHoh, 1, 1
    If Flag = True Then
        RetWe = .ShowModalEx(KalLi - KaBre - 4 - 4, KalOb, KaBre + 4, KaHoh + 4, mHwnd)
    Else
        RetWe = .ShowModalEx(-1, 20, KaBre + 4, KaHoh + 4, mHwnd)
    End If
End With

If RetWe = True Then
    If MoKa2.Selection.BlocksCount > 0 Then
        Datu1 = MoKa2.Selection.Blocks(0).DateBegin()
        If IsDate(Datu1) Then
            FKaSh = CDate(Datu1)
        Else
            FKaSh = NeuDa
        End If
    Else
        FKaSh = NeuDa
    End If
Else
    FKaSh = NeuDa
End If

Set MoKa2 = Nothing

Exit Function

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaSh " & Err.Number
Resume Next

End Function
Private Sub FKata()
On Error GoTo LiErr

Dim KatNr As Long
Dim Mld1, Tit1 As String

Set TxID0 = Me.txtID0
Set TxDa1 = Me.txtDatu1

Mld1 = "Es wurde noch kein Patient zugeordnet"
Tit1 = "Kei Patint"

If TxID0.Text <> vbNullString Then
    If IsNumeric(TxID0.Text) = True Then
        If CLng(TxID0.Text) > 0 Then
            KatNr = S_AdIdi(CLng(TxID0.Text), "ID3")
            TrMain KatNr
        End If
    End If
Else
    WindowMess Mld1, Dial3, Tit1, Me.hwnd
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKata " & Err.Number
Resume Next

End Sub

Private Sub FKran(Optional ByVal CoIdx As Long)
On Error GoTo PoErr
'Änderungen im Krankneblatt

Dim GesBe As Single
Dim EinBe As Single
Dim Fakto As Single
Dim Anzal As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermVo
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

Select Case CoIdx
Case TeL_Typ:
Case TeL_IDKurz:
Case TeL_Anz:
Case TeL_Multi:
Case TeL_Betrag:
Case Else: Exit Sub
End Select

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then

        Set RpCol = RpCls.Find(TeL_IDKurz)
        RpRow.Record(RpCol.ItemIndex).Tag = "@IDKurz"
        If IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then RpRow.Record(RpCol.ItemIndex).Value = "..."
        Set RpCol = RpCls.Find(TeL_Multi)
        If IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then RpRow.Record(RpCol.ItemIndex).Value = GlWa3
        Fakto = CSng(RpRow.Record(RpCol.ItemIndex).Value)
        Set RpCol = RpCls.Find(TeL_Anz)
        If IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then RpRow.Record(RpCol.ItemIndex).Value = 1
        Anzal = RpRow.Record(RpCol.ItemIndex).Value
        
        Set RpCol = RpCls.Find(TeL_Betrag)
        If IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then
            Set RpCol = RpCls.Find(TeL_Gesamt)
            If Not IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then
                GesBe = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(TeL_Betrag)
                RpRow.Record(RpCol.ItemIndex).Value = GesBe
            Else
                Set RpCol = RpCls.Find(TeL_Betrag)
                RpRow.Record(RpCol.ItemIndex).Value = 0
            End If
        Else
            EinBe = CSng(RpRow.Record(RpCol.ItemIndex).Value)
            RpRow.Record(RpCol.ItemIndex).Value = EinBe
        End If
        
        Set RpCol = RpCls.Find(TeL_Gesamt)
        If IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then
            Set RpCol = RpCls.Find(TeL_Betrag)
            If Not IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then
                EinBe = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(TeL_Gesamt)
                RpRow.Record(RpCol.ItemIndex).Value = EinBe
            Else
                Set RpCol = RpCls.Find(TeL_Gesamt)
                RpRow.Record(RpCol.ItemIndex).Value = 0
            End If
        End If
        
        Set RpCol = RpCls.Find(TeL_Betrag)
        EinBe = CSng(RpRow.Record(RpCol.ItemIndex).Value)
        RpRow.Record(RpCol.ItemIndex).Value = Format$(EinBe, GlWa1)
        RpRow.Record(RpCol.ItemIndex).Tag = "@Preis1" 'Tag geändert
        Set RpCol = RpCls.Find(TeL_Gesamt)
        RpRow.Record(RpCol.ItemIndex).Value = Format$(EinBe * Fakto * Anzal, GlWa1)
        RpRow.Record(RpCol.ItemIndex).Tag = "@Preis2" 'Tag geändert
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKrAn " & Err.Number
Resume Next

End Sub
Private Sub FKrCa(ByVal CoIdx As Long)
On Error GoTo PoErr
'Änderungen im Krankneblatt

Dim ZeSta As Date
Dim ZeEnd As Date
Dim AdMin As Long
Dim RmuNr As Long
Dim NeuDa As Date
Dim StaZe As Date
Dim EndZe As Date
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermVo
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set ChRau = FM.chkRauZu
Set CmRmu = FM.cmbRaum1
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns
Set RpSel = RpCo6.SelectedRows

ZeSta = TimeValue(VoZei.Text)
ZeEnd = TimeValue(BiZei.Text)
AdMin = DateDiff("n", TimeValue(ZeSta), TimeValue(ZeEnd))

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Select Case CoIdx
        Case 2:
            Set RpCol = RpCls.Find(2)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                NeuDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                RpRow.Record(RpCol.ItemIndex).Value = NeuDa
                Set RpCol = RpCls.Find(1)
                RpRow.Record(RpCol.ItemIndex).Value = Format$(NeuDa, "dddd")
            End If
            DoEvents
            If ChRau.Value = xtpChecked Then
                Set RpCol = RpCls.Find(9)
                RmuNr = RpRow.Record(RpCol.ItemIndex).Value
                Ter_Spa RmuNr
            Else
                Ter_Spa
            End If
        Case 3:
            Set RpCol = RpCls.Find(3)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                StaZe = TimeValue(RpRow.Record(RpCol.ItemIndex).Value)
                EndZe = DateAdd("n", AdMin, StaZe)
                Set RpCol = RpCls.Find(4)
                RpRow.Record(RpCol.ItemIndex).Value = Format$(EndZe, "hh:mm")
            Else
                RpRow.Record(RpCol.ItemIndex).Value = "08:00"
            End If
            DoEvents
            If ChRau.Value = xtpChecked Then
                Set RpCol = RpCls.Find(9)
                RmuNr = RpRow.Record(RpCol.ItemIndex).Value
                Ter_Zei RmuNr
            Else
                Ter_Zei
            End If
        End Select
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo6 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKrCa " & Err.Number
Resume Next

End Sub
Private Sub FOpt(Optional ByVal SeOpt As Boolean = False)
On Error Resume Next

Set ChDop = Me.chkDopTe
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4

Select Case GlTvM
Case "M1": Rahm1.Visible = True
           Rahm2.Visible = False
           Rahm3.Visible = False
           Rahm4.Visible = False
           ChDop.Enabled = True
Case "M2": Rahm1.Visible = False
           Rahm2.Visible = True
           Rahm3.Visible = False
           Rahm4.Visible = False
           ChDop.Enabled = True
Case "M3": Rahm1.Visible = False
           Rahm2.Visible = False
           Rahm3.Visible = True
           Rahm4.Visible = False
           ChDop.Enabled = False
           ChDop.Value = xtpUnchecked
Case "M4": Rahm1.Visible = False
           Rahm2.Visible = False
           Rahm3.Visible = False
           Rahm4.Visible = True
           ChDop.Enabled = False
           ChDop.Value = xtpUnchecked
End Select

IniSetVal "Layout", "TeVoMu", GlTvM

End Sub
Private Sub FTaEd(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
On Error Resume Next

Dim RpCo1 As XtremeReportControl.ReportControl

Set FM = frmTermVo
Set RpCo1 = FM.repCont1
Set RpSel = RpCo1.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If Shift = 0 Then
            Select Case KeyCode
            Case vbKeyF2: RpCo1.Navigator.BeginEdit
            Case vbKeyTab:
            Case vbKeyReturn:
            Case vbKeyDown:
            Case vbKeyUp:
            Case vbKeyPageDown:
            Case vbKeyPageUp:
            End Select
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo1 = Nothing

End Sub
Private Sub FTeLo()
On Error GoTo KoErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo6 As XtremeReportControl.ReportControl

If GlAkt = False Then
    Set FM = frmTermVo
    Set ZyEn2 = FM.optZyEn2
    Set ZyEn3 = FM.optZyEn3
    Set RpCo6 = FM.repCont6
    Set CmBrs = FM.comBar02
    Set CmAcs = CmBrs.Actions
    Set RpRcs = RpCo6.Records

    With RpCo6
        .EditItem Nothing, Nothing
        If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
        If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
        If .Records.Count > 0 Then .Records.DeleteAll
        .Populate
    End With
    
    CmAcs(AD_Termin_Vorschau).Enabled = True
    CmAcs(AD_Termin_Save).Enabled = False
    CmAcs(AD_Termin_Reset).Enabled = False
    CmAcs(AD_Termin_Freie).Enabled = True
    
    FStat
    
    Set CmAcs = Nothing
    Set RpRcs = Nothing
    Set CmBrs = Nothing
    Set RpCo6 = Nothing
End If

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeLo " & Err.Number
Resume Next

End Sub
Private Sub FStat()
On Error GoTo KoErr

Dim GesZa As Integer
Dim TerZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmTermVo
Set CmBrs = FM.comBar02
Set RpCo6 = FM.repCont6
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RpRcs = RpCo6.Records

GesZa = RpRcs.Count

If GesZa > 0 Then
    For Each RpRec In RpRcs
        If RpRec.Item(5).Checked = True Then
            TerZa = TerZa + 1
        End If
    Next RpRec
    CmSta.Pane(1).Text = "Generiert : " & GesZa & " - Termine : " & TerZa
Else
    CmSta.Pane(1).Text = "Generiert : 0 - Termine : 0"
End If

Set CmAcs = Nothing
Set RpRcs = Nothing
Set CmBrs = Nothing
Set RpCo6 = Nothing

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStat " & Err.Number
Resume Next

End Sub
Private Sub chkRauZu_Click()

Set ChRau = Me.chkRauZu
Set CmRmu = Me.cmbRaum1
    
If GlSeF = False Then
    If ChRau.Value = xtpChecked Then
        IniSetVal "TerSys", "RmuBer", -1
    Else
        IniSetVal "TerSys", "RmuBer", 0
    End If
    
    If CmRmu.ItemData(CmRmu.ListIndex) = 0 Then
        CmRmu.ListIndex = 0
    End If
End If

FTeLo

End Sub

Private Sub chkSprZe_Click()

Set ChSpr = Me.chkSprZe
    
If GlSeF = False Then 'Formular wird geladen
    If ChSpr.Value = xtpChecked Then
        GlSpP = True 'Überprüfung der Sprechzeiten
        IniSetVal "TerSys", "SpreBe", -1
    Else
        GlSpP = False
        IniSetVal "TerSys", "SpreBe", 0
    End If
End If

FTeLo

End Sub
Private Sub chkTeSpl_Click()

Set ChSpl = Me.chkTeSpl
Set ChDop = Me.chkDopTe
Set TxSp1 = Me.txoSplAn
Set TxSp2 = Me.txoSplPa

If ChSpl.Value = xtpChecked Then
    TxSp1.Enabled = True
    TxSp2.Enabled = True
Else
    TxSp1.Enabled = False
    TxSp2.Enabled = False
End If

FTeLo

End Sub

Private Sub cmbAbger_Click()

TagWe = Mid$(Me.cmbAbger.Tag, 2, Len(Me.cmbAbger.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbAbger.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub cmbAbger_GotFocus()
    RetWe = SendMessage(Me.cmbAbger.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbAbger_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbGanzt_Click()

If GlTeF = False Then 'Formular wird geladen
    GlTSa = True
End If

End Sub

Private Sub cmbGanzt_GotFocus()
    RetWe = SendMessage(Me.cmbGanzt.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbGanzt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbMitar_Click()

TagWe = Mid$(Me.cmbMitar.Tag, 2, Len(Me.cmbMitar.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    FNoVa
    Me.cmbMitar.Tag = 1 & TagWe
    GlTSa = True
End If

FTeLo

End Sub

Private Sub cmbNotVa_Click()

TagWe = Mid$(Me.cmbNotVa.Tag, 2, Len(Me.cmbNotVa.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbNotVa.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub cmbNotVa_GotFocus()
    RetWe = SendMessage(Me.cmbNotVa.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbNotVa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbTeTyp_Click()

TagWe = Mid$(Me.cmbTeTyp.Tag, 2, Len(Me.cmbTeTyp.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbTeTyp.Tag = 1 & TagWe
    GlTSa = True
End If

FTeLo

End Sub

Private Sub cmbTeTyp_GotFocus()
    RetWe = SendMessage(Me.cmbTeTyp.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbTeTyp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbTeTyp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbTeTyp.SelLength = 0
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)

Dim AktTa As Long
Dim AktZa As Integer

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTeV = True Then 'Termine vorhanden
    If GlTpV = True Then 'Kalendermarker vorhanden
        For AktTa = 0 To GlKMa - 1 'Anzahl Kalendermatker
            If Day = Left$(GlTEr(0, AktTa), 10) Then
                For AktZa = 1 To UBound(GlTep) 'Kalendermarker
                    If GlTep(AktZa, 0) = GlTEr(1, AktTa) Then
                        Metrics.BackColor = GlTep(AktZa, 2)
                        Exit For
                    End If
                Next AktZa
            End If
        Next AktTa
    End If
End If

End Sub
Private Sub dtpDatu1_SelectionChanged()
    If GlSeF = False Then
        FDatu
        TeVoSt
    End If
End Sub
Private Sub Form_Load()

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 12000
    .ClientMaxWidth = 17000
    .ClientMinHeight = 8000
    .ClientMinWidth = 14400
    .TopMost = True
End With

Set FrmEx = Nothing

End Sub
Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)

Dim AktZa As Integer

If Row.GroupRow = False Then
    If IsNumeric(Item.Record(TeL_Typ).Value) Then
        For AktZa = 1 To UBound(GlKrA)
            If Item.Record(TeL_Typ).Value = GlKrA(AktZa, 0) Then
                Metrics.ForeColor = GlKrA(AktZa, 3)
                Exit For
            End If
        Next AktZa
    End If
End If

End Sub
Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
    FTaEd KeyCode, Shift
End Sub

Private Sub repCont1_RecordsDropped(ByVal TargetRecord As XtremeReportControl.IReportRecord, ByVal Records As XtremeReportControl.IReportRecords, ByVal Above As Boolean)
    FDrop
End Sub
Private Sub repCont1_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String

If GlTeF = False Then 'Formular wird geladen
    TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
    Item.Tag = "@" & TmTag
    FKran Column.ItemIndex
End If

End Sub

Private Sub txoSplAn_Click()

Dim AnzSp As Integer

Set TxSp1 = Me.txoSplAn

AnzSp = TxSp1.ItemData(TxSp1.ListIndex)

IniSetVal "TerSys", "TeSpAn", TxSp1.ListIndex

End Sub
Private Sub txoSplPa_Click()

Dim PauSp As Integer

Set TxSp2 = Me.txoSplPa

PauSp = TxSp2.ItemData(TxSp2.ListIndex)

IniSetVal "TerSys", "TeSpPa", TxSp2.ListIndex

End Sub
Private Sub txoZwZei_Click()

Set ZwZei = Me.txoZwZei

If ZwZei.Text <> vbNullString Then
    IniSetVal "TerSys", "ZwZeit", Format$(ZwZei.Text, "hh:mm")
End If

FTeLo

End Sub
Private Sub txtAdres_Change()

TagWe = Mid$(Me.txtAdres.Tag, 2, Len(Me.txtAdres.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtAdres.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtAdres_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtAdres.SelStart = 0
        Me.txtAdres.SelLength = Len(Me.txtAdres.Text)
    End If
End Sub
Private Sub txtAdres_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtAdres_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtAdres.SelLength = 0
End Sub
Private Sub txtAdres_LostFocus()
On Error Resume Next

Dim GesZa As Long
Dim SuStr As String
Dim FLis1 As XtremeSuiteControls.ListBox
    
Set TxID0 = Me.txtID0
Set TxAdr = Me.txtAdres

SuStr = TxAdr.Text

If SuStr <> vbNullString Then
    GesZa = Ter_Adr(SuStr)
    If GesZa > 1 Then
        Load frmTermAnh
        Set FM = frmTermAnh
        Set FLis1 = FM.lstList1
        FM.Show
        FLis1.SetFocus
    End If
End If

End Sub
Private Sub cmbBehan_Click()

TagWe = Mid$(Me.cmbBehan.Tag, 2, Len(Me.cmbBehan.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbBehan.Tag = 1 & TagWe
    GlTSa = True
End If

FTeLo

End Sub
Private Sub txtBisZe_Click()
On Error Resume Next

TagWe = Mid$(Me.txtBisZe.Tag, 2, Len(Me.txtBisZe.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtBisZe.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtBisZe_GotFocus()
    Me.txtBisZe.SelStart = 0
    Me.txtBisZe.SelLength = Len(Me.txtBisZe.Text)
End Sub

Private Sub txtBisZe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtBisZe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtBisZe.SelLength = 0
End Sub
Private Sub txtBisZe_LostFocus()
On Error Resume Next

Dim TmVon As Date
Dim TmBis As Date
Dim ZeiVo As Integer
Dim AktZa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre

If InStrRev(BiZei.Text, "_", -1, 1) > 0 Then
    BiZei.Text = "00:00"
End If

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then ZeiVo = 15

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
Else
    VoZei.Text = Format$(TimeValue(Now), "hh:mm")
    TmVon = Now
End If

If BiZei.Text <> vbNullString Then
    TmBis = TimeValue(BiZei.Text)
Else
    BiZei.Text = Format$(TimeValue(Now), "hh:mm")
    TmVon = Now
End If

If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
    If ZeiVo > 0 Then
        TmVon = DateAdd("n", -ZeiVo, TmBis)
        VoZei.Text = Format$(TmVon, "hh:mm")
    Else
        If TmVon > TmBis Then
            BiZei.Text = VoZei.Text
        End If
    End If
Else
    If TmVon > TmBis Then
        BiZei.Text = VoZei.Text
    End If
End If

TagWe = Mid$(Me.txtBisZe.Tag, 2, Len(Me.txtBisZe.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtBisZe.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub cmbRemin_Click()

TagWe = Mid$(Me.cmbRemin.Tag, 2, Len(Me.cmbRemin.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbRemin.Tag = 1 & TagWe
    GlTSa = True
End If

FTeLo

End Sub

Private Sub cmbRemin_GotFocus()
    RetWe = SendMessage(Me.cmbRemin.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbRemin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbStatu_Click()

TagWe = Mid$(Me.cmbStatu.Tag, 2, Len(Me.cmbStatu.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbStatu.Tag = 1 & TagWe
    GlTSa = True
End If

FTeLo

End Sub

Private Sub cmbStatu_GotFocus()
    RetWe = SendMessage(Me.cmbStatu.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbStatu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbStatu_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbStatu.SelLength = 0
End Sub
Private Sub txtDatu1_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtDatu1.SelStart = 0
        Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
    End If
End Sub
Private Sub txtDatu1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtDatu1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtDatu1.SelLength = 0
End Sub

Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub

Private Sub txtRaum1_Change()

TagWe = Mid$(Me.txtRaum1.Tag, 2, Len(Me.txtRaum1.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtRaum1.Tag = 1 & TagWe
    GlTSa = True
End If

FTeLo

End Sub
Private Sub txtRaum1_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtRaum1.SelStart = 0
        Me.txtRaum1.SelLength = Len(Me.txtRaum1.Text)
    End If
End Sub
Private Sub txtRaum1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRaum1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtRaum1.SelLength = 0
    End If
End Sub

Private Sub txtDatu1_Change()

TagWe = Mid$(Me.txtDatu1.Tag, 2, Len(Me.txtDatu1.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtDatu1.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtVonZe_Click()
On Error Resume Next

TagWe = Mid$(Me.txtVonZe.Tag, 2, Len(Me.txtVonZe.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtVonZe.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtVonZe_GotFocus()
    Me.txtVonZe.SelStart = 0
    Me.txtVonZe.SelLength = Len(Me.txtVonZe.Text)
End Sub
Private Sub txtVonZe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtVonZe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtVonZe.SelLength = 0
End Sub
Private Sub txtVonZe_LostFocus()
On Error Resume Next

Dim TmVon As Date
Dim TmBis As Date
Dim ZeiVo As Integer
Dim AktZa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre

If InStrRev(VoZei.Text, "_", -1, 1) > 0 Then
    VoZei.Text = "00:00"
End If

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then ZeiVo = 15

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
Else
    TmVon = Now
    VoZei.Text = Format$(TimeValue(Now), "hh:mm")
End If

If BiZei.Text <> vbNullString Then
    TmBis = TimeValue(BiZei.Text)
Else
    TmVon = Now
    BiZei.Text = Format$(TimeValue(Now), "hh:mm")
End If

If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
    If ZeiVo > 0 Then
        TmBis = DateAdd("n", ZeiVo, TmVon)
        BiZei.Text = Format$(TmBis, "hh:mm")
    Else
        If TmVon > TmBis Then
            BiZei.Text = VoZei.Text
        End If
    End If
Else
    If TmVon > TmBis Then
        BiZei.Text = VoZei.Text
    End If
End If

TagWe = Mid$(Me.txtBisZe.Tag, 2, Len(Me.txtBisZe.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtBisZe.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlTeF = False Then 'Formular wird geladen
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    TeVoPo
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub

Private Sub FClos()
On Error GoTo LiErr

If WindowLoad("frmTerKat") = True Then
    Unload frmTerKat
End If

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlIdi = False Then 'Idiotenmodus
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "TermVo", "FenLin", clFen.FeLin
        IniSetVal "TermVo", "FenObe", clFen.FeObn
        IniSetVal "TermVo", "FenBre", clFen.FeBre
        IniSetVal "TermVo", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FNoVa(Optional ByVal NoSet As Boolean = False)
On Error GoTo LiErr

Dim NotVa As Integer
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTermVo
Set CmMit = FM.cmbMitar
Set CmMan = Me.cmbBehan
Set CmNot = FM.cmbNotVa

Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If GlMiT(MiIdx, 39) > 0 Then
        NotVa = GlMiT(MiIdx, 39)
    Else
        NotVa = 24
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If GlMaT(MaIdx, 25) > 0 Then
        NotVa = GlMaT(MaIdx, 25)
    Else
        NotVa = 24
    End If
End If

If NoSet = True Then
    If CmAcs(AD_Termin_Notify).Checked = False Then
        CmAcs(AD_Termin_Notify).Checked = True
        CmNot.Enabled = True
        CmNot.ListIndex = NotVa
    Else
        CmAcs(AD_Termin_Notify).Checked = False
        CmNot.Enabled = False
        CmNot.ListIndex = 0
    End If
Else
    CmNot.ListIndex = NotVa
End If

TagWe = Mid$(CmNot.Tag, 2, Len(CmNot.Tag) - 1)
CmNot.Tag = 1 & TagWe

GlTSa = True

Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNoVa " & Err.Number
Resume Next

End Sub

Private Sub FErin()
On Error Resume Next

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTermVo
Set CmRem = FM.cmbRemin
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

If CmAcs(AD_Termin_Remind).Checked = False Then
    CmAcs(AD_Termin_Remind).Checked = True
    CmRem.Enabled = True
    CmRem.ListIndex = 5
Else
    CmAcs(AD_Termin_Remind).Checked = False
    CmRem.Enabled = False
    CmRem.ListIndex = 0
End If

TagWe = Mid$(CmRem.Tag, 2, Len(CmRem.Tag) - 1)
CmRem.Tag = 1 & TagWe

GlTSa = True

Set CmSta = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FFarb(ByVal Flag As Integer)
On Error Resume Next

Dim Farbe As Long

Set FM = frmTermVo
Set TxFar = FM.txtFarbe

TxFar.Text = Flag

TagWe = Mid$(TxFar.Tag, 2, Len(TxFar.Tag) - 1)
TxFar.Tag = 1 & TagWe

GlTSa = True

TeFarb Flag, 4

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmTermVo = Nothing
End Sub

Private Sub FSave()
On Error GoTo SaErr
'Überprüft, ob der Eintrag geändert wurde und speichert dieses ab

Dim MasTe As Long
Dim AktZa As Integer
Dim WiVor As Boolean
Dim Mld1, Mld2, Mld3, Tit1 As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CaCol As XtremeCalendarControl.CalendarControl

Set FM = frmTermVo
Set TxID0 = FM.txtID0
Set CmBet = FM.txtBetre
Set CmRmu = FM.cmbRaum1
Set RpCo1 = FM.repCont1
Set RpRcs = RpCo1.Records

Tit1 = "Kein Betreff vorhanden"
Mld1 = "Sie müssen mind. einen Betreff eintragen, damit ein Terminvorschlag erstellt werden kann"
Mld2 = "Es wurden noch keine Gebührenziffern oder andere Leistungen hinzugefügt"
Mld3 = "Es wurde noch kein Patient zugeordnet"

Select Case GlBut
Case RibTab_Rechnungen:
    If TxID0.Text = vbNullString Then
        WindowMess Mld3, Dial2, Tit1, FM.hwnd
        Exit Sub
    Else
        If TxID0.Text <= 0 Then
            WindowMess Mld3, Dial2, Tit1, FM.hwnd
            Exit Sub
        End If
    End If
    If RpRcs.Count = 0 Then
        WindowMess Mld2, Dial2, Tit1, FM.hwnd
        Exit Sub
    End If
    WiVor = True
Case RibTab_Abrechnung:
    If TxID0.Text = vbNullString Then
        WindowMess Mld3, Dial2, Tit1, FM.hwnd
        Exit Sub
    Else
        If TxID0.Text <= 0 Then
            WindowMess Mld3, Dial2, Tit1, FM.hwnd
            Exit Sub
        End If
    End If
    If RpRcs.Count = 0 Then
        WindowMess Mld2, Dial2, Tit1, FM.hwnd
        Exit Sub
    End If
    WiVor = True
End Select

MasTe = Ter_VoT
DoEvents

If CmBet.Text = vbNullString Then
    If TxID0.Text = vbNullString Then
        WindowMess Mld1, Dial2, Tit1, FM.hwnd
        Exit Sub
    End If
End If

Screen.MousePointer = vbHourglass
DoEvents

If CmRmu.ItemData(CmRmu.ListIndex) = 0 Then
    For AktZa = 1 To UBound(GlRmu)
        Ter_VoS MasTe, GlRmu(AktZa, 0), WiVor
    Next AktZa
Else
    Ter_VoS MasTe, 0, WiVor
End If

Select Case GlBut
Case RibTab_Rechnungen:
    P_TeSpl
    P_List "ReSe", 0, 1
    DoEvents
Case RibTab_Abrechnung:
    P_TeSpl
    P_List "ReSe", 0, 1
    DoEvents
Case Else:
    DoEvents
    S_TeLi
    DoEvents
    S_TePi 'Kalndermarker setzen
    DoEvents
    SUpTe
    DoEvents
End Select

DoEvents
Screen.MousePointer = vbNormal

DoEvents
Unload Me

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FSet()
On Error Resume Next

Dim CmBrs As XtremeCommandBars.CommandBars

Set ZyEn2 = Me.optZyEn2
Set ZyEn3 = Me.optZyEn3
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

CmAcs(AD_Termin_Vorschau).Enabled = True
CmAcs(AD_Termin_Save).Enabled = False
CmAcs(AD_Termin_Reset).Enabled = False
CmAcs(AD_Termin_Freie).Enabled = False

Set CmAcs = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: FVors
Case KY_F5: frmAdrSuch.Show vbModal
Case KY_F7: FTeLo
Case KY_F8: FSave
Case KY_F10: FDruck
Case KY_F11: Unload Me
Case AD_Termin_Vorschau: FVors
Case AD_Termin_Save: FSave
Case AD_Termin_Reset: FTeLo
Case AD_Termin_Freie: FVors True
Case AD_Termin_Remind: FErin
Case AD_Termin_Notify: FNoVa True
Case TE_Termin_Beenden: Unload Me
Case TE_Termin_Hilfe: FHilfe
Case TE_Termin_Drucken: FDruck
Case AD_Termin_Ketten: FKata
Case AD_Termin_StaKett: FEiKe
Case AD_Termin_Abrechnen: Ter_Rec
Case AD_Termin_EintLoe: FDel
Case FaLei01: FFarb 1
Case FaLei02: FFarb 2
Case FaLei03: FFarb 3
Case FaLei04: FFarb 4
Case FaLei05: FFarb 5
Case FaLei06: FFarb 6
Case FaLei07: FFarb 7
Case FaLei08: FFarb 8
Case FaLei09: FFarb 9
Case FaLei10: FFarb 10
Case FaLei11: FFarb 11
Case FaLei12: FFarb 12
Case FaLei13: FFarb 13
Case FaLei14: FFarb 14
Case FaLei15: FFarb 15
Case FaLei16: FFarb 16
Case FaLei17: FFarb 17
Case FaLei18: FFarb 18
Case FaLei19: FFarb 19
Case FaLei20: FFarb 20
End Select

GlToo = False

End Sub
Private Sub FVors(Optional ByVal FreTe As Boolean = False)
On Error GoTo InErr

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim RmuNr As Long
Dim MndNr As Long
Dim TeBtr As String
Dim TreZe As String
Dim SpBtr As String
Dim PaMin As Integer
Dim AdMin As Integer
Dim SplAn As Integer
Dim SplPa As Integer
Dim AktPo As Integer
Dim StaPo As Integer
Dim AktZa As Integer
Dim Behin As Integer
Dim RauId As Integer
Dim TrSpl As Boolean
Dim Mld1, Mld6 As String
Dim Mld2, Mld3 As String
Dim Mld4, Mld5 As String
Dim Tit1 As String

Set TxID0 = Me.txtID0
Set TxBeG = Me.txtBehin
Set CmBet = Me.txtBetre
Set ChTer = Me.chkFreTe
Set ChRau = Me.chkRauZu
Set CmRmu = Me.cmbRaum1
Set CmMan = Me.cmbBehan
Set ChSpl = Me.chkTeSpl
Set TxSp1 = Me.txoSplAn
Set TxSp2 = Me.txoSplPa
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

Tit1 = "Terminvorschlag"
Mld1 = "Sie müssen mind. einen Betreff eintragen, damit ein Terminvorschlag erstellt werden kann"
Mld2 = "Die Dauer eines jeweiligen Teiltermins beträgt weniger as 5 Minuten. Erhöhen Sie entweder den Terminzeitraum oder verringern Sie die Länge bzw. die Anzahl der Pausen"
Mld3 = "Die Gesamtdauer der Pausen für die jeweiligen Teiltermine ist größer als die Terminzeit selbst. Erhöhen Sie entweder den Terminzeitraum oder verringern Sie die Länge bzw. die Anzahl der Pausen"
Mld4 = "Die errechnete Anzahl des durch das Trennzeichnen getrennten Terminbetreffs stimmt nicht mit der festgelegten Anzahl der Teiltermine überein. Bitte überprüfen Sie an Anzahl der eingesetzten Trennzeichen"
Mld6 = "Der ausgewählte Raum, ist für den Behinderungsgrad des Patienten nicht geeignet"

TreZe = IniGetVal("TerSys", "TreZei")

If ChSpl.Value = xtpChecked Then
    TrSpl = True 'Termine Splitten
End If

StaPo = 1
AlDa1 = TimeValue(VoZei.Text)
AlDa2 = TimeValue(BiZei.Text)
RmuNr = CmRmu.ItemData(CmRmu.ListIndex)
MndNr = CmMan.ItemData(CmMan.ListIndex)
SplAn = TxSp1.ItemData(TxSp1.ListIndex)
SplPa = TxSp2.ItemData(TxSp2.ListIndex)
AdMin = DateDiff("n", AlDa1, AlDa2)
PaMin = (SplAn - 1) * SplPa

RauId = CmRmu.ListIndex + 1

If RauId > UBound(GlRmu) Then
    RauId = RauId - 1
End If

If CmBet.Text = vbNullString Then
    If TxID0.Text = vbNullString Then
        WindowMess Mld1, Dial2, Tit1, Me.hwnd
        Exit Sub
    End If
Else
    TeBtr = CmBet.Text
End If

If TxBeG.Text <> vbNullString Then
    If IsNumeric(TxBeG.Text) Then
        Behin = CInt(TxBeG.Text)
    Else
        Behin = 0
    End If
Else
    Behin = 0
End If

If Behin > 0 Then
    If GlRmu(RauId, 3) > 0 Then
        If GlBeG(Behin, 0) > GlRmu(RauId, 3) Then
            SPopu Tit1, Mld6, IC48_Warning
            Exit Sub
        End If
    End If
End If

If TrSpl = True Then 'Termine Splitten
    If PaMin >= AdMin Then
        WindowMess Mld3, Dial2, Tit1, Me.hwnd
        Exit Sub
    End If
    If ((AdMin - PaMin) / SplAn) < 5 Then
        WindowMess Mld2, Dial2, Tit1, Me.hwnd
        Exit Sub
    End If
    
    AktZa = 0
    AktPo = InStr(StaPo, TeBtr, TreZe, 1)
    If AktPo > 0 Then
        AktZa = AktZa + 1
        StaPo = AktPo + 1
        Do While AktPo <> 0
            AktPo = InStr(StaPo, TeBtr, TreZe, 1)
            If AktPo > 0 Then
                AktZa = AktZa + 1
                StaPo = AktPo + 1
            End If
        Loop
        If SplAn <> (AktZa + 1) Then
            WindowMess Mld4, Dial2, Tit1, Me.hwnd
            Exit Sub
        End If
    End If
End If

If ChTer.Value = xtpChecked Then 'bereits belegte Termine berücksichtiigen
    If ChRau.Value = xtpChecked Then 'Raumbelegung berücksichtigen
        Ter_Vor False, RmuNr, FreTe
    Else
        Ter_Vor False, 0, FreTe
    End If
Else
    If ChRau.Value = xtpChecked Then
        Ter_Vor True, RmuNr, FreTe
    Else
        Ter_Vor True, 0, FreTe
    End If
End If

FStat

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FVors " & Err.Number
Resume Next

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlTeF = False Then 'Formular wird geladen
        If Control.id = XTP_ID_RIBBONCONTROLTAB Then
            FTabu
        Else
            FTool Control.id
        End If
    End If
End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmTermVo
Set CmBrs = FM.comBar02
Set Rahm8 = FM.frmRahm8
Set Rahm9 = FM.frmRahm9
Set RpCo1 = FM.repCont1
Set RpCo6 = FM.repCont6
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case RbTab.id
Case RibTab_Ter_Haupt:
    Rahm9.Visible = True
    Rahm8.Visible = True
    RpCo6.Visible = True
    RpCo1.Visible = False
Case RibTab_Ter_Leist:
    Rahm9.Visible = False
    Rahm8.Visible = False
    RpCo6.Visible = False
    RpCo1.Visible = True
    DoEvents
    TeVoSu
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set RpCo1 = Nothing
Set RpCo6 = Nothing
Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub txtBetre_Change()

TagWe = Mid$(Me.txtBetre.Tag, 2, Len(Me.txtBetre.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtBetre.Tag = 1 & TagWe
    GlTSa = True
    FBetr
End If

End Sub
Private Sub txtBetre_Click()

TagWe = Mid$(Me.txtBetre.Tag, 2, Len(Me.txtBetre.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtBetre.Tag = 1 & TagWe
    GlTSa = True
    FBetr
End If

End Sub

Private Sub txtBetre_GotFocus()
    GlTeF = False 'Formular wird geladen
End Sub

Private Sub txtBetre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
    GlTeF = False 'Formular wird geladen
End Sub

Private Sub txtKomme_Change()

TagWe = Mid$(Me.txtKomme.Tag, 2, Len(Me.txtKomme.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtKomme.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtKomme_GotFocus()
    Me.txtKomme.SelStart = 0
    Me.txtKomme.SelLength = Len(Me.txtKomme.Text)
    GlTeF = False
End Sub

Private Sub txtKomme_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtKomme.SelLength = 0
End Sub

Private Sub cmbRaum1_Click()

TagWe = Mid$(Me.cmbRaum1.Tag, 2, Len(Me.cmbRaum1.Tag) - 1)

If GlTeF = False Then
    Me.cmbRaum1.Tag = 1 & TagWe
    GlTSa = True
End If

FTeLo

End Sub

Private Sub cmbRaum1_GotFocus()
    RetWe = SendMessage(Me.cmbRaum1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbRaum1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub updCont2_DownClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim ManNr As Long
Dim MitNr As Long
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MitNr = CmMit.ListIndex + 1
        Else
            MitNr = GlSmI
        End If
    Else
        MitNr = GlSmI
    End If
    If GlMiA(MitNr, 8) > 0 Then
        ZeiRa = GlMiA(MitNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            ManNr = CmMan.ListIndex + 1
        Else
            ManNr = GlSMa
        End If
    Else
        ManNr = GlSMa
    End If
    If GlMan(ManNr, 8) > 0 Then
        ZeiRa = GlMan(ManNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

Select Case ZeiRa
Case 1: MiDif = 10
Case 2: MiDif = 15
Case 3: MiDif = 20
Case 4: MiDif = 30
Case 5: MiDif = 45
Case 6: MiDif = 50
Case 7: MiDif = 60
Case 8: MiDif = 90
End Select

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then ZeiVo = 15

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
        TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
        
        TmVon = DateAdd("n", -MiDif, AlDa1)
        VoZei.Text = Format$(TmVon, "hh:mm")
        VoZei.Tag = 1 & TagWe
        GlTSa = True
        
        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmBis = DateAdd("n", ZeiVo, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
                TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)
                BiZei.Tag = 1 & TagWe
            End If
        End If
    End If
End If

End Sub
Private Sub updCont2_UpClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim ManNr As Long
Dim MitNr As Long
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MitNr = CmMit.ListIndex + 1
        Else
            MitNr = GlSmI
        End If
    Else
        MitNr = GlSmI
    End If
    If GlMiA(MitNr, 8) > 0 Then
        ZeiRa = GlMiA(MitNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            ManNr = CmMan.ListIndex + 1
        Else
            ManNr = GlSMa
        End If
    Else
        ManNr = GlSMa
    End If
    If GlMan(ManNr, 8) > 0 Then
        ZeiRa = GlMan(ManNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

Select Case ZeiRa
Case 1: MiDif = 10
Case 2: MiDif = 15
Case 3: MiDif = 20
Case 4: MiDif = 30
Case 5: MiDif = 45
Case 6: MiDif = 50
Case 7: MiDif = 60
Case 8: MiDif = 90
End Select

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then ZeiVo = 15

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
        TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
        
        TmVon = DateAdd("n", MiDif, AlDa1)
        VoZei.Text = Format$(TmVon, "hh:mm")
        VoZei.Tag = 1 & TagWe
        GlTSa = True
        
        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmBis = DateAdd("n", ZeiVo, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
                TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)
                BiZei.Tag = 1 & TagWe
            End If
        Else
            If AlDa1 >= AlDa2 Then
                TmBis = DateAdd("n", MiDif, AlDa1)
                BiZei.Text = Format$(TmBis, "hh:mm")
                TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)
                BiZei.Tag = 1 & TagWe
            End If
        End If
    End If
End If

End Sub
Private Sub updCont3_DownClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim ManNr As Long
Dim MitNr As Long
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MitNr = CmMit.ListIndex + 1
        Else
            MitNr = GlSmI
        End If
    Else
        MitNr = GlSmI
    End If
    If GlMiA(MitNr, 8) > 0 Then
        ZeiRa = GlMiA(MitNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            ManNr = CmMan.ListIndex + 1
        Else
            ManNr = GlSMa
        End If
    Else
        ManNr = GlSMa
    End If
    If GlMan(ManNr, 8) > 0 Then
        ZeiRa = GlMan(ManNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

Select Case ZeiRa
Case 1: MiDif = 10
Case 2: MiDif = 15
Case 3: MiDif = 20
Case 4: MiDif = 30
Case 5: MiDif = 45
Case 6: MiDif = 50
Case 7: MiDif = 60
Case 8: MiDif = 90
End Select

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then ZeiVo = 15

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
        TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)
        
        TmBis = DateAdd("n", -MiDif, AlDa2)
        BiZei.Text = Format$(TmBis, "hh:mm")
        BiZei.Tag = 1 & TagWe
        GlTSa = True
        
        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmVon = DateAdd("n", -ZeiVo, TmBis)
                VoZei.Text = Format$(TmVon, "hh:mm")
                TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
                VoZei.Tag = 1 & TagWe
            End If
        Else
            If AlDa1 >= AlDa2 Then
                TmVon = DateAdd("n", -MiDif, AlDa2)
                VoZei.Text = Format$(TmVon, "hh:mm")
                TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
                VoZei.Tag = 1 & TagWe
            End If
        End If
    End If
End If

End Sub
Private Sub updCont3_UpClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim ManNr As Long
Dim MitNr As Long
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MitNr = CmMit.ListIndex + 1
        Else
            MitNr = GlSmI
        End If
    Else
        MitNr = GlSmI
    End If
    If GlMiA(MitNr, 8) > 0 Then
        ZeiRa = GlMiA(MitNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            ManNr = CmMan.ListIndex + 1
        Else
            ManNr = GlSMa
        End If
    Else
        ManNr = GlSMa
    End If
    If GlMan(ManNr, 8) > 0 Then
        ZeiRa = GlMan(ManNr, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

Select Case ZeiRa
Case 1: MiDif = 10
Case 2: MiDif = 15
Case 3: MiDif = 20
Case 4: MiDif = 30
Case 5: MiDif = 45
Case 6: MiDif = 50
Case 7: MiDif = 60
Case 8: MiDif = 90
End Select

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then ZeiVo = 15

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
        TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)
        
        TmBis = DateAdd("n", MiDif, AlDa2)
        BiZei.Text = Format$(TmBis, "hh:mm")
        BiZei.Tag = 1 & TagWe
        GlTSa = True
        
        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmVon = DateAdd("n", -ZeiVo, TmBis)
                VoZei.Text = Format$(TmVon, "hh:mm")
                TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
                VoZei.Tag = 1 & TagWe
            End If
        End If
    End If
End If

End Sub
Private Sub chkFreTe_Click()
    
Set ChTer = Me.chkFreTe
Set CmRmu = Me.cmbRaum1
    
If GlSeF = False Then
    If ChTer.Value = xtpChecked Then
        IniSetVal "TerSys", "TerSer", -1
    Else
        IniSetVal "TerSys", "TerSer", 0
    End If
    
    If CmRmu.ItemData(CmRmu.ListIndex) = 0 Then
        CmRmu.ListIndex = 0
    End If
End If
    
FTeLo
    
End Sub
Private Sub txtDatu4_LostFocus()
On Error Resume Next

Dim Datu1 As Date
Dim Datu2 As Date
Dim LiIdx As Integer
    
Set ZyEnT = Me.cmbZyEn1
Set TxDa1 = Me.txtDatu1
Set TxDa4 = Me.txtDatu4
  
Datu1 = CDate(TxDa1.Text)
Datu2 = CDate(TxDa4.Text)

KalWa = 4
FDaKo

If Datu2 < Datu1 Then
    Datu2 = DateAdd("d", ZyEnT.ItemData(ZyEnT.ListIndex), Datu1)
    TxDa4.Text = Datu2
End If

If GlSeF = False Then
    TeVoSt
End If

End Sub
Private Sub chkDopTe_Click()

Set ChSpl = Me.chkTeSpl
Set ChDop = Me.chkDopTe

TeVoSt
FDoTe
FTeLo

End Sub
Private Sub dtpDatu3_SelectionChanged()
    If GlSeF = False Then
        FDatu
        TeVoSt
    End If
End Sub

Private Sub btnDatu4_Click()
    KalWa = 4
    FKale
End Sub
Private Sub optZykl1_Click()
    If GlAkt = False Then
        GlTvM = "M1"
        FOpt
        FTeLo
        TeVoSt
    End If
End Sub
Private Sub optZykl2_Click()
    If GlAkt = False Then
        GlTvM = "M2"
        FOpt
        FTeLo
        TeVoSt
    End If
End Sub
Private Sub optZykl3_Click()
    If GlAkt = False Then
        GlTvM = "M3"
        FOpt
        FTeLo
        TeVoSt
    End If
End Sub
Private Sub optZykl4_Click()
    If GlAkt = False Then
        GlTvM = "M4"
        FOpt
        FTeLo
        TeVoSt
    End If
End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub
Private Sub choTaDin_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub choTaDon_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub choTaFre_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub choTaMit_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub choTaMon_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub choTaSam_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub choTaSon_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmbJahr1_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmbMona1_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmbMona3_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmbMonat_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmbWoche_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmbZyEn1_Click()

Dim NeuDa As Date
Dim AnzVo As Integer

Set ZyEnT = Me.cmbZyEn1
Set TxDa1 = Me.txtDatu1
Set TxDa4 = Me.txtDatu4

AnzVo = ZyEnT.ItemData(ZyEnT.ListIndex)

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

TxDa4.Text = DateAdd("d", AnzVo, NeuDa)

IniSetVal "TerSys", "AnzVor", ZyEnT.ListIndex

FTeLo

End Sub
Private Sub cmoJahr1_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmoJahr2_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmoJahr3_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmoJahr4_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmoMona1_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub cmoMona2_Click()
    If GlSeF = False Then TeVoSt
End Sub
Private Sub dtpDatu3_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)

Dim AktTa As Long
Dim AktZa As Integer

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTeV = True Then 'Termine vorhanden
    If GlTpV = True Then 'Kalendermarker vorhanden
        For AktTa = 0 To GlKMa - 1 'Anzahl Kalendermatker
            If Day = Left$(GlTEr(0, AktTa), 10) Then
                For AktZa = 1 To UBound(GlTep) 'Kalendermarker
                    If GlTep(AktZa, 0) = GlTEr(1, AktTa) Then
                        Metrics.BackColor = GlTep(AktZa, 2)
                        Exit For
                    End If
                Next AktZa
            End If
        Next AktTa
    End If
End If

End Sub
Private Sub repCont6_InplaceButtonDown(ByVal Button As XtremeReportControl.IReportInplaceButton)
    If Button.Column.ItemIndex = 2 Then
        FKaRo Button
    End If
End Sub
Private Sub repCont6_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

If Row.GroupRow = False Then
    If Row.Record(6).Value <> vbNullString Then
        If IsNumeric(Row.Record(6).Value) = True Then
            FrbZa = Row.Record(6).Value
            If FrbZa > 0 Then
                Metrics.BackColor = FrbZa
            End If
        End If
    End If
End If

End Sub
Private Sub repCont6_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Item.Index = 5 Then
        FStat
    End If
End Sub
Private Sub repCont6_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    FKrCa Column.ItemIndex
End Sub
Private Sub optZyEn2_Click()
    FTeLo
    FSet
End Sub
Private Sub optZyEn3_Click()
    FTeLo
    FSet
End Sub
Private Sub optZyTa1_Click()
    If GlSeF = False Then
        TeVoSt
        FTeLo
    End If
End Sub
Private Sub optZyTa2_Click()
    Me.optZykl1.Value = True
    If GlSeF = False Then
        TeVoSt
        FTeLo
    End If
End Sub
