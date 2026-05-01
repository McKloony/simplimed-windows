VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmBuSer 
   Caption         =   "Serienbuchung"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14955
   ControlBox      =   0   'False
   Icon            =   "frmBuSer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   14955
   Begin XtremeReportControl.ReportControl repCont6 
      Height          =   1500
      Left            =   1320
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6000
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
      TabIndex        =   18
      Top             =   100
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.FlatEdit txoDummy 
      Height          =   200
      Left            =   200
      TabIndex        =   0
      Top             =   12000
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
   Begin XtremeSuiteControls.GroupBox frmRahm8 
      Height          =   4200
      Left            =   6840
      TabIndex        =   25
      Top             =   960
      Width           =   7200
      _Version        =   1048579
      _ExtentX        =   12700
      _ExtentY        =   7408
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox frmRahm6 
         Height          =   840
         Left            =   100
         TabIndex        =   24
         Top             =   2100
         Width           =   7000
         _Version        =   1048579
         _ExtentX        =   12347
         _ExtentY        =   1482
         _StockProps     =   79
         Caption         =   "Seriendauer"
         UseVisualStyle  =   -1  'True
         Begin XtremeCalendarControl.DatePicker dtpDatu3 
            Height          =   405
            Left            =   6500
            TabIndex        =   33
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
            Left            =   6130
            TabIndex        =   31
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
            Left            =   3800
            TabIndex        =   28
            Top             =   350
            Width           =   1000
            _Version        =   1048579
            _ExtentX        =   1764
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Endet am :"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZyEn2 
            Height          =   220
            Left            =   400
            TabIndex        =   26
            Top             =   350
            Width           =   1200
            _Version        =   1048579
            _ExtentX        =   2117
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Endet nach :"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDatu4 
            Height          =   350
            Left            =   4900
            TabIndex        =   29
            Top             =   300
            Width           =   1200
            _Version        =   1048579
            _ExtentX        =   2117
            _ExtentY        =   617
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            BackColor       =   16777215
            Alignment       =   2
         End
         Begin XtremeSuiteControls.ComboBox cmbZyEn1 
            Height          =   310
            Left            =   1600
            TabIndex        =   27
            Top             =   300
            Width           =   1500
            _Version        =   1048579
            _ExtentX        =   2646
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            Style           =   2
            Text            =   "ComboBox1"
         End
      End
      Begin XtremeSuiteControls.GroupBox frmRahm5 
         Height          =   1710
         Left            =   100
         TabIndex        =   23
         Top             =   200
         Width           =   7000
         _Version        =   1048579
         _ExtentX        =   12347
         _ExtentY        =   3016
         _StockProps     =   79
         Caption         =   "Serienmuster"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.GroupBox frmRahm1 
            Height          =   1545
            Left            =   1900
            TabIndex        =   19
            Top             =   120
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2725
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            BorderStyle     =   2
            Begin XtremeSuiteControls.FlatEdit txoTage1 
               Height          =   350
               Left            =   870
               TabIndex        =   30
               Top             =   150
               Width           =   500
               _Version        =   1048579
               _ExtentX        =   882
               _ExtentY        =   617
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Text            =   "1"
               BackColor       =   16777215
               Alignment       =   2
            End
            Begin XtremeSuiteControls.RadioButton optZyTa2 
               Height          =   220
               Left            =   150
               TabIndex        =   32
               Top             =   560
               Width           =   1600
               _Version        =   1048579
               _ExtentX        =   2822
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Jeden Arbeitstag"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton optZyTa1 
               Height          =   220
               Left            =   150
               TabIndex        =   35
               Top             =   200
               Width           =   600
               _Version        =   1048579
               _ExtentX        =   1058
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Alle"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
            Begin VB.Label Lab14 
               BackStyle       =   0  'Transparent
               Caption         =   "Tag(e)"
               Height          =   220
               Left            =   1460
               TabIndex        =   36
               Top             =   210
               Width           =   600
            End
         End
         Begin XtremeSuiteControls.GroupBox frmRahm2 
            Height          =   1545
            Left            =   1900
            TabIndex        =   20
            Top             =   120
            Visible         =   0   'False
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2725
            _StockProps     =   79
            Appearance      =   6
            BorderStyle     =   2
            Begin XtremeSuiteControls.CheckBox choTaSam 
               Height          =   220
               Left            =   2600
               TabIndex        =   37
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
               TabIndex        =   38
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
               TabIndex        =   39
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
               TabIndex        =   40
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
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   43
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
               TabIndex        =   44
               Top             =   150
               Width           =   1900
               _Version        =   1048579
               _ExtentX        =   3334
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Text            =   "ComboBox1"
            End
         End
         Begin XtremeSuiteControls.GroupBox frmRahm3 
            Height          =   1545
            Left            =   1900
            TabIndex        =   21
            Top             =   120
            Visible         =   0   'False
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2725
            _StockProps     =   79
            Appearance      =   6
            BorderStyle     =   2
            Begin XtremeSuiteControls.RadioButton optZyMo2 
               Height          =   220
               Left            =   150
               TabIndex        =   45
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
               TabIndex        =   46
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
               TabIndex        =   47
               Top             =   700
               Width           =   1000
               _Version        =   1048579
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Text            =   "ComboBox5"
            End
            Begin XtremeSuiteControls.ComboBox cmoMona2 
               Height          =   310
               Left            =   1860
               TabIndex        =   48
               Top             =   700
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Text            =   "ComboBox6"
            End
            Begin XtremeSuiteControls.ComboBox cmbMonat 
               Height          =   310
               Left            =   1860
               TabIndex        =   49
               Top             =   150
               Width           =   1800
               _Version        =   1048579
               _ExtentX        =   3175
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cmbMona3 
               Height          =   315
               Left            =   3120
               TabIndex        =   50
               Top             =   700
               Width           =   1800
               _Version        =   1048579
               _ExtentX        =   3175
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cmbMona1 
               Height          =   310
               Left            =   800
               TabIndex        =   51
               Top             =   150
               Width           =   1000
               _Version        =   1048579
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Text            =   "ComboBox1"
            End
         End
         Begin XtremeSuiteControls.RadioButton optZykl4 
            Height          =   220
            Left            =   400
            TabIndex        =   52
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
            TabIndex        =   53
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
            TabIndex        =   54
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
            TabIndex        =   55
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
            Height          =   1545
            Left            =   1900
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   5000
            _Version        =   1048579
            _ExtentX        =   8819
            _ExtentY        =   2725
            _StockProps     =   79
            Appearance      =   6
            BorderStyle     =   2
            Begin XtremeSuiteControls.RadioButton optZyJa2 
               Height          =   220
               Left            =   150
               TabIndex        =   56
               Top             =   750
               Width           =   740
               _Version        =   1048579
               _ExtentX        =   1305
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Am"
               Appearance      =   6
            End
            Begin XtremeSuiteControls.RadioButton optZyJa1 
               Height          =   220
               Left            =   150
               TabIndex        =   57
               Top             =   200
               Width           =   740
               _Version        =   1048579
               _ExtentX        =   1305
               _ExtentY        =   388
               _StockProps     =   79
               Caption         =   "Jeden"
               Appearance      =   6
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.ComboBox cmoJahr1 
               Height          =   310
               Left            =   2100
               TabIndex        =   58
               Top             =   150
               Width           =   1300
               _Version        =   1048579
               _ExtentX        =   2302
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cmoJahr2 
               Height          =   310
               Left            =   1000
               TabIndex        =   59
               Top             =   700
               Width           =   1000
               _Version        =   1048579
               _ExtentX        =   1746
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox2"
            End
            Begin XtremeSuiteControls.ComboBox cmoJahr3 
               Height          =   310
               Left            =   2100
               TabIndex        =   60
               Top             =   700
               Width           =   1300
               _Version        =   1048579
               _ExtentX        =   2302
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox3"
            End
            Begin XtremeSuiteControls.ComboBox cmoJahr4 
               Height          =   310
               Left            =   3735
               TabIndex        =   61
               Top             =   700
               Width           =   1200
               _Version        =   1048579
               _ExtentX        =   2117
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
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
               BackColor       =   -2147483643
               BackColor       =   -2147483643
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "im"
               Height          =   220
               Left            =   3500
               TabIndex        =   63
               Top             =   720
               Width           =   200
            End
         End
      End
      Begin XtremeSuiteControls.GroupBox frmRahm7 
         Height          =   840
         Left            =   100
         TabIndex        =   73
         Top             =   3120
         Width           =   7000
         _Version        =   1048579
         _ExtentX        =   12347
         _ExtentY        =   1482
         _StockProps     =   79
         Caption         =   "Mitarbeiter"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cmbMitar 
            Height          =   315
            Left            =   1600
            TabIndex        =   34
            Top             =   300
            Width           =   2925
            _Version        =   1048579
            _ExtentX        =   5159
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            Style           =   2
            Text            =   "ComboBox6"
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Mitarbeiter :"
            Height          =   210
            Left            =   400
            TabIndex        =   74
            Top             =   350
            Width           =   1100
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm9 
      Height          =   5000
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6600
      _Version        =   1048579
      _ExtentX        =   11642
      _ExtentY        =   8819
      _StockProps     =   79
      Caption         =   "Termindaten"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbKtoRa 
         Height          =   315
         Left            =   2700
         TabIndex        =   10
         Top             =   2530
         Width           =   2920
         _Version        =   1048579
         _ExtentX        =   5159
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGegen 
         Height          =   315
         Left            =   900
         TabIndex        =   9
         Top             =   2530
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox5"
      End
      Begin XtremeSuiteControls.FlatEdit txtHaben 
         Height          =   310
         Left            =   900
         TabIndex        =   8
         Top             =   2530
         Width           =   4740
         _Version        =   1048579
         _ExtentX        =   8361
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtKonto 
         Height          =   350
         Left            =   900
         TabIndex        =   6
         Top             =   1130
         Width           =   4740
         _Version        =   1048579
         _ExtentX        =   8361
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu1 
         Height          =   495
         Left            =   0
         TabIndex        =   64
         Top             =   2600
         Visible         =   0   'False
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   64
         Show3DBorder    =   2
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   2130
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   430
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbBuTyp 
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   3230
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbBuStu 
         Height          =   315
         Left            =   2700
         TabIndex        =   12
         Top             =   3230
         Width           =   2920
         _Version        =   1048579
         _ExtentX        =   5159
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox cmbBeha4 
         Height          =   315
         Left            =   2700
         TabIndex        =   14
         Top             =   3930
         Width           =   2920
         _Version        =   1048579
         _ExtentX        =   5159
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox6"
      End
      Begin XtremeSuiteControls.ComboBox cmbWarun 
         Height          =   315
         Left            =   4240
         TabIndex        =   5
         Top             =   430
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbBuTex 
         Height          =   315
         Left            =   900
         TabIndex        =   7
         Top             =   1830
         Width           =   4740
         _Version        =   1048579
         _ExtentX        =   8361
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   900
         TabIndex        =   2
         Top             =   430
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBuBet 
         Height          =   350
         Left            =   2700
         TabIndex        =   4
         Top             =   430
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu5 
         Height          =   350
         Left            =   900
         TabIndex        =   13
         Top             =   3930
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkGewEr 
         Height          =   225
         Left            =   2705
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4500
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Keine Auswertung bei Erlösermittlung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   210
         Left            =   905
         TabIndex        =   79
         Top             =   2300
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Geldkonto :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   210
         Left            =   905
         TabIndex        =   78
         Top             =   880
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Sachkonto :"
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab35 
         BackStyle       =   0  'Transparent
         Caption         =   "Startdatum :"
         Height          =   210
         Left            =   905
         TabIndex        =   75
         Top             =   3680
         Width           =   1100
      End
      Begin VB.Label lblLab59 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   210
         Left            =   2705
         TabIndex        =   71
         Top             =   3680
         Width           =   900
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Beginnt am :"
         Height          =   210
         Left            =   905
         TabIndex        =   70
         Top             =   200
         Width           =   900
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   "Währung :"
         Height          =   210
         Left            =   4245
         TabIndex        =   69
         Top             =   200
         Width           =   900
      End
      Begin VB.Label lblLab34 
         BackStyle       =   0  'Transparent
         Caption         =   "Buchungstyp :"
         Height          =   210
         Left            =   905
         TabIndex        =   68
         Top             =   3000
         Width           =   1100
      End
      Begin VB.Label lblLab04 
         BackStyle       =   0  'Transparent
         Caption         =   "Buchungstext :"
         Height          =   210
         Left            =   905
         TabIndex        =   67
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label lblLab02 
         BackStyle       =   0  'Transparent
         Caption         =   "Betrag :"
         Height          =   210
         Left            =   2705
         TabIndex        =   66
         Top             =   200
         Width           =   900
      End
      Begin VB.Label lblLab33 
         BackStyle       =   0  'Transparent
         Caption         =   "Steuersatz :"
         Height          =   210
         Left            =   2705
         TabIndex        =   65
         Top             =   3000
         Width           =   1100
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtBezei 
      Height          =   200
      Left            =   1005
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   11955
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoNr 
      Height          =   200
      Left            =   600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   11955
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoHa 
      Height          =   200
      Left            =   1440
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   11955
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtBezHa 
      Height          =   200
      Left            =   1845
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   11955
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtIdxNr 
      Height          =   200
      Left            =   2400
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   12000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Left            =   0
      Top             =   480
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBuSer"
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
Private Rahm8 As XtremeSuiteControls.GroupBox
Private Rahm9 As XtremeSuiteControls.GroupBox

Private TxID0 As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa4 As XtremeSuiteControls.FlatEdit
Private TxDa5 As XtremeSuiteControls.FlatEdit
Private TxBuB As XtremeSuiteControls.FlatEdit
Private TxKto As XtremeSuiteControls.FlatEdit
Private TxHab As XtremeSuiteControls.FlatEdit
Private ChAsw As XtremeSuiteControls.CheckBox

Private CmRam As XtremeSuiteControls.ComboBox
Private FeWar As XtremeSuiteControls.ComboBox
Private FeGeg As XtremeSuiteControls.ComboBox
Private CmBuT As XtremeSuiteControls.ComboBox
Private CmBuS As XtremeSuiteControls.ComboBox
Private CmBe4 As XtremeSuiteControls.ComboBox
Private CmBTe As XtremeSuiteControls.ComboBox
Private ChTer As XtremeSuiteControls.CheckBox
Private ChMon As XtremeSuiteControls.CheckBox
Private CmRem As XtremeSuiteControls.ComboBox
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
Private KntRa As Integer
Private TabId As Integer
Private FoLad As Boolean

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
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50151)
TeMai = IniGetOpt("Hilfe", 50152)
TeInh = IniGetOpt("Hilfe", 50153)
TeFus = IniGetOpt("Hilfe", 50154)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Function FKaSh(ByVal KalLi As Long, ByVal KalOb As Long, ByVal NeuDa As Date, ByVal mHwnd As Long, Optional ByVal Flag As Boolean = False) As Date
On Error GoTo LaErr

Dim Datu1 As Date
Dim DayFi As Date
Dim DayLa As Date
Dim KaBre As Long
Dim KaHoh As Long
Dim RetWe As Boolean

Set FM = frmBuSer
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

Set FM = frmBuSer
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns
Set RpSel = RpCo6.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If RpRow.Record(1).Value <> vbNullString Then
            AltDa = CDate(RpRow.Record(1).Value)
        Else
            AltDa = Date
        End If
        RpBut.GetRect ItmLi, ItmOb, ItmRe, ItmHo
        ItmBr = ItmRe
        ItmTo = ItmHo + 1
        NeuDa = FKaSh(ItmBr, ItmTo, AltDa, RpCo6.hwnd, True)
        If IsDate(NeuDa) Then
            RpCo6.EditItem Nothing, Nothing
            RpRow.Record(1).Value = CDate(NeuDa)
            RpCo6.Populate
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
Private Sub FKrCa(ByVal CoIdx As Long)
On Error GoTo PoErr
'Änderungen im Krankneblatt

Dim NeuDa As Date
Dim BuBet As Single
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmBuSer
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns
Set RpSel = RpCo6.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Select Case CoIdx
        Case 1:
            Set RpCol = RpCls.Find(1)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                NeuDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                RpRow.Record(RpCol.ItemIndex).Value = NeuDa
                Set RpCol = RpCls.Find(1)
                RpRow.Record(RpCol.ItemIndex).Value = Format$(NeuDa, "dddd")
            End If
            DoEvents
        Case 3:
            Set RpCol = RpCls.Find(3)
            If Not IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then
                BuBet = CSng(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                RpRow.Record(RpCol.ItemIndex).Value = Format$(BuBet, GlWa1)
            End If
            DoEvents
        Case 4:
            Set RpCol = RpCls.Find(4)
            If Not IsNull(RpRow.Record(RpCol.ItemIndex).Value) Then
                BuBet = CSng(Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1))
                RpRow.Record(RpCol.ItemIndex).Value = Format$(BuBet, GlWa1)
            End If
            DoEvents
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
Private Sub FMand()
On Error GoTo OrErr

Dim ManNr As Long
Dim StaRa As Integer
Dim AktZa As Integer

Set CmBe4 = Me.cmbBeha4
Set CmRam = Me.cmbKtoRa

ManNr = CmBe4.ItemData(CmBe4.ListIndex)

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            If GlMan(AktZa, 25) <> vbNullString Then
                KntRa = GlMan(AktZa, 25) 'Standardkontenrahmen
            Else
                KntRa = GlKtR
            End If
            Exit For
        End If
    Next AktZa
    CmRam.ListIndex = KntRa - 1
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMand " & Err.Number
Resume Next

End Sub

Private Sub FOpt(ByVal Flag As Integer)
On Error Resume Next

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4

Select Case Flag
Case 1: Rahm1.Visible = True
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
Case 2: Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
Case 3: Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = True
        Rahm4.Visible = False
Case 4: Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = True
End Select

End Sub
Private Sub FTeLo()
On Error GoTo KoErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmBuSer
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

FStat

Set CmAcs = Nothing
Set RpRcs = Nothing
Set CmBrs = Nothing
Set RpCo6 = Nothing

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeLo " & Err.Number
Resume Next

End Sub

Private Sub cmbBeha4_Click()
    If FoLad = False Then
        FMand
    End If
End Sub

Private Sub cmbKtoRa_Click()

Set CmRam = Me.cmbKtoRa
    
If FoLad = False Then
    KntRa = CmRam.ListIndex + 1
End If

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
        SeBuSt
    End If
End Sub
Private Sub Form_Load()

FoLad = True

Set FrmEx = Me.frmExtde
Set CmRam = Me.cmbKtoRa

With FrmEx
    .ClientMaxHeight = 12000
    .ClientMaxWidth = 13820
    .ClientMinHeight = 8000
    .ClientMinWidth = 13820
    .TopMost = True
End With

KntRa = GlKtR

FoLad = False

Set FrmEx = Nothing

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
    Select Case KeyCode
    Case vbKeyF2: Me.txtDatu1.SelLength = 0
    Case vbKeyDown: Me.txtBuBet.SetFocus
    Case vbKeyUp: Me.txtKonto.SetFocus
    End Select
End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub
Private Sub txtDatu1_Change()

If GlTeF = False Then 'Formular wird geladen
    Me.txtDatu1.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlTeF = False Then 'Formular wird geladen
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    SeBuPo
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub FClos()
On Error GoTo LiErr

If WindowLoad("frmBuSer") = True Then
    Unload frmBuSer
End If

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlRes = False Then 'Reset der Einstellungen
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "BuSeVo", "FenLin", clFen.FeLin
        IniSetVal "BuSeVo", "FenObe", clFen.FeObn
        IniSetVal "BuSeVo", "FenBre", clFen.FeBre
        IniSetVal "BuSeVo", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmBuSer = Nothing
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

Set CmAcs = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FStat()
On Error GoTo KoErr

Dim GesZa As Integer
Dim TerZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmBuSer
Set CmBrs = FM.comBar02
Set RpCo6 = FM.repCont6
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RpRcs = RpCo6.Records

GesZa = RpRcs.Count

If GesZa > 0 Then
    For Each RpRec In RpRcs
        If RpRec.Item(Buh_RechNr).Checked = True Then
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
Private Sub FTool(ByVal TolId As Long)

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: FVors
Case KY_F5: frmAdrSuch.Show vbModal
Case KY_F7: FTeLo
Case KY_F8: FSave
Case KY_F10:
Case KY_F11: Unload Me
Case AD_Termin_Vorschau: FVors
Case AD_Termin_Save: FSave
Case AD_Termin_Reset: FTeLo
Case TE_Termin_Beenden: Unload Me
Case TE_Adresse_Hinzufu: SAdre 1
Case TE_Adresse_Bearbeit:
Case TE_Adresse_Suchen: frmAdrSuch.Show vbModal
Case TE_Termin_Hilfe: FHilfe
Case AD_TerSpei_Norma: FSave TolId
Case AD_Termin_Close: FSave TolId
End Select

End Sub
Private Sub FSave(Optional ByVal TolId As Long)
On Error GoTo SaErr
'Überprüft, ob der Eintrag geändert wurde und speichert dieses ab

Ser_VoS TolId
DoEvents

Select Case TolId
Case 20121:
    SUpBu
Case 20122:
    K_BuVpl "BuSe"
    P_List "BuSe", 0, 1
Case Else:
    P_List "BuSe", 0, 2
End Select

DoEvents
Unload Me

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FVors()
On Error GoTo InErr

If Me.txtBuBet.Text <> vbNullString Then
    If IsNumeric(Me.txtBuBet.Text) = True Then
        If CDbl(Me.txtBuBet.Text) > 0 Then
            If Me.txtKtoNr.Text <> vbNullString Then
                If IsNumeric(Me.txtKtoNr.Text) = True Then
                    If CLng(Me.txtBuBet.Text) > 0 Then
                        If Me.txtBezei.Text <> vbNullString Then
                            If Me.cmbBuTex.Text <> vbNullString Then
                                Ser_Vor
                            Else
                                SPopu "Kein Buchungstext", "Es wurde kein Buchungstext eingegeben.", IC48_Forbidden
                            End If
                        Else
                            SPopu "Kein Sachkonto", "Es wurde kein Sachkonto eingegeben.", IC48_Forbidden
                        End If
                    Else
                        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
                    End If
                Else
                    SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
                End If
            Else
                SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
            End If
        Else
            SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
        End If
    Else
        SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
    End If
Else
    SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
End If

FStat

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FVors " & Err.Number
Resume Next

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlTeF = False Then 'Formular wird geladen
        FTool Control.id
    End If
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

LiIdx = Datu2 - Datu1 - 2
RetWe = SendMessage(ZyEnT.hwnd, CB_SETCURSEL, LiIdx, ByVal 0&)

If GlSeF = False Then SeBuSt

End Sub
Private Sub dtpDatu3_SelectionChanged()
    If GlSeF = False Then
        FDatu
        SeBuSt
    End If
End Sub

Private Sub btnDatu4_Click()
    KalWa = 4
    FKale
End Sub
Private Sub optZykl1_Click()
    If GlAkt = False Then
        FOpt 1
        FTeLo
        SeBuSt
    End If
End Sub
Private Sub optZykl2_Click()
    If GlAkt = False Then
        FOpt 2
        FTeLo
        SeBuSt
    End If
End Sub
Private Sub optZykl3_Click()
    If GlAkt = False Then
        FOpt 3
        FTeLo
        SeBuSt
    End If
End Sub
Private Sub optZykl4_Click()
    If GlAkt = False Then
        FOpt 4
        FTeLo
        SeBuSt
    End If
End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub
Private Sub choTaDin_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub choTaDon_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub choTaFre_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub choTaMit_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub choTaMon_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub choTaSam_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub choTaSon_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmbJahr1_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmbMona1_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmbMona3_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmbMonat_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmbWoche_Click()
    If GlSeF = False Then SeBuSt
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
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmoJahr2_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmoJahr3_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmoJahr4_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmoMona1_Click()
    If GlSeF = False Then SeBuSt
End Sub
Private Sub cmoMona2_Click()
    If GlSeF = False Then SeBuSt
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
    If Button.Column.ItemIndex = 1 Then
        FKaRo Button
    End If
End Sub
Private Sub repCont6_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

If Row.GroupRow = False Then
    If IsNumeric(Row.Record(7).Value) Then
        FrbZa = Row.Record(7).Value
        If FrbZa > 0 Then
            Metrics.BackColor = FrbZa
        End If
    End If
End If

If CBool(GlGeK(Row.Record(Buh_IDB).Value, 5)) = True Then
    Metrics.ForeColor = 16711680
End If

End Sub
Private Sub repCont6_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Item.Index = 8 Then FStat
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
        SeBuSt
        FTeLo
    End If
End Sub
Private Sub optZyTa2_Click()
    Me.optZykl1.Value = True
    If GlSeF = False Then
        SeBuSt
        FTeLo
    End If
End Sub
Private Sub cmbBeha4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBeha4_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.txtKonto.SetFocus
    Case vbKeyUp: Me.txtKonto.SetFocus
    End Select
End Sub
Private Sub cmbBuTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuStu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuStu_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.cmbGegen.SetFocus
    Case vbKeyUp: Me.cmbBuTex.SetFocus
    End Select
End Sub
Private Sub cmbBuTyp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuTyp_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.cmbBuTex.SetFocus
    Case vbKeyUp: Me.cmbWarun.SetFocus
    End Select
End Sub

Private Sub cmbGegen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbGegen_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.txtKonto.SetFocus
    Case vbKeyUp: Me.cmbBuStu.SetFocus
    End Select
End Sub
Private Sub cmbWarun_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbWarun_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: Me.cmbBuTyp.SetFocus
    Case vbKeyUp: Me.txtBuBet.SetFocus
    End Select
End Sub
Private Sub txtBuBet_GotFocus()
    Me.txtBuBet.SelStart = 0
    Me.txtBuBet.SelLength = Len(Me.txtBuBet.Text)
End Sub
Private Sub txtBuBet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBuBet_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBuBet.SelLength = 0
    Case vbKeyDown: Me.cmbWarun.SetFocus
    Case vbKeyUp: Me.txtDatu1.SetFocus
    End Select
End Sub
Private Sub txtBuBet_LostFocus()

Dim Betra As Single

If Me.txtBuBet.Text <> vbNullString Then
    If IsNumeric(Me.txtBuBet.Text) Then
        Betra = Me.txtBuBet.Text
        Me.txtBuBet.Text = Format$(Betra, GlWa1)
    End If
End If

End Sub

Private Sub txtHaben_GotFocus()
    Me.txtHaben.SelStart = 0
    Me.txtHaben.SelLength = Len(Me.txtHaben.Text)
End Sub

Private Sub txtHaben_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn: KeyAscii = 0
    Case vbKeyTab: KeyAscii = 0
    End Select
End Sub

Private Sub txtHaben_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2:
            Me.txtKonto.SelLength = 0
    Case vbKeyDown:
            Me.cmbBuTyp.SetFocus
    Case vbKeyUp:
            Me.txtKonto.SetFocus
    Case vbKeyReturn:
            GlBuF = 7 'Buchungsdialog
            S_KtSu "BuSe", KntRa
    End Select
End Sub
Private Sub txtKonto_GotFocus()
    Me.txtKonto.SelStart = 0
    Me.txtKonto.SelLength = Len(Me.txtKonto.Text)
End Sub
Private Sub txtKonto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn: KeyAscii = 0
    Case vbKeyTab: KeyAscii = 0
    End Select
End Sub
Private Sub txtKonto_KeyUp(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
    Case vbKeyF2:
            Me.txtKonto.SelLength = 0
    Case vbKeyDown:
            Me.cmbBuTex.SetFocus
    Case vbKeyUp:
            Me.cmbWarun.SetFocus
    Case vbKeyReturn:
            GlBuF = 3 'Buchungsdialog
            S_KtSu "BuSe", KntRa
    End Select
End Sub
