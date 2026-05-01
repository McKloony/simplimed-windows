VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Object = "{5B44EC52-B95B-45CF-98FF-A49DFEED5A92}#16.3#0"; "Codejock.PropertyGrid.v16.3.1.ocx"
Begin VB.Form frmAdress 
   Caption         =   "Stammdaten"
   ClientHeight    =   10125
   ClientLeft      =   165
   ClientTop       =   -2715
   ClientWidth     =   13350
   ControlBox      =   0   'False
   Icon            =   "frmAdress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   13350
   Begin XtremeReportControl.ReportControl repCont5 
      Height          =   795
      Left            =   2400
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   8900
      Visible         =   0   'False
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1411
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont4 
      Height          =   795
      Left            =   240
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   8900
      Visible         =   0   'False
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1411
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont2 
      Height          =   795
      Left            =   1320
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   8000
      Visible         =   0   'False
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1411
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   795
      Left            =   240
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   8000
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1411
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont3 
      Height          =   795
      Left            =   1320
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   8900
      Visible         =   0   'False
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1411
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   1600
      Left            =   9300
      TabIndex        =   171
      Top             =   8000
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   2822
      _StockProps     =   79
      Caption         =   "Bemerkung"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtS3F02 
         Height          =   1000
         Left            =   100
         TabIndex        =   174
         TabStop         =   0   'False
         Tag             =   "0Bemerkung"
         Top             =   400
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   1764
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtS3F03 
         Height          =   1000
         Left            =   3000
         TabIndex        =   175
         TabStop         =   0   'False
         Tag             =   "0Hinweis"
         Top             =   400
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   1764
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.Label lblS2L34 
         Height          =   225
         Left            =   30
         TabIndex        =   173
         Top             =   100
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Bemerkung :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblS2L35 
         Height          =   225
         Left            =   3000
         TabIndex        =   172
         Top             =   100
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Hinweise :"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   7200
      Left            =   13500
      TabIndex        =   85
      Top             =   600
      Visible         =   0   'False
      Width           =   4200
      _Version        =   1048579
      _ExtentX        =   7408
      _ExtentY        =   12700
      _StockProps     =   79
      Caption         =   "Zugehörige"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtS4F19 
         Height          =   350
         Left            =   1160
         TabIndex        =   97
         Top             =   4030
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnPost3 
         Height          =   350
         Left            =   1890
         TabIndex        =   93
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet die Postleitzahlensuche"
         Top             =   2620
         Width           =   315
         _Version        =   1048579
         _ExtentX        =   556
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox txtS4F02 
         Height          =   315
         Left            =   1155
         TabIndex        =   87
         Top             =   760
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F09 
         Height          =   350
         Left            =   2240
         TabIndex        =   94
         Top             =   2620
         Width           =   1720
         _Version        =   1048579
         _ExtentX        =   3034
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F08 
         Height          =   350
         Left            =   1160
         TabIndex        =   92
         Top             =   2620
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F06 
         Height          =   350
         Left            =   1160
         TabIndex        =   91
         Top             =   2160
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F05 
         Height          =   350
         Left            =   1160
         TabIndex        =   90
         Top             =   1690
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F04 
         Height          =   350
         Left            =   1160
         TabIndex        =   89
         Top             =   1230
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F03 
         Height          =   350
         Left            =   3010
         TabIndex        =   88
         Top             =   760
         Width           =   940
         _Version        =   1048579
         _ExtentX        =   1658
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F01 
         Height          =   350
         Left            =   1160
         TabIndex        =   86
         Top             =   300
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F16 
         Height          =   350
         Left            =   1540
         TabIndex        =   103
         Top             =   5410
         Width           =   2410
         _Version        =   1048579
         _ExtentX        =   4251
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F15 
         Height          =   350
         Left            =   1540
         TabIndex        =   101
         Top             =   4950
         Width           =   2410
         _Version        =   1048579
         _ExtentX        =   4251
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnTele9 
         Height          =   350
         Left            =   1160
         TabIndex        =   102
         TabStop         =   0   'False
         ToolTipText     =   "Wählt die nebenstehende Rufnummer"
         Top             =   5410
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTele8 
         Height          =   350
         Left            =   1160
         TabIndex        =   100
         TabStop         =   0   'False
         ToolTipText     =   "Wählt die nebenstehende Rufnummer"
         Top             =   4950
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbS4F11 
         Height          =   315
         Left            =   1160
         TabIndex        =   96
         Top             =   3550
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbS4F12 
         Height          =   315
         Left            =   1160
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F17 
         Height          =   1200
         Left            =   1160
         TabIndex        =   104
         Top             =   5860
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   2117
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F18 
         Height          =   350
         Left            =   1160
         TabIndex        =   98
         Top             =   4500
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
      Begin XtremeSuiteControls.CheckBox chkOpti3 
         Height          =   220
         Left            =   2500
         TabIndex        =   99
         TabStop         =   0   'False
         Tag             =   "0Mailing"
         Top             =   4560
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Emailverteiler"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Suchbegriff :"
         Height          =   240
         Left            =   0
         TabIndex        =   180
         Top             =   4090
         Width           =   1000
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren :"
         Height          =   240
         Left            =   100
         TabIndex        =   179
         Top             =   4530
         Width           =   1000
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   240
         Left            =   100
         TabIndex        =   178
         Top             =   5460
         Width           =   1000
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Telefon :"
         Height          =   240
         Left            =   100
         TabIndex        =   177
         Top             =   5000
         Width           =   1000
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar :"
         Height          =   240
         Left            =   100
         TabIndex        =   176
         Top             =   5880
         Width           =   1000
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Briefanrede :"
         Height          =   240
         Left            =   100
         TabIndex        =   169
         Top             =   3600
         Width           =   1000
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Land :"
         Height          =   240
         Left            =   100
         TabIndex        =   168
         Top             =   3140
         Width           =   1000
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PLZ / Ort :"
         Height          =   240
         Left            =   100
         TabIndex        =   167
         Top             =   2670
         Width           =   1000
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Straße :"
         Height          =   240
         Left            =   100
         TabIndex        =   166
         Top             =   2200
         Width           =   1000
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Nachname :"
         Height          =   240
         Left            =   100
         TabIndex        =   165
         Top             =   1750
         Width           =   1000
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname :"
         Height          =   240
         Left            =   100
         TabIndex        =   164
         Top             =   1290
         Width           =   1000
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Titel :"
         Height          =   240
         Left            =   2580
         TabIndex        =   163
         Top             =   800
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anrede :"
         Height          =   240
         Left            =   100
         TabIndex        =   162
         Top             =   800
         Width           =   1000
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Firma/Instit.:"
         Height          =   240
         Left            =   100
         TabIndex        =   161
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zugehöriger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   160
         Top             =   40
         Width           =   2000
      End
   End
   Begin XtremePropertyGrid.PropertyGrid prpGrid3 
      Height          =   1800
      Left            =   7080
      TabIndex        =   84
      Top             =   8000
      Visible         =   0   'False
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   3175
      _StockProps     =   68
      ToolBarVisible  =   0   'False
      HelpVisible     =   -1  'True
      PropertySort    =   0
   End
   Begin XtremePropertyGrid.PropertyGrid prpGrid2 
      Height          =   1800
      Left            =   5475
      TabIndex        =   83
      Top             =   8000
      Visible         =   0   'False
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   3175
      _StockProps     =   68
      ToolBarVisible  =   0   'False
      HelpVisible     =   -1  'True
      PropertySort    =   0
   End
   Begin XtremePropertyGrid.PropertyGrid prpGrid1 
      Height          =   1800
      Left            =   3885
      TabIndex        =   82
      Top             =   8000
      Visible         =   0   'False
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   3175
      _StockProps     =   68
      ToolBarVisible  =   0   'False
      HelpVisible     =   -1  'True
      PropertySort    =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtTreKey 
      Height          =   200
      Left            =   100
      TabIndex        =   4
      Tag             =   "0TreKey"
      Top             =   15000
      Width           =   100
      _Version        =   1048579
      _ExtentX        =   176
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      MaxLength       =   250
      Appearance      =   1
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   7200
      Left            =   9000
      TabIndex        =   3
      Top             =   600
      Width           =   4200
      _Version        =   1048579
      _ExtentX        =   7408
      _ExtentY        =   12700
      _StockProps     =   79
      Caption         =   "Sonstiges"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtS2F35 
         Height          =   350
         Left            =   1160
         TabIndex        =   77
         Tag             =   "0BIC"
         Top             =   6340
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnGutha 
         Height          =   350
         Left            =   2280
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "Trägt eine Akonto- bzw. Barzahlung ein"
         Top             =   6810
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTele7 
         Height          =   350
         Left            =   2020
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "Erstellt eine neue Email an die nebenstehende Emailadresse"
         Top             =   2620
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F33 
         Height          =   350
         Left            =   1160
         TabIndex        =   76
         Tag             =   "0IBAN"
         Top             =   5880
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnTele6 
         Height          =   350
         Left            =   2020
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Zeigt die nebenstehende Internetseite"
         Top             =   3090
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTele5 
         Height          =   350
         Left            =   2020
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "Erstellt eine neue Email an die nebenstehende Emailadresse"
         Top             =   2160
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F05 
         Height          =   350
         Left            =   1160
         TabIndex        =   74
         Tag             =   "0Konto"
         Top             =   4950
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F03 
         Height          =   350
         Left            =   1160
         TabIndex        =   75
         Tag             =   "0Bank"
         Top             =   5410
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F27 
         Height          =   350
         Left            =   2400
         TabIndex        =   70
         Tag             =   "0Internet"
         Top             =   3090
         Width           =   1550
         _Version        =   1048579
         _ExtentX        =   2734
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F19 
         Height          =   350
         Left            =   2400
         TabIndex        =   64
         Tag             =   "0Telefon5"
         Top             =   2160
         Width           =   1550
         _Version        =   1048579
         _ExtentX        =   2734
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F18 
         Height          =   350
         Left            =   2400
         TabIndex        =   61
         Tag             =   "0Telefon4"
         Top             =   1690
         Width           =   1550
         _Version        =   1048579
         _ExtentX        =   2734
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F17 
         Height          =   350
         Left            =   2400
         TabIndex        =   58
         Tag             =   "0Telefon3"
         Top             =   1230
         Width           =   1550
         _Version        =   1048579
         _ExtentX        =   2734
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F16 
         Height          =   350
         Left            =   2400
         TabIndex        =   55
         Tag             =   "0Telefon2"
         Top             =   760
         Width           =   1550
         _Version        =   1048579
         _ExtentX        =   2734
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F15 
         Height          =   350
         Left            =   2400
         TabIndex        =   52
         Tag             =   "0Telefon1"
         Top             =   300
         Width           =   1550
         _Version        =   1048579
         _ExtentX        =   2734
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F10 
         Height          =   315
         Left            =   1155
         TabIndex        =   72
         Tag             =   "0IDP"
         Top             =   4030
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox11"
         DropDownItemCount=   15
      End
      Begin XtremeSuiteControls.PushButton btnTele4 
         Height          =   350
         Left            =   2020
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Sendet eine SMS an die nebenstehende Rufnummer"
         Top             =   1690
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTele3 
         Height          =   350
         Left            =   2020
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Wählt die nebenstehende Rufnummer"
         Top             =   1230
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTele2 
         Height          =   350
         Left            =   2020
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Wählt die nebenstehende Rufnummer"
         Top             =   760
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTele1 
         Height          =   350
         Left            =   2020
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Wählt die nebenstehende Rufnummer"
         Top             =   300
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F07 
         Height          =   315
         Left            =   1155
         TabIndex        =   71
         Tag             =   "0Währung"
         Top             =   3550
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox12"
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F34 
         Height          =   350
         Left            =   2400
         TabIndex        =   67
         Tag             =   "0Telefon6"
         Top             =   2620
         Width           =   1550
         _Version        =   1048579
         _ExtentX        =   2734
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F33 
         Height          =   350
         Left            =   1160
         TabIndex        =   78
         Tag             =   "0Guthaben"
         Top             =   6810
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F27 
         Height          =   315
         Left            =   375
         TabIndex        =   68
         Tag             =   "0TelTyp7"
         Top             =   3090
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   11
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F26 
         Height          =   315
         Left            =   375
         TabIndex        =   65
         Tag             =   "0TelTyp6"
         Top             =   2620
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   11
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F25 
         Height          =   315
         Left            =   375
         TabIndex        =   62
         Tag             =   "0TelTyp5"
         Top             =   2160
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   11
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F21 
         Height          =   315
         Left            =   375
         TabIndex        =   50
         Tag             =   "0TelTyp1"
         Top             =   300
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   11
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F22 
         Height          =   315
         Left            =   375
         TabIndex        =   53
         Tag             =   "0TelTyp2"
         Top             =   760
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   11
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F23 
         Height          =   315
         Left            =   375
         TabIndex        =   56
         Tag             =   "0TelTyp3"
         Top             =   1230
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   11
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F24 
         Height          =   315
         Left            =   375
         TabIndex        =   59
         Tag             =   "0TelTyp4"
         Top             =   1690
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   11
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F29 
         Height          =   315
         Left            =   1155
         TabIndex        =   73
         Tag             =   "0Behindert"
         Top             =   4500
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F36 
         Height          =   315
         Left            =   2740
         TabIndex        =   80
         Tag             =   "0GebTyp"
         Top             =   6810
         Width           =   1220
         _Version        =   1048579
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblS1L35 
         Height          =   240
         Left            =   100
         TabIndex        =   151
         Top             =   6870
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Guthaben :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblS1L34 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "BIC :"
         Height          =   240
         Left            =   100
         TabIndex        =   150
         Top             =   6400
         Width           =   1000
      End
      Begin VB.Label lblS1L33 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "IBAN :"
         Height          =   240
         Left            =   100
         TabIndex        =   149
         Top             =   5920
         Width           =   1000
      End
      Begin VB.Label lblS2L33 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kreditinstitut :"
         Height          =   240
         Left            =   100
         TabIndex        =   148
         Top             =   5460
         Width           =   1000
      End
      Begin VB.Label lblS2L06 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kontoinhaber :"
         Height          =   240
         Left            =   100
         TabIndex        =   145
         Top             =   5000
         Width           =   1000
      End
      Begin VB.Label lblS2L27 
         BackStyle       =   0  'Transparent
         Caption         =   "Sonstiges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   400
         TabIndex        =   112
         Top             =   40
         Width           =   2000
      End
      Begin VB.Label lblS2L23 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Behinderung :"
         Height          =   240
         Left            =   100
         TabIndex        =   111
         Top             =   4530
         Width           =   1000
      End
      Begin VB.Label lblS2L19 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   240
         Left            =   100
         TabIndex        =   110
         Top             =   4090
         Width           =   1000
      End
      Begin VB.Label lblS2L04 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Währung :"
         Height          =   240
         Left            =   100
         TabIndex        =   109
         Top             =   3600
         Width           =   1000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   7200
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   4200
      _Version        =   1048579
      _ExtentX        =   7408
      _ExtentY        =   12700
      _StockProps     =   79
      Caption         =   "Rechnungsempfänger"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtS2F22 
         Height          =   350
         Left            =   1160
         TabIndex        =   38
         Tag             =   "0R_Land"
         Top             =   3090
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnPost2 
         Height          =   350
         Left            =   1890
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet die Postleitzahlensuche"
         Top             =   2620
         Width           =   315
         _Version        =   1048579
         _ExtentX        =   556
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   340
         Left            =   2050
         TabIndex        =   48
         Top             =   6810
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Min             =   1
         Value           =   1
         Max             =   10
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtS1F14"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.ComboBox txtS2F12 
         Height          =   315
         Left            =   1155
         TabIndex        =   30
         Tag             =   "0R_Anrede"
         Top             =   760
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkOpti2 
         Height          =   220
         Left            =   2500
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "0Edit"
         Top             =   4560
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Synchronisation"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F20 
         Height          =   300
         Left            =   1160
         TabIndex        =   81
         Tag             =   "0R_Briefanrede"
         Top             =   19000
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   6
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F19 
         Height          =   350
         Left            =   2240
         TabIndex        =   37
         Tag             =   "0R_Ort"
         Top             =   2620
         Width           =   1720
         _Version        =   1048579
         _ExtentX        =   3034
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F18 
         Height          =   350
         Left            =   1160
         TabIndex        =   35
         Tag             =   "0R_PLZ"
         Top             =   2620
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F16 
         Height          =   350
         Left            =   1160
         TabIndex        =   34
         Tag             =   "0R_Straße"
         Top             =   2160
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F15 
         Height          =   350
         Left            =   1160
         TabIndex        =   33
         Tag             =   "0R_Name"
         Top             =   1690
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F14 
         Height          =   350
         Left            =   1160
         TabIndex        =   32
         Tag             =   "0R_Vorname"
         Top             =   1230
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F13 
         Height          =   350
         Left            =   3020
         TabIndex        =   31
         Tag             =   "0R_Titel"
         Top             =   760
         Width           =   940
         _Version        =   1048579
         _ExtentX        =   1658
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F11 
         Height          =   350
         Left            =   1160
         TabIndex        =   29
         Tag             =   "0R_Firma1"
         Top             =   300
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F09 
         Height          =   315
         Left            =   1160
         TabIndex        =   46
         Tag             =   "0IDZ"
         Top             =   6340
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F06 
         Height          =   315
         Left            =   1160
         TabIndex        =   45
         Tag             =   "0ID3"
         Top             =   5880
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox txtS2F08 
         Height          =   315
         Left            =   1160
         TabIndex        =   43
         Tag             =   "0Behandelt"
         Top             =   4950
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         DropDownItemCount=   12
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F14 
         Height          =   350
         Left            =   1160
         TabIndex        =   47
         Tag             =   "0Kopien"
         Top             =   6810
         Width           =   880
         _Version        =   1048579
         _ExtentX        =   1552
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F07 
         Height          =   315
         Left            =   1160
         TabIndex        =   44
         Tag             =   "0IDV"
         Top             =   5410
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F25 
         Height          =   350
         Left            =   1160
         TabIndex        =   41
         Tag             =   "0R_Geboren"
         Top             =   4500
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
      Begin XtremeSuiteControls.ComboBox cmbS1F10 
         Height          =   315
         Left            =   1155
         TabIndex        =   39
         Tag             =   "0Briefanrede"
         Top             =   3550
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox txtS1F22 
         Height          =   315
         Left            =   1155
         TabIndex        =   40
         TabStop         =   0   'False
         Tag             =   "0Postfach"
         Top             =   4030
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F28 
         Height          =   315
         Left            =   2500
         TabIndex        =   49
         Tag             =   "0Versand"
         Top             =   6810
         Width           =   1440
         _Version        =   1048579
         _ExtentX        =   2540
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblS2L20 
         Height          =   240
         Left            =   100
         TabIndex        =   182
         Top             =   4090
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Bundesland :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblS1L19 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Zahlungsw.:"
         Height          =   240
         Left            =   120
         TabIndex        =   147
         Top             =   6400
         Width           =   1000
      End
      Begin VB.Label lblS1L26 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Ausdrucke :"
         Height          =   240
         Left            =   100
         TabIndex        =   144
         Top             =   6870
         Width           =   1000
      End
      Begin VB.Label lblS2L26 
         BackStyle       =   0  'Transparent
         Caption         =   "Rechnungsempfänger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   126
         Top             =   40
         Width           =   2000
      End
      Begin VB.Label lblS2L05 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Verwandt :"
         Height          =   240
         Left            =   100
         TabIndex        =   125
         Top             =   5000
         Width           =   1000
      End
      Begin VB.Label lblS1L18 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Katalog :"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   100
         TabIndex        =   124
         Top             =   5920
         Width           =   1000
      End
      Begin VB.Label lblS2L07 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PKV-Tarif :"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   100
         TabIndex        =   123
         Top             =   5460
         Width           =   1000
      End
      Begin VB.Label lblS2L22 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren :"
         Height          =   240
         Left            =   100
         TabIndex        =   122
         Top             =   4530
         Width           =   1000
      End
      Begin VB.Label lblS2L18 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Land :"
         Height          =   240
         Left            =   100
         TabIndex        =   121
         Top             =   3140
         Width           =   1000
      End
      Begin VB.Label lblS2L10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Organisation :"
         Height          =   240
         Left            =   100
         TabIndex        =   120
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label lblS2L11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anrede :"
         Height          =   240
         Left            =   100
         TabIndex        =   119
         Top             =   800
         Width           =   1000
      End
      Begin VB.Label lblS2L12 
         BackStyle       =   0  'Transparent
         Caption         =   "Titel :"
         Height          =   240
         Left            =   2580
         TabIndex        =   118
         Top             =   800
         Width           =   420
      End
      Begin VB.Label lblS2L13 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname :"
         Height          =   240
         Left            =   100
         TabIndex        =   117
         Top             =   1290
         Width           =   1000
      End
      Begin VB.Label lblS2L14 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Nachname :"
         Height          =   240
         Left            =   100
         TabIndex        =   116
         Top             =   1750
         Width           =   1000
      End
      Begin VB.Label lblS2L15 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Straße :"
         Height          =   240
         Left            =   100
         TabIndex        =   115
         Top             =   2200
         Width           =   1000
      End
      Begin VB.Label lblS2L16 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PLZ / Ort :"
         Height          =   240
         Left            =   100
         TabIndex        =   114
         Top             =   2670
         Width           =   1000
      End
      Begin VB.Label lblS2L17 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Briefanrede :"
         Height          =   240
         Left            =   100
         TabIndex        =   113
         Top             =   3600
         Width           =   1000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   7200
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4365
      _Version        =   1048579
      _ExtentX        =   7699
      _ExtentY        =   12700
      _StockProps     =   79
      Caption         =   "Patientendaten"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnPost1 
         Height          =   350
         Left            =   1890
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet die Postleitzahlensuche"
         Top             =   2620
         Width           =   315
         _Version        =   1048579
         _ExtentX        =   556
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox txtS1F02 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Tag             =   "0Anrede"
         Top             =   760
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F32 
         Height          =   350
         Left            =   3040
         TabIndex        =   26
         Tag             =   "0Gewicht"
         Top             =   6340
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
      Begin XtremeSuiteControls.FlatEdit txtS1F31 
         Height          =   350
         Left            =   1160
         TabIndex        =   25
         Tag             =   "0Größe"
         Top             =   6340
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
      Begin XtremeSuiteControls.FlatEdit txtS2F24 
         Height          =   350
         Left            =   1160
         TabIndex        =   18
         Tag             =   "0Beruf"
         Top             =   4030
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.CheckBox chkOpti1 
         Height          =   220
         Left            =   2500
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "0Mailing"
         Top             =   4560
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Serienbriefadresse"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F11 
         Height          =   350
         Left            =   1160
         TabIndex        =   17
         Tag             =   "0IDKurz"
         Top             =   3550
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F09 
         Height          =   350
         Left            =   2240
         TabIndex        =   15
         Tag             =   "0Ort"
         Top             =   2620
         Width           =   1720
         _Version        =   1048579
         _ExtentX        =   3034
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F08 
         Height          =   350
         Left            =   1160
         TabIndex        =   13
         Tag             =   "0PLZ"
         Top             =   2620
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F06 
         Height          =   350
         Left            =   1160
         TabIndex        =   12
         Tag             =   "0Straße"
         Top             =   2160
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F05 
         Height          =   350
         Left            =   1160
         TabIndex        =   11
         Tag             =   "0Name"
         Top             =   1690
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F04 
         Height          =   350
         Left            =   1160
         TabIndex        =   10
         Tag             =   "0Vorname"
         Top             =   1230
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F03 
         Height          =   350
         Left            =   3010
         TabIndex        =   9
         Tag             =   "0Titel"
         Top             =   760
         Width           =   940
         _Version        =   1048579
         _ExtentX        =   1658
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F01 
         Height          =   350
         Left            =   1160
         TabIndex        =   7
         Tag             =   "0Firma1"
         Top             =   300
         Width           =   2800
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox txtS2F26 
         Height          =   315
         Left            =   2740
         TabIndex        =   22
         Tag             =   "0Familienstand"
         Top             =   4950
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F37 
         Height          =   350
         Left            =   1160
         TabIndex        =   27
         Tag             =   "0Blutgruppe"
         Top             =   6810
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
      Begin XtremeSuiteControls.FlatEdit txtS1F30 
         Height          =   350
         Left            =   3040
         TabIndex        =   28
         Tag             =   "0Mandant"
         Top             =   6810
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
      Begin XtremeSuiteControls.FlatEdit txtS1F13 
         Height          =   350
         Left            =   1160
         TabIndex        =   19
         Tag             =   "0Geboren"
         Top             =   4500
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
      Begin XtremeSuiteControls.ComboBox txtS1F12 
         Height          =   315
         Left            =   1155
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "0Land"
         Top             =   3090
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbS1F08 
         Height          =   315
         Left            =   1160
         TabIndex        =   21
         Tag             =   "0GeschlTyp"
         Top             =   4950
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F30 
         Height          =   315
         Left            =   1160
         TabIndex        =   23
         Tag             =   "0Abteilung"
         Top             =   5410
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F31 
         Height          =   315
         Left            =   1160
         TabIndex        =   24
         Tag             =   "0BGNummer"
         Top             =   5880
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.Label lblS1L13 
         Height          =   240
         Left            =   100
         TabIndex        =   181
         Top             =   5460
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Vertragsart :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblS1L31 
         Height          =   240
         Left            =   2280
         TabIndex        =   146
         Top             =   6870
         Width           =   705
         _Version        =   1048579
         _ExtentX        =   1244
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "PIN :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblS1L28 
         Height          =   240
         Left            =   100
         TabIndex        =   143
         Top             =   6870
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Blutgruppe :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblS2L25 
         BackStyle       =   0  'Transparent
         Caption         =   "Patientendaten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   142
         Top             =   40
         Width           =   2000
      End
      Begin VB.Label lblS1L27 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Gewicht :"
         Height          =   240
         Left            =   2280
         TabIndex        =   141
         Top             =   6400
         Width           =   700
      End
      Begin VB.Label lblS1L12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Körpergröße :"
         Height          =   240
         Left            =   100
         TabIndex        =   140
         Top             =   6400
         Width           =   1000
      End
      Begin VB.Label lblS1L10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Hausarzt :"
         Height          =   240
         Left            =   100
         TabIndex        =   139
         Top             =   5920
         Width           =   1000
      End
      Begin VB.Label lblS2L08 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   240
         Left            =   100
         TabIndex        =   138
         Top             =   5000
         Width           =   1000
      End
      Begin VB.Label lblS1L01 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Organisation :"
         Height          =   240
         Left            =   100
         TabIndex        =   137
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label lblS1L02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anrede :"
         Height          =   240
         Left            =   100
         TabIndex        =   136
         Top             =   800
         Width           =   1000
      End
      Begin VB.Label lblS1L03 
         BackStyle       =   0  'Transparent
         Caption         =   "Titel :"
         Height          =   240
         Left            =   2580
         TabIndex        =   135
         Top             =   800
         Width           =   420
      End
      Begin VB.Label lblS1L04 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname :"
         Height          =   240
         Left            =   100
         TabIndex        =   134
         Top             =   1290
         Width           =   1000
      End
      Begin VB.Label lblS1L05 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Nachname :"
         Height          =   240
         Left            =   100
         TabIndex        =   133
         Top             =   1750
         Width           =   1000
      End
      Begin VB.Label lblS1L06 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Straße :"
         Height          =   240
         Left            =   100
         TabIndex        =   132
         Top             =   2200
         Width           =   1000
      End
      Begin VB.Label lblS1L07 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PLZ / Ort :"
         Height          =   240
         Left            =   100
         TabIndex        =   131
         Top             =   2670
         Width           =   1000
      End
      Begin VB.Label lblS1L08 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Land :"
         Height          =   240
         Left            =   100
         TabIndex        =   130
         Top             =   3140
         Width           =   1000
      End
      Begin VB.Label lblS1L09 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Suchbegriff :"
         Height          =   240
         Left            =   100
         TabIndex        =   129
         Top             =   3600
         Width           =   1000
      End
      Begin VB.Label lblS1L11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Beruf :"
         Height          =   240
         Left            =   100
         TabIndex        =   128
         Top             =   4090
         Width           =   1000
      End
      Begin VB.Label lblS1L25 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren :"
         Height          =   240
         Left            =   100
         TabIndex        =   127
         Top             =   4530
         Width           =   1000
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtS1F20 
      Height          =   200
      Left            =   0
      TabIndex        =   152
      TabStop         =   0   'False
      Tag             =   "0DuSie"
      Top             =   19000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtGesch 
      Height          =   200
      Left            =   0
      TabIndex        =   153
      TabStop         =   0   'False
      Tag             =   "0Geschlecht"
      Top             =   15000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtFirma 
      Height          =   200
      Left            =   300
      TabIndex        =   154
      TabStop         =   0   'False
      Tag             =   "0Firma2"
      Top             =   15000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS3F01 
      Height          =   200
      Left            =   600
      TabIndex        =   155
      TabStop         =   0   'False
      Tag             =   "0Anschrift"
      Top             =   15000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAnsch 
      Height          =   200
      Left            =   1400
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   15000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F27 
      Height          =   200
      Left            =   0
      TabIndex        =   157
      TabStop         =   0   'False
      Tag             =   "0Geschlecht"
      Top             =   15000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtVersi 
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Tag             =   "0Versicherung"
      Top             =   15000
      Visible         =   0   'False
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtVerNr 
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Tag             =   "0Kartennummer"
      Top             =   15000
      Visible         =   0   'False
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtKarGu 
      Height          =   195
      Left            =   960
      TabIndex        =   158
      TabStop         =   0   'False
      Tag             =   "0Kartengultig"
      Top             =   15000
      Visible         =   0   'False
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAdrGr 
      Height          =   200
      Left            =   0
      TabIndex        =   159
      TabStop         =   0   'False
      Tag             =   "0AdrGruppe"
      Top             =   15000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   15000
      Width           =   100
      _Version        =   1048579
      _ExtentX        =   176
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   720
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars comBar01 
      Left            =   120
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   4
   End
End
Attribute VB_Name = "frmAdress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private FeKop As XtremeSuiteControls.FlatEdit
Private FePIN As XtremeSuiteControls.FlatEdit
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions
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
Private TabCo As XtremeSuiteControls.TabControl
Private TabIt As XtremeSuiteControls.TabControlItem
Private PrGr1 As XtremePropertyGrid.PropertyGrid
Private PrGr2 As XtremePropertyGrid.PropertyGrid
Private PrGr3 As XtremePropertyGrid.PropertyGrid
Private PrItm As XtremePropertyGrid.PropertyGridItem
Private RpCo1 As XtremeReportControl.ReportControl
Private RpCo2 As XtremeReportControl.ReportControl
Private RpCo3 As XtremeReportControl.ReportControl
Private RpCo4 As XtremeReportControl.ReportControl
Private RpCo5 As XtremeReportControl.ReportControl
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Const CB_SHOWDROPDOWN = &H14F
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const GWL_WNDPROC = (-4)
Private Const KEYEVENTF_KEYUP = &H2

Private RetWe As Long
Private TagWe As String
Private TabId As Integer
Private Const MinBr = 6500
Private Const MinHo = 6600

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private clFen As clsFenster

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub btnGutha_Click()
    frmAnzahl.Show vbModal
End Sub
Private Sub btnPost1_Click()
    If GlAdL = False Then
        Adr_Pos False, False, True
    End If
End Sub

Private Sub btnPost2_Click()
    If GlAdL = False Then
        Adr_Pos True, False, True
    End If
End Sub
Private Sub btnPost3_Click()
    Adr_Poz True
End Sub
Private Sub btnTele1_Click()
    If Me.txtS1F15.Text <> vbNullString Then
        AdTel Me.txtS1F15.Text, GlPrg
    End If
End Sub

Private Sub btnTele2_Click()
    If Me.txtS1F16.Text <> vbNullString Then
        AdTel Me.txtS1F16.Text, GlPrg
    End If
End Sub
Private Sub btnTele3_Click()
    If Me.txtS1F17.Text <> vbNullString Then
        AdTel Me.txtS1F17.Text, GlPrg
    End If
End Sub
Private Sub btnTele4_Click()
On Error Resume Next

Dim TmStr As String

If Me.txtS1F18.Text <> vbNullString Then
    Me.txtS1F18.Text = SRufn(Me.txtS1F18.Text) 'Formatiert die Rufnummer
    DoEvents
    TmStr = SMSTe(Me.txtS1F18.Text) 'Testen des SMS Rufnummernformates
    If TmStr <> vbNullString Then
        SPopu "Richtiges Rufnummernformat", TmStr, IC48_Information
    Else
        SPopu "Falsches Rufnummernformat", "Die eingegebene Rufnummer hat das falsche Format!", IC48_Forbidden
    End If
End If

End Sub
Private Sub btnTele5_Click()
    If Me.txtS1F19.Text <> vbNullString Then
        SMaNe GlAId, Me.txtS1F19.Text, vbNullString, Me.cmbS1F10.Text
        Unload Me
    End If
End Sub

Private Sub btnTele6_Click()
    AEmail 3
End Sub

Private Sub btnTele7_Click()
    If Me.txtS2F34.Text <> vbNullString Then
        SMaNe GlAId, Me.txtS2F34.Text, vbNullString, Me.cmbS1F10.Text
        Unload Me
    End If
End Sub

Private Sub btnTele8_Click()
    If Me.txtS4F15.Text <> vbNullString Then
        AdTel Me.txtS4F15.Text, GlPrg
    End If
End Sub

Private Sub btnTele9_Click()
    If Me.txtS4F16.Text <> vbNullString Then
        SMaNe GlAId, Me.txtS4F16.Text, vbNullString, Me.cmbS4F11.Text
        Unload Me
    End If
End Sub
Private Sub chkOpti3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GlAdL = False
End Sub

Private Sub cmbS1F07_Click()

TagWe = Mid$(Me.cmbS1F07.Tag, 2, Len(Me.cmbS1F07.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F07.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
    If GlFri = 2 Then 'Heilpraktiker
        Adr_Trf
    Else
        SPopu "Nur für Heilpraktikerabrechnung", "Diese Auswahl ist nur für die Heilpraktikerabrechnung vorgesehen.", IC48_Information
    End If
End If

End Sub
Private Sub cmbS1F07_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS1F07_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.cmbS1F07.SelLength = 0
    End If
End Sub


Private Sub cmbS1F10_GotFocus()
On Error Resume Next

TagWe = Mid$(Me.txtS2F20.Tag, 2, Len(Me.txtS2F20.Tag) - 1)

If GlAdL = False Then
    AKopi 'Kopieren der Adressdaten in die Rechungsanschrift
    DoEvents
    AdBrf
    DoEvents
    Me.txtS2F20.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
    Me.txtS2F20.Text = Me.cmbS1F10.Text
End If
    
End Sub

Private Sub cmbS1F21_Change()

TagWe = Mid$(Me.cmbS1F21.Tag, 2, Len(Me.cmbS1F21.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F21.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS1F22_Change()

TagWe = Mid$(Me.cmbS1F22.Tag, 2, Len(Me.cmbS1F22.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F22.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS1F23_Change()

TagWe = Mid$(Me.cmbS1F23.Tag, 2, Len(Me.cmbS1F23.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F23.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS1F24_Change()

TagWe = Mid$(Me.cmbS1F24.Tag, 2, Len(Me.cmbS1F24.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F24.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS1F25_Change()

TagWe = Mid$(Me.cmbS1F25.Tag, 2, Len(Me.cmbS1F25.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F25.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F25_Click()

TagWe = Mid$(Me.cmbS1F25.Tag, 2, Len(Me.cmbS1F25.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F25.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F25_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS1F25_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F25.SelLength = 0
End Sub
Private Sub cmbS1F26_Change()

TagWe = Mid$(Me.cmbS1F26.Tag, 2, Len(Me.cmbS1F26.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F26.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F26_Click()

TagWe = Mid$(Me.cmbS1F26.Tag, 2, Len(Me.cmbS1F26.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F26.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F26_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS1F26_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F26.SelLength = 0
End Sub
Private Sub cmbS1F27_Change()

TagWe = Mid$(Me.cmbS1F27.Tag, 2, Len(Me.cmbS1F27.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F27.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F27_Click()

TagWe = Mid$(Me.cmbS1F27.Tag, 2, Len(Me.cmbS1F27.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F27.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS1F27_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS1F27_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.cmbS1F27.SelLength = 0
    End If
End Sub

Private Sub cmbS1F28_Click()

TagWe = Mid$(Me.cmbS1F28.Tag, 2, Len(Me.cmbS1F28.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F28.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F28_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS2F36_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS2F36_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS2F36.SelLength = 0
End Sub

Private Sub cmbS2F29_Click()

TagWe = Mid$(Me.cmbS2F29.Tag, 2, Len(Me.cmbS2F29.Tag) - 1)

If GlAdL = False Then
    Me.cmbS2F29.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS2F29_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS2F29_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS2F29.SelLength = 0
End Sub
Private Sub cmbS2F31_Click()

TagWe = Mid$(Me.cmbS2F31.Tag, 2, Len(Me.cmbS2F31.Tag) - 1)

If GlAdL = False Then
    Me.cmbS2F31.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS2F31_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS2F31_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS2F31.SelLength = 0
End Sub

Private Sub cmbS2F36_Click()

TagWe = Mid$(Me.cmbS2F36.Tag, 2, Len(Me.cmbS2F36.Tag) - 1)

If GlAdL = False Then
    Me.cmbS2F36.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS4F11_Click()
    GlAdZ = True
End Sub
Private Sub cmbS4F11_GotFocus()
    If GlAdL = False Then
        AdBri
    End If
End Sub
Private Sub cmbS4F11_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS4F12_Click()
    GlAdZ = True
End Sub
Private Sub cmbS4F12_GotFocus()
    GlAdL = False
End Sub
Private Sub cmbS4F12_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub Form_Load()

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 14000
    .ClientMaxWidth = 18000
    .ClientMinHeight = 10200
    .ClientMinWidth = 14000
    .TopMost = GlAVo
End With

TabId = RibTab_Adr_Haupt

Set FrmEx = Nothing

End Sub

Private Sub prpGrid1_ValueChanged(ByVal Item As XtremePropertyGrid.IPropertyGridItem)
On Error Resume Next

Dim IdxNr As Long
Dim FeNam As String

Set PrGr1 = Me.prpGrid1

IdxNr = Item.id
FeNam = Right$(Item.Tag, Len(Item.Tag) - 1)
TagWe = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

If GlAdL = False Then
    Select Case IdxNr
    Case 9901:
        If Item.Value <> vbNullString Then
            frmPasswort.PaStr = Item.Value
            frmPasswort.Show
        End If
        Item.Tag = "1" & TagWe
        GlAdS = True 'Speichern der Adresse erforderlich
    Case 9902:
        Set PrItm = PrGr1.FindItem(9901)
        If CBool(Item.Value) = True Then
            PrItm.PasswordMask = False
        Else
            PrItm.PasswordMask = True
        End If
    Case Else:
        Item.Tag = "1" & TagWe
        GlAdS = True 'Speichern der Adresse erforderlich
    End Select
End If

End Sub
Private Sub prpGrid2_ValueChanged(ByVal Item As XtremePropertyGrid.IPropertyGridItem)
    
TagWe = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

If GlAdL = False Then
    Item.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub prpGrid3_ValueChanged(ByVal Item As XtremePropertyGrid.IPropertyGridItem)

If GlAdL = False Then
    GlAdS = True 'Speichern der Adresse erforderlich
    Adr_EiSa Item
End If

End Sub
Private Sub repCont2_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If GlSta = False Then
    If Row.GroupRow = False Then
        If CBool(GlGeK(Row.Record(Buh_IDB).Value, 5)) = True Then
            Metrics.ForeColor = 16711680
        End If
    End If
End If

End Sub
Private Sub repCont2_ColumnOrderChangedEx(ByVal Column As XtremeReportControl.IReportColumn, ByVal Reason As XtremeReportControl.XTPReportColumnOrderChangedReason)
    If GlAdL = False Then ASpSv
End Sub
Private Sub repCont2_ColumnWidthChanged(ByVal Column As XtremeReportControl.IReportColumn, ByVal PrevWidth As Long, ByVal NewWidth As Long)
    If GlAdL = False Then ASpSv
End Sub
Private Sub repCont3_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

If GlSta = False Then
    If Row.Record.ItemCount > 30 Then
        If Row.Record(OPo_Selekt).Value = 0 Then
            Metrics.Font.Strikethrough = False
            If Row.Record(OPo_Mahnbar).Value = 0 Then
                Metrics.ForeColor = 12632256
            Else
                If Row.Record(OPo_IBAN).Value <> vbNullString Then
                    Metrics.ForeColor = 16711680
                ElseIf Row.Record(OPo_Konto).Value <> vbNullString Then
                    Metrics.ForeColor = 16711680
                Else
                    If Row.Record(OPo_Mahnfrist).Value < Date Then
                        Metrics.ForeColor = 210
                    Else
                        Metrics.ForeColor = 44800
                    End If
                End If
            End If
            If Row.Record(OPo_Beleg).Value <> vbNullString Then
                If Row.Record(OPo_Beleg).Value <> "0" Then
                    Metrics.Font.Bold = True
                End If
            End If
        Else
            Metrics.ForeColor = 8421504
            Metrics.Font.Strikethrough = True
        End If
    End If
End If

End Sub

Private Sub repCont3_ColumnOrderChangedEx(ByVal Column As XtremeReportControl.IReportColumn, ByVal Reason As XtremeReportControl.XTPReportColumnOrderChangedReason)
    If GlAdL = False Then ASpSv
End Sub
Private Sub repCont3_ColumnWidthChanged(ByVal Column As XtremeReportControl.IReportColumn, ByVal PrevWidth As Long, ByVal NewWidth As Long)
    If GlAdL = False Then ASpSv
End Sub
Private Sub repCont4_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

If GlSta = False Then
    If Row.GroupRow = False Then
        If IsNumeric(Row.Record(Ter_Farbe).Value) Then
            FrbZa = Row.Record(Ter_Farbe).Value
            If FrbZa > 1 And FrbZa <= 20 Then
                Metrics.BackColor = GlTmF(FrbZa, 1)
            End If
        End If
        If Row.Record(Ter_VonDat).Value >= Date Then
            Metrics.Font.Bold = True
        End If
    End If
End If

End Sub
Private Sub repCont5_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn: AZuLa
    Case vbKeyDown: AZuLa
    Case vbKeyUp: AZuLa
    End Select
End Sub
Private Sub repCont5_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    AZuLa
End Sub


Private Sub txtS1F22_Click()

TagWe = Mid$(Me.txtS1F22.Tag, 2, Len(Me.txtS1F22.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F22.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub txtS1F22_GotFocus()
    GlAdL = False
End Sub

Private Sub txtS1F22_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F22_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtS1F22.SelLength = 0
    End If
End Sub

Private Sub txtS1F30_LostFocus()
On Error Resume Next

Dim AdPIN As Long

Set FM = frmAdress
Set FePIN = FM.txtS1F30

If FePIN.Text = vbNullString Then
    AdPIN = Adr_Let()
    FePIN.Text = Format$(AdPIN, "000000")
    DoEvents
ElseIf IsNumeric(FePIN.Text) = False Then
    AdPIN = Adr_Let()
    FePIN.Text = Format$(AdPIN, "000000")
    DoEvents
ElseIf CLng(FePIN.Text) = 0 Then
    AdPIN = Adr_Let()
    FePIN.Text = Format$(AdPIN, "000000")
    DoEvents
End If

End Sub

Private Sub txtS2F11_GotFocus()
    If GlAdL = False Then
        Me.txtS2F11.SelStart = 0
        Me.txtS2F11.SelLength = Len(Me.txtS2F11.Text)
    End If
End Sub
Private Sub txtS2F33_LostFocus()
On Error Resume Next

Dim TmStr As String

If GlAdL = False Then
    If Me.txtS2F33.Text <> vbNullString Then
        TmStr = Me.txtS2F33.Text
        Me.txtS2F33.Text = SNaUm(TmStr)
        If Len(TmStr) <> 22 Then
            SPopu "IBAN ist falsch", "Die IBAN hat die falsche Länge", IC48_Forbidden
        End If
    End If
End If

End Sub
Private Sub txtS2F35_Change()

TagWe = Mid$(Me.txtS2F35.Tag, 2, Len(Me.txtS2F35.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F35.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub txtS2F35_GotFocus()
    Me.txtS2F35.SelStart = 0
    Me.txtS2F35.SelLength = Len(Me.txtS2F35.Text)
    GlAdL = False
End Sub


Private Sub txtS2F35_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F35_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F35.SelLength = 0
    Case vbKeyDown: Me.cmbS1F08.SetFocus
    Case vbKeyUp: Me.txtS2F33.SetFocus
    End Select
End Sub

Private Sub txtS2F35_LostFocus()
On Error Resume Next

Dim TmStr As String

If GlAdL = False Then
    If Me.txtS2F35.Text <> vbNullString Then
        TmStr = Me.txtS2F35.Text
        Me.txtS2F35.Text = SNaUm(TmStr)
        If Len(TmStr) <> 11 Then
            SPopu "BIC ist falsch", "Die BIC hat die falsche Länge", IC48_Forbidden
        End If
    End If
End If

End Sub
Private Sub txtS3F02_Change()

TagWe = Mid$(Me.txtS3F02.Tag, 2, Len(Me.txtS3F02.Tag) - 1)

If GlAdL = False Then
    Me.txtS3F02.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub cmbS1F08_Click()

TagWe = Mid$(Me.cmbS1F08.Tag, 2, Len(Me.cmbS1F08.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F08.Tag = "1" & TagWe
    
    If Me.cmbS1F08.Text <> vbNullString Then
        Me.txtGesch.Text = Me.cmbS1F08.Text
        TagWe = Mid$(Me.txtGesch.Tag, 2, Len(Me.txtGesch.Tag) - 1)
        Me.txtGesch.Tag = "1" & TagWe
    End If
    
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F08_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS1F08_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F08.SelLength = 0
End Sub
Private Sub FClip(ByVal AnTyp As Integer)
On Error GoTo WoErr
'Kopiert Adresse in die Zwischenablage

Dim TmStr As Variant

AErAd 'Erstellt die Anschrift im Anschriftenfeld
DoEvents

Select Case AnTyp
Case 1: TmStr = Me.txtAnsch.Text
Case 2: TmStr = Me.txtS1F11.Text
End Select

Clipboard.Clear
Clipboard.SetText TmStr

Exit Sub

WoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClip" & Err.Number
Resume Next

End Sub
Private Sub FClos()
On Error GoTo SaErr

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlIdi = False Then 'Idiotenmodus
    If GlRes = False Then 'Reset der Einstellungen
        clFen.FenSav
        If clFen.FeSta = 0 Then
            IniSetVal "AdrForm", "FenLin", clFen.FeLin
            IniSetVal "AdrForm", "FenObe", clFen.FeObn
            IniSetVal "AdrForm", "FenBre", clFen.FeBre
            IniSetVal "AdrForm", "FenHoh", clFen.FeHoh
        End If
    End If
End If

Set clFen = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FEnde()
On Error GoTo SaErr

If GlAdS = True Then 'Speichern der Adresse erforderlich
    AKopi 'Kopieren der Adressdaten in die Rechungsanschrift
    DoEvents
    AErAd 'Erstellt die Anschrift im Anschriftenfeld
    DoEvents
    FSave True
End If
DoEvents
Unload Me

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEnde " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case TabId
Case RibTab_Adr_Haupt:
    TeTit = IniGetOpt("Hilfe", 50021)
    TeMai = IniGetOpt("Hilfe", 50022)
    TeInh = IniGetOpt("Hilfe", 50023)
    TeFus = IniGetOpt("Hilfe", 50024)
Case RibTab_Adr_Dokum:
    TeTit = IniGetOpt("Hilfe", 50971)
    TeMai = IniGetOpt("Hilfe", 50972)
    TeInh = IniGetOpt("Hilfe", 50973)
    TeFus = IniGetOpt("Hilfe", 50974)
Case RibTab_Adr_Eigen:
    TeTit = IniGetOpt("Hilfe", 50981)
    TeMai = IniGetOpt("Hilfe", 50982)
    TeInh = IniGetOpt("Hilfe", 50983)
    TeFus = IniGetOpt("Hilfe", 50984)
Case RibTab_Adr_Booki:
    TeTit = IniGetOpt("Hilfe", 50991)
    TeMai = IniGetOpt("Hilfe", 50992)
    TeInh = IniGetOpt("Hilfe", 50993)
    TeFus = IniGetOpt("Hilfe", 50994)
Case RibTab_Adr_Membe:
    TeTit = IniGetOpt("Hilfe", 51001)
    TeMai = IniGetOpt("Hilfe", 51002)
    TeInh = IniGetOpt("Hilfe", 51003)
    TeFus = IniGetOpt("Hilfe", 51004)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub

Private Sub chkOpti1_Click()

TagWe = Mid$(Me.chkOpti1.Tag, 2, Len(Me.chkOpti1.Tag) - 1)

If GlAdL = False Then
    Me.chkOpti1.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub
Private Sub chkOpti1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GlAdL = False
End Sub
Private Sub chkOpti2_Click()

TagWe = Mid$(Me.chkOpti2.Tag, 2, Len(Me.chkOpti2.Tag) - 1)

If GlAdL = False Then
    Me.chkOpti2.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub chkOpti2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GlAdL = False
End Sub
Private Sub cmbS1F21_Click()

TagWe = Mid$(Me.cmbS1F21.Tag, 2, Len(Me.cmbS1F21.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F21.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F21_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS1F21_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F21.SelLength = 0
End Sub

Private Sub cmbS1F22_Click()

TagWe = Mid$(Me.cmbS1F22.Tag, 2, Len(Me.cmbS1F22.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F22.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F22_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS1F22_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F22.SelLength = 0
End Sub

Private Sub cmbS1F23_Click()

TagWe = Mid$(Me.cmbS1F23.Tag, 2, Len(Me.cmbS1F23.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F23.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F23_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS1F23_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F23.SelLength = 0
End Sub

Private Sub cmbS1F24_Click()

TagWe = Mid$(Me.cmbS1F24.Tag, 2, Len(Me.cmbS1F24.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F24.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
End If

End Sub

Private Sub cmbS1F24_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS1F24_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F24.SelLength = 0
End Sub

Private Sub cmbS1F06_Click()

TagWe = Mid$(Me.cmbS1F06.Tag, 2, Len(Me.cmbS1F06.Tag) - 1)

If GlAdL = False Then
    Me.cmbS1F06.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbS1F06_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS1F06_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS1F06.SelLength = 0
End Sub

Private Sub cmbS2F07_Click()

TagWe = Mid$(Me.cmbS2F07.Tag, 2, Len(Me.cmbS2F07.Tag) - 1)

If GlAdL = False Then
    Me.cmbS2F07.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbS2F07_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS2F07_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS2F07.SelLength = 0
End Sub

Private Sub cmbS2F09_Click()

TagWe = Mid$(Me.cmbS2F09.Tag, 2, Len(Me.cmbS2F09.Tag) - 1)

If GlAdL = False Then
    Me.cmbS2F09.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbS2F09_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS2F10_Click()

TagWe = Mid$(Me.cmbS2F10.Tag, 2, Len(Me.cmbS2F10.Tag) - 1)

Me.cmbS2F10.Tag = "1" & TagWe

GlAdS = True 'Speichern der Adresse erforderlich

End Sub

Private Sub cmbS2F10_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS2F10_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS2F10.SelLength = 0
End Sub

Private Sub comBar01_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlAdL = False Then
        If Control.id = XTP_ID_RIBBONCONTROLTAB Then
            FTabu
        Else
            FTool Control.id
        End If
    End If
End Sub
Private Sub comBar01_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlAdL = False Then
    RetWe = SendMessage(Me.hwnd, WM_SETREDRAW, False, 0&)
    AdPos
    RetWe = SendMessage(Me.hwnd, WM_SETREDRAW, True, 0&)
    RetWe = GetClientRect(Me.hwnd, ClRe)
    RetWe = RedrawWindow(Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
    GlRzA = True
End If

End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Akont False
End Sub
Private Sub txtS1F01_GotFocus()
    If GlAdL = False Then
        Me.txtS1F01.SelStart = 0
        Me.txtS1F01.SelLength = Len(Me.txtS1F01.Text)
    End If
End Sub

Private Sub txtS1F02_Change()

TagWe = Mid$(Me.txtS1F02.Tag, 2, Len(Me.txtS1F02.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F02.Tag = "1" & TagWe
    GlAdS = True
    AAnre
End If

End Sub
Private Sub txtS1F05_LostFocus()

If GlAdL = False Then
    AKopi
    DoEvents
    AdBrf
    DoEvents
    Me.txtS2F20.Text = Me.cmbS1F10.Text
End If

End Sub

Private Sub txtS1F08_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F08_LostFocus()
    If GlAdL = False Then
        Adr_Pos False
    End If
End Sub
Private Sub txtS1F09_Click()
    If GlAdL = False Then
        If Me.txtS1F08.Text <> vbNullString Then
            If Me.txtS1F09.Text = vbNullString Then
                Adr_Pos False
            End If
        End If
    End If
End Sub
Private Sub txtS1F09_LostFocus()
On Error Resume Next

If Me.txtS2F12.Text = vbNullString Then
    If Me.txtS1F02.Text <> vbNullString Then
        Me.txtS2F12.Text = Me.txtS1F02.Text
        TagWe = Mid$(Me.txtS2F12.Tag, 2, Len(Me.txtS2F12.Tag) - 1)
        Me.txtS2F12.Tag = "1" & TagWe
    End If
End If
If Me.txtS2F11.Text = vbNullString Then
    If Me.txtS1F01.Text <> vbNullString Then
        Me.txtS2F11.Text = Me.txtS1F01.Text
    End If
End If
If Me.txtS1F02.Text = Me.txtS2F12.Text Then
    If Me.txtS2F13.Text = vbNullString Then If Me.txtS1F03.Text <> vbNullString Then Me.txtS2F13.Text = Me.txtS1F03.Text
    If Me.txtS2F14.Text = vbNullString Then If Me.txtS1F04.Text <> vbNullString Then Me.txtS2F14.Text = Me.txtS1F04.Text
    If Me.txtS2F15.Text = vbNullString Then If Me.txtS1F05.Text <> vbNullString Then Me.txtS2F15.Text = Me.txtS1F05.Text
End If
If Me.txtS2F16.Text = vbNullString Then If Me.txtS1F06.Text <> vbNullString Then Me.txtS2F16.Text = Me.txtS1F06.Text
If Me.txtS2F18.Text = vbNullString Then If Me.txtS1F08.Text <> vbNullString Then Me.txtS2F18.Text = Me.txtS1F08.Text
If Me.txtS2F19.Text = vbNullString Then If Me.txtS1F09.Text <> vbNullString Then Me.txtS2F19.Text = Me.txtS1F09.Text

End Sub
Private Sub cmbS1F10_Click()

If GlAdL = False Then
    TagWe = Mid$(Me.cmbS1F10.Tag, 2, Len(Me.cmbS1F10.Tag) - 1)
    Me.cmbS1F10.Tag = "1" & TagWe
    GlAdS = True 'Speichern der Adresse erforderlich
    
    TagWe = Mid$(Me.txtS1F20.Tag, 2, Len(Me.txtS1F20.Tag) - 1)
    Me.txtS1F20.Tag = "1" & TagWe
    GlAdS = True
    Me.txtS1F20.Text = Me.cmbS1F10.ItemData(Me.cmbS1F10.ListIndex)

    TagWe = Mid$(Me.txtS2F20.Tag, 2, Len(Me.txtS2F20.Tag) - 1)
    Me.txtS2F20.Tag = "1" & TagWe
    GlAdS = True
    Me.txtS2F20.Text = Me.cmbS1F10.Text
End If

End Sub
Private Sub txtS1F12_Click()

TagWe = Mid$(Me.txtS1F12.Tag, 2, Len(Me.txtS1F12.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F12.Tag = "1" & TagWe
    GlAdS = True
    Me.txtS2F22.Text = Me.txtS1F12.Text
End If

End Sub
Private Sub txtS1F13_LostFocus()
On Error Resume Next

Dim NeuDa As Date

If IsDate(Me.txtS1F13.Text) Then
    NeuDa = Me.txtS1F13.Text
    If Year(NeuDa) > 1900 Then
        Me.txtS1F13.Text = NeuDa
        If Me.txtS2F25.Text = vbNullString Then
            If Me.txtS1F13.Text <> vbNullString Then
                Me.txtS2F25.Text = Me.txtS1F13.Text
            End If
        End If
        AAnre
    Else
        Me.txtS1F13.Text = vbNullString
        SPopu "Falsches Geburtsdatum", "Das Geburtshagr ist älter als 1900!", IC48_Forbidden
    End If
End If

End Sub
Private Sub txtS1F30_Validate(Cancel As Boolean)
    If (Not txtS1F30.isValid) Then Cancel = True
End Sub

Private Sub txtS1F33_Change()

TagWe = Mid$(Me.txtS1F33.Tag, 2, Len(Me.txtS1F33.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F33.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F33_GotFocus()
    Me.txtS1F33.SelStart = 0
    Me.txtS1F33.SelLength = Len(Me.txtS1F33.Text)
    GlAdL = False
End Sub

Private Sub txtS1F33_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F33_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F33.SelLength = 0
    Case vbKeyDown: Me.cmbS2F36.SetFocus
    Case vbKeyUp: Me.txtS2F33.SetFocus
    End Select
End Sub
Private Sub txtS1F37_Change()

TagWe = Mid$(Me.txtS1F37.Tag, 2, Len(Me.txtS1F37.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F37.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F37_GotFocus()
    Me.txtS1F37.SelStart = 0
    Me.txtS1F37.SelLength = Len(Me.txtS1F37.Text)
    GlAdL = False
End Sub
Private Sub txtS1F37_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F37_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F37.SelLength = 0
    Case vbKeyDown: Me.txtS1F30.SelLength = 0
    Case vbKeyUp: Me.txtS1F32.SetFocus
    End Select
End Sub

Private Sub txtS2F12_Change()

TagWe = Mid$(Me.txtS2F12.Tag, 2, Len(Me.txtS2F12.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F12.Tag = "1" & TagWe
    GlAdS = True
    AdBrf
End If

End Sub
Private Sub txtS2F13_LostFocus()
    If GlAdL = False Then
        AdBrf
    End If
End Sub

Private Sub txtS2F14_LostFocus()
    If GlAdL = False Then
        AdBrf
    End If
End Sub

Private Sub txtS2F15_LostFocus()
    If GlAdL = False Then
        AdBrf
    End If
End Sub

Private Sub txtS2F18_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F18_LostFocus()
    If GlAdL = False Then
        Adr_Pos True
    End If
End Sub
Private Sub txtS2F19_Click()
    If GlAdL = False Then
        If Me.txtS2F18.Text <> vbNullString Then
            If Me.txtS2F19.Text = vbNullString Then
                Adr_Pos True
            End If
        End If
    End If
End Sub
Private Sub txtS2F22_Change()

TagWe = Mid$(Me.txtS2F22.Tag, 2, Len(Me.txtS2F22.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F22.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F25_LostFocus()
On Error Resume Next

Dim NeuDa As Date

If IsDate(Me.txtS2F25.Text) Then
    NeuDa = Me.txtS2F25.Text
    If Year(NeuDa) > 1900 Then
        Me.txtS2F25.Text = NeuDa
    Else
        Me.txtS2F25.Text = vbNullString
        SPopu "Falsches Geburtsdatum", "Das Geburtshagr ist älter als 1900!", IC48_Forbidden
    End If
End If

End Sub
Private Sub txtS2F26_Click()

TagWe = Mid$(Me.txtS2F26.Tag, 2, Len(Me.txtS2F26.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F26.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F26_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F26_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtS2F26.SelLength = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    TimEnde 2
    If GlRes = False Then 'Reset der Einstellungen
        If GlAdS = True Then 'Speichern der Adresse erforderlich
            AKopi 'Kopieren der Adressdaten in die Rechungsanschrift
            DoEvents
            AErAd 'Erstellt die Anschrift im Anschriftenfeld
            DoEvents
            FSave True
        End If
    End If
    FClos
    AMeAc True 'Schaltet das Menü ein / aus
    Set frmAdress = Nothing
End Sub
Private Sub FSave(Optional ByVal SaFra As Boolean = False)
On Error GoTo SaErr
'Überprüft, ob der Eintrag geändert wurde und speichert dieses ab

Dim AdPIN As Long
Dim RowNr As Long
Dim PaStr As String
Dim DoSav As Boolean
Dim StSei As Boolean 'Willkommensseite
Dim Frage As Integer
Dim Mld1, Mld2, Tit1 As String
Dim FeKur As XtremeSuiteControls.FlatEdit
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim PuGut As XtremeSuiteControls.PushButton

Set FM = frmAdress
Set PuGut = FM.btnGutha
Set FePIN = FM.txtS1F30
Set FeKur = FM.txtS1F11
Set PrGr1 = FM.prpGrid1
Set PrGr2 = FM.prpGrid2
Set PrGr3 = FM.prpGrid3
Set CmBrs = FM.comBar01
Set RpCo2 = frmMain.repCont2
Set RpCo4 = frmMain.repCont4
Set RpCo3 = frmMain.repCont3
Set RpSel = RpCo2.SelectedRows

Tit1 = "Adresse Speichern"
Mld1 = "Soll diese Adresse gespeichert werden?"
Mld2 = "Diese Adresse existiert bereits. Soll diese trotzdem gespeichert werden?"

If PrGr1.Visible = True Then
    PrGr1.SetFocus
ElseIf PrGr3.Visible = True Then
    If GlAdN = False Then
        PrGr3.SetFocus
    End If
End If

If GlTza = True Then 'Testzeit abgelaufen
    SPopu "Lizenzierung erforderlich!", "Es ist keine bzw. keine gültige Seriennummer vorhanden oder die Testzeit ist abgelaufen.", IC48_Forbidden
    Exit Sub
End If

If GlAdS = True Then 'Speichern der Adresse erforderlich
    If FeKur.Text <> vbNullString Then
        PaStr = FeKur.Text
        If SaFra = True Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        Else
            Frage = 6
        End If
        If Frage = 6 Then
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                RowNr = RpRow.Index
            End If

            AEnab True
            If GlAdN = True Then 'Neue Adresse anlegen
                If Adr_Dop(PaStr) = True Then
                    Frage = WindowMess(Mld2, Dial1, Tit1, FM.hwnd)
                    If Frage = 6 Then
                        DoSav = True
                    End If
                Else
                    DoSav = True
                End If
                If DoSav = True Then
                    If Adr_San = True Then
                        Select Case GlBut:
                        Case RibTab_Startseite:
                                    StSei = True
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    Select Case GlAdO
                                    Case 0: SReZe GlAdr
                                    Case 1: SKrZe GlAdr
                                    Case 2: SKrZe GlAdr
                                    End Select
                        Case RibTab_Adressen:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpAd True
                        Case RibTab_Mandanten:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpAd True
                        Case RibTab_Verordner:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpAd True
                        Case RibTab_Mitarbeit:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpAd True
                        Case RibTab_Fragebogen:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpAn
                        Case RibTab_Krankenbla:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpKr
                        Case RibTab_Abrechnung:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpAb
                        Case RibTab_Vorbereit:

                        Case RibTab_Tagesproto:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                        Case RibTab_Ter_Kalend:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    If WindowLoad("frmTermin") = True Then
                                        Set FM = frmTermin
                                        FM.txtID0.Text = GlAdr
                                        FM.txtAdres.Text = PaStr
                                        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
                                        FM.txtID0.Tag = 1 & TagWe
                                        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
                                        FM.txtAdres.Tag = 1 & TagWe
                                        GlTSa = True 'WICHTIG!
                                    ElseIf WindowLoad("frmTermVo") = True Then
                                        frmTermVo.txtID0.Text = GlAdr
                                        frmTermVo.txtAdres.Text = PaStr
                                    Else
                                        If GlWaN = True Then 'Wartenden Erfassen und Aktualisieren
                                            GlWaN = False
                                            Ter_Edi GlAdr, True 'in Warteliste aufnehmen
                                            DoEvents
                                            P_List "TeDe", 0, 2
                                        End If
                                    End If
                        Case RibTab_Ter_Raeume:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    If WindowLoad("frmTermin") = True Then
                                        Set FM = frmTermin
                                        FM.txtID0.Text = GlAdr
                                        FM.txtAdres.Text = PaStr
                                        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
                                        FM.txtID0.Tag = 1 & TagWe
                                        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
                                        FM.txtAdres.Tag = 1 & TagWe
                                    ElseIf WindowLoad("frmTermVo") = True Then
                                        frmTermVo.txtID0.Text = GlAdr
                                        frmTermVo.txtAdres.Text = PaStr
                                    Else
                                        If GlWaN = True Then
                                            GlWaN = False
                                            Ter_Edi GlAdr, True 'in Warteliste aufnehmen
                                            DoEvents
                                            P_List "TeDe", 0, 2
                                        End If
                                    End If
                        Case RibTab_Ter_Mitarb:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    If WindowLoad("frmTermin") = True Then
                                        Set FM = frmTermin
                                        FM.txtID0.Text = GlAdr
                                        FM.txtAdres.Text = PaStr
                                        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
                                        FM.txtID0.Tag = 1 & TagWe
                                        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
                                        FM.txtAdres.Tag = 1 & TagWe
                                    ElseIf WindowLoad("frmTermVo") = True Then
                                        frmTermVo.txtID0.Text = GlAdr
                                        frmTermVo.txtAdres.Text = PaStr
                                    Else
                                        If GlWaN = True Then
                                            GlWaN = False
                                            Ter_Edi GlAdr, True 'in Warteliste aufnehmen
                                            DoEvents
                                            P_List "TeDe", 0, 2
                                        End If
                                    End If
                        Case RibTab_Ter_Listen:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    If WindowLoad("frmTermin") = True Then
                                        Set FM = frmTermin
                                        FM.txtID0.Text = GlAdr
                                        FM.txtAdres.Text = PaStr
                                        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
                                        FM.txtID0.Tag = 1 & TagWe
                                        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
                                        FM.txtAdres.Tag = 1 & TagWe
                                    ElseIf WindowLoad("frmTermVo") = True Then
                                        frmTermVo.txtID0.Text = GlAdr
                                        frmTermVo.txtAdres.Text = PaStr
                                    End If
                        Case RibTab_Ter_Akont:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    If WindowLoad("frmTermin") = True Then
                                        Set FM = frmTermin
                                        FM.txtID0.Text = GlAdr
                                        FM.txtAdres.Text = PaStr
                                        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
                                        FM.txtID0.Tag = 1 & TagWe
                                        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
                                        FM.txtAdres.Tag = 1 & TagWe
                                    ElseIf WindowLoad("frmTermVo") = True Then
                                        frmTermVo.txtID0.Text = GlAdr
                                        frmTermVo.txtAdres.Text = PaStr
                                    End If
                        Case RibTab_Ter_Warte:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    If WindowLoad("frmTermin") = True Then
                                        Set FM = frmTermin
                                        FM.txtID0.Text = GlAdr
                                        FM.txtAdres.Text = PaStr
                                        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
                                        FM.txtID0.Tag = 1 & TagWe
                                        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
                                        FM.txtAdres.Tag = 1 & TagWe
                                    ElseIf WindowLoad("frmTermVo") = True Then
                                        frmTermVo.txtID0.Text = GlAdr
                                        frmTermVo.txtAdres.Text = PaStr
                                    Else
                                        If GlWaN = True Then
                                            GlWaN = False
                                            Ter_Edi GlAdr, True 'in Warteliste aufnehmen
                                            DoEvents
                                            P_List "TeDe", 0, 2
                                        End If
                                    End If
                        Case RibTab_Tex_Email:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    S_MaMa 5, GlAdr, PaStr
                                    SUpAd
                        Case Else:
                                    GlAdr = S_AdGui(GlAdG, "ID0")
                                    SUpAd
                        End Select
                        GlAId = GlAdr
                        GlTDa = vbNullString 'Wichtig für Textverarbeitung
                        PrGr3.Enabled = True
                    Else
                        GlAdN = False
                    End If
                Else
                    GlAdN = False
                End If
            Else
                AdBrf
                DoEvents
                Adr_Sav
                DoEvents
                Select Case GlBut:
                Case RibTab_Adressen:
                            SUpAd False, RowNr
                Case RibTab_Mandanten:
                            SUpAd False, RowNr
                Case RibTab_Verordner:
                            SUpAd False, RowNr
                Case RibTab_Mitarbeit:
                            SUpAd False, RowNr
                Case RibTab_Fragebogen:
                            SUpAn
                Case RibTab_Krankenbla:
                            SUpKr
                Case RibTab_Abrechnung:
                            Set RpSel = RpCo3.SelectedRows
                            If RpSel.Count > 0 Then
                                Set RpRow = RpSel(0)
                                RowNr = RpRow.Index
                                SUpAb RowNr
                            Else
                                SUpAb
                            End If
                            Set RpSel = RpCo4.SelectedRows
                            If RpSel.Count > 0 Then
                                Set RpRow = RpSel(0)
                                RowNr = RpRow.Index
                                SUpRe RowNr
                            Else
                                SUpRe
                            End If
                Case RibTab_Ter_Kalend:
                            If GlWaN = True Then
                                GlWaN = False
                                P_List "TeDe", 0, 2
                            End If
                Case RibTab_Ter_Raeume:
                            If GlWaN = True Then
                                GlWaN = False
                                P_List "TeDe", 0, 2
                            End If
                Case RibTab_Ter_Mitarb:
                            If GlWaN = True Then
                                GlWaN = False
                                P_List "TeDe", 0, 2
                            End If
                
                Case RibTab_Ter_Warte:
                            If GlWaN = True Then
                                GlWaN = False
                                P_List "TeDe", 0, 2
                            End If
                Case Else:
                            SUpAd
                End Select
                PuGut.Enabled = True
            End If
        End If
    End If
End If

GlAdS = False

If StSei = True Then Unload FM

Set CmBrs = Nothing
Set RpSel = Nothing
Set RpRow = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSaZu(Optional ByVal SaFra As Boolean = False)
On Error GoTo SaErr
'Überprüft, ob der Eintrag geändert wurde und speichert dieses ab

Dim RowNr As Long
Dim RowFi As Long
Dim PaStr As String
Dim FeKur As XtremeSuiteControls.FlatEdit
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmAdress
Set FeKur = FM.txtS4F19
Set CmBrs = FM.comBar01
Set RpCo5 = FM.repCont5
Set RpRws = RpCo5.Rows

FeKur.Text = AErZu(1) 'Kurzbezeichnung erstellen
DoEvents

If GlAdZ = True Then 'Speichern der Zugehörigen erforderlich
    If FeKur.Text <> vbNullString Then
        PaStr = FeKur.Text
        If GlAzN = True Then 'Neue Zugeordnete anlegen
            If Adr_ZuSa(True) = True Then
                RowNr = Adr_ZuLa(GlAzG) 'Zugehörigen Laden
                If RowNr > 0 Then
                    If RpRws.Count > 0 Then
                        RowFi = RpCo5.TopRowIndex
                    End If
                End If
                If RowNr >= RpRws.Count Then
                    RowNr = RpRws.Count - 1
                End If
                RpCo5.TopRowIndex = RowFi
                RpRws.Row(0).Selected = False
                RpRws.Row(RowNr).EnsureVisible
                RpRws.Row(RowNr).Selected = True
                If GlFoc = True Then
                    Set RpCo5.FocusedRow = RpRws.Row(RowNr)
                End If
                AZuLa
            Else
                GlAzN = False
            End If
        Else
            If Adr_ZuSa(False) = True Then
                RowNr = Adr_ZuLa(GlAzG) 'Zugehörigen Laden
                If RowNr > 0 Then
                    If RpRws.Count > 0 Then
                        RowFi = RpCo5.TopRowIndex
                    End If
                End If
                If RowNr >= RpRws.Count Then
                    RowNr = RpRws.Count - 1
                End If
                RpCo5.TopRowIndex = RowFi
                RpRws.Row(0).Selected = False
                RpRws.Row(RowNr).EnsureVisible
                RpRws.Row(RowNr).Selected = True
                If GlFoc = True Then
                    Set RpCo5.FocusedRow = RpRws.Row(RowNr)
                End If
                AZuLa
            Else
                GlAzN = False
            End If
        End If
    End If
End If

GlAdZ = False

Set CmBrs = Nothing
Set RpRws = Nothing
Set RpCo5 = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSaZu " & Err.Number
Resume Next

End Sub

Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim TxDum As XtremeSuiteControls.FlatEdit

Set FM = frmAdress
Set CmBrs = FM.comBar01
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set PrGr1 = FM.prpGrid1
Set PrGr2 = FM.prpGrid2
Set PrGr3 = FM.prpGrid3
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set TxDum = FM.txtDummy
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

TxDum.SetFocus
DoEvents

GlAdL = True

TabId = RbTab.id

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case TabId
Case RibTab_Adr_Haupt:
    Rahm1.Visible = True
    Rahm2.Visible = True
    Rahm3.Visible = True
    Rahm4.Visible = False
    Rahm5.Visible = False
    PrGr1.Visible = False
    PrGr2.Visible = False
    PrGr3.Visible = False
    RpCo2.Visible = False
    RpCo3.Visible = False
    RpCo4.Visible = False
    RpCo5.Visible = False
Case RibTab_Adr_Dokum:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    PrGr1.Visible = True
    PrGr2.Visible = True
    PrGr3.Visible = False
    RpCo2.Visible = False
    RpCo3.Visible = False
    RpCo4.Visible = False
    RpCo5.Visible = False
Case RibTab_Adr_Eigen:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    PrGr1.Visible = False
    PrGr2.Visible = False
    PrGr3.Visible = True
    RpCo2.Visible = False
    RpCo3.Visible = False
    RpCo4.Visible = False
    RpCo5.Visible = False
Case RibTab_Adr_Booki:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    PrGr1.Visible = False
    PrGr2.Visible = False
    PrGr3.Visible = False
    RpCo2.Visible = True
    RpCo3.Visible = True
    RpCo4.Visible = True
    RpCo5.Visible = False
Case RibTab_Adr_Membe:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = True
    Rahm5.Visible = False
    PrGr1.Visible = False
    PrGr2.Visible = False
    PrGr3.Visible = False
    RpCo2.Visible = False
    RpCo3.Visible = False
    RpCo4.Visible = False
    RpCo5.Visible = True
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set PrGr1 = Nothing
Set PrGr2 = Nothing
Set PrGr3 = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

GlAdL = False

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

Set FM = frmAdress

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: ASper True
            ANeue True
            Kon_Lis
Case KY_F5: frmAdrSuch.Show vbModal
Case KY_F6: frmGrupp.Show vbModal
Case KY_F7: AOutl
Case KY_F8: AKopi
            AErAd 'Erstellt die Anschrift im Anschriftenfeld
            FSave
Case KY_F11: FEnde
Case AM_Patient_Speichern: AKopi
                           AErAd 'Erstellt die Anschrift im Anschriftenfeld
                           FSave
Case AM_Hilfe: FHilfe
Case AM_Patient_Gruppe: frmGrupp.Show vbModal
Case AM_Beenden: FEnde
Case AM_Patient_Copy: AdKop
Case AM_Patient_Del: Adr_Loe
Case AM_Patient_Such: frmAdrSuch.Show vbModal
Case AM_Patient_Clip1: FClip 1
Case AM_Patient_Clip2: FClip 2
Case AM_Notiz_Neu: Akont True
Case AM_Notiz_Bearbeit: Akont False
Case AM_Notiz_Loeschen: Kon_Loe 2
Case AM_Extras_Optionen: frmOptions.Show
Case AM_Programmhilfe: FHilfe
Case AM_Geburtsdatum: AGebu
                      AKopi
                      AErAd 'Erstellt die Anschrift im Anschriftenfeld
                      FSave
Case AD_Patient_Add: ASper True
                     ANeue True
                     Kon_Lis
Case AM_Guthaben: AGuth
                  FSave
Case AD_Patient_Copy: AdKop
Case AD_Patient_Del: Adr_Loe
Case AD_Adressen_Chipkarte: AChip
Case AD_Patienten_Suchen: frmAdrSuch.Show vbModal
Case AD_Patienten_Save: AKopi
                        AErAd 'Erstellt die Anschrift im Anschriftenfeld
                        FSave
Case AD_Patienten_Gruppe: frmGrupp.Show vbModal
Case AD_Speichern_Nroma: AKopi
                         AErAd 'Erstellt die Anschrift im Anschriftenfeld
                         FSave
Case AD_Speichern_Close: FEnde
Case AD_Adressen_SMS: ATerm
Case AD_Adressen_Warten: frmWaKom.Show vbModal
Case AD_Adressen_GDT_Ex: S_AdGDT
Case AD_Adressen_GDT_Im: S_GDT True
Case AD_Adressen_Passw: ATerm True
Case AD_Member_Add: AZuNe
Case AD_Member_Del: Adr_ZuLo
Case AD_Member_Copy: Adr_ZuKo
Case AD_Member_Orig: Adr_ZuHa
Case AD_Member_Save: FSaZu True
Case Else:
    If TolId < 0 Then
        APaSe TolId
    End If
End Select

GlToo = False

End Sub
Private Sub txtS2F25_Change()

TagWe = Mid$(Me.txtS2F25.Tag, 2, Len(Me.txtS2F25.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F25.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F25_GotFocus()
    Me.txtS2F25.SelStart = 0
    Me.txtS2F25.SelLength = Len(Me.txtS2F25.Text)
    GlAdL = False
End Sub
Private Sub txtS2F25_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F25_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F25.SelLength = 0
    Case vbKeyDown: Me.txtS2F08.SetFocus
    Case vbKeyUp: Me.txtS1F22.SetFocus
    End Select
End Sub
Private Sub txtS1F01_Change()

TagWe = Mid$(Me.txtS1F01.Tag, 2, Len(Me.txtS1F01.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F01.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F01_KeyDown(KeyCode As Integer, Shift As Integer)

TagWe = Mid$(Me.txtS1F01.Tag, 2, Len(Me.txtS1F01.Tag) - 1)

If Left$(Me.txtS1F01.Tag, 1) = 0 Then
    Me.txtS1F01.Tag = "1" & TagWe
    GlAdL = False
    GlAdS = True
End If

End Sub

Private Sub txtS1F01_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F01_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F01.SelLength = 0
    Case vbKeyDown: Me.txtS1F02.SetFocus
    Case vbKeyUp:
    End Select
End Sub

Private Sub txtS1F02_Click()

TagWe = Mid$(Me.txtS1F02.Tag, 2, Len(Me.txtS1F02.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F02.Tag = "1" & TagWe
    GlAdS = True
    AAnre
End If

End Sub
Private Sub txtS1F02_GotFocus()
    GlAdL = False
End Sub

Private Sub txtS1F02_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F02_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F02.SelLength = 0
    Case vbKeyDown:
    Case vbKeyUp:
    End Select
End Sub
Private Sub txtS1F03_Change()

TagWe = Mid$(Me.txtS1F03.Tag, 2, Len(Me.txtS1F03.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F03.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F03_GotFocus()
    Me.txtS1F03.SelStart = 0
    Me.txtS1F03.SelLength = Len(Me.txtS1F03.Text)
    GlAdL = False
End Sub

Private Sub txtS1F03_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F03_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F03.SelLength = 0
    Case vbKeyDown: Me.txtS1F04.SetFocus
    Case vbKeyUp: Me.txtS1F02.SetFocus
    End Select
End Sub
Private Sub txtS1F04_Change()

TagWe = Mid$(Me.txtS1F04.Tag, 2, Len(Me.txtS1F04.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F04.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F04_GotFocus()
    Me.txtS1F04.SelStart = 0
    Me.txtS1F04.SelLength = Len(Me.txtS1F04.Text)
    GlAdL = False
End Sub

Private Sub txtS1F04_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F04_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F04.SelLength = 0
    Case vbKeyDown: Me.txtS1F05.SetFocus
    Case vbKeyUp: Me.txtS1F03.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS1F04.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS1F04.Text = UCase(Chr$(KeyCode))
                            Me.txtS1F04.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub
Private Sub txtS1F05_Change()

TagWe = Mid$(Me.txtS1F05.Tag, 2, Len(Me.txtS1F05.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F05.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F05_GotFocus()
    Me.txtS1F05.SelStart = 0
    Me.txtS1F05.SelLength = Len(Me.txtS1F05.Text)
    GlAdL = False
End Sub

Private Sub txtS1F05_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F05_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F05.SelLength = 0
    Case vbKeyDown: Me.txtS1F06.SetFocus
    Case vbKeyUp: Me.txtS1F04.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS1F05.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS1F05.Text = UCase(Chr$(KeyCode))
                            Me.txtS1F05.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub
Private Sub txtS1F06_Change()

TagWe = Mid$(Me.txtS1F06.Tag, 2, Len(Me.txtS1F06.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F06.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F06_GotFocus()
    Me.txtS1F06.SelStart = 0
    Me.txtS1F06.SelLength = Len(Me.txtS1F06.Text)
    GlAdL = False
End Sub
Private Sub txtS1F06_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F06_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F06.SelLength = 0
    Case vbKeyDown: Me.txtS1F08.SetFocus
    Case vbKeyUp: Me.txtS1F05.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS1F06.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS1F06.Text = UCase(Chr$(KeyCode))
                            Me.txtS1F06.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub
Private Sub txtS1F08_Change()

TagWe = Mid$(Me.txtS1F08.Tag, 2, Len(Me.txtS1F08.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F08.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F08_GotFocus()
    Me.txtS1F08.SelStart = 0
    Me.txtS1F08.SelLength = Len(Me.txtS1F08.Text)
    GlAdL = False
End Sub

Private Sub txtS1F08_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F08.SelLength = 0
    Case vbKeyDown: Me.txtS1F09.SetFocus
    Case vbKeyUp: Me.txtS1F06.SetFocus
    End Select
End Sub
Private Sub txtS1F09_Change()

TagWe = Mid$(Me.txtS1F09.Tag, 2, Len(Me.txtS1F09.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F09.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F09_GotFocus()
    Me.txtS1F09.SelStart = 0
    Me.txtS1F09.SelLength = Len(Me.txtS1F09.Text)
    GlAdL = False
End Sub

Private Sub txtS1F09_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F09_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F09.SelLength = 0
    Case vbKeyDown: Me.txtS1F12.SetFocus
    Case vbKeyUp: Me.txtS1F08.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS1F09.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS1F09.Text = UCase(Chr$(KeyCode))
                            Me.txtS1F09.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub
Private Sub cmbS1F10_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F11_Change()

TagWe = Mid$(Me.txtS1F11.Tag, 2, Len(Me.txtS1F11.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F11.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F11_GotFocus()
    Me.txtS1F11.SelStart = 0
    Me.txtS1F11.SelLength = Len(Me.txtS1F11.Text)
    GlAdL = False
End Sub

Private Sub txtS1F11_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F11_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F11.SelLength = 0
    Case vbKeyDown: Me.txtS2F24.SetFocus
    Case vbKeyUp: Me.txtS1F12.SetFocus
    End Select
End Sub

Private Sub txtS1F12_GotFocus()
    GlAdL = False
End Sub
Private Sub txtS1F20_Change()

TagWe = Mid$(Me.txtS2F20.Tag, 2, Len(Me.txtS2F20.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F20.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F12_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F12_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F12.SelLength = 0
    Case vbKeyDown:
    Case vbKeyUp:
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS1F12.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS1F12.Text = UCase(Chr$(KeyCode))
                            Me.txtS1F12.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub
Private Sub txtS1F13_Change()

TagWe = Mid$(Me.txtS1F13.Tag, 2, Len(Me.txtS1F13.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F13.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F13_GotFocus()
    Me.txtS1F13.SelStart = 0
    Me.txtS1F13.SelLength = Len(Me.txtS1F13.Text)
    GlAdL = False
End Sub
Private Sub txtS1F13_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F13_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F13.SelLength = 0
    Case vbKeyDown: Me.cmbS1F08.SetFocus
    Case vbKeyUp: Me.txtS2F24.SetFocus
    End Select
End Sub
Private Sub txtS1F14_Change()

TagWe = Mid$(Me.txtS1F14.Tag, 2, Len(Me.txtS1F14.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F14.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F14_GotFocus()
    Me.txtS1F14.SelStart = 0
    Me.txtS1F14.SelLength = Len(Me.txtS1F14.Text)
    GlAdL = False
End Sub

Private Sub txtS1F14_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F14_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F14.SelLength = 0
    Case vbKeyDown: Me.cmbS1F21.SetFocus
    Case vbKeyUp: Me.cmbS1F07.SetFocus
    End Select
End Sub

Private Sub txtS1F15_Change()

TagWe = Mid$(Me.txtS1F15.Tag, 2, Len(Me.txtS1F15.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F15.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F15_GotFocus()
    Me.txtS1F15.SelStart = Len(Me.txtS1F15.Text)
    Me.txtS1F15.SelLength = 0
    GlAdL = False
End Sub

Private Sub txtS1F15_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F15_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F15.SelLength = 0
    Case vbKeyDown: Me.txtS1F16.SetFocus
    Case vbKeyUp: Me.txtS1F13.SetFocus
    End Select
End Sub

Private Sub txtS1F16_Change()

TagWe = Mid$(Me.txtS1F16.Tag, 2, Len(Me.txtS1F16.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F16.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F16_GotFocus()
    Me.txtS1F16.SelStart = Len(Me.txtS1F16.Text)
    Me.txtS1F16.SelLength = 0
    GlAdL = False
End Sub

Private Sub txtS1F16_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F16_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F16.SelLength = 0
    Case vbKeyDown: Me.txtS1F17.SetFocus
    Case vbKeyUp: Me.txtS1F15.SetFocus
    End Select
End Sub

Private Sub txtS1F17_Change()

TagWe = Mid$(Me.txtS1F17.Tag, 2, Len(Me.txtS1F17.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F17.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F17_GotFocus()
    Me.txtS1F17.SelStart = Len(Me.txtS1F17.Text)
    Me.txtS1F17.SelLength = 0
    GlAdL = False
End Sub

Private Sub txtS1F17_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F17_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F17.SelLength = 0
    Case vbKeyDown: Me.txtS1F18.SetFocus
    Case vbKeyUp: Me.txtS1F16.SetFocus
    End Select
End Sub

Private Sub txtS1F18_Change()

TagWe = Mid$(Me.txtS1F18.Tag, 2, Len(Me.txtS1F18.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F18.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F18_GotFocus()
    Me.txtS1F18.SelStart = Len(Me.txtS1F18.Text)
    Me.txtS1F18.SelLength = 0
    GlAdL = False
End Sub

Private Sub txtS1F18_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F18_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F18.SelLength = 0
    Case vbKeyDown: Me.txtS1F19.SetFocus
    Case vbKeyUp: Me.txtS1F17.SetFocus
    End Select
End Sub

Private Sub txtS1F19_Change()

TagWe = Mid$(Me.txtS1F19.Tag, 2, Len(Me.txtS1F19.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F19.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F19_GotFocus()
    Me.txtS1F19.SelStart = 0
    Me.txtS1F19.SelLength = Len(Me.txtS1F19.Text)
    GlAdL = False
End Sub

Private Sub txtS1F19_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F19_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F19.SelLength = 0
    Case vbKeyDown: Me.txtS2F34.SetFocus
    Case vbKeyUp: Me.txtS1F18.SetFocus
    End Select
End Sub

Private Sub txtS1F27_Change()

TagWe = Mid$(Me.txtS1F27.Tag, 2, Len(Me.txtS1F27.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F27.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F27_GotFocus()
    Me.txtS1F27.SelStart = 0
    Me.txtS1F27.SelLength = Len(Me.txtS1F27.Text)
    GlAdL = False
End Sub

Private Sub txtS1F27_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F27_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F27.SelLength = 0
    Case vbKeyDown: Me.cmbS2F07.SetFocus
    Case vbKeyUp: Me.txtS2F34.SetFocus
    End Select
End Sub

Private Sub txtS1F30_Change()
On Error Resume Next

Set FePIN = Me.txtS1F30

TagWe = Mid$(FePIN.Tag, 2, Len(FePIN.Tag) - 1)

If GlAdL = False Then
    FePIN.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F30_GotFocus()
    Me.txtS1F30.SelStart = 0
    Me.txtS1F30.SelLength = Len(Me.txtS1F30.Text)
    GlAdL = False
End Sub

Private Sub txtS1F30_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F30_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode
Case vbKeyDelete:
Case vbKeyBack:
Case vbKeyF2: Me.txtS1F30.SelLength = 0
Case vbKeyDown: Me.txtS2F11.SetFocus
Case vbKeyUp: Me.txtS1F37.SetFocus
Case vbKeyTab: Me.txtS2F11.SetFocus
Case 48 To 57:
Case 96 To 105:
Case Else: SPopu "Nur Zahlen erlaubt", "In diesem Feld werden nur numerische Werte gespeichert.", IC48_Warning
End Select

End Sub
Private Sub txtS1F31_Change()

TagWe = Mid$(Me.txtS1F31.Tag, 2, Len(Me.txtS1F31.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F31.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F31_GotFocus()
    Me.txtS1F32.SelStart = 0
    Me.txtS1F32.SelLength = Len(Me.txtS1F32.Text)
    GlAdL = False
End Sub
Private Sub txtS1F31_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F31_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F31.SelLength = 0
    Case vbKeyDown: Me.txtS1F32.SetFocus
    Case vbKeyUp: Me.cmbS2F31.SetFocus
    End Select
End Sub
Private Sub txtS1F32_Change()

TagWe = Mid$(Me.txtS1F32.Tag, 2, Len(Me.txtS1F32.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F32.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F32_GotFocus()
    Me.txtS1F32.SelStart = 0
    Me.txtS1F32.SelLength = Len(Me.txtS1F32.Text)
    GlAdL = False
End Sub
Private Sub txtS1F32_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F32_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS1F32.SelLength = 0
    Case vbKeyDown: Me.txtS1F37.SetFocus
    Case vbKeyUp: Me.txtS1F31.SetFocus
    End Select
End Sub
Private Sub txtS2F03_Change()

TagWe = Mid$(Me.txtS2F03.Tag, 2, Len(Me.txtS2F03.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F03.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F03_GotFocus()
    Me.txtS2F03.SelStart = 0
    Me.txtS2F03.SelLength = Len(Me.txtS2F03.Text)
    GlAdL = False
End Sub

Private Sub txtS2F03_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F03_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F03.SelLength = 0
    Case vbKeyDown: Me.txtS2F33.SetFocus
    Case vbKeyUp: Me.txtS2F05.SetFocus
    End Select
End Sub

Private Sub txtS2F05_Change()

TagWe = Mid$(Me.txtS2F05.Tag, 2, Len(Me.txtS2F05.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F05.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F05_GotFocus()
    Me.txtS2F05.SelStart = 0
    Me.txtS2F05.SelLength = Len(Me.txtS2F05.Text)
    GlAdL = False
End Sub

Private Sub txtS2F05_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F05_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F05.SelLength = 0
    Case vbKeyDown: Me.txtS2F03.SetFocus
    Case vbKeyUp: Me.cmbS2F29.SetFocus
    End Select
End Sub

Private Sub txtS2F08_Click()

TagWe = Mid$(Me.txtS2F08.Tag, 2, Len(Me.txtS2F08.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F08.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F08_GotFocus()
    GlAdL = False
End Sub
Private Sub txtS2F08_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F08_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtS2F08.SelLength = 0
End Sub

Private Sub txtS2F11_Change()

TagWe = Mid$(Me.txtS2F11.Tag, 2, Len(Me.txtS2F11.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F11.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F11_KeyDown(KeyCode As Integer, Shift As Integer)

TagWe = Mid$(Me.txtS2F11.Tag, 2, Len(Me.txtS2F11.Tag) - 1)

If Left$(Me.txtS2F11.Tag, 1) = 0 Then
    Me.txtS2F11.Tag = "1" & TagWe
    GlAdL = False
    GlAdS = True
End If

End Sub

Private Sub txtS2F11_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F11_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F11.SelLength = 0
    Case vbKeyDown: Me.txtS2F12.SetFocus
    Case vbKeyUp: Me.txtS1F32.SetFocus
    End Select
End Sub

Private Sub txtS2F12_Click()

TagWe = Mid$(Me.txtS2F12.Tag, 2, Len(Me.txtS2F12.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F12.Tag = "1" & TagWe
    GlAdS = True
    AdBrf
End If

End Sub
Private Sub txtS2F12_GotFocus()
    GlAdL = False
End Sub
Private Sub txtS2F12_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F12_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F12.SelLength = 0
    Case vbKeyDown:
    Case vbKeyUp:
    End Select
End Sub

Private Sub txtS2F13_Change()

TagWe = Mid$(Me.txtS2F13.Tag, 2, Len(Me.txtS2F13.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F13.Tag = "1" & TagWe
    GlAdS = True
    AdBrf
End If

End Sub
Private Sub txtS2F13_GotFocus()
    Me.txtS2F13.SelStart = 0
    Me.txtS2F13.SelLength = Len(Me.txtS2F13.Text)
    GlAdL = False
End Sub

Private Sub txtS2F13_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F13_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F13.SelLength = 0
    Case vbKeyDown: Me.txtS2F14.SetFocus
    Case vbKeyUp: Me.txtS2F12.SetFocus
    End Select
End Sub

Private Sub txtS2F14_Change()

TagWe = Mid$(Me.txtS2F14.Tag, 2, Len(Me.txtS2F14.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F14.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F14_GotFocus()
    Me.txtS2F14.SelStart = 0
    Me.txtS2F14.SelLength = Len(Me.txtS2F14.Text)
    GlAdL = False
End Sub

Private Sub txtS2F14_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F14_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F14.SelLength = 0
    Case vbKeyDown: Me.txtS2F15.SetFocus
    Case vbKeyUp: Me.txtS2F13.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS2F14.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS2F14.Text = UCase(Chr$(KeyCode))
                            Me.txtS2F14.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub

Private Sub txtS2F15_Change()

TagWe = Mid$(Me.txtS2F15.Tag, 2, Len(Me.txtS2F15.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F15.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F15_GotFocus()
    Me.txtS2F15.SelStart = 0
    Me.txtS2F15.SelLength = Len(Me.txtS2F15.Text)
    GlAdL = False
End Sub

Private Sub txtS2F15_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F15_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F15.SelLength = 0
    Case vbKeyDown: Me.txtS2F16.SetFocus
    Case vbKeyUp: Me.txtS2F14.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS2F15.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS2F15.Text = UCase(Chr$(KeyCode))
                            Me.txtS2F15.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub

Private Sub txtS2F16_Change()

TagWe = Mid$(Me.txtS2F16.Tag, 2, Len(Me.txtS2F16.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F16.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F16_GotFocus()
    Me.txtS2F16.SelStart = 0
    Me.txtS2F16.SelLength = Len(Me.txtS2F16.Text)
    GlAdL = False
End Sub

Private Sub txtS2F16_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F16_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F16.SelLength = 0
    Case vbKeyDown: Me.txtS2F18.SetFocus
    Case vbKeyUp: Me.txtS2F15.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS2F16.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS2F16.Text = UCase(Chr$(KeyCode))
                            Me.txtS2F16.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub

Private Sub txtS2F18_Change()

TagWe = Mid$(Me.txtS2F18.Tag, 2, Len(Me.txtS2F18.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F18.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F18_GotFocus()
    Me.txtS2F18.SelStart = 0
    Me.txtS2F18.SelLength = Len(Me.txtS2F18.Text)
    GlAdL = False
End Sub

Private Sub txtS2F18_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F18.SelLength = 0
    Case vbKeyDown: Me.txtS2F19.SetFocus
    Case vbKeyUp: Me.txtS2F16.SetFocus
    End Select
End Sub

Private Sub txtS2F19_Change()

TagWe = Mid$(Me.txtS2F19.Tag, 2, Len(Me.txtS2F19.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F19.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F19_GotFocus()
    Me.txtS2F19.SelStart = 0
    Me.txtS2F19.SelLength = Len(Me.txtS2F19.Text)
    GlAdL = False
End Sub

Private Sub txtS2F19_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F19_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F19.SelLength = 0
    Case vbKeyDown: Me.txtS2F22.SetFocus
    Case vbKeyUp: Me.txtS2F18.SetFocus
    Case Else:
            If GlRDP = False Then
                If Shift = 0 Then
                    If Len(Me.txtS2F19.Text) = 1 Then
                        If KeyCode > 47 Then
                            Me.txtS2F19.Text = UCase(Chr$(KeyCode))
                            Me.txtS2F19.SelStart = 1
                        End If
                    End If
                End If
            End If
    End Select
End Sub

Private Sub txtS2F22_GotFocus()
    Me.txtS2F22.SelStart = 0
    Me.txtS2F22.SelLength = Len(Me.txtS1F12.Text)
    GlAdL = False
End Sub

Private Sub txtS2F22_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F22_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F22.SelLength = 0
    Case vbKeyDown: Me.cmbS1F10.SetFocus
    Case vbKeyUp: Me.txtS2F25.SetFocus
    End Select
End Sub

Private Sub txtS2F24_Change()

TagWe = Mid$(Me.txtS2F24.Tag, 2, Len(Me.txtS2F24.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F24.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F24_GotFocus()
    GlAdL = False
End Sub
Private Sub txtS2F24_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F24_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F24.SelLength = 0
    Case vbKeyDown:  Me.txtS1F13.SetFocus
    Case vbKeyUp: Me.txtS1F12.SetFocus
    End Select
End Sub
Private Sub txtS2F27_Change()

TagWe = Mid$(Me.txtS2F27.Tag, 2, Len(Me.txtS2F27.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F27.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F27_GotFocus()
    Me.txtS2F27.SelStart = 0
    Me.txtS2F27.SelLength = Len(Me.txtS2F27.Text)
    GlAdL = False
End Sub
Private Sub txtS2F27_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F27_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F27.SelLength = 0
    Case vbKeyUp: Me.cmbS2F10.SetFocus
    End Select
End Sub
Private Sub txtS2F27_LostFocus()

Dim NeuDa As Date

If IsDate(Me.txtS2F27.Text) Then
    NeuDa = Me.txtS2F27.Text
    Me.txtS2F27.Text = NeuDa
End If

End Sub

Private Sub txtS2F33_Change()

TagWe = Mid$(Me.txtS2F33.Tag, 2, Len(Me.txtS2F33.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F33.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F33_GotFocus()
    Me.txtS2F33.SelStart = 0
    Me.txtS2F33.SelLength = Len(Me.txtS2F33.Text)
    GlAdL = False
End Sub

Private Sub txtS2F33_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F33_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F33.SelLength = 0
    Case vbKeyDown: Me.txtS2F35.SetFocus
    Case vbKeyUp: Me.txtS2F03.SetFocus
    End Select
End Sub
Private Sub txtS2F34_Change()

TagWe = Mid$(Me.txtS2F34.Tag, 2, Len(Me.txtS2F34.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F34.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F34_GotFocus()
    Me.txtS2F34.SelStart = 0
    Me.txtS2F34.SelLength = Len(Me.txtS2F34.Text)
    GlAdL = False
End Sub

Private Sub txtS2F34_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F34_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F34.SelLength = 0
    Case vbKeyDown: Me.txtS1F27.SetFocus
    Case vbKeyUp: Me.txtS1F19.SetFocus
    End Select
End Sub

Private Sub txtS3F03_Change()

TagWe = Mid$(Me.txtS3F03.Tag, 2, Len(Me.txtS3F03.Tag) - 1)

If GlAdL = False Then
    Me.txtS3F03.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub cmbS2F30_Click()

TagWe = Mid$(Me.cmbS2F30.Tag, 2, Len(Me.cmbS2F30.Tag) - 1)

If GlAdL = False Then
    Me.cmbS2F30.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub cmbS2F30_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbS2F30_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbS2F30.SelLength = 0
End Sub

Private Sub txtS4F01_Change()
    GlAdZ = True
End Sub

Private Sub txtS4F01_GotFocus()
    If GlAdL = False Then
        Me.txtS4F01.SelStart = 0
        Me.txtS4F01.SelLength = Len(Me.txtS4F01.Text)
    End If
End Sub

Private Sub txtS4F01_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F01_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F01.SelLength = 0
    Case vbKeyDown: Me.txtS4F02.SetFocus
    Case vbKeyUp:
    End Select
End Sub
Private Sub txtS4F02_Change()
    If GlAdL = False Then
        GlAdZ = True
        AdBri
    End If
End Sub
Private Sub txtS4F02_Click()
    If GlAdL = False Then
        GlAdZ = True
        AdBri
    End If
End Sub
Private Sub txtS4F02_GotFocus()
    GlAdL = False
End Sub
Private Sub txtS4F02_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS4F02_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F02.SelLength = 0
    Case vbKeyDown:
    Case vbKeyUp:
    End Select
End Sub

Private Sub txtS4F03_Change()
    If GlAdL = False Then
        GlAdZ = True
        AdBri
    End If
End Sub
Private Sub txtS4F03_GotFocus()
    Me.txtS4F03.SelStart = 0
    Me.txtS4F03.SelLength = Len(Me.txtS4F03.Text)
    GlAdL = False
End Sub
Private Sub txtS4F03_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F03_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F03.SelLength = 0
    Case vbKeyDown: Me.txtS4F04.SetFocus
    Case vbKeyUp: Me.txtS4F02.SetFocus
    End Select
End Sub
Private Sub txtS4F03_LostFocus()
    If GlAdL = False Then
        AdBri
    End If
End Sub
Private Sub txtS4F04_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F04_GotFocus()
    Me.txtS4F04.SelStart = 0
    Me.txtS4F04.SelLength = Len(Me.txtS4F04.Text)
    GlAdL = False
End Sub
Private Sub txtS4F04_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F04_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F04.SelLength = 0
    Case vbKeyDown: Me.txtS4F05.SetFocus
    Case vbKeyUp: Me.txtS4F03.SetFocus
    End Select
End Sub
Private Sub txtS4F04_LostFocus()
    If GlAdL = False Then
        AdBri
    End If
End Sub
Private Sub txtS4F05_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F05_GotFocus()
    Me.txtS4F05.SelStart = 0
    Me.txtS4F05.SelLength = Len(Me.txtS4F05.Text)
    GlAdL = False
End Sub
Private Sub txtS4F05_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F05_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F05.SelLength = 0
    Case vbKeyDown: Me.txtS4F06.SetFocus
    Case vbKeyUp: Me.txtS4F04.SetFocus
    End Select
End Sub
Private Sub txtS4F05_LostFocus()
    If GlAdL = False Then
        AdBri
    End If
End Sub
Private Sub txtS4F06_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F06_GotFocus()
    Me.txtS4F06.SelStart = 0
    Me.txtS4F06.SelLength = Len(Me.txtS4F06.Text)
    GlAdL = False
End Sub
Private Sub txtS4F06_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F06_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F06.SelLength = 0
    Case vbKeyDown: Me.txtS4F08.SetFocus
    Case vbKeyUp: Me.txtS4F05.SetFocus
    End Select
End Sub
Private Sub txtS4F08_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F08_GotFocus()
    Me.txtS4F08.SelStart = 0
    Me.txtS4F08.SelLength = Len(Me.txtS4F08.Text)
    GlAdL = False
End Sub
Private Sub txtS4F08_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F08_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F08.SelLength = 0
    Case vbKeyDown: Me.txtS4F09.SetFocus
    Case vbKeyUp: Me.txtS4F08.SetFocus
    End Select
End Sub

Private Sub txtS4F08_LostFocus()
    If GlAdL = False Then
        If Me.txtS4F09.Text = vbNullString Then
            Adr_Poz False
        End If
    End If
End Sub
Private Sub txtS4F09_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F09_Click()
    If Me.txtS4F08.Text <> vbNullString Then
        If Me.txtS4F09.Text = vbNullString Then
            Adr_Poz False
        End If
    End If
End Sub
Private Sub txtS4F09_GotFocus()
    Me.txtS4F09.SelStart = 0
    Me.txtS4F09.SelLength = Len(Me.txtS4F09.Text)
    GlAdL = False
End Sub
Private Sub txtS4F09_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F09_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F09.SelLength = 0
    Case vbKeyDown: Me.cmbS4F12.SetFocus
    Case vbKeyUp: Me.txtS4F08.SetFocus
    End Select
End Sub
Private Sub txtS4F15_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F15_GotFocus()
    Me.txtS4F15.SelLength = 0
    Me.txtS4F15.SelStart = Len(Me.txtS4F15.Text)
    GlAdL = False
End Sub
Private Sub txtS4F15_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F15_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F15.SelLength = 0
    Case vbKeyDown: Me.txtS4F16.SetFocus
    Case vbKeyUp: Me.txtS4F18.SetFocus
    End Select
End Sub
Private Sub txtS4F16_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F16_GotFocus()
    Me.txtS4F16.SelLength = 0
    Me.txtS4F16.SelStart = Len(Me.txtS4F16.Text)
    GlAdL = False
End Sub
Private Sub txtS4F16_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F16_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F16.SelLength = 0
    Case vbKeyDown: Me.txtS4F17.SetFocus
    Case vbKeyUp: Me.txtS4F16.SetFocus
    End Select
End Sub
Private Sub txtS4F17_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F18_Change()
    GlAdZ = True
End Sub
Private Sub txtS4F18_GotFocus()
    Me.txtS4F18.SelStart = Len(Me.txtS4F18.Text)
    Me.txtS4F18.SelLength = 0
    GlAdL = False
End Sub
Private Sub txtS4F18_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F18_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F18.SelLength = 0
    Case vbKeyDown: Me.txtS4F15.SetFocus
    Case vbKeyUp: Me.cmbS4F11.SetFocus
    End Select
End Sub


Private Sub txtS4F19_GotFocus()
    Me.txtS4F19.SelStart = 0
    Me.txtS4F19.SelLength = Len(Me.txtS4F19.Text)
End Sub
Private Sub txtS4F19_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
