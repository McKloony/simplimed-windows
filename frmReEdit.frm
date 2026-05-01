VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReEdit 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnungseigenschaften"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   36
      Top             =   8000
      Width           =   8400
      _Version        =   1048579
      _ExtentX        =   14817
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   6400
         TabIndex        =   40
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
      Begin XtremeSuiteControls.PushButton btnWieter 
         Default         =   -1  'True
         Height          =   400
         Left            =   5000
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Speichern"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton cmdReInh 
         Height          =   400
         Left            =   3600
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "Rechn.-&Inhalt"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   2300
         TabIndex        =   37
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
   Begin XtremeSuiteControls.ComboBox txtKomFe 
      Height          =   315
      Left            =   2820
      TabIndex        =   33
      Top             =   6795
      Width           =   4845
      _Version        =   1048579
      _ExtentX        =   8546
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.UpDown updCont1 
      Height          =   350
      Left            =   2220
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2600
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
      BuddyControl    =   "txtKopie"
      BuddyProperty   =   ""
   End
   Begin XtremeSuiteControls.FlatEdit txtReEmp 
      Height          =   350
      Left            =   2820
      TabIndex        =   2
      Top             =   500
      Width           =   2910
      _Version        =   1048579
      _ExtentX        =   5133
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtKopie 
      Height          =   350
      Left            =   700
      TabIndex        =   12
      Top             =   2600
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Text            =   "1"
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtBetra 
      Height          =   350
      Left            =   6060
      TabIndex        =   3
      Top             =   500
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnDatu2 
      Height          =   350
      Left            =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Öffnet den Auswahlkalender"
      Top             =   1200
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   617
      _ExtentY        =   617
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnDatu1 
      Height          =   350
      Left            =   2200
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Öffnet den Auswahlkalender"
      Top             =   1200
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   617
      _ExtentY        =   617
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtReNum 
      Height          =   350
      Left            =   700
      TabIndex        =   1
      Top             =   500
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   10000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbReTyp 
      Height          =   315
      Left            =   2820
      TabIndex        =   10
      Top             =   1900
      Width           =   2920
      _Version        =   1048579
      _ExtentX        =   5159
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbVersi 
      Height          =   310
      Left            =   2820
      TabIndex        =   14
      Top             =   2600
      Width           =   2920
      _Version        =   1048579
      _ExtentX        =   5159
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.ComboBox cmbReStu 
      Height          =   315
      Left            =   700
      TabIndex        =   9
      Top             =   1900
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "0"
   End
   Begin XtremeSuiteControls.ComboBox cmbZaZie 
      Height          =   310
      Left            =   2820
      TabIndex        =   17
      Top             =   3300
      Width           =   2920
      _Version        =   1048579
      _ExtentX        =   5159
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox4"
   End
   Begin XtremeSuiteControls.ComboBox cmbBeha2 
      Height          =   315
      Left            =   2820
      TabIndex        =   20
      Top             =   4000
      Width           =   2920
      _Version        =   1048579
      _ExtentX        =   5159
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox5"
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   0
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
      VisualTheme     =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtDatu2 
      Height          =   350
      Left            =   2820
      TabIndex        =   6
      Top             =   1200
      Width           =   1470
      _Version        =   1048579
      _ExtentX        =   2593
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDatu1 
      Height          =   350
      Left            =   700
      TabIndex        =   4
      Top             =   1200
      Width           =   1470
      _Version        =   1048579
      _ExtentX        =   2593
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtBezah 
      Height          =   350
      Left            =   6060
      TabIndex        =   11
      Top             =   1900
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtAnzah 
      Height          =   350
      Left            =   6060
      TabIndex        =   8
      Top             =   1200
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtExtra 
      Height          =   350
      Left            =   6060
      TabIndex        =   15
      Top             =   2600
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtTerNr 
      Height          =   350
      Left            =   6060
      TabIndex        =   18
      Top             =   3300
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.ComboBox cmbBeTyp 
      Height          =   315
      Left            =   700
      TabIndex        =   19
      Top             =   4000
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "0"
   End
   Begin XtremeSuiteControls.FlatEdit txtRabat 
      Height          =   350
      Left            =   700
      TabIndex        =   16
      Top             =   3300
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cmbMitar 
      Height          =   315
      Left            =   2820
      TabIndex        =   23
      Top             =   4700
      Width           =   2925
      _Version        =   1048579
      _ExtentX        =   5159
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox5"
   End
   Begin XtremeSuiteControls.FlatEdit txtRecId 
      Height          =   350
      Left            =   6060
      TabIndex        =   21
      Top             =   4000
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.ComboBox cmbWarun 
      Height          =   315
      Left            =   6060
      TabIndex        =   24
      Top             =   4700
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2831
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtGuStr 
      Height          =   350
      Left            =   6060
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtAuStr 
      Height          =   350
      Left            =   700
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtRzDat 
      Height          =   350
      Left            =   700
      TabIndex        =   28
      Top             =   6100
      Width           =   1470
      _Version        =   1048579
      _ExtentX        =   2593
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.PushButton btnDatu3 
      Height          =   350
      Left            =   2200
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Öffnet den Auswahlkalender"
      Top             =   6100
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   617
      _ExtentY        =   617
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbArzNr 
      Height          =   315
      Left            =   2820
      TabIndex        =   26
      Top             =   5400
      Width           =   2925
      _Version        =   1048579
      _ExtentX        =   5159
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox5"
   End
   Begin XtremeSuiteControls.FlatEdit txtRzNum 
      Height          =   350
      Left            =   6060
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6100
      Width           =   1605
      _Version        =   1048579
      _ExtentX        =   2831
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.ComboBox cmbBeGru 
      Height          =   315
      Left            =   700
      TabIndex        =   22
      Top             =   4700
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbGruTe 
      Height          =   315
      Left            =   700
      TabIndex        =   32
      Top             =   6800
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbVersa 
      Height          =   315
      Left            =   2820
      TabIndex        =   30
      Top             =   6100
      Width           =   2925
      _Version        =   1048579
      _ExtentX        =   5159
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.ComboBox cmbTheEn 
      Height          =   315
      Left            =   700
      TabIndex        =   34
      Top             =   7500
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox3"
   End
   Begin XtremeSuiteControls.FlatEdit txtRzTex 
      Height          =   350
      Left            =   2820
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7500
      Width           =   4840
      _Version        =   1048579
      _ExtentX        =   8537
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.Label lblLab33 
      Height          =   210
      Left            =   705
      TabIndex        =   72
      Top             =   7270
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Therapieende :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab32 
      Height          =   210
      Left            =   2828
      TabIndex        =   71
      Top             =   6570
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Rechnungsfreitext :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab31 
      Height          =   210
      Left            =   705
      TabIndex        =   70
      Top             =   6570
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Gruppentherapie :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab30 
      Height          =   210
      Left            =   705
      TabIndex        =   69
      Top             =   4470
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Behandlungsgrund :"
      Transparent     =   -1  'True
   End
   Begin VB.Label lblLab29 
      BackStyle       =   0  'Transparent
      Caption         =   "Verordnungtext :"
      Height          =   210
      Left            =   2828
      TabIndex        =   68
      Top             =   7270
      Width           =   1305
   End
   Begin VB.Label lblLab28 
      BackStyle       =   0  'Transparent
      Caption         =   "Verordnungsbeleg :"
      Height          =   210
      Left            =   6065
      TabIndex        =   67
      Top             =   5870
      Width           =   1500
   End
   Begin VB.Label lblLab27 
      BackStyle       =   0  'Transparent
      Caption         =   "Hausarzt / Verordner :"
      Height          =   210
      Left            =   2828
      TabIndex        =   66
      Top             =   5170
      Width           =   1700
   End
   Begin VB.Label lblLab26 
      BackStyle       =   0  'Transparent
      Caption         =   "Verordnungsdatum :"
      Height          =   210
      Left            =   705
      TabIndex        =   65
      Top             =   5870
      Width           =   1500
   End
   Begin VB.Label lblLab23 
      BackStyle       =   0  'Transparent
      Caption         =   "Auftragsnummer :"
      Height          =   210
      Left            =   705
      TabIndex        =   64
      Top             =   5170
      Width           =   1500
   End
   Begin VB.Label lblLab24 
      BackStyle       =   0  'Transparent
      Caption         =   "Gutschriftnummer :"
      Height          =   210
      Left            =   6065
      TabIndex        =   63
      Top             =   5170
      Width           =   1500
   End
   Begin XtremeSuiteControls.Label lblLab22 
      Height          =   210
      Left            =   6065
      TabIndex        =   62
      Top             =   4470
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Währung :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab19 
      Height          =   210
      Left            =   6065
      TabIndex        =   61
      Top             =   3770
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Rechnungszähler: "
      Transparent     =   -1  'True
   End
   Begin VB.Label lblLab18 
      BackStyle       =   0  'Transparent
      Caption         =   "Mitarbeiter :"
      Height          =   210
      Left            =   2825
      TabIndex        =   60
      Top             =   4470
      Width           =   1500
   End
   Begin VB.Label lblLab15 
      BackStyle       =   0  'Transparent
      Caption         =   "Mandant :"
      Height          =   210
      Left            =   2825
      TabIndex        =   59
      Top             =   3770
      Width           =   1500
   End
   Begin VB.Label lblLab17 
      BackStyle       =   0  'Transparent
      Caption         =   "Empfängertyp :"
      Height          =   210
      Left            =   705
      TabIndex        =   58
      Top             =   3770
      Width           =   1500
   End
   Begin XtremeSuiteControls.Label lblLab16 
      Height          =   210
      Left            =   6065
      TabIndex        =   57
      Top             =   3030
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Terminserie :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab13 
      Height          =   210
      Left            =   6065
      TabIndex        =   56
      Top             =   2370
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Extragebühr :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab10 
      Height          =   210
      Left            =   6065
      TabIndex        =   55
      Top             =   1670
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Bezahlt :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab07 
      Height          =   210
      Left            =   6065
      TabIndex        =   54
      Top             =   970
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Differenz :"
      Transparent     =   -1  'True
   End
   Begin VB.Label lblLab01 
      BackStyle       =   0  'Transparent
      Caption         =   "Belegnummer :"
      Height          =   210
      Left            =   705
      TabIndex        =   52
      Top             =   270
      Width           =   1200
   End
   Begin VB.Label lblLab02 
      BackStyle       =   0  'Transparent
      Caption         =   "Belegdatum :"
      Height          =   210
      Left            =   705
      TabIndex        =   51
      Top             =   970
      Width           =   1500
   End
   Begin VB.Label lblLab03 
      BackStyle       =   0  'Transparent
      Caption         =   "Fälligkeit :"
      Height          =   210
      Left            =   2825
      TabIndex        =   50
      Top             =   970
      Width           =   1500
   End
   Begin VB.Label lblLab04 
      BackStyle       =   0  'Transparent
      Caption         =   "Betrag :"
      Height          =   210
      Left            =   6065
      TabIndex        =   49
      Top             =   270
      Width           =   1500
   End
   Begin VB.Label lblLab06 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient :"
      Height          =   210
      Left            =   2825
      TabIndex        =   48
      Top             =   270
      Width           =   900
   End
   Begin VB.Label lblLab09 
      BackStyle       =   0  'Transparent
      Caption         =   "Katalog :"
      Height          =   210
      Left            =   2825
      TabIndex        =   47
      Top             =   2370
      Width           =   1500
   End
   Begin VB.Label lblLab12 
      BackStyle       =   0  'Transparent
      Caption         =   "Zahlungsw.:"
      Height          =   210
      Left            =   2825
      TabIndex        =   46
      Top             =   3030
      Width           =   1500
   End
   Begin VB.Label lblLab05 
      BackStyle       =   0  'Transparent
      Caption         =   "Belegtyp :"
      Height          =   210
      Left            =   2825
      TabIndex        =   45
      Top             =   1670
      Width           =   1500
   End
   Begin VB.Label lblLab25 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnungsversand :"
      Height          =   210
      Left            =   2828
      TabIndex        =   44
      Top             =   5870
      Width           =   1600
   End
   Begin VB.Label lblLab11 
      BackStyle       =   0  'Transparent
      Caption         =   "Ausdrucke :"
      Height          =   210
      Left            =   705
      TabIndex        =   43
      Top             =   2370
      Width           =   1500
   End
   Begin VB.Label lblLab08 
      BackStyle       =   0  'Transparent
      Caption         =   "Steuersatz :"
      Height          =   210
      Left            =   705
      TabIndex        =   42
      Top             =   1670
      Width           =   1500
   End
   Begin VB.Label lblLab14 
      BackStyle       =   0  'Transparent
      Caption         =   "Minderung (%) :"
      Height          =   210
      Left            =   705
      TabIndex        =   41
      Top             =   3030
      Width           =   1500
   End
End
Attribute VB_Name = "frmReEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private CmVrs As XtremeSuiteControls.ComboBox
Private CmZil As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmReS As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmArz As XtremeSuiteControls.ComboBox
Private CmKom As XtremeSuiteControls.ComboBox
Private CmBTy As XtremeSuiteControls.ComboBox
Private CmVeA As XtremeSuiteControls.ComboBox
Private CmTEn As XtremeSuiteControls.ComboBox
Private CmWar As XtremeSuiteControls.ComboBox
Private CmGrn As XtremeSuiteControls.ComboBox
Private CmGru As XtremeSuiteControls.ComboBox
Private CmVer As XtremeSuiteControls.ComboBox
Private CmEnd As XtremeSuiteControls.ComboBox
Private TxRen As XtremeSuiteControls.FlatEdit
Private TxAuf As XtremeSuiteControls.FlatEdit
Private TxRzn As XtremeSuiteControls.FlatEdit
Private TxRzt As XtremeSuiteControls.FlatEdit
Private TxKop As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private MoKal As XtremeCalendarControl.DatePicker
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow

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

Private ReDat As Date
Private KalWa As Integer
Private ReAnp As Boolean
Private FoLad As Boolean

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private clFen As clsFenster
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date
Dim ReStr As String

Set TxRen = Me.txtReNum
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtRzDat
Set MoKal = Me.dtpDatu1

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
End If

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
        If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
            If ReStr <> "-" Then
                If Year(NeuDa) < Year(Date) Then
                    NeuDa = ReDat
                    SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
                End If
            End If
        End If
        TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
        If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
            If ReStr <> "-" Then
                If Year(NeuDa) < Year(Date) Then
                    NeuDa = ReDat
                    SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
                End If
            End If
        End If
        TxDa2.Text = Format$(NeuDa, "dd.mm.yyyy")
    End If
Case 3:
    If IsDate(TxDa3.Text) Then
        NeuDa = TxDa3.Text
        TxDa3.Text = Format$(NeuDa, "dd.mm.yyyy")
    End If
End Select

With MoKal
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

If NeuDa > Date Then
    SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date
Dim ReStr As String

Set TxRen = Me.txtReNum
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
End If

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
        If ReStr <> "-" Then
            If Year(NeuDa) < Year(Date) Then
                NeuDa = ReDat
                SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
            End If
        End If
    End If
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
Private Sub FInit()
On Error GoTo SuErr

Dim AktZa As Integer
Dim GesZa As Integer

Set FM = frmReEdit
Set Rahm0 = FM.frmRahm0
Set TxRen = FM.txtReNum
Set TxKop = FM.txtKopie
Set TxRzn = FM.txtRzNum
Set CmKom = FM.txtKomFe
Set CmVrs = FM.cmbVersi
Set CmZil = FM.cmbZaZie
Set CmTyp = FM.cmbReTyp
Set CmReS = FM.cmbReStu
Set CmMan = FM.cmbBeha2
Set CmMit = FM.cmbMitar
Set CmArz = FM.cmbArzNr
Set CmBTy = FM.cmbBeTyp
Set CmWar = FM.cmbWarun
Set CmGrn = FM.cmbBeGru
Set CmGru = FM.cmbGruTe
Set CmVer = FM.cmbVersa
Set CmEnd = FM.cmbTheEn
Set MoKal = FM.dtpDatu1
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnDatu2
Set PuBu3 = FM.btnDatu3
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set TxDa3 = FM.txtRzDat
Set ImMan = frmMain.imgManag

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

For AktZa = 1 To UBound(GlGKa)
    CmVrs.AddItem GlGKa(AktZa, 1)
    CmVrs.ItemData(AktZa - 1) = GlGKa(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlZah)
    CmZil.AddItem GlZah(AktZa, 1)
    CmZil.ItemData(AktZa - 1) = GlZah(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlStu)
    CmReS.AddItem GlStu(AktZa, 2)
    CmReS.ItemData(AktZa - 1) = GlStu(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlReK)
    CmKom.AddItem GlReK(AktZa, 1)
    CmKom.ItemData(AktZa - 1) = GlReK(AktZa, 0)
Next AktZa

With CmTyp
    .AddItem "R - Standardrechnung"
    .ItemData(0) = 1
    .AddItem "V - Kostenvoranschlag"
    .ItemData(1) = 2
    .AddItem "L - Laborrechnung"
    .ItemData(2) = 3
    .AddItem "A - Abrechnungsstelle"
    .ItemData(3) = 4
    .AddItem "U - Gutschrift"
    .ItemData(4) = 5
    .AddItem "M - Rechnungsauftrag"
    .ItemData(5) = 6
    .AddItem "G - Gewerberechnung"
    .ItemData(6) = 7
    .AddItem "I - Importrechnung"
    .ItemData(7) = 8
End With

With CmBTy
    .AddItem "Privat Inland"
    .ItemData(0) = 1
    .AddItem "Privat Europa"
    .ItemData(1) = 2
    .AddItem "Privat Ausland"
    .ItemData(2) = 3
    .AddItem "Gewerb. Inland"
    .ItemData(3) = 4
    .AddItem "Gewerb. Europa"
    .ItemData(4) = 5
    .AddItem "Gewerb. Ausland"
    .ItemData(5) = 6
    .ListIndex = 1
End With

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
    CmMit.AddItem GlMiK(AktZa, 1)
    CmMit.ItemData(AktZa - 1) = GlMiK(AktZa, 2)
Next AktZa

If GlArV = True Then
   For AktZa = 1 To UBound(GlArz) 'Verordner
        CmArz.AddItem GlArz(AktZa, 8)
        CmArz.ItemData(AktZa - 1) = GlArz(AktZa, 0)
    Next AktZa
End If

For AktZa = 1 To UBound(GlWar)
    CmWar.AddItem GlWar(AktZa, 1)
    CmWar.ItemData(AktZa - 1) = GlWar(AktZa, 0)
Next AktZa

With CmGru
    .AddItem "Einzeltherapie"
    .ItemData(0) = 1
    .AddItem "Gruppentherapie"
    .ItemData(1) = 2
    .ListIndex = 0
End With

With CmVer
    .AddItem "Postversand"
    .ItemData(0) = 0
    .AddItem "Emailversand"
    .ItemData(1) = 1
    .AddItem "Downloadlink"
    .ItemData(2) = 2
    .ListIndex = GlRVs - 1 'Standard-Rechnungsversandweg
End With

With CmEnd
    .AddItem "Andauerend"
    .ItemData(0) = 1
    .AddItem "Beendet"
    .ItemData(1) = 2
    .ListIndex = 0
End With

If GlMaV = False Then 'Mandanten vorhanden
    CmMan.Enabled = False
Else
    If UBound(GlMan) <= 1 Then
        CmMan.Enabled = False
    End If
End If

For AktZa = 0 To UBound(GlThG)
    With CmGrn
        .AddItem GlThG(AktZa, 0)
        .ItemData(AktZa) = GlThG(AktZa, 1)
    End With
Next AktZa

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

With TxKop
    .Pattern = "\d*"
    .SetMask "00", "__"
End With

With TxRzn
    .Pattern = "\d*"
    .SetMask "000000", "__________"
End With

Select Case GlRFm 'Rechnungsnummernformat
Case 2: TxRen.SetMask "00-00-000000", "__-__-______"
Case 3: TxRen.SetMask "0000-000000", "____-______"
Case 4: TxRen.SetMask "00-000000", "__-______"
Case 5: TxRen.SetMask "0000-0000", "____-____"
End Select

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu3.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

If GlSpl = True Then 'Steuerspalte
    CmReS.Enabled = False
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub

Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date
Dim Datu3 As Date
Dim ReStr As String
Dim ZaZil As Integer
Dim ZiZal As Integer
Dim AktZa As Integer

Set TxRen = Me.txtReNum
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtRzDat
Set CmZil = FM.cmbZaZie
Set MoKal = Me.dtpDatu1

If CmZil.ListIndex >= 0 Then
    ZaZil = CmZil.ItemData(CmZil.ListIndex)
    For AktZa = 1 To UBound(GlZah)
        If GlZah(AktZa, 0) = ZaZil Then
            ZiZal = GlZah(AktZa, 2)
            Exit For
        End If
    Next AktZa
End If

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
End If

If ZiZal = 0 Then
    ZiZal = 1
End If

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
                    If IsDate(TxDa1.Text) Then
                        NeuDa = .Selection.Blocks(0).DateBegin
                        If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
                            If ReStr <> "-" Then
                                If Year(NeuDa) < Year(Date) Then
                                    NeuDa = ReDat
                                    SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
                                End If
                            End If
                        End If
                        TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
                    End If
                End If
            End If
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    If IsDate(TxDa2.Text) Then
                        NeuDa = .Selection.Blocks(0).DateBegin
                        If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
                            If ReStr <> "-" Then
                                If Year(NeuDa) < Year(Date) Then
                                    NeuDa = ReDat
                                    SPopu "Rechnungsdatum ist zu alt", "Das Rechnungsdatum darf sich nicht auf abgeschlossene Geschäftsjahre beziehen.", IC48_Information
                                End If
                            End If
                        End If
                        TxDa2.Text = Format$(NeuDa, "dd.mm.yyyy")
                    End If
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

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        Datu1 = TxDa1.Text
    End If
End If
If TxDa2.Text <> vbNullString Then
    If IsDate(TxDa2.Text) = True Then
        Datu2 = TxDa2.Text
    End If
End If
If TxDa3.Text <> vbNullString Then
    If IsDate(TxDa3.Text) = True Then
        Datu3 = TxDa3.Text
    End If
End If

If Datu2 < Datu1 Then
    TxDa1.Text = Datu2
End If

If Datu3 > Datu1 Then
    TxDa3.Text = Datu1
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FLoad()
On Error GoTo SuErr

Set FM = frmReEdit
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set TxDa3 = FM.txtRzDat
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnDatu2
Set PuBu3 = FM.btnDatu3

ReDat = TxDa1.Text

If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
    If IsDate(TxDa1.Text) = True Then
        If Year(ReDat) < Year(Date) Then
            TxDa1.Enabled = False
            PuBu1.Enabled = False
        End If
    End If
    If IsDate(TxDa2.Text) = True Then
        If Year(ReDat) < Year(Date) Then
            TxDa2.Enabled = False
            PuBu2.Enabled = False
        End If
    End If
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FReNu()
On Error GoTo SuErr

Dim Mld1, Mld2 As String
Dim Mld3, Mld4 As String

Set FM = frmReEdit

Mld1 = "Rechnungsnummer  wurde geändert"
Mld2 = "Soll jetzt eine Reindizierung der Rechnungsnummern durchgeführt werden?"
Mld3 = "Sie haben eine Änderung an der automatisch generierten Rechnungsnummer vorgenommen. Damit auf Basis dieser neunen Rechnungsnummer auch zukünftig eine Rechnungsnummer generiert werden kann, muss eine Reindizierung der Rechnungsnummern vorgenommen werden. Dieses kann einige Minuten in Anspruch nehmen."
Mld4 = "HINWEIS! Änderungen an der automatisch vorgeschlagenen Rechnungsnummer werden protokolliert und können über das Tagesprotokoll eingesehen werden."

If GlReN = True Then 'Rechnungsnummern sofort erzeugen
    If ReAnp = True Then 'Rechnungsnummer Reindizieren
        SMeFr Mld1, Mld2, Mld3, Mld4, False, 0, , FM.hwnd
        If GlMes = 33565 Then
            S_ReNm
            DoEvents
        End If
    End If
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FReNu " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FTyPr(Optional ByVal ManPr As Boolean = False)
On Error GoTo LaErr
'Kontrolliert den Belegtyp

Dim NeuDa As Date
Dim AnzRe As Long
Dim ReNum As Long
Dim MaNum As Long
Dim MiNum As Long
Dim ReStr As String
Dim AuStr As String
Dim ReTyp As String
Dim TyStr As String
Dim ReAbg As Boolean
Dim TypNr As Integer
Dim TyNum As Integer
Dim Mld1, Tit1 As String
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set CmTyp = Me.cmbReTyp
Set TxRen = Me.txtReNum
Set TxAuf = Me.txtAuStr
Set TxDa1 = Me.txtDatu1
Set CmMan = Me.cmbBeha2
Set CmMit = Me.cmbMitar

TypNr = CmTyp.ListIndex 'Neuer Belegtyp
If TxRen.Text <> vbNullString Then
    AuStr = TxRen.Text
Else
    AuStr = "-"
End If
MaNum = CmMan.ItemData(CmMan.ListIndex)
MiNum = CmMit.ItemData(CmMit.ListIndex)

Tit1 = "Belegtyp Ändern"
Mld1 = "Das Ändern des von Ihnen gewünschten Belegtyps ist nicht möglich"

Select Case TypNr 'Neuer Belegtyp
Case 0: ReTyp = "R"
Case 1: ReTyp = "V"
Case 2: ReTyp = "L"
Case 3: ReTyp = "A"
Case 4: ReTyp = "U"
Case 5: ReTyp = "M"
Case 6: ReTyp = "G"
Case 7: ReTyp = "I"
End Select

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        NeuDa = CDate(TxDa1.Text)
    Else
        NeuDa = Date
    End If
Else
    NeuDa = Date
End If

Select Case GlBut
Case RibTab_Startseite:
            Set RpCls = RpCo4.Columns
            Set RpSel = RpCo4.SelectedRows
            AnzRe = RpSel.Count
Case RibTab_Abrechnung:
            Set RpCls = RpCo3.Columns
            Set RpSel = RpCo3.SelectedRows
            AnzRe = RpSel.Count
Case RibTab_Rechnungen:
            Set RpCls = RpCo4.Columns
            Set RpSel = RpCo4.SelectedRows
            AnzRe = RpSel.Count
End Select

If AnzRe > 0 Then
    Set RpRow = RpSel(0)
    Set RpCol = RpCls.Find(Rec_ID1)
    ReNum = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_RechNr)
    ReStr = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_Type)
    TyStr = RpRow.Record(RpCol.ItemIndex).Value
    Select Case GlBut
    Case RibTab_Abrechnung:
        Set RpCol = RpCls.Find(Rec_Selekt)
        ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
    Case RibTab_Rechnungen:
        Set RpCol = RpCls.Find(Rec_Selekt)
        If LCase(RpRow.Record(RpCol.ItemIndex).Value) = "ja" Then
            ReAbg = True
        Else
            ReAbg = False
        End If
    End Select
    
    Select Case TyStr 'Alter Belegtyp
    Case "R": TyNum = 0
    Case "V": TyNum = 1
    Case "L": TyNum = 2
    Case "A": TyNum = 3
    Case "U": TyNum = 4
    Case "M": TyNum = 5
    Case "G": TyNum = 6
    Case "I": TyNum = 7
    End Select

    If ManPr = False Then
        Select Case ReTyp 'neuer Belegtyp
        Case "V":
                CmTyp.ListIndex = TyNum
                SPopu Tit1, Mld1, IC48_Lock
        Case "U":
                CmTyp.ListIndex = TyNum
                SPopu Tit1, Mld1, IC48_Lock
        Case "M":
                CmTyp.ListIndex = TyNum
                SPopu Tit1, Mld1, IC48_Lock
        Case Else:
            Select Case TyStr 'alter Belegtyp
            Case "U": 'Gutschrift
                CmTyp.ListIndex = TyNum
                SPopu Tit1, Mld1, IC48_Lock
            Case "V":
                TxAuf.Text = AuStr
                TxRen.Text = S_ReVo(NeuDa, ReTyp, MaNum, MiNum, True)
            Case "M":
                TxAuf.Text = AuStr
                TxRen.Text = S_ReVo(NeuDa, ReTyp, MaNum, MiNum, True)
            Case Else:
                TxAuf.Text = vbNullString
                TxRen.Text = ReStr
            End Select
        End Select
    End If
End If

Set RpSel = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTyPr " & Err.Number
Resume Next

End Sub
Private Sub FZaZi()
On Error GoTo LaErr

Dim NeuDa As Date
Dim ZaZil As Integer
Dim ZiZal As Integer
Dim AktZa As Integer

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set CmZil = Me.cmbZaZie

 If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

If CmZil.ListIndex >= 0 Then
    ZaZil = CmZil.ItemData(CmZil.ListIndex)
    For AktZa = 1 To UBound(GlZah)
        If GlZah(AktZa, 0) = ZaZil Then
            ZiZal = GlZah(AktZa, 2)
            Exit For
        End If
    Next AktZa
End If

If ZiZal = 0 Then ZiZal = 1

TxDa2.Text = NeuDa + ZiZal

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZaZi " & Err.Number
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
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50801)
TeMai = IniGetOpt("Hilfe", 50802)
TeInh = IniGetOpt("Hilfe", 50803)
TeFus = IniGetOpt("Hilfe", 50804)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWieter_Click()
On Error Resume Next

Dim NeuDa As Date
Dim ManNr As Long
Dim ReStr As String
Dim ReTyp As String
Dim TypNr As Integer
Dim Lange As Integer
Dim Mld1, Mld2, Tit1 As String

Set FM = frmReEdit
Set TxDa1 = Me.txtDatu1
Set TxRen = FM.txtReNum
Set CmTyp = Me.cmbReTyp
Set CmMan = FM.cmbBeha2

Select Case GlRFm 'Rechnungsnummernformat
Case 2: Lange = 12
Case 3: Lange = 11
Case 4: Lange = 9
Case 5: Lange = 9
End Select

TypNr = CmTyp.ListIndex
ManNr = CmMan.ItemData(CmMan.ListIndex)

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
End If

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        NeuDa = CDate(TxDa1.Text)
    Else
        NeuDa = Date
    End If
Else
    NeuDa = Date
End If

Select Case TypNr
Case 0: ReTyp = "R"
Case 1: ReTyp = "V"
Case 2: ReTyp = "L"
Case 3: ReTyp = "A"
Case 4: ReTyp = "U"
Case 5: ReTyp = "M"
Case 6: ReTyp = "G"
Case 7: ReTyp = "I"
End Select

Tit1 = "Falsches Rechnungsnummernformat"
Mld1 = "Die von Ihnen eingestellte Rechnungsnummer hat das falsche Format"
Mld2 = "Die von Ihnen eingestellte Rechnungsnummer existiert bereits"

If GlReN = True Then 'Rechnungsnummern sofort erzeugen
    If ReStr = "-" Then
        S_Save
        Unload Me
    ElseIf InStrRev(ReStr, "_", -1, 1) > 0 Then
        SPopu Tit1, Mld1, IC48_Warning
        Exit Sub
    ElseIf Len(ReStr) <> Lange Then
        SPopu Tit1, Mld1, IC48_Warning
        Exit Sub
    Else
        If ReAnp = True Then
            If ReStr <> vbNullString Then
                If S_ReVr(ReStr, ReTyp, ManNr, NeuDa) = True Then
                    SPopu Tit1, Mld2, IC48_Warning
                    Exit Sub
                Else
                    S_Save
                    FReNu
                    Unload Me
                End If
            Else
                SPopu Tit1, Mld1, IC48_Warning
                Exit Sub
            End If
        Else
            S_Save
            Unload Me
        End If
    End If
Else
    S_Save
    Unload Me
End If

S_ReKon
DoEvents

If WindowLoad("frmAufga") = True Then
    If GlWaT = RibTab_Wart_Beha Then
        WaSpl RibTab_Wart_Beha
        S_WaLa RibTab_Wart_Beha
    End If
End If

DoEvents
SAnza

End Sub

Private Sub cmbArzNr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbArzNr_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.txtRzNum.SetFocus
    Case vbKeyUp: 'Me.txtRzDat.SetFocus
    End Select
End Sub


Private Sub cmbBeha2_Click()
    If FoLad = False Then
        FTyPr True
    End If
End Sub

Private Sub cmbBeha2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBeha2_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.txtKomFe.SetFocus
    Case vbKeyUp: 'Me.txtKopie.SetFocus
    End Select
End Sub

Private Sub cmbBeTyp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbBeTyp_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.cmbVeArt.SetFocus
    Case vbKeyUp: 'Me.txtKomFe.SetFocus
    End Select
End Sub

Private Sub cmbReStu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbReStu_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.cmbZaZie.SetFocus
    Case vbKeyUp: 'Me.cmbVersi.SetFocus
    End Select
End Sub

Private Sub cmbReTyp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbReTyp_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.cmbVersi.SetFocus
    Case vbKeyUp: 'Me.txtReEmp.SetFocus
    End Select
End Sub

Private Sub cmbVersi_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbVersi_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.cmbReStu.SetFocus
    Case vbKeyUp: 'Me.cmbReTyp.SetFocus
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
    Case vbKeyDown: 'Me.txtKomFe.SetFocus
    Case vbKeyUp: 'Me.cmbThEnd.SetFocus
    End Select
End Sub

Private Sub cmbZaZie_Click()
    If FoLad = False Then
        FZaZi
    End If
End Sub

Private Sub cmbZaZie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbZaZie_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.txtKopie.SetFocus
    Case vbKeyUp: 'Me.cmbReStu.SetFocus
    End Select
End Sub

Private Sub cmdReInh_Click()
    SReZe
    Unload Me
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

FInit
S_Posi
FLoad

clFen.FenVor

Set clFen = Nothing

FoLad = False

AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmReEdit = Nothing
End Sub

Private Sub txtAnzah_GotFocus()
    Me.txtAnzah.SelStart = 0
    Me.txtAnzah.SelLength = Len(Me.txtAnzah.Text)
End Sub

Private Sub txtAnzah_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtAnzah_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtAnzah.SelLength = 0
    Case vbKeyDown: 'Me.txtBezah.SetFocus
    Case vbKeyUp: 'Me.txtKomFe.SetFocus
    End Select
End Sub

Private Sub txtBezah_GotFocus()
    Me.txtBezah.SelStart = 0
    Me.txtBezah.SelLength = Len(Me.txtBezah.Text)
End Sub

Private Sub txtBezah_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtBezah_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBezah.SelLength = 0
    Case vbKeyDown: 'Me.txtReNum.SetFocus
    Case vbKeyUp: 'Me.txtAnzah.SetFocus
    End Select
End Sub
Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
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
    Case vbKeyDown: 'Me.txtDatu2.SetFocus
    Case vbKeyUp: 'Me.txtReNum.SetFocus
    End Select
End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub
Private Sub txtDatu2_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu2.SelLength = Len(Me.txtDatu2.Text)
End Sub
Private Sub txtDatu2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtDatu2_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtDatu2.SelLength = 0
    Case vbKeyDown: 'Me.txtBetra.SetFocus
    Case vbKeyUp: 'Me.txtDatu1.SetFocus
    End Select
End Sub
Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
End Sub

Private Sub txtExtra_GotFocus()
    Me.txtExtra.SelStart = 0
    Me.txtExtra.SelLength = Len(Me.txtExtra.Text)
End Sub
Private Sub txtExtra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtGuStr_GotFocus()
    Me.txtGuStr.SelStart = 0
    Me.txtGuStr.SelLength = Len(Me.txtGuStr.Text)
End Sub


Private Sub txtGuStr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtRabat_GotFocus()
    Me.txtRabat.SelStart = 0
    Me.txtRabat.SelLength = Len(Me.txtRabat.Text)
End Sub
Private Sub txtRabat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRabat_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtRabat.SelLength = 0
    Case vbKeyDown: 'Me.cmbBeha2.SetFocus
    Case vbKeyUp: 'Me.txtExtra.SetFocus
    End Select
End Sub

Private Sub txtRecId_GotFocus()
    Me.txtRecId.SelStart = 0
    Me.txtRecId.SelLength = Len(Me.txtRecId.Text)
End Sub
Private Sub txtRecId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtReNum_KeyDown(KeyCode As Integer, Shift As Integer)
    ReAnp = True
End Sub
Private Sub txtReNum_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtReNum.SelLength = 0
    Case vbKeyDown: 'Me.txtDatu1.SetFocus
    Case vbKeyUp: 'Me.txtKomFe.SetFocus
    End Select
End Sub
Private Sub txtBetra_GotFocus()
    Me.txtBetra.SelStart = 0
    Me.txtBetra.SelLength = Len(Me.txtBetra.Text)
End Sub
Private Sub txtBetra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBetra_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBetra.SelLength = 0
    Case vbKeyDown: 'Me.txtReEmp.SetFocus
    Case vbKeyUp: 'Me.txtBetra.SetFocus
    End Select
End Sub
Private Sub txtKomFe_GotFocus()
    Me.txtKomFe.SelStart = 0
    Me.txtKomFe.SelLength = Len(Me.txtKomFe.Text)
End Sub
Private Sub txtKomFe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKomFe_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtKomFe.SelLength = 0
    Case vbKeyDown: 'Me.cmbBeTyp.SetFocus
    Case vbKeyUp: 'Me.txtRzTex.SetFocus
    End Select
End Sub
Private Sub txtKopie_GotFocus()
    Me.txtKopie.SelStart = 0
    Me.txtKopie.SelLength = Len(Me.txtKopie.Text)
End Sub
Private Sub txtKopie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKopie_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtKopie.SelLength = 0
    Case vbKeyDown: 'Me.cmbBeha2.SetFocus
    Case vbKeyUp: 'Me.cmbZaZie.SetFocus
    End Select
End Sub
Private Sub txtReEmp_GotFocus()
    Me.txtReEmp.SelStart = 0
    Me.txtReEmp.SelLength = Len(Me.txtReEmp.Text)
End Sub
Private Sub txtReEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtReEmp_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtReEmp.SelLength = 0
    Case vbKeyDown: 'Me.cmbReTyp.SetFocus
    Case vbKeyUp: 'Me.txtBetra.SetFocus
    End Select
End Sub
Private Sub txtReNum_GotFocus()
    Me.txtReNum.SelStart = 0
    Me.txtReNum.SelLength = Len(Me.txtReNum.Text)
End Sub
Private Sub txtReNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRzDat_GotFocus()
    Me.txtRzDat.SelStart = 0
    Me.txtRzDat.SelLength = Len(Me.txtRzDat.Text)
End Sub

Private Sub txtRzDat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRzDat_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtRzDat.SelLength = 0
    Case vbKeyDown: 'Me.cmbArzNr.SetFocus
    Case vbKeyUp: 'Me.txtGuStr.SetFocus
    End Select
End Sub

Private Sub txtRzDat_LostFocus()
    KalWa = 3
    FDaKo
End Sub

Private Sub txtRzNum_GotFocus()
    Me.txtRzNum.SelStart = 0
    Me.txtRzNum.SelLength = Len(Me.txtRzNum.Text)
End Sub

Private Sub txtRzNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRzNum_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtRzNum.SelLength = 0
    Case vbKeyDown: 'Me.txtRzTex.SetFocus
    Case vbKeyUp: 'Me.cmbArzNr.SetFocus
    End Select
End Sub

Private Sub txtRzTex_GotFocus()
    Me.txtRzTex.SelStart = 0
    Me.txtRzTex.SelLength = Len(Me.txtRzTex.Text)
End Sub

Private Sub txtRzTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRzTex_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtRzTex.SelLength = 0
    Case vbKeyDown: 'Me.txtKomFe.SetFocus
    Case vbKeyUp: 'Me.txtRzNum.SetFocus
    End Select
End Sub
Private Sub txtTerNr_GotFocus()
    Me.txtTerNr.SelStart = 0
    Me.txtTerNr.SelLength = Len(Me.txtTerNr.Text)
End Sub
Private Sub txtTerNr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbReTyp_Click()
    If FoLad = False Then
        FTyPr
    End If
End Sub

