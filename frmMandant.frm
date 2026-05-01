VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmMandant 
   Caption         =   "Stammdaten"
   ClientHeight    =   12090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17880
   ControlBox      =   0   'False
   Icon            =   "frmMandant.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12090
   ScaleWidth      =   17880
   Begin XtremeSuiteControls.GroupBox frmRahm7 
      Height          =   3800
      Left            =   11000
      TabIndex        =   7
      Top             =   8400
      Visible         =   0   'False
      Width           =   11000
      _Version        =   1048579
      _ExtentX        =   19403
      _ExtentY        =   6703
      _StockProps     =   79
      Caption         =   "Sprechzeiten"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cmbNotVa 
         Height          =   315
         Left            =   8940
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   740
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
      Begin XtremeSuiteControls.CheckBox chkDefra 
         Height          =   255
         Left            =   7800
         TabIndex        =   96
         TabStop         =   0   'False
         Tag             =   "0Versand"
         Top             =   3040
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Multi-Terminbetreff Auswahl"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKaAus 
         Height          =   255
         Left            =   7800
         TabIndex        =   98
         TabStop         =   0   'False
         Tag             =   "0Gesperrt"
         Top             =   3040
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Kalenderspalte ausblenden"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox14 
         Height          =   220
         Left            =   4860
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   3040
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox13 
         Height          =   220
         Left            =   4860
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2580
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox12 
         Height          =   220
         Left            =   4860
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2140
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox11 
         Height          =   220
         Left            =   4860
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1680
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox10 
         Height          =   220
         Left            =   4860
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1220
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox09 
         Height          =   220
         Left            =   4860
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   780
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox08 
         Height          =   220
         Left            =   4860
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   340
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox07 
         Height          =   220
         Left            =   1440
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   3040
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox06 
         Height          =   220
         Left            =   1440
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2580
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox05 
         Height          =   220
         Left            =   1440
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2140
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox04 
         Height          =   220
         Left            =   1440
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1680
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox03 
         Height          =   220
         Left            =   1440
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1220
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox02 
         Height          =   220
         Left            =   1440
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   780
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBox01 
         Height          =   220
         Left            =   1440
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   340
         Width           =   220
         _Version        =   1048579
         _ExtentX        =   388
         _ExtentY        =   388
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ01 
         Height          =   320
         Left            =   1720
         TabIndex        =   48
         Top             =   300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ05 
         Height          =   320
         Left            =   1720
         TabIndex        =   54
         Top             =   740
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ09 
         Height          =   320
         Left            =   1720
         TabIndex        =   60
         Top             =   1180
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ13 
         Height          =   320
         Left            =   1720
         TabIndex        =   66
         Top             =   1640
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ17 
         Height          =   320
         Left            =   1720
         TabIndex        =   72
         Top             =   2080
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox5"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ21 
         Height          =   320
         Left            =   1720
         TabIndex        =   78
         Top             =   2540
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox6"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ25 
         Height          =   320
         Left            =   1720
         TabIndex        =   84
         Top             =   3000
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox7"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ02 
         Height          =   320
         Left            =   3100
         TabIndex        =   49
         Top             =   300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox8"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ06 
         Height          =   320
         Left            =   3100
         TabIndex        =   55
         Top             =   740
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox9"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ10 
         Height          =   320
         Left            =   3100
         TabIndex        =   61
         Top             =   1180
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox10"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ14 
         Height          =   320
         Left            =   3100
         TabIndex        =   67
         Top             =   1640
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox11"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ18 
         Height          =   320
         Left            =   3100
         TabIndex        =   73
         Top             =   2080
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox12"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ22 
         Height          =   315
         Left            =   3100
         TabIndex        =   79
         Top             =   2520
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox13"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ26 
         Height          =   320
         Left            =   3100
         TabIndex        =   85
         Top             =   3000
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox14"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ03 
         Height          =   320
         Left            =   5150
         TabIndex        =   51
         Top             =   300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox15"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ07 
         Height          =   320
         Left            =   5150
         TabIndex        =   57
         Top             =   740
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox16"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ11 
         Height          =   320
         Left            =   5150
         TabIndex        =   63
         Top             =   1180
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox17"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ15 
         Height          =   320
         Left            =   5150
         TabIndex        =   69
         Top             =   1640
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox18"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ19 
         Height          =   320
         Left            =   5150
         TabIndex        =   75
         Top             =   2080
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox19"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ23 
         Height          =   320
         Left            =   5150
         TabIndex        =   81
         Top             =   2540
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox20"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ27 
         Height          =   320
         Left            =   5150
         TabIndex        =   87
         Top             =   3000
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox21"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ04 
         Height          =   320
         Left            =   6510
         TabIndex        =   52
         Top             =   300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox22"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ08 
         Height          =   320
         Left            =   6510
         TabIndex        =   58
         Top             =   740
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox23"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ12 
         Height          =   320
         Left            =   6510
         TabIndex        =   64
         Top             =   1180
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox24"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ16 
         Height          =   320
         Left            =   6510
         TabIndex        =   70
         Top             =   1640
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox25"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ20 
         Height          =   315
         Left            =   6510
         TabIndex        =   76
         Top             =   2080
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox26"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ24 
         Height          =   320
         Left            =   6510
         TabIndex        =   82
         Top             =   2540
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox27"
      End
      Begin XtremeSuiteControls.ComboBox cmbSpZ28 
         Height          =   320
         Left            =   6510
         TabIndex        =   88
         Top             =   3000
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox28"
      End
      Begin XtremeSuiteControls.ComboBox cmbRast1 
         Height          =   315
         Left            =   8940
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox28"
         DropDownItemCount=   12
      End
      Begin XtremeSuiteControls.ComboBox cmbMaxTe 
         Height          =   315
         Left            =   8940
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   740
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
      Begin XtremeSuiteControls.ComboBox cmbVorLa 
         Height          =   315
         Left            =   8940
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1640
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkOnlTe 
         Height          =   255
         Left            =   7800
         TabIndex        =   95
         TabStop         =   0   'False
         Tag             =   "0OnlTer"
         Top             =   2660
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Online-Terminbuchungs System"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbRast2 
         Height          =   315
         Left            =   8940
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
         DropDownItemCount=   12
      End
      Begin XtremeSuiteControls.ComboBox cmbMaxPa 
         Height          =   315
         Left            =   8940
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1180
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
      Begin XtremeSuiteControls.ComboBox cmbBuRad 
         Height          =   315
         Left            =   8940
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   2080
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
      Begin XtremeSuiteControls.Label lblLab82 
         Height          =   255
         Left            =   7600
         TabIndex        =   229
         Top             =   780
         Width           =   1270
         _Version        =   1048579
         _ExtentX        =   2240
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Emailerinnerung :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab80 
         Height          =   255
         Left            =   7600
         TabIndex        =   226
         Top             =   2140
         Width           =   1270
         _Version        =   1048579
         _ExtentX        =   2240
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Buchungsradius :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab74 
         Height          =   255
         Left            =   7600
         TabIndex        =   211
         Top             =   1220
         Width           =   1270
         _Version        =   1048579
         _ExtentX        =   2240
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Max / Patient :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab72 
         Height          =   255
         Left            =   7600
         TabIndex        =   197
         Top             =   1680
         Width           =   1270
         _Version        =   1048579
         _ExtentX        =   2240
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Vorlaufstunden :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab71 
         Height          =   255
         Left            =   7600
         TabIndex        =   196
         Top             =   780
         Width           =   1270
         _Version        =   1048579
         _ExtentX        =   2240
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Max. / Tag :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab70 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Zeitraster :"
         Height          =   240
         Left            =   7600
         TabIndex        =   190
         Top             =   340
         Width           =   1270
      End
      Begin VB.Label lblLab46 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Sonntag von :"
         Height          =   240
         Left            =   100
         TabIndex        =   187
         Top             =   3040
         Width           =   1240
      End
      Begin VB.Label lblLab41 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Dienstag von :"
         Height          =   240
         Left            =   100
         TabIndex        =   186
         Top             =   780
         Width           =   1240
      End
      Begin VB.Label lblLab40 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Montag von :"
         Height          =   240
         Left            =   100
         TabIndex        =   185
         Top             =   340
         Width           =   1240
      End
      Begin VB.Label lblLab43 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Donnerstag von :"
         Height          =   240
         Left            =   100
         TabIndex        =   184
         Top             =   1680
         Width           =   1240
      End
      Begin VB.Label lblLab44 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Freitag von :"
         Height          =   240
         Left            =   100
         TabIndex        =   183
         Top             =   2140
         Width           =   1240
      End
      Begin VB.Label lblLab45 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Samstag von :"
         Height          =   240
         Left            =   100
         TabIndex        =   182
         Top             =   2580
         Width           =   1240
      End
      Begin VB.Label lblLab42 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mittwoch von :"
         Height          =   240
         Left            =   100
         TabIndex        =   181
         Top             =   1220
         Width           =   1240
      End
      Begin VB.Label lblLab53 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2730
         TabIndex        =   180
         Top             =   3040
         Width           =   300
      End
      Begin VB.Label lblLab48 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2730
         TabIndex        =   179
         Top             =   780
         Width           =   300
      End
      Begin VB.Label lblLab47 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2730
         TabIndex        =   178
         Top             =   340
         Width           =   300
      End
      Begin VB.Label lblLab50 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2730
         TabIndex        =   177
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label lblLab51 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2730
         TabIndex        =   176
         Top             =   2140
         Width           =   300
      End
      Begin VB.Label lblLab52 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2730
         TabIndex        =   175
         Top             =   2580
         Width           =   300
      End
      Begin VB.Label lblLab49 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   2730
         TabIndex        =   174
         Top             =   1220
         Width           =   300
      End
      Begin VB.Label lblLab67 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   6150
         TabIndex        =   173
         Top             =   3040
         Width           =   300
      End
      Begin VB.Label lblLab62 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   6150
         TabIndex        =   172
         Top             =   780
         Width           =   300
      End
      Begin VB.Label lblLab61 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   6150
         TabIndex        =   171
         Top             =   340
         Width           =   300
      End
      Begin VB.Label lblLab64 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   6150
         TabIndex        =   170
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label lblLab65 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   6150
         TabIndex        =   169
         Top             =   2140
         Width           =   300
      End
      Begin VB.Label lblLab66 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   6150
         TabIndex        =   168
         Top             =   2580
         Width           =   300
      End
      Begin VB.Label lblLab63 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis :"
         Height          =   240
         Left            =   6150
         TabIndex        =   167
         Top             =   1220
         Width           =   300
      End
      Begin VB.Label lblLab60 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "und von :"
         Height          =   240
         Left            =   4120
         TabIndex        =   166
         Top             =   3040
         Width           =   700
      End
      Begin VB.Label lblLab55 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "und von :"
         Height          =   240
         Left            =   4120
         TabIndex        =   165
         Top             =   780
         Width           =   700
      End
      Begin VB.Label lblLab54 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "und von :"
         Height          =   240
         Left            =   4120
         TabIndex        =   164
         Top             =   340
         Width           =   700
      End
      Begin VB.Label lblLab57 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "und von :"
         Height          =   240
         Left            =   4120
         TabIndex        =   163
         Top             =   1680
         Width           =   700
      End
      Begin VB.Label lblLab58 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "und von :"
         Height          =   240
         Left            =   4120
         TabIndex        =   162
         Top             =   2140
         Width           =   700
      End
      Begin VB.Label lblLab59 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "und von :"
         Height          =   240
         Left            =   4120
         TabIndex        =   161
         Top             =   2580
         Width           =   700
      End
      Begin VB.Label lblLab56 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "und von :"
         Height          =   240
         Left            =   4120
         TabIndex        =   160
         Top             =   1220
         Width           =   700
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm8 
      Height          =   3800
      Left            =   120
      TabIndex        =   191
      Top             =   8400
      Visible         =   0   'False
      Width           =   11000
      _Version        =   1048579
      _ExtentX        =   19403
      _ExtentY        =   6703
      _StockProps     =   79
      Caption         =   "Mandantenvorgaben"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtKoste 
         Height          =   350
         Left            =   6970
         TabIndex        =   121
         Tag             =   "0AbrBereich"
         Top             =   2620
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtS3F02 
         Height          =   600
         Left            =   1800
         TabIndex        =   122
         Tag             =   "0Anamnese"
         Top             =   3100
         Width           =   8700
         _Version        =   1048579
         _ExtentX        =   15346
         _ExtentY        =   1058
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbGbKat 
         Height          =   315
         Left            =   6970
         TabIndex        =   116
         Tag             =   "0StaGeb"
         Top             =   300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGbKet 
         Height          =   315
         Left            =   6970
         TabIndex        =   117
         Tag             =   "0StaKet"
         Top             =   760
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbSteue 
         Height          =   315
         Left            =   6970
         TabIndex        =   120
         Tag             =   "0StaStu"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbKtoRa 
         Height          =   315
         Left            =   1800
         TabIndex        =   110
         Tag             =   "0StaRam"
         Top             =   300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbKtoEr 
         Height          =   315
         Left            =   1800
         TabIndex        =   111
         Tag             =   "0StaKon"
         Top             =   760
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbReTyp 
         Height          =   315
         Left            =   6970
         TabIndex        =   119
         Tag             =   "0SteRet"
         Top             =   1690
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbKtoEk 
         Height          =   315
         Left            =   1800
         TabIndex        =   112
         Tag             =   "0StaKo2"
         Top             =   1230
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGeKt1 
         Height          =   315
         Left            =   1800
         TabIndex        =   113
         Tag             =   "0StaGk1"
         Top             =   1690
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGeKt2 
         Height          =   315
         Left            =   1800
         TabIndex        =   114
         Tag             =   "0StaGk2"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbKtoSt 
         Height          =   315
         Left            =   1800
         TabIndex        =   115
         Tag             =   "0StStKt"
         Top             =   2620
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGbKe2 
         Height          =   315
         Left            =   6970
         TabIndex        =   118
         Tag             =   "0Kanton"
         Top             =   1230
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label lblLab84 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Gebührenkette 2 :"
         Height          =   240
         Left            =   5270
         TabIndex        =   231
         Top             =   1290
         Width           =   1605
      End
      Begin XtremeSuiteControls.Label lblLab83 
         Height          =   255
         Left            =   100
         TabIndex        =   230
         Top             =   2660
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Umsatzsteuerkonto :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab81 
         Height          =   240
         Left            =   5270
         TabIndex        =   228
         Top             =   2660
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Kostenstelle :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geldkonto (Kasse) :"
         Height          =   240
         Left            =   100
         TabIndex        =   209
         Top             =   2200
         Width           =   1600
      End
      Begin VB.Label lblLab11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geldkonto (Bank) :"
         Height          =   240
         Left            =   100
         TabIndex        =   208
         Top             =   1750
         Width           =   1600
      End
      Begin VB.Label lblLab10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Erlöskonto (Kasse) :"
         Height          =   240
         Left            =   100
         TabIndex        =   207
         Top             =   1290
         Width           =   1600
      End
      Begin VB.Label lblLab09 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Steuersatz :"
         Height          =   240
         Left            =   5270
         TabIndex        =   206
         Top             =   2200
         Width           =   1600
      End
      Begin VB.Label lblLab08 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Rechnungstyp :"
         Height          =   240
         Left            =   5270
         TabIndex        =   205
         Top             =   1750
         Width           =   1600
      End
      Begin VB.Label lblLab07 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Erlöskonto (Bank) :"
         Height          =   240
         Left            =   100
         TabIndex        =   204
         Top             =   800
         Width           =   1600
      End
      Begin VB.Label lblLab06 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kontenrahmen :"
         Height          =   240
         Left            =   100
         TabIndex        =   203
         Top             =   360
         Width           =   1600
      End
      Begin VB.Label lblLab05 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Gebührenkette 1 :"
         Height          =   240
         Left            =   5270
         TabIndex        =   202
         Top             =   800
         Width           =   1600
      End
      Begin VB.Label lblLab04 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Gebührenkatalog :"
         Height          =   240
         Left            =   5270
         TabIndex        =   201
         Top             =   360
         Width           =   1600
      End
      Begin VB.Label lblLab03 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Therapientexte :"
         Height          =   240
         Left            =   100
         TabIndex        =   200
         Top             =   3100
         Width           =   1600
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F20 
      Height          =   300
      Left            =   4080
      TabIndex        =   123
      TabStop         =   0   'False
      Tag             =   "0R_Briefanrede"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F19 
      Height          =   300
      Left            =   3720
      TabIndex        =   124
      TabStop         =   0   'False
      Tag             =   "0R_Ort"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F18 
      Height          =   300
      Left            =   3360
      TabIndex        =   125
      TabStop         =   0   'False
      Tag             =   "0R_PLZ"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F16 
      Height          =   300
      Left            =   3000
      TabIndex        =   126
      TabStop         =   0   'False
      Tag             =   "0R_Straße"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F15 
      Height          =   300
      Left            =   2280
      TabIndex        =   127
      TabStop         =   0   'False
      Tag             =   "0R_Name"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F14 
      Height          =   300
      Left            =   5160
      TabIndex        =   128
      TabStop         =   0   'False
      Tag             =   "0R_Vorname"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F13 
      Height          =   300
      Left            =   4440
      TabIndex        =   129
      TabStop         =   0   'False
      Tag             =   "0R_Titel"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F12 
      Height          =   300
      Left            =   4800
      TabIndex        =   130
      TabStop         =   0   'False
      Tag             =   "0R_Anrede"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS1F01 
      Height          =   300
      Left            =   2640
      TabIndex        =   131
      TabStop         =   0   'False
      Tag             =   "0Firma1"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   3560
      Left            =   5760
      TabIndex        =   2
      Top             =   1500
      Width           =   5505
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   6279
      _StockProps     =   79
      Caption         =   "Sonstiges"
      UseVisualStyle  =   -1  'True
      Begin XtremeReportControl.ReportControl repCont5 
         Height          =   1800
         Left            =   1360
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   380
         Visible         =   0   'False
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   3175
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIKNum 
         Height          =   350
         Left            =   1360
         TabIndex        =   31
         Tag             =   "0KVNummer"
         Top             =   1690
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox txtS1F12 
         Height          =   315
         Left            =   1365
         TabIndex        =   26
         Tag             =   "0Land"
         Top             =   760
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbBuLnd 
         Height          =   315
         Left            =   1360
         TabIndex        =   30
         Tag             =   "0BunLan"
         Top             =   1230
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbAbrBz 
         Height          =   315
         Left            =   1360
         TabIndex        =   32
         Tag             =   "0KVBez"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbKatal 
         Height          =   315
         Left            =   1365
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "0ID3"
         Top             =   300
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbKanto 
         Height          =   315
         Left            =   1360
         TabIndex        =   29
         Tag             =   "0Kanton"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F13 
         Height          =   350
         Left            =   1360
         TabIndex        =   33
         Tag             =   "0Geboren"
         Top             =   2620
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
      Begin XtremeSuiteControls.ComboBox txtS2F26 
         Height          =   315
         Left            =   1360
         TabIndex        =   35
         Tag             =   "0Familienstand"
         Top             =   3090
         Width           =   1590
         _Version        =   1048579
         _ExtentX        =   2805
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F37 
         Height          =   315
         Left            =   1360
         TabIndex        =   37
         Tag             =   "0Blutgruppe"
         Top             =   3090
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         PasswordChar    =   "*"
      End
      Begin XtremeSuiteControls.CheckBox chkOpti2 
         Height          =   225
         Left            =   3400
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "0Edit"
         Top             =   3140
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Synchronisation"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F30 
         Height          =   350
         Left            =   3900
         TabIndex        =   34
         Tag             =   "0Mandant"
         Top             =   2620
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1676
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtGLNum 
         Height          =   315
         Left            =   1360
         TabIndex        =   27
         Tag             =   "0GLN"
         Top             =   1230
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   13
      End
      Begin XtremeSuiteControls.FlatEdit txtZSRnr 
         Height          =   350
         Left            =   1360
         TabIndex        =   28
         Tag             =   "0ZSR"
         Top             =   1690
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   7
      End
      Begin XtremeSuiteControls.Label lblLab18 
         Height          =   240
         Left            =   100
         TabIndex        =   216
         Top             =   800
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Land :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab16 
         Height          =   240
         Left            =   100
         TabIndex        =   215
         Top             =   3140
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Passwort :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab15 
         Height          =   240
         Left            =   100
         TabIndex        =   214
         Top             =   2200
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "KV Bezirk :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab13 
         Height          =   240
         Left            =   100
         TabIndex        =   213
         Top             =   1290
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Bundesland :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab14 
         Height          =   240
         Left            =   100
         TabIndex        =   212
         Top             =   1750
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "LANR :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Nummer :"
         Height          =   240
         Left            =   3140
         TabIndex        =   199
         Top             =   2670
         Width           =   700
      End
      Begin XtremeSuiteControls.Label lblLab17 
         Height          =   240
         Left            =   100
         TabIndex        =   133
         Top             =   360
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Grundvorgabe :"
         ForeColor       =   192
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren :"
         Height          =   240
         Left            =   100
         TabIndex        =   132
         Top             =   2670
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   3070
      Left            =   5760
      TabIndex        =   4
      Top             =   5200
      Width           =   5505
      _Version        =   1048579
      _ExtentX        =   9710
      _ExtentY        =   5415
      _StockProps     =   79
      Caption         =   "Finanzdaten 1"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtS1F39 
         Height          =   350
         Left            =   1360
         TabIndex        =   42
         Tag             =   "0ZSR"
         Top             =   760
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F38 
         Height          =   350
         Left            =   1360
         TabIndex        =   41
         Tag             =   "0GLN"
         Top             =   300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F34 
         Height          =   350
         Left            =   1360
         TabIndex        =   45
         Tag             =   "0BIC"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F33 
         Height          =   350
         Left            =   1360
         TabIndex        =   44
         Tag             =   "0IBAN"
         Top             =   1690
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F04 
         Height          =   350
         Left            =   1360
         TabIndex        =   39
         Tag             =   "0BLZ"
         Top             =   300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F05 
         Height          =   350
         Left            =   1360
         TabIndex        =   40
         Tag             =   "0Konto"
         Top             =   760
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F03 
         Height          =   350
         Left            =   1360
         TabIndex        =   43
         Tag             =   "0Bank"
         Top             =   1230
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F22 
         Height          =   350
         Left            =   1360
         TabIndex        =   46
         Tag             =   "0Abteilung"
         Top             =   2620
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.Label lblLab79 
         Height          =   240
         Left            =   100
         TabIndex        =   225
         Top             =   2200
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "BIC :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab78 
         Height          =   240
         Left            =   100
         TabIndex        =   224
         Top             =   1750
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "IBAN :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab77 
         Height          =   240
         Left            =   100
         TabIndex        =   223
         Top             =   1290
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Kreditinstitut :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab76 
         Height          =   240
         Left            =   100
         TabIndex        =   222
         Top             =   800
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Kontoinhaber :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab75 
         Height          =   240
         Left            =   100
         TabIndex        =   221
         Top             =   360
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Bankleitzahl :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab22 
         Height          =   240
         Left            =   100
         TabIndex        =   220
         Top             =   2670
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Steuernummer :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3070
      Left            =   120
      TabIndex        =   3
      Top             =   5200
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   5415
      _StockProps     =   79
      Caption         =   "Kommunikation"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtS1F17 
         Height          =   350
         Left            =   1360
         TabIndex        =   19
         Tag             =   "0Telefon3"
         Top             =   760
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F16 
         Height          =   350
         Left            =   1360
         TabIndex        =   18
         Tag             =   "0Telefon2"
         Top             =   300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F27 
         Height          =   350
         Left            =   1360
         TabIndex        =   21
         Tag             =   "0Internet"
         Top             =   1690
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
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
         Left            =   1360
         TabIndex        =   20
         Tag             =   "0Telefon5"
         Top             =   1230
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
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
      Begin XtremeSuiteControls.FlatEdit txtS1F23 
         Height          =   350
         Left            =   1360
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "0Postfach"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F15 
         Height          =   350
         Left            =   1360
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "0Telefon1"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbS2F10 
         Height          =   315
         Left            =   1360
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "0IDP"
         Top             =   2620
         Visible         =   0   'False
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox11"
         DropDownItemCount=   15
      End
      Begin XtremeSuiteControls.Label lblLab20 
         Height          =   240
         Left            =   100
         TabIndex        =   218
         Top             =   2200
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "PVS-Nr.:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab24 
         Height          =   240
         Left            =   100
         TabIndex        =   195
         Top             =   2670
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Mandant :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Telefax :"
         Height          =   240
         Left            =   100
         TabIndex        =   137
         Top             =   800
         Width           =   1200
      End
      Begin VB.Label Label03 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Telefon :"
         Height          =   240
         Left            =   100
         TabIndex        =   136
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Internet :"
         Height          =   240
         Left            =   100
         TabIndex        =   135
         Top             =   1750
         Width           =   1200
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   240
         Left            =   100
         TabIndex        =   134
         Top             =   1290
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3560
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   5505
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   6279
      _StockProps     =   79
      Caption         =   "Adressdaten"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnPost1 
         Height          =   350
         Left            =   2190
         TabIndex        =   15
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
      Begin XtremeSuiteControls.FlatEdit txtS1F09 
         Height          =   350
         Left            =   2540
         TabIndex        =   16
         Tag             =   "0Ort"
         Top             =   2620
         Width           =   2340
         _Version        =   1048579
         _ExtentX        =   4128
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F08 
         Height          =   350
         Left            =   1360
         TabIndex        =   14
         Tag             =   "0PLZ"
         Top             =   2620
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F06 
         Height          =   350
         Left            =   1360
         TabIndex        =   13
         Tag             =   "0Straße"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F03 
         Height          =   350
         Left            =   3400
         TabIndex        =   10
         Tag             =   "0Titel"
         Top             =   760
         Width           =   1440
         _Version        =   1048579
         _ExtentX        =   2540
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F05 
         Height          =   350
         Left            =   1360
         TabIndex        =   12
         Tag             =   "0Name"
         Top             =   1690
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F04 
         Height          =   350
         Left            =   1360
         TabIndex        =   11
         Tag             =   "0Vorname"
         Top             =   1230
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F11 
         Height          =   350
         Left            =   1360
         TabIndex        =   8
         Tag             =   "0R_Firma1"
         Top             =   300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox txtS1F02 
         Height          =   315
         Left            =   1360
         TabIndex        =   9
         Tag             =   "0Anrede"
         Top             =   760
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2275
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F24 
         Height          =   350
         Left            =   1360
         TabIndex        =   17
         Tag             =   "0Beruf"
         Top             =   3090
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.Label lblLab19 
         Height          =   240
         Left            =   100
         TabIndex        =   217
         Top             =   360
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Praxis / Firma :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname :"
         Height          =   240
         Left            =   100
         TabIndex        =   144
         Top             =   1290
         Width           =   1200
      End
      Begin VB.Label Label02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PLZ / Ort :"
         Height          =   240
         Left            =   100
         TabIndex        =   143
         Top             =   2670
         Width           =   1200
      End
      Begin VB.Label Label01 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Straße :"
         Height          =   240
         Left            =   100
         TabIndex        =   142
         Top             =   2200
         Width           =   1200
      End
      Begin VB.Label lblLab73 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Nachname :"
         Height          =   240
         Left            =   100
         TabIndex        =   141
         Top             =   1750
         Width           =   1200
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anrede :"
         Height          =   240
         Left            =   100
         TabIndex        =   140
         Top             =   800
         Width           =   1200
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Titel :"
         Height          =   240
         Left            =   2700
         TabIndex        =   139
         Top             =   800
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Berufsbezeich.:"
         Height          =   240
         Left            =   100
         TabIndex        =   138
         Top             =   3140
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F27 
      Height          =   300
      Left            =   5520
      TabIndex        =   145
      TabStop         =   0   'False
      Tag             =   "0Datum"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F22 
      Height          =   300
      Left            =   1200
      TabIndex        =   146
      TabStop         =   0   'False
      Tag             =   "0R_Land"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS2F25 
      Height          =   300
      Left            =   6240
      TabIndex        =   147
      TabStop         =   0   'False
      Tag             =   "0R_Geboren"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS3F01 
      Height          =   300
      Left            =   6600
      TabIndex        =   148
      TabStop         =   0   'False
      Tag             =   "0Anschrift"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS1F20 
      Height          =   195
      Left            =   720
      TabIndex        =   149
      TabStop         =   0   'False
      Tag             =   "0DuSie"
      Top             =   13000
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
   Begin XtremeSuiteControls.ComboBox cmbS1F10 
      Height          =   315
      Left            =   7800
      TabIndex        =   150
      TabStop         =   0   'False
      Tag             =   "0Briefanrede"
      Top             =   12300
      Width           =   600
      _Version        =   1048579
      _ExtentX        =   1058
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Appearance      =   6
   End
   Begin XtremeSuiteControls.FlatEdit txtGesch 
      Height          =   195
      Left            =   240
      TabIndex        =   151
      TabStop         =   0   'False
      Tag             =   "0Geschlecht"
      Top             =   12300
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtS4F01 
      Height          =   300
      Left            =   5880
      TabIndex        =   152
      TabStop         =   0   'False
      Tag             =   "0Kontoinhaber"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbS1F08 
      Height          =   315
      Left            =   8500
      TabIndex        =   153
      TabStop         =   0   'False
      Tag             =   "0GeschlTyp"
      Top             =   12300
      Width           =   600
      _Version        =   1048579
      _ExtentX        =   1058
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   3070
      Left            =   11400
      TabIndex        =   6
      Top             =   5200
      Visible         =   0   'False
      Width           =   5505
      _Version        =   1048579
      _ExtentX        =   9710
      _ExtentY        =   5415
      _StockProps     =   79
      Caption         =   "Finanzdaten 2"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtGIDNr 
         Height          =   350
         Left            =   1360
         TabIndex        =   109
         Tag             =   "0GID"
         Top             =   2620
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtBICN2 
         Height          =   350
         Left            =   1360
         TabIndex        =   108
         Tag             =   "0BIC2"
         Top             =   2160
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtIBAN2 
         Height          =   350
         Left            =   1360
         TabIndex        =   107
         Tag             =   "0IBAN2"
         Top             =   1690
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtBLZ02 
         Height          =   350
         Left            =   1360
         TabIndex        =   104
         Tag             =   "0BLZ2"
         Top             =   300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtKont2 
         Height          =   350
         Left            =   1360
         TabIndex        =   105
         Tag             =   "0Konto2"
         Top             =   760
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtBank2 
         Height          =   350
         Left            =   1360
         TabIndex        =   106
         Tag             =   "0Bank2"
         Top             =   1230
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.Label lblLab23 
         Height          =   240
         Left            =   100
         TabIndex        =   194
         Top             =   2670
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Gläubiger-ID :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label33 
         Height          =   240
         Left            =   100
         TabIndex        =   193
         Top             =   2200
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "BIC :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kreditinstitut :"
         Height          =   240
         Left            =   100
         TabIndex        =   158
         Top             =   1290
         Width           =   1200
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bankleitzahl :"
         Height          =   240
         Left            =   100
         TabIndex        =   157
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kontoinhaber :"
         Height          =   240
         Left            =   100
         TabIndex        =   156
         Top             =   800
         Width           =   1200
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "IBAN :"
         Height          =   240
         Left            =   100
         TabIndex        =   155
         Top             =   1750
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm6 
      Height          =   3560
      Left            =   11400
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   5505
      _Version        =   1048579
      _ExtentX        =   9710
      _ExtentY        =   6279
      _StockProps     =   79
      Caption         =   "Benennungen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtS1F10 
         Height          =   350
         Left            =   1360
         TabIndex        =   100
         Tag             =   "0Firma2"
         Top             =   760
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         MaxLength       =   30
      End
      Begin XtremeSuiteControls.FlatEdit txtS2F23 
         Height          =   350
         Left            =   1360
         TabIndex        =   102
         TabStop         =   0   'False
         Tag             =   "0Telefon6"
         Top             =   2620
         Width           =   3150
         _Version        =   1048579
         _ExtentX        =   5556
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtS1F11 
         Height          =   350
         Left            =   1360
         TabIndex        =   99
         Tag             =   "0IDKurz"
         Top             =   300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   150
      End
      Begin XtremeSuiteControls.PushButton btnSign1 
         Height          =   350
         Left            =   4550
         TabIndex        =   103
         TabStop         =   0   'False
         ToolTipText     =   "Ordnen dem Mitarbeiter eine Signaturdatei zu"
         Top             =   2620
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtEmSig 
         Height          =   1230
         Left            =   1360
         TabIndex        =   101
         Tag             =   "0Objekt"
         Top             =   1230
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   2170
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.Label lblLab21 
         Height          =   240
         Left            =   100
         TabIndex        =   219
         Top             =   2670
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Signaturdatei :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Verkehrsname :"
         Height          =   240
         Left            =   100
         TabIndex        =   210
         Top             =   800
         Width           =   1200
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Emailsignatur :"
         Height          =   240
         Left            =   100
         TabIndex        =   192
         Top             =   1290
         Width           =   1200
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anzeigename :"
         Height          =   240
         Left            =   100
         TabIndex        =   159
         Top             =   360
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtZeit1 
      Height          =   300
      Left            =   6960
      TabIndex        =   188
      TabStop         =   0   'False
      Tag             =   "0Sprechzeiten"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAdrGr 
      Height          =   200
      Left            =   0
      TabIndex        =   189
      TabStop         =   0   'False
      Tag             =   "0AdrGruppe"
      Top             =   12000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtZeit2 
      Height          =   300
      Left            =   7340
      TabIndex        =   198
      TabStop         =   0   'False
      Tag             =   "0Buchungszeiten"
      Top             =   12300
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   615
      Left            =   17640
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   840
      Width           =   615
      _Version        =   1048579
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   64
      Show3DBorder    =   2
      VisualTheme     =   0
   End
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   735
      Left            =   240
      TabIndex        =   154
      Top             =   480
      Width           =   11175
      _Version        =   1048579
      _ExtentX        =   19711
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   $"frmMandant.frx":6852
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeCommandBars.CommandBars comBar01 
      Left            =   120
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   4
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   720
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Shape shpLabl1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   120
      Top             =   360
      Width           =   11700
   End
End
Attribute VB_Name = "frmMandant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private FS As Form
Private AktCo As VB.Control
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Rahm7 As XtremeSuiteControls.GroupBox
Private Rahm8 As XtremeSuiteControls.GroupBox
Private ChKaAu As XtremeSuiteControls.CheckBox
Private ChDefr As XtremeSuiteControls.CheckBox
Private ChOnTe As XtremeSuiteControls.CheckBox
Private Labl71 As XtremeSuiteControls.Label
Private Labl72 As XtremeSuiteControls.Label
Private Labl74 As XtremeSuiteControls.Label
Private Labl80 As XtremeSuiteControls.Label
Private Labl82 As XtremeSuiteControls.Label
Private ChS01 As XtremeSuiteControls.CheckBox
Private ChS02 As XtremeSuiteControls.CheckBox
Private ChS03 As XtremeSuiteControls.CheckBox
Private ChS04 As XtremeSuiteControls.CheckBox
Private ChS05 As XtremeSuiteControls.CheckBox
Private ChS06 As XtremeSuiteControls.CheckBox
Private ChS07 As XtremeSuiteControls.CheckBox
Private ChS08 As XtremeSuiteControls.CheckBox
Private ChS09 As XtremeSuiteControls.CheckBox
Private ChS10 As XtremeSuiteControls.CheckBox
Private ChS11 As XtremeSuiteControls.CheckBox
Private ChS12 As XtremeSuiteControls.CheckBox
Private ChS13 As XtremeSuiteControls.CheckBox
Private ChS14 As XtremeSuiteControls.CheckBox
Private TxDum As XtremeSuiteControls.FlatEdit
Private txGebo As XtremeSuiteControls.FlatEdit
Private txVorn As XtremeSuiteControls.FlatEdit
Private txName As XtremeSuiteControls.FlatEdit
Private txStra As XtremeSuiteControls.FlatEdit
Private txPost As XtremeSuiteControls.FlatEdit
Private txOrte As XtremeSuiteControls.FlatEdit
Private txTele As XtremeSuiteControls.FlatEdit
Private txFaxe As XtremeSuiteControls.FlatEdit
Private txBank As XtremeSuiteControls.FlatEdit
Private txBan2 As XtremeSuiteControls.FlatEdit
Private txBaLZ As XtremeSuiteControls.FlatEdit
Private txBaL2 As XtremeSuiteControls.FlatEdit
Private txKont As XtremeSuiteControls.FlatEdit
Private txKon2 As XtremeSuiteControls.FlatEdit
Private txSteu As XtremeSuiteControls.FlatEdit
Private txIKNr As XtremeSuiteControls.FlatEdit
Private txBeru As XtremeSuiteControls.FlatEdit
Private txTite As XtremeSuiteControls.FlatEdit
Private TxEmai As XtremeSuiteControls.FlatEdit
Private TxIntr As XtremeSuiteControls.FlatEdit
Private TxPrax As XtremeSuiteControls.FlatEdit
Private TxIBAN As XtremeSuiteControls.FlatEdit
Private TxBIC1 As XtremeSuiteControls.FlatEdit
Private TxIBA2 As XtremeSuiteControls.FlatEdit
Private TxBIC2 As XtremeSuiteControls.FlatEdit
Private txGlID As XtremeSuiteControls.FlatEdit
Private TxAbre As XtremeSuiteControls.FlatEdit
Private txBeme As XtremeSuiteControls.FlatEdit
Private TxLand As XtremeSuiteControls.ComboBox
Private TxAnre As XtremeSuiteControls.ComboBox
Private cmFach As XtremeSuiteControls.ComboBox
Private cmBuLa As XtremeSuiteControls.ComboBox
Private cmKVBz As XtremeSuiteControls.ComboBox
Private cmKant As XtremeSuiteControls.ComboBox
Private cmKata As XtremeSuiteControls.ComboBox
Private cmKett As XtremeSuiteControls.ComboBox
Private cmKet2 As XtremeSuiteControls.ComboBox
Private cmRahm As XtremeSuiteControls.ComboBox
Private cmKont As XtremeSuiteControls.ComboBox
Private cmReTy As XtremeSuiteControls.ComboBox
Private cmSteu As XtremeSuiteControls.ComboBox
Private cmRas1 As XtremeSuiteControls.ComboBox
Private cmRas2 As XtremeSuiteControls.ComboBox
Private cmMaxT As XtremeSuiteControls.ComboBox
Private cmMaxP As XtremeSuiteControls.ComboBox
Private cmVorl As XtremeSuiteControls.ComboBox
Private cmBuRa As XtremeSuiteControls.ComboBox
Private cmNoti As XtremeSuiteControls.ComboBox
Private cmbS01 As XtremeSuiteControls.ComboBox
Private cmbS02 As XtremeSuiteControls.ComboBox
Private cmbS03 As XtremeSuiteControls.ComboBox
Private cmbS04 As XtremeSuiteControls.ComboBox
Private cmbS05 As XtremeSuiteControls.ComboBox
Private cmbS06 As XtremeSuiteControls.ComboBox
Private cmbS07 As XtremeSuiteControls.ComboBox
Private cmbS08 As XtremeSuiteControls.ComboBox
Private cmbS09 As XtremeSuiteControls.ComboBox
Private cmbS10 As XtremeSuiteControls.ComboBox
Private cmbS11 As XtremeSuiteControls.ComboBox
Private cmbS12 As XtremeSuiteControls.ComboBox
Private cmbS13 As XtremeSuiteControls.ComboBox
Private cmbS14 As XtremeSuiteControls.ComboBox
Private cmbS15 As XtremeSuiteControls.ComboBox
Private cmbS16 As XtremeSuiteControls.ComboBox
Private cmbS17 As XtremeSuiteControls.ComboBox
Private cmbS18 As XtremeSuiteControls.ComboBox
Private cmbS19 As XtremeSuiteControls.ComboBox
Private cmbS20 As XtremeSuiteControls.ComboBox
Private cmbS21 As XtremeSuiteControls.ComboBox
Private cmbS22 As XtremeSuiteControls.ComboBox
Private cmbS23 As XtremeSuiteControls.ComboBox
Private cmbS24 As XtremeSuiteControls.ComboBox
Private cmbS25 As XtremeSuiteControls.ComboBox
Private cmbS26 As XtremeSuiteControls.ComboBox
Private cmbS27 As XtremeSuiteControls.ComboBox
Private cmbS28 As XtremeSuiteControls.ComboBox
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private CoDia As XtremeSuiteControls.CommonDialog
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Public AdAnd As Boolean 'Adressenänderung

Private Const KEYEVENTF_KEYUP = &H2
Private Const CB_SHOWDROPDOWN = &H14F
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const GWL_WNDPROC = (-4)

Private RetWe As Long
Private TbOld As Long
Private TagWe As String
Private LogLa As Boolean
Private SpNeu As Boolean 'Neue Sprechzeiten

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private clFil As clsFile
Private clFen As clsFenster

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub btnPost1_Click()
    If GlAdL = False Then
        Adr_Pos False, True, True
    End If
End Sub

Private Sub btnSign1_Click()
    FOpn
End Sub

Private Sub chkDefra_Click()

TagWe = Mid$(Me.chkDefra.Tag, 2, Len(Me.chkDefra.Tag) - 1)

If GlAdL = False Then
    Me.chkDefra.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub chkKaAus_Click()
On Error Resume Next

Set ChOnTe = Me.chkOnlTe
Set ChKaAu = Me.chkKaAus

TagWe = Mid$(ChKaAu.Tag, 2, Len(ChKaAu.Tag) - 1)

If GlAdL = False Then 'Formular wird geladen
    If ChKaAu.Value = xtpChecked Then
        If ChOnTe.Value = xtpChecked Then
            GlAdL = True
            ChOnTe.Value = xtpUnchecked
            ChOnTe.Tag = "1OnlTer"
            GlAdL = False
        End If
        ChOnTe.Enabled = False
    Else
        If GlOTS = True Then 'Online-Terminbuchungs System aktivieren
            If ChOnTe.Enabled = False Then
                ChOnTe.Enabled = True
            End If
        End If
    End If
    ChKaAu.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub chkOnlTe_Click()
On Error Resume Next

Set ChOnTe = Me.chkOnlTe
Set ChKaAu = Me.chkKaAus

TagWe = Mid$(ChOnTe.Tag, 2, Len(ChOnTe.Tag) - 1)

If GlAdL = False Then
    ChOnTe.Tag = "1" & TagWe
    GlAdS = True
    If ChOnTe.Value = xtpChecked Then
        If ChKaAu.Value = xtpChecked Then
            GlAdL = True
            ChKaAu.Value = xtpUnchecked
            ChKaAu.Tag = "1Gesperrt"
            GlAdL = False
        End If
        ChKaAu.Enabled = False
    Else
        If UBound(GlMiA) > 1 Then 'Aktive Mitarbeiter
            If ChKaAu.Enabled = False Then
                ChKaAu.Enabled = True
            End If
        End If
    End If
End If

End Sub
Private Sub cmbAbrBz_Click()

TagWe = Mid$(Me.cmbAbrBz.Tag, 2, Len(Me.cmbAbrBz.Tag) - 1)

If GlAdL = False Then
    Me.cmbAbrBz.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub cmbBuLnd_Click()

TagWe = Mid$(Me.cmbBuLnd.Tag, 2, Len(Me.cmbBuLnd.Tag) - 1)

If GlAdL = False Then
    Me.cmbBuLnd.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbBuRad_Click()
    If GlAdL = False Then
        GlAdS = True
    End If
End Sub
Private Sub cmbGbKat_Click()

Dim StaKa As Long

TagWe = Mid$(Me.cmbGbKat.Tag, 2, Len(Me.cmbGbKat.Tag) - 1)

If GlAdL = False Then
    StaKa = Me.cmbGbKat.ItemData(Me.cmbGbKat.ListIndex)
    Me.cmbGbKat.Tag = "1" & TagWe
    GlAdS = True
    DoEvents
    Man_Cmb 2, StaKa
End If

End Sub
Private Sub cmbGbKat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbGbKe2_Click()

TagWe = Mid$(Me.cmbGbKe2.Tag, 2, Len(Me.cmbGbKe2.Tag) - 1)

If GlAdL = False Then
    Me.cmbGbKe2.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbGbKe2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbGbKet_Click()

TagWe = Mid$(Me.cmbGbKet.Tag, 2, Len(Me.cmbGbKet.Tag) - 1)

If GlAdL = False Then
    Me.cmbGbKet.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbGbKet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbGeKt1_Click()

TagWe = Mid$(Me.cmbGeKt1.Tag, 2, Len(Me.cmbGeKt1.Tag) - 1)

If GlAdL = False Then
    Me.cmbGeKt1.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbGeKt1_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbGeKt2_Click()

TagWe = Mid$(Me.cmbGeKt2.Tag, 2, Len(Me.cmbGeKt2.Tag) - 1)

If GlAdL = False Then
    Me.cmbGeKt2.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbGeKt2_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbKanto_Click()

TagWe = Mid$(Me.cmbKanto.Tag, 2, Len(Me.cmbKanto.Tag) - 1)

If GlAdL = False Then
    Me.cmbKanto.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub cmbKatal_Click()

Set cmFach = Me.cmbKatal

TagWe = Mid$(Me.cmbKatal.Tag, 2, Len(Me.cmbKatal.Tag) - 1)

If GlAdL = False Then
    Me.cmbKatal.Tag = "1" & TagWe
    GlAdS = True
    MOpn
End If
    
End Sub
Private Sub cmbKtoEk_Click()

TagWe = Mid$(Me.cmbKtoEk.Tag, 2, Len(Me.cmbKtoEk.Tag) - 1)

If GlAdL = False Then
    Me.cmbKtoEk.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbKtoEk_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbKtoEr_Click()

TagWe = Mid$(Me.cmbKtoEr.Tag, 2, Len(Me.cmbKtoEr.Tag) - 1)

If GlAdL = False Then
    Me.cmbKtoEr.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbKtoEr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbKtoRa_Click()

Dim StaRa As Long

TagWe = Mid$(Me.cmbKtoRa.Tag, 2, Len(Me.cmbKtoRa.Tag) - 1)

If GlAdL = False Then
    StaRa = Me.cmbKtoRa.ItemData(Me.cmbKtoRa.ListIndex)
    Me.cmbKtoRa.Tag = "1" & TagWe
    GlAdS = True
    DoEvents
    Man_Cmb 1, StaRa
End If

End Sub

Private Sub cmbKtoRa_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbKtoSt_Click()

TagWe = Mid$(Me.cmbKtoSt.Tag, 2, Len(Me.cmbKtoSt.Tag) - 1)

If GlAdL = False Then
    Me.cmbKtoSt.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbKtoSt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbMaxPa_Click()
    If GlAdL = False Then
        GlAdS = True
    End If
End Sub
Private Sub cmbMaxTe_Click()
    If GlAdL = False Then
        GlAdS = True
    End If
End Sub

Private Sub cmbNotVa_Click()
    If GlAdL = False Then
        GlAdS = True
    End If
End Sub
Private Sub cmbRast2_Click()
    If GlAdL = False Then
        FRast
        GlAdS = True
    End If
End Sub

Private Sub cmbReTyp_Click()

TagWe = Mid$(Me.cmbReTyp.Tag, 2, Len(Me.cmbReTyp.Tag) - 1)

If GlAdL = False Then
    Me.cmbReTyp.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbReTyp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
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
    
    GlAdS = True
End If

End Sub
Private Sub cmbS1F08_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS2F10_Click()

TagWe = Mid$(Me.cmbS2F10.Tag, 2, Len(Me.cmbS2F10.Tag) - 1)

Me.cmbS2F10.Tag = "1" & TagWe

GlAdS = True

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


Private Sub cmbSteue_Click()

TagWe = Mid$(Me.cmbSteue.Tag, 2, Len(Me.cmbSteue.Tag) - 1)

If GlAdL = False Then
    Me.cmbSteue.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub cmbSteue_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbVorLa_Click()
    If GlAdL = False Then
        GlAdS = True
    End If
End Sub

Private Sub Form_Activate()

TbOld = RibTab_Adr_Haupt

End Sub
Private Sub Form_Load()

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 11000
    .ClientMaxWidth = 13000
    .ClientMinHeight = 10300
    .ClientMinWidth = 11500
    If GlMId < 0 Then
        .TopMost = True
    End If
End With

If GlMId = -2 Then
    LogLa = True
End If

Set FrmEx = Nothing

End Sub

Private Sub repCont5_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAdL = False Then
        GlAdS = True
    End If
End Sub
Private Sub repCont5_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAdL = False Then
        GlAdS = True
    End If
End Sub

Private Sub txtGLNum_Change()

TagWe = Mid$(Me.txtGLNum.Tag, 2, Len(Me.txtGLNum.Tag) - 1)

If GlAdL = False Then
    Me.txtGLNum.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtGLNum_GotFocus()
    Me.txtGLNum.SelStart = 0
    Me.txtGLNum.SelLength = Len(Me.txtGLNum.Text)
End Sub
Private Sub txtBICN2_Change()

TagWe = Mid$(Me.txtBICN2.Tag, 2, Len(Me.txtBICN2.Tag) - 1)

If GlAdL = False Then
    Me.txtBICN2.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtBICN2_GotFocus()
    Me.txtBICN2.SelStart = 0
    Me.txtBICN2.SelLength = Len(Me.txtBICN2.Text)
End Sub

Private Sub txtBICN2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtBICN2_LostFocus()
    If GlAdL = False Then
        If Me.txtBICN2.Text <> vbNullString Then
            Me.txtBICN2.Text = SNaUm(Me.txtBICN2.Text)
        End If
    End If
End Sub

Private Sub txtEmSig_Change()

TagWe = Mid$(Me.txtEmSig.Tag, 2, Len(Me.txtEmSig.Tag) - 1)

If GlAdL = False Then
    Me.txtEmSig.Tag = "1" & TagWe
    GlAdS = True
End If
    
End Sub
Private Sub txtGIDNr_Change()

TagWe = Mid$(Me.txtGIDNr.Tag, 2, Len(Me.txtGIDNr.Tag) - 1)

If GlAdL = False Then
    Me.txtGIDNr.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtGIDNr_GotFocus()
    Me.txtGIDNr.SelStart = 0
    Me.txtGIDNr.SelLength = Len(Me.txtGIDNr.Text)
End Sub

Private Sub txtGIDNr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtGLNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtIBAN2_LostFocus()
    If GlAdL = False Then 'Formular wird geladen
        If Me.txtIBAN2.Text <> vbNullString Then
            Me.txtIBAN2.Text = SNaUm(Me.txtIBAN2.Text)
        End If
    End If
End Sub

Private Sub txtKoste_Change()

TagWe = Mid$(Me.txtKoste.Tag, 2, Len(Me.txtKoste.Tag) - 1)

If GlAdL = False Then
    Me.txtKoste.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtKoste_GotFocus()
    Me.txtKoste.SelStart = 0
    Me.txtKoste.SelLength = Len(Me.txtKoste.Text)
End Sub

Private Sub txtKoste_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F05_LostFocus()
On Error Resume Next

If GlAdL = False Then
    MKopi
    DoEvents
    AdBrf , True
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
        If Me.txtS1F08.Text <> vbNullString Then
            Adr_Pos False, True
        End If
    End If
End Sub
Private Sub txtS1F09_Click()
    If GlAdL = False Then
        If Me.txtS1F08.Text <> vbNullString Then
            If Me.txtS1F09.Text = vbNullString Then
                Adr_Pos False, True
            End If
        End If
    End If
End Sub
Private Sub txtS1F10_Change()

TagWe = Mid$(Me.txtS1F10.Tag, 2, Len(Me.txtS1F10.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F10.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F10_GotFocus()
    Me.txtS1F10.SelStart = 0
    Me.txtS1F10.SelLength = Len(Me.txtS1F10.Text)
End Sub
Private Sub txtS1F10_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F11_GotFocus()
    Me.txtS1F11.SelStart = 0
    Me.txtS1F11.SelLength = Len(Me.txtS1F11.Text)
End Sub
Private Sub txtS1F11_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F12_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
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

Private Sub txtS1F13_LostFocus()
Dim NeuDa As Date

If IsDate(Me.txtS1F13.Text) Then
    NeuDa = Me.txtS1F13.Text
    Me.txtS1F13.Text = NeuDa
    If Me.txtS2F25.Text = vbNullString Then
        If Me.txtS1F13.Text <> vbNullString Then
            Me.txtS2F25.Text = Me.txtS1F13.Text
        End If
    End If
End If

End Sub
Private Sub txtS1F20_Change()

TagWe = Mid$(Me.txtS2F20.Tag, 2, Len(Me.txtS2F20.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F20.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FSpDe()
On Error GoTo DaErr
'Lädt die Sprechzeitendetails

Dim IdxNr As Long
Dim TabId As Long
Dim TmStr As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Set CmCom = CmBrs.FindControl(CmCom, AD_Sprechzeit_Auswa, , True)

TabId = RbTab.id

IdxNr = CmCom.ItemData(CmCom.ListIndex)

TmStr = Man_SpD

If TmStr <> vbNullString Then
    GlAdL = True

    Screen.MousePointer = vbHourglass
    clFen.FenDsk 2
    
    MRast TmStr
            
    clFen.FenDsk 3
    Screen.MousePointer = vbNormal

    GlAdL = False
End If

Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpDe " & Err.Number
Resume Next

End Sub
Private Sub FSpNe()
On Error GoTo NeErr
'Legt neue Sprechzeiten an

Dim StaDa As Date
Dim TabId As Long
Dim TmStr As String
Dim GeSta As Integer
Dim GeDat As Integer
Dim AktWo As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmDat As XtremeCommandBars.CommandBarComboBox
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmDat = CmBrs.FindControl(CmDat, AD_Sprechzeit_Datum, , True)
Set CmCom = CmBrs.FindControl(CmCom, AD_Sprechzeit_Auswa, , True)

TabId = RbTab.id

GeSta = CmCom.ListCount
GeDat = CmDat.ListCount

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If GeSta > 0 Then
    TmStr = CmCom.List(GeSta)
    StaDa = CDate(Right$(TmStr, 10))
    For AktWo = 1 To GeDat
        If StaDa = CDate(CmDat.List(AktWo)) Then
            CmDat.ListIndex = AktWo + 1
            Exit For
        End If
    Next AktWo
Else
    CmDat.ListIndex = 1
End If

If CmDat.Enabled = False Then CmDat.Enabled = True

GlAdL = True

MRast

GlAdL = False

clFen.FenDsk 3
Screen.MousePointer = vbNormal

SpNeu = True

Set clFen = Nothing

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpNe " & Err.Number
Resume Next

End Sub
Private Sub FSpLo()
On Error GoTo DaErr
'Löscht die Sprechzeitendetails

Dim IdxNr As Long
Dim TabId As Long
Dim TmStr As String
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Set CmCom = CmBrs.FindControl(CmCom, AD_Sprechzeit_Auswa, , True)

Tit1 = "Sprechzeiteneintrag Entfernen"
Mld1 = "Möchten Sie den ausgewählten Sprechzeiteneintrag wirklich entfernen?"

TabId = RbTab.id

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then

    Man_SpK
    DoEvents

    If CmCom.ListCount > 0 Then
        CmCom.Clear
        DoEvents
    End If

    Man_SpL
    DoEvents

    IdxNr = CmCom.ItemData(CmCom.ListIndex)

    TmStr = Man_SpD

    If TmStr <> vbNullString Then
        GlAdL = True
    
        Screen.MousePointer = vbHourglass
        clFen.FenDsk 2
        
        MRast TmStr
                    
        clFen.FenDsk 3
        Screen.MousePointer = vbNormal
        
        GlAdL = False
    End If
End If

Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpLo " & Err.Number
Resume Next

End Sub


Private Sub FSpSa()
On Error GoTo NeErr
'Speichert neue Sprechzeiten

Dim IdxNr As Long
Dim TabId As Long
Dim StaDa As Date
Dim TmStr As String
Dim BuStr As String
Dim TxZe1 As XtremeSuiteControls.FlatEdit
Dim TxZe2 As XtremeSuiteControls.FlatEdit
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmDat As XtremeCommandBars.CommandBarComboBox
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMandant
Set TxDum = FM.txtDummy
Set TxZe1 = FM.txtZeit1
Set TxZe2 = FM.txtZeit2
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Set CmDat = CmBrs.FindControl(CmDat, AD_Sprechzeit_Datum, , True)
Set CmCom = CmBrs.FindControl(CmCom, AD_Sprechzeit_Auswa, , True)

TabId = RbTab.id

TxDum.SetFocus

If SpNeu = True Then
    If CmDat.Text <> vbNullString Then
        If IsDate(CmDat.Text) = True Then
            StaDa = CDate(CmDat.Text)
            If StaDa < Date Then
                SPopu "Falsches Startdatum", "Das Startdatum liegt nicht in der Zukunft", IC48_Information
                CmDat.Text = Format$(Date, "dd.mm.yyyy")
                Exit Sub
            End If
        Else
            SPopu "Falsches Datumsformat", "Die EIngabe des Startdatums hat das falsche Format", IC48_Information
            CmDat.Text = Format$(Date, "dd.mm.yyyy")
            Exit Sub
        End If
    Else
        SPopu "Kein Startdatum", "Es wurde kein Startdatum eingegeben", IC48_Information
        CmDat.Text = Format$(Date, "dd.mm.yyyy")
        Exit Sub
    End If
End If

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case TabId
Case RibTab_Adr_Dokum:
            TmStr = FSaZe
            BuStr = GlSZe
Case RibTab_Adr_Booki:
            TmStr = GlSZe
            BuStr = FSaZe
End Select

If SpNeu = True Then
    Man_SpN TmStr, BuStr, StaDa
    
    If CmCom.ListCount > 0 Then
        CmCom.Clear
        DoEvents
    End If
    
    Man_SpL
    DoEvents
    IdxNr = CmCom.ItemData(CmCom.ListIndex)
    TmStr = Man_SpD
    
    GlAdL = True
    MRast TmStr
    GlAdL = False
Else
    TmStr = FSaZe
    Select Case TabId
    Case RibTab_Adr_Dokum:
        TxZe1.Text = TmStr
        Man_SpS TmStr
    Case RibTab_Adr_Booki:
        TxZe2.Text = BuStr
        Man_SpS BuStr
    End Select
End If

If SpNeu = True Then
    If CmDat.Enabled = True Then CmDat.Enabled = False
    SpNeu = False
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpSa " & Err.Number
Resume Next

End Sub
Private Sub FSpTy()
On Error GoTo NeErr
'Wechselt des Sprechzeitentyp

Dim TabId As Long
Dim TmStr As String
Dim SpTyp As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmDat As XtremeCommandBars.CommandBarComboBox
Dim CmCb1 As XtremeCommandBars.CommandBarControl
Dim CmCb2 As XtremeCommandBars.CommandBarControl
Dim CmCb3 As XtremeCommandBars.CommandBarControl
Dim CmAkt As XtremeCommandBars.CommandBarComboBox
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmCb1 = CmBrs.FindControl(CmCb1, AD_Sprechzeit_Add, , True)
Set CmCb2 = CmBrs.FindControl(CmCb2, AD_Sprechzeit_Save, , True)
Set CmCb3 = CmBrs.FindControl(CmCb3, AD_Sprechzeit_Del, , True)
Set CmDat = CmBrs.FindControl(CmDat, AD_Sprechzeit_Datum, , True)
Set CmAkt = CmBrs.FindControl(CmAkt, AD_Sprechzeit_Typen, , True)
Set CmCom = CmBrs.FindControl(CmCom, AD_Sprechzeit_Auswa, , True)

TabId = RbTab.id

SpTyp = CmAkt.ListIndex

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Select Case SpTyp
Case 1:
    GlSpT = False
    CmDat.Enabled = False
    CmCom.Enabled = False
    CmCb1.Enabled = False
    CmCb2.Enabled = False
    CmCb3.Enabled = False
    S_SeSe 76, , , , False
    IniSetVal "TerSys", "SprTyp", 0
Case 2:
    GlSpT = True
    CmCom.Enabled = True
    CmCb1.Enabled = True
    CmCb2.Enabled = True
    CmCb3.Enabled = True
    S_SeSe 76, , , , True
    IniSetVal "TerSys", "SprTyp", -1
    TmStr = Man_SpD
End Select

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

GlAdL = True

Select Case SpTyp
Case 1: MRast
Case 2: MRast TmStr
End Select

GlAdL = False

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpTy " & Err.Number
Resume Next

End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim TabId As Long
Dim TmStr As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set Rahm8 = FM.frmRahm8

Set cmRas1 = FM.cmbRast1
Set cmRas2 = FM.cmbRast2
Set cmMaxT = FM.cmbMaxTe
Set cmMaxP = FM.cmbMaxPa
Set cmVorl = FM.cmbVorla
Set cmBuRa = FM.cmbBuRad
Set cmNoti = FM.cmbNotVa
Set cmRahm = FM.cmbKtoRa
Set cmKata = FM.cmbGbKat
Set ChOnTe = FM.chkOnlTe
Set ChDefr = FM.chkDefra
Set ChKaAu = FM.chkKaAus
Set Labl71 = FM.lblLab71
Set Labl72 = FM.lblLab72
Set Labl74 = FM.lblLab74
Set Labl80 = FM.lblLab80
Set Labl82 = FM.lblLab82

Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set CmCom = CmBrs.FindControl(CmCom, AD_Sprechzeit_Auswa, , True)

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
    Rahm4.Visible = True
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = False
Case RibTab_Adr_Dokum:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    Rahm6.Visible = True
    Rahm7.Visible = True
    Rahm8.Visible = False
    cmRas1.Visible = True
    cmRas2.Visible = False
    cmMaxT.Visible = False
    cmMaxP.Visible = False
    cmVorl.Visible = False
    cmBuRa.Visible = False
    cmNoti.Visible = True
    Labl71.Visible = False
    Labl72.Visible = False
    Labl74.Visible = False
    Labl80.Visible = False
    Labl82.Visible = True
    ChOnTe.Visible = True
    ChDefr.Visible = False
    ChKaAu.Visible = True
    Select Case GlBut
    Case RibTab_Mandanten: Rahm7.Caption = "Sprechzeiten"
    Case RibTab_Verordner: Rahm7.Caption = "Sprechzeiten"
    Case RibTab_Mitarbeit: Rahm7.Caption = "Arbeitszeiten"
    Case Else: Rahm7.Caption = "Sprechzeiten"
    End Select
Case RibTab_Adr_Eigen:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = False
    Rahm6.Visible = False
    Rahm7.Visible = False
    Rahm8.Visible = True
Case RibTab_Adr_Booki:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    Rahm5.Visible = True
    Rahm6.Visible = True
    Rahm7.Visible = True
    Rahm8.Visible = False
    cmRas1.Visible = False
    cmRas2.Visible = True
    cmMaxT.Visible = True
    cmMaxP.Visible = True
    cmVorl.Visible = True
    cmBuRa.Visible = True
    cmNoti.Visible = False
    Labl71.Visible = True
    Labl72.Visible = True
    Labl74.Visible = True
    Labl80.Visible = True
    Labl82.Visible = False
    ChOnTe.Visible = False
    ChDefr.Visible = True
    ChKaAu.Visible = False
    Rahm7.Caption = "Buchungszeiten"
End Select

GlAdL = True

If GlSpT = False Then 'starre oder flexible Sprechzeiten verwenden
    Select Case TabId
    Case RibTab_Adr_Dokum: MRast
    Case RibTab_Adr_Booki: MRast
    End Select
Else
    If GlBut = RibTab_Verordner Then
        MRast
    Else
        Select Case TabId
        Case RibTab_Adr_Dokum:
            TmStr = Man_SpD
            MRast TmStr
        Case RibTab_Adr_Booki:
            MRast
        End Select
    End If
End If

GlAdL = False

TbOld = TabId

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Set FM = frmMandant

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: GlAdL = True
            ASper True, True
            MNeu True
            GlAdL = False
Case KY_F5: frmAdrSuch.Show vbModal
Case KY_F8: FEnde True
Case KY_F11: FEnde
Case AM_Hilfe: FHilfe
Case AM_Patient_Speichern: FEnde True
Case AM_Beenden: FEnde
Case AM_Patient_Copy:
Case AM_Patient_Del:
Case AM_Patient_Such: frmAdrSuch.Show vbModal
Case AM_Programmhilfe: FHilfe
Case AM_Geburtsdatum: AGebu True
                      MKopi
                      AErAd True 'Erstellt die Anschrift im Anschriftenfeld
                      FSave
Case AD_Patient_Add: GlAdL = True
                     MNeu True
                     GlAdL = False
Case AD_Patient_Copy:
Case AD_Patient_Del:
Case AD_Patienten_Suchen: frmAdrSuch.Show vbModal
Case AD_Patienten_Save: FEnde True
Case AD_Speichern_Nroma: FEnde True
Case AD_Speichern_Close: FEnde True
Case AD_Sprechzeit_Add: FSpNe
Case AD_Sprechzeit_Del: FSpLo
Case AD_Sprechzeit_Save: FSpSa
Case AD_Sprechzeit_Datum:
Case AD_Sprechzeit_Auswa: FSpDe
Case AD_Sprechzeit_Typen: FSpTy
Case Else:
    If TolId < 0 Then
        APaSe TolId
    End If
End Select

GlToo = False

End Sub
Private Sub FClos()
On Error GoTo SaErr

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlIdi = False Then 'Idiotenmodus
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "ManForm", "FenLin", clFen.FeLin
        IniSetVal "ManForm", "FenObe", clFen.FeObn
        IniSetVal "ManForm", "FenBre", clFen.FeBre
        IniSetVal "ManForm", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case GlBut
Case RibTab_Mandanten:
    TeTit = IniGetOpt("Hilfe", 50941)
    TeMai = IniGetOpt("Hilfe", 50942)
    TeInh = IniGetOpt("Hilfe", 50943)
    TeFus = IniGetOpt("Hilfe", 50944)
Case RibTab_Mitarbeit:
    TeTit = IniGetOpt("Hilfe", 50951)
    TeMai = IniGetOpt("Hilfe", 50952)
    TeInh = IniGetOpt("Hilfe", 50953)
    TeFus = IniGetOpt("Hilfe", 50954)
Case RibTab_Verordner:
    TeTit = IniGetOpt("Hilfe", 50961)
    TeMai = IniGetOpt("Hilfe", 50962)
    TeInh = IniGetOpt("Hilfe", 50963)
    TeFus = IniGetOpt("Hilfe", 50964)
End Select

If TeTit <> vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
End If

End Sub

Private Sub FEnde(Optional ByVal MaSav As Boolean = False)
On Error GoTo SaErr

Set txVorn = Me.txtS1F04
Set txName = Me.txtS1F05

If LogLa = True Then
    If txVorn.Text = vbNullString Then
        SPopu "Benutzerdaten", "Bitte Name und Vorname ausfüllen", IC48_Information
        Exit Sub
    End If
End If

If LogLa = True Then
    If txName.Text = vbNullString Then
        SPopu "Benutzerdaten", "Bitte Name und Vorname ausfüllen", IC48_Information
        Exit Sub
    End If
End If

If CBool(IniGetVal("Vorgabe", "StMaVo")) = False Then
    IniSetVal "Vorgabe", "StMaVo", -1  'Standardmandant vorhanden
End If

If MaSav = True Then
    If GlAdS = True Then 'Speichern der Adresse erforderlich
        FInSa
        DoEvents
        MKopi
        DoEvents
        If AdAnd = True Then 'Adressenänderung
            AErAd True 'Erstellt die Anschrift im Anschriftenfeld
        End If
        DoEvents
        FSpSa
        DoEvents
        FSave
    End If
End If

DoEvents
Unload Me

If LogLa = True Then
    If CBool(GlSet(4, 9)) = True Then 'Benutzeranmeldung beim Start
        frmLogin.Show vbModal
    End If
End If

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEnde " & Err.Number
Resume Next

End Sub
Private Sub FInSa()
On Error Resume Next

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Dim TmStr As String
Dim TmIdx As Integer
Dim KaIdx As Integer
Dim FaIdx As Integer
Dim IdRa1 As Integer
Dim IdRa2 As Integer
Dim MaxTe As Integer
Dim MaxPa As Integer
Dim Vorla As Integer
Dim BuRad As Integer
Dim Notif As Integer

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set txIKNr = FM.txtIKNum
Set txGebo = FM.txtS1F13
Set TxPrax = FM.txtS2F11
Set txVorn = FM.txtS1F04
Set txName = FM.txtS1F05
Set txStra = FM.txtS1F06
Set txPost = FM.txtS1F08
Set txOrte = FM.txtS1F09
Set txTele = FM.txtS1F16
Set txFaxe = FM.txtS1F17
Set txBank = FM.txtS2F03
Set txBaLZ = FM.txtS2F04
Set txKont = FM.txtS2F05
Set txBeru = FM.txtS2F24
Set txTite = FM.txtS1F03
Set TxEmai = FM.txtS1F19
Set TxIntr = FM.txtS1F27
Set TxIBAN = FM.txtS2F33
Set TxBIC1 = FM.txtS2F34
Set TxBIC2 = FM.txtBICN2
Set txGlID = FM.txtGIDNr
Set TxLand = FM.txtS1F12
Set TxAnre = FM.txtS1F02
Set TxAbre = FM.txtS1F23
Set txBan2 = FM.txtBank2
Set txBaL2 = FM.txtBLZ02
Set txKon2 = FM.txtKont2
Set TxIBA2 = FM.txtIBAN2
Set txBeme = FM.txtS3F02
Set cmBuLa = FM.cmbBuLnd
Set cmKVBz = FM.cmbAbrBz
Set cmKant = FM.cmbKanto
Set cmFach = FM.cmbKatal
Set cmRas1 = FM.cmbRast1
Set cmRas2 = FM.cmbRast2
Set cmMaxT = FM.cmbMaxTe
Set cmMaxP = FM.cmbMaxPa
Set cmVorl = FM.cmbVorla
Set cmBuRa = FM.cmbBuRad
Set cmNoti = FM.cmbNotVa
Set ChKaAu = FM.chkKaAus
Set ChDefr = FM.chkDefra
Set ChOnTe = FM.chkOnlTe

TmStr = FSaZe

If GlMId < 0 Then
    If GlOTS = True Then 'Online-Terminbuchungs System
        If RbTab.id = RibTab_Adr_Booki Then
            IniSetVal "Adress", "ASpr2", TmStr
        Else
            IniSetVal "Adress", "ASpr1", TmStr
        End If
    Else
        IniSetVal "Adress", "ASpr1", TmStr
    End If

    If cmRas1.Text <> vbNullString Then
        IdRa1 = cmRas1.ItemData(cmRas1.ListIndex)
    Else
        IdRa1 = GlZeR 'Zeitrasterindex
    End If
    
    If cmRas2.Text <> vbNullString Then
        IdRa2 = cmRas2.ItemData(cmRas2.ListIndex)
    Else
        IdRa2 = GlZeR 'Zeitrasterindex
    End If
    
    If cmMaxT.Text <> vbNullString Then
        MaxTe = cmMaxT.ItemData(cmMaxT.ListIndex)
    Else
        MaxTe = 0
    End If
    
    If cmMaxP.Text <> vbNullString Then
        MaxPa = cmMaxP.ItemData(cmMaxP.ListIndex)
    Else
        MaxPa = 1
    End If

    If cmVorl.Text <> vbNullString Then
        Vorla = cmVorl.ItemData(cmVorl.ListIndex)
    Else
        Vorla = 2
    End If
    
    If cmBuRa.Text <> vbNullString Then
        BuRad = cmBuRa.ItemData(cmBuRa.ListIndex)
    Else
        BuRad = 12
    End If
    
    If cmNoti.Text <> vbNullString Then
        Notif = cmNoti.ItemData(cmNoti.ListIndex)
    Else
        Notif = 24
    End If
    
    If cmFach.Text <> vbNullString Then
        TmIdx = cmFach.ItemData(cmFach.ListIndex)
    Else
        TmIdx = GlFri
    End If
    
    Select Case TmIdx
    Case 1: KaIdx = 1
    Case 2: KaIdx = 20
    Case 3: KaIdx = 10
    Case 4: KaIdx = 11
    Case 5: KaIdx = 32
    Case 22: KaIdx = 30
    Case 50: KaIdx = 40
    Case Else: KaIdx = 20
    End Select
    
    IniSetVal "Adress", "AAnre", TxAnre.Text
    IniSetVal "Adress", "AVoNa", txVorn.Text
    IniSetVal "Adress", "AName", txName.Text
    IniSetVal "Adress", "AStra", txStra.Text
    IniSetVal "Adress", "APLZ", txPost.Text
    IniSetVal "Adress", "AOrt", txOrte.Text
    IniSetVal "Adress", "ATele", txTele.Text
    IniSetVal "Adress", "AFax", txFaxe.Text
    IniSetVal "Adress", "ABank", txBank.Text
    IniSetVal "Adress", "AnBLZ", txBaLZ.Text
    IniSetVal "Adress", "AnKto", txKont.Text
    IniSetVal "Adress", "AIKNr", txIKNr.Text
    IniSetVal "Adress", "ABeru", txBeru.Text
    IniSetVal "Adress", "ATite", txTite.Text
    IniSetVal "Adress", "AEmail", TxEmai.Text
    IniSetVal "Adress", "AInter", TxIntr.Text
    IniSetVal "Adress", "APraxis", TxPrax.Text
    IniSetVal "Adress", "AIBANr", TxIBAN.Text
    IniSetVal "Adress", "ABICNr", TxBIC1.Text
    IniSetVal "Adress", "AGlIDr", txGlID.Text
    IniSetVal "Adress", "ALand", TxLand.Text
    IniSetVal "Adress", "AGebo", txGebo.Text
    IniSetVal "Adress", "ABan2", txBan2.Text
    IniSetVal "Adress", "AnBL2", txBaL2.Text
    IniSetVal "Adress", "AnKt2", txKon2.Text
    IniSetVal "Adress", "AIBAN2", TxIBA2.Text
    IniSetVal "Adress", "ABeme", txBeme.Text
    IniSetVal "Adress", "ABICN2", TxBIC2.Text
    IniSetVal "Adress", "AnFar", "F" & FaIdx
    IniSetVal "Adress", "ARas1", "R" & IdRa1
    IniSetVal "Adress", "ARas2", "R" & IdRa2
    IniSetVal "Adress", "AMaxt", "R" & MaxTe
    IniSetVal "Adress", "AVorla", "R" & Vorla
    IniSetVal "Adress", "ABuRad", "R" & BuRad
    IniSetVal "Adress", "ANotif", "R" & Notif
    
    IniSetVal "System", "PVSNum", TxAbre.Text
    
    If ChKaAu.Value = xtpChecked Then
        IniSetVal "Adress", "ATmPl", True
    Else
        IniSetVal "Adress", "ATmPl", False
    End If
    If ChDefr.Value = xtpChecked Then
        IniSetVal "Adress", "ADefr", True
    Else
        IniSetVal "Adress", "ADefr", False
    End If
    If ChOnTe.Value = xtpChecked Then
        IniSetVal "Adress", "AOnTe", True
    Else
        IniSetVal "Adress", "AOnTe", False
    End If
    
    IniSetVal "Vorgabe", "StaKat", KaIdx
End If

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FOpn()
On Error GoTo ReErr
'Weißt eine Emailsignaturdatei zu

Dim TagWe As String
Dim FiNam As String
Dim NeNam As String
Dim DaNam As String
Dim DaExt As String
Dim TxSig As XtremeSuiteControls.FlatEdit
Dim CoDia As XtremeSuiteControls.CommonDialog

Set FM = frmMandant
Set TxSig = FM.txtS2F23
Set CoDia = frmMain.comDialo

Set clFil = New clsFile

With CoDia
    .CancelError = True
    .DialogStyle = 1
    .DialogTitle = "Bitte Name und Ordner der Datei angeben"
    .DefaultExt = "*.txn"
    .Filter = "Newslettervorlage (.txn)|*.txn|Joint Photographic Experts Group (.jpg)|*.jpg|Portable Network Graphics (.png)|*.png|Portable Network Graphics (.png)|*.png|Alle Dateien (*.*)|*.*"
    .InitDir = GlVor
    .FileName = vbNullString
    .ShowOpen
    FiNam = .FileName
    If .FileTitle = vbNullString Then
        Set CoDia = Nothing
        Set clFil = Nothing
        Exit Sub
    End If
End With

If Not IsNull(FiNam) And Not FiNam = vbNullString Then
    With clFil
        .FilPfa FiNam
        DaNam = .DaNam
        DaExt = .DaExt
    End With
    If LCase(DaExt) <> "txn" Then
        NeNam = GlVor & DaNam
        With clFil
            .DaCop = FiNam & ";" & NeNam & vbNullChar
            If .FilCop(1) = False Then
                SPopu "Datei schreibgeschützt", "Die Datei kann nicht kopiert werden", IC48_Warning
            End If
        End With
    End If
    TxSig.Text = DaNam
    TagWe = Mid$(TxSig.Tag, 2, Len(TxSig.Tag) - 1)
    TxSig.Tag = "1" & TagWe
    GlAdS = True
End If

Set clFil = Nothing

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpn " & Err.Number
Resume Next

End Sub
Private Sub FRast()
On Error GoTo AnErr

Set FM = frmMandant

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

GlAdS = True
GlAdL = True
MRast

GlAdL = False

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRast " & Err.Number
Resume Next

End Sub
Private Sub FRcht()
On Error GoTo SaErr

Dim Recht As String
Dim AktZa As Integer
Dim RpRec As XtremeReportControl.ReportRecord
Dim RpRcs As XtremeReportControl.ReportRecords
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmMandant
Set RpCo5 = FM.repCont5
Set RpRcs = RpCo5.Records

For Each RpRec In RpRcs
    If RpRec.Item(0).Checked = True Then
        Recht = Recht & "1"
    Else
        Recht = Recht & "0"
    End If
Next RpRec

If Recht = vbNullString Then
    Recht = GlStR 'Rechtestring
End If

If IsNumeric(Recht) = False Then
    Recht = GlStR 'Rechtestring
End If

If Len(Recht) <> GlZaR Then 'Rechteanzahl
    Recht = GlStR 'Rechtestring
End If
        
For AktZa = 0 To GlZaR - 1 'Rechteanzahl
    If Mid$(Recht, AktZa + 1, 1) = "1" Then
        GlRch(0, AktZa) = 1
    Else
        GlRch(0, AktZa) = 0
    End If
Next AktZa

DoEvents
SRecht

Set RpRcs = Nothing
Set RpCo5 = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRcht " & Err.Number
Resume Next

End Sub
Private Sub FSave(Optional ByVal SaFra As Boolean = False)
On Error GoTo SaErr
'Überprüft, ob der Eintrag geändert wurde und speichert dieses ab

Dim RowNr As Long
Dim SuStr As String
Dim RetBo As Boolean
Dim DoSav As Boolean
Dim ErSta As Boolean
Dim IdRa1 As Integer
Dim IdRa2 As Integer
Dim MaxTe As Integer
Dim MaxPa As Integer
Dim Vorla As Integer
Dim BuRad As Integer
Dim Notif As Integer
Dim Frage As Integer
Dim Mld1, Mld2, Tit1 As String

Dim FeKur As XtremeSuiteControls.FlatEdit
Dim RpCo2 As XtremeReportControl.ReportControl
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim CmBrM As XtremeCommandBars.CommandBars
Dim CmBrs As XtremeCommandBars.CommandBars
Dim MsBar As XtremeCommandBars.MessageBar
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set FeKur = FM.txtS1F11
Set TxDum = FM.txtDummy
Set cmRas1 = FM.cmbRast1
Set cmRas2 = FM.cmbRast2
Set cmMaxT = FM.cmbMaxTe
Set cmMaxP = FM.cmbMaxPa
Set cmVorl = FM.cmbVorla
Set cmBuRa = FM.cmbBuRad
Set cmNoti = FM.cmbNotVa

Set CmBrM = frmMain.comBar01
Set MsBar = CmBrM.MessageBar
Set RpCo2 = frmMain.repCont2
Set CaCol = frmMain.calCont1
Set RpSel = RpCo2.SelectedRows

Tit1 = "Datensatz Speichern"
Mld1 = "Soll der Datensatz gespeichert werden?"
Mld2 = "Dieser Mandant existiert bereits. Soll diese trotzdem gespeichert werden?"

If TxDum.Text <> vbNullString Then
    If TxDum.Text = -1 Then
        If GlMaV = True Then 'Mandanten vorhanden
            Exit Sub
        Else
            GlMId = -2
            GlAdN = True
            GlAdS = True
            GlAdG = CreateID("M")
        End If
    End If
ElseIf GlMaV = False Then 'Mandanten vorhanden
    GlMId = -2
    GlAdN = True
    GlAdS = True
    GlAdG = CreateID("M")
End If

If GlMId = -2 Then
    ErSta = True 'Erststart
End If

If GlAdS = True Then
    If cmRas1.Text <> vbNullString Then 'Zeitrasterindex
        IdRa1 = cmRas1.ItemData(cmRas1.ListIndex)
    Else
        IdRa1 = GlZeR 'Zeitrasterindex
    End If
    
    If cmRas2.Text <> vbNullString Then 'Zeitrasterindex
        IdRa2 = cmRas2.ItemData(cmRas2.ListIndex)
    Else
        IdRa2 = GlZeR 'Zeitrasterindex
    End If
    
    If cmMaxT.Text <> vbNullString Then
        MaxTe = cmMaxT.ItemData(cmMaxT.ListIndex)
    Else
        MaxTe = 0
    End If
    
    If cmMaxP.Text <> vbNullString Then
        MaxPa = cmMaxP.ItemData(cmMaxP.ListIndex)
    Else
        MaxPa = 1
    End If
    
    If cmVorl.Text <> vbNullString Then
        Vorla = cmVorl.ItemData(cmVorl.ListIndex)
    Else
        Vorla = 2
    End If
    
    If cmBuRa.Text <> vbNullString Then
        BuRad = cmBuRa.ItemData(cmBuRa.ListIndex)
    Else
        BuRad = 3
    End If
    
    If cmNoti.Text <> vbNullString Then
        Notif = cmNoti.ItemData(cmNoti.ListIndex)
    Else
        Notif = 24
    End If

    If FeKur.Text <> vbNullString Then
        SuStr = FeKur.Text
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
            If GlAdN = True Then
                If Adr_Dop(SuStr, 1) = True Then
                    Frage = WindowMess(Mld2, Dial1, Tit1, FM.hwnd)
                    If Frage = 6 Then
                        DoSav = True
                    End If
                Else
                    DoSav = True
                End If
                If DoSav = True Then
                    If ErSta = True Then
                        GlMId = -3
                        RetBo = Man_San()
                        FRcht
                        GlMId = -2
                        GlAdG = CreateID("M")
                    End If
                    If Man_San = True Then
                        DoEvents
                        S_Ary1
                        DoEvents
                        S_Ary3
                        DoEvents
                        GlAdr = S_AdGui(GlAdG, "ID0")
                        GlMId = GlAdr
                        DoEvents
                        Unload Me
                        DoEvents
                        DBWaKl
                        DoEvents
                        S_AdSpl
                        S_ReSpl
                        S_AbSpl
                        SUpAd
                        If GlMiV = True Then
                            If MsBar.Visible = True Then
                                MsBar.Visible = False
                            End If
                        End If
                    Else
                        GlAdN = False
                    End If
                Else
                    GlAdN = False
                End If
            Else
                Man_Sav
                S_Ary1
                DoEvents
                S_Ary3
                DoEvents
                SUpAd False
            End If
        End If
    End If
    CaCol.RedrawControl
    GlTDa = vbNullString 'Wichtig für Textverarbeitung
End If

GlAdS = False

Set RpCo2 = Nothing
Set RpSel = Nothing
Set CmBrM = Nothing
Set CmBrs = Nothing
Set RbBar = Nothing
Set RbTab = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Function FSaZe() As String
On Error GoTo SaErr
'Überprüft, ob der Eintrag geändert wurde und gibt eine Zeichenkette zurück

Dim TabId As Long
Dim TmStr As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim TxZe1 As XtremeSuiteControls.FlatEdit
Dim TxZe2 As XtremeSuiteControls.FlatEdit

Set FM = frmMandant
Set TxZe1 = FM.txtZeit1
Set TxZe2 = FM.txtZeit2
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set ChS01 = FM.chkBox01
Set ChS02 = FM.chkBox02
Set ChS03 = FM.chkBox03
Set ChS04 = FM.chkBox04
Set ChS05 = FM.chkBox05
Set ChS06 = FM.chkBox06
Set ChS07 = FM.chkBox07
Set ChS08 = FM.chkBox08
Set ChS09 = FM.chkBox09
Set ChS10 = FM.chkBox10
Set ChS11 = FM.chkBox11
Set ChS12 = FM.chkBox12
Set ChS13 = FM.chkBox13
Set ChS14 = FM.chkBox14

Set cmbS01 = FM.cmbSpZ01
Set cmbS02 = FM.cmbSpZ02
Set cmbS03 = FM.cmbSpZ03
Set cmbS04 = FM.cmbSpZ04
Set cmbS05 = FM.cmbSpZ05
Set cmbS06 = FM.cmbSpZ06
Set cmbS07 = FM.cmbSpZ07
Set cmbS08 = FM.cmbSpZ08
Set cmbS09 = FM.cmbSpZ09
Set cmbS10 = FM.cmbSpZ10
Set cmbS11 = FM.cmbSpZ11
Set cmbS12 = FM.cmbSpZ12
Set cmbS13 = FM.cmbSpZ13
Set cmbS14 = FM.cmbSpZ14
Set cmbS15 = FM.cmbSpZ15
Set cmbS16 = FM.cmbSpZ16
Set cmbS17 = FM.cmbSpZ17
Set cmbS18 = FM.cmbSpZ18
Set cmbS19 = FM.cmbSpZ19
Set cmbS20 = FM.cmbSpZ20
Set cmbS21 = FM.cmbSpZ21
Set cmbS22 = FM.cmbSpZ22
Set cmbS23 = FM.cmbSpZ23
Set cmbS24 = FM.cmbSpZ24
Set cmbS25 = FM.cmbSpZ25
Set cmbS26 = FM.cmbSpZ26
Set cmbS27 = FM.cmbSpZ27
Set cmbS28 = FM.cmbSpZ28

TabId = RbTab.id

If ChS01.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS01.Text & "_"
TmStr = TmStr & cmbS02.Text

If ChS08.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS03.Text & "_"
TmStr = TmStr & cmbS04.Text

If ChS02.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS05.Text & "_"
TmStr = TmStr & cmbS06.Text

If ChS09.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS07.Text & "_"
TmStr = TmStr & cmbS08.Text

If ChS03.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS09.Text & "_"
TmStr = TmStr & cmbS10.Text

If ChS10.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS11.Text & "_"
TmStr = TmStr & cmbS12.Text

If ChS04.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS13.Text & "_"
TmStr = TmStr & cmbS14.Text

If ChS11.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS15.Text & "_"
TmStr = TmStr & cmbS16.Text

If ChS05.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS17.Text & "_"
TmStr = TmStr & cmbS18.Text

If ChS12.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS19.Text & "_"
TmStr = TmStr & cmbS20.Text

If ChS06.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS21.Text & "_"
TmStr = TmStr & cmbS22.Text

If ChS13.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS23.Text & "_"
TmStr = TmStr & cmbS24.Text

If ChS07.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS25.Text & "_"
TmStr = TmStr & cmbS26.Text

If ChS14.Value = xtpChecked Then
    TmStr = TmStr & "A"
Else
    TmStr = TmStr & "B"
End If

TmStr = TmStr & cmbS27.Text & "_"
TmStr = TmStr & cmbS28.Text

If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
    Select Case TabId
    Case RibTab_Adr_Dokum:
        TxZe1.Text = TmStr
        TagWe = Mid$(TxZe1.Tag, 2, Len(TxZe1.Tag) - 1)
        TxZe1.Tag = "1" & TagWe
    Case RibTab_Adr_Booki:
        TxZe2.Text = TmStr
        TagWe = Mid$(TxZe2.Tag, 2, Len(TxZe2.Tag) - 1)
        TxZe2.Tag = "1" & TagWe
    End Select
End If

FSaZe = TmStr

Set CmBrs = Nothing
Set RbBar = Nothing
Set RbTab = Nothing

Exit Function

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSaZe " & Err.Number
Resume Next

End Function
Private Sub chkOpti2_Click()

TagWe = Mid$(Me.chkOpti2.Tag, 2, Len(Me.chkOpti2.Tag) - 1)

If GlAdL = False Then
    Me.chkOpti2.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

FClos

GlAdS = False

Set frmMandant = Nothing

End Sub
Private Sub txtS1F23_Change()

TagWe = Mid$(Me.txtS1F23.Tag, 2, Len(Me.txtS1F23.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F23.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F23_GotFocus()
    Me.txtS1F23.SelStart = 0
    Me.txtS1F23.SelLength = Len(Me.txtS1F23.Text)
End Sub

Private Sub txtS1F23_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F30_Change()

TagWe = Mid$(Me.txtS1F30.Tag, 2, Len(Me.txtS1F30.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F30.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F30_GotFocus()
    Me.txtS1F30.SelStart = 0
    Me.txtS1F30.SelLength = Len(Me.txtS1F30.Text)
End Sub

Private Sub txtS1F30_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F30_Validate(Cancel As Boolean)
     If (Not txtS1F30.isValid) Then Cancel = True
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
End Sub
Private Sub txtS1F37_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F37_LostFocus()
On Error Resume Next

Set FS = frmPasswort

If Me.txtS1F37.Text <> vbNullString Then
    FS.PaStr = Me.txtS1F37.Text
    FS.Show
End If

End Sub
Private Sub txtS1F38_Change()

TagWe = Mid$(Me.txtS1F38.Tag, 2, Len(Me.txtS1F38.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F38.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F38_GotFocus()
    Me.txtS1F38.SelStart = 0
    Me.txtS1F38.SelLength = Len(Me.txtS1F38.Text)
End Sub

Private Sub txtS1F38_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F39_Change()

TagWe = Mid$(Me.txtS1F39.Tag, 2, Len(Me.txtS1F39.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F39.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F39_GotFocus()
    Me.txtS1F39.SelStart = 0
    Me.txtS1F39.SelLength = Len(Me.txtS1F39.Text)
End Sub

Private Sub txtS1F39_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F11_GotFocus()
    Me.txtS2F11.SelStart = 0
    Me.txtS2F11.SelLength = Len(Me.txtS2F11.Text)
End Sub

Private Sub txtS2F11_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F23_Change()

TagWe = Mid$(Me.txtS2F23.Tag, 2, Len(Me.txtS2F23.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F23.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F23_GotFocus()
    Me.txtS2F23.SelStart = Len(Me.txtS2F23.Text)
    Me.txtS2F23.SelLength = 0
End Sub
Private Sub txtS2F23_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
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
End Sub

Private Sub txtS2F33_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtIKNum_Change()

TagWe = Mid$(Me.txtIKNum.Tag, 2, Len(Me.txtIKNum.Tag) - 1)

If GlAdL = False Then
    Me.txtIKNum.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F01_Change()

TagWe = Mid$(Me.txtS1F01.Tag, 2, Len(Me.txtS1F01.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F01.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F02_Change()

TagWe = Mid$(Me.txtS1F02.Tag, 2, Len(Me.txtS1F02.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F02.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F02_Click()

TagWe = Mid$(Me.txtS1F02.Tag, 2, Len(Me.txtS1F02.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F02.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F02_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F03_Change()
On Error Resume Next

TagWe = Mid$(Me.txtS1F03.Tag, 2, Len(Me.txtS1F03.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F03.Tag = "1" & TagWe
    GlAdS = True
    AdBrf , True
End If

End Sub
Private Sub txtS1F04_Change()

TagWe = Mid$(Me.txtS1F04.Tag, 2, Len(Me.txtS1F04.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F04.Tag = "1" & TagWe
    GlAdS = True
    AdAnd = True 'Adressenänderung
End If

End Sub
Private Sub txtS1F05_Change()

TagWe = Mid$(Me.txtS1F05.Tag, 2, Len(Me.txtS1F05.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F05.Tag = "1" & TagWe
    GlAdS = True
    AdAnd = True 'Adressenänderung
End If

End Sub
Private Sub txtS1F06_Change()

TagWe = Mid$(Me.txtS1F06.Tag, 2, Len(Me.txtS1F06.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F06.Tag = "1" & TagWe
    GlAdS = True
    AdAnd = True 'Adressenänderung
End If

End Sub
Private Sub txtS1F08_Change()

TagWe = Mid$(Me.txtS1F08.Tag, 2, Len(Me.txtS1F08.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F08.Tag = "1" & TagWe
    GlAdS = True
    AdAnd = True 'Adressenänderung
End If

End Sub

Private Sub txtS1F09_Change()

TagWe = Mid$(Me.txtS1F09.Tag, 2, Len(Me.txtS1F09.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F09.Tag = "1" & TagWe
    GlAdS = True
    AdAnd = True 'Adressenänderung
End If

End Sub
Private Sub txtS1F11_Change()

TagWe = Mid$(Me.txtS1F11.Tag, 2, Len(Me.txtS1F11.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F11.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F12_Click()

TagWe = Mid$(Me.txtS1F12.Tag, 2, Len(Me.txtS1F12.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F12.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F13_Change()

TagWe = Mid$(Me.txtS1F13.Tag, 2, Len(Me.txtS1F13.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F13.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F15_Change()

TagWe = Mid$(Me.txtS1F15.Tag, 2, Len(Me.txtS1F15.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F15.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F16_Change()

TagWe = Mid$(Me.txtS1F16.Tag, 2, Len(Me.txtS1F16.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F16.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F17_Change()

TagWe = Mid$(Me.txtS1F17.Tag, 2, Len(Me.txtS1F17.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F17.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS1F19_Change()

TagWe = Mid$(Me.txtS1F19.Tag, 2, Len(Me.txtS1F19.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F19.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F22_Change()

TagWe = Mid$(Me.txtS1F22.Tag, 2, Len(Me.txtS1F22.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F22.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS1F22_GotFocus()
    Me.txtS1F22.SelStart = 0
    Me.txtS1F22.SelLength = Len(Me.txtS1F22.Text)
End Sub

Private Sub txtS1F22_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F27_Change()

TagWe = Mid$(Me.txtS1F27.Tag, 2, Len(Me.txtS1F27.Tag) - 1)

If GlAdL = False Then
    Me.txtS1F27.Tag = "1" & TagWe
    GlAdS = True
End If

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
End Sub

Private Sub txtS2F03_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F04_Change()

TagWe = Mid$(Me.txtS2F04.Tag, 2, Len(Me.txtS2F04.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F04.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F04_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS2F04.SelLength = 0
    Case vbKeyDown: Me.txtS1F39.SetFocus
    Case vbKeyUp: Me.txtS2F03.SetFocus
    End Select
End Sub
Private Sub txtS2F05_Change()

TagWe = Mid$(Me.txtS2F05.Tag, 2, Len(Me.txtS2F05.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F05.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F11_Change()

TagWe = Mid$(Me.txtS2F11.Tag, 2, Len(Me.txtS2F11.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F11.Tag = "1" & TagWe
    GlAdS = True
    AdAnd = True 'Adressenänderung
End If

End Sub

Private Sub txtS2F12_Change()

TagWe = Mid$(Me.txtS2F12.Tag, 2, Len(Me.txtS2F12.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F12.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F13_Change()

TagWe = Mid$(Me.txtS2F13.Tag, 2, Len(Me.txtS2F13.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F13.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F14_Change()

TagWe = Mid$(Me.txtS2F14.Tag, 2, Len(Me.txtS2F14.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F14.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F15_Change()

TagWe = Mid$(Me.txtS2F15.Tag, 2, Len(Me.txtS2F15.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F15.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F16_Change()

TagWe = Mid$(Me.txtS2F16.Tag, 2, Len(Me.txtS2F16.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F16.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F18_Change()

TagWe = Mid$(Me.txtS2F18.Tag, 2, Len(Me.txtS2F18.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F18.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F19_Change()

TagWe = Mid$(Me.txtS2F19.Tag, 2, Len(Me.txtS2F19.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F19.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F20_Change()

TagWe = Mid$(Me.txtS2F20.Tag, 2, Len(Me.txtS2F20.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F20.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F22_Change()

TagWe = Mid$(Me.txtS2F22.Tag, 2, Len(Me.txtS2F22.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F22.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F24_Change()

TagWe = Mid$(Me.txtS2F24.Tag, 2, Len(Me.txtS2F24.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F24.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F24_GotFocus()
    Me.txtS2F24.SelStart = 0
    Me.txtS2F24.SelLength = Len(Me.txtS2F24.Text)
End Sub
Private Sub txtS2F24_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F04_GotFocus()
    Me.txtS2F04.SelStart = 0
    Me.txtS2F04.SelLength = Len(Me.txtS2F04.Text)
End Sub
Private Sub txtS2F04_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F19_GotFocus()
    Me.txtS1F19.SelStart = 0
    Me.txtS1F19.SelLength = Len(Me.txtS1F19.Text)
End Sub


Private Sub txtS1F19_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F17_GotFocus()
    Me.txtS1F17.SelStart = Len(Me.txtS1F17.Text)
    Me.txtS1F17.SelLength = 0
End Sub
Private Sub txtS1F17_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtIKNum_GotFocus()
    Me.txtIKNum.SelStart = 0
    Me.txtIKNum.SelLength = Len(Me.txtIKNum.Text)
End Sub
Private Sub txtIKNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F27_GotFocus()
    Me.txtS1F27.SelStart = 0
    Me.txtS1F27.SelLength = Len(Me.txtS1F27.Text)
End Sub

Private Sub txtS1F27_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
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
Private Sub txtS1F05_GotFocus()
    Me.txtS1F05.SelStart = 0
    Me.txtS1F05.SelLength = Len(Me.txtS1F05.Text)
End Sub
Private Sub txtS1F05_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
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
Private Sub txtS1F08_GotFocus()
    Me.txtS1F08.SelStart = 0
    Me.txtS1F08.SelLength = Len(Me.txtS1F08.Text)
End Sub
Private Sub txtS1F01_GotFocus()
    Me.txtS1F01.SelStart = 0
    Me.txtS1F01.SelLength = Len(Me.txtS1F01.Text)
End Sub

Private Sub txtS1F01_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS1F06_GotFocus()
    Me.txtS1F06.SelStart = 0
    Me.txtS1F06.SelLength = Len(Me.txtS1F06.Text)
End Sub
Private Sub txtS1F06_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F16_GotFocus()
    Me.txtS1F16.SelStart = Len(Me.txtS1F16.Text)
    Me.txtS1F16.SelLength = 0
End Sub
Private Sub txtS1F16_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F03_GotFocus()
    Me.txtS1F03.SelStart = 0
    Me.txtS1F03.SelLength = Len(Me.txtS1F03.Text)
End Sub
Private Sub txtS1F03_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS1F04_GotFocus()
    Me.txtS1F04.SelStart = 0
    Me.txtS1F04.SelLength = Len(Me.txtS1F04.Text)
End Sub
Private Sub txtS1F04_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS2F25_Change()

TagWe = Mid$(Me.txtS2F25.Tag, 2, Len(Me.txtS2F25.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F25.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub
Private Sub txtS2F27_Change()

TagWe = Mid$(Me.txtS2F27.Tag, 2, Len(Me.txtS2F27.Tag) - 1)

If GlAdL = False Then
    Me.txtS2F27.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtS2F33_LostFocus()
On Error Resume Next

Dim TmStr As String

If GlAdL = False Then 'Formular wird geladen
    If Me.txtS2F33.Text <> vbNullString Then
        TmStr = Me.txtS2F33.Text
        Me.txtS2F33.Text = SNaFi(TmStr, True)
        If Len(TmStr) <> 22 Then
            SPopu "IBAN ist falsch", "Die IBAN hat die falsche Länge", IC48_Forbidden
        End If
    End If
End If

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
End Sub

Private Sub txtS2F34_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS2F34_LostFocus()
On Error Resume Next

Dim TmStr As String

If GlAdL = False Then
    If Me.txtS2F34.Text <> vbNullString Then
        TmStr = Me.txtS2F34.Text
        Me.txtS2F34.Text = SNaUm(TmStr)
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
    GlAdS = True
End If

End Sub

Private Sub txtS4F01_Change()

TagWe = Mid$(Me.txtS4F01.Tag, 2, Len(Me.txtS4F01.Tag) - 1)

If GlAdL = False Then
    Me.txtS4F01.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub comBar01_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlAdL = False Then
    RetWe = SendMessage(Me.hwnd, WM_SETREDRAW, False, 0&)
    MPosi
    RetWe = SendMessage(Me.hwnd, WM_SETREDRAW, True, 0&)
    RetWe = GetClientRect(Me.hwnd, ClRe)
    RetWe = RedrawWindow(Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
    GlRzA = True
End If

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

Private Sub chkBox01_Click()

Set ChS01 = Me.chkBox01
Set cmbS01 = Me.cmbSpZ01
Set cmbS02 = Me.cmbSpZ02

If ChS01.Value = xtpChecked Then
    cmbS01.Enabled = True
    cmbS02.Enabled = True
    cmbS01.Visible = True
    cmbS02.Visible = True
Else
    cmbS01.Enabled = False
    cmbS02.Enabled = False
    cmbS01.Visible = False
    cmbS02.Visible = False
End If

GlAdS = True

End Sub
Private Sub chkBox02_Click()

Set ChS02 = Me.chkBox02
Set cmbS05 = Me.cmbSpZ05
Set cmbS06 = Me.cmbSpZ06

If ChS02.Value = xtpChecked Then
    cmbS05.Enabled = True
    cmbS06.Enabled = True
    cmbS05.Visible = True
    cmbS06.Visible = True
Else
    cmbS05.Enabled = False
    cmbS06.Enabled = False
    cmbS05.Visible = False
    cmbS06.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox03_Click()

Set ChS03 = Me.chkBox03
Set cmbS09 = Me.cmbSpZ09
Set cmbS10 = Me.cmbSpZ10

If ChS03.Value = xtpChecked Then
    cmbS09.Enabled = True
    cmbS10.Enabled = True
    cmbS09.Visible = True
    cmbS10.Visible = True
Else
    cmbS09.Enabled = False
    cmbS10.Enabled = False
    cmbS09.Visible = False
    cmbS10.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox04_Click()

Set ChS04 = Me.chkBox04
Set cmbS13 = Me.cmbSpZ13
Set cmbS14 = Me.cmbSpZ14

If ChS04.Value = xtpChecked Then
    cmbS13.Enabled = True
    cmbS14.Enabled = True
    cmbS13.Visible = True
    cmbS14.Visible = True
Else
    cmbS13.Enabled = False
    cmbS14.Enabled = False
    cmbS13.Visible = False
    cmbS14.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox05_Click()

Set ChS05 = Me.chkBox05
Set cmbS17 = Me.cmbSpZ17
Set cmbS18 = Me.cmbSpZ18

If ChS05.Value = xtpChecked Then
    cmbS17.Enabled = True
    cmbS18.Enabled = True
    cmbS17.Visible = True
    cmbS18.Visible = True
Else
    cmbS17.Enabled = False
    cmbS18.Enabled = False
    cmbS17.Visible = False
    cmbS18.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox06_Click()

Set ChS06 = Me.chkBox06
Set cmbS21 = Me.cmbSpZ21
Set cmbS22 = Me.cmbSpZ22

If ChS06.Value = xtpChecked Then
    cmbS21.Enabled = True
    cmbS22.Enabled = True
    cmbS21.Visible = True
    cmbS22.Visible = True
Else
    cmbS21.Enabled = False
    cmbS22.Enabled = False
    cmbS21.Visible = False
    cmbS22.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox07_Click()

Set ChS07 = Me.chkBox07
Set cmbS25 = Me.cmbSpZ25
Set cmbS26 = Me.cmbSpZ26

If ChS07.Value = xtpChecked Then
    cmbS25.Enabled = True
    cmbS26.Enabled = True
    cmbS25.Visible = True
    cmbS26.Visible = True
Else
    cmbS25.Enabled = False
    cmbS26.Enabled = False
    cmbS25.Visible = False
    cmbS26.Visible = False
End If

GlAdS = True

End Sub
Private Sub chkBox08_Click()

Set ChS08 = Me.chkBox08
Set cmbS03 = Me.cmbSpZ03
Set cmbS04 = Me.cmbSpZ04

If ChS08.Value = xtpChecked Then
    cmbS03.Enabled = True
    cmbS04.Enabled = True
    cmbS03.Visible = True
    cmbS04.Visible = True
Else
    cmbS03.Enabled = False
    cmbS04.Enabled = False
    cmbS03.Visible = False
    cmbS04.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox09_Click()

Set ChS09 = Me.chkBox09
Set cmbS07 = Me.cmbSpZ07
Set cmbS08 = Me.cmbSpZ08

If ChS09.Value = xtpChecked Then
    cmbS07.Enabled = True
    cmbS08.Enabled = True
    cmbS07.Visible = True
    cmbS08.Visible = True
Else
    cmbS07.Enabled = False
    cmbS08.Enabled = False
    cmbS07.Visible = False
    cmbS08.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox10_Click()

Set ChS10 = Me.chkBox10
Set cmbS11 = Me.cmbSpZ11
Set cmbS12 = Me.cmbSpZ12

If ChS10.Value = xtpChecked Then
    cmbS11.Enabled = True
    cmbS12.Enabled = True
    cmbS11.Visible = True
    cmbS12.Visible = True
Else
    cmbS11.Enabled = False
    cmbS12.Enabled = False
    cmbS11.Visible = False
    cmbS12.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox11_Click()

Set ChS11 = Me.chkBox11
Set cmbS15 = Me.cmbSpZ15
Set cmbS16 = Me.cmbSpZ16

If ChS11.Value = xtpChecked Then
    cmbS15.Enabled = True
    cmbS16.Enabled = True
    cmbS15.Visible = True
    cmbS16.Visible = True
Else
    cmbS15.Enabled = False
    cmbS16.Enabled = False
    cmbS15.Visible = False
    cmbS16.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox12_Click()

Set ChS12 = Me.chkBox12
Set cmbS19 = Me.cmbSpZ19
Set cmbS20 = Me.cmbSpZ20

If ChS12.Value = xtpChecked Then
    cmbS19.Enabled = True
    cmbS20.Enabled = True
    cmbS19.Visible = True
    cmbS20.Visible = True
Else
    cmbS19.Enabled = False
    cmbS20.Enabled = False
    cmbS19.Visible = False
    cmbS20.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox13_Click()

Set ChS13 = Me.chkBox13
Set cmbS23 = Me.cmbSpZ23
Set cmbS24 = Me.cmbSpZ24

If ChS13.Value = xtpChecked Then
    cmbS23.Enabled = True
    cmbS24.Enabled = True
    cmbS23.Visible = True
    cmbS24.Visible = True
Else
    cmbS23.Enabled = False
    cmbS24.Enabled = False
    cmbS23.Visible = False
    cmbS24.Visible = False
End If

GlAdS = True

End Sub

Private Sub chkBox14_Click()

Set ChS14 = Me.chkBox14
Set cmbS27 = Me.cmbSpZ27
Set cmbS28 = Me.cmbSpZ28

If ChS14.Value = xtpChecked Then
    cmbS27.Enabled = True
    cmbS28.Enabled = True
    cmbS27.Visible = True
    cmbS28.Visible = True
Else
    cmbS27.Enabled = False
    cmbS28.Enabled = False
    cmbS27.Visible = False
    cmbS28.Visible = False
End If

GlAdS = True

End Sub
Private Sub txtBank2_Change()

TagWe = Mid$(Me.txtBank2.Tag, 2, Len(Me.txtBank2.Tag) - 1)

If GlAdL = False Then
    Me.txtBank2.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtBank2_GotFocus()
    Me.txtBank2.SelStart = 0
    Me.txtBank2.SelLength = Len(Me.txtBank2.Text)
End Sub

Private Sub txtBank2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBLZ02_Change()

TagWe = Mid$(Me.txtBLZ02.Tag, 2, Len(Me.txtBLZ02.Tag) - 1)

If GlAdL = False Then
    Me.txtBLZ02.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtBLZ02_GotFocus()
    Me.txtBLZ02.SelStart = 0
    Me.txtBLZ02.SelLength = Len(Me.txtBLZ02.Text)
End Sub

Private Sub txtBLZ02_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtIBAN2_Change()

TagWe = Mid$(Me.txtIBAN2.Tag, 2, Len(Me.txtIBAN2.Tag) - 1)

If GlAdL = False Then
    Me.txtIBAN2.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtIBAN2_GotFocus()
    Me.txtIBAN2.SelStart = 0
    Me.txtIBAN2.SelLength = Len(Me.txtIBAN2.Text)
End Sub

Private Sub txtIBAN2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKont2_Change()

TagWe = Mid$(Me.txtKont2.Tag, 2, Len(Me.txtKont2.Tag) - 1)

If GlAdL = False Then
    Me.txtKont2.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtKont2_GotFocus()
    Me.txtKont2.SelStart = 0
    Me.txtKont2.SelLength = Len(Me.txtKont2.Text)
End Sub


Private Sub txtKont2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub


Private Sub cmbSpZ01_GotFocus()
    Me.cmbSpZ01.SelStart = 0
    Me.cmbSpZ01.SelLength = Len(Me.cmbSpZ01.Text)
End Sub

Private Sub cmbSpZ01_LostFocus()
    If InStrRev(Me.cmbSpZ01.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ01.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ02_LostFocus()
    If InStrRev(Me.cmbSpZ02.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ02.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ03_LostFocus()
    If InStrRev(Me.cmbSpZ03.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ03.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ04_LostFocus()
    If InStrRev(Me.cmbSpZ04.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ04.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ05_LostFocus()
    If InStrRev(Me.cmbSpZ05.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ05.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ06_LostFocus()
    If InStrRev(Me.cmbSpZ06.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ06.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ07_LostFocus()
    If InStrRev(Me.cmbSpZ07.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ07.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ08_LostFocus()
    If InStrRev(Me.cmbSpZ08.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ08.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ09_LostFocus()
    If InStrRev(Me.cmbSpZ09.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ09.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ10_LostFocus()
    If InStrRev(Me.cmbSpZ10.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ10.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ11_LostFocus()
    If InStrRev(Me.cmbSpZ11.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ11.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ12_LostFocus()
    If InStrRev(Me.cmbSpZ12.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ12.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ13_LostFocus()
    If InStrRev(Me.cmbSpZ13.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ13.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ14_LostFocus()
    If InStrRev(Me.cmbSpZ14.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ14.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ15_LostFocus()
    If InStrRev(Me.cmbSpZ15.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ15.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ16_LostFocus()
    If InStrRev(Me.cmbSpZ16.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ16.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ17_LostFocus()
    If InStrRev(Me.cmbSpZ17.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ17.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ18_LostFocus()
    If InStrRev(Me.cmbSpZ18.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ18.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ19_LostFocus()
    If InStrRev(Me.cmbSpZ19.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ19.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ20_LostFocus()
    If InStrRev(Me.cmbSpZ20.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ20.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ21_LostFocus()
    If InStrRev(Me.cmbSpZ21.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ21.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ22_LostFocus()
    If InStrRev(Me.cmbSpZ22.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ22.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ23_LostFocus()
    If InStrRev(Me.cmbSpZ23.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ23.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ24_LostFocus()
    If InStrRev(Me.cmbSpZ24.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ24.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ25_LostFocus()
    If InStrRev(Me.cmbSpZ25.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ25.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ26_LostFocus()
    If InStrRev(Me.cmbSpZ26.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ26.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ27_LostFocus()
    If InStrRev(Me.cmbSpZ27.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ27.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ28_LostFocus()
    If InStrRev(Me.cmbSpZ28.Text, "_", -1, 1) > 0 Then
        Me.cmbSpZ28.Text = "00:00"
    End If
End Sub
Private Sub cmbSpZ02_GotFocus()
    Me.cmbSpZ02.SelStart = 0
    Me.cmbSpZ02.SelLength = Len(Me.cmbSpZ02.Text)
End Sub
Private Sub cmbSpZ03_GotFocus()
    Me.cmbSpZ03.SelStart = 0
    Me.cmbSpZ03.SelLength = Len(Me.cmbSpZ03.Text)
End Sub
Private Sub cmbSpZ04_GotFocus()
    Me.cmbSpZ04.SelStart = 0
    Me.cmbSpZ04.SelLength = Len(Me.cmbSpZ04.Text)
End Sub
Private Sub cmbSpZ05_GotFocus()
    Me.cmbSpZ05.SelStart = 0
    Me.cmbSpZ05.SelLength = Len(Me.cmbSpZ05.Text)
End Sub
Private Sub cmbSpZ06_GotFocus()
    Me.cmbSpZ06.SelStart = 0
    Me.cmbSpZ06.SelLength = Len(Me.cmbSpZ06.Text)
End Sub
Private Sub cmbSpZ07_GotFocus()
    Me.cmbSpZ07.SelStart = 0
    Me.cmbSpZ07.SelLength = Len(Me.cmbSpZ07.Text)
End Sub
Private Sub cmbSpZ08_GotFocus()
    Me.cmbSpZ08.SelStart = 0
    Me.cmbSpZ08.SelLength = Len(Me.cmbSpZ08.Text)
End Sub
Private Sub cmbSpZ09_GotFocus()
    Me.cmbSpZ09.SelStart = 0
    Me.cmbSpZ09.SelLength = Len(Me.cmbSpZ09.Text)
End Sub
Private Sub cmbSpZ10_GotFocus()
    Me.cmbSpZ10.SelStart = 0
    Me.cmbSpZ10.SelLength = Len(Me.cmbSpZ10.Text)
End Sub
Private Sub cmbSpZ11_GotFocus()
    Me.cmbSpZ11.SelStart = 0
    Me.cmbSpZ11.SelLength = Len(Me.cmbSpZ11.Text)
End Sub
Private Sub cmbSpZ12_GotFocus()
    Me.cmbSpZ12.SelStart = 0
    Me.cmbSpZ12.SelLength = Len(Me.cmbSpZ12.Text)
End Sub
Private Sub cmbSpZ13_GotFocus()
    Me.cmbSpZ13.SelStart = 0
    Me.cmbSpZ13.SelLength = Len(Me.cmbSpZ13.Text)
End Sub
Private Sub cmbSpZ14_GotFocus()
    Me.cmbSpZ14.SelStart = 0
    Me.cmbSpZ14.SelLength = Len(Me.cmbSpZ14.Text)
End Sub
Private Sub cmbSpZ15_GotFocus()
    Me.cmbSpZ15.SelStart = 0
    Me.cmbSpZ15.SelLength = Len(Me.cmbSpZ15.Text)
End Sub
Private Sub cmbSpZ16_GotFocus()
    Me.cmbSpZ16.SelStart = 0
    Me.cmbSpZ16.SelLength = Len(Me.cmbSpZ16.Text)
End Sub
Private Sub cmbSpZ17_GotFocus()
    Me.cmbSpZ17.SelStart = 0
    Me.cmbSpZ17.SelLength = Len(Me.cmbSpZ17.Text)
End Sub
Private Sub cmbSpZ18_GotFocus()
    Me.cmbSpZ18.SelStart = 0
    Me.cmbSpZ18.SelLength = Len(Me.cmbSpZ18.Text)
End Sub
Private Sub cmbSpZ19_GotFocus()
    Me.cmbSpZ19.SelStart = 0
    Me.cmbSpZ19.SelLength = Len(Me.cmbSpZ19.Text)
End Sub
Private Sub cmbSpZ20_GotFocus()
    Me.cmbSpZ20.SelStart = 0
    Me.cmbSpZ20.SelLength = Len(Me.cmbSpZ20.Text)
End Sub
Private Sub cmbSpZ21_GotFocus()
    Me.cmbSpZ21.SelStart = 0
    Me.cmbSpZ21.SelLength = Len(Me.cmbSpZ21.Text)
End Sub
Private Sub cmbSpZ22_GotFocus()
    Me.cmbSpZ22.SelStart = 0
    Me.cmbSpZ22.SelLength = Len(Me.cmbSpZ22.Text)
End Sub
Private Sub cmbSpZ23_GotFocus()
    Me.cmbSpZ23.SelStart = 0
    Me.cmbSpZ23.SelLength = Len(Me.cmbSpZ23.Text)
End Sub
Private Sub cmbSpZ24_GotFocus()
    Me.cmbSpZ24.SelStart = 0
    Me.cmbSpZ24.SelLength = Len(Me.cmbSpZ24.Text)
End Sub
Private Sub cmbSpZ25_GotFocus()
    Me.cmbSpZ25.SelStart = 0
    Me.cmbSpZ25.SelLength = Len(Me.cmbSpZ25.Text)
End Sub
Private Sub cmbSpZ26_GotFocus()
    Me.cmbSpZ26.SelStart = 0
    Me.cmbSpZ26.SelLength = Len(Me.cmbSpZ26.Text)
End Sub
Private Sub cmbSpZ27_GotFocus()
    Me.cmbSpZ27.SelStart = 0
    Me.cmbSpZ27.SelLength = Len(Me.cmbSpZ27.Text)
End Sub
Private Sub cmbSpZ28_GotFocus()
    Me.cmbSpZ28.SelStart = 0
    Me.cmbSpZ28.SelLength = Len(Me.cmbSpZ28.Text)
End Sub

Private Sub cmbSpZ01_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ01.Text) > TimeValue(Me.cmbSpZ02.Text) Then Me.cmbSpZ02.Text = Me.cmbSpZ01.Text
    If TimeValue(Me.cmbSpZ01.Text) > TimeValue(Me.cmbSpZ03.Text) Then Me.cmbSpZ03.Text = Me.cmbSpZ01.Text
    If TimeValue(Me.cmbSpZ01.Text) > TimeValue(Me.cmbSpZ04.Text) Then Me.cmbSpZ04.Text = Me.cmbSpZ01.Text
End If
    
GlAdS = True
    
End Sub
Private Sub cmbSpZ02_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ02.Text) < TimeValue(Me.cmbSpZ01.Text) Then Me.cmbSpZ01.Text = Me.cmbSpZ02.Text
    If TimeValue(Me.cmbSpZ02.Text) > TimeValue(Me.cmbSpZ03.Text) Then Me.cmbSpZ03.Text = Me.cmbSpZ02.Text
    If TimeValue(Me.cmbSpZ02.Text) > TimeValue(Me.cmbSpZ04.Text) Then Me.cmbSpZ04.Text = Me.cmbSpZ02.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ03_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ03.Text) < TimeValue(Me.cmbSpZ01.Text) Then Me.cmbSpZ01.Text = Me.cmbSpZ03.Text
    If TimeValue(Me.cmbSpZ03.Text) < TimeValue(Me.cmbSpZ02.Text) Then Me.cmbSpZ02.Text = Me.cmbSpZ03.Text
    If TimeValue(Me.cmbSpZ03.Text) > TimeValue(Me.cmbSpZ04.Text) Then Me.cmbSpZ04.Text = Me.cmbSpZ03.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ04_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ04.Text) < TimeValue(Me.cmbSpZ01.Text) Then Me.cmbSpZ01.Text = Me.cmbSpZ04.Text
    If TimeValue(Me.cmbSpZ04.Text) < TimeValue(Me.cmbSpZ02.Text) Then Me.cmbSpZ02.Text = Me.cmbSpZ04.Text
    If TimeValue(Me.cmbSpZ04.Text) < TimeValue(Me.cmbSpZ03.Text) Then Me.cmbSpZ03.Text = Me.cmbSpZ04.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ05_Click()
    
If GlAdL = False Then
    If TimeValue(Me.cmbSpZ05.Text) > TimeValue(Me.cmbSpZ06.Text) Then Me.cmbSpZ06.Text = Me.cmbSpZ05.Text
    If TimeValue(Me.cmbSpZ05.Text) > TimeValue(Me.cmbSpZ07.Text) Then Me.cmbSpZ07.Text = Me.cmbSpZ05.Text
    If TimeValue(Me.cmbSpZ05.Text) > TimeValue(Me.cmbSpZ08.Text) Then Me.cmbSpZ08.Text = Me.cmbSpZ05.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ06_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ06.Text) < TimeValue(Me.cmbSpZ05.Text) Then Me.cmbSpZ05.Text = Me.cmbSpZ06.Text
    If TimeValue(Me.cmbSpZ06.Text) > TimeValue(Me.cmbSpZ07.Text) Then Me.cmbSpZ07.Text = Me.cmbSpZ06.Text
    If TimeValue(Me.cmbSpZ06.Text) > TimeValue(Me.cmbSpZ08.Text) Then Me.cmbSpZ08.Text = Me.cmbSpZ06.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ07_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ07.Text) < TimeValue(Me.cmbSpZ05.Text) Then Me.cmbSpZ05.Text = Me.cmbSpZ07.Text
    If TimeValue(Me.cmbSpZ07.Text) < TimeValue(Me.cmbSpZ06.Text) Then Me.cmbSpZ06.Text = Me.cmbSpZ07.Text
    If TimeValue(Me.cmbSpZ07.Text) > TimeValue(Me.cmbSpZ08.Text) Then Me.cmbSpZ08.Text = Me.cmbSpZ07.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ08_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ08.Text) < TimeValue(Me.cmbSpZ05.Text) Then Me.cmbSpZ05.Text = Me.cmbSpZ08.Text
    If TimeValue(Me.cmbSpZ08.Text) < TimeValue(Me.cmbSpZ06.Text) Then Me.cmbSpZ06.Text = Me.cmbSpZ08.Text
    If TimeValue(Me.cmbSpZ08.Text) < TimeValue(Me.cmbSpZ07.Text) Then Me.cmbSpZ07.Text = Me.cmbSpZ08.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ09_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ09.Text) > TimeValue(Me.cmbSpZ10.Text) Then Me.cmbSpZ10.Text = Me.cmbSpZ09.Text
    If TimeValue(Me.cmbSpZ09.Text) > TimeValue(Me.cmbSpZ11.Text) Then Me.cmbSpZ11.Text = Me.cmbSpZ09.Text
    If TimeValue(Me.cmbSpZ09.Text) > TimeValue(Me.cmbSpZ12.Text) Then Me.cmbSpZ12.Text = Me.cmbSpZ09.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ10_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ10.Text) < TimeValue(Me.cmbSpZ09.Text) Then Me.cmbSpZ09.Text = Me.cmbSpZ10.Text
    If TimeValue(Me.cmbSpZ10.Text) > TimeValue(Me.cmbSpZ11.Text) Then Me.cmbSpZ11.Text = Me.cmbSpZ10.Text
    If TimeValue(Me.cmbSpZ10.Text) > TimeValue(Me.cmbSpZ12.Text) Then Me.cmbSpZ12.Text = Me.cmbSpZ10.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ11_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ11.Text) < TimeValue(Me.cmbSpZ09.Text) Then Me.cmbSpZ09.Text = Me.cmbSpZ11.Text
    If TimeValue(Me.cmbSpZ11.Text) < TimeValue(Me.cmbSpZ10.Text) Then Me.cmbSpZ10.Text = Me.cmbSpZ11.Text
    If TimeValue(Me.cmbSpZ11.Text) > TimeValue(Me.cmbSpZ12.Text) Then Me.cmbSpZ12.Text = Me.cmbSpZ11.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ12_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ12.Text) < TimeValue(Me.cmbSpZ09.Text) Then Me.cmbSpZ09.Text = Me.cmbSpZ12.Text
    If TimeValue(Me.cmbSpZ12.Text) < TimeValue(Me.cmbSpZ10.Text) Then Me.cmbSpZ10.Text = Me.cmbSpZ12.Text
    If TimeValue(Me.cmbSpZ12.Text) < TimeValue(Me.cmbSpZ11.Text) Then Me.cmbSpZ11.Text = Me.cmbSpZ12.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ13_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ13.Text) > TimeValue(Me.cmbSpZ14.Text) Then Me.cmbSpZ14.Text = Me.cmbSpZ13.Text
    If TimeValue(Me.cmbSpZ13.Text) > TimeValue(Me.cmbSpZ15.Text) Then Me.cmbSpZ15.Text = Me.cmbSpZ13.Text
    If TimeValue(Me.cmbSpZ13.Text) > TimeValue(Me.cmbSpZ16.Text) Then Me.cmbSpZ16.Text = Me.cmbSpZ13.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ14_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ14.Text) < TimeValue(Me.cmbSpZ13.Text) Then Me.cmbSpZ13.Text = Me.cmbSpZ14.Text
    If TimeValue(Me.cmbSpZ14.Text) > TimeValue(Me.cmbSpZ15.Text) Then Me.cmbSpZ15.Text = Me.cmbSpZ14.Text
    If TimeValue(Me.cmbSpZ14.Text) > TimeValue(Me.cmbSpZ16.Text) Then Me.cmbSpZ16.Text = Me.cmbSpZ14.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ15_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ15.Text) < TimeValue(Me.cmbSpZ13.Text) Then Me.cmbSpZ13.Text = Me.cmbSpZ15.Text
    If TimeValue(Me.cmbSpZ15.Text) < TimeValue(Me.cmbSpZ14.Text) Then Me.cmbSpZ14.Text = Me.cmbSpZ15.Text
    If TimeValue(Me.cmbSpZ15.Text) > TimeValue(Me.cmbSpZ16.Text) Then Me.cmbSpZ16.Text = Me.cmbSpZ15.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ16_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ16.Text) < TimeValue(Me.cmbSpZ13.Text) Then Me.cmbSpZ13.Text = Me.cmbSpZ16.Text
    If TimeValue(Me.cmbSpZ16.Text) < TimeValue(Me.cmbSpZ14.Text) Then Me.cmbSpZ14.Text = Me.cmbSpZ16.Text
    If TimeValue(Me.cmbSpZ16.Text) < TimeValue(Me.cmbSpZ15.Text) Then Me.cmbSpZ15.Text = Me.cmbSpZ16.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ17_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ17.Text) > TimeValue(Me.cmbSpZ18.Text) Then Me.cmbSpZ18.Text = Me.cmbSpZ17.Text
    If TimeValue(Me.cmbSpZ17.Text) > TimeValue(Me.cmbSpZ19.Text) Then Me.cmbSpZ19.Text = Me.cmbSpZ17.Text
    If TimeValue(Me.cmbSpZ17.Text) > TimeValue(Me.cmbSpZ20.Text) Then Me.cmbSpZ20.Text = Me.cmbSpZ17.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ18_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ18.Text) < TimeValue(Me.cmbSpZ17.Text) Then Me.cmbSpZ17.Text = Me.cmbSpZ18.Text
    If TimeValue(Me.cmbSpZ18.Text) > TimeValue(Me.cmbSpZ19.Text) Then Me.cmbSpZ19.Text = Me.cmbSpZ18.Text
    If TimeValue(Me.cmbSpZ18.Text) > TimeValue(Me.cmbSpZ20.Text) Then Me.cmbSpZ20.Text = Me.cmbSpZ18.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ19_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ19.Text) < TimeValue(Me.cmbSpZ17.Text) Then Me.cmbSpZ17.Text = Me.cmbSpZ19.Text
    If TimeValue(Me.cmbSpZ19.Text) < TimeValue(Me.cmbSpZ18.Text) Then Me.cmbSpZ18.Text = Me.cmbSpZ19.Text
    If TimeValue(Me.cmbSpZ19.Text) > TimeValue(Me.cmbSpZ20.Text) Then Me.cmbSpZ20.Text = Me.cmbSpZ19.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ20_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ20.Text) < TimeValue(Me.cmbSpZ17.Text) Then Me.cmbSpZ17.Text = Me.cmbSpZ20.Text
    If TimeValue(Me.cmbSpZ20.Text) < TimeValue(Me.cmbSpZ18.Text) Then Me.cmbSpZ18.Text = Me.cmbSpZ20.Text
    If TimeValue(Me.cmbSpZ20.Text) < TimeValue(Me.cmbSpZ19.Text) Then Me.cmbSpZ19.Text = Me.cmbSpZ20.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ21_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ21.Text) > TimeValue(Me.cmbSpZ22.Text) Then Me.cmbSpZ22.Text = Me.cmbSpZ21.Text
    If TimeValue(Me.cmbSpZ21.Text) > TimeValue(Me.cmbSpZ23.Text) Then Me.cmbSpZ23.Text = Me.cmbSpZ21.Text
    If TimeValue(Me.cmbSpZ21.Text) > TimeValue(Me.cmbSpZ24.Text) Then Me.cmbSpZ24.Text = Me.cmbSpZ21.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ22_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ22.Text) < TimeValue(Me.cmbSpZ21.Text) Then Me.cmbSpZ21.Text = Me.cmbSpZ22.Text
    If TimeValue(Me.cmbSpZ22.Text) > TimeValue(Me.cmbSpZ23.Text) Then Me.cmbSpZ23.Text = Me.cmbSpZ22.Text
    If TimeValue(Me.cmbSpZ22.Text) > TimeValue(Me.cmbSpZ24.Text) Then Me.cmbSpZ24.Text = Me.cmbSpZ22.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ23_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ23.Text) < TimeValue(Me.cmbSpZ21.Text) Then Me.cmbSpZ21.Text = Me.cmbSpZ23.Text
    If TimeValue(Me.cmbSpZ23.Text) < TimeValue(Me.cmbSpZ22.Text) Then Me.cmbSpZ22.Text = Me.cmbSpZ23.Text
    If TimeValue(Me.cmbSpZ23.Text) > TimeValue(Me.cmbSpZ24.Text) Then Me.cmbSpZ24.Text = Me.cmbSpZ23.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ24_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ24.Text) < TimeValue(Me.cmbSpZ21.Text) Then Me.cmbSpZ21.Text = Me.cmbSpZ24.Text
    If TimeValue(Me.cmbSpZ24.Text) < TimeValue(Me.cmbSpZ22.Text) Then Me.cmbSpZ22.Text = Me.cmbSpZ24.Text
    If TimeValue(Me.cmbSpZ24.Text) < TimeValue(Me.cmbSpZ23.Text) Then Me.cmbSpZ23.Text = Me.cmbSpZ24.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ25_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ25.Text) > TimeValue(Me.cmbSpZ26.Text) Then Me.cmbSpZ26.Text = Me.cmbSpZ25.Text
    If TimeValue(Me.cmbSpZ25.Text) > TimeValue(Me.cmbSpZ27.Text) Then Me.cmbSpZ27.Text = Me.cmbSpZ25.Text
    If TimeValue(Me.cmbSpZ25.Text) > TimeValue(Me.cmbSpZ28.Text) Then Me.cmbSpZ28.Text = Me.cmbSpZ25.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ26_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ26.Text) < TimeValue(Me.cmbSpZ25.Text) Then Me.cmbSpZ25.Text = Me.cmbSpZ26.Text
    If TimeValue(Me.cmbSpZ26.Text) > TimeValue(Me.cmbSpZ27.Text) Then Me.cmbSpZ27.Text = Me.cmbSpZ26.Text
    If TimeValue(Me.cmbSpZ26.Text) > TimeValue(Me.cmbSpZ28.Text) Then Me.cmbSpZ28.Text = Me.cmbSpZ26.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ27_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ27.Text) < TimeValue(Me.cmbSpZ25.Text) Then Me.cmbSpZ25.Text = Me.cmbSpZ27.Text
    If TimeValue(Me.cmbSpZ27.Text) < TimeValue(Me.cmbSpZ26.Text) Then Me.cmbSpZ26.Text = Me.cmbSpZ27.Text
    If TimeValue(Me.cmbSpZ27.Text) > TimeValue(Me.cmbSpZ28.Text) Then Me.cmbSpZ28.Text = Me.cmbSpZ27.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbSpZ28_Click()

If GlAdL = False Then
    If TimeValue(Me.cmbSpZ28.Text) < TimeValue(Me.cmbSpZ25.Text) Then Me.cmbSpZ25.Text = Me.cmbSpZ28.Text
    If TimeValue(Me.cmbSpZ28.Text) < TimeValue(Me.cmbSpZ26.Text) Then Me.cmbSpZ26.Text = Me.cmbSpZ28.Text
    If TimeValue(Me.cmbSpZ28.Text) < TimeValue(Me.cmbSpZ27.Text) Then Me.cmbSpZ27.Text = Me.cmbSpZ28.Text
End If
    
GlAdS = True

End Sub
Private Sub cmbRast1_Click()
    If GlAdL = False Then
        FRast
        GlAdS = True
    End If
End Sub

Private Sub txtZSRnr_Change()

TagWe = Mid$(Me.txtZSRnr.Tag, 2, Len(Me.txtZSRnr.Tag) - 1)

If GlAdL = False Then
    Me.txtZSRnr.Tag = "1" & TagWe
    GlAdS = True
End If

End Sub

Private Sub txtZSRnr_GotFocus()
    Me.txtZSRnr.SelStart = 0
    Me.txtZSRnr.SelLength = Len(Me.txtZSRnr.Text)
End Sub
Private Sub txtZSRnr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

