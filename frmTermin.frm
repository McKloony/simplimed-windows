VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmTermin 
   Caption         =   "Termineigenschaften"
   ClientHeight    =   11595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   ControlBox      =   0   'False
   Icon            =   "frmTermin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11595
   ScaleWidth      =   14880
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1020
      Left            =   12000
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   1799
      _StockProps     =   64
      BorderStyle     =   3
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont2 
      Height          =   1020
      Left            =   12000
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   1799
      _StockProps     =   64
      BorderStyle     =   3
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   5800
      Left            =   495
      TabIndex        =   38
      Top             =   5600
      Visible         =   0   'False
      Width           =   11000
      _Version        =   1048579
      _ExtentX        =   19403
      _ExtentY        =   10231
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnDatu3 
         Height          =   350
         Left            =   8420
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1800
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtRzTex 
         Height          =   350
         Left            =   7100
         TabIndex        =   58
         Tag             =   "0RzTex"
         Top             =   1280
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtGuiID 
         Height          =   350
         Left            =   7100
         TabIndex        =   56
         TabStop         =   0   'False
         Tag             =   "0GuiID"
         Top             =   280
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnPost3 
         Height          =   350
         Left            =   2230
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet die Postleitzahlensuche"
         Top             =   2800
         Width           =   330
         _Version        =   1048579
         _ExtentX        =   582
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox txtS4F02 
         Height          =   315
         Left            =   1500
         TabIndex        =   40
         Tag             =   "0Anrede"
         Top             =   780
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F09 
         Height          =   350
         Left            =   2600
         TabIndex        =   47
         Tag             =   "0Ort"
         Top             =   2800
         Width           =   2380
         _Version        =   1048579
         _ExtentX        =   4198
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F08 
         Height          =   350
         Left            =   1500
         TabIndex        =   45
         Tag             =   "0PLZ"
         Top             =   2800
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
         Left            =   1500
         TabIndex        =   44
         Tag             =   "0Straße"
         Top             =   2300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F05 
         Height          =   350
         Left            =   1500
         TabIndex        =   43
         Tag             =   "0Name"
         Top             =   1800
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F04 
         Height          =   350
         Left            =   1500
         TabIndex        =   42
         Tag             =   "0Vorname"
         Top             =   1280
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F03 
         Height          =   350
         Left            =   3800
         TabIndex        =   41
         Tag             =   "0Titel"
         Top             =   780
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F01 
         Height          =   350
         Left            =   1500
         TabIndex        =   39
         Tag             =   "0Firma1"
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
      Begin XtremeSuiteControls.ComboBox cmbS4F12 
         Height          =   315
         Left            =   1500
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "0Land"
         Top             =   3300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F16 
         Height          =   350
         Left            =   1500
         TabIndex        =   52
         Tag             =   "0Telefon5"
         Top             =   4800
         Width           =   3060
         _Version        =   1048579
         _ExtentX        =   5397
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F15 
         Height          =   350
         Left            =   1500
         TabIndex        =   50
         Tag             =   "0Telefon1"
         Top             =   4300
         Width           =   3060
         _Version        =   1048579
         _ExtentX        =   5397
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnTele9 
         Height          =   350
         Left            =   4600
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Wählt die nebenstehende Rufnummer"
         Top             =   4800
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTele8 
         Height          =   350
         Left            =   4600
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Wählt die nebenstehende Rufnummer"
         Top             =   4300
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbS4F11 
         Height          =   315
         Left            =   1500
         TabIndex        =   49
         Tag             =   "0Briefanrede"
         Top             =   3800
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F18 
         Height          =   350
         Left            =   1500
         TabIndex        =   54
         Tag             =   "0Geboren"
         Top             =   5300
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
      Begin XtremeSuiteControls.ComboBox cmbArzNr 
         Height          =   315
         Left            =   7100
         TabIndex        =   57
         Tag             =   "0IDO"
         Top             =   780
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6191
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtRzDat 
         Height          =   350
         Left            =   7100
         TabIndex        =   59
         Tag             =   "0RzDat"
         Top             =   1800
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRzNum 
         Height          =   350
         Left            =   7100
         TabIndex        =   62
         Tag             =   "0RzNum"
         Top             =   2800
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtZeiAn 
         Height          =   350
         Left            =   7100
         TabIndex        =   63
         Tag             =   "0ZeiAn"
         Top             =   3300
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNotDa 
         Height          =   350
         Left            =   7100
         TabIndex        =   66
         Tag             =   "0NotifySendDat"
         Top             =   4300
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNotZe 
         Height          =   350
         Left            =   8500
         TabIndex        =   67
         Tag             =   "0NotifySendTime"
         Top             =   4300
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNoDat 
         Height          =   350
         Left            =   7100
         TabIndex        =   64
         TabStop         =   0   'False
         Tag             =   "0NotifySetDate"
         Top             =   3800
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNoTim 
         Height          =   350
         Left            =   8500
         TabIndex        =   65
         TabStop         =   0   'False
         Tag             =   "0NotifySetTime"
         Top             =   3800
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRzAnz 
         Height          =   350
         Left            =   7100
         TabIndex        =   61
         Tag             =   "0RzAnz"
         Top             =   2300
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtOnBuT 
         Height          =   350
         Left            =   8500
         TabIndex        =   69
         Tag             =   "0OnlBook"
         Top             =   4800
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtOnSyT 
         Height          =   350
         Left            =   8500
         TabIndex        =   71
         Tag             =   "0OnlSync"
         Top             =   5300
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtOnBuD 
         Height          =   350
         Left            =   7100
         TabIndex        =   68
         Tag             =   "0OnlBook"
         Top             =   4800
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtOnSyD 
         Height          =   350
         Left            =   7100
         TabIndex        =   70
         Tag             =   "0OnlSync"
         Top             =   5300
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtS4F20 
         Height          =   350
         Left            =   3500
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   5300
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblLab47 
         Height          =   240
         Left            =   2860
         TabIndex        =   132
         Top             =   5350
         Width           =   600
         _Version        =   1048579
         _ExtentX        =   1058
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "PIN :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab46 
         Height          =   240
         Left            =   5300
         TabIndex        =   131
         Top             =   5350
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Synchronisiert :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab45 
         Height          =   240
         Left            =   5300
         TabIndex        =   130
         Top             =   4850
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Onlinegebucht :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab44 
         Height          =   240
         Left            =   5300
         TabIndex        =   129
         Top             =   4350
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Gesendet :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab37 
         Height          =   240
         Left            =   5300
         TabIndex        =   121
         Top             =   3850
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Emailerinnerung :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab36 
         Height          =   240
         Left            =   5300
         TabIndex        =   120
         Top             =   3350
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Wartezimmerzeit :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab35 
         Height          =   255
         Left            =   5300
         TabIndex        =   118
         Top             =   1340
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Verordnungsdiagnose :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab34 
         Height          =   240
         Left            =   5300
         TabIndex        =   117
         Top             =   2340
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Verordnungsmenge :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab33 
         Height          =   240
         Left            =   5300
         TabIndex        =   116
         Top             =   1850
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Verordnungsdatum :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab32 
         Height          =   240
         Left            =   5300
         TabIndex        =   115
         Top             =   2850
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Verordnungsbeleg :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab31 
         Height          =   240
         Left            =   5300
         TabIndex        =   114
         Top             =   850
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Verordner / Hausarzt :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblGuiID 
         Height          =   240
         Left            =   5300
         TabIndex        =   113
         Top             =   360
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Terminkennung :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab27 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Briefanrede :"
         Height          =   240
         Left            =   420
         TabIndex        =   111
         Top             =   3850
         Width           =   1020
      End
      Begin VB.Label lblLab29 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Telefon :"
         Height          =   240
         Left            =   420
         TabIndex        =   110
         Top             =   4350
         Width           =   1020
      End
      Begin VB.Label lblLab30 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   240
         Left            =   420
         TabIndex        =   109
         Top             =   4850
         Width           =   1020
      End
      Begin VB.Label lblLab28 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren :"
         Height          =   240
         Left            =   420
         TabIndex        =   108
         Top             =   5350
         Width           =   1020
      End
      Begin VB.Label lblLab26 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Land :"
         Height          =   240
         Left            =   420
         TabIndex        =   107
         Top             =   3350
         Width           =   1020
      End
      Begin VB.Label lblLab20 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Firma/Instit.:"
         Height          =   240
         Left            =   420
         TabIndex        =   106
         Top             =   350
         Width           =   1020
      End
      Begin VB.Label lblLab21 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anrede :"
         Height          =   240
         Left            =   420
         TabIndex        =   105
         Top             =   850
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Titel :"
         Height          =   240
         Left            =   3300
         TabIndex        =   104
         Top             =   850
         Width           =   420
      End
      Begin VB.Label lblLab22 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname :"
         Height          =   240
         Left            =   420
         TabIndex        =   103
         Top             =   1340
         Width           =   1020
      End
      Begin VB.Label lblLab23 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Nachname :"
         Height          =   240
         Left            =   420
         TabIndex        =   102
         Top             =   1850
         Width           =   1020
      End
      Begin VB.Label lblLab24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Straße :"
         Height          =   240
         Left            =   420
         TabIndex        =   101
         Top             =   2350
         Width           =   1020
      End
      Begin VB.Label lblLab25 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PLZ / Ort :"
         Height          =   240
         Left            =   420
         TabIndex        =   100
         Top             =   2850
         Width           =   1020
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   5400
      Left            =   500
      TabIndex        =   1
      Top             =   120
      Width           =   11000
      _Version        =   1048579
      _ExtentX        =   19403
      _ExtentY        =   9525
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown updCont3 
         Height          =   350
         Left            =   7210
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
         Left            =   5710
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
         Height          =   405
         Left            =   360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   5000
         Visible         =   0   'False
         Width           =   400
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   706
         _StockProps     =   64
         Show3DBorder    =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRaum1 
         Height          =   315
         Left            =   4800
         TabIndex        =   25
         Tag             =   "0Raum"
         Top             =   3800
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtAdres 
         Height          =   350
         Left            =   1500
         TabIndex        =   2
         Tag             =   "0Patient"
         Top             =   300
         Width           =   9300
         _Version        =   1048579
         _ExtentX        =   16404
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKomme 
         Height          =   800
         Left            =   1500
         TabIndex        =   27
         Tag             =   "0Kommentar"
         Top             =   4300
         Width           =   9300
         _Version        =   1048579
         _ExtentX        =   16404
         _ExtentY        =   1411
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   350
         Left            =   3180
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   1800
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   3180
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
         Left            =   1500
         TabIndex        =   3
         Tag             =   "0IDKurz"
         Top             =   800
         Width           =   9300
         _Version        =   1048579
         _ExtentX        =   16404
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         MaxLength       =   250
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbStatu 
         Height          =   310
         Left            =   1500
         TabIndex        =   15
         Tag             =   "0Farbtyp"
         Top             =   2300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox cmbPrior 
         Height          =   315
         Left            =   1485
         TabIndex        =   18
         Tag             =   "0Priorität"
         Top             =   2800
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox5"
      End
      Begin XtremeSuiteControls.ComboBox cmbRaum1 
         Height          =   315
         Left            =   4800
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "0IDR"
         Top             =   1800
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox6"
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   310
         Left            =   4800
         TabIndex        =   19
         Tag             =   "0IDP"
         Top             =   2800
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox7"
      End
      Begin XtremeSuiteControls.ComboBox cmbRemin 
         Height          =   310
         Left            =   9400
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "0Vorwarn"
         Top             =   1300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox8"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   350
         Left            =   1500
         TabIndex        =   11
         Tag             =   "0BisDat"
         Top             =   1800
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
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   1500
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
         Left            =   4800
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
         Left            =   6300
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
      Begin XtremeSuiteControls.FlatEdit txtRefNr 
         Height          =   350
         Left            =   9400
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "0MasTer"
         Top             =   3800
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbTeTyp 
         Height          =   315
         Left            =   1500
         TabIndex        =   21
         Tag             =   "0TerTyp"
         Top             =   3300
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
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   4800
         TabIndex        =   22
         Tag             =   "0IDM"
         Top             =   3300
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox7"
      End
      Begin XtremeSuiteControls.ComboBox cmbGesch 
         Height          =   315
         Left            =   1500
         TabIndex        =   24
         Tag             =   "0Geschlecht"
         Top             =   3800
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
      Begin XtremeSuiteControls.ComboBox cmbAbger 
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Tag             =   "0Aufgabe"
         Top             =   2300
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
      Begin XtremeSuiteControls.ComboBox cmbGanzt 
         Height          =   310
         Left            =   9400
         TabIndex        =   17
         Top             =   2300
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
      Begin XtremeSuiteControls.ComboBox cmbAbgeh 
         Height          =   310
         Left            =   9400
         TabIndex        =   20
         Top             =   2800
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbNotVa 
         Height          =   315
         Left            =   9400
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "0NotifyValue"
         Top             =   1800
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbOnlTe 
         Height          =   315
         Left            =   9400
         TabIndex        =   23
         Tag             =   "0OnlEmp"
         Top             =   3300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab43 
         Height          =   240
         Left            =   7940
         TabIndex        =   128
         Top             =   3350
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Onlinegebucht :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab42 
         Height          =   240
         Left            =   7940
         TabIndex        =   126
         Top             =   2850
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Abgehakt :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab41 
         Height          =   240
         Left            =   7940
         TabIndex        =   125
         Top             =   2350
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Ganztagstermin :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab40 
         Height          =   240
         Left            =   7940
         TabIndex        =   124
         Top             =   1850
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Emailerinnerung :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab39 
         Height          =   240
         Left            =   420
         TabIndex        =   123
         Top             =   2850
         Width           =   1020
         _Version        =   1048579
         _ExtentX        =   1799
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Priorität :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   240
         Left            =   420
         TabIndex        =   99
         Top             =   850
         Width           =   1020
         _Version        =   1048579
         _ExtentX        =   1799
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Betreff :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab17 
         Height          =   240
         Left            =   3800
         TabIndex        =   98
         Top             =   3840
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Terminort :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab12 
         Height          =   240
         Left            =   420
         TabIndex        =   97
         Top             =   3840
         Width           =   1020
         _Version        =   1048579
         _ExtentX        =   1799
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Geschlecht :"
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab16 
         Height          =   240
         Left            =   3800
         TabIndex        =   91
         Top             =   3350
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Mitarbeiter :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab14 
         Height          =   240
         Left            =   3800
         TabIndex        =   90
         Top             =   2350
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Leistungen :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   240
         Left            =   7940
         TabIndex        =   89
         Top             =   1350
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Terminerinnerung :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   240
         Left            =   7940
         TabIndex        =   88
         Top             =   3840
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Seriennummer :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab13 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar :"
         Height          =   240
         Left            =   420
         TabIndex        =   86
         Top             =   4350
         Width           =   1020
      End
      Begin VB.Label lblLab09 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Terminstatus :"
         Height          =   240
         Left            =   420
         TabIndex        =   85
         Top             =   2350
         Width           =   1020
      End
      Begin VB.Label lblLab15 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   240
         Left            =   3800
         TabIndex        =   84
         Top             =   2850
         Width           =   930
      End
      Begin VB.Label lblLab11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Marker :"
         Height          =   240
         Left            =   420
         TabIndex        =   83
         Top             =   3350
         Width           =   1020
      End
      Begin VB.Label lblLab04 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Terminende :"
         Height          =   240
         Left            =   420
         TabIndex        =   82
         Top             =   1850
         Width           =   1020
      End
      Begin VB.Label lblLab01 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Patient :"
         Height          =   240
         Left            =   420
         TabIndex        =   81
         Top             =   350
         Width           =   1020
      End
      Begin VB.Label lblLab03 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Terminstart :"
         Height          =   240
         Left            =   420
         TabIndex        =   80
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label lblLab05 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Terminzeit :"
         Height          =   240
         Left            =   3800
         TabIndex        =   79
         Top             =   1350
         Width           =   930
      End
      Begin VB.Label lblLab10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Raumplan :"
         Height          =   240
         Left            =   3800
         TabIndex        =   78
         Top             =   1850
         Width           =   930
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
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
   Begin XtremeSuiteControls.FlatEdit txtID2 
      Height          =   195
      Left            =   1200
      TabIndex        =   72
      TabStop         =   0   'False
      Tag             =   "0ID2"
      Top             =   12000
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
      TabIndex        =   73
      TabStop         =   0   'False
      Tag             =   "0Farbe"
      Top             =   12000
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
      TabIndex        =   74
      TabStop         =   0   'False
      Tag             =   "0ID0"
      Top             =   12000
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
      TabIndex        =   75
      TabStop         =   0   'False
      Tag             =   "0IDSer"
      Top             =   12000
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
      TabIndex        =   76
      TabStop         =   0   'False
      Tag             =   "0MasTer"
      Top             =   12000
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
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "0SerTyp"
      Top             =   12000
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
   Begin XtremeSuiteControls.FlatEdit txtPaTel 
      Height          =   195
      Left            =   3120
      TabIndex        =   87
      TabStop         =   0   'False
      Tag             =   "0Datei"
      Top             =   12000
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
      Left            =   5500
      TabIndex        =   92
      TabStop         =   0   'False
      Tag             =   "0Wiederholung"
      Top             =   12000
      Width           =   220
      _Version        =   1048579
      _ExtentX        =   388
      _ExtentY        =   388
      _StockProps     =   79
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtFall1 
      Height          =   195
      Left            =   3800
      TabIndex        =   93
      TabStop         =   0   'False
      Tag             =   "0Fallig1"
      Top             =   12000
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
      Left            =   4200
      TabIndex        =   94
      TabStop         =   0   'False
      Tag             =   "0Fallig2"
      Top             =   12000
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
      Left            =   4600
      TabIndex        =   95
      TabStop         =   0   'False
      Tag             =   "0GesBetrag"
      Top             =   12000
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
      Left            =   5000
      TabIndex        =   96
      TabStop         =   0   'False
      Tag             =   "0Behindert"
      Top             =   12000
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
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   405
      Left            =   12000
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   6495
      _Version        =   1048579
      _ExtentX        =   11465
      _ExtentY        =   714
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbTypen 
         Height          =   315
         Left            =   40
         TabIndex        =   30
         Top             =   60
         Width           =   720
         _Version        =   1048579
         _ExtentX        =   1270
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbZiffe 
         Height          =   315
         Left            =   800
         TabIndex        =   31
         Top             =   60
         Width           =   1050
         _Version        =   1048579
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         DropDownItemCount=   20
      End
      Begin XtremeSuiteControls.ComboBox cmbBezei 
         Height          =   315
         Left            =   1880
         TabIndex        =   32
         Top             =   60
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3545
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         AutoComplete    =   -1  'True
         DropDownItemCount=   20
      End
      Begin XtremeSuiteControls.FlatEdit txtEinze 
         Height          =   350
         Left            =   5060
         TabIndex        =   35
         Top             =   60
         Width           =   795
         _Version        =   1048579
         _ExtentX        =   1402
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtAnzal 
         Height          =   350
         Left            =   3900
         TabIndex        =   33
         Top             =   60
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   873
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMulti 
         Height          =   350
         Left            =   4430
         TabIndex        =   34
         Top             =   60
         Width           =   600
         _Version        =   1048579
         _ExtentX        =   1058
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtS3F01 
      Height          =   195
      Left            =   6000
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   12000
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
   Begin XtremeSuiteControls.FlatEdit txtRzAkt 
      Height          =   195
      Left            =   3500
      TabIndex        =   119
      TabStop         =   0   'False
      Tag             =   "0RzAkt"
      Top             =   12000
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
   Begin XtremeSuiteControls.FlatEdit txtDatum 
      Height          =   195
      Left            =   6400
      TabIndex        =   122
      TabStop         =   0   'False
      Tag             =   "0Datum"
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
   Begin XtremeSuiteControls.FlatEdit txtNoSta 
      Height          =   195
      Left            =   6800
      TabIndex        =   127
      TabStop         =   0   'False
      Tag             =   "0NotifyStatus"
      Top             =   12000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   344
      _ExtentY        =   344
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
Attribute VB_Name = "frmTermin"
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
Private TxID0 As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private TxKom As XtremeSuiteControls.FlatEdit
Private TxID2 As XtremeSuiteControls.FlatEdit
Private TxFar As XtremeSuiteControls.FlatEdit
Private TxAdr As XtremeSuiteControls.FlatEdit
Private TxMas As XtremeSuiteControls.FlatEdit
Private TxSTy As XtremeSuiteControls.FlatEdit
Private VoZei As XtremeSuiteControls.FlatEdit
Private TxAnz As XtremeSuiteControls.FlatEdit
Private TxMul As XtremeSuiteControls.FlatEdit
Private TxEin As XtremeSuiteControls.FlatEdit
Private BiZei As XtremeSuiteControls.FlatEdit
Private TxNoS As XtremeSuiteControls.FlatEdit
Private TxNoD As XtremeSuiteControls.FlatEdit
Private TxNoZ As XtremeSuiteControls.FlatEdit
Private TxRak As XtremeSuiteControls.FlatEdit
Private TxRzA As XtremeSuiteControls.FlatEdit
Private CmRem As XtremeSuiteControls.ComboBox
Private CmBet As XtremeSuiteControls.ComboBox
Private CmRmu As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmPri As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmGes As XtremeSuiteControls.ComboBox
Private CmETy As XtremeSuiteControls.ComboBox
Private CmZif As XtremeSuiteControls.ComboBox
Private CmBez As XtremeSuiteControls.ComboBox
Private CmRzA As XtremeSuiteControls.ComboBox
Private CmGan As XtremeSuiteControls.ComboBox
Private CmSpe As XtremeSuiteControls.ComboBox
Private CmAbg As XtremeSuiteControls.ComboBox
Private CmNot As XtremeSuiteControls.ComboBox
Private CmStu As XtremeSuiteControls.ComboBox
Private CmAbr As XtremeSuiteControls.ComboBox
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
Private MoKal As XtremeCalendarControl.DatePicker
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private RetWe As Long
Private TagWe As String
Private PaStr As String
Private KalWa As Integer
Private TabId As Integer

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
Private Sub FAdre()
On Error GoTo LiErr

Dim Mld1, Tit1 As String

Set TxID0 = Me.txtID0

Mld1 = "Es wurde noch kein Patient zugeordnet"
Tit1 = "Kei Patint"

If TxID0.Text <> vbNullString Then
    If IsNumeric(TxID0.Text) = True Then
        If CLng(TxID0.Text) > 0 Then
            AMain CLng(TxID0.Text)
        Else
            WindowMess Mld1, Dial3, Tit1, Me.hwnd
        End If
    Else
        WindowMess Mld1, Dial3, Tit1, Me.hwnd
    End If
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
Dim TmStr As String
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

Set FM = frmMain
Set CaCol = FM.calCont1
Set DaPro = CaCol.DataProvider
Set CaLbs = DaPro.LabelList

RmIdx = 0
MiIdx = 0
IdxNr = CmBet.ListIndex + 1

If CmBet.Text <> vbNullString Then
    TmStr = CmBet.Text
End If

If IdxNr > 0 Then
    If LCase(GlBtr(IdxNr, 1)) = LCase(TmStr) Then
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
                If RmuNr = GlRmu(AktZa, 2) Then
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
            For AktZa = 1 To UBound(GlMiT)
                If MitNr = GlMiT(AktZa, 2) Then
                    ManNr = GlMiT(AktZa, 7)
                    MiIdx = AktZa - 1
                    CmMit.ListIndex = MiIdx
                    Exit For
                End If
            Next AktZa
    
            For AktZa = 1 To UBound(GlMaT)
                If ManNr = GlMaT(AktZa, 2) Then
                    MaIdx = AktZa - 1
                    CmMan.ListIndex = MaIdx
                    Exit For
                End If
            Next AktZa
        End If
    Else
        FarWe = vbWhite
        ZeiVo = 0
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
    End If
Else
    'FarWe = vbWhite
    ZeiVo = 0
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
Private Sub FClip(ByVal AnTyp As Integer)
On Error GoTo WoErr
'Kopiert Adresse in die Zwischenablage

Dim TmStr As Variant

Select Case AnTyp
Case 1:
    TmStr = Me.txtID0.Text
Case 2:
    TmStr = Me.txtGuiID.Text
End Select

Clipboard.Clear
Clipboard.SetText TmStr

SPopu "Zwischenablage", "Die Informationen wurden in die Zwischenablage kopiert", IC48_Information

Exit Sub

WoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClip" & Err.Number
Resume Next

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
        IniSetVal "Termin", "FenLin", clFen.FeLin
        IniSetVal "Termin", "FenObe", clFen.FeObn
        IniSetVal "Termin", "FenBre", clFen.FeBre
        IniSetVal "Termin", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub

Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date

Set CmGan = Me.cmbGanzt
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtRzDat
Set MoKal = Me.dtpDatu1

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
        TxDa1.Text = NeuDa
        TxDa2.Text = NeuDa
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

If IsDate(TxDa1.Text) = True Then
    Datu1 = CDate(TxDa1.Text)
Else
    Exit Sub
End If

If IsDate(TxDa2.Text) = True Then
    Datu2 = CDate(TxDa2.Text)
Else
    Exit Sub
End If

If GlOTS = True Then 'Online-Terminbuchungs Sytem
    If Datu2 <> Datu1 Then
        TxDa2.Text = Datu1
    End If
Else
    If Datu2 > Datu1 Then
        CmGan.ListIndex = 1
    End If
End If

With MoKal
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

Set MoKal = Nothing

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

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtRzDat
Set MoKal = Me.dtpDatu1
Set CmGan = Me.cmbGanzt

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
    
    Datu1 = CDate(TxDa1.Text)
    Datu2 = CDate(TxDa2.Text)
    
    If Datu2 > Datu1 Then
        CmGan.ListIndex = 1
    End If
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FDrop()
On Error GoTo OrErr

Dim RowNr As Long
Dim KrRow As Long
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
Set CmAcs = CmBrs.Actions

Tr_Einf
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpTe RowNr
End If

Set RpSel = RpCon.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    KrRow = RpRow.Index
Else
    KrRow = 1
End If

TUpAb KrRow

If WindowLoad("frmTermin") = True Then
    CmAcs(AD_Termin_Abrechnen).Enabled = True
    CmAcs(AD_Termin_EintLoe).Enabled = True
End If

Set RpCon = Nothing
Set RpCo1 = Nothing
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
Private Sub FEiKe(Optional ByVal StaKe As Integer)
On Error GoTo OrErr
'Standardkette einfügen

Dim TmVon As Date
Dim TmBis As Date
Dim RowNr As Long
Dim KrRow As Long
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

Tr_Einf AdMin, StaKe
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpTe RowNr
End If

Set RpSel = RpCon.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    KrRow = RpRow.Index
Else
    KrRow = 1
End If

TUpAb KrRow
CmAcs(AD_Termin_Abrechnen).Enabled = True
CmAcs(AD_Termin_EintLoe).Enabled = True

GlTSa = True

Set RpCon = Nothing
Set RpCo1 = Nothing
Set RpCls = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEiKe " & Err.Number
Resume Next

End Sub
Private Sub FEinf(ByVal EngTy As Integer)
On Error GoTo OrErr

Dim KrRow As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTermin
Set RpCon = FM.repCont2
Set RpSel = RpCon.SelectedRows

Ter_Ein EngTy
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    KrRow = RpRow.Index
Else
    KrRow = 1
End If

TUpAb KrRow

GlTSa = True

Set RpCon = Nothing
Set RpCls = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSum " & Err.Number
Resume Next

End Sub
Private Sub FEmai()
On Error GoTo WoErr
'Kopiert Adresse in die Zwischenablage

Dim PatNr As Long
Dim TmStr As String
Dim EmTex As String
Dim EmBet As String
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmTermin
Set TxID0 = FM.txtID0
Set RpCo2 = FM.repCont2
Set RpRws = RpCo2.Rows

If TxID0.Text <> vbNullString Then
    If IsNumeric(TxID0.Text) = True Then
        If CLng(TxID0.Text) > 0 Then
            PatNr = CLng(TxID0.Text)
        End If
    End If
End If

EmBet = "Terminprotokoll " & Format$(Now, "YYYYMMDD_HHMM")

If RpRws.Count > 0 Then
    For Each RpRow In RpRws
        If RpRow.GroupRow = False Then
            TmStr = TmStr & RpRow.Record(TeP_ID2).Value & ";" & RpRow.Record(TeP_TerID).Value & ";" & RpRow.Record(TeP_Datum).Value & ";" & RpRow.Record(TeP_Zeit).Value & ";" & RpRow.Record(TeP_IDKurz).Value & ";" & RpRow.Record(TeP_Kommen).Value & vbCrLf
        End If
    Next RpRow
    
    EmTex = vbCrLf & vbCrLf & TmStr & vbCrLf
    SMaNe PatNr, , , EmTex, EmBet
    DoEvents
        
    Unload FM
End If

Set RpRws = Nothing
Set RpCo2 = Nothing

Exit Sub

WoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEmai" & Err.Number
Resume Next

End Sub
Private Sub FExpo()
On Error GoTo WoErr
'Kopiert Adresse in die Zwischenablage

Dim TmStr As String
Dim DaNam As String
Dim FilNa As String
Dim Frage As Integer
Dim RetWe As Boolean
Dim Mld1, Tit1 As String
Dim CoDia As XtremeSuiteControls.CommonDialog
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmTermin
Set RpCo2 = FM.repCont2
Set RpRws = RpCo2.Rows
Set CoDia = frmMain.comDialo

DaNam = "Terminprotokoll_" & Format$(Now, "YYYYMMDD_HHMM") & ".txt"
Mld1 = "Die Datei existiert bereits, soll diese überschrieben werden?"
Tit1 = "Terminprotokollexport"

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

If RpRws.Count > 0 Then
    For Each RpRow In RpRws
        If RpRow.GroupRow = False Then
            TmStr = TmStr & RpRow.Record(TeP_ID2).Value & ";" & RpRow.Record(TeP_TerID).Value & ";" & RpRow.Record(TeP_Datum).Value & ";" & RpRow.Record(TeP_Zeit).Value & ";" & RpRow.Record(TeP_IDKurz).Value & ";" & RpRow.Record(TeP_Kommen).Value & vbCrLf
        End If
    Next RpRow

    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.txt"
        .Filter = "Ascii-Text Format (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
        .DialogTitle = "Bitte Name und Ordner der Protokolldatei angeben"
        If GlRDP = True Then
            .FileName = GlIPf & DaNam
            .InitDir = GlIPf
        Else
            .FileName = GlEPf & DaNam
            .InitDir = GlEPf
        End If
        .ShowSave
        FilNa = .FileName
        If .FileTitle = vbNullString Then
            Set clFil = Nothing
            Set CoDia = Nothing
            Set RpRws = Nothing
            Set RpCo2 = Nothing
            Exit Sub
        End If
    End With
    If Right$(FilNa, 4) <> ".txt" Then FilNa = FilNa & ".txt"
    
    With clFil
        If .FilVor(FilNa) = True Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                .DaLoe = FilNa & vbNullChar
                .FilLoe
            Else
                Set clFil = Nothing
                Set CoDia = Nothing
                Set RpRws = Nothing
                Set RpCo2 = Nothing
                Exit Sub
            End If
        End If
    End With
    
    With clFil
        .FilPfa FilNa
        .StrDa = TmStr
        RetWe = .FilWrSt
        .StrDa = vbNullString
    End With
End If

Set clFil = Nothing

Set CoDia = Nothing
Set RpRws = Nothing
Set RpCo2 = Nothing

Exit Sub

WoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FExpo" & Err.Number
Resume Next

End Sub
Public Sub FGeLo()
On Error GoTo LiErr
'Löscht einen Gebühreneitrag

Dim TerNr As Long
Dim RowNr As Long
Dim KrRow As Long
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermin
Set TxID2 = FM.txtID2
Set RpCon = FM.repCont1
Set RpRcs = RpCon.Records

If TxID2.Text <> vbNullString Then
    TerNr = TxID2.Text
Else
    TerNr = GlTem
End If

Set RpSel = RpCon.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    KrRow = RpRow.Index
    Ter_Del
    DoEvents
    TUpAb KrRow, TerNr
End If

Set RpCo1 = frmMain.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

Ter_Lei TerNr, True
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpTe RowNr
End If

Set RpRcs = Nothing
Set RpSel = Nothing
Set RpCon = Nothing
Set RpCo1 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FGeLo " & Err.Number
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
Set TxDa3 = Me.txtRzDat
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
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 3: .Top = TxDa3.Top + TxDa2.Height
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

If Datu2 < Datu1 Then
    TxDa1.Text = Datu2
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
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

Dim TerNr As Long
Dim MasTe As Long
Dim KrRow As Long
Dim GesBe As Double
Dim EinBe As Double
Dim Fakto As Single
Dim RowNr As Integer
Dim Anzal As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermin
Set TxID2 = FM.txtID2
Set TxMas = FM.txtRefNr
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

If TxID2.Text <> vbNullString Then
    TerNr = TxID2.Text
Else
    TerNr = GlTem
End If

If TxMas.Text <> vbNullString Then
    MasTe = TxMas.Text
Else
    MasTe = 0
End If

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
                GesBe = CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(TeL_Betrag)
                RpRow.Record(RpCol.ItemIndex).Value = GesBe
            Else
                Set RpCol = RpCls.Find(TeL_Betrag)
                RpRow.Record(RpCol.ItemIndex).Value = 0
            End If
        Else
            EinBe = CDbl(RpRow.Record(RpCol.ItemIndex).Value)
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
        EinBe = CDbl(RpRow.Record(RpCol.ItemIndex).Value)
        RpRow.Record(RpCol.ItemIndex).Value = Round(Format$(EinBe, GlWa1), 2)
        RpRow.Record(RpCol.ItemIndex).Tag = "@Preis1" 'Tag geändert
        Set RpCol = RpCls.Find(TeL_Gesamt)
        RpRow.Record(RpCol.ItemIndex).Value = Round(Format$(EinBe * Fakto * Anzal, GlWa1), 2)
        RpRow.Record(RpCol.ItemIndex).Tag = "@Preis2" 'Tag geändert
    End If
End If

DoEvents
Ter_LeS TerNr
DoEvents

Ter_Akt TerNr, MasTe
DoEvents

TUpAb RpRow.Index, TerNr

Set RpCon = frmMain.repCont1
Set RpSel = RpCon.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    SUpTe RowNr
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCon = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKran " & Err.Number
Resume Next

End Sub
Private Sub FKrTy()
On Error GoTo OrErr

Dim CoIdx As Integer

Set CmETy = Me.cmbTypen
Set CmZif = Me.cmbZiffe
Set CmBez = Me.cmbBezei

CoIdx = CmETy.ListIndex + 1

CmZif.Clear
CmBez.Clear

Select Case CoIdx
Case 1:
    Ter_Com
    CmBez.SetFocus
Case 6:
    
Case Else:
    Ter_Com
    CmZif.SetFocus
End Select

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKrTy " & Err.Number
Resume Next

End Sub
Private Sub FMita()
On Error GoTo LiErr

Dim MitNr As Long
Dim ManNr As Long
Dim MaIdx As Integer
Dim AktZa As Integer

Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMan.ListCount > 1 Then
        MitNr = CmMit.ItemData(CmMit.ListIndex)
        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter
            If MitNr = GlMiT(AktZa, 2) Then
                ManNr = GlMiT(AktZa, 7) 'zugeordnete Mandantennummer
                Exit For
            End If
        Next AktZa
        For AktZa = 1 To UBound(GlMaT) 'Aktive Mitarbeiter
            If ManNr = GlMaT(AktZa, 2) Then
                MaIdx = AktZa - 1
                CmMan.ListIndex = MaIdx
                Exit For
            End If
        Next AktZa
    End If
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMita " & Err.Number
Resume Next

End Sub
Private Sub FNoSe()
On Error GoTo WoErr
'Errechnet und speichert das Emailerinnerungsdatum

Dim AkDat As Date
Dim StaZe As Date
Dim NotDa As String
Dim NotZe As String
Dim NotVa As Integer

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set TxNoD = Me.txtNoDat
Set TxNoZ = Me.txtNoTim
Set CmNot = Me.cmbNotVa

NotVa = CmNot.ItemData(CmNot.ListIndex)

If NotVa = 0 Then
    NotVa = 24
End If

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        If VoZei.Text <> vbNullString Then
            If IsDate(VoZei.Text) = True Then
    
                AkDat = CDate(TxDa1.Text)
                StaZe = TimeValue(VoZei.Text)

                NotDa = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "dd.mm.yyyy")
                NotZe = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "hh:mm")
                
                TxNoD.Text = NotDa
                TxNoZ.Text = NotZe

                TagWe = Mid$(TxNoD.Tag, 2, Len(TxNoD.Tag) - 1)
                TxNoD.Tag = "1" & TagWe
                
                TagWe = Mid$(TxNoZ.Tag, 2, Len(TxNoZ.Tag) - 1)
                TxNoZ.Tag = "1" & TagWe
            End If
        End If
    End If
End If

Exit Sub

WoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNoSe" & Err.Number
Resume Next

End Sub
Private Sub FNoVa(Optional ByVal NtSet As Boolean = False)
On Error GoTo LiErr
'Ändert dem Notification Wert

Dim DatSt As String
Dim ZeiSt As String
Dim NotDa As String
Dim NotZe As String
Dim NotSt As String
Dim NotVa As Integer
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTermin
Set TxDa1 = FM.txtDatu1
Set VoZei = FM.txtVonZe
Set CmMit = FM.cmbMitar
Set CmMan = FM.cmbBehan
Set CmNot = FM.cmbNotVa
Set TxNoS = FM.txtNoSta
Set TxNoD = FM.txtNoDat
Set TxNoZ = FM.txtNoTim

Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) Then
        DatSt = Format$(TxDa1.Text, "dd.mm.yyyy")
    End If
End If

If VoZei.Text <> vbNullString Then
    If IsDate(VoZei.Text) Then
        ZeiSt = Format$(VoZei.Text, "hh:mm:ss")
    End If
End If

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

NotDa = Format$(CDate(DateAdd("h", -NotVa, DatSt & " " & ZeiSt)), "dd.mm.yyyy")
NotZe = Format$(CDate(DateAdd("h", -NotVa, DatSt & " " & ZeiSt)), "hh:mm")
NotSt = NotDa & Chr$(32) & NotZe

If NtSet = True Then
    If CmAcs(AD_Termin_Notify).Checked = False Then
        If NotSt <> vbNullString Then
            If IsDate(NotSt) = True Then
                If CDate(NotSt) > Now Then
                    TxNoS.Text = 3 'Senden
                Else
                    TxNoS.Text = 1 'Gesendet
                End If
            Else
                TxNoS.Text = 0 'Nicht Senden
            End If
        Else
            TxNoS.Text = 0 'Nicht Senden
        End If
    Else
        TxNoS.Text = 0 'Nicht Senden
        TxNoD.Text = vbNullString
        TxNoZ.Text = vbNullString
        TagWe = Mid$(TxNoD.Tag, 2, Len(TxNoD.Tag) - 1)
        TxNoD.Tag = "1" & TagWe
        TagWe = Mid$(TxNoZ.Tag, 2, Len(TxNoZ.Tag) - 1)
        TxNoZ.Tag = "1" & TagWe
    End If
Else
    CmNot.ListIndex = NotVa
    If NotSt <> vbNullString Then
        If IsDate(NotSt) = True Then
            If CDate(NotSt) > Now Then
                TxNoS.Text = 3 'Senden
            Else
                TxNoS.Text = 1 'Gesendet
            End If
        Else
            TxNoS.Text = 0 'Nicht Senden
        End If
    Else
        TxNoS.Text = 0 'Nicht Senden
    End If
End If

If CInt(TxNoS.Text) > 2 Then
    CmAcs(AD_Termin_Notify).Checked = True
    CmNot.Enabled = True
    CmNot.ListIndex = NotVa
Else
    CmAcs(AD_Termin_Notify).Checked = False
    CmNot.Enabled = False
    CmNot.ListIndex = 0
End If

TagWe = Mid$(CmNot.Tag, 2, Len(CmNot.Tag) - 1)
CmNot.Tag = 1 & TagWe

TagWe = Mid$(TxNoS.Tag, 2, Len(TxNoS.Tag) - 1)
TxNoS.Tag = 1 & TagWe

GlTSa = True

Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNoVa " & Err.Number
Resume Next

End Sub
Private Sub FOpen()
On Error GoTo AnErr

If TabId = RibTab_Ter_WarZi Then
    Ter_Bea
End If

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FreTe()
On Error GoTo AnErr

Dim TagWe As String

Set FM = frmTermin
Set CmAbr = FM.cmbAbger

TagWe = Mid$(CmAbr.Tag, 2, Len(CmAbr.Tag) - 1)

Ter_Rec

CmAbr.ListIndex = 2

CmAbr.Tag = 1 & TagWe
GlTSa = True

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FReTe " & Err.Number
Resume Next

End Sub

Private Sub FTaEd(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
On Error Resume Next

Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTermin
Set RpCon = FM.repCont1
Set RpSel = RpCon.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If Shift = 0 Then
            Select Case KeyCode
            Case vbKeyF2: RpCon.Navigator.BeginEdit
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
Set RpCon = Nothing

End Sub
Public Sub FTeLo()
On Error GoTo LiErr
'Löscht den termin

Dim TerNr As Long
Dim Frage As Integer
Dim Tit1, Mld1 As String

Set FM = frmTermin
Set TxID2 = FM.txtID2

Tit1 = "Termin Entfernen"
Mld1 = "Möchten Sie den markierten Termin wirklich entfernen?"

If TxID2.Text <> vbNullString Then
    TerNr = TxID2.Text
Else
    TerNr = GlTem
End If

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    S_Term TerNr, 2
    DoEvents
    S_TeLi
    DoEvents
    S_TePi 'Kalndermarker setzen
    DoEvents
    Unload Me
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeLo " & Err.Number
Resume Next

End Sub

Private Sub FZKo1()
On Error GoTo PoErr

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

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZKo1 " & Err.Number
Resume Next

End Sub
Private Sub FZKo2()
On Error GoTo PoErr

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

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZKo2 " & Err.Number
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


Private Sub btnPost3_Click()
    Ter_Poz True
End Sub

Private Sub btnTele8_Click()
On Error Resume Next

Dim TmStr As String

If Me.txtS4F15.Text <> vbNullString Then
    TmStr = SMSTe(Me.txtS4F15.Text) 'Testen des SMS Rufnummernformates
    If TmStr <> vbNullString Then
        SPopu "Richtiges Rufnummernformat", TmStr, IC48_Information
    Else
        SPopu "Falsches Rufnummernformat", "Die eingegebene Rufnummer hat das falsche Format!", IC48_Forbidden
    End If
End If

End Sub

Private Sub cmbAbgeh_Click()

If GlTeF = False Then 'Formular wird geladen
    GlTSa = True
End If

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

Private Sub cmbAbger_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbAbger.SelLength = 0
End Sub

Private Sub cmbArzNr_Click()

TagWe = Mid$(Me.cmbArzNr.Tag, 2, Len(Me.cmbArzNr.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbArzNr.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub cmbArzNr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbBehan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBezei_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.cmbBezei.Text = vbNullString
    End If
End Sub
Private Sub cmbBezei_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GlTeF = False Then 'Formular wird geladen
            FEinf 2
        End If
    End If
End Sub

Private Sub cmbGanzt_Click()
On Error Resume Next

Set CmGan = Me.cmbGanzt
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

If GlTeF = False Then 'Formular wird geladen
    GlTSa = True
    If CmGan.ListIndex = 0 Then
        If VoZei.Text <> vbNullString Then
            If VoZei.Text = "00:00" Then
                VoZei.Text = "08:00"
                VoZei.Tag = 1 & TagWe
            End If
        End If
        If BiZei.Text <> vbNullString Then
            If BiZei.Text = "00:00" Then
                BiZei.Text = "09:00"
                BiZei.Tag = 1 & TagWe
            End If
        End If
    End If
End If

End Sub
Private Sub cmbGesch_Click()

TagWe = Mid$(Me.cmbGesch.Tag, 2, Len(Me.cmbGesch.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbGesch.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub cmbGesch_GotFocus()
    RetWe = SendMessage(Me.cmbGesch.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbGesch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbGesch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbGesch.SelLength = 0
End Sub
Private Sub cmbMitar_Click()

TagWe = Mid$(Me.cmbMitar.Tag, 2, Len(Me.cmbMitar.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    FMita
    FNoVa
    Me.cmbMitar.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub cmbMitar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbNotVa_Click()

TagWe = Mid$(Me.cmbNotVa.Tag, 2, Len(Me.cmbNotVa.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbNotVa.Tag = 1 & TagWe
    GlTSa = True
    FNoSe
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

Private Sub cmbOnlTe_Click()
TagWe = Mid$(Me.cmbOnlTe.Tag, 2, Len(Me.cmbOnlTe.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbOnlTe.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub cmbS4F11_Click()

TagWe = Mid$(Me.cmbS4F11.Tag, 2, Len(Me.cmbS4F11.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbS4F11.Tag = "1" & TagWe
    GlTSa = True
End If

End Sub
Private Sub cmbS4F11_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Ter_Brz
    End If
End Sub
Private Sub cmbS4F11_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbS4F12_Click()

TagWe = Mid$(Me.cmbS4F12.Tag, 2, Len(Me.cmbS4F12.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbS4F12.Tag = "1" & TagWe
    GlTSa = True
End If

End Sub
Private Sub cmbS4F12_KeyPress(KeyAscii As Integer)
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

Private Sub cmbTypen_Click()
    If GlTeF = False Then 'Formular wird geladen
        FKrTy
    End If
End Sub
Private Sub cmbTypen_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlTeF = False Then 'Formular wird geladen
        Select Case Chr$(KeyCode)
        Case "D": FKrTy
        Case "G": FKrTy
        Case "L": FKrTy
        Case "M": FKrTy
        Case "B": FKrTy
        Case "Z": FKrTy
        Case "P": FKrTy
        Case "I": FKrTy
        Case "U": FKrTy
        End Select
    End If
End Sub

Private Sub cmbZiffe_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.cmbZiffe.Text = vbNullString
        Me.cmbBezei.Text = vbNullString
    End If
End Sub

Private Sub cmbZiffe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GlTeF = False Then 'Formular wird geladen
            FEinf 1
        End If
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
    FDatu
End Sub

Private Sub Form_Load()

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 10000
    .ClientMaxWidth = 14000
    .ClientMinHeight = 8300
    .ClientMinWidth = 11600
    .TopMost = True
End With

TabId = RibTab_Ter_Haupt

Set FrmEx = Nothing

End Sub


Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

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
Private Sub repCont2_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FOpen
End Sub

Private Sub txtAdres_Change()

TagWe = Mid$(Me.txtAdres.Tag, 2, Len(Me.txtAdres.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtAdres.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtAdres_GotFocus()
On Error Resume Next

Set TxID0 = Me.txtID0
Set TxAdr = Me.txtAdres

If TxID0.Text <> vbNullString Then
    If TxAdr.Text <> vbNullString Then
        PaStr = TxAdr.Text
    End If
End If

If GlTeF = False Then 'Formular wird geladen
    TxAdr.SelStart = 0
    TxAdr.SelLength = Len(TxAdr.Text)
End If

End Sub
Private Sub txtAdres_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtAdres_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtAdres.SelLength = 0
    End If
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
Else
    If TxID0.Text <> vbNullString Then
        If PaStr <> vbNullString Then
            TxAdr.Text = PaStr
        End If
    End If
End If

End Sub
Private Sub cmbBehan_Click()

TagWe = Mid$(Me.cmbBehan.Tag, 2, Len(Me.cmbBehan.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbBehan.Tag = 1 & TagWe
    GlTSa = True
End If

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
    If KeyCode = vbKeyF2 Then
        Me.txtBisZe.SelLength = 0
    ElseIf KeyCode = vbKeyReturn Then
        FZKo2
    End If
End Sub
Private Sub txtBisZe_LostFocus()
    FZKo2
    FNoSe
End Sub
Private Sub cmbPrior_Click()

TagWe = Mid$(Me.cmbPrior.Tag, 2, Len(Me.cmbPrior.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbPrior.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub cmbPrior_GotFocus()
    RetWe = SendMessage(Me.cmbPrior.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbPrior_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbPrior_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbPrior.SelLength = 0
End Sub

Private Sub cmbRemin_Click()

TagWe = Mid$(Me.cmbRemin.Tag, 2, Len(Me.cmbRemin.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.cmbRemin.Tag = 1 & TagWe
    GlTSa = True
End If

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
On Error Resume Next

Dim CmBrs As XtremeCommandBars.CommandBars

Dim TerNr As Long
Dim IdStu As Integer
Dim Frage As Integer
Dim Mld1, Tit1 As String

Set FM = frmTermin
Set CmBrs = FM.comBar02
Set TxID2 = FM.txtID2
Set CmStu = FM.cmbStatu
Set TxNoS = FM.txtNoSta
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

IdStu = CmStu.ItemData(CmStu.ListIndex)

Tit1 = "E-Mail-Erinnerung deaktivieren"
Mld1 = "Soll die E-Mail-Erinnerung zu diesem Termin deaktviert werden?"

If TxID2.Text <> vbNullString Then
    TerNr = TxID2.Text
Else
    TerNr = GlTem
End If

TagWe = Mid$(CmStu.Tag, 2, Len(CmStu.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    CmStu.Tag = 1 & TagWe
    GlTSa = True
End If
 
If CmAcs(AD_Termin_Notify).Checked = True Then
    If IdStu = 1 Or IdStu = 4 Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            CmAcs(AD_Termin_Notify).Checked = False
            TxNoS.Text = 0 'Nicht Senden
            TagWe = Mid$(TxNoS.Tag, 2, Len(TxNoS.Tag) - 1)
            TxNoS.Tag = 1 & TagWe
            DoEvents
        End If
    End If
End If

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
    FNoSe
End Sub
Private Sub txtDatu2_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtDatu2.SelStart = 0
        Me.txtDatu2.SelLength = Len(Me.txtDatu2.Text)
    End If
End Sub
Private Sub txtDatu2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtDatu2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtDatu2.SelLength = 0
End Sub

Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
    FNoSe
End Sub

Private Sub txtID0_Change()
On Error Resume Next

Set TxID0 = Me.txtID0

TagWe = Mid$(TxID0.Tag, 2, Len(TxID0.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    TxID0.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtNoDat_Change()

TagWe = Mid$(Me.txtNoDat.Tag, 2, Len(Me.txtNoDat.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtNoDat.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtNoSta_Change()

TagWe = Mid$(Me.txtNoSta.Tag, 2, Len(Me.txtNoSta.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtNoSta.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtNoTim_Change()

TagWe = Mid$(Me.txtNoDat.Tag, 2, Len(Me.txtNoDat.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtNoDat.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtRaum1_Change()

TagWe = Mid$(Me.txtRaum1.Tag, 2, Len(Me.txtRaum1.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtRaum1.Tag = 1 & TagWe
    GlTSa = True
End If

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
    If KeyCode = vbKeyF2 Then Me.txtRaum1.SelLength = 0
End Sub

Private Sub txtDatu1_Change()

TagWe = Mid$(Me.txtDatu1.Tag, 2, Len(Me.txtDatu1.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtDatu1.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtDatu2_Change()

TagWe = Mid$(Me.txtDatu2.Tag, 2, Len(Me.txtDatu2.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtDatu2.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtRefNr_Change()

TagWe = Mid$(Me.txtRefNr.Tag, 2, Len(Me.txtRefNr.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtRefNr.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtRefNr_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtRefNr.SelStart = 0
        Me.txtRefNr.SelLength = Len(Me.txtRefNr.Text)
    End If
End Sub
Private Sub txtRefNr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtRefNr_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtRefNr.SelLength = 0
End Sub

Private Sub txtRzAnz_Change()

Set TxRzA = Me.txtRzAnz
Set TxRak = Me.txtRzAkt

TagWe = Mid$(TxRzA.Tag, 2, Len(TxRzA.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    TxRzA.Tag = 1 & TagWe
    If TxRak.Text = vbNullString Then
        TagWe = Mid$(TxRak.Tag, 2, Len(TxRak.Tag) - 1)
        TxRak.Text = 1
        TxRak.Tag = 1 & TagWe
    ElseIf IsNumeric(TxRak.Text) = True Then
        If CInt(TxRak.Text) = 0 Then
            TagWe = Mid$(TxRak.Tag, 2, Len(TxRak.Tag) - 1)
            TxRak.Text = 1
            TxRak.Tag = 1 & TagWe
        End If
    End If
    GlTSa = True
End If

End Sub

Private Sub txtRzAnz_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtRzAnz.SelStart = 0
        Me.txtRzAnz.SelLength = Len(Me.txtRzAnz.Text)
    End If
End Sub

Private Sub txtRzAnz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtRzDat_Change()

TagWe = Mid$(Me.txtRzDat.Tag, 2, Len(Me.txtRzDat.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtRzDat.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtRzDat_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtRzDat.SelStart = 0
        Me.txtRzDat.SelLength = Len(Me.txtRzDat.Text)
    End If
End Sub

Private Sub txtRzDat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtRzNum_Change()

TagWe = Mid$(Me.txtRzNum.Tag, 2, Len(Me.txtRzNum.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtRzNum.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtRzNum_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtRzNum.SelStart = 0
        Me.txtRzNum.SelLength = Len(Me.txtRzNum.Text)
    End If
End Sub

Private Sub txtRzNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRzTex_Change()

TagWe = Mid$(Me.txtRzTex.Tag, 2, Len(Me.txtRzTex.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtRzTex.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtRzTex_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtRzTex.SelStart = 0
        Me.txtRzTex.SelLength = Len(Me.txtRzTex.Text)
    End If
End Sub


Private Sub txtRzTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtS4F01_Change()

TagWe = Mid$(Me.txtS4F01.Tag, 2, Len(Me.txtS4F01.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F01.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtS4F01_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
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
    
TagWe = Mid$(Me.txtS4F02.Tag, 2, Len(Me.txtS4F02.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F02.Tag = 1 & TagWe
    GlTSa = True
    Ter_Brz
End If

End Sub
Private Sub txtS4F02_Click()

TagWe = Mid$(Me.txtS4F02.Tag, 2, Len(Me.txtS4F02.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F02.Tag = 1 & TagWe
    GlTSa = True
    Ter_Brz
End If

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
    
TagWe = Mid$(Me.txtS4F03.Tag, 2, Len(Me.txtS4F03.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F03.Tag = 1 & TagWe
    GlTSa = True
    Ter_Brz
End If

End Sub
Private Sub txtS4F03_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F03.SelStart = 0
        Me.txtS4F03.SelLength = Len(Me.txtS4F03.Text)
    End If
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
    If GlTeF = False Then 'Formular wird geladen
        Ter_Brz
    End If
End Sub

Private Sub txtS4F04_Change()

TagWe = Mid$(Me.txtS4F04.Tag, 2, Len(Me.txtS4F04.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F04.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtS4F04_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F04.SelStart = 0
        Me.txtS4F04.SelLength = Len(Me.txtS4F04.Text)
    End If
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
    If GlTeF = False Then 'Formular wird geladen
        Ter_Brz
    End If
End Sub

Private Sub txtS4F05_Change()

TagWe = Mid$(Me.txtS4F05.Tag, 2, Len(Me.txtS4F05.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F05.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub

Private Sub txtS4F05_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F05.SelStart = 0
        Me.txtS4F05.SelLength = Len(Me.txtS4F05.Text)
    End If
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
    If GlTeF = False Then 'Formular wird geladen
        Ter_Brz
    End If
End Sub

Private Sub txtS4F06_Change()

TagWe = Mid$(Me.txtS4F06.Tag, 2, Len(Me.txtS4F06.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F06.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtS4F06_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F06.SelStart = 0
        Me.txtS4F06.SelLength = Len(Me.txtS4F06.Text)
    End If
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

TagWe = Mid$(Me.txtS4F08.Tag, 2, Len(Me.txtS4F08.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F08.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtS4F08_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F08.SelStart = 0
        Me.txtS4F08.SelLength = Len(Me.txtS4F08.Text)
    End If
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
    If GlTeF = False Then 'Formular wird geladen
        If Me.txtS4F09.Text = vbNullString Then
            Ter_Poz False
        End If
    End If
End Sub

Private Sub txtS4F09_Change()

TagWe = Mid$(Me.txtS4F09.Tag, 2, Len(Me.txtS4F09.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F09.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtS4F09_Click()
    If Me.txtS4F08.Text <> vbNullString Then
        If Me.txtS4F09.Text = vbNullString Then
            Ter_Poz False
        End If
    End If
End Sub
Private Sub txtS4F09_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F09.SelStart = 0
        Me.txtS4F09.SelLength = Len(Me.txtS4F09.Text)
    End If
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

TagWe = Mid$(Me.txtS4F15.Tag, 2, Len(Me.txtS4F15.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F15.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtS4F15_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F15.SelStart = 0
        Me.txtS4F15.SelLength = Len(Me.txtS4F15.Text)
    End If
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

TagWe = Mid$(Me.txtS4F16.Tag, 2, Len(Me.txtS4F16.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F16.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtS4F16_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F16.SelStart = 0
        Me.txtS4F16.SelLength = Len(Me.txtS4F16.Text)
    End If
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
    Case vbKeyDown: Me.txtS4F01.SetFocus
    Case vbKeyUp: Me.txtS4F16.SetFocus
    End Select
End Sub

Private Sub txtS4F18_Change()

TagWe = Mid$(Me.txtS4F18.Tag, 2, Len(Me.txtS4F18.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    Me.txtS4F18.Tag = 1 & TagWe
    GlTSa = True
End If

End Sub
Private Sub txtS4F18_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F18.SelStart = 0
        Me.txtS4F18.SelLength = Len(Me.txtS4F18.Text)
    End If
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
    Case vbKeyDown: Me.txtS4F20.SetFocus
    Case vbKeyUp: Me.cmbS4F11.SetFocus
    End Select
End Sub

Private Sub txtS4F20_GotFocus()
    If GlTeF = False Then 'Formular wird geladen
        Me.txtS4F20.SelStart = 0
        Me.txtS4F20.SelLength = Len(Me.txtS4F20.Text)
    End If
End Sub

Private Sub txtS4F20_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtS4F20_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtS4F20.SelLength = 0
    Case vbKeyDown: Me.txtS4F15.SetFocus
    Case vbKeyUp: Me.txtS4F18.SetFocus
    End Select
End Sub

Private Sub txtS4F20_Validate(Cancel As Boolean)
    If (Not txtS4F20.isValid) Then Cancel = True
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
    If KeyCode = vbKeyF2 Then
        Me.txtVonZe.SelLength = 0
    ElseIf KeyCode = vbKeyReturn Then
        FZKo1
    End If
End Sub
Private Sub txtVonZe_LostFocus()
    FZKo1
    FNoSe
End Sub

Private Sub FErin()
On Error Resume Next

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTermin
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

Set FM = frmTermin
Set TxFar = FM.txtFarbe

TxFar.Text = Flag

TagWe = Mid$(TxFar.Tag, 2, Len(TxFar.Tag) - 1)
TxFar.Tag = 1 & TagWe

GlTSa = True

TeFarb Flag, 1

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case TabId
Case RibTab_Ter_Haupt:
    TeTit = IniGetOpt("Hilfe", 50701)
    TeMai = IniGetOpt("Hilfe", 50702)
    TeInh = IniGetOpt("Hilfe", 50703)
    TeFus = IniGetOpt("Hilfe", 50704)
Case RibTab_Ter_Adres:
    TeTit = IniGetOpt("Hilfe", 51011)
    TeMai = IniGetOpt("Hilfe", 51012)
    TeInh = IniGetOpt("Hilfe", 51013)
    TeFus = IniGetOpt("Hilfe", 51014)
Case RibTab_Ter_Leist:
    TeTit = IniGetOpt("Hilfe", 51021)
    TeMai = IniGetOpt("Hilfe", 51022)
    TeInh = IniGetOpt("Hilfe", 51023)
    TeFus = IniGetOpt("Hilfe", 51024)
Case RibTab_Ter_WarZi:
    TeTit = IniGetOpt("Hilfe", 51031)
    TeMai = IniGetOpt("Hilfe", 51032)
    TeInh = IniGetOpt("Hilfe", 51033)
    TeFus = IniGetOpt("Hilfe", 51034)
Case RibTab_Ter_Proto:
    TeTit = IniGetOpt("Hilfe", 51041)
    TeMai = IniGetOpt("Hilfe", 51042)
    TeInh = IniGetOpt("Hilfe", 51043)
    TeFus = IniGetOpt("Hilfe", 51044)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    TMeAc True
    Set frmTermin = Nothing
End Sub

Private Sub FSave(Optional ByVal TeClo As Boolean = False, Optional ByVal AdImp As Boolean = False)
On Error GoTo SaErr
'Überprüft, ob der Eintrag geändert wurde und speichert dieses ab

Dim NeuDa As Date
Dim RowNr As Long
Dim RmuNr As Long
Dim PatNr As Long
Dim Frage As Integer
Dim AktZa As Integer
Dim GesZa As Integer
Dim Dialo As Boolean
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CaCol As XtremeCalendarControl.CalendarControl

Set FM = frmMain
Set CaCol = FM.calCont1
Set RpCo1 = FM.repCont1
Set TxID0 = Me.txtID0
Set TxAdr = Me.txtAdres
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmRmu = Me.cmbRaum1
Set CmMan = Me.cmbBehan
Set TxAnz = Me.txtAnzal
Set TxMul = Me.txtMulti
Set TxEin = Me.txtEinze
Set CmETy = Me.cmbTypen
Set CmZif = Me.cmbZiffe
Set CmBez = Me.cmbBezei
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

If GlTza = True Then 'Testzeit abgelaufen
    SPopu "Lizenzierung erforderlich!", "Es ist keine bzw. keine gültige Seriennummer vorhanden oder die Testzeit ist abgelaufen.", IC48_Forbidden
    Exit Sub
End If

Tit1 = "Termin Speichern"
GesZa = UBound(GlRmu)

If Left$(VoZei.Tag, 1) = "1" Then
    FZKo1
    DoEvents
End If

If Left$(BiZei.Tag, 1) = "1" Then
    FZKo2
    DoEvents
End If

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

If CmRmu.ListCount > 0 Then
    If CmRmu.ListIndex >= 0 Then
        RmuNr = CmRmu.ItemData(CmRmu.ListIndex)
    Else
        RmuNr = 1
    End If
Else
    RmuNr = 1
End If

GlSeF = True

If AdImp = True Then
    PatNr = Ter_AdI() 'Terminadressenimport
    DoEvents
    Adr_Er1 , PatNr
    DoEvents
    CmAcs(TE_Adresse_Ubertrag).Enabled = False
    CmAcs(TE_Adresse_Bearbeit).Enabled = True
    TxID0.Text = PatNr
    DoEvents
End If

If GlTSa = True Then
    If CmBet.Text = vbNullString Then
        If TxAdr.Text = vbNullString Then
            Mld1 = "Das Feld Betreff muß erst ausgefüllt werden, damit dieser Datensatz gespeichern werden kann"
            Tit1 = "Fehlende Angaben"
            SPopu Tit1, Mld1, IC48_Forbidden
            Exit Sub
        End If
    End If

    If TxID0.Text <> vbNullString Then
        If TxAdr.Text = vbNullString Then
            If PaStr <> vbNullString Then
                TxAdr.Text = PaStr
            End If
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    Ter_Brz
    DoEvents
    
    If GlTeN = True Then 'neuen Termin hinzufügen
        If Ter_San() = True Then
            With GlSuT
                .SuIdx = -1
                .GuiID = GlTeG 'Termin GuiID
            End With
            DoEvents
            GlTSa = False
            SSuch
            GlTeN = False
        Else
            Dialo = True
        End If
    Else
        If Ter_Sav() = True Then
            DoEvents
            GlTSa = False
        Else
            Dialo = True
        End If
       
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpTe RowNr
        End If

        For Each AktCo In Me.Controls
            If AktCo.Tag <> vbNullString Then
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                AktCo.Tag = 0 & TagWe
            End If
        Next AktCo
    End If

    If TxID0.Text <> vbNullString Then
        If IsNumeric(TxID0.Text) = True Then
            If CLng(TxID0.Text) > 0 Then
                PatNr = CLng(TxID0.Text)
            End If
        End If
    End If
        
    If PatNr > 0 Then
        CmAcs(AD_Termin_Ketten).Enabled = True
        CmAcs(AD_Termin_StaKett).Enabled = True
        CmAcs(AD_Termin_Abrechnen).Enabled = True
        CmAcs(AD_Termin_EintLoe).Enabled = True
        TxAnz.Enabled = True
        TxMul.Enabled = True
        TxEin.Enabled = True
        CmETy.Enabled = True
        CmZif.Enabled = True
        CmBez.Enabled = True
        DoEvents
        Ter_Com
        DoEvents
        Ter_Edi PatNr, False 'aus Warteliste entfernen
        DoEvents
        P_List "TeDe", 0, 2
    End If
    DoEvents
        
    S_TeLi
    DoEvents
    S_TePi 'Kalndermarker setzen
    DoEvents
    
    If GlBut = RibTab_Startseite Then
        STaSe ShoCut_Termin, RibTab_Ter_Kalend
        If GlBut <> RibTab_Ter_Kalend Then
            GlBut = RibTab_Ter_Kalend
            SButt
            SButD
            SBuLa
            DoEvents
            SPosi
        End If
    End If
    
End If

Screen.MousePointer = vbNormal

GlSeF = False

If Dialo = False Then
    If TeClo = False Then
        Unload Me
    Else
        CmAcs(AD_Termin_Close).Enabled = False
        CmAcs(AD_Termin_Delete).Enabled = False
    End If
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Set CmBrs = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: TeNew
Case KY_F5: frmAdrSuch.Show vbModal
Case KY_F7: TeOut
Case KY_F8: FSave
Case KY_F10: FDruck
Case KY_F11: Unload Me
Case TE_Termin_Hinzufu: TeNew
Case TE_Termin_Loeschen: FTeLo
Case TE_Termin_Speichern: FSave
Case TE_Termin_Beenden: Unload Me
Case TE_Adresse_Hinzufu: SAdre 1
Case TE_Adresse_Bearbeit: FAdre
Case TE_Adresse_Suchen: frmAdrSuch.Show vbModal
Case TE_Adresse_Ubertrag: FSave True, True
Case TE_Termin_Hilfe: FHilfe
Case TE_Termin_Drucken: FDruck
Case SY_TE_Termin_Outlook: TeOut
Case AD_Termin_Copy: STeKo
Case AD_Termin_Add: TeNew
Case AD_Termin_Delete: FTeLo
Case AD_Termin_Save: FSave True
Case AD_Termin_Close: FSave
Case AD_Termin_Remind: FErin
Case AD_Termin_Notify: FNoVa True
Case AD_Termin_Ketten: FKata
Case AD_Termin_StaKett: FEiKe
Case AD_Termin_StaKet2: FEiKe 2
Case AD_Termin_Abrechnen: FreTe
Case AD_Termin_EintLoe: FGeLo
Case AD_Termin_TermID: FExpo
Case AD_Termin_ProEmail: FEmai
Case AD_Termin_WarSet: Ter_Bea
Case AD_Termin_WarNeu: frmAdrSuch.Show vbModal
Case AD_Termin_WarDel: FTWal
Case AD_Termin_Clip1: FClip 1
Case AD_Termin_Clip2: FClip 2
Case SY_TE_Termin_Docume:
Case SY_TE_Termin_EmlBes: TeNach 6
Case SY_TE_Termin_EmlEri: TeNach 7
Case SY_TE_Termin_EmlAbs: TeNach 8
Case SY_TE_Termin_EmlVrs: TeNach 9
Case SY_TE_Termin_SMSBes: TeNach 10
Case SY_TE_Termin_SMSEri: TeNach 11
Case SY_TE_Termin_SMSAbs: TeNach 12
Case SY_TE_Termin_SMSVrs: TeNach 13
Case SY_TE_Termin_EmlSto: TeNach 15
Case SY_TE_Termin_SMSSto: TeNach 16
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
Private Sub FTWal()
On Error GoTo LoErr
'Eintrag aus Warteliste entfernen

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim BaItm As XtremeShortcutBar.ShortcutBarItem
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Dim AnzPo As Long
Dim IdxNr As Long
Dim Frage As Integer
Dim Mld1, Tit1 As String

Set FM = frmTermin
Set CmBrs = FM.comBar02
Set RbBar = CmBrs.Item(1)
Set RpCo2 = FM.repCont2
Set RpCls = RpCo2.Columns
Set RpRcs = RpCo2.Records
Set RpSel = RpCo2.SelectedRows

AnzPo = RpSel.Count

Tit1 = "Wartenden Entfernen"
Mld1 = "Möchten Sie die markierte Adresse wirklich entfernen?"

If GlRch(0, 15) = 0 Then
    WindowMess "Sie besitzen keine Berechtigung für diesen Vorgang", Dial3, Tit1, FM.hwnd
    Exit Sub
End If

If AnzPo > 0 Then
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Adr_ID0)
            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
            Ter_Edi IdxNr, False 'aus Warteliste entfernen
            DoEvents
            Ter_WaL
        End If
    End If
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpRcs = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

LoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTWal " & Err.Number
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

Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTermin
Set CmBrs = FM.comBar02
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

TabId = RbTab.id

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2
DoEvents

Select Case TabId
Case RibTab_Ter_Haupt:
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
    RpCo1.Visible = False
    RpCo2.Visible = False
Case RibTab_Ter_Adres:
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    RpCo1.Visible = False
    RpCo2.Visible = False
Case RibTab_Ter_Leist:
    Rahm1.Visible = False
    Rahm2.Visible = True
    Rahm3.Visible = False
    RpCo1.Visible = True
    RpCo2.Visible = False
Case RibTab_Ter_WarZi:
    TeSpa TabId
    DoEvents
    Ter_WaL
    DoEvents
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    RpCo1.Visible = False
    RpCo2.Visible = True
Case RibTab_Ter_Proto:
    TeSpa TabId
    DoEvents
    Ter_SeD
    DoEvents
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    RpCo1.Visible = False
    RpCo2.Visible = True
End Select

DoEvents
clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set RpCo1 = Nothing
Set RpCo2 = Nothing
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
    GlTeF = False 'Formular wird geladen
End Sub

Private Sub txtKomme_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.txtKomme.SelLength = 0
End Sub

Private Sub cmbRaum1_Click()
On Error Resume Next

Dim RmuNr As Long
Dim MitNr As Long
Dim MiIdx As Integer
Dim AkZa1 As Integer
Dim AkZa2 As Integer
Dim Mld1, Tit1 As String

Set CmRmu = Me.cmbRaum1
Set CmMit = Me.cmbMitar

Tit1 = "Mitarbeiteranpassung"
Mld1 = "Der dem ausgewählten Raum zugeordnete Mitarbeiter wurde aktualisiert"

RmuNr = CmRmu.ItemData(CmRmu.ListIndex)

If GlTeF = False Then
    For AkZa1 = 1 To UBound(GlRmu)
        If RmuNr = GlRmu(AkZa1, 2) Then
            MitNr = GlRmu(AkZa1, 4)
            If MitNr > 0 Then
                For AkZa2 = 1 To UBound(GlMiT)
                    If MitNr = GlMiT(AkZa2, 2) Then
                        MiIdx = AkZa2 - 1
                        CmMit.ListIndex = MiIdx
                        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
                        CmMit.Tag = 1 & TagWe
                        SPopu Tit1, Mld1, IC48_Information
                        Exit For
                    End If
                Next AkZa2
            End If
            Exit For
        End If
    Next AkZa1
End If

TagWe = Mid$(CmRmu.Tag, 2, Len(CmRmu.Tag) - 1)

If GlTeF = False Then 'Formular wird geladen
    CmRmu.Tag = 1 & TagWe
    GlTSa = True
End If

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
Dim MaIdx As Integer
Dim MiIdx As Integer
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
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If GlMiT(MiIdx, 8) > 0 Then
        ZeiRa = GlMiT(MiIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
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
    If GlMaT(MaIdx, 8) > 0 Then
        ZeiRa = GlMaT(MaIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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
        Else
            If TmVon >= AlDa2 Then
                TmBis = DateAdd("n", MiDif, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
                TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)
                BiZei.Tag = 1 & TagWe
            End If
        End If
        DoEvents
        FNoSe
    End If
End If

End Sub
Private Sub updCont2_UpClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
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
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If GlMiT(MiIdx, 8) > 0 Then
        ZeiRa = GlMiT(MiIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
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
    If GlMaT(MaIdx, 8) > 0 Then
        ZeiRa = GlMaT(MaIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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
            If TmVon >= AlDa2 Then
                TmBis = DateAdd("n", MiDif, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
                TagWe = Mid$(BiZei.Tag, 2, Len(BiZei.Tag) - 1)
                BiZei.Tag = 1 & TagWe
            End If
        End If
        DoEvents
        FNoSe
    End If
End If

End Sub
Private Sub updCont3_DownClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
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
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If GlMiT(MiIdx, 8) > 0 Then
        ZeiRa = GlMiT(MiIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
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
    If GlMaT(MaIdx, 8) > 0 Then
        ZeiRa = GlMaT(MaIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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
            If TmBis <= AlDa1 Then
                TmVon = DateAdd("n", -MiDif, AlDa1)
                VoZei.Text = Format$(TmVon, "hh:mm")
                TagWe = Mid$(VoZei.Tag, 2, Len(VoZei.Tag) - 1)
                VoZei.Tag = 1 & TagWe
            End If
        End If
        DoEvents
        FNoSe
    End If
End If

End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlTeF = False Then 'Formular wird geladen
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    TePos
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub


Private Sub updCont3_UpClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
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
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If GlMiT(MiIdx, 8) > 0 Then
        ZeiRa = GlMiT(MiIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
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
    If GlMaT(MaIdx, 8) > 0 Then
        ZeiRa = GlMaT(MaIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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
        DoEvents
        FNoSe
    End If
End If

End Sub
Private Sub txtAnzal_GotFocus()
    Me.txtAnzal.SelStart = 0
    Me.txtAnzal.SelLength = Len(Me.txtAnzal.Text)
End Sub
Private Sub txtAnzal_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.txtMulti.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        If GlTeF = False Then 'Formular wird geladen
            FEinf 3
        End If
    End If
End Sub
Private Sub txtEinze_GotFocus()
    Me.txtEinze.SelStart = 0
    Me.txtEinze.SelLength = Len(Me.txtEinze.Text)
End Sub
Private Sub txtEinze_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.txtDatu1.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        If GlTeF = False Then 'Formular wird geladen
            FEinf 3
        End If
    End If
End Sub
Private Sub txtMulti_GotFocus()
    Me.txtMulti.SelStart = 0
    Me.txtMulti.SelLength = Len(Me.txtMulti.Text)
End Sub
Private Sub txtMulti_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.txtEinze.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        If GlTeF = False Then 'Formular wird geladen
            FEinf 3
        End If
    End If
End Sub
