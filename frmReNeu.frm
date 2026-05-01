VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmReNeu 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnung Hinzufügen"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   75
      Top             =   6000
      Width           =   6600
      _Version        =   1048579
      _ExtentX        =   11642
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4600
         TabIndex        =   76
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
      Begin XtremeSuiteControls.PushButton btnWeite 
         Default         =   -1  'True
         Height          =   400
         Left            =   3200
         TabIndex        =   77
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
      Begin XtremeSuiteControls.PushButton btnZuruk 
         Height          =   400
         Left            =   1800
         TabIndex        =   78
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
         Left            =   500
         TabIndex        =   79
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
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   600
      Left            =   800
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   8000
      Visible         =   0   'False
      Width           =   600
      _Version        =   1048579
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   5800
      Left            =   400
      TabIndex        =   43
      Top             =   100
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   10231
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkPhyAb 
         Height          =   255
         Left            =   1300
         TabIndex        =   5
         Top             =   2100
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ergänzungsdaten erfassen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkReNum 
         Height          =   255
         Left            =   1300
         TabIndex        =   4
         Top             =   1700
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Belegnummer jetzt erzeugen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   310
         Left            =   1300
         TabIndex        =   6
         Top             =   3000
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbReStu 
         Height          =   315
         Left            =   1300
         TabIndex        =   9
         Top             =   5100
         Width           =   2100
         _Version        =   1048579
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   1300
         TabIndex        =   7
         Top             =   3700
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkDiaRe 
         Height          =   255
         Left            =   1300
         TabIndex        =   3
         Top             =   1300
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rechnungsdiagnose übernehmen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkDiaKr 
         Height          =   255
         Left            =   1300
         TabIndex        =   2
         Top             =   920
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Krankenblattdiagnose übernehmen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox txtKomFe 
         Height          =   310
         Left            =   1300
         TabIndex        =   8
         Top             =   4400
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   195
         Left            =   1320
         TabIndex        =   70
         Top             =   4150
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Rechnungsfreitext :"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   610
         Left            =   800
         TabIndex        =   63
         Top             =   100
         Width           =   4800
         _Version        =   1048579
         _ExtentX        =   8467
         _ExtentY        =   1076
         _StockProps     =   79
         Caption         =   "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Rechnung anlegen? Bitte wählen Sie die gewünschten Optionen."
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   200
         Left            =   1320
         TabIndex        =   62
         Top             =   4850
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Steuersatz :"
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab04 
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   200
         Left            =   1320
         TabIndex        =   60
         Top             =   3450
         Width           =   1395
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   200
         Left            =   1320
         TabIndex        =   44
         Top             =   2750
         Width           =   1400
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   5800
      Left            =   400
      TabIndex        =   12
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   10231
      _StockProps     =   79
      Caption         =   "GroupBox2"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtBemer 
         Height          =   350
         Left            =   900
         TabIndex        =   17
         Top             =   2580
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtPost 
         Height          =   350
         Left            =   900
         TabIndex        =   16
         Top             =   1860
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   350
         Left            =   900
         TabIndex        =   15
         Top             =   1150
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin VB.Label lblLab19 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte geben Sie das gewünschte Suchkriterium ein und klicken auf Weiter oder bestätigen mit der ENTER-Taste."
         Height          =   450
         Left            =   700
         TabIndex        =   59
         Top             =   100
         Width           =   5100
      End
      Begin VB.Label lblLab22 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Anmerkung"
         Height          =   195
         Left            =   910
         TabIndex        =   18
         Top             =   2340
         Width           =   3000
      End
      Begin VB.Label lblLab21 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Postleitzahl"
         Height          =   195
         Left            =   910
         TabIndex        =   14
         Top             =   1620
         Width           =   3000
      End
      Begin VB.Label lblLab20 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Kurzbezeichnung"
         Height          =   195
         Left            =   910
         TabIndex        =   13
         Top             =   920
         Width           =   3000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   5800
      Left            =   400
      TabIndex        =   1
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   10231
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListBox lstList1 
         Height          =   4400
         Left            =   500
         TabIndex        =   11
         Top             =   1150
         Width           =   4500
         _Version        =   1048579
         _ExtentX        =   7937
         _ExtentY        =   7761
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         SelectionBackColor=   14925219
         SelectionForeColor=   4473924
      End
      Begin VB.Label lblLab24 
         BackStyle       =   0  'Transparent
         Caption         =   "Gefundene Einträge :"
         Height          =   200
         Left            =   510
         TabIndex        =   51
         Top             =   920
         Width           =   1600
      End
      Begin VB.Label lblLab23 
         BackStyle       =   0  'Transparent
         Caption         =   "Folgende Einträge wurden gefunden. Bitte wählen Sie den gewünschten Patienten und klicken auf Weiter."
         Height          =   450
         Left            =   700
         TabIndex        =   10
         Top             =   100
         Width           =   5100
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   8000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtIdKur 
      Height          =   195
      Left            =   240
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   8000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPatWa 
      Height          =   195
      Left            =   480
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   8000
      Width           =   195
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   5805
      Left            =   400
      TabIndex        =   19
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   10239
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont5 
         Height          =   920
         Left            =   700
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4800
         Width           =   4400
         _Version        =   1048579
         _ExtentX        =   7761
         _ExtentY        =   1623
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   4800
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2570
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
         BuddyControl    =   "txtReKop"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtReKop 
         Height          =   350
         Left            =   3500
         TabIndex        =   27
         Top             =   2570
         Width           =   1290
         _Version        =   1048579
         _ExtentX        =   2275
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "1"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   4810
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   1150
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
         TabIndex        =   21
         Top             =   1150
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbZaZie 
         Height          =   310
         Left            =   700
         TabIndex        =   24
         Top             =   1860
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4392
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbReTyp 
         Height          =   315
         Left            =   700
         TabIndex        =   26
         Top             =   2570
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4392
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   3500
         TabIndex        =   22
         Top             =   1150
         Width           =   1290
         _Version        =   1048579
         _ExtentX        =   2275
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbVersi 
         Height          =   315
         Left            =   705
         TabIndex        =   31
         Top             =   3980
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4419
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbBeTyp 
         Height          =   315
         Left            =   700
         TabIndex        =   29
         Top             =   3290
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4392
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "0"
      End
      Begin XtremeSuiteControls.ComboBox cmbReWar 
         Height          =   315
         Left            =   3495
         TabIndex        =   25
         Top             =   1860
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbBeGru 
         Height          =   315
         Left            =   3495
         TabIndex        =   30
         Top             =   3285
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "0"
      End
      Begin XtremeSuiteControls.ComboBox cmbVersa 
         Height          =   315
         Left            =   3495
         TabIndex        =   32
         Top             =   3975
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab32 
         Height          =   210
         Left            =   3510
         TabIndex        =   74
         Top             =   3740
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Versandart :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab31 
         Height          =   210
         Left            =   3510
         TabIndex        =   73
         Top             =   3040
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Behandlungsgrund :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab10 
         Height          =   210
         Left            =   3510
         TabIndex        =   66
         Top             =   900
         Width           =   1700
         _Version        =   1048579
         _ExtentX        =   2999
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Rechnungsdatum :"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab09 
         Height          =   210
         Left            =   710
         TabIndex        =   65
         Top             =   900
         Width           =   1900
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Rechnungsnummer :"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   600
         Left            =   700
         TabIndex        =   64
         Top             =   100
         Width           =   5100
         _Version        =   1048579
         _ExtentX        =   8996
         _ExtentY        =   1058
         _StockProps     =   79
         Caption         =   $"frmReNeu.frx":0000
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLab18 
         BackStyle       =   0  'Transparent
         Caption         =   "Rechnungsnummernbasis :"
         Height          =   210
         Left            =   710
         TabIndex        =   61
         Top             =   4560
         Width           =   2000
      End
      Begin VB.Label lblLab16 
         BackStyle       =   0  'Transparent
         Caption         =   "Empfängertyp :"
         Height          =   210
         Left            =   710
         TabIndex        =   50
         Top             =   3040
         Width           =   1305
      End
      Begin VB.Label lblLab15 
         BackStyle       =   0  'Transparent
         Caption         =   "Katalog :"
         Height          =   210
         Left            =   705
         TabIndex        =   49
         Top             =   3740
         Width           =   1395
      End
      Begin VB.Label lblLab14 
         BackStyle       =   0  'Transparent
         Caption         =   "Währung :"
         Height          =   210
         Left            =   3510
         TabIndex        =   46
         Top             =   1620
         Width           =   1700
      End
      Begin VB.Label lblLab13 
         BackStyle       =   0  'Transparent
         Caption         =   "Belegtyp :"
         Height          =   210
         Left            =   710
         TabIndex        =   42
         Top             =   2330
         Width           =   1395
      End
      Begin VB.Label lblLab11 
         BackStyle       =   0  'Transparent
         Caption         =   "Zahlungsziel :"
         Height          =   210
         Left            =   710
         TabIndex        =   41
         Top             =   1620
         Width           =   1600
      End
      Begin VB.Label lblLab12 
         BackStyle       =   0  'Transparent
         Caption         =   "Ausdrucke :"
         Height          =   210
         Left            =   3510
         TabIndex        =   20
         Top             =   2330
         Width           =   1700
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   5800
      Left            =   400
      TabIndex        =   47
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   10231
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkTheBe 
         Height          =   255
         Left            =   1300
         TabIndex        =   35
         Top             =   1700
         Width           =   3105
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Therapie beendet"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkGrThe 
         Height          =   255
         Left            =   1300
         TabIndex        =   34
         Top             =   1300
         Width           =   3105
         _Version        =   1048579
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Gruppentherapie"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   350
         Left            =   3120
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Öffnet den Auswahlkalender"
         Top             =   2580
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   350
         Left            =   1300
         TabIndex        =   36
         Top             =   2580
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbArzNr 
         Height          =   315
         Left            =   1300
         TabIndex        =   39
         Top             =   3990
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtRzNum 
         Height          =   350
         Left            =   1300
         TabIndex        =   38
         Top             =   3290
         Width           =   2100
         _Version        =   1048579
         _ExtentX        =   3704
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRzTex 
         Height          =   660
         Left            =   1300
         TabIndex        =   40
         ToolTipText     =   "4740"
         Top             =   4690
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   1164
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.Label lblLab30 
         Height          =   195
         Left            =   1320
         TabIndex        =   72
         Top             =   4440
         Width           =   1905
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Verordnungstext :"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab28 
         Height          =   195
         Left            =   1320
         TabIndex        =   71
         Top             =   3040
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Belegnummer :"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab27 
         Height          =   195
         Left            =   1320
         TabIndex        =   69
         Top             =   3740
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Verordner :"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab26 
         Height          =   195
         Left            =   1320
         TabIndex        =   68
         Top             =   2340
         Width           =   1905
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Verordnungsdatum :"
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLab25 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte erfassen Sie die ergänzenden Rechnungsdaten. Diese werden standardmäßig auf dem Rechnungsformular mit ausgegeben."
         Height          =   640
         Left            =   700
         TabIndex        =   48
         Top             =   100
         Width           =   5100
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm6 
      Height          =   5800
      Left            =   400
      TabIndex        =   52
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   10231
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.RadioButton optBehLe 
         Height          =   220
         Left            =   1200
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   4850
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Letzter Behandlungstag kopieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optBehEr 
         Height          =   220
         Left            =   1200
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   4500
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Erster Behandlungstag kopieren"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu2 
         Height          =   2595
         Left            =   100
         TabIndex        =   53
         Top             =   1150
         Width           =   5500
         _Version        =   1048579
         _ExtentX        =   9701
         _ExtentY        =   4577
         _StockProps     =   64
         Show3DBorder    =   0
         ColumnCount     =   2
      End
      Begin XtremeSuiteControls.CheckBox chkExpMo 
         Height          =   220
         Left            =   1200
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4000
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Multiselektion in Kalenderauswahl"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab29 
         Height          =   610
         Left            =   700
         TabIndex        =   67
         Top             =   100
         Width           =   5100
         _Version        =   1048579
         _ExtentX        =   8996
         _ExtentY        =   1076
         _StockProps     =   79
         Caption         =   $"frmReNeu.frx":00BD
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmReNeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control

Private Lbl02 As XtremeSuiteControls.Label
Private Lbl05 As XtremeSuiteControls.Label
Private Lbl08 As XtremeSuiteControls.Label
Private Lbl09 As XtremeSuiteControls.Label
Private Lbl10 As XtremeSuiteControls.Label
Private Lbl29 As XtremeSuiteControls.Label
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxKur As XtremeSuiteControls.FlatEdit
Private TxPLZ As XtremeSuiteControls.FlatEdit
Private TxBem As XtremeSuiteControls.FlatEdit
Private TxRen As XtremeSuiteControls.FlatEdit
Private TxKop As XtremeSuiteControls.FlatEdit
Private TxDum As XtremeSuiteControls.FlatEdit
Private TxWar As XtremeSuiteControls.FlatEdit
Private TxRec As XtremeSuiteControls.FlatEdit
Private TxRzn As XtremeSuiteControls.FlatEdit
Private TxRzt As XtremeSuiteControls.FlatEdit
Private FLis1 As XtremeSuiteControls.ListBox
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private CmKom As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmZil As XtremeSuiteControls.ComboBox
Private CmStu As XtremeSuiteControls.ComboBox
Private CmWar As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmArz As XtremeSuiteControls.ComboBox
Private CmVer As XtremeSuiteControls.ComboBox
Private CmBTy As XtremeSuiteControls.ComboBox
Private CmGru As XtremeSuiteControls.ComboBox
Private CmVrs As XtremeSuiteControls.ComboBox
Private ChRnm As XtremeSuiteControls.CheckBox
Private ChExp As XtremeSuiteControls.CheckBox
Private ChDiR As XtremeSuiteControls.CheckBox
Private ChDiK As XtremeSuiteControls.CheckBox
Private ChPhy As XtremeSuiteControls.CheckBox
Private ChGru As XtremeSuiteControls.CheckBox
Private ChThe As XtremeSuiteControls.CheckBox
Private OpBEr As XtremeSuiteControls.RadioButton
Private OpBLe As XtremeSuiteControls.RadioButton
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private PuBu4 As XtremeSuiteControls.PushButton
Private DaPi1 As XtremeCalendarControl.DatePicker
Private DaPi2 As XtremeCalendarControl.DatePicker
Private RpRow As XtremeReportControl.ReportRow
Private RpSel As XtremeReportControl.ReportSelectedRows
Private ImMan As XtremeCommandBars.ImageManager

Private mPaNr As Long 'Patientennummer
Private mReNr As Long 'Rechnungsnummer
Private mMaNr As Long 'Mandantennummer
Private ReTyp As String
Private mPaKu As String
Private mReSt As String
Private mReBe As Single
Private mReBz As Single
Private mPaGu As Single
Private KalWa As Integer
Private FStSa As Integer
Private AbExp As Boolean
Private ReAnp As Boolean
Private FoLad As Boolean

Private Sub FCapt()
On Error GoTo SuErr

Set Lbl02 = Me.lblLab02
Set Lbl05 = Me.lblLab05
Set Lbl08 = Me.lblLab08
Set Lbl09 = Me.lblLab09
Set Lbl10 = Me.lblLab10
Set Lbl29 = Me.lblLab29
Set ChDiR = Me.chkDiaRe
Set ChDiK = Me.chkDiaKr
Set ChRnm = Me.chkReNum
Set CmZil = Me.cmbZaZie
Set CmVer = Me.cmbVersi
Set CmTyp = Me.cmbReTyp

If ChDiK.Enabled = False Then ChDiK.Enabled = True
If ChDiR.Enabled = False Then ChDiR.Enabled = True
If CmZil.Enabled = False Then CmZil.Enabled = True
If CmVer.Enabled = False Then CmVer.Enabled = True

Select Case UCase(ReTyp)
Case "R":
    If GlRKo = True Then 'Rechnung Kopieren
        Me.Caption = "Rechnung kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Rechnung kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Rechnung hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Rechnung anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Der neuen Rechnung wird folgende Rechnungsnummer und folgendes Rechnungsdatum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des ersten Behandlungstages der vorherigen Rechnung in die neue Rechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Rechnungsfreitext :"
    Lbl09.Caption = "Rechnungsnummer :"
    Lbl10.Caption = "Rechnungsdatum :"
    ChRnm.Caption = "Rechnungsnummer jetzt erzeugen"
Case "V":
    If GlRKo = True Then
        Me.Caption = "Kostenvoranschlag kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie einen neuen Kostenvoranschlag kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Kostenvoranschlag hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie einen neuen Kostenvoranschlag anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Dem neuen Kostenvoranschlag wird folgende Nummer und folgendes Datum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des ersten Behandlungstages der vorherigen Rechnung in die neue Rechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Kostenvoranschlagfreitext :"
    Lbl09.Caption = "Voranschlagsnummer :"
    Lbl10.Caption = "Voranschlagsdatum :"
    ChRnm.Caption = "Voranschlagsnummer jetzt erzeugen"
Case "L":
    If GlRKo = True Then
        Me.Caption = "Laborrechnung kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Laborrechnung kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Laborrechnung hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Laborrechnung anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Der neuen Laborrechnung wird folgende Rechnungsnummer und folgendes Datum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des ersten Behandlungstages der vorherigen Rechnung in die neue Rechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Laborrechnungsfreitext :"
    Lbl09.Caption = "Laborrechnungsnummer :"
    Lbl10.Caption = "Laborrechnungsdatum :"
    ChRnm.Caption = "Laborrechnungsnummer jetzt erzeugen"
Case "A":
    If GlRKo = True Then
        Me.Caption = "Exportrechnung kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Exportrechnung kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Exportrechnung hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Exportrechnung anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Der neuen Exportrechnung wird folgende Rechnungsnummer und folgendes Datum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des ersten Behandlungstages der vorherigen Rechnung in die neue Rechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Exportrechnungsfreitext :"
    Lbl09.Caption = "Exportrechnungsnummer :"
    Lbl10.Caption = "Exportrechnungsdatum :"
    ChRnm.Caption = "Exportrechnungsnummer jetzt erzeugen"
Case "U":
    If GlRKo = True Then
        Me.Caption = "Gutschrift kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Gutschrift kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Gutschrift hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Gutschrift anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Der neuen Gutschrift wird folgende Nummer und folgendes Datum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des ersten Behandlungstages der vorherigen Rechnung in die neue Rechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Gutschriftfreitext :"
    Lbl09.Caption = "Gutschriftsnummer :"
    Lbl10.Caption = "Gutschriftsdatum :"
    ChRnm.Caption = "Gutschriftsnummer jetzt erzeugen"
    ChDiK.Enabled = False
    ChDiR.Enabled = False
    CmVer.Enabled = False
Case "M":
    If GlRKo = True Then
        Me.Caption = "Rechnungsauftrag kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie einen neuen Rechnungsauftrag kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Rechnungsauftrag hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie einen neuen Rechnungsauftrag anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Dem neuen Rechnungsauftrag wird folgende Auftragsnummer und folgendes Auftragsdatum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des ersten Behandlungstages der vorherigen Rechnung in die neue Rechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Auftragsfreitext :"
    Lbl09.Caption = "Auftragsnummer :"
    Lbl10.Caption = "Auftragsdatum :"
    ChRnm.Caption = "Auftragsnummer jetzt erzeugen"
Case "G":
    If GlRKo = True Then
        Me.Caption = "Gewerberechnung kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Gewerberechnung kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Gewerberechnung hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Gewerberechnung anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Der neuen Gewerberechnung wird folgende Rechnungsnummer und folgendes Datum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des ersten Behandlungstages der vorherigen Rechnung in die neue Rechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Gewerberechnungsfreitext :"
    Lbl09.Caption = "Rechnungsnummer :"
    Lbl10.Caption = "Rechnungsdatum :"
    ChRnm.Caption = "Rechnungsnummer jetzt erzeugen"
    ChDiK.Enabled = False
    ChDiR.Enabled = False
    CmVer.Enabled = False
Case "I":
    If GlRKo = True Then
        Me.Caption = "Importrechnung kopieren"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Importrechnung kopieren? Wählen Sie die gewünschten Optionen."
    Else
        Me.Caption = "Importrechnung hinzufügen"
        Lbl02.Caption = "Für welchen Mandanten und Mitarbeiter möchten Sie eine neue Importrechnung anlegen? Wählen Sie die gewünschten Optionen."
    End If
    Lbl08.Caption = "Der neuen Importrechnung wird folgende Rechnungsnummer und folgendes Datum zugewiesen. Falls erforderlich, können Sie diese Angaben jetzt ändern."
    Lbl29.Caption = "Wenn Sie die Leistungen des letzten Behandlungstages in die neue Importrechnung kopieren möchten, markieren Sie bitte den oder die neuen Behandlungstermine."
    Lbl05.Caption = "Importrechnungsfreitext :"
    Lbl09.Caption = "Importrechnungsnummer :"
    Lbl10.Caption = "Importrechnungsdatum :"
    ChRnm.Caption = "Importrechnungsnummer jetzt erzeugen"
End Select

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FCapt " & Err.Number
Resume Next

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date
Dim ReStr As String

Set TxRen = Me.txtReNum
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set DaPi1 = Me.dtpDatu1

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
End If

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
        TxDa1.Text = FDaPr(NeuDa, ReStr)
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
        TxDa2.Text = Format$(NeuDa, "dd.mm.yyyy")
    End If
End Select

With DaPi1
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

DoEvents
FReKo

If NeuDa > Date Then
    SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
End If

Set DaPi1 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FDat1()
On Error GoTo OrErr

Dim NeuDa As Date
Dim ReStr As String

Set TxRen = Me.txtReNum
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set DaPi1 = Me.dtpDatu1

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
End If

If DaPi1.Selection.BlocksCount > 0 Then
    NeuDa = DaPi1.Selection.Blocks(0).DateBegin
    Select Case KalWa
    Case 1: TxDa1.Text = FDaPr(NeuDa, ReStr)
    Case 2: TxDa2.Text = FDaPr(NeuDa, ReStr)
    End Select
    TxDa1.SetFocus
End If

DoEvents
FReKo

Set DaPi1 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDat1 " & Err.Number
Resume Next

End Sub
Private Sub FDat2()
On Error GoTo OrErr

Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim AnzBl As Long
Dim AktBl As Long
Dim AktTa As Long
Dim BloTa As Long

Set DaPi2 = Me.dtpDatu2

AktTa = 0
AnzBl = DaPi2.Selection.BlocksCount

If AnzBl = 1 Then
    DaBeg = DaPi2.Selection(0).DateBegin
    DaEnd = DaPi2.Selection(0).DateEnd
    If DaEnd > DaBeg Then
        Do
        DaAkt = DaBeg + AktTa
        AktTa = AktTa + 1
        ReDim Preserve GlTag(AktTa)
        GlTag(AktTa) = DaAkt
        Loop Until DaAkt >= DaEnd
    Else
        ReDim Preserve GlTag(1)
        GlTag(1) = DaBeg
    End If
ElseIf AnzBl > 1 Then
    For AktBl = 0 To AnzBl - 1
        DaBeg = DaPi2.Selection.Blocks(AktBl).DateBegin
        DaEnd = DaPi2.Selection.Blocks(AktBl).DateEnd
        If DaEnd > DaBeg Then
            BloTa = 0
            Do
            DaAkt = DaBeg + BloTa
            AktTa = AktTa + 1
            BloTa = BloTa + 1
            ReDim Preserve GlTag(AktTa)
            GlTag(AktTa) = DaAkt
            Loop Until DaAkt >= DaEnd
        Else
            AktTa = AktTa + 1
            ReDim Preserve GlTag(AktTa)
            GlTag(AktTa) = DaBeg
        End If
    Next AktBl
End If

If GlPop = True Then
    S_AbTa
    S_AbDo
End If

Set DaPi2 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDat2 " & Err.Number
Resume Next

End Sub
Private Function FDaPr(ByVal NeuDa As Date, ByVal ReStr As String) As Date
On Error GoTo OrErr

If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
    If Year(NeuDa) <> Year(Date) Then
        If ReStr <> "-" Then
            FDaPr = Date
            SPopu "Ungültiges Rechnungsdatum", "Das Rechnungsdatum muss sich innerhalb des aktuellen Geschäftsjahrs befinden", IC48_Information
        Else
            FDaPr = NeuDa
        End If
    Else
        FDaPr = NeuDa
    End If
Else
    FDaPr = NeuDa
End If

Exit Function

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaPr " & Err.Number
Resume Next

End Function
Private Sub FErst()
On Error GoTo InErr

Set OpBEr = Me.optBehEr
Set OpBLe = Me.optBehLe

If FoLad = False Then
    If OpBEr.Value = True Then
        IniSetVal "System", "ReNeEr", -1
    Else
        IniSetVal "System", "ReNeEr", 0
    End If
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FErst " & Err.Number
Resume Next

End Sub

Private Sub FExp()
On Error GoTo InErr

Set ChExp = Me.chkExpMo
Set DaPi2 = Me.dtpDatu2

If FoLad = False Then
    If ChExp.Value = 1 Then
        IniSetVal "System", "KopExp", -1
        AbExp = True
    Else
        IniSetVal "System", "KopExp", 0
        AbExp = False
    End If
    DaPi2.MultiSelectionMode = AbExp
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FExp " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo InErr

Dim LiLin As Boolean
Dim FeGeg As VB.ComboBox
Dim CmBuT As VB.ComboBox
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCon = Me.repCont5
Set ImMan = frmMain.imgManag
Set RpCls = RpCon.Columns

LiLin = False

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
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = False
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
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = GlAnZ
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = True
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
    .PreviewMode = False
    .ShowHeader = False
    .SelectionEnable = False
    .ScrollModeH = xtpReportScrollModeNone
    .ScrollModeV = xtpReportScrollModeBlock
End With

Set RpCls = RpCon.Columns
With RpCls
    Set RpCol = .Add(0, "Rechnung", 0, True)
    Set RpCol = .Add(1, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(2, "Abgeschlossen", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(3, "Patient", 0, True)
    If RpCon.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(4, "Storniert", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(5, "Type", 0, False)
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

If GlTFt.SIZE > 10 Then
    RpCls(0).Width = 135
    RpCls(1).Width = 105
Else
    RpCls(0).Width = 105
    RpCls(1).Width = 75
End If
RpCls(2).Width = 0
RpCls(3).Width = 220
RpCls(4).Width = 0
RpCls(5).Width = 0

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub

Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim ReStr As String

Set TxRen = Me.txtReNum
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set DaPi1 = Me.dtpDatu1
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
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
End Select

With DaPi1
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    Select Case KalWa
    Case 1:
        .Top = Rahm4.Top + TxDa1.Top + TxDa1.Height
        .Left = Rahm4.Left + TxDa1.Left
        If .ShowModal(1, 1) Then
            If .Selection.BlocksCount > 0 Then
                If IsDate(TxDa1.Text) Then
                    NeuDa = .Selection.Blocks(0).DateBegin
                    TxDa1.Text = FDaPr(NeuDa, ReStr)
                End If
            End If
        End If
    Case 2:
        .Top = Rahm5.Top + TxDa2.Top + TxDa2.Height
        .Left = Rahm5.Left + TxDa2.Left
        If .ShowModal(1, 1) Then
            If .Selection.BlocksCount > 0 Then
                If IsDate(TxDa2.Text) Then
                    NeuDa = .Selection.Blocks(0).DateBegin
                    TxDa2.Text = FDaPr(NeuDa, ReStr)
                End If
            End If
        End If
    End Select
End With

Set DaPi1 = Nothing

FReKo

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo SuErr

Dim NeuDa As Date
Dim DayFi As Date
Dim DayLa As Date
Dim AnzRe As Long
Dim ArNum As Long
Dim KatNr As Long
Dim RetWe As Long
Dim StMan As Long
Dim StMit As Long
Dim GuBet As Single
Dim MaIdx As Integer
Dim ReKop As Integer
Dim ZaZil As Integer
Dim AktZa As Integer
Dim Kopie As Integer
Dim LauZa As Integer
Dim PaWar As Integer
Dim LiIdx As Integer
Dim TeWer As Variant
Dim TeGut As Variant
Dim TeTyp As Variant
Dim ArzNr As Variant
Dim BeVor As Boolean
Dim VrsAr As Integer
Dim BehEr As Boolean
Dim Mld1, Tit1 As String
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set TxDum = Me.txtDummy
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxRzn = Me.txtRzNum
Set CmKom = Me.txtKomFe
Set CmVer = Me.cmbVersi
Set CmTyp = Me.cmbReTyp
Set CmZil = Me.cmbZaZie
Set CmStu = Me.cmbReStu
Set CmWar = Me.cmbReWar
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set CmArz = Me.cmbArzNr
Set CmBTy = Me.cmbBeTyp
Set CmVrs = Me.cmbVersa
Set CmGru = Me.cmbBeGru
Set ChRnm = Me.chkReNum
Set ChExp = Me.chkExpMo
Set ChDiR = Me.chkDiaRe
Set ChDiK = Me.chkDiaKr
Set ChPhy = Me.chkPhyAb
Set ChGru = Me.chkGrThe
Set ChThe = Me.chkTheBe
Set TxRen = Me.txtReNum
Set TxKop = Me.txtReKop
Set OpBEr = Me.optBehEr
Set OpBLe = Me.optBehLe
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set PuBu1 = Me.btnDatu1
Set PuBu2 = Me.btnDatu2
Set ImMan = FM.imgManag
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set DaPi1 = Me.dtpDatu1
Set DaPi2 = Me.dtpDatu2

StMan = GlMan(GlSMa, 2)
StMit = GlMiA(GlSmI, 2)
AbExp = CBool(IniGetVal("System", "KopExp"))
BehEr = CBool(IniGetVal("System", "ReNeEr"))

If GlRKo = True Then 'Rechnung Kopieren
    Tit1 = "Rechnung Kopieren"
    Mld1 = "Für diesen Patienten existiert noch eine nicht abgeschlossene Rechnung." & vbCrLf & "Diese ggf. erst abschließen."
Else
    Tit1 = "Neue Rechnung"
    Mld1 = "Für diesen Patienten existiert noch eine nicht abgeschlossene Rechnung." & vbCrLf & "Diese ggf. erst abschließen."
End If

With DaPi1
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
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Behandlungstag"
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

With DaPi2
    .AllowNoncontinuousSelection = True
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    .BorderStyle = 0
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    .MaxSelectionCount = GlMxK 'Maximal slektierbare Kalendertage
    .MultiSelectionMode = AbExp
    .RightToLeft = False
    .ShowNoneButton = True
    .ShowNonMonthDays = True
    .ShowTodayButton = True
    .ShowWeekNumbers = False
    .TextNoneButton = "Markiere Keine"
    .TextTodayButton = "Markiere Heute"
    .ToolTipText = "Markieren Sie bitte hier die Behandlungstage des Patienten"
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
    If GlStS > 1 Then 'Standardsteuersatz
        If GlReT = vbNullString Then 'Standardbelegtyp
            ReTyp = "G"
        End If
    End If
    Select Case ReTyp
    Case "R": LiIdx = 0
    Case "V": LiIdx = 1
    Case "L": LiIdx = 2
    Case "A": LiIdx = 3
    Case "U": LiIdx = 4
    Case "M": LiIdx = 5
    Case "G": LiIdx = 6
    Case "I": LiIdx = 7
    End Select
    .ListIndex = LiIdx
End With

For AktZa = 1 To UBound(GlZah)
    CmZil.AddItem GlZah(AktZa, 1)
    CmZil.ItemData(AktZa - 1) = GlZah(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlStu)
    CmStu.AddItem GlStu(AktZa, 2)
    CmStu.ItemData(AktZa - 1) = GlStu(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlWar)
    CmWar.AddItem GlWar(AktZa, 1)
    CmWar.ItemData(AktZa - 1) = GlWar(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlReK)
    CmKom.AddItem GlReK(AktZa, 1)
    CmKom.ItemData(AktZa - 1) = GlReK(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlGKa)
    CmVer.AddItem GlGKa(AktZa, 1)
    CmVer.ItemData(AktZa - 1) = GlGKa(AktZa, 0)
Next AktZa

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
    .ListIndex = 0
End With

With CmVrs
    .AddItem "Postversand"
    .ItemData(0) = 0
    .AddItem "Emailversand"
    .ItemData(1) = 1
    .AddItem "Downloadlink"
    .ItemData(2) = 2
    .ListIndex = GlRVs - 1 'Standard-Rechnungsversandweg
End With

For AktZa = 0 To UBound(GlThG)
    With CmGru
        .AddItem GlThG(AktZa, 0)
        .ItemData(AktZa) = GlThG(AktZa, 1)
    End With
Next AktZa
CmGru.ListIndex = 0

If GlReN = True Then
    ChRnm.Value = xtpChecked 'Rechnungsnummern sofort erzeugen
End If
If GlDiR = True Then
    ChDiR.Value = xtpChecked
End If
If GlDiK = True Then
    ChDiK.Value = xtpChecked
End If
If GlPhs = True Then
    ChPhy.Value = xtpChecked
End If

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

Select Case GlRFm 'Rechnungsnummernformat
Case 2: TxRen.SetMask "00-00-000000", "__-__-______"
Case 3: TxRen.SetMask "0000-000000", "____-______"
Case 4: TxRen.SetMask "00-000000", "__-______"
Case 5: TxRen.SetMask "0000-0000", "____-____"
End Select

With TxKop
    .Pattern = "\d*"
    .SetMask "00", "__"
End With

With TxRzn
    .Pattern = "\d*"
    .Text = "000000"
    .SetMask "000000", "______"
End With

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
    Set RpCol = RpCls.Find(Rec_ID0)
    mPaNr = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_ID1)
    mReNr = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_Datum)
    NeuDa = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_RechNr)
    mReSt = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_Betrag)
    mReBe = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_Bezahlt)
    mReBz = RpRow.Record(RpCol.ItemIndex).Value
Else
    mPaNr = GlAdr
    NeuDa = Date
    Select Case GlBut
    Case RibTab_Abrechnung:
                If S_ReOf(mPaNr, 0, ReTyp) > 0 Then
                    WindowMess Mld1, Dial2, Tit1, FM.hwnd
                    FReNu 'Reindizierung
                End If
    Case RibTab_Rechnungen:
    End Select
End If

For AktZa = 1 To UBound(GlMaA)
    LauZa = LauZa + 1
    With CmMan
        .AddItem GlMaA(AktZa, 1)
        .ItemData(LauZa - 1) = GlMaA(AktZa, 2)
    End With
Next AktZa

For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
    CmMit.AddItem GlMiA(AktZa, 1)
    CmMit.ItemData(AktZa - 1) = GlMiA(AktZa, 2)
Next AktZa
CmMit.ListIndex = GlSmI - 1

If GlArV = True Then
   For AktZa = 1 To UBound(GlArz) 'Verordner
        CmArz.AddItem GlArz(AktZa, 8)
        CmArz.ItemData(AktZa - 1) = GlArz(AktZa, 0)
    Next AktZa
End If

If mPaNr > 0 Then
    S_AdDe mPaNr 'Adressendetails
    With GlADt
        ZaZil = .AdZil
        KatNr = .AdKat
        TeWer = .AdMan
        TeTyp = .AdTyp
        ReKop = .AdKop
        TeGut = .AdGut
        mPaKu = .AdKur
        PaWar = .AdWar
        ArzNr = .AdBGn
        VrsAr = .AdVer
    End With
    DoEvents

    If ReKop > 0 Then
        TxKop.Text = Format$(ReKop, "00")
    Else
        TxKop.Text = "01"
    End If
    If ZaZil > 0 Then
        CmZil.ListIndex = SCmb(CmZil, ZaZil)
    Else
        CmZil.ListIndex = 0
    End If
    If PaWar > 0 Then
        For AktZa = 1 To UBound(GlWar)
            If CInt(GlWar(AktZa, 0)) = PaWar Then
                CmWar.ListIndex = AktZa - 1
                Exit For
            End If
        Next AktZa
    Else
        CmWar.ListIndex = 0
    End If
    If KatNr > 0 Then
        CmVer.ListIndex = SCmb(CmVer, KatNr)
    Else
        CmVer.ListIndex = SCmb(CmVer, GlStK)
    End If
    CmVrs.ListIndex = VrsAr
    If TeTyp <> vbNullString Then
        If IsNumeric(TeTyp) = True Then
            CmBTy.ListIndex = TeTyp - 1
        Else
            CmBTy.ListIndex = GlStP
        End If
    Else
        CmBTy.ListIndex = GlStP 'Standardempfängertyp
    End If

    If GlRst = False Then 'Mandantenbezogene Datenbegrenzung
        Select Case GlMaR 'Mandant neue(s) Rechnung/Rezept
        Case "J1": 'Standardmandant aus Optionsdialog
            mMaNr = GlMan(GlSMa, 2)
        Case "J2": 'Mandant aus Adresseneingabemaske
            If TeWer <> vbNullString Then
                mMaNr = CLng(TeWer)
                For AktZa = 1 To UBound(GlMan)
                    If mMaNr = GlMan(AktZa, 2) Then
                        BeVor = True
                        If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
                            If FoLad = False Then
                                ReTyp = GlMaA(AktZa, 30) 'Standardbelegtyp
                                FStSa = GlMaA(AktZa, 24) 'Standardsteuersatz
                                Select Case ReTyp
                                Case "R": LiIdx = 0
                                Case "V": LiIdx = 1
                                Case "L": LiIdx = 2
                                Case "A": LiIdx = 3
                                Case "U": LiIdx = 4
                                Case "M": LiIdx = 5
                                Case "G": LiIdx = 6
                                Case "I": LiIdx = 7
                                End Select
                                CmTyp.ListIndex = LiIdx
                                CmStu.ListIndex = FStSa - 1
                                FTyPr
                                FCapt
                                FReRc
                            End If
                        End If
                        Exit For
                    End If
                Next AktZa
                If BeVor = True Then
                    If CBool(GlMan(AktZa, 5)) = False Then 'Passiv / Aktiv
                        mMaNr = GlThe(AktZa, 0)
                    Else
                        mMaNr = GlMan(GlSMa, 2)
                    End If
                Else
                    mMaNr = GlMan(GlSMa, 2)
                End If
            Else
                mMaNr = GlMan(GlSMa, 2)
            End If
        Case "J3": 'Mandant aus Mitarbeitereingabemaske
            mMaNr = GlMiA(GlSmI, 7)
        End Select
    Else
        mMaNr = GlMiA(GlSmI, 7)
    End If

    If GlArV = True Then 'Verordner vorhanden
        If ArzNr <> vbNullString Then
            If IsNumeric(ArzNr) = True Then
                ArNum = SCmb(CmArz, CLng(ArzNr))
                If ArNum > 1 Then
                    CmArz.ListIndex = ArNum
                End If
            End If
        End If
    End If
Else
    SPopu "Keine Patienten", "Sie müssen erst einen Patienten anlegen, bevor Sie eine Rechnung anlegen können", IC48_Forbidden
End If

MaIdx = SCmb(CmMan, mMaNr)
If MaIdx < 0 Then
    MaIdx = 0
End If
With CmMan
    .ListIndex = MaIdx
    .Enabled = True
End With

If GlSpl = True Then 'Steuerspalte
    With CmStu
        .Enabled = False
        .ListIndex = 0
    End With
Else
    CmStu.ListIndex = GlStS - 1 'Standardsteuersatz
End If

With DaPi2
    .EnsureVisible NeuDa
    DayFi = .FirstDayOfWeek
    DayLa = .LastVisibleDay
End With

S_AbTe DayFi, DayLa

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Rahm6.BackColor = GlBak
ChRnm.BackColor = GlBak
ChExp.BackColor = GlBak
ChDiR.BackColor = GlBak
ChDiK.BackColor = GlBak
ChPhy.BackColor = GlBak
ChGru.BackColor = GlBak
ChThe.BackColor = GlBak
OpBEr.BackColor = GlBak
OpBLe.BackColor = GlBak

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

If TeGut <> vbNullString Then
    GuBet = CSng(TeGut)
    If GuBet > 0 Then
        mPaGu = GuBet
    End If
End If

GlNeR = GlNeX

If AbExp = True Then
    ChExp.Value = 1
End If

If BehEr = True Then
    OpBEr.Value = True
Else
    OpBLe.Value = True
End If

Set ImMan = Nothing
Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FReKo()
On Error GoTo SuErr
'Rechnungsnummernkontrolle

Dim NeuDa As Date
Dim MaNum As Long
Dim MiNum As Long
Dim ReStr As String
Dim ReTyp As String

Set CmTyp = Me.cmbReTyp
Set TxDa1 = Me.txtDatu1
Set TxRen = Me.txtReNum
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar

If TxRen.Text <> vbNullString Then
    ReStr = TxRen.Text
Else
    ReStr = vbNullString
End If

MaNum = CmMan.ItemData(CmMan.ListIndex)
MiNum = CmMit.ItemData(CmMit.ListIndex)

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        NeuDa = CDate(TxDa1.Text)
    Else
        NeuDa = Date
    End If
Else
    NeuDa = Date
End If

Select Case CmTyp.ListIndex
Case 0: ReTyp = "R"
Case 1: ReTyp = "V"
Case 2: ReTyp = "L"
Case 3: ReTyp = "A"
Case 4: ReTyp = "U"
Case 5: ReTyp = "M"
Case 6: ReTyp = "G"
Case 7: ReTyp = "I"
End Select

If GlReN = True Then 'Rechnungsnummern sofort erzeugen
    If ReStr <> "-" Then
        TxRen.Text = S_ReVo(NeuDa, ReTyp, MaNum, MiNum, True)
        FReRc
    End If
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FReKo " & Err.Number
Resume Next

End Sub
Private Sub FReNu()
On Error GoTo SuErr

Dim Mld1, Mld2 As String
Dim Mld3, Mld4 As String

Set FM = frmReNeu

Mld1 = "Rechnungsnummer wurde geändert"
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
Private Sub FReRc()
On Error Resume Next
'Füllt die Tabelle mit den letzten drei Rechnungen

Dim AktZa As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpRec As XtremeReportControl.ReportRecord
Dim RpRcs As XtremeReportControl.ReportRecords
Dim RpItm As XtremeReportControl.ReportRecordItem

Set RpCon = Me.repCont5
Set TxRen = Me.txtReNum
Set ImMan = frmMain.imgManag
Set RpRcs = RpCon.Records

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    .Populate
End With

For AktZa = 0 To UBound(GlRec) - 1 'Gefundene Rechnungen
    Set RpRec = RpRcs.Add()
    Set RpItm = RpRec.AddItem(GlRec(AktZa, 1))
    RpItm.Focusable = False
    If GlRec(AktZa, 5) = 0 Then
        RpItm.Icon = IC16_Mail_Open
    Else
        RpItm.Icon = IC16_Mail_Close
    End If
    Set RpItm = RpRec.AddItem(GlRec(AktZa, 3))
    RpItm.Focusable = False
    Set RpItm = RpRec.AddItem(GlRec(AktZa, 5))
    RpItm.Focusable = False
    Set RpItm = RpRec.AddItem(GlRec(AktZa, 2))
    RpItm.Focusable = False
    Set RpItm = RpRec.AddItem(GlRec(AktZa, 4))
    RpItm.Focusable = False
    Set RpItm = RpRec.AddItem(GlRec(AktZa, 6))
    RpItm.Focusable = False
Next AktZa

RpCon.Populate

If UBound(GlRec) = 0 Then
    TxRen.Enabled = True
Else
    TxRen.Enabled = False
End If

End Sub
Private Sub FRest()
On Error Resume Next

Set TxKur = Me.txtKurz
Set TxPLZ = Me.txtPost
Set TxBem = Me.txtBemer
Set TxRen = Me.txtReNum

TxKur.Text = vbNullString
TxPLZ.Text = vbNullString
TxBem.Text = vbNullString
TxRen.Text = vbNullString

End Sub

Private Sub FSuda()
On Error GoTo SuErr

Dim Mld1, Tit1 As String

Set TxKur = Me.txtKurz
Set TxPLZ = Me.txtPost
Set TxBem = Me.txtBemer
Set FLis1 = Me.lstList1
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5

If TxKur.Text <> vbNullString Then
    S_AdFin TxKur.Text, 1, 1
ElseIf TxPLZ.Text <> vbNullString Then
    S_AdFin TxPLZ.Text, 2, 1
ElseIf TxBem.Text <> vbNullString Then
    S_AdFin TxBem.Text, 3, 1
End If

If FLis1.ListCount > 0 Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
    Rahm5.Visible = False
    FLis1.SetFocus
    FLis1.Selected(0) = True
Else
    Mld1 = "Das von Ihnen eingegebene Suchkriterium brachte leider keine Suchergebnisse"
    Tit1 = "Adressuche"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuda " & Err.Number
Resume Next

End Sub
Private Sub FTyPr()
On Error GoTo LaErr
'Kontrolliert den Belegtyp

Dim ReDat As Date
Dim AnzRe As Long
Dim ReNum As Long
Dim MaNum As Long
Dim MiNum As Long
Dim ReStr As String
Dim ReTyp As String
Dim TyStr As String
Dim BeBez As Single
Dim ReAbg As Boolean
Dim TypNr As Integer
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
Set TxDa1 = Me.txtDatu1
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar

TypNr = CmTyp.ListIndex
MaNum = CmMan.ItemData(CmMan.ListIndex)
MiNum = CmMit.ItemData(CmMit.ListIndex)

Select Case CmTyp.ListIndex
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
        ReDat = CDate(TxDa1.Text)
    Else
        ReDat = Date
    End If
Else
    ReDat = Date
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
    Set RpCol = RpCls.Find(Rec_Selekt)
    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
        ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
    Else
        ReAbg = True
    End If
    Set RpCol = RpCls.Find(Rec_Type)
    TyStr = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Rec_Bezahlt)
    BeBez = RpRow.Record(RpCol.ItemIndex).Value
    
    If TypNr = 4 Then 'Gutschrift
        If ReAbg = False Then
            GlRtT = vbNullString 'WICHTIG!
            CmTyp.ListIndex = 0
            Tit1 = "Rechnung nicht abgeschlossen"
            Mld1 = "Die Rechnung " & ReStr & " wurde noch nicht abgeschlossen und kann daher nicht gutgeschrieben werden, da diese noch korrigiert werden kann."
            WindowMess Mld1, Dial2, Tit1, FM.hwnd
        Else
            If BeBez = 0 Then
                GlRtT = vbNullString 'WICHTIG!
                CmTyp.ListIndex = 0
                Tit1 = "Rechnung nicht bezahlt"
                Mld1 = "Für die Rechnung " & ReStr & " wurde noch kei Erlös gebucht und brauch daher nicht gutgeschrieben werden."
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
            Else
                Select Case UCase(TyStr)
                Case "R":
                Case "G":
                Case "L":
                Case "I":
                Case Else:
                    GlRtT = vbNullString 'WICHTIG!
                    Tit1 = "Falscher Belegtyp"
                    Mld1 = "Vom Beleg " & ReStr & " kann keine Gutschrift erstellt werden, da es keine Rechnung ist."
                    WindowMess Mld1, Dial2, Tit1, FM.hwnd
                End Select
            End If
        End If
    End If
End If

TxRen.Text = S_ReVo(ReDat, ReTyp, MaNum, MiNum, True)

If TyStr = "U" Or TyStr = "V" Or TyStr = "M" Then
    TxRen.Enabled = False
Else
    If GlReN = False Then 'Rechnungsnummern sofort erzeugen
        TxRen.Enabled = False
    Else
        If UBound(GlRec) > 0 Then
            TxRen.Enabled = False
        Else
            TxRen.Enabled = True
        End If
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
Private Sub FVoSe()
On Error GoTo LaErr

Dim MaIdx As Integer
Dim LiIdx As Integer

Set CmMan = Me.cmbManda
Set CmStu = Me.cmbReStu
Set CmTyp = Me.cmbReTyp

If FoLad = False Then
    If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
        MaIdx = CmMan.ListIndex + 1
        ReTyp = GlMaA(MaIdx, 30) 'Standardbelegtyp
        FStSa = GlMaA(MaIdx, 24) 'Standardsteuersatz
        Select Case ReTyp
        Case "R": LiIdx = 0
        Case "V": LiIdx = 1
        Case "L": LiIdx = 2
        Case "A": LiIdx = 3
        Case "U": LiIdx = 4
        Case "M": LiIdx = 5
        Case "G": LiIdx = 6
        Case "I": LiIdx = 7
        End Select
        CmTyp.ListIndex = LiIdx
        CmStu.ListIndex = FStSa - 1
        FCapt
    End If
End If

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FVoSe " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo SuErr

Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim DayFi As Date
Dim DayLa As Date
Dim ReDat As Date
Dim RzDat As Date
Dim RowNr As Long
Dim TmpNr As Long
Dim KatNr As Long
Dim MaNum As Long
Dim MiNum As Long
Dim AktTa As Long
Dim BloTa As Long
Dim AnzBl As Long
Dim AktBl As Long
Dim NeuRe As Long
Dim ArNum As Long
Dim RzNum As Long
Dim ReBet As Single
Dim ReStu As Single
Dim AnzBe As Single
Dim ReRab As Single
Dim GuBet As Single
Dim GuiID As String
Dim ReKom As String
Dim EiTex As String
Dim ReStr As String
Dim TyStr As String
Dim RzTex As String
Dim NeReB As Double
Dim BeBet As Double
Dim BeGru As Integer
Dim ReKop As Integer
Dim AktZa As Integer
Dim StuZa As Integer
Dim ZaZil As Integer
Dim ZaInt As Integer
Dim ZiTag As Integer
Dim Lange As Integer
Dim PaWar As Integer
Dim ReTyp As Integer
Dim BeTyp As Integer
Dim VerAr As Integer
Dim ZiMah As Boolean
Dim GrThe As Boolean
Dim ThBen As Boolean
Dim ReAbg As Boolean
Dim RetWe As Boolean
Dim BehEr As Boolean
Dim ArzNr As Variant
Dim Mld1, Mld2, Tit1 As String
Dim Frage As Integer
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set TxDum = Me.txtDummy
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxKur = Me.txtKurz
Set TxPLZ = Me.txtPost
Set TxWar = Me.txtPatWa
Set TxBem = Me.txtBemer
Set TxRen = Me.txtReNum
Set TxKop = Me.txtReKop
Set TxRzn = Me.txtRzNum
Set TxRzt = Me.txtRzTex
Set CmKom = Me.txtKomFe
Set CmTyp = Me.cmbReTyp
Set CmZil = Me.cmbZaZie
Set CmVer = Me.cmbVersi
Set CmStu = Me.cmbReStu
Set CmWar = Me.cmbReWar
Set CmMan = Me.cmbManda
Set CmMit = Me.cmbMitar
Set CmBTy = Me.cmbBeTyp
Set CmGru = Me.cmbBeGru
Set CmArz = Me.cmbArzNr
Set CmVrs = Me.cmbVersa
Set ChRnm = Me.chkReNum
Set ChDiR = Me.chkDiaRe
Set ChDiK = Me.chkDiaKr
Set ChPhy = Me.chkPhyAb
Set ChGru = Me.chkGrThe
Set ChThe = Me.chkTheBe
Set OpBEr = Me.optBehEr
Set OpBLe = Me.optBehLe
Set FLis1 = Me.lstList1
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set DaPi1 = Me.dtpDatu1
Set DaPi2 = Me.dtpDatu2
Set PuBu3 = Me.btnWeite
Set PuBu4 = Me.btnZuruk
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4

If GlTza = True Then 'Testzeit abgelaufen
    SPopu "Lizenzierung erforderlich!", "Es ist keine bzw. keine gültige Seriennummer vorhanden oder die Testzeit ist abgelaufen.", IC48_Forbidden
    Exit Sub
End If

Tit1 = "Rechnung erst abschließen!"
Mld1 = "Für diesen Patienten existiert noch eine nicht abgeschlossene Rechnung. Diese sollte erst abgeschlossen werden."

MaNum = CmMan.ItemData(CmMan.ListIndex)
MiNum = CmMit.ItemData(CmMit.ListIndex)
ArNum = CmArz.ItemData(CmArz.ListIndex)

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        ReDat = CDate(TxDa1.Text)
    Else
        ReDat = Date
    End If
Else
    ReDat = Date
End If

If TxDa2.Text <> vbNullString Then
    If IsDate(TxDa2.Text) = True Then
        RzDat = CDate(TxDa2.Text)
    Else
        RzDat = Date
    End If
Else
    RzDat = Date
End If

Select Case CmTyp.ListIndex
Case 0: TyStr = "R"
Case 1: TyStr = "V"
Case 2: TyStr = "L"
Case 3: TyStr = "A"
Case 4: TyStr = "U"
Case 5: TyStr = "M"
Case 6: TyStr = "G"
Case 7: TyStr = "I"
End Select

Screen.MousePointer = vbHourglass
PuBu3.Enabled = False

If Rahm1.Visible = True Then
    If GlPhs = True Then
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = True
    Else
        Select Case GlBut
        Case RibTab_Startseite:
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
            TxKur.SetFocus
        Case RibTab_Abrechnung:
            Rahm1.Visible = False
            Rahm2.Visible = False
            Rahm3.Visible = False
            Rahm4.Visible = True
            Rahm5.Visible = False
            Rahm6.Visible = False
            DoEvents
            If ChRnm.Value = xtpChecked Then 'Rechnungsnummern sofort erzeugen
                TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, True)
                FReRc
            Else
                If TyStr = "U" Or TyStr = "V" Or TyStr = "M" Then
                    TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, True)
                Else
                    TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, False)
                End If
                FReRc
            End If
            If GlRMa = True Then 'getrennter Mandentenrechnungsnummernkreis
                If S_ReOf(mPaNr, MaNum, TyStr) > 0 Then
                    SPopu Tit1, Mld1, IC48_Information
                End If
            Else
                If S_ReOf(mPaNr, 0, TyStr) > 0 Then
                    SPopu Tit1, Mld1, IC48_Information
                End If
            End If
            If TyStr = "U" Or TyStr = "V" Or TyStr = "M" Then
                TxRen.Enabled = False
            Else
                If GlReN = False Then 'Rechnungsnummern sofort erzeugen
                    TxRen.Enabled = False
                Else
                    If UBound(GlRec) > 0 Then
                        TxRen.Enabled = False
                    Else
                        TxRen.Enabled = True
                    End If
                End If
            End If
        Case RibTab_Rechnungen:
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
            TxKur.SetFocus
        End Select
    End If
    PuBu4.Enabled = True
    
ElseIf Rahm2.Visible = True Then
    
    FSuda
    
ElseIf Rahm3.Visible = True Then
    
    mPaNr = FLis1.ItemData(FLis1.ListIndex)
    GlAdr = mPaNr
    GlTDa = vbNullString 'Wichtig für Textverarbeitung
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = True
    Rahm5.Visible = False
    Rahm6.Visible = False
    DoEvents
    If ChRnm.Value = xtpChecked Then 'Rechnungsnummern sofort erzeugen
        TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, True)
        FReRc
    Else
        If TyStr = "U" Or TyStr = "V" Or TyStr = "M" Then
            TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, True)
        Else
            TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, False)
        End If
        FReRc
    End If
    If GlRMa = True Then 'getrennter Mandentenrechnungsnummernkreis
        If S_ReOf(mPaNr, MaNum, TyStr) > 0 Then
            SPopu Tit1, Mld1, IC48_Information
        End If
    Else
        If S_ReOf(mPaNr, 0, TyStr) > 0 Then
            SPopu Tit1, Mld1, IC48_Information
        End If
    End If
    ReKop = S_AdIdi(mPaNr, "Kopien")
    If ReKop < 1 Then
        TxKop.Text = "01"
    Else
        If ReKop > 0 Then
            TxKop.Text = Format$(ReKop, "00")
        Else
            TxKop.Text = "01"
        End If
    End If
    
    If TyStr = "U" Or TyStr = "V" Or TyStr = "M" Then
        TxRen.Enabled = False
    Else
        If GlReN = False Then 'Rechnungsnummern sofort erzeugen
            TxRen.Enabled = False
        Else
            If UBound(GlRec) > 0 Then
                TxRen.Enabled = False
            Else
                TxRen.Enabled = True
            End If
        End If
    End If
    
    S_AdDe mPaNr 'Adressendetails
    With GlADt
        ZaZil = .AdZil
        KatNr = .AdKat
        mPaKu = .AdKur
        PaWar = .AdWar
        ArzNr = .AdBGn
    End With

    If ZaZil > 0 Then
        CmZil.ListIndex = SCmb(CmZil, ZaZil)
    Else
        CmZil.ListIndex = 0
    End If
    If PaWar > 0 Then
        CmWar.ListIndex = PaWar - 1
    Else
        CmWar.ListIndex = 0
    End If
    If KatNr > 0 Then
        CmVer.ListIndex = SCmb(CmVer, KatNr)
    Else
        CmVer.ListIndex = SCmb(CmVer, GlStK)
    End If
    If GlArV = True Then
        If ArzNr <> vbNullString Then
            CmArz.ListIndex = SCmb(CmArz, CLng(ArzNr))
        End If
    End If
    
ElseIf Rahm4.Visible = True Then

    Select Case GlRFm 'Rechnungsnummernformat
    Case 2: Lange = 12
    Case 3: Lange = 11
    Case 4: Lange = 9
    Case 5: Lange = 9
    End Select

    Tit1 = "Falsches Rechnungsnummernformat"
    Mld1 = "Die von Ihnen eingestellte Rechnungsnummer hat das falsche Format"
    Mld2 = "Die von Ihnen eingestellte Rechnungsnummer existiert bereits"
    
    If TxRen.Text <> vbNullString Then
        ReStr = TxRen.Text
    Else
        ReStr = vbNullString
    End If
    
    If TxRzn.Text <> vbNullString Then
        RzNum = Val(TxRzn.Text)
    Else
        RzNum = 0
    End If
    
    If TxRzt.Text <> vbNullString Then
        RzTex = TxRzt.Text
    Else
        RzTex = vbNullString
    End If

    If GlReN = True Then 'Rechnungsnummern sofort erzeugen
        If InStrRev(ReStr, "_", -1, 1) > 0 Then
            Screen.MousePointer = vbNormal
            SPopu Tit1, Mld1, IC48_Warning
            Exit Sub
        ElseIf Len(ReStr) <> Lange Then
            Screen.MousePointer = vbNormal
            SPopu Tit1, Mld1, IC48_Warning
            Exit Sub
        Else
            If ReStr <> vbNullString Then
                If S_ReVr(ReStr, TyStr, MaNum, ReDat) = True Then
                    Screen.MousePointer = vbNormal
                    SPopu Tit1, Mld2, IC48_Warning
                    Exit Sub
                End If
            Else
                Screen.MousePointer = vbNormal
                SPopu Tit1, Mld1, IC48_Warning
                Exit Sub
            End If
        End If
    End If
    
    If CmKom.Text <> vbNullString Then
        ReKom = CmKom.Text
    End If
    
    If CmZil.Text <> vbNullString Then
        ZaZil = CmZil.ItemData(CmZil.ListIndex)
    Else
        ZaZil = CmZil.ItemData(1)
    End If
    If CmVer.Text <> vbNullString Then
        KatNr = CmVer.ItemData(CmVer.ListIndex)
    Else
        KatNr = CmVer.ItemData(1)
    End If
    If CmWar.Text <> vbNullString Then
        PaWar = CmWar.ItemData(CmWar.ListIndex)
    Else
        PaWar = CmWar.ItemData(1)
    End If

    For AktZa = 1 To UBound(GlZah)
        If GlZah(AktZa, 0) = ZaZil Then
            Screen.MousePointer = vbNormal
            ZiTag = GlZah(AktZa, 2)
            ZiMah = GlZah(AktZa, 3)
            ZaInt = GlZah(AktZa, 4)
            Exit For
        End If
    Next AktZa
    
    If CmStu.Text <> vbNullString Then
        StuZa = CmStu.ItemData(CmStu.ListIndex)
    Else
        StuZa = CmStu.ItemData(1)
    End If
    For AktZa = 1 To UBound(GlStu)
        If GlStu(AktZa, 0) = StuZa Then
            Screen.MousePointer = vbNormal
            ReStu = CSng(GlStu(AktZa, 1))
            If ReStu > 0 And ReStu < 1 Then
                ReStu = ReStu * 100
            End If
            Exit For
        End If
    Next AktZa
    If TxKop.Text > 0 Then
        ReKop = TxKop.Text
    Else
        ReKop = 1
    End If
    
    ReTyp = CmTyp.ListIndex
    BeTyp = CmBTy.ListIndex
    BeGru = CmGru.ListIndex
    VerAr = CmVrs.ListIndex
    ReStr = TxRen.Text

    If ChThe.Value = xtpChecked Then ThBen = True
    If ChGru.Value = xtpChecked Then GrThe = True

    With GlNeR
        .PatNr = mPaNr
        .ReDat = ReDat
        .ReStr = ReStr
        .ReKop = ReKop
        .ReStu = ReStu
        .ReZah = ZaZil
        .KatNr = KatNr
        .ReZie = ZiTag
        .ReInt = ZaInt
        .ReMah = ZiMah
        .ReKom = ReKom
        .VerAr = VerAr
        .ReTyp = ReTyp
        .BeTyp = BeTyp
        .MaNum = MaNum
        .MiNum = MiNum
        .ArNum = ArNum
        .PaStr = mPaKu
        .PaWar = PaWar
        .RzDat = RzDat
        .RzTex = RzTex
        .BeGru = BeGru
        .GrThe = GrThe
        .ThBen = ThBen
        If RzNum > 0 Then
            .RzNum = RzNum
        End If
        If GlRKo = True Then 'Rechnung Kopieren
            .ReBas = mReNr
        End If
        If ReTyp = 4 Then 'Gutschrift
            .GutNr = mReNr
            .GuStr = mReSt
        End If
    End With

    If GlRKo = True Then 'Rechnung Kopieren
        S_ReNe
        DoEvents
        S_ReKop
    Else
        S_ReNe
    End If

    Select Case GlBut
    Case RibTab_Startseite:
                SUpAb
                SUpRe
                DoEvents
                SAnza
                DoEvents
                SReZe GlAId
    Case RibTab_Abrechnung:
                SUpAb
                SUpRe
                DoEvents
                SAnza
    Case RibTab_Rechnungen:
                SUpRe
    End Select
    
    DoEvents
    KGeKa KatNr 'Gebührenkatalogaktualisierung
    DoEvents
    
    Select Case ReTyp
    Case 1: 'Kostenvoranschlag
        FReNu 'Reindizierung
        DoEvents
        Select Case GlBut
        Case RibTab_Abrechnung:
        Case RibTab_Rechnungen: SReZe mPaNr
        End Select
        DoEvents
        Unload Me
    Case 4: 'Gutschrift
        If mReNr > 0 Then
            Rahm1.Visible = False
            Rahm2.Visible = False
            Rahm3.Visible = False
            Rahm4.Visible = False
            Rahm5.Visible = False
            Rahm6.Visible = True
        Else
            FReNu 'Reindizierung
            DoEvents
            Select Case GlBut
            Case RibTab_Abrechnung:
            Case RibTab_Rechnungen: SReZe mPaNr
            End Select
            DoEvents
            Unload Me
        End If
    Case Else:
        If mReNr > 0 Then
            If GlRKo = True Then 'Rechnung Kopieren
                FReNu 'Reindizierung
                DoEvents
                Select Case GlBut
                Case RibTab_Abrechnung:
                Case RibTab_Rechnungen: SReZe mPaNr
                End Select
                DoEvents
                Unload Me
            Else
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = False
                Rahm4.Visible = False
                Rahm5.Visible = False
                Rahm6.Visible = True
            End If
        Else
            FReNu 'Reindizierung
            DoEvents
            Select Case GlBut
            Case RibTab_Abrechnung:
            Case RibTab_Rechnungen: SReZe mPaNr
            End Select
            DoEvents
            Unload Me
        End If
    End Select
    
ElseIf Rahm5.Visible = True Then

    Select Case GlBut 'identisch zu Rahm1
    Case RibTab_Startseite:
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        TxKur.SetFocus
    Case RibTab_Abrechnung:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = True
        Rahm5.Visible = False
        Rahm6.Visible = False
        DoEvents
        If ChRnm.Value = xtpChecked Then 'Rechnungsnummern sofort erzeugen
            TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, True)
            FReRc
        Else
            If TyStr = "U" Or TyStr = "V" Or TyStr = "M" Then
                TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, True)
            Else
                TxRen.Text = S_ReVo(ReDat, TyStr, MaNum, MiNum, False)
            End If
            FReRc
        End If
        If GlRMa = True Then 'getrennter Mandentenrechnungsnummernkreis
            If S_ReOf(mPaNr, MaNum, TyStr) > 0 Then
                SPopu Tit1, Mld1, IC48_Information
            End If
        Else
            If S_ReOf(mPaNr, 0, TyStr) > 0 Then
                SPopu Tit1, Mld1, IC48_Information
            End If
        End If
    Case RibTab_Rechnungen:
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        TxKur.SetFocus
    End Select
    
    If TyStr = "U" Or TyStr = "V" Or TyStr = "M" Then
        TxRen.Enabled = False
    Else
        If GlReN = False Then 'Rechnungsnummern sofort erzeugen
            TxRen.Enabled = False
        Else
            If UBound(GlRec) > 0 Then
                TxRen.Enabled = False
            Else
                TxRen.Enabled = True
            End If
        End If
    End If

ElseIf Rahm6.Visible = True Then
                        
        Select Case GlBut
        Case RibTab_Abrechnung:
                Set RpCls = RpCo3.Columns
                Set RpSel = RpCo3.SelectedRows
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rec_IDR)
                NeuRe = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Betrag)
                ReBet = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Rec_Selekt)
                ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Rec_Rabatt)
                ReRab = Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1)
                Set RpCol = RpCls.Find(Rec_Bezahlt)
                BeBet = RpRow.Record(RpCol.ItemIndex).Value
        Case RibTab_Rechnungen:
                Set RpCls = RpCo4.Columns
                Set RpSel = RpCo4.SelectedRows
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rec_IDR)
                NeuRe = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Betrag)
                ReBet = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Rec_Rabatt)
                ReRab = Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1)
                Set RpCol = RpCls.Find(Rec_Bezahlt)
                BeBet = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Selekt)
                If LCase(RpRow.Record(RpCol.ItemIndex).Value) = "ja" Then
                    ReAbg = True
                Else
                    ReAbg = False
                End If
        End Select
    
        AktTa = 0
        ReTyp = CmTyp.ListIndex
        BehEr = CBool(OpBEr.Value)

        AnzBl = DaPi2.Selection.BlocksCount
        With DaPi2
            DayFi = .FirstDayOfWeek
            DayLa = .LastVisibleDay
            If ReTyp = 4 Then 'Gutschrift
                .EnsureVisible Date
                .Select Date
            End If
        End With

        If AnzBl = 1 Then
            DaBeg = DaPi2.Selection(0).DateBegin
            DaEnd = DaPi2.Selection(0).DateEnd
            If DaEnd > DaBeg Then
                Do
                DaAkt = DaBeg + AktTa
                AktTa = AktTa + 1
                ReDim Preserve GlTag(AktTa)
                GlTag(AktTa) = DaAkt
                Loop Until DaAkt >= DaEnd
            Else
                ReDim Preserve GlTag(1)
                GlTag(1) = DaBeg
            End If
        ElseIf AnzBl > 1 Then
            For AktBl = 0 To AnzBl - 1
                DaBeg = DaPi2.Selection.Blocks(AktBl).DateBegin
                DaEnd = DaPi2.Selection.Blocks(AktBl).DateEnd
                If DaEnd > DaBeg Then
                    BloTa = 0
                    Do
                    DaAkt = DaBeg + BloTa
                    AktTa = AktTa + 1
                    BloTa = BloTa + 1
                    ReDim Preserve GlTag(AktTa)
                    GlTag(AktTa) = DaAkt
                    Loop Until DaAkt >= DaEnd
                Else
                    AktTa = AktTa + 1
                    ReDim Preserve GlTag(AktTa)
                    GlTag(AktTa) = DaBeg
                End If
            Next AktBl
        End If

        If ReTyp = 4 Then 'Gutschrift
            GuiID = CreateID("R")
            If mReBz > 0 Then
                GuBet = mReBz
            Else
                GuBet = mReBe
            End If
            S_KrGu GlTag(1), GuBet, GuiID, mReSt 'legt einen Gutschrifteintrag an
            DoEvents
            NeReB = S_ReBet(NeuRe, Round(ReBet, 2), ReAbg, ReRab) 'passt den Rechnungsbetrag an
            DoEvents
            RetWe = S_KrBe(NeuRe, ZaZil, ReBet, BeBet) 'Rechnet den Endbetrag aus
            DoEvents
            SUpAb
            DoEvents
            SAnza 'Sperrt den Akontobutton
            DoEvents
            GlVzA = True 'Rechnungsübersichtverzögerung Aktualisieren
            
            FReNu 'Reindizierung
            DoEvents
            Select Case GlBut
            Case RibTab_Abrechnung:
            Case RibTab_Rechnungen: SReZe mPaNr
            End Select
            DoEvents
            Unload Me
            
        Else

            If AnzBl > 0 Then
                If mReNr > 0 Then
                    S_KrKp mReNr, NeuRe, mPaNr, BehEr 'kopiert Positionen aus der vorherigen Rechnung
                    DoEvents
                    S_AbTe DayFi, DayLa 'markiert die belegten Tage im Kalender Bold
                    DoEvents
                    NeReB = S_ReBet(NeuRe, Round(ReBet, 2), ReAbg, ReRab) 'passt den Rechnungsbetrag an
                    DoEvents
                    RetWe = S_KrBe(NeuRe, ZaZil, ReBet, BeBet) 'rechnet den Endbetrag aus
                    DoEvents
                    SUpAb
                    DoEvents
                    SAnza 'sperrt den Akontobutton
                    DoEvents
                    ReDim Preserve GlTag(1)
                    GlTag(1) = Date
                End If
                
                GlVzA = True 'Rechnungsübersichtverzögerung Aktualisieren
            End If

            FReNu 'Reindizierung
            DoEvents
            Select Case GlBut
            Case RibTab_Abrechnung:
            Case RibTab_Rechnungen: SReZe mPaNr
            End Select
            DoEvents
            Unload Me
        End If

End If

PuBu3.Enabled = True
Screen.MousePointer = vbNormal

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub FZuru()
On Error Resume Next

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set Rahm6 = Me.frmRahm6
Set PuBu4 = Me.btnZuruk

If Rahm2.Visible = True Then
    Rahm6.Visible = False
    Rahm5.Visible = False
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = True
    PuBu4.Enabled = False
ElseIf Rahm3.Visible = True Then
    Rahm6.Visible = False
    Rahm5.Visible = False
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = True
    Rahm1.Visible = False
ElseIf Rahm4.Visible = True Then
    Select Case GlBut
    Case RibTab_Startseite:
        Rahm6.Visible = False
        Rahm5.Visible = False
        Rahm4.Visible = False
        Rahm3.Visible = True
        Rahm2.Visible = False
        Rahm1.Visible = False
    Case RibTab_Abrechnung:
        Rahm6.Visible = False
        Rahm5.Visible = False
        Rahm4.Visible = False
        Rahm3.Visible = False
        Rahm2.Visible = False
        Rahm1.Visible = True
        PuBu4.Enabled = False
    Case RibTab_Rechnungen:
        Rahm6.Visible = False
        Rahm5.Visible = False
        Rahm4.Visible = False
        Rahm3.Visible = True
        Rahm2.Visible = False
        Rahm1.Visible = False
    End Select
ElseIf Rahm6.Visible = True Then
    Rahm6.Visible = False
    Rahm5.Visible = False
    Rahm4.Visible = True
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = False
ElseIf Rahm5.Visible = True Then
    Rahm6.Visible = False
    Rahm5.Visible = False
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = True
    PuBu4.Enabled = False
End If

Screen.MousePointer = vbNormal

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

If GlRKo = True Then 'Rechnung Kopieren
    TeTit = IniGetOpt("Hilfe", 50731)
    TeMai = IniGetOpt("Hilfe", 50732)
    TeInh = IniGetOpt("Hilfe", 50733)
    TeFus = IniGetOpt("Hilfe", 50734)
Else
    Select Case GlBut
    Case RibTab_Startseite:
        TeTit = IniGetOpt("Hilfe", 50721)
        TeMai = IniGetOpt("Hilfe", 50722)
        TeInh = IniGetOpt("Hilfe", 50723)
        TeFus = IniGetOpt("Hilfe", 50724)
    Case RibTab_Abrechnung:
        TeTit = IniGetOpt("Hilfe", 50711)
        TeMai = IniGetOpt("Hilfe", 50712)
        TeInh = IniGetOpt("Hilfe", 50713)
        TeFus = IniGetOpt("Hilfe", 50714)
    Case RibTab_Rechnungen:
        TeTit = IniGetOpt("Hilfe", 50721)
        TeMai = IniGetOpt("Hilfe", 50722)
        TeInh = IniGetOpt("Hilfe", 50723)
        TeFus = IniGetOpt("Hilfe", 50724)
    End Select
End If
    
SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    FReNu 'Reindizierung
    DoEvents
    Unload Me
End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub

Private Sub btnWeite_Click()
    FWeit
End Sub
Private Sub btnZuruk_Click()
    FZuru
End Sub
Private Sub chkDiaKr_Click()

Set ChDiK = Me.chkDiaKr

If ChDiK.Value = xtpChecked Then
    IniSetVal "System", "DiKrUb", -1
    GlDiK = True
Else
    IniSetVal "System", "DiKrUb", 0
    GlDiK = False
End If

End Sub

Private Sub chkDiaRe_Click()

Set ChDiR = Me.chkDiaRe

If ChDiR.Value = xtpChecked Then
    IniSetVal "System", "DiReUb", -1
    GlDiR = True
Else
    IniSetVal "System", "DiReUb", 0
    GlDiR = False
End If

End Sub

Private Sub chkPhyAb_Click()

Set ChPhy = Me.chkPhyAb

If ChPhy.Value = xtpChecked Then
    IniSetVal "System", "PhyAbr", -1
    GlPhs = True
Else
    IniSetVal "System", "PhyAbr", 0
    GlPhs = False
End If

End Sub

Private Sub chkReNum_Click()
    
Set ChRnm = Me.chkReNum

If ChRnm.Value = xtpChecked Then
    GlReN = True
Else
    GlReN = False
End If

S_SeSe 5, , , , GlReN
    
End Sub

Private Sub cmbManda_Click()
    If FoLad = False Then
        FVoSe
    End If
End Sub
Private Sub cmbReStu_Click()
On Error Resume Next

Set CmStu = Me.cmbReStu
Set CmTyp = Me.cmbReTyp

If FoLad = False Then
    If CmStu.ListIndex > 0 Then
        If ReTyp = vbNullString Then
            CmTyp.ListIndex = 6
            ReTyp = "G"
        End If
    End If
End If

End Sub

Private Sub cmbReTyp_Click()
On Error Resume Next

Dim LiIdx As Integer
    
Set CmTyp = Me.cmbReTyp

If FoLad = False Then
    LiIdx = CmTyp.ListIndex
    Select Case LiIdx
    Case 0: ReTyp = "R"
    Case 1: ReTyp = "V"
    Case 2: ReTyp = "L"
    Case 3: ReTyp = "A"
    Case 4: ReTyp = "U"
    Case 5: ReTyp = "M"
    Case 6: ReTyp = "G"
    Case 7: ReTyp = "I"
    End Select
    FTyPr
    FCapt
    FReRc
End If

End Sub

Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDat1
End Sub

Private Sub dtpDatu2_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
On Error Resume Next

Dim AktTa As Long

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTaV > 0 Then
    For AktTa = 1 To GlTaV
        If Day = GlBet(AktTa) Then
            Metrics.BackColor = GlMkr
        End If
    Next AktTa
End If

End Sub
Private Sub dtpDatu2_MonthChanged()
On Error Resume Next

Dim DayFi As Date
Dim DayLa As Date

Set DaPi2 = Me.dtpDatu2

With DaPi2
    DayFi = .FirstDayOfWeek
    DayLa = .LastVisibleDay
End With

S_AbTe DayFi, DayLa

Set DaPi1 = Nothing

End Sub

Private Sub dtpDatu2_SelectionChanged()
    FDat2
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF11: Unload Me
    End Select
End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True
If GlRtT <> vbNullString Then 'Standardbelegtyp Temporär
    ReTyp = GlRtT
Else
    ReTyp = GlReT 'Standardbelegtyp
End If
FInit
FCapt
FKonf
FoLad = False
AFont Me
SFrame 1, Me.hwnd

GlRtT = vbNullString 'WICHTIG!

End Sub

Private Sub Form_Unload(Cancel As Integer)
    GlRKo = False 'Rechnung kopieren
    Set frmReNeu = Nothing
End Sub

Private Sub optBehEr_Click()
    FErst
End Sub

Private Sub optBehLe_Click()
    FErst
End Sub
Private Sub repCont5_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If Row.GroupRow = False Then
    Select Case Row.Record(5).Value
    Case "M": Metrics.ForeColor = 16744448
    Case "L": Metrics.ForeColor = 33023
    Case "V": Metrics.ForeColor = 8421631
    Case "I": Metrics.ForeColor = 13138080
    Case "U": Metrics.ForeColor = 6604830
    Case Else:
            If Row.Record(2).Value = 0 Then
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
    End Select
    If Row.Record(2).Value = 0 Then
        Metrics.Font.Bold = True
    End If
    If Row.Record(4).Value = True Then
        Metrics.Font.Strikethrough = True
        Metrics.ForeColor = 8421504
    End If
End If

End Sub

Private Sub txtBemer_GotFocus()
    FRest
End Sub
Private Sub txtBemer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub
Private Sub txtDatu1_Validate(Cancel As Boolean)
    FReKo
End Sub
Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
End Sub


Private Sub txtKurz_GotFocus()
    FRest
End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtPost_GotFocus()
    FRest
End Sub
Private Sub txtPost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtReKop_GotFocus()
    Me.txtReKop.SelStart = 0
    Me.txtReKop.SelLength = Len(Me.txtReKop.Text)
End Sub

Private Sub txtReNum_GotFocus()
    Me.txtReNum.SelStart = 0
    Me.txtReNum.SelLength = Len(Me.txtReNum.Text)
End Sub
Private Sub txtReNum_KeyDown(KeyCode As Integer, Shift As Integer)
    ReAnp = True
End Sub
Private Sub chkExpMo_Click()
    FExp
End Sub

Private Sub txtRzNum_GotFocus()
    Me.txtRzNum.SelStart = 0
    Me.txtRzNum.SelLength = Len(Me.txtRzNum.Text)
End Sub
