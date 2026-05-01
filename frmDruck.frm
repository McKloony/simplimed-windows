VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmDruck 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Drucken"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   47
      Top             =   4900
      Width           =   8500
      _Version        =   1048579
      _ExtentX        =   14993
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   6500
         TabIndex        =   48
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
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   5100
         TabIndex        =   49
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
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   3800
         TabIndex        =   50
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
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   2700
      Left            =   700
      TabIndex        =   41
      Top             =   700
      Visible         =   0   'False
      Width           =   7100
      _Version        =   1048579
      _ExtentX        =   12524
      _ExtentY        =   4762
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox teiRahm3 
         Height          =   2600
         Left            =   3700
         TabIndex        =   44
         Top             =   0
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   4586
         _StockProps     =   79
         Caption         =   "Ausgabeeinstellungen"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox chkXRech 
            Height          =   220
            Left            =   500
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Gibt die markieren Rechnungen im X-Rechnungsformat aus"
            Top             =   780
            Width           =   2500
            _Version        =   1048579
            _ExtentX        =   4410
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "E-Rechnungsausgabe"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkDrDia 
            Height          =   220
            Left            =   500
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Zeigt einen Dialog zur Auswahl des Druckers und reduziert die Anzahl der Ausdrucke auf eins"
            Top             =   420
            Width           =   2500
            _Version        =   1048579
            _ExtentX        =   4410
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Druckerauswahldialog"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeCalendarControl.DatePicker dtpDatu1 
            Height          =   500
            Left            =   2800
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   100
            Visible         =   0   'False
            Width           =   500
            _Version        =   1048579
            _ExtentX        =   882
            _ExtentY        =   882
            _StockProps     =   64
            Show3DBorder    =   2
         End
         Begin XtremeSuiteControls.UpDown updCont1 
            Height          =   350
            Left            =   1810
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1860
            Width           =   255
            _Version        =   1048579
            _ExtentX        =   450
            _ExtentY        =   600
            _StockProps     =   64
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDatu1"
            BuddyProperty   =   ""
         End
         Begin XtremeSuiteControls.FlatEdit txtDatu1 
            Height          =   350
            Left            =   500
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1860
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
         Begin XtremeSuiteControls.PushButton btnDatu1 
            Height          =   350
            Left            =   2090
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Öffnet den Auswahlkalender"
            Top             =   1860
            Width           =   350
            _Version        =   1048579
            _ExtentX        =   617
            _ExtentY        =   617
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkRePru 
            Height          =   220
            Left            =   500
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Zeigt die Dateinamen der Formulare anstelle dessen Bezeichnungen"
            Top             =   1110
            Width           =   2500
            _Version        =   1048579
            _ExtentX        =   4410
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Formulardateinamen zeigen"
            UseVisualStyle  =   -1  'True
            MultiLine       =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkReDat 
            Height          =   220
            Left            =   500
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Ändert das Rechnungsdatum bei noch nicht abgeschlossenen Rechnungen"
            Top             =   1480
            Width           =   2500
            _Version        =   1048579
            _ExtentX        =   4410
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Rechnungsdatum angleichen"
            UseVisualStyle  =   -1  'True
            MultiLine       =   0   'False
         End
      End
      Begin XtremeSuiteControls.GroupBox teiRahm2 
         Height          =   1200
         Left            =   100
         TabIndex        =   43
         Top             =   1400
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   2117
         _StockProps     =   79
         Caption         =   "Formularauswahl"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cmbReFor 
            Height          =   310
            Left            =   300
            TabIndex        =   3
            Top             =   460
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
      End
      Begin XtremeSuiteControls.GroupBox teiRahm1 
         Height          =   1300
         Left            =   100
         TabIndex        =   42
         Top             =   0
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   2293
         _StockProps     =   79
         Caption         =   "Rechnungsabschluss"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton optOpti2 
            Height          =   220
            Left            =   300
            TabIndex        =   2
            ToolTipText     =   "Belässt die Rechnung als nicht gedruckt und verriegelt diese nicht"
            Top             =   780
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Rechnung noch nicht abschließen"
            ForeColor       =   192
            UseVisualStyle  =   -1  'True
            MultiLine       =   0   'False
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optOpti1 
            Height          =   220
            Left            =   300
            TabIndex        =   1
            ToolTipText     =   "Kennzeichnet die Rechnung als gedruckt und verriegelt diese"
            Top             =   380
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Rechnung jetzt abschließen"
            ForeColor       =   49152
            UseVisualStyle  =   -1  'True
            MultiLine       =   0   'False
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   2700
      Left            =   700
      TabIndex        =   37
      Top             =   700
      Visible         =   0   'False
      Width           =   7100
      _Version        =   1048579
      _ExtentX        =   12524
      _ExtentY        =   4762
      _StockProps     =   79
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox teiRahm6 
         Height          =   2600
         Left            =   3700
         TabIndex        =   40
         Top             =   0
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   4586
         _StockProps     =   79
         Caption         =   "Ausgabeoptionen"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox chkRechn 
            Height          =   220
            Left            =   300
            TabIndex        =   34
            Top             =   1130
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Rechnung mit ausgeben"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkMaGeb 
            Height          =   220
            Left            =   300
            TabIndex        =   35
            Top             =   1500
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Mahngebühr hinzufügen :"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkMeRec 
            Height          =   220
            Left            =   300
            TabIndex        =   32
            Top             =   380
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Mahnfristdatum autom. anpassen"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkMaLis 
            Height          =   220
            Left            =   300
            TabIndex        =   33
            Top             =   760
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Offene-Posten Liste drucken"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtMahng 
            Height          =   310
            Left            =   300
            TabIndex        =   36
            Top             =   1860
            Width           =   1200
            _Version        =   1048579
            _ExtentX        =   2117
            _ExtentY        =   547
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            Enabled         =   0   'False
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox teiRahm5 
         Height          =   1200
         Left            =   100
         TabIndex        =   39
         Top             =   1400
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   2117
         _StockProps     =   79
         Caption         =   "Formularauswahl"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cmbMaFor 
            Height          =   310
            Left            =   300
            TabIndex        =   31
            Top             =   460
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
      End
      Begin XtremeSuiteControls.GroupBox teiRahm4 
         Height          =   1300
         Left            =   100
         TabIndex        =   38
         Top             =   0
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   2293
         _StockProps     =   79
         Caption         =   "Mahnstufe"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton optOpti4 
            Height          =   220
            Left            =   300
            TabIndex        =   30
            Top             =   780
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Mahnstufe noch nicht erhöhen"
            ForeColor       =   192
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optOpti3 
            Height          =   220
            Left            =   300
            TabIndex        =   29
            Top             =   380
            Width           =   3000
            _Version        =   1048579
            _ExtentX        =   5292
            _ExtentY        =   388
            _StockProps     =   79
            Caption         =   "Mahnstufe automatisch erhöhen"
            ForeColor       =   49152
            UseVisualStyle  =   -1  'True
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   1300
      Left            =   800
      TabIndex        =   14
      Top             =   3460
      Width           =   6900
      _Version        =   1048579
      _ExtentX        =   12171
      _ExtentY        =   2293
      _StockProps     =   79
      Caption         =   "Ausgabeeinstellungen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cmbEmail 
         Height          =   315
         Left            =   300
         TabIndex        =   26
         Top             =   420
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.CheckBox chkDruSe 
         Height          =   225
         Left            =   4000
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Druckt jede Rechnung in einem separaten Druckauftrag und nicht ein einem Druckauftrag"
         Top             =   420
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Druckauftrags-Separierung"
         UseVisualStyle  =   -1  'True
         MultiLine       =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkDuble 
         Height          =   225
         Left            =   4000
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Aktiviert den zweiseitigen Druck, falls der Druckertreiber dieses unterstützt"
         Top             =   800
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Duplexdruck forcieren"
         UseVisualStyle  =   -1  'True
         MultiLine       =   0   'False
      End
   End
   Begin VB.TextBox txoDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   7500
      Width           =   80
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   2600
      Left            =   800
      TabIndex        =   13
      Top             =   700
      Visible         =   0   'False
      Width           =   6900
      _Version        =   1048579
      _ExtentX        =   12171
      _ExtentY        =   4586
      _StockProps     =   79
      Caption         =   "Bitte wählen Sie eine Druckoption:"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox teiRahm7 
         Height          =   1500
         Left            =   100
         TabIndex        =   46
         Top             =   1000
         Visible         =   0   'False
         Width           =   6700
         _Version        =   1048579
         _ExtentX        =   11818
         _ExtentY        =   2646
         _StockProps     =   79
         Caption         =   "Einnahmebuchung"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.CheckBox chkGutsh 
            Height          =   230
            Left            =   4700
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Die Buchung wird ohne Patientenzuordnung generiert, so dass dieser später zugeordnet werden kann"
            Top             =   1000
            Width           =   1900
            _Version        =   1048579
            _ExtentX        =   3351
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gutscheinbuchung"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkGutha 
            Height          =   230
            Left            =   4700
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Dem zugeordneten Patienten wird der Buchungsbetrag als Guthaben hinzugefügt"
            Top             =   600
            Width           =   1900
            _Version        =   1048579
            _ExtentX        =   3351
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Guthabenbuchung"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbKonto 
            Height          =   310
            Left            =   700
            TabIndex        =   17
            Top             =   500
            Width           =   3600
            _Version        =   1048579
            _ExtentX        =   6350
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Style           =   2
            Text            =   "ComboBox2"
         End
         Begin XtremeSuiteControls.ComboBox cmbGegen 
            Height          =   310
            Left            =   700
            TabIndex        =   16
            Top             =   0
            Width           =   3600
            _Version        =   1048579
            _ExtentX        =   6350
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Style           =   2
            Text            =   "ComboBox3"
         End
         Begin XtremeSuiteControls.ComboBox cmbStKto 
            Height          =   315
            Left            =   700
            TabIndex        =   18
            Top             =   1000
            Width           =   3600
            _Version        =   1048579
            _ExtentX        =   6350
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            Enabled         =   0   'False
            Style           =   2
            Text            =   "ComboBox1"
         End
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   340
         Left            =   5120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1830
         Visible         =   0   'False
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Min             =   1
         Value           =   2
         Max             =   999
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtKopie"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.RadioButton optAdre9 
         Height          =   225
         Left            =   800
         TabIndex        =   23
         Top             =   1900
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Den markierten Eintrag mehrfach drucken"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optAdre8 
         Height          =   220
         Left            =   800
         TabIndex        =   22
         Top             =   1550
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Alle Einträge in der Auswahl drucken"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optAdre7 
         Height          =   220
         Left            =   800
         TabIndex        =   21
         Top             =   1200
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Nur den/die markieren Einträge drucken"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbLiDru 
         Height          =   310
         Left            =   800
         TabIndex        =   15
         Top             =   500
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
         DropDownItemCount=   12
      End
      Begin XtremeSuiteControls.FlatEdit txtKopie 
         Height          =   310
         Left            =   4300
         TabIndex        =   24
         Top             =   1850
         Visible         =   0   'False
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Text            =   "2"
         Alignment       =   2
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   2700
      Left            =   700
      TabIndex        =   45
      Top             =   700
      Visible         =   0   'False
      Width           =   7100
      _Version        =   1048579
      _ExtentX        =   12524
      _ExtentY        =   4762
      _StockProps     =   79
      BorderStyle     =   2
   End
   Begin VB.Label Lab01 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte markieren Sie, welche Aktionen nun durchgeführt werden sollen und klicken auf Weiter"
      Height          =   220
      Left            =   720
      TabIndex        =   12
      Top             =   120
      Width           =   7000
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   500
      Left            =   0
      Top             =   0
      Width           =   8500
   End
End
Attribute VB_Name = "frmDruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private UpDo1 As XtremeSuiteControls.UpDown
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxKop As XtremeSuiteControls.FlatEdit
Private TxMah As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private TeRa1 As XtremeSuiteControls.GroupBox
Private TeRa2 As XtremeSuiteControls.GroupBox
Private TeRa3 As XtremeSuiteControls.GroupBox
Private TeRa4 As XtremeSuiteControls.GroupBox
Private TeRa5 As XtremeSuiteControls.GroupBox
Private TeRa6 As XtremeSuiteControls.GroupBox
Private TeRa7 As XtremeSuiteControls.GroupBox
Private CmReF As XtremeSuiteControls.ComboBox
Private CmMaF As XtremeSuiteControls.ComboBox
Private CmLis As XtremeSuiteControls.ComboBox
Private CmEml As XtremeSuiteControls.ComboBox
Private CmKto As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmStu As XtremeSuiteControls.ComboBox
Private ChMaL As XtremeSuiteControls.CheckBox
Private ChReD As XtremeSuiteControls.CheckBox
Private ChMaR As XtremeSuiteControls.CheckBox
Private ChMaG As XtremeSuiteControls.CheckBox
Private ChXRe As XtremeSuiteControls.CheckBox
Private ChePr As XtremeSuiteControls.CheckBox
Private CheDu As XtremeSuiteControls.CheckBox
Private CheSe As XtremeSuiteControls.CheckBox
Private CheGu As XtremeSuiteControls.CheckBox
Private CheGs As XtremeSuiteControls.CheckBox
Private CheRe As XtremeSuiteControls.CheckBox
Private CheDr As XtremeSuiteControls.CheckBox
Private MoKal As XtremeCalendarControl.DatePicker
Private OpOp1 As XtremeSuiteControls.RadioButton
Private OpOp2 As XtremeSuiteControls.RadioButton
Private OpOp3 As XtremeSuiteControls.RadioButton
Private OpOp4 As XtremeSuiteControls.RadioButton
Private OpAd7 As XtremeSuiteControls.RadioButton
Private OpAd8 As XtremeSuiteControls.RadioButton
Private OpAd9 As XtremeSuiteControls.RadioButton
Private PuBu1 As XtremeSuiteControls.PushButton
Private CoDia As XtremeSuiteControls.CommonDialog
Private ImMan As XtremeCommandBars.ImageManager
Private RpRow As XtremeReportControl.ReportRow
Private RpSel As XtremeReportControl.ReportSelectedRows

Public EmVer As Integer
Public DrTyp As Integer
Public EmSep As Boolean

Private AlGeb As Single
Private ReTyp As String
Private ReIdx As Integer
Private MaIdx As Integer
Private EtIdx As Integer
Private AdIdx As Integer
Private TeIdx As Integer
Private LsIdx As Integer
Private DrFrm As Boolean
Private FrLoa As Boolean

Private clFil As clsFile
Private clLis As clsLisLab
Private clFen As clsFenster

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E

Private WithEvents clDru As clsDruck
Attribute clDru.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolliert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
    TxDa1.Text = FDaPr(NeuDa)
    With MoKal
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
    If NeuDa > Date Then
        SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
    End If
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

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    TxDa1.Text = FDaPr(NeuDa)
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Function FDaPr(ByVal NeuDa As Date) As Date
On Error GoTo OrErr

Dim ReDat As Date
Dim AnzPo As Integer
Dim ReAbg As Boolean
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select
Set RpRow = RpSel(0)

If RpRow.GroupRow = False Then
    Set RpCol = RpCls.Find(Rec_Datum)
    ReDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
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
Else
    ReDat = Date
    ReAbg = False
End If

AnzPo = RpSel.Count

If GlRnm = True Then 'Neustart der Rechnungsnummer am Jahresanfang
    If AnzPo > 1 Then
        If Year(NeuDa) <> Year(Date) Then
            FDaPr = Date
            SPopu "Ungültiges Rechnungsdatum", "Das Rechnungsdatum muss sich innerhalb des aktuellen Geschäftsjahrs befinden", IC48_Information
        Else
            FDaPr = NeuDa
        End If
    ElseIf AnzPo = 1 Then
        If ReAbg = False Then
            If Year(NeuDa) <> Year(ReDat) Then
                FDaPr = ReDat
                SPopu "Ungültiges Rechnungsdatum", "Das Rechnungsdatum muss sich innerhalb des aktuellen Geschäftsjahrs befinden", IC48_Information
            Else
                FDaPr = NeuDa
            End If
        Else
            FDaPr = NeuDa
        End If
    End If
Else
    FDaPr = NeuDa
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Function

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaPr " & Err.Number
Resume Next

End Function
Private Sub FFoSe(Optional ByVal LiDru As Boolean = False)
On Error GoTo OrErr

Dim LiIdx As Integer

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set CmReF = Me.cmbReFor
Set CmMaF = Me.cmbMaFor
Set CmLis = Me.cmbLiDru
Set CheRe = Me.chkRechn
Set ChXRe = Me.chkXRech
Set TxKop = Me.txtKopie
Set OpAd7 = Me.optAdre7
Set OpAd9 = Me.optAdre9

If Rahm4.Visible = True Then
    If WindowLoad("frmAdress") = True Then
        LiIdx = CmLis.ListIndex + 1
        IniSetVal "System", "DrForm", "D" & Format$(LiIdx, "00")
    Else
        Select Case GlBut
        Case RibTab_Adressen:
            LiIdx = CmLis.ListIndex + 1
            IniSetVal "System", "AdForm", "A" & Format$(LiIdx, "00")
        Case RibTab_Mandanten:
            LiIdx = CmLis.ListIndex + 1
            IniSetVal "System", "AdForm", "A" & Format$(LiIdx, "00")
        Case RibTab_Verordner:
            LiIdx = CmLis.ListIndex + 1
            IniSetVal "System", "AdForm", "A" & Format$(LiIdx, "00")
        Case RibTab_Mitarbeit:
             LiIdx = CmLis.ListIndex + 1
             IniSetVal "System", "AdForm", "A" & Format$(LiIdx, "00")
        Case RibTab_Vorbereit:
             LiIdx = CmLis.ListIndex + 1
             IniSetVal "System", "AdForm", "A" & Format$(LiIdx, "00")
        Case RibTab_Rechnungen:
            Select Case DrTyp
            Case 1:
                LiIdx = CmLis.ListIndex + 1
                IniSetVal "System", "DrForm", "D" & Format$(LiIdx, "00")
            Case 4:
                LiIdx = CmLis.ListIndex + 1
                IniSetVal "System", "LiForm", "L" & Format$(LiIdx, "00")
            End Select
        Case RibTab_Mahnwesen:
            Select Case DrTyp
            Case 1:
                LiIdx = CmLis.ListIndex + 1
                IniSetVal "System", "DrForm", "D" & Format$(LiIdx, "00")
            Case 5:
            End Select
        Case RibTab_Ter_Listen:
            LiIdx = CmLis.ListIndex + 1
            IniSetVal "System", "TeForm", "T" & Format$(LiIdx, "00")
        Case RibTab_Ter_Akont:
            LiIdx = CmLis.ListIndex + 1
            IniSetVal "System", "TeForm", "T" & Format$(LiIdx, "00")
        Case RibTab_Ter_Warte:
            LiIdx = CmLis.ListIndex + 1
            IniSetVal "System", "TeForm", "T" & Format$(LiIdx, "00")
        Case RibTab_Rezeptmodul:
            Select Case DrTyp
            Case 6:
                LiIdx = CmLis.ListIndex + 1
            Case 1:
                LiIdx = CmLis.ListIndex + 1
                IniSetVal "System", "DrForm", "D" & Format$(LiIdx, "00")
            End Select
        Case RibTab_Belegmodul:
            Select Case DrTyp
            Case 6:
                LiIdx = CmLis.ListIndex + 1
            Case 1:
                LiIdx = CmLis.ListIndex + 1
                IniSetVal "System", "DrForm", "D" & Format$(LiIdx, "00")
            End Select
        Case Else:
            LiIdx = CmLis.ListIndex + 1
            IniSetVal "System", "DrForm", "D" & Format$(LiIdx, "00")
        End Select
    End If
ElseIf Rahm1.Visible = True Then
    Select Case GlBut
    Case RibTab_Abrechnung:
        LiIdx = CmReF.ListIndex + 1
        IniSetVal "System", "ReForm", "R" & Format$(LiIdx, "00")
    Case RibTab_Rechnungen:
        LiIdx = CmReF.ListIndex + 1
        IniSetVal "System", "ReForm", "R" & Format$(LiIdx, "00")
    End Select
ElseIf Rahm2.Visible = True Then
    LiIdx = CmMaF.ListIndex + 1
    IniSetVal "System", "MaForm", "M" & Format$(LiIdx, "00")
    If LiIdx = 2 Then
        CheRe.Enabled = False
        CheRe.Value = xtpUnchecked
    Else
        CheRe.Enabled = True
    End If
End If

If LiDru = True Then
    Select Case GlBut
    Case RibTab_Adressen:
        Select Case LiIdx
        Case 2:
            OpAd9.Enabled = True
        Case 6:
            OpAd9.Enabled = True
        Case Else:
            If OpAd9.Value = True Then OpAd7.Value = True
            OpAd9.Enabled = False
        End Select
    Case RibTab_Ter_Listen:
        If LiIdx = 3 Then
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = False
        Else
            OpAd7.Enabled = True
            OpAd8.Enabled = True
            OpAd9.Enabled = False
        End If
    Case RibTab_Ter_Akont:
        If LiIdx = 3 Then
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = False
        Else
            OpAd7.Enabled = True
            OpAd8.Enabled = True
            OpAd9.Enabled = False
        End If
    Case RibTab_Ter_Warte:
        If LiIdx = 3 Then
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = False
        Else
            OpAd7.Enabled = True
            OpAd8.Enabled = True
            OpAd9.Enabled = False
        End If
    End Select
    
    If OpAd9.Value = True Then
        TxKop.Visible = True
    Else
        TxKop.Visible = False
    End If
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FFoSe " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo InErr

Dim ReDat As Date
Dim RetWe As Long
Dim AnzRe As Long
Dim AnzRz As Long
Dim ManNr As Long
Dim StaGe As Long
Dim StaKt As Long
Dim FoIdx As Integer
Dim StaRa As Integer
Dim IdStK As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim GesZa As Integer
Dim ReAbs As Boolean
Dim MahSt As Boolean
Dim ReAbg As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set CmBrs = FM.comBar01
Set ImMan = FM.imgManag
Set CoDia = FM.comDialo
Set RpCo1 = FM.repCont1
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set TxMah = Me.txtMahng
Set TxDa1 = Me.txtDatu1
Set TxKop = Me.txtKopie
Set CmReF = Me.cmbReFor
Set CmMaF = Me.cmbMaFor
Set CmEml = Me.cmbEmail
Set CmKto = Me.cmbKonto
Set CmGeg = Me.cmbGegen
Set CmStu = Me.cmbStKto
Set OpAd7 = Me.optAdre7
Set OpAd8 = Me.optAdre8
Set OpAd9 = Me.optAdre9
Set ChReD = Me.chkReDat
Set ChXRe = Me.chkXRech
Set ChMaR = Me.chkMeRec
Set ChMaL = Me.chkMaLis
Set CheRe = Me.chkRechn
Set CheDr = Me.chkDrDia
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set TeRa1 = Me.teiRahm1
Set TeRa2 = Me.teiRahm2
Set TeRa3 = Me.teiRahm3
Set TeRa4 = Me.teiRahm4
Set TeRa5 = Me.teiRahm5
Set TeRa6 = Me.teiRahm6
Set TeRa7 = Me.teiRahm7
Set MoKal = Me.dtpDatu1
Set OpOp1 = Me.optOpti1
Set OpOp2 = Me.optOpti2
Set OpOp3 = Me.optOpti3
Set OpOp4 = Me.optOpti4
Set CmLis = Me.cmbLiDru
Set ChMaG = Me.chkMaGeb
Set ChePr = Me.chkRePru
Set CheDu = Me.chkDuble
Set CheSe = Me.chkDruSe
Set CheGu = Me.chkGutha
Set CheGs = Me.chkGutsh
Set PuBu1 = Me.btnDatu1
Set UpDo1 = Me.updCont1

Set clDru = New clsDruck
Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

ReAbs = CBool(IniGetVal("System", "RechAb"))
MahSt = CBool(IniGetVal("System", "MahnSt"))
DrFrm = CBool(IniGetVal("System", "DruFrm"))

ReIdx = Right$(IniGetVal("System", "ReForm"), 2) - 1
MaIdx = Right$(IniGetVal("System", "MaForm"), 2) - 1
EtIdx = Right$(IniGetVal("System", "DrForm"), 2) - 1
AdIdx = Right$(IniGetVal("System", "AdForm"), 2) - 1
TeIdx = Right$(IniGetVal("System", "TeForm"), 2) - 1
LsIdx = Right$(IniGetVal("System", "LiForm"), 2) - 1

Select Case GlBut
Case RibTab_Abrechnung:
            Set RpCls = RpCo3.Columns
            Set RpSel = RpCo3.SelectedRows
            AnzRe = RpSel.Count
            If AnzRe = 1 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rec_Type)
                ReTyp = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Datum)
                ReDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Rec_Selekt)
                ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                ReDat = Date
                ReAbg = False
                CheDr.Enabled = False
            End If
Case RibTab_Rechnungen:
            Set RpCls = RpCo4.Columns
            Set RpSel = RpCo4.SelectedRows
            AnzRe = RpSel.Count
            If AnzRe = 1 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rec_Type)
                ReTyp = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Datum)
                ReDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                Set RpCol = RpCls.Find(Rec_Selekt)
                If LCase(RpRow.Record(RpCol.ItemIndex).Value) = "ja" Then
                    ReAbg = True
                Else
                    ReAbg = False
                End If
            Else
                ReDat = Date
                ReAbg = False
                CheDr.Enabled = False
            End If
Case RibTab_Belegmodul:
            Set RpCls = RpCo5.Columns
            Set RpSel = RpCo5.SelectedRows
            AnzRz = RpSel.Count
            If AnzRz = 1 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rzp_IDP)
                ManNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rzp_Datum)
                ReDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                ReDat = Date
                ReAbg = False
                CheDr.Enabled = False
            End If
Case RibTab_Mahnwesen:
            Set RpCls = RpCo1.Columns
            Set RpSel = RpCo1.SelectedRows
            GesZa = RpSel.Count
            If GesZa > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(OPo_Gebuehr)
                If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                    AlGeb = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                End If
                If GesZa > 1 Then
                    CheRe.Enabled = False
                    ChMaL.Enabled = True
                End If
            End If
            ReDat = Date
            ReAbg = False
End Select

If ReAbg = False Then
    If Year(ReDat) = Year(Date) Then
        ReDat = Date
    ElseIf Year(ReDat) < Year(Date) Then
        ReDat = "31.12." & Year(ReDat)
    End If
End If

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

Select Case EmVer
Case 1: 'E-Mail-Versand
    With CmEml
        .AddItem "Emailversand an einen Patienten"
        .ItemData(0) = 1
        .AddItem "Emailversand an den Mandanten"
        .ItemData(1) = 2
        .AddItem "Emailversand an alle Patienten"
        .ItemData(2) = 3
        .ListIndex = GlEmA 'Emailafdresse Rechnungsversand
    End With
Case 6: 'Downloadlink
    With CmEml
        .AddItem "Downloadlink an einen Patienten"
        .ItemData(0) = 1
        .AddItem "Downloadlink an den Mandanten"
        .ItemData(1) = 2
        .AddItem "Downloadlink an alle Patienten"
        .ItemData(2) = 3
        .ListIndex = GlEmA 'Emailafdresse Rechnungsversand
    End With
Case Else
    With CmEml
        .AddItem "Beleg wird gedruckt"
        .ItemData(0) = 1
        .ListIndex = 0
    End With
End Select

With CmReF
    If DrFrm = True Then
        .AddItem GlFrm(1, 0)
        .ItemData(0) = 0
        .AddItem GlFrm(1, 1)
        .ItemData(1) = 1
        .AddItem GlFrm(1, 2)
        .ItemData(2) = 2
        .AddItem GlFrm(1, 3)
        .ItemData(3) = 3
        .AddItem GlFrm(1, 4)
        .ItemData(4) = 4
        .AddItem GlFrm(1, 5)
        .ItemData(5) = 5
        .AddItem GlFrm(1, 14)
        .ItemData(6) = 14
        .AddItem GlFrm(1, 9)
        .ItemData(7) = 9
        .AddItem GlFrm(1, 15)
        .ItemData(8) = 15
        .AddItem GlFrm(1, 6)
        .ItemData(9) = 6
        .AddItem GlFrm(1, 7)
        .ItemData(10) = 7
        .AddItem GlFrm(1, 16)
        .ItemData(11) = 16
        .AddItem GlFrm(1, 13)
        .ItemData(12) = 13
        .AddItem GlFrm(1, 17)
        .ItemData(13) = 17
        .AddItem GlFrm(1, 19)
        .ItemData(14) = 19
        .AddItem GlFrm(1, 78)
        .ItemData(15) = 78
        .AddItem GlFrm(1, 18)
        .ItemData(16) = 18
        .AddItem GlFrm(1, 96)
        .ItemData(17) = 96
    Else
        .AddItem GlFrm(0, 0)
        .ItemData(0) = 0
        .AddItem GlFrm(0, 1)
        .ItemData(1) = 1
        .AddItem GlFrm(0, 2)
        .ItemData(2) = 2
        .AddItem GlFrm(0, 3)
        .ItemData(3) = 3
        .AddItem GlFrm(0, 4)
        .ItemData(4) = 4
        .AddItem GlFrm(0, 5)
        .ItemData(5) = 5
        .AddItem GlFrm(0, 14)
        .ItemData(6) = 14
        .AddItem GlFrm(0, 9)
        .ItemData(7) = 9
        .AddItem GlFrm(0, 15)
        .ItemData(8) = 15
        .AddItem GlFrm(0, 6)
        .ItemData(9) = 6
        .AddItem GlFrm(0, 7)
        .ItemData(10) = 7
        .AddItem GlFrm(0, 16)
        .ItemData(11) = 16
        .AddItem GlFrm(0, 13)
        .ItemData(12) = 13
        .AddItem GlFrm(0, 17)
        .ItemData(13) = 17
        .AddItem GlFrm(0, 19)
        .ItemData(14) = 19
        .AddItem GlFrm(0, 78)
        .ItemData(15) = 78
        .AddItem GlFrm(0, 18)
        .ItemData(16) = 18
        .AddItem GlFrm(0, 96)
        .ItemData(17) = 96
    End If
End With

If ReTyp <> vbNullString Then
    Select Case ReTyp
    Case "L": RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, 2, ByVal 0&)
    Case "V": RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, 7, ByVal 0&)
    Case "U": RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, 8, ByVal 0&)
    Case Else: RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, ReIdx, ByVal 0&)
    End Select
Else
    RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, ReIdx, ByVal 0&)
End If

With CmMaF
    .AddItem "Einzelmahnungsformular"
    .ItemData(0) = 1
    .AddItem "Sammelmahnungsformular"
    .ItemData(1) = 2
    .AddItem "Einzelmahnung (Alternativ)"
    .ItemData(2) = 3
End With
RetWe = SendMessage(CmMaF.hwnd, CB_SETCURSEL, MaIdx, ByVal 0&)
If MaIdx = 1 Then
    CheRe.Enabled = False
End If

If WindowLoad("frmAdress") = True Then
    With CmLis
        .AddItem "Adressetiketten"
        .ItemData(0) = 1
        .AddItem "Photoetiketten"
        .ItemData(1) = 2
        .AddItem "Versicherungsetiketten"
        .ItemData(2) = 3
        .AddItem "Karteikartenetiketten"
        .ItemData(3) = 4
        .AddItem "Krankenaktenausdruck"
        .ItemData(4) = 5
        .AddItem "Diagnose-Etiketten"
        .ItemData(5) = 6
        .AddItem "Photoetiketten (Alternativ)"
        .ItemData(6) = 7
    End With
    RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, EtIdx, ByVal 0&)
    OpAd7.Enabled = False
    OpAd8.Enabled = False
    OpAd9.Enabled = True
Else
    Select Case GlBut
    Case RibTab_Adressen:
            With CmLis
                .AddItem "Adressenliste"
                .ItemData(0) = 1
                .AddItem "Adressetiketten"
                .ItemData(1) = 2
                .AddItem "Photoetiketten"
                .ItemData(2) = 3
                .AddItem "Patientenkrankenblatt"
                .ItemData(3) = 4
                .AddItem "Versicherungsetiketten"
                .ItemData(4) = 5
                .AddItem "Karteikartenetiketten"
                .ItemData(5) = 6
                .AddItem "Krankenaktenausdruck"
                .ItemData(6) = 7
                .AddItem "Diagnose-Etiketten"
                .ItemData(7) = 8
                .AddItem "Photoetiketten (Alternativ)"
                .ItemData(8) = 9
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, AdIdx, ByVal 0&)
            Select Case AdIdx
            Case 1:
                OpAd9.Enabled = True
            Case 5:
                OpAd9.Enabled = True
            Case Else:
                If OpAd9.Value = True Then OpAd7.Value = True
                OpAd9.Enabled = False
            End Select
    Case RibTab_Mandanten:
            With CmLis
                .AddItem "Adressenliste"
                .ItemData(0) = 1
                .AddItem "Adressetiketten"
                .ItemData(1) = 2
                .AddItem "Photoetiketten"
                .ItemData(2) = 3
                .AddItem "Patientenkrankenblatt"
                .ItemData(3) = 4
                .AddItem "Versicherungsetiketten"
                .ItemData(4) = 5
                .AddItem "Karteikartenetiketten"
                .ItemData(5) = 6
                .AddItem "Krankenaktenausdruck"
                .ItemData(6) = 7
                .AddItem "Diagnose-Etiketten"
                .ItemData(7) = 8
                .AddItem "Photoetiketten (Alternativ)"
                .ItemData(8) = 9
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, AdIdx, ByVal 0&)
    Case RibTab_Verordner:
            With CmLis
                .AddItem "Adressenliste"
                .ItemData(0) = 1
                .AddItem "Adressetiketten"
                .ItemData(1) = 2
                .AddItem "Photoetiketten"
                .ItemData(2) = 3
                .AddItem "Patientenkrankenblatt"
                .ItemData(3) = 4
                .AddItem "Versicherungsetiketten"
                .ItemData(4) = 5
                .AddItem "Karteikartenetiketten"
                .ItemData(5) = 6
                .AddItem "Krankenaktenausdruck"
                .ItemData(6) = 7
                .AddItem "Diagnose-Etiketten"
                .ItemData(7) = 8
                .AddItem "Photoetiketten (Alternativ)"
                .ItemData(8) = 9
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, AdIdx, ByVal 0&)
    Case RibTab_Mitarbeit:
            With CmLis
                .AddItem "Adressenliste"
                .ItemData(0) = 1
                .AddItem "Adressetiketten"
                .ItemData(1) = 2
                .AddItem "Photoetiketten"
                .ItemData(2) = 3
                .AddItem "Patientenkrankenblatt"
                .ItemData(3) = 4
                .AddItem "Versicherungsetiketten"
                .ItemData(4) = 5
                .AddItem "Karteikartenetiketten"
                .ItemData(5) = 6
                .AddItem "Krankenaktenausdruck"
                .ItemData(6) = 7
                .AddItem "Diagnose-Etiketten"
                .ItemData(7) = 8
                .AddItem "Photoetiketten (Alternativ)"
                .ItemData(8) = 9
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, AdIdx, ByVal 0&)
    Case RibTab_Rechnungen:
            Select Case DrTyp
            Case 1:
                With CmLis
                    .AddItem "Adressetiketten"
                    .ItemData(0) = 1
                    .AddItem "Photoetiketten"
                    .ItemData(1) = 2
                    .AddItem "Versicherungsetiketten"
                    .ItemData(2) = 3
                    .AddItem "Karteikartenetiketten"
                    .ItemData(3) = 4
                    .AddItem "Krankenaktenausdruck"
                    .ItemData(4) = 5
                    .AddItem "Diagnose-Etiketten"
                    .ItemData(5) = 6
                    .AddItem "Photoetiketten (Alternativ)"
                    .ItemData(6) = 7
                End With
                RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, EtIdx, ByVal 0&)
            Case 4:
                With CmLis
                    .AddItem "Rechnungsliste"
                    .ItemData(0) = 4
                    .AddItem "Rechnungsexportliste"
                    .ItemData(1) = 5
                    .AddItem "Rechnungsschnellübersicht"
                    .ItemData(2) = 6
                End With
                RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, LsIdx, ByVal 0&)
            End Select
    Case RibTab_Mahnwesen:
            Select Case DrTyp
            Case 1:
                With CmLis
                    .AddItem "Adressetiketten"
                    .ItemData(0) = 1
                    .AddItem "Photoetiketten"
                    .ItemData(1) = 2
                    .AddItem "Versicherungsetiketten"
                    .ItemData(2) = 3
                    .AddItem "Karteikartenetiketten"
                    .ItemData(3) = 4
                    .AddItem "Krankenaktenausdruck"
                    .ItemData(4) = 5
                    .AddItem "Diagnose-Etiketten"
                    .ItemData(5) = 6
                    .AddItem "Photoetiketten (Alternativ)"
                    .ItemData(6) = 7
                End With
                RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, EtIdx, ByVal 0&)
            Case 5:
                With CmLis
                    .AddItem "Offene-Postenliste"
                    .ItemData(0) = 1
                    .AddItem "Gruppierte Postenliste"
                    .ItemData(1) = 2
                End With
            End Select
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, 0, ByVal 0&)
    Case RibTab_HomeBanki:
            With CmLis
                .AddItem "Kontoauszug"
                .ItemData(0) = 1
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, 0, ByVal 0&)
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = False
    Case RibTab_Ter_Kalend:
            With CmLis
                .AddItem "Terminzettel"
                .ItemData(0) = 1
                .AddItem "Terminkalender"
                .ItemData(1) = 2
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, 0, ByVal 0&)
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = False
    Case RibTab_Ter_Raeume:
            With CmLis
                .AddItem "Terminzettel"
                .ItemData(0) = 1
                .AddItem "Terminkalender"
                .ItemData(1) = 2
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, 0, ByVal 0&)
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = False
    Case RibTab_Ter_Mitarb:
            With CmLis
                .AddItem "Terminzettel"
                .ItemData(0) = 1
                .AddItem "Terminkalender"
                .ItemData(1) = 2
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, 0, ByVal 0&)
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = False
    Case RibTab_Ter_Listen:
            If TeIdx = 0 Then TeIdx = 1
            With CmLis
                .AddItem "Einfache Terminliste"
                .ItemData(0) = 1
                .AddItem "Patienten Terminserie"
                .ItemData(1) = 2
                .AddItem "Patienten Terminzettel"
                .ItemData(2) = 3
                .AddItem "Überweisungsträger 1"
                .ItemData(3) = 4
                .AddItem "Überweisungsträger 2"
                .ItemData(4) = 5
                .AddItem "Überweisungsträger 3"
                .ItemData(5) = 6
                .AddItem "Überweisungsträger 4"
                .ItemData(6) = 7
                .AddItem "Quittung Terminserie 1"
                .ItemData(7) = 8
                .AddItem "Quittung Terminserie 2"
                .ItemData(8) = 9
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, TeIdx, ByVal 0&)
            If TeIdx = 3 Then
                OpAd7.Enabled = False
                OpAd8.Enabled = False
                OpAd9.Enabled = False
            Else
                OpAd7.Enabled = True
                OpAd8.Enabled = True
                OpAd9.Enabled = False
            End If
    Case RibTab_Ter_Akont:
            If TeIdx = 0 Then TeIdx = 1
            With CmLis
                .AddItem "Einfache Terminliste"
                .ItemData(0) = 1
                .AddItem "Patienten Terminserie"
                .ItemData(1) = 2
                .AddItem "Patienten Terminzettel"
                .ItemData(2) = 3
                .AddItem "Überweisungsträger 1"
                .ItemData(3) = 4
                .AddItem "Überweisungsträger 2"
                .ItemData(4) = 5
                .AddItem "Überweisungsträger 3"
                .ItemData(5) = 6
                .AddItem "Überweisungsträger 4"
                .ItemData(6) = 7
                .AddItem "Quittung Terminserie 1"
                .ItemData(7) = 8
                .AddItem "Quittung Terminserie 2"
                .ItemData(8) = 9
                .AddItem "Termin Einzelmahnung"
                .ItemData(9) = 10
                .AddItem "Termin Sammelmahnung"
                .ItemData(10) = 11
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, TeIdx, ByVal 0&)
            If TeIdx = 3 Then
                OpAd7.Enabled = False
                OpAd8.Enabled = False
                OpAd9.Enabled = False
            Else
                OpAd7.Enabled = True
                OpAd8.Enabled = True
                OpAd9.Enabled = False
            End If
    Case RibTab_Ter_Warte:
            If TeIdx = 0 Then TeIdx = 1
            With CmLis
                .AddItem "Einfache Terminliste"
                .ItemData(0) = 1
                .AddItem "Patienten Terminserie"
                .ItemData(1) = 2
                .AddItem "Patienten Terminzettel"
                .ItemData(2) = 3
                .AddItem "Überweisungsträger 1"
                .ItemData(3) = 4
                .AddItem "Überweisungsträger 2"
                .ItemData(4) = 5
                .AddItem "Überweisungsträger 3"
                .ItemData(5) = 6
                .AddItem "Überweisungsträger 4"
                .ItemData(6) = 7
                .AddItem "Quittung Terminserie 1"
                .ItemData(7) = 8
                .AddItem "Quittung Terminserie 2"
                .ItemData(8) = 9
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, TeIdx, ByVal 0&)
            If TeIdx = 3 Then
                OpAd7.Enabled = False
                OpAd8.Enabled = False
                OpAd9.Enabled = False
            Else
                OpAd7.Enabled = True
                OpAd8.Enabled = True
                OpAd9.Enabled = False
            End If
    Case RibTab_Rezeptmodul:
            Select Case DrTyp
            Case 6:
                With CmLis
                    .AddItem "inklusive Hintergrundgrafik"
                    .ItemData(0) = 1
                    .AddItem "exklusive Hintergrundgrafik"
                    .ItemData(1) = 2
                End With
                RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, 0, ByVal 0&)
                OpAd7.Enabled = False
                OpAd8.Enabled = False
                OpAd9.Enabled = False
            Case 1:
                With CmLis
                    .AddItem "Adressetiketten"
                    .ItemData(0) = 1
                    .AddItem "Photoetiketten"
                    .ItemData(1) = 2
                    .AddItem "Versicherungsetiketten"
                    .ItemData(2) = 3
                    .AddItem "Karteikartenetiketten"
                    .ItemData(3) = 4
                    .AddItem "Krankenaktenausdruck"
                    .ItemData(4) = 5
                    .AddItem "Diagnose-Etiketten"
                    .ItemData(5) = 6
                    .AddItem "Photoetiketten (Alternativ)"
                    .ItemData(6) = 7
                End With
                RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, EtIdx, ByVal 0&)
            End Select
    Case RibTab_Belegmodul:
            Set CmCom = CmBrs.FindControl(CmCom, SY_RZ_Beleg_Vorlage, , True)
            FoIdx = CmCom.ListIndex
            Select Case FoIdx
            Case 4: TeRa7.Visible = True
            Case 6: TeRa7.Visible = True
            End Select
    
            Select Case DrTyp
            Case 6:
                With CmKto
                    If GlMVo = False Then 'mandantenbezogene Vorgaben verwenden
                        For AktZa = 1 To UBound(GlErK)
                            .AddItem GlErK(AktZa, 1)
                            .ItemData(.NewIndex) = GlErK(AktZa, 0) '[IDK]
                        Next AktZa
                    End If
                End With
                                                                
                With CmGeg
                    If GlBuc = True Then 'einfache Buchhaltung verwenden
                        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                            .AddItem GlGeK(AktZa, 3)
                            .ItemData(AktZa - 1) = GlGeK(AktZa, 0)
                        Next AktZa
                    Else
                        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                                    .AddItem GlSaK(AktKo, 3)
                                    .ItemData(AktZa - 1) = GlSaK(AktKo, 6) '[IDB]
                                End If
                            Next AktKo
                        Next AktZa
                        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
                            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                                .AddItem GlGeK(AktZa, 3)
                                .ItemData(AktZa - 1) = GlGeK(AktZa, 0)
                            Next AktZa
                        End If
                    End If
                End With
                
                With CmStu
                    For AktZa = 1 To UBound(GlSaU) 'Sachkonten mit Steuerkontenzuordnung
                        .AddItem GlSaU(AktZa, 3)
                        .ItemData(AktZa - 1) = GlSaU(AktZa, 6) '[IDI]
                    Next AktZa
                End With

                IdStK = SCmb(CmStu, GlSKo) 'Standardsteuerkonto
                If IdStK >= 0 Then
                    CmStu.ListIndex = IdStK
                Else
                    CmStu.ListIndex = 0
                End If

                If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
                    S_KoMa 3, ManNr
                Else
                    StaKt = SCmb(CmKto, GlSE1) 'Standarderlöskonto Kasse
                    If StaKt >= 0 Then
                        CmKto.ListIndex = StaKt
                    Else
                        CmKto.ListIndex = 0
                    End If
                    If StaGe = 0 Then
                        StaGe = SCmb(CmGeg, GlGkK) 'Standardgeldkonto Kasse
                        If StaGe >= 0 Then
                            If CmGeg.ListCount > 0 Then
                                CmGeg.ListIndex = StaGe
                            End If
                        Else
                            CmGeg.ListIndex = 0
                        End If
                    Else
                        StaGe = SCmb(CmGeg, StaGe)
                        If CmGeg.ListCount > 0 Then
                            CmGeg.ListIndex = StaGe
                        End If
                    End If
                End If
                
                If CmKto.ListIndex < 0 Then
                    CmKto.ListIndex = 0
                End If
                
                If CmGeg.ListIndex < 0 Then
                    CmGeg.ListIndex = 0
                End If

                With CmLis
                    .AddItem "inklusive Hintergrundgrafik"
                    .ItemData(0) = 1
                    .AddItem "exklusive Hintergrundgrafik"
                    .ItemData(1) = 2
                End With
                RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, 0, ByVal 0&)
                OpAd7.Enabled = False
                OpAd8.Enabled = False
                OpAd9.Enabled = False
            Case 1:
                With CmLis
                    .AddItem "Adressetiketten"
                    .ItemData(0) = 1
                    .AddItem "Photoetiketten"
                    .ItemData(1) = 2
                    .AddItem "Versicherungsetiketten"
                    .ItemData(2) = 3
                    .AddItem "Karteikartenetiketten"
                    .ItemData(3) = 4
                    .AddItem "Krankenaktenausdruck"
                    .ItemData(4) = 5
                    .AddItem "Diagnose-Etiketten"
                    .ItemData(5) = 6
                    .AddItem "Photoetiketten (Alternativ)"
                    .ItemData(6) = 7
                End With
                RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, EtIdx, ByVal 0&)
            End Select
    Case Else:
            With CmLis
                .AddItem "Adressetiketten"
                .ItemData(0) = 1
                .AddItem "Photoetiketten"
                .ItemData(1) = 2
                .AddItem "Versicherungsetiketten"
                .ItemData(2) = 3
                .AddItem "Karteikartenetiketten"
                .ItemData(3) = 4
                .AddItem "Krankenaktenausdruck"
                .ItemData(4) = 5
                .AddItem "Diagnose-Etiketten"
                .ItemData(5) = 6
                .AddItem "Adressetiketten (Alternativ)"
                .ItemData(6) = 7
                .AddItem "Photoetiketten (Alternativ)"
                .ItemData(7) = 8
            End With
            RetWe = SendMessage(CmLis.hwnd, CB_SETCURSEL, EtIdx, ByVal 0&)
            OpAd7.Enabled = False
            OpAd8.Enabled = False
            OpAd9.Enabled = True
    End Select
End If

Select Case DrTyp
Case 1:
    Rahm3.Visible = True
    Rahm4.Visible = True
Case 2:
    Rahm1.Visible = True
    If ReAbs = True Then
        OpOp1.Value = True
        ChReD.Value = xtpChecked
        ChReD.Enabled = True
        TxDa1.Enabled = True
        UpDo1.Enabled = True
        PuBu1.Enabled = True
    Else
        OpOp2.Value = True
        ChReD.Enabled = False
        TxDa1.Enabled = False
        UpDo1.Enabled = False
        PuBu1.Enabled = False
    End If
Case 3:
    Rahm2.Visible = True
    ChReD.Enabled = False
    TxDa1.Enabled = False
    PuBu1.Enabled = False
    If MahSt = True Then
        OpOp3.Value = True
    Else
        ChMaG.Enabled = False
    End If
Case 4:
    Rahm3.Visible = True
    Rahm4.Visible = True
Case 5:
    Rahm3.Visible = True
    Rahm4.Visible = True
Case 6:
    Rahm3.Visible = True
    Rahm4.Visible = True
Case 7:
    Rahm3.Visible = True
    Rahm4.Visible = True
End Select

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(ReDat, "dd.mm.yyyy")
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

With CmReF
    .AutoComplete = False
    .DropDownItemCount = 18
End With

With CmLis
    .AutoComplete = False
    .DropDownItemCount = 12
End With

With TxKop
    .Pattern = "\d*"
    .SetMask "0", "_"
    .Text = 2
End With

ChMaR.Value = xtpChecked

If DrFrm = True Then
    ChePr.Value = xtpChecked
End If

Select Case EmVer
Case 0: 'Drucken
    CmEml.Enabled = False
    ChXRe.Enabled = False
Case 1: 'E-Mail-Versand
    CmEml.Enabled = True
    If EmSep = True Then
        If GlEmA = 0 Then 'Emailafdresse Rechnungsversand
            CmEml.ListIndex = 2
        End If
    End If
Case 5:
    CmEml.Enabled = False
Case 6: 'Download-Link
    CmEml.Enabled = True
    If EmSep = True Then
        If GlEmA = 0 Then 'Emailafdresse Rechnungsversand
            CmEml.ListIndex = 2
        End If
    End If
End Select

If EmVer > 0 Then
    CheDu.Enabled = False
    CheSe.Enabled = False
Else
    If GlDub = True Then
        CheDu.Value = xtpChecked
    End If
    If GlDrS = True Then
        CheSe.Value = xtpChecked
    End If
End If

Select Case GlBut
Case RibTab_Abrechnung:
    Select Case EmVer
    Case 1: OpOp1.Value = True 'E-Mail-Versand
    Case 6: OpOp1.Value = True 'downloadlink
    End Select
Case RibTab_Rechnungen:
    Select Case EmVer
    Case 1: OpOp1.Value = True 'E-Mail-Versand
    Case 6: OpOp1.Value = True 'downloadlink
    End Select
Case RibTab_Mahnwesen:
    Select Case EmVer
    Case 1: OpOp3.Value = True 'E-Mail-Versand
    Case 6: OpOp3.Value = True 'downloadlink
    End Select
End Select

If GlGut = True Then 'Guthaben bei Quittungsdruck automatisch erhöhen
    CheGu.Value = xtpChecked
End If

CmStu.Enabled = GlSpB 'Umsatzsteuer Splittbuchung

TxMah.Text = Format$(0, GlWa1)

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
TeRa1.BackColor = GlBak
TeRa2.BackColor = GlBak
TeRa3.BackColor = GlBak
TeRa4.BackColor = GlBak
TeRa5.BackColor = GlBak
TeRa6.BackColor = GlBak
TeRa7.BackColor = GlBak
OpOp1.BackColor = GlBak
OpOp2.BackColor = GlBak
OpOp3.BackColor = GlBak
OpOp4.BackColor = GlBak
ChMaL.BackColor = GlBak
ChMaR.BackColor = GlBak
ChMaG.BackColor = GlBak
OpAd7.BackColor = GlBak
OpAd8.BackColor = GlBak
OpAd9.BackColor = GlBak
ChReD.BackColor = GlBak
ChXRe.BackColor = GlBak
ChePr.BackColor = GlBak
CheDu.BackColor = GlBak
CheSe.BackColor = GlBak
CheGu.BackColor = GlBak
CheGs.BackColor = GlBak
CheRe.BackColor = GlBak
CheDr.BackColor = GlBak

clFen.FenVor

Set clFen = Nothing
Set clDru = Nothing
Set ImMan = Nothing

Set RpCo1 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1
Set TeRa3 = Me.teiRahm3

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = TeRa3.Top + TxDa1.Top + TxDa1.Height
    .Left = TxDa1.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TxDa1.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKont()
On Error GoTo InErr

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set TxDa1 = FM.txtDatu1
Set OpOp1 = Me.optOpti1
Set OpOp2 = Me.optOpti2
Set OpOp3 = Me.optOpti3
Set OpOp4 = Me.optOpti4
Set ChMaL = Me.chkMaLis
Set ChReD = Me.chkReDat
Set ChMaR = Me.chkMeRec
Set ChMaG = Me.chkMaGeb
Set ChXRe = Me.chkXRech
Set TxMah = Me.txtMahng
Set UpDo1 = Me.updCont1
Set PuBu1 = Me.btnDatu1

If Rahm1.Visible = True Then
    If OpOp1.Value = True Then
        ChReD.Value = xtpChecked
        ChReD.Enabled = True
        TxDa1.Enabled = True
        UpDo1.Enabled = True
        PuBu1.Enabled = True
        IniSetVal "System", "RechAb", -1
    Else
        ChReD.Value = xtpUnchecked
        ChReD.Enabled = False
        TxDa1.Enabled = False
        UpDo1.Enabled = False
        PuBu1.Enabled = False
        IniSetVal "System", "RechAb", 0
    End If
ElseIf Rahm2.Visible = True Then
    If OpOp3.Value = True Then
        ChMaG.Enabled = True
        TxMah.Enabled = True
        If ChMaG.Value = xtpChecked Then
            TxMah.Text = Format$(5, GlWa1)
        Else
            TxMah.Enabled = False
            TxMah.Text = Format$(0, GlWa1)
        End If
        ChMaR.Value = xtpChecked
        IniSetVal "System", "MahnSt", -1
    ElseIf OpOp4.Value = True Then
        ChMaG.Enabled = False
        TxMah.Enabled = False
        TxMah.Text = Format$(0, GlWa1)
        ChMaG.Value = xtpUnchecked
        ChMaR.Value = xtpUnchecked
        IniSetVal "System", "MahnSt", 0
    End If
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKont " & Err.Number
Resume Next

End Sub
Private Sub FoDat()
On Error GoTo OrErr

Dim RetWe As Long

Set CmReF = Me.cmbReFor
Set ChePr = Me.chkRePru

CmReF.Clear
DoEvents

If ChePr.Value = xtpChecked Then
    DrFrm = True
Else
    DrFrm = False
End If

IniSetVal "System", "DruFrm", DrFrm

With CmReF
    If DrFrm = True Then
        .AddItem GlFrm(1, 0)
        .ItemData(0) = 1
        .AddItem GlFrm(1, 1)
        .ItemData(1) = 2
        .AddItem GlFrm(1, 2)
        .ItemData(2) = 3
        .AddItem GlFrm(1, 3)
        .ItemData(3) = 4
        .AddItem GlFrm(1, 4)
        .ItemData(4) = 5
        .AddItem GlFrm(1, 5)
        .ItemData(5) = 6
        .AddItem GlFrm(1, 14)
        .ItemData(6) = 7
        .AddItem GlFrm(1, 9)
        .ItemData(7) = 8
        .AddItem GlFrm(1, 15)
        .ItemData(8) = 9
        .AddItem GlFrm(1, 6)
        .ItemData(9) = 10
        .AddItem GlFrm(1, 7)
        .ItemData(10) = 11
        .AddItem GlFrm(1, 16)
        .ItemData(11) = 12
        .AddItem GlFrm(1, 13)
        .ItemData(12) = 13
        .AddItem GlFrm(1, 17)
        .ItemData(13) = 14
        .AddItem GlFrm(1, 19)
        .ItemData(14) = 15
        .AddItem GlFrm(1, 78)
        .ItemData(15) = 16
        .AddItem GlFrm(1, 18)
        .ItemData(16) = 17
        .AddItem GlFrm(1, 96)
        .ItemData(17) = 18
    Else
        .AddItem GlFrm(0, 0)
        .ItemData(0) = 1
        .AddItem GlFrm(0, 1)
        .ItemData(1) = 2
        .AddItem GlFrm(0, 2)
        .ItemData(2) = 3
        .AddItem GlFrm(0, 3)
        .ItemData(3) = 4
        .AddItem GlFrm(0, 4)
        .ItemData(4) = 5
        .AddItem GlFrm(0, 5)
        .ItemData(5) = 6
        .AddItem GlFrm(0, 14)
        .ItemData(6) = 7
        .AddItem GlFrm(0, 9)
        .ItemData(7) = 8
        .AddItem GlFrm(0, 15)
        .ItemData(8) = 9
        .AddItem GlFrm(0, 6)
        .ItemData(9) = 10
        .AddItem GlFrm(0, 7)
        .ItemData(10) = 11
        .AddItem GlFrm(0, 16)
        .ItemData(11) = 12
        .AddItem GlFrm(0, 13)
        .ItemData(12) = 13
        .AddItem GlFrm(0, 17)
        .ItemData(13) = 14
        .AddItem GlFrm(0, 19)
        .ItemData(14) = 15
        .AddItem GlFrm(0, 78)
        .ItemData(15) = 16
        .AddItem GlFrm(0, 18)
        .ItemData(16) = 17
        .AddItem GlFrm(0, 96)
        .ItemData(17) = 18
    End If
End With

If ReTyp <> vbNullString Then
    Select Case ReTyp
    Case "L": RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, 2, ByVal 0&)
    Case "V": RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, 7, ByVal 0&)
    Case "U": RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, 8, ByVal 0&)
    Case Else: RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, ReIdx, ByVal 0&)
    End Select
Else
    RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, ReIdx, ByVal 0&)
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FoDat " & Err.Number
Resume Next

End Sub
Private Sub FOpti()
On Error Resume Next

Set UpDo1 = Me.updCont2
Set TxKop = Me.txtKopie
Set OpAd9 = Me.optAdre9
    
If OpAd9.Value = True Then
    TxKop.Visible = True
    UpDo1.Visible = True
Else
    TxKop.Visible = False
    UpDo1.Visible = False
End If

End Sub
Private Sub FWeit()
On Error GoTo InErr

Dim ReDat As Date
Dim NeuDa As Date
Dim ErKId As Long
Dim StKID As Long
Dim ErKNr As Long
Dim GeKNr As Long
Dim StKNr As Long
Dim ForNa As String
Dim ErKBe As String
Dim GeKBe As String
Dim StKBe As String
Dim MaGeb As Single
Dim NeGeb As Single
Dim ReAbs As Boolean
Dim DrLis As Boolean
Dim DaAnp As Boolean
Dim KeiSe As Boolean
Dim Gesam As Boolean
Dim DrSof As Boolean
Dim EiDru As Boolean
Dim Gutha As Boolean
Dim Gutsh As Boolean
Dim ReAus As Boolean
Dim DrDia As Boolean
Dim XRech As Boolean
Dim GeKId As Integer
Dim LiIdx As Integer
Dim IdxWe As Integer
Dim EmSen As Integer
Dim AnzKo As Integer

Set FM = frmDruck
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set TeRa7 = FM.teiRahm7
Set TxDa1 = FM.txtDatu1
Set TxKop = FM.txtKopie
Set CmReF = FM.cmbReFor
Set CmMaF = FM.cmbMaFor
Set OpOp1 = FM.optOpti1
Set OpOp2 = FM.optOpti2
Set OpOp3 = FM.optOpti3
Set OpOp4 = FM.optOpti4
Set OpAd7 = FM.optAdre7
Set OpAd8 = FM.optAdre8
Set OpAd9 = FM.optAdre9
Set ChMaL = FM.chkMaLis
Set ChReD = FM.chkReDat
Set ChXRe = FM.chkXRech
Set ChMaR = FM.chkMeRec
Set ChMaG = FM.chkMaGeb
Set TxMah = FM.txtMahng
Set CmLis = FM.cmbLiDru
Set CmEml = FM.cmbEmail
Set CmKto = FM.cmbKonto
Set CmGeg = FM.cmbGegen
Set ChePr = FM.chkRePru
Set CheGu = FM.chkGutha
Set CheGs = FM.chkGutsh
Set CheRe = FM.chkRechn
Set CheDr = Me.chkDrDia

DrSof = False

GlDru = GlDrX

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

EmSen = CmEml.ListIndex + 1

Select Case EmSen
Case 1: EmSep = False 'an einen Patienten
Case 2: EmSep = False 'an den Mandanten
Case 3: EmSep = True 'an alle Patienten
End Select

If CheDr.Value = xtpChecked Then
    DrDia = True
End If

If Rahm4.Visible = True Then

    If WindowLoad("frmAdress") = True Then
        EiDru = True
        LiIdx = CmLis.ListIndex + 1
        Select Case LiIdx
        Case 1: ForNa = "AdrEti"
        Case 2: ForNa = "PhoEti"
        Case 3: ForNa = "VerEti"
        Case 4: ForNa = "KarEt1"
        Case 5: ForNa = "KarEt2"
        Case 6: ForNa = "DiaEti"
        Case 7: ForNa = "StaEt2"
        End Select
    Else
        Select Case GlBut
        Case RibTab_Adressen:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "AdrLis"
            Case 2: ForNa = "AdrEti"
            Case 3: ForNa = "PhoEti"
            Case 4: ForNa = "KraBla"
            Case 5: ForNa = "VerEti"
            Case 6: ForNa = "KarEt1"
            Case 7: ForNa = "KarEt2"
            Case 8: ForNa = "DiaEti"
            Case 9: ForNa = "StaEt2"
            End Select
        Case RibTab_Mandanten:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "AdrLis"
            Case 2: ForNa = "AdrEti"
            Case 3: ForNa = "PhoEti"
            Case 4: ForNa = "KraBla"
            Case 5: ForNa = "VerEti"
            Case 6: ForNa = "KarEt1"
            Case 7: ForNa = "KarEt2"
            Case 8: ForNa = "DiaEti"
            Case 9: ForNa = "StaEt2"
            End Select
        Case RibTab_Verordner:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "AdrLis"
            Case 2: ForNa = "AdrEti"
            Case 3: ForNa = "PhoEti"
            Case 4: ForNa = "KraBla"
            Case 5: ForNa = "VerEti"
            Case 6: ForNa = "KarEt1"
            Case 7: ForNa = "KarEt2"
            Case 8: ForNa = "DiaEti"
            Case 9: ForNa = "StaEt2"
            End Select
        Case RibTab_Mitarbeit:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "AdrLis"
            Case 2: ForNa = "AdrEti"
            Case 3: ForNa = "PhoEti"
            Case 4: ForNa = "KraBla"
            Case 5: ForNa = "VerEti"
            Case 6: ForNa = "KarEt1"
            Case 7: ForNa = "KarEt2"
            Case 8: ForNa = "DiaEti"
            Case 9: ForNa = "StaEt2"
            End Select
        Case RibTab_Vorbereit:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "AdrLis"
            Case 2: ForNa = "AdrEti"
            Case 3: ForNa = "PhoEti"
            Case 4: ForNa = "KraBla"
            Case 5: ForNa = "VerEti"
            Case 6: ForNa = "KarEt1"
            Case 7: ForNa = "KarEt2"
            Case 8: ForNa = "DiaEti"
            Case 9: ForNa = "StaEt2"
            End Select
        Case RibTab_Rechnungen:
            Select Case DrTyp
            Case 1:
                EiDru = True
                LiIdx = CmLis.ListIndex + 1
                Select Case LiIdx
                Case 1: ForNa = "AdrEti"
                Case 2: ForNa = "PhoEti"
                Case 3: ForNa = "VerEti"
                Case 4: ForNa = "KarEt1"
                Case 5: ForNa = "KarEt2"
                Case 6: ForNa = "DiaEti"
                Case 7: ForNa = "StaEt2"
                End Select
            Case 4:
                LiIdx = CmLis.ListIndex + 1
                Select Case CmLis.ListIndex
                Case 0: ForNa = "RechLi"
                Case 1: ForNa = "ReList"
                Case 2: ForNa = "ResUbe"
                End Select
            End Select
        Case RibTab_Mahnwesen:
            Select Case DrTyp
            Case 1:
                EiDru = True
                LiIdx = CmLis.ListIndex + 1
                Select Case LiIdx
                Case 1: ForNa = "AdrEti"
                Case 2: ForNa = "PhoEti"
                Case 3: ForNa = "VerEti"
                Case 4: ForNa = "KarEt1"
                Case 5: ForNa = "KarEt2"
                Case 6: ForNa = "DiaEti"
                Case 7: ForNa = "StaEt2"
                End Select
            Case 5:
                Select Case CmLis.ListIndex
                Case 0: ForNa = "PostLi"
                Case 1: ForNa = "PostGr"
                End Select
            End Select
        Case RibTab_HomeBanki:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "StaEti"
            End Select
        Case RibTab_Ter_Kalend:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = vbNullString
                    Unload Me
                    Select Case EmVer
                    Case 0: STeDr "TerPat"
                    Case 1: STeDr "TerPat", , , , EmSen
                    Case 5: STeDr "TerPat", , , , EmVer
                    End Select
            Case 2: ForNa = vbNullString
                    Unload Me
                    STePr
            End Select
        Case RibTab_Ter_Raeume:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = vbNullString
                    Unload Me
                    Select Case EmVer
                    Case 0: STeDr "TerPat"
                    Case 1: STeDr "TerPat", , , , EmSen
                    Case 5: STeDr "TerPat", , , , EmVer
                    End Select
            Case 2: ForNa = vbNullString
                    Unload Me
                    STePr
            End Select
        Case RibTab_Ter_Mitarb:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = vbNullString
                    Unload Me
                    Select Case EmVer
                    Case 0: STeDr "TerPat"
                    Case 1: STeDr "TerPat", , , , EmSen
                    Case 5: STeDr "TerPat", , , , EmVer
                    End Select
            Case 2: ForNa = vbNullString
                    Unload Me
                    STePr
            End Select
        Case RibTab_Ter_Listen:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "TerLis"
            Case 2: ForNa = vbNullString
                    Unload Me
                    STeDr "TerSer"
            Case 3: ForNa = vbNullString
                    Unload Me
                    STeDr "TerPat"
            Case 4: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbNo"
            Case 5: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbEu"
            Case 6: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbAN"
            Case 7: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbAE"
            Case 8: ForNa = vbNullString
                    Unload Me
                    STeDr "QuiTer"
            Case 9: ForNa = vbNullString
                    Unload Me
                    STeDr "QuiTem"
            Case 10: ForNa = vbNullString
                    Unload Me
                    STeDr "TeEiMa"
            Case 11: ForNa = vbNullString
                    Unload Me
                    STeDr "TeEiSa"
            End Select
        Case RibTab_Ter_Akont:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "TerLis"
            Case 2: ForNa = vbNullString
                    Unload Me
                    STeDr "TerSer"
            Case 3: ForNa = vbNullString
                    Unload Me
                    STeDr "TerPat"
            Case 4: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbNo"
            Case 5: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbEu"
            Case 6: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbAN"
            Case 7: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbAE"
            Case 8: ForNa = vbNullString
                    Unload Me
                    STeDr "QuiTer"
            Case 9: ForNa = vbNullString
                    Unload Me
                    STeDr "QuiTem"
            Case 10: ForNa = vbNullString
                    Unload Me
                    STeDr "TeEiMa"
            Case 11: ForNa = vbNullString
                    Unload Me
                    STeDr "TeEiSa"
            End Select
        Case RibTab_Ter_Warte:
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "TerLis"
            Case 2: ForNa = vbNullString
                    Unload Me
                    STeDr "TerSer"
            Case 3: ForNa = vbNullString
                    Unload Me
                    STeDr "TerPat"
            Case 4: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbNo"
            Case 5: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbEu"
            Case 6: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbAN"
            Case 7: ForNa = vbNullString
                    Unload Me
                    STeDr "TeUbAE"
            Case 8: ForNa = vbNullString
                    Unload Me
                    STeDr "QuiTer"
            Case 9: ForNa = vbNullString
                    Unload Me
                    STeDr "QuiTem"
            Case 10: ForNa = vbNullString
                    Unload Me
                    STeDr "TeEiMa"
            Case 11: ForNa = vbNullString
                    Unload Me
                    STeDr "TeEiSa"
            End Select
        Case RibTab_Rezeptmodul:
            Select Case DrTyp
            Case 6:
                LiIdx = CmLis.ListIndex + 1
                Select Case LiIdx
                Case 1: Unload Me
                        Select Case EmVer
                        Case 0: SRzDr False, 1, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh
                        Case 1: SRzDr True, 1, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmSen
                        Case 6: SRzDr True, 1, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmVer
                        End Select
                Case 2: Unload Me
                        Select Case EmVer
                        Case 0: SRzDr False, 2, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh
                        Case 1: SRzDr True, 2, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmSen
                        Case 6: SRzDr True, 2, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmVer
                        End Select
                End Select
            Case 1:
                EiDru = True
                LiIdx = CmLis.ListIndex + 1
                Select Case LiIdx
                Case 1: ForNa = "AdrEti"
                Case 2: ForNa = "PhoEti"
                Case 3: ForNa = "VerEti"
                Case 4: ForNa = "KarEt1"
                Case 5: ForNa = "KarEt2"
                Case 6: ForNa = "DiaEti"
                Case 7: ForNa = "StaEt2"
                End Select
            End Select
        Case RibTab_Belegmodul:
            Select Case DrTyp
            Case 6:
                If CheGu.Value = xtpChecked Then
                    Gutha = True
                Else
                    Gutha = False
                End If
                If CheGs.Value = xtpChecked Then
                    Gutsh = True
                Else
                    Gutsh = False
                End If
                If TeRa7.Visible = True Then
                    ErKId = CmKto.ItemData(CmKto.ListIndex)
                    GeKId = CmGeg.ItemData(CmGeg.ListIndex)
                    StKID = CmStu.ItemData(CmStu.ListIndex)
                    If GlKnF = True Then 'Sachkontenformatierung sechsstellig
                        ErKNr = Left$(CmKto.Text, 6)
                        GeKNr = Left$(CmGeg.Text, 6)
                        StKNr = Left$(CmStu.Text, 6)
                        ErKBe = Mid$(CmKto.Text, 8, Len(CmKto.Text) - 7)
                        GeKBe = Mid$(CmGeg.Text, 8, Len(CmGeg.Text) - 7)
                        StKBe = Mid$(CmStu.Text, 8, Len(CmStu.Text) - 7)
                    Else
                        ErKNr = Left$(CmKto.Text, 4)
                        GeKNr = Left$(CmGeg.Text, 4)
                        StKNr = Left$(CmStu.Text, 4)
                        ErKBe = Mid$(CmKto.Text, 6, Len(CmKto.Text) - 5)
                        GeKBe = Mid$(CmGeg.Text, 6, Len(CmGeg.Text) - 5)
                        StKBe = Mid$(CmStu.Text, 6, Len(CmStu.Text) - 5)
                    End If
                Else
                    ErKId = 0
                    GeKId = 0
                    StKID = 0
                    ErKNr = 0
                    GeKNr = 0
                    StKNr = 0
                End If
                LiIdx = CmLis.ListIndex + 1
                Select Case LiIdx
                Case 1: Unload Me
                        Select Case EmVer
                        Case 0: SRzDr False, LiIdx, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh
                        Case 1: SRzDr True, LiIdx, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmSen
                        Case 6: SRzDr True, LiIdx, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmVer
                        End Select
                Case 2: Unload Me
                        Select Case EmVer
                        Case 0: SRzDr False, LiIdx, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh
                        Case 1: SRzDr True, LiIdx, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmSen
                        Case 6: SRzDr True, LiIdx, ErKId, ErKNr, ErKBe, GeKId, GeKNr, GeKBe, StKID, StKNr, StKBe, Gutha, Gutsh, EmVer
                        End Select
                End Select
            Case 1:
                EiDru = True
                LiIdx = CmLis.ListIndex + 1
                Select Case LiIdx
                Case 1: ForNa = "AdrEti"
                Case 2: ForNa = "PhoEti"
                Case 3: ForNa = "VerEti"
                Case 4: ForNa = "KarEt1"
                Case 5: ForNa = "KarEt2"
                Case 6: ForNa = "DiaEti"
                Case 7: ForNa = "StaEt2"
                End Select
            End Select
        Case Else:
            EiDru = True
            LiIdx = CmLis.ListIndex + 1
            Select Case LiIdx
            Case 1: ForNa = "AdrEti"
            Case 2: ForNa = "PhoEti"
            Case 3: ForNa = "VerEti"
            Case 4: ForNa = "KarEt1"
            Case 5: ForNa = "KarEt2"
            Case 6: ForNa = "DiaEti"
            Case 7: ForNa = "StaEt2"
            End Select
        End Select
    End If

    If ForNa <> vbNullString Then
        If OpAd8.Value = True Then
            Gesam = True
        End If
        If OpAd9.Value = True Then
            If IsNumeric(TxKop.Text) Then
                AnzKo = TxKop.Text
            Else
                AnzKo = 1
            End If
            With GlDru
                .ForNa = ForNa
                .Wiede = True
                .WieAn = AnzKo
                If EmVer = 1 Then
                    .EmVer = EmSen
                Else
                    .EmVer = EmVer
                End If
            End With
        End If
        Unload Me
        DoEvents
        Select Case ForNa
        Case "ReList":
                Select Case EmVer
                Case 0: S_BeDat ForNa, False, GlDrV, , Gesam
                Case 1: S_BeDat ForNa, False, False, EmSen, Gesam
                Case 5: S_BeDat ForNa, False, False, EmVer, Gesam
                End Select
        Case "ResUbe":
                Select Case EmVer
                Case 0: S_BeDat ForNa, False, GlDrV, , Gesam
                Case 1: S_BeDat ForNa, False, False, EmSen, Gesam
                Case 5: S_BeDat ForNa, False, False, EmVer, Gesam
                End Select
        Case Else:
            If EiDru = True Then
                If WindowLoad("frmAdress") = True Then Unload frmAdress
                FoDru ForNa
            Else
                If ForNa = "KraBla" Then
                    SKrDr
                Else
                    Select Case EmVer
                    Case 0: SDruck ForNa, GlDrV, , Gesam
                    Case 1: SDruck ForNa, GlDrV, , Gesam, , , EmSen
                    Case 5: SDruck ForNa, GlDrV, , Gesam, , , EmVer
                    End Select
                End If
            End If
        End Select
    End If
    
ElseIf Rahm1.Visible = True Then 'Rechnungen
                
    LiIdx = CmReF.ListIndex
    IdxWe = CmReF.ItemData(LiIdx)
    ForNa = GlFrm(2, IdxWe)

    If ChXRe.Value = xtpChecked Then XRech = True
    If OpOp1.Value = True Then ReAbs = True
    If ChReD.Value = 1 Then DaAnp = True
    If ChReD.Value = 1 Then ReDat = NeuDa

    Unload Me
    DoEvents
    With GlDru
        .ForNa = ForNa
        .DruVo = GlDrV
        .DrSof = DrSof
        .ReAbs = ReAbs
        .DrLis = DrLis
        .DaAnp = DaAnp
        .ReDat = ReDat
        .DrDia = DrDia
        .EmSep = EmSep
        .XRech = XRech
        .GoBDk = True
        If EmVer = 1 Then
            .EmVer = EmSen
        Else
            .EmVer = EmVer
        End If
    End With
    SDrDia

ElseIf Rahm2.Visible = True Then 'Mahnungen

    GlDrV = True 'Druckvorschau
    LiIdx = CmMaF.ListIndex + 1
    If OpOp3.Value = True Then ReAbs = True
    If ChMaL.Value = xtpChecked Then DrLis = True
    If ChMaR.Value = xtpChecked Then KeiSe = True
    If CheRe.Value = xtpChecked Then ReAus = True
    
    Select Case LiIdx
    Case 1: ForNa = "EiMahn"
    Case 2: ForNa = "SaMahn"
    Case 3: ForNa = "StMahn"
    End Select

    If TxMah.Text <> vbNullString Then
        If IsNumeric(TxMah.Text) = True Then
            MaGeb = CSng(TxMah.Text)
        End If
    End If
    If AlGeb > 0 Then
        MaGeb = MaGeb + AlGeb
    End If

    Unload Me
    DoEvents
    With GlDru
        .ForNa = ForNa
        .DruVo = GlDrV
        .DrSof = DrSof
        .ReAbs = ReAbs
        .DrLis = DrLis
        .KeiSe = KeiSe
        .MaGeb = MaGeb
        .DrDia = False
        .EmSep = EmSep
        .MaReD = ReAus
        If EmVer = 1 Then
            .EmVer = EmSen
        Else
            .EmVer = EmVer
        End If
    End With
    SDrDia
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub FoDru(ByVal ForNa As String)
On Error GoTo DrErr
'Druckt den aktuellen Eintrag

Dim FiNam As String
Dim LoNam As String
Dim Formu As Boolean

Set FM = frmDruck
Set clLis = New clsLisLab
Set clFil = New clsFile

If GlFrn <> vbNullString Then 'Formulareordner
    FiNam = GlFrn & S_FoCh(ForNa)
Else
    FiNam = GlFrO & S_FoCh(ForNa)
End If

If clFil.FilVor(FiNam) = True Then
    Formu = True
Else
    Formu = False
    SMeFr GlMeT, GlMeM, GlMeI, GlMeF, False, 1, True, FM.hwnd
End If

If Formu = True Then
    Select Case GlBut
    Case RibTab_Adressen: S_AdMa False, False
    Case RibTab_Mandanten: S_AdMa False, False
    Case RibTab_Verordner: S_AdMa False, False
    Case RibTab_Mitarbeit: S_AdMa False, False
    Case RibTab_Rechnungen: S_AdMa False, False
    Case RibTab_Mahnwesen: S_AdMa False, False
    Case Else: S_AdMa False, False, GlAdr
    End Select

    With clLis
        .ForNam = ForNa
        .FilNam = FiNam
        .PfaTmp = GlTmp
        .Gesamt = False
        If ForNa = "StaEt2" Then
            .DruVor = False
            .DruDia = False
            .Wieder = False
            .WieAnz = 1
        Else
            .DruVor = GlDrV
            .DruDia = True
            .Wieder = GlDru.Wiede
            .WieAnz = GlDru.WieAn
        End If
        .MandVo = True
        .MitaVo = GlMiV
        .ArztVo = GlArV
        .LLPrLa
    End With
    DoEvents

    S_AdMa False, True
End If

Set clLis = Nothing
Set clFil = Nothing

Exit Sub

DrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FoDru " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    If FrLoa = False Then
        FKale
    End If
End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5

If Rahm4.Visible = True Then
    Select Case GlBut
    Case RibTab_Rezeptmodul:
        TeTit = IniGetOpt("Hilfe", 50171)
        TeMai = IniGetOpt("Hilfe", 50172)
        TeInh = IniGetOpt("Hilfe", 50173)
        TeFus = IniGetOpt("Hilfe", 50174)
    Case RibTab_Belegmodul:
        TeTit = IniGetOpt("Hilfe", 50181)
        TeMai = IniGetOpt("Hilfe", 50182)
        TeInh = IniGetOpt("Hilfe", 50183)
        TeFus = IniGetOpt("Hilfe", 50184)
    Case Else:
        TeTit = IniGetOpt("Hilfe", 50191)
        TeMai = IniGetOpt("Hilfe", 50192)
        TeInh = IniGetOpt("Hilfe", 50193)
        TeFus = IniGetOpt("Hilfe", 50194)
    End Select
ElseIf Rahm1.Visible = True Then
    Select Case GlBut
    Case RibTab_Abrechnung:
        TeTit = IniGetOpt("Hilfe", 50201)
        TeMai = IniGetOpt("Hilfe", 50202)
        TeInh = IniGetOpt("Hilfe", 50203)
        TeFus = IniGetOpt("Hilfe", 50204)
    Case RibTab_Rechnungen:
        TeTit = IniGetOpt("Hilfe", 50211)
        TeMai = IniGetOpt("Hilfe", 50212)
        TeInh = IniGetOpt("Hilfe", 50213)
        TeFus = IniGetOpt("Hilfe", 50214)
    End Select
ElseIf Rahm2.Visible = True Then
    TeTit = IniGetOpt("Hilfe", 50221)
    TeMai = IniGetOpt("Hilfe", 50222)
    TeInh = IniGetOpt("Hilfe", 50223)
    TeFus = IniGetOpt("Hilfe", 50224)
End If

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub

Private Sub chkDruSe_Click()
On Error Resume Next

Set CheSe = Me.chkDruSe

If FrLoa = False Then
    GlDrS = Not GlDrS
    
    If GlDrS = True Then
        CheSe.Value = xtpChecked
    Else
        CheSe.Value = xtpUnchecked
    End If
End If

End Sub

Private Sub chkDuble_Click()
On Error Resume Next

Set CheDu = Me.chkDuble

If FrLoa = False Then
    GlDub = Not GlDub
    
    If GlDub = True Then
        CheDu.Value = xtpChecked
    Else
        CheDu.Value = xtpUnchecked
    End If
End If

End Sub


Private Sub chkGutha_Click()
On Error Resume Next

Set CheGu = Me.chkGutha
Set CheGs = Me.chkGutsh

If CheGu.Value = xtpChecked Then
    If CheGs.Value = xtpChecked Then
        CheGs.Value = xtpUnchecked
    End If
    GlGut = True
    IniSetVal "System", "GuthEr", -1
Else
    GlGut = False
    IniSetVal "System", "GuthEr", 0
End If

End Sub

Private Sub chkGutsh_Click()
On Error Resume Next

Set CheGu = Me.chkGutha
Set CheGs = Me.chkGutsh

If CheGs.Value = xtpChecked Then
    If CheGu.Value = xtpChecked Then
        CheGu.Value = xtpUnchecked
    End If
End If

End Sub
Private Sub chkMaGeb_Click()
    If FrLoa = False Then
        FKont
    End If
End Sub

Private Sub chkReDat_Click()
On Error Resume Next

Set FM = frmDruck
Set ChReD = FM.chkReDat
Set TxDa1 = FM.txtDatu1
Set UpDo1 = FM.updCont1
Set PuBu1 = FM.btnDatu1

If FrLoa = False Then
    If ChReD.Value = xtpChecked Then
        TxDa1.Enabled = True
        UpDo1.Enabled = True
        PuBu1.Enabled = True
    Else
        TxDa1.Enabled = False
        UpDo1.Enabled = False
        PuBu1.Enabled = False
    End If
End If

End Sub

Private Sub chkRePru_Click()
    If FrLoa = False Then
        FoDat
    End If
End Sub

Private Sub chkXRech_Click()
On Error Resume Next

Dim RetWe As Long

Set ChXRe = Me.chkXRech
Set CmEml = Me.cmbEmail
Set CmReF = Me.cmbReFor

 RetWe = SendMessage(CmReF.hwnd, CB_SETCURSEL, 0, ByVal 0&)

If GlVar <> "PS3" Then
    If GlRDP = False Then
        ChXRe.Value = xtpUnchecked
        SPopu "E-Rechnung Export", "Die E-Rechnung Schnittstelle wurde noch nicht freigeschaltet bzw. lizenziert!", IC48_Forbidden
    End If
End If

End Sub

Private Sub cmbEmail_Click()
On Error Resume Next

Set ChXRe = Me.chkXRech
Set CmEml = Me.cmbEmail

End Sub
Private Sub cmbLiDru_Click()
    If FrLoa = False Then
        FFoSe True
    End If
End Sub
Private Sub cmbMaFor_Click()
    If FrLoa = False Then
        FFoSe
    End If
End Sub

Private Sub cmbReFor_Click()
    If FrLoa = False Then
        FFoSe
    End If
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub

Private Sub dtpDatu1_SelectionChanged()
    If FrLoa = False Then
        FDatu
    End If
End Sub

Private Sub optAdre7_Click()
    FOpti
End Sub
Private Sub optAdre8_Click()
    FOpti
End Sub
Private Sub optAdre9_Click()
    FOpti
End Sub
Private Sub Form_Load()
On Error Resume Next

FrLoa = True

FInit
AFont Me

FrLoa = False

SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmDruck = Nothing
End Sub
Private Sub optOpti1_Click()
    If FrLoa = False Then
        FKont
    End If
End Sub
Private Sub optOpti2_Click()
    If FrLoa = False Then
        FKont
    End If
End Sub
Private Sub optOpti3_Click()
    If FrLoa = False Then
        FKont
    End If
End Sub
Private Sub optOpti4_Click()
    If FrLoa = False Then
        FKont
    End If
End Sub
Private Sub txtDatu1_LostFocus()
    If FrLoa = False Then
        FDaKo
    End If
End Sub

Private Sub txtMahng_GotFocus()
    Me.txtMahng.SelStart = 0
    Me.txtMahng.SelLength = Len(Me.txtMahng.Text)
End Sub

Private Sub txtMahng_LostFocus()
On Error Resume Next

Dim Betra As Double

If Me.txtMahng.Text <> vbNullString Then
    If IsNumeric(Me.txtMahng.Text) = True Then
        Betra = CDbl(Me.txtMahng.Text)
        If Betra < 0 Then
            Betra = Betra * (-1)
        End If
        Me.txtMahng.Text = Format$(Betra, GlWa1)
    End If
End If

End Sub
Private Sub updCont1_DownClick()
On Error Resume Next

Dim AltDa As Date
Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text
NeuDa = DateAdd("d", -1, AltDa)

TxDa1.Text = FDaPr(NeuDa)

End Sub
Private Sub updCont1_UpClick()
On Error Resume Next

Dim AltDa As Date
Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text
NeuDa = DateAdd("d", 1, AltDa)

TxDa1.Text = FDaPr(NeuDa)

End Sub
