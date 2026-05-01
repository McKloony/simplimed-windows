VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Begin VB.Form frmAdrFilt 
   Caption         =   "Adressenfilter"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4035
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   7108
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   340
         Left            =   4710
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1490
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu2"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   340
         Left            =   4710
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1130
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu1"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.CheckBox chkFilt13 
         Height          =   220
         Left            =   900
         TabIndex        =   9
         Top             =   1515
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressen, die nach dem"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtWoOrt 
         Height          =   315
         Left            =   3400
         TabIndex        =   21
         Top             =   3300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtTele2 
         Height          =   315
         Left            =   3400
         TabIndex        =   19
         Top             =   2940
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtWoPLZ 
         Height          =   315
         Left            =   3400
         TabIndex        =   17
         Top             =   2580
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtBemer 
         Height          =   315
         Left            =   3400
         TabIndex        =   15
         Top             =   2230
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.CheckBox chkFilt10 
         Height          =   220
         Left            =   900
         TabIndex        =   20
         Top             =   3315
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressen mit dem Ort"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt9 
         Height          =   220
         Left            =   900
         TabIndex        =   18
         Top             =   2955
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressen, deren Telefon mit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt4 
         Height          =   220
         Left            =   900
         TabIndex        =   16
         Top             =   2595
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressen, deren PLZ mit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt3 
         Height          =   220
         Left            =   900
         TabIndex        =   14
         Top             =   2235
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressen, die als Bemerkung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt2 
         Height          =   220
         Left            =   900
         TabIndex        =   12
         Top             =   1875
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressen, denen der Katalog"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt1 
         Height          =   220
         Left            =   900
         TabIndex        =   6
         Top             =   1155
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressen, die vor dem"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbKata1 
         Height          =   315
         Left            =   3405
         TabIndex        =   13
         Top             =   1860
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   315
         Left            =   3400
         TabIndex        =   10
         Top             =   1500
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   315
         Left            =   3400
         TabIndex        =   7
         Top             =   1140
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin VB.Label lblLab8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrFilt.frx":0000
         Height          =   615
         Left            =   900
         TabIndex        =   70
         Top             =   300
         Width           =   6000
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "zugeordnet ist."
         Height          =   220
         Left            =   5600
         TabIndex        =   69
         Top             =   1875
         Width           =   1800
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "geändert wurden."
         Height          =   220
         Left            =   5100
         TabIndex        =   68
         Top             =   1185
         Width           =   1800
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "beginnt."
         Height          =   220
         Left            =   5600
         TabIndex        =   67
         Top             =   2595
         Width           =   1800
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "aufgeführt haben."
         Height          =   220
         Left            =   5600
         TabIndex        =   66
         Top             =   2235
         Width           =   1800
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "beginnt."
         Height          =   220
         Left            =   5600
         TabIndex        =   65
         Top             =   2955
         Width           =   1800
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   220
         Left            =   5100
         TabIndex        =   64
         Top             =   1515
         Width           =   1455
         _Version        =   1048579
         _ExtentX        =   2566
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "geändert wurden."
         Transparent     =   -1  'True
      End
   End
   Begin VB.TextBox txtStart 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   120
      TabIndex        =   0
      Top             =   9000
      Width           =   80
   End
   Begin VB.PictureBox picBild1 
      BorderStyle     =   0  'Kein
      Height          =   240
      Left            =   360
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   9000
      Visible         =   0   'False
      Width           =   240
      Begin VB.TextBox txoDummy 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         Height          =   200
         Left            =   0
         TabIndex        =   63
         Top             =   6000
         Width           =   80
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4035
      Left            =   7920
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   7108
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown updCont6 
         Height          =   340
         Left            =   3810
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2210
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         Min             =   1900
         Max             =   3000
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtGeb02"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont5 
         Height          =   340
         Left            =   3810
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1850
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         Min             =   1900
         Max             =   3000
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtGeb01"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont4 
         Height          =   340
         Left            =   3810
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1490
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         Min             =   1
         Max             =   60
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtKalWo"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont3 
         Height          =   340
         Left            =   3810
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1130
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         Min             =   1
         Max             =   12
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtMonat"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtGeb02 
         Height          =   310
         Left            =   3000
         TabIndex        =   32
         Top             =   2220
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtGeb01 
         Height          =   310
         Left            =   3000
         TabIndex        =   29
         Top             =   1860
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtKalWo 
         Height          =   310
         Left            =   3000
         TabIndex        =   26
         Top             =   1500
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMonat 
         Height          =   310
         Left            =   3000
         TabIndex        =   23
         Top             =   1140
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkFilt11 
         Height          =   220
         Left            =   900
         TabIndex        =   34
         Top             =   2595
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Personen, die am"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt8 
         Height          =   220
         Left            =   900
         TabIndex        =   31
         Top             =   2235
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Personen, die vor"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt7 
         Height          =   220
         Left            =   900
         TabIndex        =   28
         Top             =   1875
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Personen, die nach"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt6 
         Height          =   220
         Left            =   900
         TabIndex        =   25
         Top             =   1515
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Personen, die in der"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFilt5 
         Height          =   220
         Left            =   900
         TabIndex        =   22
         Top             =   1155
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Personen, die im Monat"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtGeb03 
         Height          =   315
         Left            =   3000
         TabIndex        =   35
         Top             =   2580
         Width           =   345
         _Version        =   1048579
         _ExtentX        =   600
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtGeb04 
         Height          =   315
         Left            =   3420
         TabIndex        =   36
         Top             =   2580
         Width           =   345
         _Version        =   1048579
         _ExtentX        =   600
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkFilt12 
         Height          =   220
         Left            =   900
         TabIndex        =   37
         Top             =   2955
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Serienbriefadresse"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMaili 
         Height          =   315
         Left            =   3000
         TabIndex        =   38
         Top             =   2940
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkFilt14 
         Height          =   225
         Left            =   900
         TabIndex        =   39
         Top             =   3315
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungsversandart"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbVersa 
         Height          =   315
         Left            =   3000
         TabIndex        =   40
         Top             =   3320
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrFilt.frx":00D2
         Height          =   615
         Left            =   900
         TabIndex        =   76
         Top             =   300
         Width           =   6000
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "geboren wurden."
         Height          =   220
         Left            =   4200
         TabIndex        =   75
         Top             =   1905
         Width           =   1800
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "KW Geburtstag haben."
         Height          =   220
         Left            =   4200
         TabIndex        =   74
         Top             =   1545
         Width           =   1800
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Geburtstag haben."
         Height          =   220
         Left            =   4200
         TabIndex        =   73
         Top             =   1185
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "geboren wurden."
         Height          =   220
         Left            =   4200
         TabIndex        =   72
         Top             =   2265
         Width           =   1800
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "geboren wurden."
         Height          =   220
         Left            =   4200
         TabIndex        =   71
         Top             =   2595
         Width           =   1800
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4035
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   7108
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtKr4 
         Height          =   350
         Left            =   4320
         TabIndex        =   78
         Top             =   3200
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKr3 
         Height          =   350
         Left            =   4320
         TabIndex        =   79
         Top             =   2700
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKr2 
         Height          =   350
         Left            =   4320
         TabIndex        =   80
         Top             =   2200
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKr1 
         Height          =   350
         Left            =   4320
         TabIndex        =   81
         Top             =   1700
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKr8 
         Height          =   350
         Left            =   6180
         TabIndex        =   59
         Top             =   3200
         Visible         =   0   'False
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKr7 
         Height          =   350
         Left            =   6180
         TabIndex        =   54
         Top             =   2700
         Visible         =   0   'False
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKr6 
         Height          =   350
         Left            =   6180
         TabIndex        =   49
         Top             =   2200
         Visible         =   0   'False
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKr5 
         Height          =   350
         Left            =   6180
         TabIndex        =   44
         Top             =   1700
         Visible         =   0   'False
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbUo2 
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   2200
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cmbUo3 
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   2700
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cmbUo4 
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   3200
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1588
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cmbDa1 
         Height          =   315
         Left            =   1130
         TabIndex        =   41
         Top             =   1700
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbDa2 
         Height          =   315
         Left            =   1130
         TabIndex        =   46
         Top             =   2200
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbDa3 
         Height          =   315
         Left            =   1130
         TabIndex        =   51
         Top             =   2700
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbDa4 
         Height          =   315
         Left            =   1130
         TabIndex        =   56
         Top             =   3200
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBe4 
         Height          =   315
         Left            =   2630
         TabIndex        =   57
         Top             =   3200
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBe3 
         Height          =   315
         Left            =   2630
         TabIndex        =   52
         Top             =   2700
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBe2 
         Height          =   315
         Left            =   2630
         TabIndex        =   47
         Top             =   2200
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbBe1 
         Height          =   315
         Left            =   2630
         TabIndex        =   42
         Top             =   1700
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbKr1 
         Height          =   320
         Left            =   4320
         TabIndex        =   43
         Top             =   1700
         Visible         =   0   'False
         Width           =   2600
         _Version        =   1048579
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbKr2 
         Height          =   320
         Left            =   4320
         TabIndex        =   48
         Top             =   2200
         Visible         =   0   'False
         Width           =   2600
         _Version        =   1048579
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbKr3 
         Height          =   320
         Left            =   4320
         TabIndex        =   53
         Top             =   2700
         Visible         =   0   'False
         Width           =   2600
         _Version        =   1048579
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbKr4 
         Height          =   320
         Left            =   4320
         TabIndex        =   58
         Top             =   3200
         Visible         =   0   'False
         Width           =   2600
         _Version        =   1048579
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
      End
      Begin VB.Label lblLab1 
         BackStyle       =   0  'Transparent
         Caption         =   "Und/Oder"
         Height          =   195
         Left            =   180
         TabIndex        =   90
         Top             =   1360
         Width           =   795
      End
      Begin VB.Label lblLab2 
         BackStyle       =   0  'Transparent
         Caption         =   "Datenfeld"
         Height          =   195
         Left            =   1140
         TabIndex        =   89
         Top             =   1360
         Width           =   795
      End
      Begin VB.Label lblLab3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bedingung"
         Height          =   195
         Left            =   2640
         TabIndex        =   88
         Top             =   1360
         Width           =   795
      End
      Begin VB.Label lblLab4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kriterium"
         Height          =   195
         Left            =   4335
         TabIndex        =   87
         Top             =   1360
         Width           =   795
      End
      Begin VB.Label lblLa5 
         BackStyle       =   0  'Transparent
         Caption         =   "und"
         Height          =   180
         Left            =   5865
         TabIndex        =   86
         Top             =   1740
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblLa6 
         BackStyle       =   0  'Transparent
         Caption         =   "und"
         Height          =   180
         Left            =   5865
         TabIndex        =   85
         Top             =   2235
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblLa7 
         BackStyle       =   0  'Transparent
         Caption         =   "und"
         Height          =   195
         Left            =   5865
         TabIndex        =   84
         Top             =   2745
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblLa8 
         BackStyle       =   0  'Transparent
         Caption         =   "und"
         Height          =   225
         Left            =   5865
         TabIndex        =   83
         Top             =   3240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblLab9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrFilt.frx":01CB
         Height          =   795
         Left            =   300
         TabIndex        =   82
         Top             =   300
         Width           =   6645
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   4035
      Left            =   7920
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   7108
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   3600
         Left            =   160
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   220
         Width           =   3800
         _Version        =   1048579
         _ExtentX        =   6703
         _ExtentY        =   6350
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         ForeColor       =   4473924
      End
      Begin XtremeSuiteControls.RadioButton optOpti2 
         Height          =   220
         Left            =   4600
         TabIndex        =   62
         Top             =   2880
         Width           =   2055
         _Version        =   1048579
         _ExtentX        =   3625
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "ODER-Verknüpfung"
         Appearance      =   12
      End
      Begin XtremeSuiteControls.RadioButton optOpti1 
         Height          =   220
         Left            =   4600
         TabIndex        =   61
         Top             =   2580
         Width           =   2055
         _Version        =   1048579
         _ExtentX        =   3625
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "UND-Verknüpfung"
         Appearance      =   12
         Value           =   -1  'True
      End
      Begin VB.Label lblLabl10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrFilt.frx":02A4
         Height          =   2000
         Left            =   4200
         TabIndex        =   77
         Top             =   300
         Width           =   3000
      End
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   240
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdrFilt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private FeLb1 As VB.Label
Private FeLb2 As VB.Label
Private FeLb3 As VB.Label
Private FeLb4 As VB.Label
Private Bild1 As PictureBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private UpCo1 As XtremeSuiteControls.UpDown
Private UpCo2 As XtremeSuiteControls.UpDown
Private UpCo3 As XtremeSuiteControls.UpDown
Private UpCo4 As XtremeSuiteControls.UpDown
Private UpCo5 As XtremeSuiteControls.UpDown
Private UpCo6 As XtremeSuiteControls.UpDown
Private Dummy As XtremeSuiteControls.FlatEdit
Private FeUo2 As XtremeSuiteControls.ComboBox
Private FeUo3 As XtremeSuiteControls.ComboBox
Private FeUo4 As XtremeSuiteControls.ComboBox
Private FeDa1 As XtremeSuiteControls.ComboBox
Private FeDa2 As XtremeSuiteControls.ComboBox
Private FeDa3 As XtremeSuiteControls.ComboBox
Private FeDa4 As XtremeSuiteControls.ComboBox
Private FeBe1 As XtremeSuiteControls.ComboBox
Private FeBe2 As XtremeSuiteControls.ComboBox
Private FeBe3 As XtremeSuiteControls.ComboBox
Private FeBe4 As XtremeSuiteControls.ComboBox
Private FeKb1 As XtremeSuiteControls.ComboBox
Private FeKb2 As XtremeSuiteControls.ComboBox
Private FeKb3 As XtremeSuiteControls.ComboBox
Private FeKb4 As XtremeSuiteControls.ComboBox
Private FeVo2 As XtremeSuiteControls.ComboBox
Private FeV13 As XtremeSuiteControls.ComboBox
Private FeVrs As XtremeSuiteControls.ComboBox
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private FeVo3 As XtremeSuiteControls.FlatEdit
Private FeVo4 As XtremeSuiteControls.FlatEdit
Private FeVo5 As XtremeSuiteControls.FlatEdit
Private FeVo6 As XtremeSuiteControls.FlatEdit
Private FeVo7 As XtremeSuiteControls.FlatEdit
Private FeVo8 As XtremeSuiteControls.FlatEdit
Private FeVo9 As XtremeSuiteControls.FlatEdit
Private FeV10 As XtremeSuiteControls.FlatEdit
Private FeV11 As XtremeSuiteControls.FlatEdit
Private FeV12 As XtremeSuiteControls.FlatEdit
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CoDia As XtremeSuiteControls.CommonDialog
Private ToTab As XtremeCommandBars.TabControlItem
Private WithEvents TbBar As XtremeCommandBars.TabToolBar
Attribute TbBar.VB_VarHelpID = -1
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private RetWe As Long
Private FelWe() As String
Private BedWe() As String
Private TabId As Integer

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const CB_SHOWDROPDOWN = &H14F
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const GWL_WNDPROC = (-4)
Private Const KEYEVENTF_KEYUP = &H2

Private clFil As clsFile
Private clFen As clsFenster

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Sub FAusf()
On Error GoTo InErr

Dim ASQL As String
Dim FSQL As String
Dim Gefun As Long
Dim TaKey As String
Dim Mld1, Tit1 As String

Set FM = frmMain
Set TrLi1 = FM.trvList1

If GlTyp < 2 Then
    ASQL = "SELECT * FROM dbo.qryAdrSu WHERE "
Else
    ASQL = "SELECT * FROM qryAdrSu WHERE "
End If

GlAkt = True

TrLi1.Nodes("P804").Selected = True

Select Case TabId
Case 0: FSQL = ASQL & FVoFi
Case 1: FSQL = ASQL & FVoFi
Case 2: FSQL = ASQL & FStri
Case 3: FSQL = ASQL & FVoFi
End Select

Gefun = InStr(1, FSQL, "WHERE ()", 1)

If Gefun = 0 Then
    If Len(FSQL) > 1 Then
        If GlBut = RibTab_Tex_Vorlag Then
            If S_TxSet(FSQL) = True Then
                Unload Me
            Else
                Mld1 = "Es wurden keine Adressen gefunden, die den von Ihnen eingegebenen Kriterien entsprechen"
                Tit1 = "Suchfilter"
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
            End If
        ElseIf GlBut = RibTab_Tex_Dokumt Then
            If S_TxSet(FSQL) = True Then
                Unload Me
            Else
                Mld1 = "Es wurden keine Adressen gefunden, die den von Ihnen eingegebenen Kriterien entsprechen"
                Tit1 = "Suchfilter"
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
            End If
        ElseIf GlBut = RibTab_Tex_NewsLe Then
            If S_TxSet(FSQL) = True Then
                Unload Me
            Else
                Mld1 = "Es wurden keine Adressen gefunden, die den von Ihnen eingegebenen Kriterien entsprechen"
                Tit1 = "Suchfilter"
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
            End If
        Else
            If S_AdFil(FSQL) = True Then
                Unload Me
            Else
                Mld1 = "Es wurden keine Adressen gefunden, die den von Ihnen eingegebenen Kriterien entsprechen"
                Tit1 = "Suchfilter"
                WindowMess Mld1, Dial2, Tit1, FM.hwnd
            End If
        End If
    End If
Else
    Mld1 = "Sie haben noch keine Filterkriterien bestimmt"
    Tit1 = "Adressen Filtern"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

GlAkt = False

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAusf " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50031)
TeMai = IniGetOpt("Hilfe", 50032)
TeInh = IniGetOpt("Hilfe", 50033)
TeFus = IniGetOpt("Hilfe", 50034)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FMenu()
On Error GoTo InErr
'Legt alle Menüs und Toolleisten an

Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmAdrFilt
Set CmBrs = FM.comBar02
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set TrLi1 = FM.trvList1
Set CmSta = CmBrs.StatusBar
Set CmAcs = CmBrs.Actions
Set CmOpt = CmBrs.Options
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

Set Opti1 = Me.optOpti1
Set Opti2 = Me.optOpti2

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmSta
    .Font.SIZE = 8
    Set CmPan = .AddPane(1)
    CmPan.Text = Date
    CmPan.Width = 100
    CmPan.Alignment = xtpAlignmentCenter
    Set CmPan = .AddPane(2)
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    .Visible = True
End With

Set TbBar = CmBrs.AddTabToolBar("TabBar")

Set ToTab = TbBar.InsertCategory(RibTab_Opti1, "Standardfilter")
With ToTab
    .Visible = True
    .Selected = True
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Patient, "Adressen Filtern")
    With CmCon
        .ToolTipText = "Führt den konfigurierten Filter aus"
        .ShortcutText = "F6"
        .IconId = IC24_Patient_Find
        .Category = "Standardfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .ToolTipText = "Speichert die Filterkriterien"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .Category = "Standardfilter"
        .Enabled = Not GlRDP
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Zurücksetzen")
    With CmCon
        .ToolTipText = "Setzt den Filter zurück"
        .BeginGroup = True
        .IconId = IC24_Nav_Down_Left
        .Category = "Standardfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .Category = "Standardfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Abbrechen")
    With CmCon
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
        .Category = "Standardfilter"
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti2, "Weitere Filter")
With ToTab
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Patient, "Adressen Filtern")
    With CmCon
        .ToolTipText = "Führt den konfigurierten Filter aus"
        .ShortcutText = "F6"
        .IconId = IC24_Patient_Find
        .Category = "Weitere Filter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .ToolTipText = "Speichert die Filterkriterien"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .Category = "Weitere Filter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Zurücksetzen")
    With CmCon
        .ToolTipText = "Setzt den Filter zurück"
        .BeginGroup = True
        .IconId = IC24_Nav_Down_Left
        .Category = "Weitere Filter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .Category = "Weitere Filter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Abbrechen")
    With CmCon
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
        .Category = "Weitere Filter"
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti3, "Kombinationsfilter")
With ToTab
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Patient, "Adressen Filtern")
    With CmCon
        .ToolTipText = "Führt den konfigurierten Filter aus"
        .ShortcutText = "F6"
        .IconId = IC24_Patient_Find
        .Category = "Kombinationsfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .ToolTipText = "Speichert die Filterkriterien"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .Category = "Kombinationsfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Zurücksetzen")
    With CmCon
        .ToolTipText = "Setzt den Filter zurück"
        .BeginGroup = True
        .IconId = IC24_Nav_Down_Left
        .Category = "Kombinationsfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .Category = "Kombinationsfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Abbrechen")
    With CmCon
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
        .Category = "Kombinationsfilter"
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti4, "Gruppenfilter")
With ToTab
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Patient, "Adressen Filtern")
    With CmCon
        .ToolTipText = "Führt den konfigurierten Filter aus"
        .ShortcutText = "F6"
        .IconId = IC24_Patient_Find
        .Category = "Gruppenfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .ToolTipText = "Speichert die Filterkriterien"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .Category = "Gruppenfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Zurücksetzen")
    With CmCon
        .ToolTipText = "Setzt den Filter zurück"
        .BeginGroup = True
        .IconId = IC24_Nav_Down_Left
        .Category = "Gruppenfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .Category = "Gruppenfilter"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Abbrechen")
    With CmCon
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
        .Category = "Gruppenfilter"
    End With
End With

For Each CmBar In CmBrs
    If CmBar.Type = xtpBarTypeNormal Then
        Set CmCoS = CmBar.Controls
        For Each CmCon In CmCoS
            CmCon.Style = xtpButtonIconAndCaption
        Next CmCon
    End If
Next CmBar

With TrLi1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .Checkboxes = True
    .Font.SIZE = GlTFt.SIZE
    .Font.Name = GlTFt.Name
    .ForeColor = -2147483641
    .FullRowSelect = False
    .HideSelection = False
    .HotTracking = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpTreeViewLabelManual
    .Scroll = True
    .ShowLines = xtpTreeViewShowLines
    .ShowPlusMinus = True
    .SingleSel = False
End With

Set Knote = TrLi1.Nodes.Add(, , "P801", "Adressen", IC16_Folder_View)
With Knote
    .Bold = True
    .Checked = False
    .Expanded = True
End With
AdGru 3, True
DoEvents
Set Knote = TrLi1.Nodes.Add("P801", 4, "P802", "Serienmailadressen", IC16_Folder_Check)
Set Knote = TrLi1.Nodes.Add("P801", 4, "P803", "Onlinesynchronisation", IC16_Folder_Up)

'---

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
    .TabPaintManager.FixedTabWidth = 110
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

DoEvents
CmBrs.RecalcLayout
DoEvents
CmBrs.PaintManager.RefreshMetrics

Opti1.BackColor = GlBak
Opti2.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Me.chkFilt1.BackColor = GlBak
Me.chkFilt2.BackColor = GlBak
Me.chkFilt3.BackColor = GlBak
Me.chkFilt4.BackColor = GlBak
Me.chkFilt5.BackColor = GlBak
Me.chkFilt6.BackColor = GlBak
Me.chkFilt7.BackColor = GlBak
Me.chkFilt8.BackColor = GlBak
Me.chkFilt9.BackColor = GlBak
Me.chkFilt10.BackColor = GlBak
Me.chkFilt11.BackColor = GlBak
Me.chkFilt12.BackColor = GlBak
Me.chkFilt13.BackColor = GlBak
Me.chkFilt14.BackColor = GlBak

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FOpen()
On Error GoTo InErr

Dim PosGe As Long
Dim PosNa As Long
Dim StaPo As Long
Dim StrWe As String
Dim NamAn As String
Dim NamFe As String
Dim NamGe As String
Dim AktZa As Integer

Set FM = frmMain
Set FeUo2 = Me.cmbUo2
Set FeUo3 = Me.cmbUo3
Set FeUo4 = Me.cmbUo4
Set FeDa1 = Me.cmbDa1
Set FeDa2 = Me.cmbDa2
Set FeDa3 = Me.cmbDa3
Set FeDa4 = Me.cmbDa4
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set FeVo2 = Me.cmbKata1
Set FeVo3 = Me.txtBemer
Set FeVo4 = Me.txtWoPLZ
Set FeVo5 = Me.txtMonat
Set FeVo6 = Me.txtKalWo
Set FeVo7 = Me.txtGeb01
Set FeVo8 = Me.txtGeb02
Set FeV11 = Me.txtGeb03
Set FeV12 = Me.txtGeb04
Set FeV13 = Me.cmbMaili
Set FeVrs = Me.cmbVersa

StrWe = "Anamnese#qryAdrSu.Anamnese;Anrede#qryAdrSu.Anrede;Anschrift#qryAdrSu.Anschrift;" & _
"Änderung#qryAdrSu.Geändert;Bank#qryAdrSu.Bank;Begrüßung#qryAdrSu.DuSie;Behandler#qryAdrSu.IDP;" & _
"Bemerkung#qryAdrSu.Bemerkung;Briefanrede#qryAdrSu.Briefanrede;Diagnose#qryAdrSu.Diagnose;" & _
"Email#qryAdrSu.Telefon5;Erstkontakt#qryAdrSu.Datum;Favorit#qryAdrSu.VIP;" & _
"Firma#qryAdrSu.Firma1;Geboren#qryAdrSu.Geboren;Geschlecht#qryAdrSu.Geschlecht;" & _
"Katalog#qryAdrSu.ID3;Kontonummer#qryAdrSu.Konto;Kopien#qryAdrSu.Kopien;" & _
"Kurzbezeich#qryAdrSu.IDKurz;Land#qryAdrSu.Land;Landeskennung#qryAdrSu.LK;Mailing#qryAdrSu.Mailing;" & _
"Nachname#qryAdrSu.Name;Ort#qryAdrSu.Ort;PatNr#qryAdrSu.Mandant;" & _
"PLZ#qryAdrSu.PLZ;Straße#qryAdrSu.Straße;Telefon1#qryAdrSu.Telefon1;" & _
"Telefon2#qryAdrSu.Telefon2;Telefon3#qryAdrSu.Telefon3;Telefon4#qryAdrSu.Telefon4;" & _
"Titel#qryAdrSu.Titel;Vorname#qryAdrSu.Vorname;Währung#qryAdrSu.Währung;Zahlungsart#qryAdrSu.IDZ;PKV#qryAdrSu.IDV;"

With FeVrs
    .AddItem "Postversand"
    .ItemData(0) = 0
    .AddItem "Emailversand"
    .ItemData(1) = 1
    .AddItem "Downloadlink"
    .ItemData(2) = 2
    .ListIndex = GlRVs - 1 'Standard-Rechnungsversandweg
End With

With FeUo2
    .AddItem "Und"
    .ItemData(0) = 1
    .AddItem "Oder"
    .ItemData(1) = 2
End With

With FeUo3
    .AddItem "Und"
    .ItemData(0) = 1
    .AddItem "Oder"
    .ItemData(1) = 2
End With

With FeUo4
    .AddItem "Und"
    .ItemData(0) = 1
    .AddItem "Oder"
    .ItemData(1) = 2
End With

With FeV13
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
    .ListIndex = 0
End With

AktZa = 0
StaPo = 1
PosGe = InStr(StaPo, StrWe, ";", 1)
If PosGe <> 0 Then
    NamGe = Mid$(StrWe, StaPo, PosGe - 1)
    PosNa = InStr(1, NamGe, "#", 1)
    If PosNa <> 0 Then
        NamAn = Left$(NamGe, PosNa - 1)
        NamFe = Mid$(NamGe, PosNa + 1, Len(NamGe) - PosNa)
        FeDa1.AddItem NamAn
        FeDa1.ItemData(AktZa) = AktZa
        ReDim FelWe(0)
        FelWe(AktZa) = NamFe
    End If
    StaPo = PosGe + 1
    Do While PosGe <> 0
    AktZa = AktZa + 1
    PosGe = InStr(StaPo, StrWe, ";", 1)
    If PosGe <> 0 Then
        NamGe = Mid$(StrWe, StaPo, PosGe - StaPo)
        PosNa = InStr(1, NamGe, "#", 1)
        If PosNa <> 0 Then
            NamAn = Left$(NamGe, PosNa - 1)
            NamFe = Mid$(NamGe, PosNa + 1, Len(NamGe) - PosNa)
            FeDa1.AddItem NamAn
            FeDa1.ItemData(AktZa) = AktZa
            ReDim Preserve FelWe(0 To UBound(FelWe) + 1)
            FelWe(AktZa) = NamFe
        End If
    End If
    StaPo = PosGe + 1
    Loop
End If

AktZa = 0
StaPo = 1
PosGe = InStr(StaPo, StrWe, ";", 1)
If PosGe <> 0 Then
    NamGe = Mid$(StrWe, StaPo, PosGe - 1)
    PosNa = InStr(1, NamGe, "#", 1)
    If PosNa <> 0 Then
        NamAn = Left$(NamGe, PosNa - 1)
        FeDa2.AddItem NamAn
        FeDa2.ItemData(AktZa) = AktZa
    End If
    StaPo = PosGe + 1
    Do While PosGe <> 0
    AktZa = AktZa + 1
    PosGe = InStr(StaPo, StrWe, ";", 1)
    If PosGe <> 0 Then
        NamGe = Mid$(StrWe, StaPo, PosGe - StaPo)
        PosNa = InStr(1, NamGe, "#", 1)
        If PosNa <> 0 Then
            NamAn = Left$(NamGe, PosNa - 1)
            FeDa2.AddItem NamAn
            FeDa2.ItemData(AktZa) = AktZa
        End If
    End If
    StaPo = PosGe + 1
    Loop
End If

AktZa = 0
StaPo = 1
PosGe = InStr(StaPo, StrWe, ";", 1)
If PosGe <> 0 Then
    NamGe = Mid$(StrWe, StaPo, PosGe - 1)
    PosNa = InStr(1, NamGe, "#", 1)
    If PosNa <> 0 Then
        NamAn = Left$(NamGe, PosNa - 1)
        FeDa3.AddItem NamAn
        FeDa3.ItemData(AktZa) = AktZa
    End If
    StaPo = PosGe + 1
    Do While PosGe <> 0
    AktZa = AktZa + 1
    PosGe = InStr(StaPo, StrWe, ";", 1)
    If PosGe <> 0 Then
        NamGe = Mid$(StrWe, StaPo, PosGe - StaPo)
        PosNa = InStr(1, NamGe, "#", 1)
        If PosNa <> 0 Then
            NamAn = Left$(NamGe, PosNa - 1)
            FeDa3.AddItem NamAn
            FeDa3.ItemData(AktZa) = AktZa
        End If
    End If
    StaPo = PosGe + 1
    Loop
End If

AktZa = 0
StaPo = 1
PosGe = InStr(StaPo, StrWe, ";", 1)
If PosGe <> 0 Then
    NamGe = Mid$(StrWe, StaPo, PosGe - 1)
    PosNa = InStr(1, NamGe, "#", 1)
    If PosNa <> 0 Then
        NamAn = Left$(NamGe, PosNa - 1)
        FeDa4.AddItem NamAn
        FeDa4.ItemData(AktZa) = AktZa
    End If
    StaPo = PosGe + 1
    Do While PosGe <> 0
    AktZa = AktZa + 1
    PosGe = InStr(StaPo, StrWe, ";", 1)
    If PosGe <> 0 Then
        NamGe = Mid$(StrWe, StaPo, PosGe - StaPo)
        PosNa = InStr(1, NamGe, "#", 1)
        If PosNa <> 0 Then
            NamAn = Left$(NamGe, PosNa - 1)
            FeDa4.AddItem NamAn
            FeDa4.ItemData(AktZa) = AktZa
        End If
    End If
    StaPo = PosGe + 1
    Loop
End If

ReDim BedWe(100 To 110)
BedWe(100) = "Like"
BedWe(101) = "Between"
BedWe(102) = "Like"
BedWe(103) = "="
BedWe(104) = "<"
BedWe(105) = "<="
BedWe(106) = ">"
BedWe(107) = ">="
BedWe(108) = "<>"
BedWe(109) = "="
BedWe(110) = "="

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Date, "dd.mm.yyyy")
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Date, "dd.mm.yyyy")
End With

With FeVo5
    .SetMask "00", "__"
    .Pattern = "\d*"
    .Text = Format$(Date, "mm")
End With

With FeVo6
    .SetMask "00", "__"
    .Pattern = "\d*"
    .Text = Format$(Date, "ww")
End With

With FeVo7
    .SetMask "0000", "____"
    .Pattern = "\d*"
    .Text = Format$(DateAdd("yyyy", -18, Date), "yyyy")
End With

With FeVo8
    .SetMask "0000", "____"
    .Pattern = "\d*"
    .Text = Format$(DateAdd("yyyy", -18, Date), "yyyy")
End With

With FeV11
    .SetMask "00", "__"
    .Pattern = "\d*"
    .Text = Format$(Date, "dd")
End With

With FeV12
    .SetMask "00", "__"
    .Pattern = "\d*"
    .Text = Format$(Date, "mm")
End With

S_Vol 16, 5

FeVo2.ListIndex = 0

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpen " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo PoErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmAdrFilt
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set CmBrs = FM.comBar02

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    If ClBre > 10 Then
        If ClHoh > 100 Then
            Rahm1.Move 0, ClObn, ClBre, ClHoh
            Rahm2.Move 0, ClObn, ClBre, ClHoh
            Rahm3.Move 0, ClObn, ClBre, ClHoh
            Rahm4.Move 0, ClObn, ClBre, ClHoh
        End If
    End If
End If

Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FRes()
On Error Resume Next

Dim AktCo As VB.Control

For Each AktCo In Me.Controls
    If TypeName(AktCo) = "TextBox" Or TypeName(AktCo) = "ComboBox" Then
        AktCo.Text = vbNullString
    End If
Next AktCo

End Sub
Private Sub FSave()
On Error GoTo StErr

Dim FSQL As String
Dim FiNam As String
Dim TrKey As String
Dim DaNam As String
Dim AktZa As Integer

Set FM = frmMain
Set CoDia = FM.comDialo

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

Select Case TabId
Case 0: FSQL = FVoFi
Case 1: FSQL = FVoFi
Case 2: FSQL = FStri
Case 3: FSQL = FVoFi
End Select

With CoDia
    .CancelError = True
    .DialogStyle = 1
    .DefaultExt = "*.sql"
    .Filter = "SQL-Anweisungen (*.sql)|*.sql|Alle Dateien (*.*)|*.*"
    .DialogTitle = "Bitte geben Sie einen gültigen Dateinamen ein"
    .FileName = vbNullString
    .InitDir = GlFPf
    .ShowSave
    FiNam = .FileName
    If FiNam = vbNullString Then Exit Sub
End With

If Right$(FiNam, 4) <> ".sql" Then
    FiNam = FiNam & ".sql"
End If

With clFil
    .FilPfa FiNam
    DaNam = .DaNam
    If .FilVor(FiNam) = True Then
        .DaLoe = FiNam & vbNullChar
        .FilLoe
    End If
    .StrDa = FSQL
    RetWe = .FilWrSt
End With

Set TrLi1 = FM.trvList1

For Each Knote In TrLi1.Nodes
    If Left$(Knote.Key, 1) = "F" Then
        AktZa = AktZa + 1
    End If
Next Knote
TrKey = "F" & Format$(AktZa + 2, "0000")
Set Knote = TrLi1.Nodes.Add("P804", 4, TrKey, DaNam, IC16_Doc_View)

Set CoDia = Nothing
Set clFil = Nothing

Unload Me

Exit Sub

StErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Exit Sub

End Sub
Private Function FStri() As String
On Error GoTo StErr
'Kreiert Filterstring für Komplexfilter

Dim SQL1 As String
Dim Dat1 As String
Dim Dat2 As String
Dim SuSt As String
Dim SoStr As String
Dim SpIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set CmBrs = frmMain.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmCom = CmBrs.FindControl(CmCom, SY_AD_Adresse_SortFeld, , True)

SpIdx = CmCom.ListIndex

Select Case SpIdx
Case 1: SoStr = "[IDKurz]"
Case 2: SoStr = "[Name]"
Case 3: SoStr = "[Vorname]"
Case 4: SoStr = "[Firma1]"
Case 5: SoStr = "[Geboren]"
Case 6: SoStr = "[Straße]"
Case 7: SoStr = "[PLZ]"
Case 8: SoStr = "[Ort]"
Case 9: SoStr = "[Telefon1]"
Case 10: SoStr = "[Telefon1]"
Case 11: SoStr = "[Telefon5]"
Case 12:
    If GlTyp < 2 Then
        SoStr = "RIGHT ('00000000' + CONVERT (varchar(10), Mandant), 8)"
    Else
        SoStr = "Format$([Mandant],'00000000')"
    End If
End Select

SQL1 = "("

If Me.txtKr5.Visible = True And Me.txtKr1.Visible = True Then
    If Me.txtKr5.Text <> vbNullString And Me.txtKr1.Text <> vbNullString Then
        If GlTyp > 1 Then
            Dat1 = DatePart("m", Me.txtKr1.Text) & "/" & DatePart("d", Me.txtKr1.Text) & "/" & DatePart("yyyy", Me.txtKr1.Text)
            Dat2 = DatePart("m", Me.txtKr5.Text) & "/" & DatePart("d", Me.txtKr5.Text) & "/" & DatePart("yyyy", Me.txtKr5.Text)
            SQL1 = SQL1 & FelWe(Me.cmbDa1.ItemData(Me.cmbDa1.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe1.ItemData(Me.cmbBe1.ListIndex)) & Chr$(32)
            SQL1 = SQL1 & "#" & Dat1 & "# And #" & Dat2 & "# "
        Else
            Dat1 = DatePart("yyyy", Me.txtKr1.Text) & "-" & DatePart("m", Me.txtKr1.Text) & "-" & DatePart("d", Me.txtKr1.Text)
            Dat2 = DatePart("yyyy", Me.txtKr5.Text) & "-" & DatePart("m", Me.txtKr5.Text) & "-" & DatePart("d", Me.txtKr5.Text)
            SQL1 = SQL1 & FelWe(Me.cmbDa1.ItemData(Me.cmbDa1.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe1.ItemData(Me.cmbBe1.ListIndex)) & Chr$(32)
            SQL1 = SQL1 & "(CONVERT(DATETIME, '" & Dat1 & "', 102)) AND (CONVERT(DATETIME, '" & Dat2 & "', 102))"
        End If
    End If
ElseIf Me.txtKr1.Visible = True Then
    If Me.txtKr1.Text <> vbNullString Then
        SQL1 = SQL1 & FelWe(Me.cmbDa1.ItemData(Me.cmbDa1.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe1.ItemData(Me.cmbBe1.ListIndex)) & Chr$(32)
        If BedWe(Me.cmbBe1.ItemData(Me.cmbBe1.ListIndex)) = "Like" Then
            SQL1 = SQL1 & "'%" & Me.txtKr1.Text & "%'"
        Else
            SQL1 = SQL1 & Me.txtKr1.Text
        End If
    End If
ElseIf Me.cmbKr1.Visible = True Then
    If Me.cmbKr1.Text <> vbNullString Then
        SQL1 = SQL1 & FelWe(Me.cmbDa1.ItemData(Me.cmbDa1.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe1.ItemData(Me.cmbBe1.ListIndex)) & Chr$(32)
        SQL1 = SQL1 & Me.cmbKr1.ItemData(Me.cmbKr1.ListIndex)
    End If
End If

If Me.cmbUo2.Text <> vbNullString Then
    If Me.txtKr6.Visible = True And Me.txtKr2.Visible = True Then
        If Me.txtKr6.Text <> vbNullString And Me.txtKr2.Text <> vbNullString Then
            If GlTyp > 1 Then
                Dat1 = DatePart("m", Me.txtKr2.Text) & "/" & DatePart("d", Me.txtKr2.Text) & "/" & DatePart("yyyy", Me.txtKr2.Text)
                Dat2 = DatePart("m", Me.txtKr6.Text) & "/" & DatePart("d", Me.txtKr6.Text) & "/" & DatePart("yyyy", Me.txtKr6.Text)
                SQL1 = SQL1 & Chr$(32) & Me.cmbUo2.Tag & Chr$(32) & FelWe(Me.cmbDa2.ItemData(Me.cmbDa2.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe2.ItemData(Me.cmbBe2.ListIndex)) & Chr$(32)
                SQL1 = SQL1 & "#" & Dat1 & "# And #" & Dat2 & "# "
            Else
                Dat1 = DatePart("yyyy", Me.txtKr2.Text) & "-" & DatePart("m", Me.txtKr2.Text) & "-" & DatePart("d", Me.txtKr2.Text)
                Dat2 = DatePart("yyyy", Me.txtKr6.Text) & "-" & DatePart("m", Me.txtKr6.Text) & "-" & DatePart("d", Me.txtKr6.Text)
                SQL1 = SQL1 & Chr$(32) & Me.cmbUo2.Tag & Chr$(32) & FelWe(Me.cmbDa2.ItemData(Me.cmbDa2.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe2.ItemData(Me.cmbBe2.ListIndex)) & Chr$(32)
                SQL1 = SQL1 & "(CONVERT(DATETIME, '" & Dat1 & "', 102)) AND (CONVERT(DATETIME, '" & Dat2 & "', 102))"
            End If
        End If
    ElseIf Me.txtKr2.Visible = True Then
        If Me.txtKr2.Text <> vbNullString Then
            SQL1 = SQL1 & Chr$(32) & Me.cmbUo2.Tag & Chr$(32) & FelWe(Me.cmbDa2.ItemData(Me.cmbDa2.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe2.ItemData(Me.cmbBe2.ListIndex)) & Chr$(32)
            If BedWe(Me.cmbBe2.ItemData(Me.cmbBe2.ListIndex)) = "Like" Then
                SQL1 = SQL1 & "'%" & Me.txtKr2.Text & "%'"
            Else
                SQL1 = SQL1 & Me.txtKr2.Text
            End If
        End If
    ElseIf Me.cmbKr2.Visible = True Then
        If Me.cmbKr2.Text <> vbNullString Then
            SQL1 = SQL1 & Chr$(32) & Me.cmbUo2.Tag & Chr$(32) & FelWe(Me.cmbDa2.ItemData(Me.cmbDa2.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe2.ItemData(Me.cmbBe2.ListIndex)) & Chr$(32)
            SQL1 = SQL1 & Me.cmbKr2.ItemData(Me.cmbKr2.ListIndex)
        End If
    End If
End If

If Me.cmbUo3.Text <> vbNullString Then
    If Me.txtKr7.Visible = True And Me.txtKr3.Visible = True Then
        If Me.txtKr7.Text <> vbNullString And Me.txtKr3.Text <> vbNullString Then
            If GlTyp > 1 Then
                Dat1 = DatePart("m", Me.txtKr3.Text) & "/" & DatePart("d", Me.txtKr3.Text) & "/" & DatePart("yyyy", Me.txtKr3.Text)
                Dat2 = DatePart("m", Me.txtKr7.Text) & "/" & DatePart("d", Me.txtKr7.Text) & "/" & DatePart("yyyy", Me.txtKr7.Text)
                SQL1 = SQL1 & Chr$(32) & Me.cmbUo3.Tag & Chr$(32) & FelWe(Me.cmbDa3.ItemData(Me.cmbDa3.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe3.ItemData(Me.cmbBe3.ListIndex)) & Chr$(32)
                SQL1 = SQL1 & "#" & Dat1 & "# And #" & Dat2 & "# "
            Else
                Dat1 = DatePart("yyyy", Me.txtKr3.Text) & "-" & DatePart("m", Me.txtKr3.Text) & "-" & DatePart("d", Me.txtKr3.Text)
                Dat2 = DatePart("yyyy", Me.txtKr7.Text) & "-" & DatePart("m", Me.txtKr7.Text) & "-" & DatePart("d", Me.txtKr7.Text)
                SQL1 = SQL1 & Chr$(32) & Me.cmbUo3.Tag & Chr$(32) & FelWe(Me.cmbDa3.ItemData(Me.cmbDa3.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe3.ItemData(Me.cmbBe3.ListIndex)) & Chr$(32)
                SQL1 = SQL1 & "(CONVERT(DATETIME, '" & Dat1 & "', 102)) AND (CONVERT(DATETIME, '" & Dat2 & "', 102))"
            End If
        End If
    ElseIf Me.txtKr3.Visible = True Then
        If Me.txtKr3.Text <> vbNullString Then
            SQL1 = SQL1 & Chr$(32) & Me.cmbUo3.Tag & Chr$(32) & FelWe(Me.cmbDa3.ItemData(Me.cmbDa3.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe3.ItemData(Me.cmbBe3.ListIndex)) & Chr$(32)
            If BedWe(Me.cmbBe3.ItemData(Me.cmbBe3.ListIndex)) = "Like" Then
                SQL1 = SQL1 & "'%" & Me.txtKr3.Text & "%'"
            Else
                SQL1 = SQL1 & Me.txtKr3.Text
            End If
        End If
    ElseIf Me.cmbKr3.Visible = True Then
        If Me.cmbKr3.Text <> vbNullString Then
            SQL1 = SQL1 & Chr$(32) & Me.cmbUo3.Tag & Chr$(32) & FelWe(Me.cmbDa3.ItemData(Me.cmbDa3.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe3.ItemData(Me.cmbBe3.ListIndex)) & Chr$(32)
            SQL1 = SQL1 & Me.cmbKr3.ItemData(Me.cmbKr3.ListIndex)
        End If
    End If
End If

If Me.cmbUo4.Text <> vbNullString Then
    If Me.txtKr8.Visible = True And Me.txtKr4.Visible = True Then
        If Me.txtKr8.Text <> vbNullString And Me.txtKr4.Text <> vbNullString Then
            If GlTyp > 1 Then
                Dat1 = DatePart("m", Me.txtKr4.Text) & "/" & DatePart("d", Me.txtKr4.Text) & "/" & DatePart("yyyy", Me.txtKr4.Text)
                Dat2 = DatePart("m", Me.txtKr8.Text) & "/" & DatePart("d", Me.txtKr8.Text) & "/" & DatePart("yyyy", Me.txtKr8.Text)
                SQL1 = SQL1 & Chr$(32) & Me.cmbUo4.Tag & Chr$(32) & FelWe(Me.cmbDa4.ItemData(Me.cmbDa4.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe4.ItemData(Me.cmbBe4.ListIndex)) & Chr$(32)
                SQL1 = SQL1 & "#" & Dat1 & "# And #" & Dat2 & "# "
            Else
                Dat1 = DatePart("yyyy", Me.txtKr4.Text) & "-" & DatePart("m", Me.txtKr4.Text) & "-" & DatePart("d", Me.txtKr4.Text)
                Dat2 = DatePart("yyyy", Me.txtKr8.Text) & "-" & DatePart("m", Me.txtKr8.Text) & "-" & DatePart("d", Me.txtKr8.Text)
                SQL1 = SQL1 & Chr$(32) & Me.cmbUo4.Tag & Chr$(32) & FelWe(Me.cmbDa4.ItemData(Me.cmbDa4.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe4.ItemData(Me.cmbBe4.ListIndex)) & Chr$(32)
                SQL1 = SQL1 & "(CONVERT(DATETIME, '" & Dat1 & "', 102)) AND (CONVERT(DATETIME, '" & Dat2 & "', 102))"
            End If
        End If
    ElseIf Me.txtKr4.Visible = True Then
        If Me.txtKr4.Text <> vbNullString Then
            SQL1 = SQL1 & Chr$(32) & Me.cmbUo4.Tag & Chr$(32) & FelWe(Me.cmbDa4.ItemData(Me.cmbDa4.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe4.ItemData(Me.cmbBe4.ListIndex)) & Chr$(32)
            If BedWe(Me.cmbBe4.ItemData(Me.cmbBe4.ListIndex)) = "Like" Then
                SQL1 = SQL1 & "'%" & Me.txtKr4.Text & "%'"
            Else
                SQL1 = SQL1 & Me.txtKr4.Text
            End If
        End If
    ElseIf Me.cmbKr4.Visible = True Then
        If Me.cmbKr4.Text <> vbNullString Then
            SQL1 = SQL1 & Chr$(32) & Me.cmbUo4.Tag & Chr$(32) & FelWe(Me.cmbDa4.ItemData(Me.cmbDa4.ListIndex)) & Chr$(32) & BedWe(Me.cmbBe4.ItemData(Me.cmbBe4.ListIndex)) & Chr$(32)
            SQL1 = SQL1 & Me.cmbKr4.ItemData(Me.cmbKr4.ListIndex)
        End If
    End If
End If

If GlTyp > 1 Then
    SQL1 = SQL1 & ") ORDER BY " & SoStr
Else
    SQL1 = SQL1 & ") ORDER BY " & SoStr & ";"
End If

FStri = SQL1

Exit Function

StErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStri " & Err.Number
Resume Next

End Function
Private Sub FTabu(ByVal TaIdx As Long)
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmDat As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmThe As XtremeCommandBars.CommandBarComboBox

Set FM = frmAdrFilt
Set CmBrs = FM.comBar02
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set CmAcs = CmBrs.Actions

TabId = TaIdx

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If GlAsL = False Then
    Select Case TaIdx
    Case 0:
            Rahm1.Visible = True
            Rahm2.Visible = False
            Rahm3.Visible = False
            Rahm4.Visible = False
    Case 1:
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
    Case 2:
            Rahm1.Visible = False
            Rahm2.Visible = False
            Rahm3.Visible = True
            Rahm4.Visible = False
    Case 3:
            Rahm1.Visible = False
            Rahm2.Visible = False
            Rahm3.Visible = False
            Rahm4.Visible = True
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
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F6: FAusf
Case KY_F8: FSave
Case KY_F11: Unload Me
Case SY_OP_Patient: FAusf
Case SY_OP_Speichern: FSave
Case SY_OP_Reset: FRes
Case SY_OP_Hilfe: FHilfe
Case SY_OP_Abbruch: Unload Me
End Select

GlToo = False

End Sub
Private Function FVoFi() As String
On Error GoTo StErr
'Kreiert Filterstring für Vorgabefilter

Dim Datu1 As Date
Dim Datu2 As Date
Dim SQL1 As String
Dim SQL2 As String
Dim DaSt1 As String
Dim DaSt2 As String
Dim GruKy As String
Dim GrIdx As String
Dim SoStr As String
Dim Versa As Integer
Dim AktZa As Integer
Dim SpIdx As Integer
Dim Kombi As Boolean
Dim GebFi As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set CmBrs = frmMain.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set FeVo2 = Me.cmbKata1
Set FeVo3 = Me.txtBemer
Set FeVo4 = Me.txtWoPLZ
Set FeVo5 = Me.txtMonat
Set FeVo6 = Me.txtKalWo
Set FeVo7 = Me.txtGeb01
Set FeVo8 = Me.txtGeb02
Set FeVo9 = Me.txtTele2
Set FeV10 = Me.txtWoOrt
Set FeV11 = Me.txtGeb03
Set FeV12 = Me.txtGeb04
Set FeV13 = Me.cmbMaili
Set FeVrs = Me.cmbVersa
Set TrLi1 = Me.trvList1
Set Opti1 = Me.optOpti1
Set Opti2 = Me.optOpti2

Set CmCom = CmBrs.FindControl(CmCom, SY_AD_Adresse_SortFeld, , True)

SpIdx = CmCom.ListIndex

Versa = FeVrs.ItemData(FeVrs.ListIndex)

Select Case SpIdx
Case 1: SoStr = "[IDKurz]"
Case 2: SoStr = "[Name]"
Case 3: SoStr = "[Vorname]"
Case 4: SoStr = "[Firma1]"
Case 5: SoStr = "[Geboren]"
Case 6: SoStr = "[Straße]"
Case 7: SoStr = "[PLZ]"
Case 8: SoStr = "[Ort]"
Case 9: SoStr = "[Telefon1]"
Case 10: SoStr = "[Telefon1]"
Case 11: SoStr = "[Telefon5]"
Case 12:
    If GlTyp < 2 Then
        SoStr = "RIGHT ('00000000' + CONVERT (varchar(10), Mandant), 8)"
    Else
        SoStr = "Format$([Mandant],'00000000')"
    End If
End Select

AktZa = 1
SQL1 = "("

If Me.chkFilt1.Value = 1 Then
    If GlTyp < 2 Then
        DaSt1 = DatePart("yyyy", TxDa1.Text) & "-" & DatePart("m", TxDa1.Text) & "-" & DatePart("d", TxDa1.Text)
        SQL1 = SQL1 & "(Geändert < CONVERT(DATETIME, '" & DaSt1 & "', 102))"
    Else
        DaSt1 = DatePart("m", TxDa1.Text) & "/" & DatePart("d", TxDa1.Text) & "/" & DatePart("yyyy", TxDa1.Text)
        SQL1 = SQL1 & "((qryAdrSu.Geändert) < #" & DaSt1 & "#)"
    End If
    Kombi = True
End If

If Me.chkFilt13.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        DaSt2 = DatePart("yyyy", TxDa2.Text) & "-" & DatePart("m", TxDa2.Text) & "-" & DatePart("d", TxDa2.Text)
        SQL1 = SQL1 & "(Geändert > CONVERT(DATETIME, '" & DaSt2 & "', 102))"
    Else
        DaSt2 = DatePart("m", TxDa2.Text) & "/" & DatePart("d", TxDa2.Text) & "/" & DatePart("yyyy", TxDa2.Text)
        SQL1 = SQL1 & "((qryAdrSu.Geändert) > #" & DaSt2 & "#)"
    End If
    Kombi = True
End If

If Me.chkFilt2.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        SQL1 = SQL1 & "((ID3)=" & FeVo2.ItemData(FeVo2.ListIndex) & ")"
    Else
        SQL1 = SQL1 & "((qryAdrSu.ID3)=" & FeVo2.ItemData(FeVo2.ListIndex) & ")"
    End If
    Kombi = True
End If

If Me.chkFilt3.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        SQL1 = SQL1 & "((Bemerkung) Like '%" & FeVo3.Text & "%')"
    Else
        SQL1 = SQL1 & "((qryAdrSu.Bemerkung) Like '%" & FeVo3.Text & "%')"
    End If
    Kombi = True
End If

If Me.chkFilt4.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        SQL1 = SQL1 & "((PLZ) Like '" & FeVo4.Text & "%')"
    Else
        SQL1 = SQL1 & "((qryAdrSu.PLZ) Like '" & FeVo4.Text & "%')"
    End If
    Kombi = True
End If

If Me.chkFilt5.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        SQL1 = SQL1 & "(MONTH(Geboren) IS NOT NULL) AND (MONTH(Geboren) = " & CInt(FeVo5.Text) & ")"
    Else
        SQL1 = SQL1 & "((qryAdrSu.Geboren) Is Not Null) AND  ((Month(qryAdrSu.Geboren))=" & CInt(FeVo5.Text) & ")"
    End If
    Kombi = True
    GebFi = True
End If

If Me.chkFilt6.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        SQL1 = SQL1 & "(DATEPART(ww, Geboren) IS NOT NULL) AND (DATEPART(ww, Geboren) = " & CInt(FeVo6.Text) & ")"
    Else
        SQL1 = SQL1 & "((qryAdrSu.Geboren) Is Not Null) AND  ((Format(qryAdrSu.Geboren,'ww', 2, 1)) Like '" & CInt(FeVo6.Text) & "')"
    End If
    Kombi = True
    GebFi = True
End If

If Me.chkFilt7.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        SQL1 = SQL1 & "(YEAR(Geboren) IS NOT NULL) AND (YEAR(Geboren) > " & CInt(FeVo7.Text) & ")"
    Else
        SQL1 = SQL1 & "((qryAdrSu.Geboren) Is Not Null) AND ((Year(qryAdrSu.Geboren)) > " & CInt(FeVo7.Text) & ")"
    End If
    Kombi = True
    GebFi = True
End If

If Me.chkFilt8.Value = 1 Then
    If Kombi = True Then
        SQL1 = SQL1 & " AND "
    End If
    If GlTyp < 2 Then
        SQL1 = SQL1 & "(YEAR(Geboren) IS NOT NULL) AND (YEAR(Geboren) < " & CInt(FeVo8.Text) & ")"
    Else
        SQL1 = SQL1 & "((qryAdrSu.Geboren) Is Not Null) AND  ((Year(qryAdrSu.Geboren)) < " & CInt(FeVo8.Text) & ")"
    End If
    Kombi = True
    GebFi = True
End If

If Me.chkFilt9.Value = 1 Then
    If Kombi = True Then SQL1 = SQL1 & " AND "
    If GlTyp > 1 Then
        SQL1 = SQL1 & "((qryAdrSu.Telefon2) Like '" & FeVo9.Text & "%')"
    Else
        SQL1 = SQL1 & "((Telefon2) Like '" & FeVo9.Text & "%')"
    End If
    Kombi = True
End If

If Me.chkFilt14.Value = 1 Then
    If Kombi = True Then SQL1 = SQL1 & " AND "
    If GlTyp > 1 Then
        SQL1 = SQL1 & "((qryAdrSu.Versand) = " & Versa & ")"
    Else
        SQL1 = SQL1 & "((Versand) = " & Versa & ")"
    End If
    Kombi = True
End If

If Me.chkFilt10.Value = 1 Then
    If Kombi = True Then SQL1 = SQL1 & " AND "
    If GlTyp > 1 Then
        SQL1 = SQL1 & "((qryAdrSu.Ort) Like '%" & FeV10.Text & "%')"
    Else
        SQL1 = SQL1 & "((Ort) Like '%" & FeV10.Text & "%')"
    End If
    Kombi = True
End If

If Me.chkFilt11.Value = 1 Then
    If Kombi = True Then SQL1 = SQL1 & " AND "
    If GlTyp > 1 Then
        SQL1 = SQL1 & "((Day(qryAdrSu.Geboren))=" & CInt(FeV11.Text) & ") AND ((DatePart('m',qryAdrSu.Geboren))=" & CInt(FeV12.Text) & ")"
    Else
        SQL1 = SQL1 & "(DATEPART(dd, Geboren) = " & CInt(FeV11.Text) & ") AND (DATEPART(mm, Geboren) = " & CInt(FeV12.Text) & ")"
    End If
    Kombi = True
    GebFi = True
End If

If Me.chkFilt12.Value = 1 Then
    If Kombi = True Then SQL1 = SQL1 & " AND "
    If FeV13.ListIndex = 0 Then
        If GlTyp > 1 Then
            SQL1 = SQL1 & "((qryAdrSu.Mailing) = -1)"
        Else
            SQL1 = SQL1 & "((Mailing) = 1)"
        End If
    Else
        If GlTyp > 1 Then
            SQL1 = SQL1 & "((qryAdrSu.Mailing) = 0)"
        Else
            SQL1 = SQL1 & "((Mailing) = 0)"
        End If
    End If
    Kombi = True
End If

For Each Knote In TrLi1.Nodes
    If Knote.Checked = True Then
        Select Case Knote.Key
        Case "P801":
        Case "P802":
                If GlTyp < 2 Then
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " AND Mailing = 1"
                    Else
                        SQL2 = SQL2 & "Mailing = 1"
                    End If
                Else
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " AND [Mailing] = -1"
                    Else
                        SQL2 = SQL2 & "[Mailing] = -1"
                    End If
                End If
        Case "P803":
                If GlTyp < 2 Then
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " AND Edit = 1"
                    Else
                        SQL2 = SQL2 & "Edit = 1"
                    End If
                Else
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " AND [Edit] = -1"
                    Else
                        SQL2 = SQL2 & "[Edit] = -1"
                    End If
                End If
        Case Else:
            GrIdx = Mid$(Knote.Key, 2, Len(Knote.Key) - 1)
            GruKy = "o" & GrIdx & "o"
            If Opti1.Value = True Then
                If GlTyp > 1 Then
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " AND TreKey Like '%" & GruKy & "%'"
                    Else
                        SQL2 = SQL2 & "TreKey Like '%" & GruKy & "%'"
                    End If
                Else
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " AND [TreKey] Like '%" & GruKy & "%'"
                    Else
                        SQL2 = SQL2 & "[TreKey] Like '%" & GruKy & "%'"
                    End If
                End If
            Else
                If GlTyp > 1 Then
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " OR TreKey Like '%" & GruKy & "%'"
                    Else
                        SQL2 = SQL2 & "TreKey Like '%" & GruKy & "%'"
                    End If
                Else
                    If AktZa > 1 Then
                        SQL2 = SQL2 & " OR [TreKey] Like '%" & GruKy & "%'"
                    Else
                        SQL2 = SQL2 & "[TreKey] Like '%" & GruKy & "%'"
                    End If
                End If
            End If
            AktZa = AktZa + 1
        End Select
    End If
Next Knote

If SQL1 <> "(" Then
    If SQL2 <> vbNullString Then
        SQL1 = SQL1 & " AND " & SQL2
    End If
Else
    SQL1 = SQL1 & SQL2
End If

If GlTyp > 1 Then
    If GebFi = True Then
        SQL1 = SQL1 & ") ORDER BY Left$([Geboren],2);"
    Else
        SQL1 = SQL1 & ") ORDER BY " & SoStr
    End If
Else
    If GebFi = True Then
        SQL1 = SQL1 & ") ORDER BY DATEPART(dd, Geboren)"
    Else
        SQL1 = SQL1 & ") ORDER BY " & SoStr & ";"
    End If
End If

FVoFi = SQL1

Exit Function

StErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FVoFi " & Err.Number
Resume Next

End Function
Private Sub FWahl(ByVal FeIdx As Integer, ByVal FelNr As Integer)
On Error GoTo WaErr

Dim FeBed As XtremeSuiteControls.ComboBox
Dim FecKr As XtremeSuiteControls.ComboBox
Dim FetKr As XtremeSuiteControls.FlatEdit
Dim FeKri As XtremeSuiteControls.FlatEdit

Select Case FelNr
Case 1: Set FeBed = Me.cmbBe1
        Set FecKr = Me.cmbKr1
        Set FetKr = Me.txtKr1
        Set FeKri = Me.txtKr5
Case 2: Set FeBed = Me.cmbBe2
        Set FecKr = Me.cmbKr2
        Set FetKr = Me.txtKr2
        Set FeKri = Me.txtKr6
Case 3: Set FeBed = Me.cmbBe3
        Set FecKr = Me.cmbKr3
        Set FetKr = Me.txtKr3
        Set FeKri = Me.txtKr7
Case 4: Set FeBed = Me.cmbBe4
        Set FecKr = Me.cmbKr4
        Set FetKr = Me.txtKr4
        Set FeKri = Me.txtKr8
End Select

FeBed.Clear
FecKr.Clear

Select Case FeIdx
Case 0: GoSub BeTyp1
Case 1: GoSub BeTyp1
Case 2: GoSub BeTyp1
Case 3: GoSub BeTyp2
Case 4: GoSub BeTyp1
Case 5: GoSub BeTyp4
Case 6: GoSub BeTyp4
Case 7: GoSub BeTyp1
Case 8: GoSub BeTyp1
Case 9: GoSub BeTyp1
Case 10: GoSub BeTyp1
Case 11: GoSub BeTyp2
Case 12: GoSub BeTyp5
Case 13: GoSub BeTyp1
Case 14: GoSub BeTyp2
Case 15: GoSub BeTyp1
Case 16: GoSub BeTyp4
Case 17: GoSub BeTyp1
Case 18: GoSub BeTyp3
Case 19: GoSub BeTyp1
Case 20: GoSub BeTyp1
Case 21: GoSub BeTyp1
Case 22: GoSub BeTyp5
Case 23: GoSub BeTyp1
Case 24: GoSub BeTyp1
Case 25: GoSub BeTyp3
Case 26: GoSub BeTyp1
Case 27: GoSub BeTyp1
Case 28: GoSub BeTyp1
Case 29: GoSub BeTyp1
Case 30: GoSub BeTyp1
Case 31: GoSub BeTyp1
Case 32: GoSub BeTyp1
Case 33: GoSub BeTyp1
Case 34: GoSub BeTyp4
Case 35: GoSub BeTyp4
Case 36: GoSub BeTyp4
End Select

Exit Sub

BeTyp1:
With FeBed
    .AddItem "enthält"
    .ItemData(0) = 100
End With
FetKr.Visible = True
FecKr.Visible = False
Return

BeTyp2:
With FeBed
    .AddItem "zwischen"
    .ItemData(0) = 101
    .AddItem "enthält"
    .ItemData(1) = 102
End With
FetKr.Visible = True
FecKr.Visible = False
Return

BeTyp3:
With FeBed
    .AddItem "ist gleich"
    .ItemData(0) = 103
    .AddItem "kleiner als"
    .ItemData(1) = 104
    .AddItem "kleiner gleich"
    .ItemData(2) = 105
    .AddItem "größer als"
    .ItemData(3) = 106
    .AddItem "größer gleich"
    .ItemData(4) = 107
    .AddItem "ist ungleich"
    .ItemData(5) = 108
End With
FetKr.Visible = True
FecKr.Visible = False
Return

BeTyp4:
With FeBed
    .AddItem "zugeordnet ist"
    .ItemData(0) = 109
End With
FetKr.Visible = False
FecKr.Visible = True
FeKri.Visible = False
S_Vol FeIdx, FelNr
Return

BeTyp5:
With FeBed
    .AddItem "ist markiert"
    .ItemData(0) = 110
End With
FetKr.Visible = False
FecKr.Visible = True
With FecKr
    .AddItem "Ja"
    .ItemData(0) = -1
    .AddItem "Nein"
    .ItemData(1) = 0
End With
Return

WaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWahl " & Err.Number
Resume Next

End Sub
Private Sub chkFilt1_Click()

Set TxDa1 = Me.txtDatu1
Set UpCo1 = Me.updCont1

Select Case Me.chkFilt1.Value
Case 1: TxDa1.Enabled = True
        UpCo1.Enabled = True
Case 0: TxDa1.Enabled = False
        UpCo1.Enabled = False
End Select

End Sub

Private Sub chkFilt10_Click()

Set FeV10 = Me.txtWoOrt

Select Case Me.chkFilt10.Value
Case 1: FeV10.Enabled = True
Case 0: FeV10.Enabled = False
End Select

End Sub
Private Sub chkFilt11_Click()

Set FeV11 = Me.txtGeb03
Set FeV12 = Me.txtGeb04

Select Case Me.chkFilt11.Value
Case 1: FeV11.Enabled = True
        FeV12.Enabled = True
Case 0: FeV11.Enabled = False
        FeV12.Enabled = False
End Select

End Sub

Private Sub chkFilt12_Click()

Set FeV13 = Me.cmbMaili

Select Case Me.chkFilt12.Value
Case 1: FeV13.Enabled = True
Case 0: FeV13.Enabled = False
End Select

End Sub
Private Sub chkFilt13_Click()

Set TxDa2 = Me.txtDatu2
Set UpCo2 = Me.updCont2

Select Case Me.chkFilt13.Value
Case 1: TxDa2.Enabled = True
        UpCo2.Enabled = True
Case 0: TxDa2.Enabled = False
        UpCo2.Enabled = False
End Select

End Sub
Private Sub chkFilt14_Click()

Set FeVrs = Me.cmbVersa

Select Case Me.chkFilt14.Value
Case 1: FeVrs.Enabled = True
Case 0: FeVrs.Enabled = False
End Select

End Sub
Private Sub chkFilt2_Click()

Set FeVo2 = Me.cmbKata1

Select Case Me.chkFilt2.Value
Case 1: FeVo2.Enabled = True
Case 0: FeVo2.Enabled = False
End Select

End Sub
Private Sub chkFilt3_Click()

Set FeVo3 = Me.txtBemer

Select Case Me.chkFilt3.Value
Case 1: FeVo3.Enabled = True
Case 0: FeVo3.Enabled = False
End Select

End Sub
Private Sub chkFilt4_Click()

Set FeVo4 = Me.txtWoPLZ

Select Case Me.chkFilt4.Value
Case 1: FeVo4.Enabled = True
Case 0: FeVo4.Enabled = False
End Select

End Sub
Private Sub chkFilt5_Click()

Set FeVo5 = Me.txtMonat
Set UpCo3 = Me.updCont3

Select Case Me.chkFilt5.Value
Case 1: FeVo5.Enabled = True
        UpCo3.Enabled = True
Case 0: FeVo5.Enabled = False
        UpCo3.Enabled = False
End Select

End Sub
Private Sub chkFilt6_Click()

Set FeVo6 = Me.txtKalWo
Set UpCo4 = Me.updCont4

Select Case Me.chkFilt6.Value
Case 1: FeVo6.Enabled = True
        UpCo4.Enabled = True
Case 0: FeVo6.Enabled = False
        UpCo4.Enabled = False
End Select

End Sub
Private Sub chkFilt7_Click()

Set FeVo7 = Me.txtGeb01
Set UpCo5 = Me.updCont5

Select Case Me.chkFilt7.Value
Case 1: FeVo7.Enabled = True
        UpCo5.Enabled = True
Case 0: FeVo7.Enabled = False
        UpCo5.Enabled = False
End Select

End Sub
Private Sub chkFilt8_Click()

Set FeVo8 = Me.txtGeb02
Set UpCo6 = Me.updCont6

Select Case Me.chkFilt8.Value
Case 1: FeVo8.Enabled = True
        UpCo6.Enabled = True
Case 0: FeVo8.Enabled = False
        UpCo6.Enabled = False
End Select

End Sub
Private Sub chkFilt9_Click()

Set FeVo9 = Me.txtTele2

Select Case Me.chkFilt9.Value
Case 1: FeVo9.Enabled = True
Case 0: FeVo9.Enabled = False
End Select

End Sub
Private Sub cmbBe1_Click()
On Error Resume Next

If Not Me.cmbBe1.Text = vbNullString Then
    If Me.txtKr1.Visible Then
        Me.txtKr1.Enabled = True
        If Me.cmbBe1.Text = "zwischen" Then
            Me.lblLa5.Visible = True
            Me.txtKr5.Visible = True
        Else
            Me.lblLa5.Visible = False
            Me.txtKr5.Visible = False
        End If
    Else
        If Me.cmbBe1.Text <> "zwischen" Then
            Me.cmbKr1.Enabled = True
        End If
    End If
Else
    If Me.txtKr1.Visible Then
        Me.txtKr1.Enabled = False
    Else
        Me.cmbKr1.Enabled = True
    End If
    Me.lblLa5.Visible = False
    Me.txtKr5.Visible = False
End If
Me.txtKr5.Text = vbNullString

End Sub
Private Sub cmbBe1_GotFocus()
    RetWe = SendMessage(Me.cmbBe1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbBe1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub cmbBe2_Click()
On Error Resume Next

If Not Me.cmbBe2.Text = vbNullString Then
    If Me.txtKr2.Visible Then
        Me.txtKr2.Enabled = True
        If Me.cmbBe2.Text = "zwischen" Then
            Me.lblLa6.Visible = True
            Me.txtKr6.Visible = True
        Else
            Me.lblLa6.Visible = False
            Me.txtKr6.Visible = False
        End If
    Else
        If Me.cmbBe2.Text <> "zwischen" Then
            Me.cmbKr2.Enabled = True
        End If
    End If
Else
    If Me.txtKr2.Visible Then
        Me.txtKr2.Enabled = False
    Else
        Me.cmbKr2.Enabled = True
    End If
    Me.lblLa6.Visible = False
    Me.txtKr6.Visible = False
End If
Me.txtKr6.Text = vbNullString

End Sub
Private Sub cmbBe2_GotFocus()
    RetWe = SendMessage(Me.cmbBe2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbBe2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBe3_Click()
On Error Resume Next

If Not Me.cmbBe3.Text = vbNullString Then
    If Me.txtKr3.Visible Then
        Me.txtKr3.Enabled = True
        If Me.cmbBe3.Text = "zwischen" Then
            Me.lblLa7.Visible = True
            Me.txtKr7.Visible = True
        Else
            Me.lblLa7.Visible = False
            Me.txtKr7.Visible = False
        End If
    Else
        If Me.cmbBe3.Text <> "zwischen" Then
            Me.cmbKr3.Enabled = True
        End If
    End If
Else
    If Me.txtKr3.Visible Then
        Me.txtKr3.Enabled = False
    Else
        Me.cmbKr3.Enabled = True
    End If
    Me.lblLa7.Visible = False
    Me.txtKr7.Visible = False
End If
Me.txtKr7.Text = vbNullString

End Sub
Private Sub cmbBe3_GotFocus()
    RetWe = SendMessage(Me.cmbBe3.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbBe3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBe4_Click()
On Error Resume Next

If Not Me.cmbBe4.Text = vbNullString Then
    If Me.txtKr4.Visible Then
        Me.txtKr4.Enabled = True
        If Me.cmbBe4.Text = "zwischen" Then
            Me.lblLa8.Visible = True
            Me.txtKr8.Visible = True
        Else
            Me.lblLa8.Visible = False
            Me.txtKr8.Visible = False
        End If
    Else
        If Me.cmbBe4.Text <> "zwischen" Then
            Me.cmbKr4.Enabled = True
        End If
    End If
Else
    If Me.txtKr4.Visible Then
        Me.txtKr4.Enabled = False
    Else
        Me.cmbKr4.Enabled = True
    End If
    Me.lblLa8.Visible = False
    Me.txtKr8.Visible = False
End If
Me.txtKr8.Text = vbNullString

End Sub
Private Sub cmbBe4_GotFocus()
    RetWe = SendMessage(Me.cmbBe4.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbBe4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbDa1_Click()

If Me.cmbDa1.Text = vbNullString Then
    Me.cmbBe1.Enabled = False
Else
    Me.cmbBe1.Enabled = True
    FWahl Me.cmbDa1.ItemData(Me.cmbDa1.ListIndex), 1
End If

End Sub
Private Sub cmbDa1_GotFocus()
    RetWe = SendMessage(Me.cmbDa1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbDa1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbDa2_Click()

If Me.cmbDa2.Text = vbNullString Then
    Me.cmbBe2.Enabled = False
Else
    Me.cmbBe2.Enabled = True
    FWahl Me.cmbDa2.ItemData(Me.cmbDa2.ListIndex), 2
End If

End Sub
Private Sub cmbDa2_GotFocus()
    RetWe = SendMessage(Me.cmbDa2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbDa2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbDa3_Click()

If Me.cmbDa3.Text = vbNullString Then
    Me.cmbBe3.Enabled = False
Else
    Me.cmbBe3.Enabled = True
    FWahl Me.cmbDa3.ItemData(Me.cmbDa3.ListIndex), 3
End If

End Sub
Private Sub cmbDa3_GotFocus()
    RetWe = SendMessage(Me.cmbDa3.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbDa3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbDa4_Click()

If Me.cmbDa4.Text = vbNullString Then
    Me.cmbBe4.Enabled = False
Else
    Me.cmbBe4.Enabled = True
    FWahl Me.cmbDa4.ItemData(Me.cmbDa4.ListIndex), 1
End If

End Sub
Private Sub cmbDa4_GotFocus()
    RetWe = SendMessage(Me.cmbDa4.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbDa4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbKr1_Click()

If Not Me.cmbKr1.Text = vbNullString Then
    If Not Me.cmbBe1.Text = "zwischen" Then
        Me.cmbUo2.Enabled = True
    End If
Else
    Me.cmbUo2.Enabled = False
End If

End Sub
Private Sub cmbKr1_GotFocus()
    RetWe = SendMessage(Me.cmbKr1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbKr1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbKr2_Click()

If Not Me.cmbKr2.Text = vbNullString Then
    If Not Me.cmbBe2.Text = "zwischen" Then
        Me.cmbUo3.Enabled = True
    End If
Else
    Me.cmbUo3.Enabled = False
End If

End Sub
Private Sub cmbKr2_GotFocus()
    RetWe = SendMessage(Me.cmbKr2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbKr2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbKr3_Click()

If Not Me.cmbKr3.Text = vbNullString Then
    If Not Me.cmbBe3.Text = "zwischen" Then
        Me.cmbUo4.Enabled = True
    End If
Else
    Me.cmbUo4.Enabled = False
End If

End Sub
Private Sub cmbKr3_GotFocus()
    RetWe = SendMessage(Me.cmbKr3.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbKr3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbKr4_GotFocus()
    RetWe = SendMessage(Me.cmbKr4.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbKr4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbUo2_Click()

If Me.cmbUo2.Text = vbNullString Then
    Me.cmbDa2.Enabled = False
Else
    Me.cmbDa2.Enabled = True
    If Me.cmbUo2.Text = "Und" Then
        Me.cmbUo2.Tag = "And"
    Else
        Me.cmbUo2.Tag = "Or"
    End If
End If

End Sub
Private Sub cmbUo2_GotFocus()
    'RetWe = SendMessage(Me.cmbUo2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbUo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbUo3_Click()

If Me.cmbUo3.Text = vbNullString Then
    Me.cmbDa3.Enabled = False
Else
    Me.cmbDa3.Enabled = True
    If Me.cmbUo3.Text = "Und" Then
        Me.cmbUo3.Tag = "And"
    Else
        Me.cmbUo3.Tag = "Or"
    End If
End If

End Sub
Private Sub cmbUo3_GotFocus()
    'RetWe = SendMessage(Me.cmbUo3.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbUo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbUo4_Click()

If Me.cmbUo4.Text = vbNullString Then
    Me.cmbDa4.Enabled = False
Else
    Me.cmbDa4.Enabled = True
    If Me.cmbUo4.Text = "Und" Then
        Me.cmbUo4.Tag = "And"
    Else
        Me.cmbUo4.Tag = "Or"
    End If
End If

End Sub
Private Sub cmbUo4_GotFocus()
    'RetWe = SendMessage(Me.cmbUo4.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub cmbUo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlAsL = False Then FTool Control.id
End Sub
Private Sub comBar02_Resize()

Dim ClRe As RECT

If GlAsL = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    FPosi
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub Form_Activate()
    FPosi
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 5800
    .ClientMaxWidth = 8400
    .ClientMinHeight = 5200
    .ClientMinWidth = 7400
    .TopMost = True
End With

FMenu
FOpen

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAdrFilt = Nothing
End Sub
Private Sub TbBar_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    FTabu Item.Index
End Sub
Private Sub trvList1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo LaErr

Set TrLi1 = Me.trvList1

If Button = vbRightButton Then
    Set TrLi1.SelectedItem = TrLi1.HitTest(x, y)
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub trvList1_NodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Dim TreKy As String

Set TrLi1 = Me.trvList1
    
For Each Knote In TrLi1.Nodes
    Knote.Image = IC16_Folder_Close
Next Knote

Node.Image = IC16_Folder_Open
TrLi1.Nodes(1).Image = IC16_Folder_View

If Node.Key = "P801" Then
    For Each Knote In TrLi1.Nodes
        Knote.Checked = Node.Checked
    Next Knote
End If

End Sub
Private Sub trvList1_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Dim TreKy As String

Set TrLi1 = Me.trvList1
    
For Each Knote In TrLi1.Nodes
    Knote.Image = IC16_Folder_Close
Next Knote

Node.Image = IC16_Folder_Open
TrLi1.Nodes(1).Image = IC16_Folder_View

End Sub
Private Sub txtBemer_GotFocus()
    Me.txtBemer.SelStart = 0
    Me.txtBemer.SelLength = Len(Me.txtBemer.Text)
End Sub
Private Sub txtDatu1_Change()
    If Len(Me.txtDatu1.Text) = 2 Or Len(Me.txtDatu1.Text) = 5 Then
        Me.txtDatu1.Text = Me.txtDatu1.Text & "."
        Me.txtDatu1.SelStart = Len(Me.txtDatu1.Text)
    End If
End Sub
Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FAusf
End Sub

Private Sub txtGeb01_GotFocus()
    Me.txtGeb01.SelStart = 0
    Me.txtGeb01.SelLength = Len(Me.txtGeb01.Text)
End Sub
Private Sub txtGeb01_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FAusf
End Sub
Private Sub txtGeb02_GotFocus()
    Me.txtGeb02.SelStart = 0
    Me.txtGeb02.SelLength = Len(Me.txtGeb02.Text)
End Sub
Private Sub txtGeb02_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FAusf
End Sub
Private Sub txtGeb03_GotFocus()
    Me.txtGeb03.SelStart = 0
    Me.txtGeb03.SelLength = Len(Me.txtGeb03.Text)
End Sub
Private Sub txtGeb03_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FAusf
End Sub

Private Sub txtGeb04_GotFocus()
    Me.txtGeb04.SelStart = 0
    Me.txtGeb04.SelLength = Len(Me.txtGeb04.Text)
End Sub
Private Sub txtGeb04_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FAusf
End Sub

Private Sub txtKalWo_GotFocus()
    Me.txtKalWo.SelStart = 0
    Me.txtKalWo.SelLength = Len(Me.txtKalWo.Text)
End Sub
Private Sub txtKalWo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FAusf
End Sub
Private Sub txtKr1_Change()

If Not Me.txtKr1.Text = vbNullString Then
    If Not Me.cmbBe1.Text = "zwischen" Then
        Me.cmbUo2.Enabled = True
    End If
Else
    Me.cmbUo2.Enabled = False
End If

End Sub
Private Sub txtKr1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKr2_Change()

If Not Me.txtKr2.Text = vbNullString Then
    If Not Me.cmbBe2.Text = "zwischen" Then
        Me.cmbUo3.Enabled = True
    End If
Else
    Me.cmbUo3.Enabled = False
End If

End Sub
Private Sub txtKr2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKr3_Change()

If Not Me.txtKr3.Text = vbNullString Then
    If Not Me.cmbBe3.Text = "zwischen" Then
        Me.cmbUo4.Enabled = True
    End If
Else
    Me.cmbUo4.Enabled = False
End If

End Sub
Private Sub txtKr3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKr4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKr5_Change()

If Not Me.txtKr1.Text = vbNullString And Not Me.txtKr5.Text = vbNullString Then
    Me.cmbUo2.Enabled = True
Else
    Me.cmbUo2.Enabled = False
End If

If Len(Me.txtKr5.Text) = 2 Or Len(Me.txtKr5.Text) = 5 Then
   Me.txtKr5.Text = Me.txtKr5.Text & "."
   Me.txtKr5.SelStart = Len(Me.txtKr5.Text)
End If

End Sub
Private Sub txtKr5_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKr6_Change()

If Not Me.txtKr2.Text = vbNullString And Not Me.txtKr6.Text = vbNullString Then
    Me.cmbUo3.Enabled = True
Else
    Me.cmbUo3.Enabled = False
End If

If Len(Me.txtKr6.Text) = 2 Or Len(Me.txtKr6.Text) = 5 Then
   Me.txtKr6.Text = Me.txtKr6.Text & "."
   Me.txtKr6.SelStart = Len(Me.txtKr6.Text)
End If

End Sub
Private Sub txtKr6_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKr7_Change()

If Not Me.txtKr3.Text = vbNullString And Not Me.txtKr7.Text = vbNullString Then
    Me.cmbUo4.Enabled = True
Else
    Me.cmbUo4.Enabled = False
End If

If Len(Me.txtKr7.Text) = 2 Or Len(Me.txtKr7.Text) = 5 Then
   Me.txtKr7.Text = Me.txtKr7.Text & "."
   Me.txtKr7.SelStart = Len(Me.txtKr7.Text)
End If

End Sub
Private Sub txtKr7_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtKr8_Change()

If Len(Me.txtKr8.Text) = 2 Or Len(Me.txtKr8.Text) = 5 Then
   Me.txtKr8.Text = Me.txtKr8.Text & "."
   Me.txtKr8.SelStart = Len(Me.txtKr8.Text)
End If

End Sub
Private Sub txtKr8_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtMonat_GotFocus()
    Me.txtMonat.SelStart = 0
    Me.txtMonat.SelLength = Len(Me.txtMonat.Text)
End Sub
Private Sub txtMonat_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FAusf
End Sub
Private Sub txtTele2_GotFocus()
    Me.txtTele2.SelStart = 0
    Me.txtTele2.SelLength = Len(Me.txtTele2.Text)
End Sub

Private Sub txtWoOrt_GotFocus()
    Me.txtWoOrt.SelStart = 0
    Me.txtWoOrt.SelLength = Len(Me.txtWoOrt.Text)
End Sub
Private Sub txtWoPLZ_GotFocus()
    Me.txtWoPLZ.SelStart = 0
    Me.txtWoPLZ.SelLength = Len(Me.txtWoPLZ.Text)
End Sub
Private Sub updCont1_DownClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text

TxDa1.Text = DateAdd("d", 1, AltDa)

End Sub

Private Sub updCont2_DownClick()

Dim AltDa As Date

Set TxDa2 = Me.txtDatu2

AltDa = TxDa2.Text

TxDa2.Text = DateAdd("d", -1, AltDa)

End Sub
Private Sub updCont2_UpClick()

Dim AltDa As Date

Set TxDa2 = Me.txtDatu2

AltDa = TxDa2.Text

TxDa2.Text = DateAdd("d", 1, AltDa)

End Sub
