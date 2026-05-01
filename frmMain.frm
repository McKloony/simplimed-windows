VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{C215CB9A-0AE1-499F-A101-48B3C370D3DF}#16.3#0"; "Codejock.ChartPro.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.DockingPane.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.ShortcutBar.v16.3.1.ocx"
Object = "{5B44EC52-B95B-45CF-98FF-A49DFEED5A92}#16.3#0"; "Codejock.PropertyGrid.v16.3.1.ocx"
Object = "{0A354FBB-9D11-4E4C-A84A-435FAC06E59A}#13.0#0"; "fldrv012.ocx"
Object = "{DCFE21FC-18E2-480F-82BF-057A7A5D6422}#13.0#0"; "filev012.ocx"
Object = "{621DDB00-A516-11E8-A658-0013D350667C}#3.2#0"; "tx4ole26.ocx"
Begin VB.Form frmMain 
   Caption         =   "SimpliMed"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12885
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   12885
   Begin XtremeReportControl.ReportControl repContK 
      Height          =   1005
      Left            =   6720
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1764
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont6 
      Height          =   1000
      Left            =   7920
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1000
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1764
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      OLEDropMode     =   1
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont8 
      Height          =   1005
      Left            =   9120
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1000
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1764
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont4 
      Height          =   1005
      Left            =   9120
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1000
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1764
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont2 
      Height          =   1005
      Left            =   7920
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1000
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1764
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1005
      Left            =   6720
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1764
      _StockProps     =   64
      AutoColumnSizing=   0   'False
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont0 
      Height          =   1005
      Left            =   10680
      TabIndex        =   94
      Top             =   4700
      Visible         =   0   'False
      Width           =   1005
      _Version        =   1048579
      _ExtentX        =   1764
      _ExtentY        =   1764
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin VB.PictureBox picRah14 
      BorderStyle     =   0  'Kein
      Height          =   500
      Left            =   10800
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   226
      Top             =   5800
      Visible         =   0   'False
      Width           =   800
      Begin XtremeSuiteControls.Label lblDeta9 
         Height          =   340
         Left            =   0
         TabIndex        =   227
         Top             =   0
         Width           =   700
         _Version        =   1048579
         _ExtentX        =   1235
         _ExtentY        =   600
         _StockProps     =   79
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
         EnableMarkup    =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm9 
      Height          =   1995
      Left            =   10680
      TabIndex        =   220
      Top             =   8040
      Visible         =   0   'False
      Width           =   2595
      _Version        =   1048579
      _ExtentX        =   4586
      _ExtentY        =   3528
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin Tx4oleLib.TXRuler TexRule1 
         Height          =   405
         Left            =   100
         TabIndex        =   221
         Top             =   0
         Width           =   2000
         _Version        =   196610
         _ExtentX        =   3528
         _ExtentY        =   714
         _StockProps     =   96
         Language        =   49
         ScaleUnits      =   0
         Appearance      =   3
         Direction       =   0
         EnablePageMargins=   -1  'True
         RightToLeft     =   0   'False
         ReadOnly        =   0   'False
         FormulaMode     =   0
      End
      Begin Tx4oleLib.TXRuler TexRule2 
         Height          =   1400
         Left            =   100
         TabIndex        =   222
         Top             =   400
         Width           =   405
         _Version        =   196610
         _ExtentX        =   714
         _ExtentY        =   2469
         _StockProps     =   96
         Language        =   49
         ScaleUnits      =   0
         Appearance      =   3
         Direction       =   1
         EnablePageMargins=   -1  'True
         RightToLeft     =   0   'False
         ReadOnly        =   0   'False
         FormulaMode     =   0
      End
      Begin Tx4oleLib.TXTextControl TexCont1 
         Height          =   1400
         Left            =   500
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   400
         Width           =   2000
         _Version        =   196610
         _ExtentX        =   3528
         _ExtentY        =   2469
         _StockProps     =   73
         BackColor       =   16777215
         Language        =   49
         BorderStyle     =   1
         BackStyle       =   1
         ControlChars    =   0   'False
         EditMode        =   0
         HideSelection   =   -1  'True
         InsertionMode   =   -1  'True
         MousePointer    =   0
         ZoomFactor      =   100
         ViewMode        =   3
         ClipChildren    =   0   'False
         ClipSiblings    =   -1  'True
         SizeMode        =   0
         TabKey          =   -1  'True
         FormatSelection =   0   'False
         VTSpellDictionary=   ""
         ScrollBars      =   3
         PageWidth       =   12240
         PageHeight      =   15840
         PageMarginL     =   1440
         PageMarginT     =   1440
         PageMarginR     =   1440
         PageMarginB     =   1440
         PrintZoom       =   100
         PrintOffset     =   0   'False
         PrintColors     =   -1  'True
         FontName        =   "Arial"
         FontSize        =   11
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Baseline        =   0
         TextBkColor     =   16777215
         Alignment       =   0
         LineSpacing     =   100
         LineSpacingT    =   0
         FrameStyle      =   32
         FrameDistance   =   0
         FrameLineWidth  =   20
         IndentL         =   0
         IndentR         =   0
         IndentFL        =   0
         IndentT         =   0
         IndentB         =   0
         Text            =   ""
         WordWrapMode    =   1
         AllowUndo       =   -1  'True
         TextFrameMarkerLines=   -1  'True
         FieldLinkTargetMarkers=   0   'False
         PageOrientation =   0
         PageViewStyle   =   1
         FontSettings    =   0
         AllowDrag       =   0   'False
         AllowDrop       =   0   'False
         SelectionViewMode=   1
         SectionRestartPageNumbering=   0
         PermanentControlChars=   16
         RightToLeft     =   0   'False
         TextDirection   =   2
         Locale          =   1031
         Justification   =   1
         FrameColor      =   16777215
         FrameLineColor  =   0
         DocumentPermissions=   31
         SelectObjects   =   -1  'True
         IsTrackChangesEnabled=   0   'False
         IsFormulaCalculationEnabled=   -1  'True
         FormulaReferenceStyle=   0
      End
   End
   Begin XtremeSuiteControls.Resizer rszRahm1 
      Height          =   2175
      Left            =   15960
      TabIndex        =   90
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
      _Version        =   1048579
      _ExtentX        =   2778
      _ExtentY        =   3836
      _StockProps     =   1
      Begin XtremeCommandBars.BackstageButton bksButt14 
         Height          =   300
         Left            =   1080
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   360
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt13 
         Height          =   300
         Left            =   1080
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   0
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt12 
         Height          =   300
         Left            =   720
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   1080
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt11 
         Height          =   300
         Left            =   360
         TabIndex        =   199
         TabStop         =   0   'False
         Top             =   1080
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageSeparator bksSpep03 
         Height          =   255
         Left            =   740
         TabIndex        =   200
         Top             =   1800
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   450
         _StockProps     =   2
         MarkupText      =   ""
         Appearance      =   12
      End
      Begin XtremeCommandBars.BackstageSeparator bksSpep02 
         Height          =   255
         Left            =   380
         TabIndex        =   201
         Top             =   1800
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   450
         _StockProps     =   2
         MarkupText      =   ""
         Appearance      =   12
      End
      Begin XtremeCommandBars.BackstageButton bksButt10 
         Height          =   300
         Left            =   0
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   1080
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt09 
         Height          =   300
         Left            =   720
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt08 
         Height          =   300
         Left            =   720
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   360
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt07 
         Height          =   300
         Left            =   720
         TabIndex        =   205
         TabStop         =   0   'False
         Top             =   0
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt06 
         Height          =   300
         Left            =   360
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt05 
         Height          =   300
         Left            =   360
         TabIndex        =   207
         TabStop         =   0   'False
         Top             =   360
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt04 
         Height          =   300
         Left            =   360
         TabIndex        =   208
         TabStop         =   0   'False
         Top             =   0
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt03 
         Height          =   300
         Left            =   0
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt02 
         Height          =   300
         Left            =   0
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   360
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt01 
         Height          =   300
         Left            =   0
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   0
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageSeparator bksSpep04 
         Height          =   255
         Left            =   1100
         TabIndex        =   212
         Top             =   1800
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   450
         _StockProps     =   2
         MarkupText      =   ""
         Appearance      =   12
      End
      Begin XtremeCommandBars.BackstageButton bksButt15 
         Height          =   300
         Left            =   1080
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageSeparator bksSpep01 
         Height          =   255
         Left            =   0
         TabIndex        =   214
         Top             =   1800
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   450
         _StockProps     =   2
         MarkupText      =   ""
         Appearance      =   12
      End
      Begin XtremeCommandBars.BackstageButton bksButt16 
         Height          =   300
         Left            =   1080
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   1080
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt17 
         Height          =   300
         Left            =   0
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   1460
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt18 
         Height          =   300
         Left            =   360
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   1460
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt19 
         Height          =   300
         Left            =   720
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   1460
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
      Begin XtremeCommandBars.BackstageButton bksButt20 
         Height          =   300
         Left            =   1080
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   1460
         Width           =   300
         _Version        =   1048579
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   79
         Enabled         =   0   'False
         Appearance      =   0
      End
   End
   Begin FileViewControl.FileView filView1 
      Height          =   735
      Left            =   16080
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   64
      ViewStyle       =   3
      CurrentFolder   =   "frmMain.frx":6852
      AllowZipFolders =   0   'False
      HideSelection   =   0   'False
      SetTextBackColor=   -1
   End
   Begin VB.PictureBox picRah13 
      BorderStyle     =   0  'Kein
      Height          =   1095
      Left            =   5880
      ScaleHeight     =   1095
      ScaleWidth      =   1005
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   1000
      Begin XtremeSuiteControls.FlatEdit txtBiKom 
         Height          =   375
         Left            =   100
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _Version        =   1048579
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
      End
   End
   Begin VB.PictureBox picRah01 
      BorderStyle     =   0  'Kein
      Height          =   2600
      Left            =   10320
      ScaleHeight     =   2595
      ScaleWidth      =   1305
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   1300
      Begin XtremeReportControl.ReportControl repCont3 
         Height          =   1005
         Left            =   0
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   1764
         _StockProps     =   64
         AutoColumnSizing=   0   'False
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeReportControl.ReportControl repCont5 
         Height          =   1005
         Left            =   0
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   1764
         _StockProps     =   64
         AutoColumnSizing=   0   'False
         FreezeColumnsAbs=   0   'False
      End
   End
   Begin VB.PictureBox picRah03 
      BorderStyle     =   0  'Kein
      Height          =   1875
      Left            =   2040
      ScaleHeight     =   1875
      ScaleWidth      =   1905
      TabIndex        =   79
      Top             =   5040
      Visible         =   0   'False
      Width           =   1900
      Begin XtremeSuiteControls.ListView lstView2 
         Height          =   645
         Left            =   120
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1080
         Width           =   795
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1147
         _StockProps     =   77
         BackColor       =   -2147483643
         OLEDropMode     =   1
      End
      Begin XtremeSuiteControls.ListView lstView3 
         Height          =   650
         Left            =   120
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   300
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1147
         _StockProps     =   77
         BackColor       =   -2147483643
         OLEDropMode     =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDeta1 
         Height          =   645
         Left            =   960
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   795
         _Version        =   1048579
         _ExtentX        =   1402
         _ExtentY        =   1147
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDeta0 
         Height          =   645
         Left            =   960
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   795
         _Version        =   1048579
         _ExtentX        =   1402
         _ExtentY        =   1147
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   225
         Left            =   120
         TabIndex        =   85
         Top             =   20
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   225
         Left            =   960
         TabIndex        =   84
         Top             =   20
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483643
      End
   End
   Begin VB.PictureBox picRah04 
      BorderStyle     =   0  'Kein
      Height          =   1040
      Left            =   4080
      ScaleHeight     =   1035
      ScaleWidth      =   1905
      TabIndex        =   76
      Top             =   5040
      Visible         =   0   'False
      Width           =   1900
      Begin XtremeSuiteControls.FlatEdit txtDeta8 
         Height          =   650
         Left            =   120
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   300
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1147
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDeta3 
         Height          =   650
         Left            =   960
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   300
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1147
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
   End
   Begin VB.PictureBox picRah06 
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   4200
      ScaleHeight     =   735
      ScaleWidth      =   1380
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   1380
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   495
         Left            =   120
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   120
         Width           =   500
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   882
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         HideSelection   =   0   'False
         OLEDropMode     =   1
      End
      Begin XtremeSuiteControls.Label lblDeta4 
         Height          =   495
         Left            =   720
         TabIndex        =   75
         Top             =   120
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   882
         _StockProps     =   79
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
         EnableMarkup    =   -1  'True
      End
   End
   Begin VB.PictureBox picRah05 
      BorderStyle     =   0  'Kein
      Height          =   1140
      Left            =   6100
      ScaleHeight     =   1140
      ScaleWidth      =   1005
      TabIndex        =   71
      Top             =   5040
      Visible         =   0   'False
      Width           =   1000
      Begin XtremeShortcutBar.ShortcutBar shtCut01 
         Height          =   885
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1561
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picRah02 
      BorderStyle     =   0  'Kein
      Height          =   1040
      Left            =   15
      ScaleHeight     =   1035
      ScaleWidth      =   1905
      TabIndex        =   65
      Top             =   5040
      Visible         =   0   'False
      Width           =   1900
      Begin XtremeSuiteControls.ListView lstView1 
         Height          =   650
         Left            =   120
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   300
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1147
         _StockProps     =   77
         BackColor       =   -2147483643
         OLEDropMode     =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDeta6 
         Height          =   650
         Left            =   960
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   300
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1402
         _ExtentY        =   1147
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDeta7 
         Height          =   200
         Left            =   0
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   8000
         Visible         =   0   'False
         Width           =   200
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   225
         Left            =   960
         TabIndex        =   70
         Top             =   20
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   225
         Left            =   120
         TabIndex        =   69
         Top             =   20
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483643
      End
   End
   Begin VB.PictureBox picRah07 
      BorderStyle     =   0  'Kein
      Height          =   1695
      Left            =   10200
      ScaleHeight     =   1695
      ScaleWidth      =   1455
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
      Begin FolderViewControl.FolderView fldView1 
         Height          =   975
         Left            =   120
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   1720
         _StockProps     =   64
         root            =   "frmMain.frx":686A
         rrr             =   "test"
         AllowZipFolders =   0   'False
         HideSelection   =   0   'False
         QueryHasSubFolders=   0   'False
         FullRowSelect   =   0   'False
         LineColor       =   -16777216
         NodeHeight      =   -1
      End
      Begin XtremeSuiteControls.TreeView trvList2 
         Height          =   495
         Left            =   120
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   500
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   882
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         HideSelection   =   0   'False
      End
   End
   Begin VB.PictureBox picRah08 
      BorderStyle     =   0  'Kein
      Height          =   7695
      Left            =   12000
      ScaleHeight     =   7695
      ScaleWidth      =   3735
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   3735
      Begin XtremeSuiteControls.GroupBox frmRahm6 
         Height          =   3600
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   3405
         _Version        =   1048579
         _ExtentX        =   6006
         _ExtentY        =   6350
         _StockProps     =   79
         Appearance      =   6
         BorderStyle     =   2
         Begin XtremeSuiteControls.RadioButton optGrup6 
            Height          =   230
            Left            =   400
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppiert nach Mandanten"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optGrup5 
            Height          =   230
            Left            =   400
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1900
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppiert nach Patienten"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optGrup4 
            Height          =   230
            Left            =   400
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppiert nach Abgeschlossen"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optGrup3 
            Height          =   230
            Left            =   400
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1100
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppiert nach Monat"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optGrup2 
            Height          =   230
            Left            =   400
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   700
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppiert nach Datum"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optGrup1 
            Height          =   230
            Left            =   400
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Keine Gruppierung"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkExpan 
            Height          =   230
            Left            =   400
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   2900
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppen Expandieren"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkStorn 
            Height          =   225
            Left            =   400
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   3300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Zeige Stornierte"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox frmRahm5 
         Height          =   3500
         Left            =   120
         TabIndex        =   18
         Top             =   3960
         Visible         =   0   'False
         Width           =   3405
         _Version        =   1048579
         _ExtentX        =   6006
         _ExtentY        =   6165
         _StockProps     =   79
         Appearance      =   6
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnDatu3 
            Height          =   315
            Left            =   2580
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2480
            Width           =   315
            _Version        =   1048579
            _ExtentX        =   547
            _ExtentY        =   547
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZeit4 
            Height          =   225
            Left            =   300
            TabIndex        =   25
            Top             =   2030
            Width           =   940
            _Version        =   1048579
            _ExtentX        =   1658
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Zeitraum"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZeit3 
            Height          =   225
            Left            =   300
            TabIndex        =   23
            Top             =   1450
            Width           =   940
            _Version        =   1048579
            _ExtentX        =   1658
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Jahr"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZeit2 
            Height          =   225
            Left            =   300
            TabIndex        =   21
            Top             =   870
            Width           =   940
            _Version        =   1048579
            _ExtentX        =   1658
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Quartal"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optZeit1 
            Height          =   225
            Left            =   300
            TabIndex        =   19
            Top             =   290
            Width           =   940
            _Version        =   1048579
            _ExtentX        =   1658
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Monat"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbMonat 
            Height          =   315
            Left            =   1360
            TabIndex        =   20
            Top             =   260
            Width           =   1500
            _Version        =   1048579
            _ExtentX        =   2646
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cmbQurta 
            Height          =   315
            Left            =   1360
            TabIndex        =   22
            Top             =   840
            Width           =   1500
            _Version        =   1048579
            _ExtentX        =   2646
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Style           =   2
            Text            =   "ComboBox2"
         End
         Begin XtremeSuiteControls.FlatEdit txtDatu3 
            Height          =   315
            Left            =   1360
            TabIndex        =   28
            Top             =   2480
            Width           =   1200
            _Version        =   1048579
            _ExtentX        =   2117
            _ExtentY        =   547
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            BackColor       =   16777215
            Alignment       =   2
         End
         Begin XtremeSuiteControls.ComboBox cmbJahre 
            Height          =   315
            Left            =   1360
            TabIndex        =   24
            Top             =   1420
            Width           =   1500
            _Version        =   1048579
            _ExtentX        =   2646
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Style           =   2
         End
         Begin XtremeCalendarControl.DatePicker dtpDatu6 
            Height          =   405
            Left            =   480
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   405
            _Version        =   1048579
            _ExtentX        =   706
            _ExtentY        =   706
            _StockProps     =   64
            Show3DBorder    =   2
            VisualTheme     =   0
         End
         Begin XtremeSuiteControls.PushButton btnDatu2 
            Height          =   315
            Left            =   2580
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2000
            Width           =   315
            _Version        =   1048579
            _ExtentX        =   547
            _ExtentY        =   547
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDatu2 
            Height          =   315
            Left            =   1360
            TabIndex        =   26
            Top             =   2000
            Width           =   1200
            _Version        =   1048579
            _ExtentX        =   2117
            _ExtentY        =   547
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            BackColor       =   16777215
            Alignment       =   2
         End
         Begin XtremeSuiteControls.ComboBox cmbVgJah 
            Height          =   315
            Left            =   1360
            TabIndex        =   32
            Top             =   3010
            Width           =   1500
            _Version        =   1048579
            _ExtentX        =   2646
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4473924
            BackColor       =   16777215
            Style           =   2
         End
         Begin XtremeSuiteControls.CheckBox chkJahre 
            Height          =   255
            Left            =   300
            TabIndex        =   31
            Top             =   3040
            Width           =   940
            _Version        =   1048579
            _ExtentX        =   1658
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Vergleich"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
      End
   End
   Begin VB.PictureBox picRah09 
      BorderStyle     =   0  'Kein
      Height          =   4000
      Left            =   120
      ScaleHeight     =   4005
      ScaleWidth      =   5505
      TabIndex        =   13
      Top             =   7000
      Visible         =   0   'False
      Width           =   5500
      Begin XtremeReportControl.ReportControl repContT 
         Height          =   855
         Left            =   4440
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   855
         _Version        =   1048579
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu7 
         Height          =   800
         Left            =   3500
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1411
         _StockProps     =   64
         Show3DBorder    =   0
         VisualTheme     =   0
      End
      Begin XtremeSuiteControls.GroupBox frmRahm7 
         Height          =   3800
         Left            =   120
         TabIndex        =   182
         Top             =   120
         Visible         =   0   'False
         Width           =   3405
         _Version        =   1048579
         _ExtentX        =   6006
         _ExtentY        =   6703
         _StockProps     =   79
         Appearance      =   6
         BorderStyle     =   2
         Begin XtremeSuiteControls.RadioButton optTeGr6 
            Height          =   230
            Left            =   400
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Mandanten"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optTeGr5 
            Height          =   230
            Left            =   400
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1900
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Patienten"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optTeGr4 
            Height          =   230
            Left            =   400
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Seriennummer"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optTeGr3 
            Height          =   230
            Left            =   400
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1100
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Monat"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optTeGr2 
            Height          =   230
            Left            =   400
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   700
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Datum"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optTeGr1 
            Height          =   230
            Left            =   400
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Keine Gruppierung"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkExpa2 
            Height          =   230
            Left            =   400
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2900
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppen Expandieren"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkStor2 
            Height          =   225
            Left            =   400
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   3300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Zeige Entfernte"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption schCapt1 
         Height          =   800
         Left            =   3500
         TabIndex        =   16
         Top             =   960
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   1411
         _StockProps     =   14
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   5
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox picRah10 
      BorderStyle     =   0  'Kein
      Height          =   5000
      Left            =   7200
      ScaleHeight     =   4995
      ScaleWidth      =   3135
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6800
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox picBild2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         Height          =   500
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.PictureBox picBild1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         Height          =   500
         Left            =   720
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   500
      End
      Begin XtremeSuiteControls.FlatEdit txtDeta5 
         Height          =   200
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   12000
         Width           =   200
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDeta2 
         Height          =   1005
         Left            =   1920
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   1764
         _StockProps     =   77
         ForeColor       =   255
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.TreeView trvList5 
         Height          =   495
         Left            =   1320
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   882
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox frmRahm8 
         Height          =   3500
         Left            =   0
         TabIndex        =   187
         Top             =   1600
         Visible         =   0   'False
         Width           =   3405
         _Version        =   1048579
         _ExtentX        =   6006
         _ExtentY        =   6174
         _StockProps     =   79
         Appearance      =   6
         BorderStyle     =   2
         Begin XtremeSuiteControls.RadioButton optLaGr6 
            Height          =   230
            Left            =   400
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   2300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Mandanten"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optLaGr5 
            Height          =   230
            Left            =   400
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   1900
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Patienten"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optLaGr4 
            Height          =   230
            Left            =   400
            TabIndex        =   190
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Seriennummer"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optLaGr3 
            Height          =   230
            Left            =   400
            TabIndex        =   191
            TabStop         =   0   'False
            Top             =   1100
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Monat"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optLaGr2 
            Height          =   230
            Left            =   400
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   700
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppieren nach Datum"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optLaGr1 
            Height          =   230
            Left            =   400
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   300
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Keine Gruppierung"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkExpa3 
            Height          =   230
            Left            =   400
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   2900
            Width           =   2700
            _Version        =   1048579
            _ExtentX        =   4762
            _ExtentY        =   406
            _StockProps     =   79
            Caption         =   "Gruppen Expandieren"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label lblDeta7 
         Height          =   495
         Left            =   120
         TabIndex        =   224
         Top             =   720
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblDeta5 
         Height          =   495
         Left            =   720
         TabIndex        =   12
         Top             =   720
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   882
         _StockProps     =   79
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblDeta6 
         Height          =   495
         Left            =   1320
         TabIndex        =   184
         Top             =   720
         Visible         =   0   'False
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   882
         _StockProps     =   79
         BackColor       =   -2147483643
         Alignment       =   4
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picRah12 
      BorderStyle     =   0  'Kein
      Height          =   1575
      Left            =   7080
      ScaleHeight     =   1575
      ScaleWidth      =   3405
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   3400
      Begin XtremeSuiteControls.TabControl TabCont1 
         Height          =   1400
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   2469
         _StockProps     =   68
         ItemCount       =   2
         Item(0).Caption =   "Einzelbrief"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabPage1"
         Item(1).Caption =   "Serienbrief"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabPage2"
         Begin XtremeSuiteControls.TabControlPage TabPage2 
            Height          =   1035
            Left            =   -69970
            TabIndex        =   3
            Top             =   330
            Visible         =   0   'False
            Width           =   2445
            _Version        =   1048579
            _ExtentX        =   4313
            _ExtentY        =   1826
            _StockProps     =   1
            Page            =   1
            Begin XtremeReportControl.ReportControl repCont9 
               Height          =   800
               Left            =   720
               TabIndex        =   4
               Top             =   120
               Width           =   1005
               _Version        =   1048579
               _ExtentX        =   1773
               _ExtentY        =   1411
               _StockProps     =   64
               AutoColumnSizing=   0   'False
               FreezeColumnsAbs=   0   'False
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabPage1 
            Height          =   1035
            Left            =   30
            TabIndex        =   5
            Top             =   330
            Width           =   2445
            _Version        =   1048579
            _ExtentX        =   4313
            _ExtentY        =   1826
            _StockProps     =   1
            Page            =   0
            Begin XtremeSuiteControls.ListView lstView4 
               Height          =   650
               Left            =   1080
               TabIndex        =   185
               TabStop         =   0   'False
               Top             =   120
               Width           =   800
               _Version        =   1048579
               _ExtentX        =   1411
               _ExtentY        =   1147
               _StockProps     =   77
               BackColor       =   -2147483643
               OLEDropMode     =   1
            End
            Begin XtremeSuiteControls.Label lblDeta8 
               Height          =   650
               Left            =   120
               TabIndex        =   225
               Top             =   120
               Width           =   800
               _Version        =   1048579
               _ExtentX        =   1411
               _ExtentY        =   1147
               _StockProps     =   79
               BackColor       =   -2147483643
               Alignment       =   4
               WordWrap        =   -1  'True
               EnableMarkup    =   -1  'True
            End
         End
      End
      Begin XtremeSuiteControls.TreeView trvList3 
         Height          =   495
         Left            =   2760
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   882
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         HideSelection   =   0   'False
      End
   End
   Begin FileViewControl.FileView filView2 
      Height          =   855
      Left            =   6720
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   64
      ViewStyle       =   1
      CurrentFolder   =   "frmMain.frx":6882
      AllowZipFolders =   0   'False
      HideSelection   =   0   'False
      SetTextBackColor=   -1
   End
   Begin XtremeCalendarControl.CalendarControl calCont1 
      Height          =   1455
      Left            =   8040
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2055
      _Version        =   1048579
      _ExtentX        =   3625
      _ExtentY        =   2566
      _StockProps     =   64
      VisualTheme     =   0
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   1575
      Left            =   16080
      TabIndex        =   100
      Top             =   200
      Visible         =   0   'False
      Width           =   1335
      _Version        =   1048579
      _ExtentX        =   2355
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "GroupBox1"
      BackColor       =   -2147483643
      Appearance      =   6
      BorderStyle     =   2
      Begin VB.PictureBox picBild5 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox picBild6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin XtremeChartControl.ChartControl chrCont1 
      Height          =   700
      Left            =   6120
      TabIndex        =   103
      Top             =   6240
      Width           =   765
      _Version        =   1048579
      _ExtentX        =   1349
      _ExtentY        =   1235
      _StockProps     =   0
   End
   Begin XtremePropertyGrid.PropertyGrid prpGrid1 
      Height          =   800
      Left            =   10800
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   7000
      Visible         =   0   'False
      Width           =   800
      _Version        =   1048579
      _ExtentX        =   1411
      _ExtentY        =   1411
      _StockProps     =   68
      ToolBarVisible  =   0   'False
      HelpVisible     =   -1  'True
      PropertySort    =   0
      HelpHeight      =   30
      VisualTheme     =   7
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu3 
      Height          =   645
      Left            =   240
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   1138
      _StockProps     =   64
      Show3DBorder    =   2
      VisualTheme     =   0
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   2895
      Left            =   16000
      TabIndex        =   106
      Top             =   5520
      Visible         =   0   'False
      Width           =   4695
      _Version        =   1048579
      _ExtentX        =   8281
      _ExtentY        =   5106
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtBetra 
         Height          =   210
         Left            =   2700
         TabIndex        =   139
         Top             =   1320
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   714
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   1
         FlatStyle       =   -1  'True
         UseVisualStyle  =   0   'False
      End
      Begin VB.TextBox txtReTex 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Height          =   495
         Left            =   240
         MultiLine       =   -1  'True
         OLEDropMode     =   1  'Manuell
         TabIndex        =   109
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin XtremeSuiteControls.FlatEdit txtDaVon 
         Height          =   210
         Left            =   600
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1320
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDaBis 
         Height          =   210
         Left            =   600
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   1560
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAbsNa 
         Height          =   210
         Left            =   120
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAnzBi 
         Height          =   210
         Left            =   2040
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   2280
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAnzWo 
         Height          =   210
         Left            =   2520
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   1800
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAnzTe 
         Height          =   210
         Left            =   2040
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   1800
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHeil1 
         Height          =   210
         Left            =   2040
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   1560
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHeil2 
         Height          =   210
         Left            =   2520
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   2520
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIndka 
         Height          =   210
         Left            =   2040
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   2040
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDiagn 
         Height          =   210
         Left            =   2520
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   2040
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTheZi 
         Height          =   210
         Left            =   2520
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   1560
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBegru 
         Height          =   210
         Left            =   2520
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   2280
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtKilom 
         Height          =   210
         Left            =   120
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   2520
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArzNr 
         Height          =   210
         Left            =   600
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   2520
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtReNum 
         Height          =   210
         Left            =   1080
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   2520
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBeNum 
         Height          =   210
         Left            =   1560
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   2520
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFakt1 
         Height          =   210
         Left            =   1560
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   2040
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHauB2 
         Height          =   210
         Left            =   1080
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   2040
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFakt3 
         Height          =   210
         Left            =   600
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   2280
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFakt2 
         Height          =   210
         Left            =   120
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   2280
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFakt5 
         Height          =   210
         Left            =   1560
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   2280
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFakt4 
         Height          =   210
         Left            =   1080
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   2280
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtGesZu 
         Height          =   210
         Left            =   120
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   1800
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtGesBr 
         Height          =   210
         Left            =   600
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   1800
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHeiP1 
         Height          =   210
         Left            =   1080
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   1800
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHeiP2 
         Height          =   210
         Left            =   1560
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   1800
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHauB1 
         Height          =   210
         Left            =   600
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   2040
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtWeGel 
         Height          =   210
         Left            =   120
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   2040
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtSumme 
         Height          =   210
         Left            =   1560
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   1560
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAnOrt 
         Height          =   210
         Left            =   600
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1080
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbLeiTe 
         Height          =   315
         Left            =   3480
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
         _Version        =   1048579
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.CheckBox chkUnfAr 
         Height          =   195
         Left            =   2160
         TabIndex        =   142
         Top             =   360
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPaNam 
         Height          =   210
         Left            =   120
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   1080
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPaAdr 
         Height          =   210
         Left            =   120
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   1320
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPaGeb 
         Height          =   210
         Left            =   120
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   1560
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatum 
         Height          =   210
         Left            =   600
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtUeber 
         Height          =   210
         Left            =   600
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDaNeu 
         Height          =   210
         Left            =   1080
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtKasse 
         Height          =   210
         Left            =   1080
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   1320
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtVerNr 
         Height          =   210
         Left            =   1080
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   1560
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtGulti 
         Height          =   210
         Left            =   1560
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtKasNr 
         Height          =   210
         Left            =   1560
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtStatu 
         Height          =   210
         Left            =   1560
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   1080
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIKNum 
         Height          =   210
         Left            =   1560
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   1320
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkSonst 
         Height          =   195
         Left            =   2400
         TabIndex        =   155
         Top             =   360
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkGebPf 
         Height          =   195
         Left            =   2640
         TabIndex        =   156
         Top             =   360
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAuId3 
         Height          =   195
         Left            =   2880
         TabIndex        =   157
         Top             =   360
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkHaus1 
         Height          =   195
         Left            =   3120
         TabIndex        =   158
         Top             =   360
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkErstb 
         Height          =   195
         Left            =   2160
         TabIndex        =   159
         Top             =   600
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkSchul 
         Height          =   195
         Left            =   2400
         TabIndex        =   160
         Top             =   600
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBVGFe 
         Height          =   195
         Left            =   2640
         TabIndex        =   161
         Top             =   600
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkEWRCH 
         Height          =   195
         Left            =   2880
         TabIndex        =   162
         Top             =   600
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkHaus2 
         Height          =   195
         Left            =   3120
         TabIndex        =   163
         Top             =   600
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFolge 
         Height          =   195
         Left            =   2160
         TabIndex        =   164
         Top             =   840
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkSchwi 
         Height          =   195
         Left            =   2400
         TabIndex        =   165
         Top             =   840
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAuId1 
         Height          =   195
         Left            =   2640
         TabIndex        =   166
         Top             =   840
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkGrupp 
         Height          =   195
         Left            =   2880
         TabIndex        =   167
         Top             =   840
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkThBe1 
         Height          =   195
         Left            =   3120
         TabIndex        =   168
         Top             =   840
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkDurch 
         Height          =   195
         Left            =   2160
         TabIndex        =   169
         Top             =   1080
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkUnfal 
         Height          =   195
         Left            =   2400
         TabIndex        =   170
         Top             =   1080
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAuId2 
         Height          =   195
         Left            =   2640
         TabIndex        =   171
         Top             =   1080
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkRegel 
         Height          =   195
         Left            =   2880
         TabIndex        =   172
         Top             =   1080
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkThBe2 
         Height          =   195
         Left            =   3120
         TabIndex        =   173
         Top             =   1080
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkArGeb 
         Height          =   195
         Left            =   2160
         TabIndex        =   174
         Top             =   1320
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkGebFr 
         Height          =   195
         Left            =   2400
         TabIndex        =   175
         Top             =   1320
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAbsAd 
         Height          =   210
         Left            =   120
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtZeiVo 
         Height          =   210
         Left            =   1080
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtZeiBi 
         Height          =   210
         Left            =   1080
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   1080
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTxKon 
         Height          =   210
         Left            =   2040
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   370
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         Appearance      =   1
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAbsen 
         Height          =   225
         Left            =   1680
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   0
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4762
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Stempel nicht ausdrucken"
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
      End
      Begin VB.PictureBox picBild4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         Height          =   800
         Left            =   3600
         ScaleHeight     =   795
         ScaleWidth      =   795
         TabIndex        =   107
         Top             =   840
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.PictureBox picBild3 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         Height          =   800
         Left            =   3600
         ScaleHeight     =   795
         ScaleWidth      =   795
         TabIndex        =   108
         Top             =   1680
         Width           =   800
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   4500
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   405
      Left            =   15
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   10035
      _Version        =   1048579
      _ExtentX        =   17701
      _ExtentY        =   714
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.UpDown UpDown1 
         Height          =   350
         Left            =   1290
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   60
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu1"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.ComboBox cmbTypen 
         Height          =   315
         Left            =   1940
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   60
         Width           =   760
         _Version        =   1048579
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbZiffe 
         Height          =   315
         Left            =   2740
         TabIndex        =   54
         Top             =   60
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         DropDownItemCount=   20
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu4 
         Height          =   780
         Left            =   9200
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   780
         _Version        =   1048579
         _ExtentX        =   1376
         _ExtentY        =   1376
         _StockProps     =   64
         Show3DBorder    =   2
         VisualTheme     =   0
      End
      Begin XtremeSuiteControls.ComboBox cmbBezei 
         Height          =   315
         Left            =   3980
         TabIndex        =   55
         Top             =   60
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         AutoComplete    =   -1  'True
         DropDownItemCount=   20
      End
      Begin XtremeSuiteControls.FlatEdit txtEinze 
         Height          =   350
         Left            =   7000
         TabIndex        =   58
         Top             =   60
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   600
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtAnzal 
         Height          =   350
         Left            =   5700
         TabIndex        =   56
         Top             =   60
         Width           =   500
         _Version        =   1048579
         _ExtentX        =   882
         _ExtentY        =   600
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMulti 
         Height          =   350
         Left            =   6300
         TabIndex        =   57
         Top             =   60
         Width           =   600
         _Version        =   1048579
         _ExtentX        =   1058
         _ExtentY        =   600
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   1560
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   60
         Width           =   340
         _Version        =   1048579
         _ExtentX        =   600
         _ExtentY        =   600
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   40
         TabIndex        =   50
         Top             =   60
         Width           =   1240
         _Version        =   1048579
         _ExtentX        =   2187
         _ExtentY        =   600
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   8000
         TabIndex        =   59
         Top             =   60
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
   Begin XtremeShortcutBar.ShortcutCaption schCapt2 
      Height          =   375
      Left            =   10800
      TabIndex        =   181
      Top             =   6500
      Visible         =   0   'False
      Width           =   800
      _Version        =   1048579
      _ExtentX        =   1411
      _ExtentY        =   661
      _StockProps     =   14
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
      ForeColor       =   16777215
   End
   Begin XtremeCommandBars.CommandBars comBar01 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgManag 
      Left            =   480
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":689A
   End
   Begin XtremeSuiteControls.PopupControl popCont3 
      Left            =   1815
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.PopupControl popCont2 
      Left            =   1335
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.PopupControl popCont1 
      Left            =   855
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.TaskDialog tskDialo 
      Left            =   2775
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
   Begin XtremeSuiteControls.CommonDialog comDialo 
      Left            =   3135
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   3495
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dcpDoc01 
      Left            =   2280
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PiR14 As VB.PictureBox
Private DocPa As XtremeDockingPane.DockingPane
Private Labl9 As XtremeSuiteControls.Label
Private TrLi1 As XtremeSuiteControls.TreeView
Private TrLi2 As XtremeSuiteControls.TreeView
Private TrLi3 As XtremeSuiteControls.TreeView
Private TrLi5 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private TxDe3 As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private TxDe6 As XtremeSuiteControls.FlatEdit
Private TxDe8 As XtremeSuiteControls.FlatEdit
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmZif As XtremeSuiteControls.ComboBox
Private CmBez As XtremeSuiteControls.ComboBox
Private CmbMo As XtremeSuiteControls.ComboBox
Private CmbQu As XtremeSuiteControls.ComboBox
Private CmbJa As XtremeSuiteControls.ComboBox
Private CmbJv As XtremeSuiteControls.ComboBox
Private TsDia As XtremeSuiteControls.TaskDialog
Private TabCo As XtremeSuiteControls.TabControl
Private CoDia As XtremeSuiteControls.CommonDialog
Private LiFld As FolderViewControl.FolderView
Private LiNod As FolderViewControl.TreeNode
Private LiFi1 As FileViewControl.FileView
Private LiFi2 As FileViewControl.FileView
Private LiFit As FileViewControl.ListItem
Private LiVw1 As XtremeSuiteControls.ListView
Private LiVw2 As XtremeSuiteControls.ListView
Private LiVw3 As XtremeSuiteControls.ListView
Private LiVw4 As XtremeSuiteControls.ListView
Private LiVw5 As XtremeSuiteControls.ListView
Private LiIts As XtremeSuiteControls.ListViewItems
Private LiItm As XtremeSuiteControls.ListViewItem
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private Rahm5 As XtremeSuiteControls.GroupBox
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private OpZei As XtremeSuiteControls.RadioButton
Private PrGr1 As XtremePropertyGrid.PropertyGrid
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmBuT As XtremeCommandBars.CommandBarButton
Private CmPgs As XtremeCommandBars.StatusBarProgressPane
Private PoItm As XtremeSuiteControls.PopupControlItem
Private MoKal As XtremeCalendarControl.DatePicker
Private DaPi3 As XtremeCalendarControl.DatePicker
Private DaPi4 As XtremeCalendarControl.DatePicker
Private DaPi6 As XtremeCalendarControl.DatePicker
Private DaPi7 As XtremeCalendarControl.DatePicker
Private CaCol As XtremeCalendarControl.CalendarControl
Private CaThe As XtremeCalendarControl.CalendarThemeOffice2007
Private CaThI As XtremeCalendarControl.CalendarThemeImageList
Private CelPa As XtremeCalendarControl.CalendarThemeDayViewCellParams
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
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private TxCoN As Tx4oleLib.TXTextControl
Private TxRu1 As Tx4oleLib.TXRuler
Private TxRu2 As Tx4oleLib.TXRuler
Private ChCon As XtremeChartControl.ChartControl
Private ChCnt As XtremeChartControl.ChartContent
Private ChLeg As XtremeChartControl.ChartLegend
Private ChSrs As XtremeChartControl.ChartSeries
Private ChDia As XtremeChartControl.ChartDiagram2D
Private ChBar As XtremeChartControl.ChartBarSeriesStyle
Private ChPie As XtremeChartControl.ChartPieSeriesStyle
Private ChPyr As XtremeChartControl.ChartPieSeriesStyle
Private ChLin As XtremeChartControl.ChartSplineSeriesStyle
Private ChAre As XtremeChartControl.ChartStackedSplineAreaSeriesStyle

Private WithEvents CmSta As XtremeCommandBars.StatusBar
Attribute CmSta.VB_VarHelpID = -1
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

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
Private Const EM_GETLINECOUNT = &HBA
Private Const GWL_WNDPROC As Long = (-4&)
Private Const WM_DESTROY As Long = &H2&
Private Const WM_NCLBUTTONDOWN As Long = &HA1&
Private Const WM_NCMOUSEMOVE As Long = &HA0&
Private Const HTMINBUTTON As Long = 8&
Private Const HTREDUCE As Long = HTMINBUTTON
Private Const HTMAXBUTTON As Long = 9&
Private Const HTZOOM As Long = HTMAXBUTTON
Private Const MF_BYPOSITION As Long = &H400&

Private MauDo As Boolean
Private KalWa As Integer
Private TxPhr As String

Private clFil As clsFile
Private clWor As clsWord
Private clAnw As clsAnwend
Private clFen As clsFenster
Private clDru As clsDruck
Private clNet As clsNetz

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub bksButt01_Click()
    SAdre 1
End Sub


Private Sub bksButt02_Click()
    GlAdU = 1
    frmAdrSuch.Show vbModal
End Sub


Private Sub bksButt03_Click()
    SReDi 1
End Sub


Private Sub bksButt04_Click()
On Error Resume Next
  
GlBu3 = RibTab_Rechnungen
STaSe ShoCut_Finanz, RibTab_Rechnungen

End Sub
Private Sub bksButt05_Click()
    SAnEd True
End Sub
Private Sub bksButt06_Click()
    GlAdU = 2
    frmAdrSuch.Show vbModal
End Sub
Private Sub bksButt07_Click()
    SReDi 3
End Sub
Private Sub bksButt08_Click()
    STerm True
End Sub


Private Sub bksButt09_Click()
On Error Resume Next
    
GlBu7 = RibTab_Kat_Ketten
STaSe ShoCut_Katalog, RibTab_Kat_Ketten

End Sub


Private Sub bksButt10_Click()
On Error Resume Next

GlBu3 = RibTab_Mahnwesen
STaSe ShoCut_Finanz, RibTab_Mahnwesen

End Sub


Private Sub bksButt11_Click()
On Error Resume Next

GlBu3 = RibTab_Buchungen
STaSe ShoCut_Finanz, RibTab_Buchungen

GlNeB = True 'neue Buchung

frmBuEdit.Show

End Sub
Private Sub bksButt12_Click()
On Error Resume Next
    
GlBu3 = RibTab_Statistik
STaSe ShoCut_Finanz, RibTab_Statistik

End Sub
Private Sub bksButt13_Click()
    STxDi
End Sub
Private Sub bksButt14_Click()
On Error Resume Next

GlBu6 = RibTab_Tex_Email
STaSe ShoCut_Texte, RibTab_Tex_Email
SMaNe

End Sub

Private Sub bksButt15_Click()
    frmChipcard.Show vbModal
End Sub
Private Sub bksButt16_Click()
    frmFormular.Show vbModal
End Sub

Private Sub bksButt17_Click()
On Error Resume Next

Dim FS As Form

If GlTSe > 0 Then
    Set FM = frmMain
    Set FS = frmTSEInit
    
    FS.Show vbModeless, FM
    DoEvents

    Select Case GlTSe
    Case 1: TSESwi
    Case 2: TSEWeb
    End Select
End If

End Sub
Private Sub bksButt18_Click()
On Error Resume Next

GlBu3 = RibTab_HomeBanki
STaSe ShoCut_Finanz, RibTab_HomeBanki

End Sub
Private Sub bksButt19_Click()
On Error Resume Next

GlBu5 = RibTab_LabBerichte
STaSe ShoCut_Labor, RibTab_LabBerichte

End Sub

Private Sub bksButt20_Click()
    SLize 3
End Sub
Private Sub btnDatu2_Click()
    If GlAkt = False Then
        KalWa = 2
        FKale
    End If
End Sub

Private Sub btnDatu3_Click()
    If GlAkt = False Then
        KalWa = 3
        FKale
    End If
End Sub


Private Sub calCont1_BeforeDrawThemeObject(ByVal eObjType As XtremeCalendarControl.CalendarBeforeDrawThemeObject, ByVal DrawParams As Variant)
On Error Resume Next

Dim SpZe1 As Date
Dim SpZe2 As Date
Dim SpZe3 As Date
Dim SpZe4 As Date
Dim StaZe As Date
Dim MitNr As Long
Dim ManNr As Long
Dim ColWO As Long
Dim ColNW As Long
Dim ColKO As Long
Dim ColL1 As Long
Dim ColL2 As Long
Dim TmSpr As String
Dim GrIdx As Integer
Dim SuIdx As Integer
Dim WoTag As Integer
Dim AktZa As Integer
Dim SpEn1 As Boolean
Dim SpEn2 As Boolean
Dim FaZei As Boolean

Set FM = frmMain
Set CaCol = FM.calCont1
Set CaThe = CaCol.Theme

FaZei = GlMZe 'Sprechzeiten des Mandanten im Terminkalender zeigen

If GlBut = RibTab_Ter_Raeume Then
    If GlMZe = True Then 'Sollen die Sprechzeiten grundsätzlich angezeigt werden
        If GlTRF = True Then
            FaZei = False
        End If
    End If
End If

If GlDat = True Then
    If GlCaS > 11 Then 'Kalenderfilterinhalt
        SuIdx = 1
    Else
        SuIdx = GlCaS
    End If

    If eObjType = xtpCalendarBeforeDraw_DayViewCell Then
        Set CelPa = DrawParams
        GrIdx = CelPa.ViewGroup.GroupIndex + 1 'Kalenderresourcenindex
        ColWO = CaThe.DayView.Day.Group.Cell.WorkCell.BackgroundColor
        ColNW = CaThe.DayView.Day.Group.Cell.NonWorkCell.BackgroundColor
        ColL1 = CaThe.DayView.Day.Group.Cell.WorkCell.BorderBottomHourColor
        ColL2 = CaThe.DayView.Day.Group.Cell.WorkCell.BorderBottomInHourColor
        WoTag = Weekday(DateValue(CelPa.BeginTime), vbMonday)
        StaZe = TimeValue(CelPa.BeginTime)

        If GrIdx <= 18 Then 'Kalenderresourcenindex WICHTIG!
            Select Case GlBut
            Case RibTab_Ter_Raeume:
                If GlSMa > UBound(GlMaT) Then
                    TmSpr = GlMaT(1, 6) 'Hier unabhängig der Einstellung immer Sprechzeiten der Mandanten laden
                Else
                    TmSpr = GlMaT(GlSMa, 6)
                End If
                If GlMFa = True Then 'Farbliche Kennzeichnung der Mandatenten
                    ColNW = GlTmH(GrIdx, 2)
                    ColKO = GlTmH(GrIdx, 3)
                    ColL1 = GlTmH(GrIdx, 4)
                    ColL2 = GlTmH(GrIdx, 5)
                End If
                
            Case RibTab_Ter_Mitarb:
            
                If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
                    If GlSpT = False Then 'starre oder flexible Sprechzeiten verwenden
                        If GrIdx <= UBound(GlMiT) Then 'aktive Mitarbeiter + Terminspalte
                            TmSpr = GlMiT(GrIdx, 6)
                        Else
                            If GlSmI > UBound(GlMiT) Then
                                TmSpr = GlMiT(1, 6)
                            Else
                                TmSpr = GlMiT(GlSmI, 6)
                            End If
                        End If
                    Else
                        If GrIdx <= UBound(GlMiT) Then 'aktive Mitarbeiter + Terminspalte
                            MitNr = GlMiT(GrIdx, 2)
                        Else
                            If GlSmI > UBound(GlMiT) Then
                                MitNr = GlMiT(1, 2)
                            Else
                                MitNr = GlMiT(GlSmI, 2)
                            End If
                        End If
                        If UBound(GlSpr) > 0 Then 'gespeicherte Sprechzeiten
                            For AktZa = 1 To UBound(GlSpr)
                                If MitNr = GlSpr(AktZa, 1) Then
                                    If CDate(GlSpr(AktZa, 4)) <= GlDFi Then
                                        TmSpr = GlSpr(AktZa, 2)
                                    End If
                                    If CDate(GlSpr(AktZa, 4)) > GlDFi Then
                                        Exit For
                                    End If
                                End If
                            Next AktZa
                            If TmSpr = vbNullString Then
                                If GrIdx <= UBound(GlMiT) Then 'Aktive Mitarbeiter + Terminspalte
                                    TmSpr = GlMiT(GrIdx, 6)
                                Else
                                    If GlSmI > UBound(GlMiT) Then
                                        TmSpr = GlMiT(1, 6)
                                    Else
                                        TmSpr = GlMiT(GlSmI, 6)
                                    End If
                                End If
                            End If
                        Else
                            If GrIdx <= UBound(GlMiT) Then 'Aktive Mitarbeiter + Terminspalte
                                TmSpr = GlMiT(GrIdx, 6)
                            Else
                                If GlSmI > UBound(GlMiT) Then
                                    TmSpr = GlMiT(1, 6)
                                Else
                                    TmSpr = GlMiT(GlSmI, 6)
                                End If
                            End If
                        End If
                    End If
                Else 'Mandantenplan
                    If GlSpT = False Then 'starre oder flexible Sprechzeiten verwenden
                        If GrIdx <= UBound(GlMaT) Then
                            TmSpr = GlMaT(GrIdx, 6)
                        Else
                            If GlSMa > UBound(GlMaT) Then
                                TmSpr = GlMaT(1, 6)
                            Else
                                TmSpr = GlMaT(GlSMa, 6)
                            End If
                        End If
                    Else
                        If GrIdx <= UBound(GlMiT) Then 'aktive Mitarbeiter + Terminspalte
                            ManNr = GlMaT(GrIdx, 2)
                        Else
                            If GlSMa > UBound(GlMaT) Then
                                ManNr = GlMaT(1, 2)
                            Else
                                ManNr = GlMaT(GlSMa, 2)
                            End If
                        End If
                        If UBound(GlSpr) > 0 Then
                            For AktZa = 1 To UBound(GlSpr)
                                If ManNr = GlSpr(AktZa, 1) Then
                                    If CDate(GlSpr(AktZa, 4)) <= GlDFi Then
                                        TmSpr = GlSpr(AktZa, 2)
                                    End If
                                    If CDate(GlSpr(AktZa, 4)) > GlDFi Then
                                        Exit For
                                    End If
                                End If
                            Next AktZa
                            If TmSpr = vbNullString Then
                                If GrIdx <= UBound(GlMaT) Then 'Aktive Mitarbeiter + Terminspalte
                                    TmSpr = GlMaT(GrIdx, 6)
                                Else
                                    If GlSMa > UBound(GlMaT) Then
                                        TmSpr = GlMaT(1, 6)
                                    Else
                                        TmSpr = GlMaT(GlSMa, 6)
                                    End If
                                End If
                            End If
                        Else
                            If GrIdx <= UBound(GlMiT) Then 'Aktive Mitarbeiter + Terminspalte
                                TmSpr = GlMaT(GrIdx, 6)
                            Else
                                If GlSMa > UBound(GlMaT) Then
                                    TmSpr = GlMaT(1, 6)
                                Else
                                    TmSpr = GlMaT(GlSMa, 6)
                                End If
                            End If
                        End If
                    End If
                End If
                If GlMFa = True Then 'Farbliche Kennzeichnung der Mandatenten / Mitarbeiter
                    ColNW = GlTmH(GrIdx, 2) 'Terminhintergrund
                    ColKO = GlTmH(GrIdx, 3)
                    ColL1 = GlTmH(GrIdx, 4)
                    ColL2 = GlTmH(GrIdx, 5)
                End If
                
            Case RibTab_Ter_Kalend:

                If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
                    If GlCaF = 3 Then 'Kalenderfilterauswahl auf Mitarbeiter
                        If GlSpT = False Then 'starre oder flexible Sprechzeiten verwenden
                            If UBound(GlMiT) > 0 Then
                                TmSpr = GlMiT(GlCaS, 6)
                            Else
                                TmSpr = GlSZe 'Sprechzietenstring
                            End If
                        Else
                            If UBound(GlMiT) > 0 Then
                                MitNr = GlMiT(GlCaS, 2)
                            Else
                                MitNr = 0
                            End If
                            If UBound(GlSpr) > 0 Then
                                For AktZa = 1 To UBound(GlSpr)
                                    If MitNr = GlSpr(AktZa, 1) Then
                                        If CDate(GlSpr(AktZa, 4)) <= GlDFi Then
                                            TmSpr = GlSpr(AktZa, 2)
                                        End If
                                        If CDate(GlSpr(AktZa, 4)) > GlDFi Then
                                            Exit For
                                        End If
                                    End If
                                Next AktZa
                                If TmSpr = vbNullString Then
                                    If UBound(GlMiT) > 0 Then
                                        TmSpr = GlMiT(GlCaS, 6)
                                    Else
                                        TmSpr = GlSZe 'Sprechzietenstring
                                    End If
                                End If
                            Else
                                If UBound(GlMiT) > 0 Then
                                    TmSpr = GlMiT(GlCaS, 6)
                                Else
                                    TmSpr = GlSZe 'Sprechzietenstring
                                End If
                            End If
                        End If
                    Else
                        If GlSpT = False Then 'starre oder flexible Sprechzeiten verwenden
                            If UBound(GlMiT) > 0 Then
                                If GlSmI > UBound(GlMiT) Then
                                    TmSpr = GlMiT(1, 6)
                                Else
                                    TmSpr = GlMiT(GlSmI, 6)
                                End If
                            Else
                                TmSpr = GlSZe 'Sprechzietenstring
                            End If
                        Else
                            If UBound(GlMiT) > 0 Then
                                If GlSmI > UBound(GlMiT) Then
                                    MitNr = GlMiT(1, 2)
                                Else
                                    MitNr = GlMiT(GlSmI, 2)
                                End If
                            Else
                                MitNr = 0
                            End If
                            If UBound(GlSpr) > 0 Then
                                For AktZa = 1 To UBound(GlSpr)
                                    If MitNr = GlSpr(AktZa, 1) Then
                                        If CDate(GlSpr(AktZa, 4)) <= GlDFi Then
                                            TmSpr = GlSpr(AktZa, 2)
                                        End If
                                        If CDate(GlSpr(AktZa, 4)) > GlDFi Then
                                            Exit For
                                        End If
                                    End If
                                Next AktZa
                                If TmSpr = vbNullString Then
                                    If UBound(GlMiT) > 0 Then
                                        If GlSmI > UBound(GlMiT) Then
                                            TmSpr = GlMiT(1, 6)
                                        Else
                                            TmSpr = GlMiT(GlSmI, 6)
                                        End If
                                    Else
                                        TmSpr = GlSZe 'Sprechzietenstring
                                    End If
                                End If
                            Else
                                If UBound(GlMiT) > 0 Then
                                    If GlSmI > UBound(GlMiT) Then
                                        TmSpr = GlMiT(1, 6)
                                    Else
                                        TmSpr = GlMiT(GlSmI, 6)
                                    End If
                                Else
                                    TmSpr = GlSZe 'Sprechzietenstring
                                End If
                            End If
                        End If
                    End If
                    Select Case GlCaF 'Kalenderfilterauswahl
                    Case 1: 'Alle Termine
                            If GlMFa = True Then
                                ColNW = GlKaB
                                ColKO = GlTmH(1, 3)
                                ColL1 = GlTmH(1, 4)
                                ColL2 = GlTmH(1, 5)
                            End If
                    Case 2: 'Raumbelegung
                            If GlMFa = True Then
                                ColNW = GlTmH(SuIdx, 2)
                                ColKO = GlTmH(SuIdx, 3)
                                ColL1 = GlTmH(SuIdx, 4)
                                ColL2 = GlTmH(SuIdx, 5)
                            End If
                    Case 3: 'Mitarbeiter
                            If GlMFa = True Then
                                ColNW = GlTmH(SuIdx, 2)
                                ColKO = GlTmH(SuIdx, 3)
                                ColL1 = GlTmH(SuIdx, 4)
                                ColL2 = GlTmH(SuIdx, 5)
                            End If
                    End Select
                    
                Else 'Mandantenplan

                    If GlCaF = 3 Then 'Kalenderfilterauswahl auf Mandanten
                        If GlSpT = False Then 'starre oder flexible Sprechzeiten verwenden
                            If UBound(GlMaT) > 0 Then
                                TmSpr = GlMaT(GlCaS, 6)
                            Else
                                TmSpr = GlSZe 'Sprechzietenstring
                            End If
                        Else
                            If UBound(GlMaT) > 0 Then
                                ManNr = GlMaT(GlCaS, 2)
                            Else
                                ManNr = 0
                            End If
                            If UBound(GlSpr) > 0 Then
                                For AktZa = 1 To UBound(GlSpr)
                                    If ManNr = GlSpr(AktZa, 1) Then
                                        If CDate(GlSpr(AktZa, 4)) <= GlDFi Then
                                            TmSpr = GlSpr(AktZa, 2)
                                        End If
                                        If CDate(GlSpr(AktZa, 4)) > GlDFi Then
                                            Exit For
                                        End If
                                    End If
                                Next AktZa
                                If TmSpr = vbNullString Then
                                    If UBound(GlMaT) > 0 Then
                                        TmSpr = GlMaT(GlCaS, 6)
                                    Else
                                        TmSpr = GlSZe 'Sprechzietenstring
                                    End If
                                End If
                            Else
                                If UBound(GlMaT) > 0 Then
                                    TmSpr = GlMiT(GlCaS, 6)
                                Else
                                    TmSpr = GlSZe 'Sprechzietenstring
                                End If
                            End If
                        End If
                    Else
                        If GlSpT = False Then 'starre oder flexible Sprechzeiten verwenden
                            If UBound(GlMaT) > 0 Then
                                If GlSMa > UBound(GlMaT) Then
                                    TmSpr = GlMaT(1, 6)
                                Else
                                    TmSpr = GlMaT(GlSMa, 6)
                                End If
                            Else
                                TmSpr = GlSZe 'Sprechzietenstring
                            End If
                        Else
                            If UBound(GlMaT) > 0 Then
                                If GlSMa > UBound(GlMaT) Then
                                    ManNr = GlMaT(1, 2)
                                Else
                                    ManNr = GlMaT(GlSMa, 2)
                                End If
                            Else
                                ManNr = 0
                            End If
                            If UBound(GlSpr) > 0 Then
                                For AktZa = 1 To UBound(GlSpr)
                                    If MitNr = GlSpr(AktZa, 1) Then
                                        If CDate(GlSpr(AktZa, 4)) <= GlDFi Then
                                            TmSpr = GlSpr(AktZa, 2)
                                        End If
                                        If CDate(GlSpr(AktZa, 4)) > GlDFi Then
                                            Exit For
                                        End If
                                    End If
                                Next AktZa
                                If TmSpr = vbNullString Then
                                    If UBound(GlMaT) > 0 Then
                                        If GlSMa > UBound(GlMaT) Then
                                            TmSpr = GlMaT(1, 6)
                                        Else
                                            TmSpr = GlMaT(GlSMa, 6)
                                        End If
                                    Else
                                        TmSpr = GlSZe 'Sprechzietenstring
                                    End If
                                End If
                            Else
                                If UBound(GlMaT) > 0 Then
                                    If GlSMa > UBound(GlMaT) Then
                                        TmSpr = GlMaT(1, 6)
                                    Else
                                        TmSpr = GlMaT(GlSMa, 6)
                                    End If
                                Else
                                    TmSpr = GlSZe 'Sprechzietenstring
                                End If
                            End If
                        End If
                    End If
                    Select Case GlCaF 'Kalenderfilterauswahl
                    Case 1: 'Alle Termine
                            If GlMFa = True Then
                                ColNW = GlKaB
                                ColKO = GlTmH(1, 3)
                                ColL1 = GlTmH(1, 4)
                                ColL2 = GlTmH(1, 5)
                            End If
                    Case 2: 'Raumbelegung
                            If GlMFa = True Then
                                ColNW = GlTmH(SuIdx, 2)
                                ColKO = GlTmH(SuIdx, 3)
                                ColL1 = GlTmH(SuIdx, 4)
                                ColL2 = GlTmH(SuIdx, 5)
                            End If
                    Case 3: 'Mitarbeiter
                            If GlMFa = True Then
                                ColNW = GlTmH(SuIdx, 2)
                                ColKO = GlTmH(SuIdx, 3)
                                ColL1 = GlTmH(SuIdx, 4)
                                ColL2 = GlTmH(SuIdx, 5)
                            End If
                    End Select
                End If
            End Select

            With CaThe.DayView.Day.Group
                If GlBut = RibTab_Ter_Kalend Then
                    If GlCaF > 1 Then 'Kalenderfilterauswahl
                        If GlMFa = True Then 'Farbliche Kennzeichnung der Mandatenten
                            .AllDayEvents.BackgroundColor = ColKO
                        End If
                        .Cell.WorkCell.BorderBottomHourColor = ColL1
                        .Cell.WorkCell.BorderBottomInHourColor = ColNW
                        .Cell.NonWorkCell.BorderBottomHourColor = ColL1
                        .Cell.NonWorkCell.BorderBottomInHourColor = ColNW
                    Else
                        If GlSty = 8 Then 'Office 2013
                            .AllDayEvents.BackgroundColor = GlKaB
                        End If
                    End If
                Else
                    If GlMFa = True Then
                        .AllDayEvents.BackgroundColor = ColKO
                    End If
                    .Cell.WorkCell.BorderBottomHourColor = ColL1
                    .Cell.WorkCell.BorderBottomInHourColor = ColNW
                    .Cell.NonWorkCell.BorderBottomHourColor = ColL1
                    .Cell.NonWorkCell.BorderBottomInHourColor = ColNW
                End If

                If FaZei = True Then 'Sprechzeiten zeigen
                    Select Case WoTag
                    Case 1: 'Montag
                        SpZe1 = TimeValue(Mid$(TmSpr, 2, 5))
                        SpZe2 = TimeValue(Mid$(TmSpr, 8, 5))
                        SpZe3 = TimeValue(Mid$(TmSpr, 14, 5))
                        SpZe4 = TimeValue(Mid$(TmSpr, 20, 5))
                        If Mid$(TmSpr, 1, 1) = "A" Then SpEn1 = True
                        If Mid$(TmSpr, 13, 1) = "A" Then SpEn2 = True
                        If SpEn1 = True Then
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        End If
                        If StaZe > SpZe2 And StaZe < SpZe3 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                        End If
                        If SpEn2 = False Then
                            If StaZe > SpZe3 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe > SpZe3 And StaZe < SpZe4 Then
                                .Cell.WorkCell.BackgroundColor = ColWO
                                .Cell.NonWorkCell.BackgroundColor = ColWO
                            End If
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                    Case 2: 'Dienstag
                        SpZe1 = TimeValue(Mid$(TmSpr, 26, 5))
                        SpZe2 = TimeValue(Mid$(TmSpr, 32, 5))
                        SpZe3 = TimeValue(Mid$(TmSpr, 38, 5))
                        SpZe4 = TimeValue(Mid$(TmSpr, 44, 5))
                        If Mid$(TmSpr, 25, 1) = "A" Then SpEn1 = True
                        If Mid$(TmSpr, 37, 1) = "A" Then SpEn2 = True
                        If SpEn1 = True Then
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        End If
                        If StaZe > SpZe2 And StaZe < SpZe3 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                        End If
                        If SpEn2 = False Then
                            If StaZe > SpZe3 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe > SpZe3 And StaZe < SpZe4 Then
                                .Cell.WorkCell.BackgroundColor = ColWO
                                .Cell.NonWorkCell.BackgroundColor = ColWO
                            End If
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                    Case 3: 'Mittwoch
                        SpZe1 = TimeValue(Mid$(TmSpr, 50, 5))
                        SpZe2 = TimeValue(Mid$(TmSpr, 56, 5))
                        SpZe3 = TimeValue(Mid$(TmSpr, 62, 5))
                        SpZe4 = TimeValue(Mid$(TmSpr, 68, 5))
                        If Mid$(TmSpr, 49, 1) = "A" Then SpEn1 = True
                        If Mid$(TmSpr, 61, 1) = "A" Then SpEn2 = True
                        If SpEn1 = True Then
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        End If
                        If StaZe > SpZe2 And StaZe < SpZe3 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                        End If
                        If SpEn2 = False Then
                            If StaZe > SpZe3 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe > SpZe3 And StaZe < SpZe4 Then
                                .Cell.WorkCell.BackgroundColor = ColWO
                                .Cell.NonWorkCell.BackgroundColor = ColWO
                            End If
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                    Case 4: 'Donnerstag
                        SpZe1 = TimeValue(Mid$(TmSpr, 74, 5))
                        SpZe2 = TimeValue(Mid$(TmSpr, 80, 5))
                        SpZe3 = TimeValue(Mid$(TmSpr, 86, 5))
                        SpZe4 = TimeValue(Mid$(TmSpr, 92, 5))
                        If Mid$(TmSpr, 73, 1) = "A" Then SpEn1 = True
                        If Mid$(TmSpr, 85, 1) = "A" Then SpEn2 = True
                        If SpEn1 = True Then
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        End If
                        If StaZe > SpZe2 And StaZe < SpZe3 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                        End If
                        If SpEn2 = False Then
                            If StaZe > SpZe3 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe > SpZe3 And StaZe < SpZe4 Then
                                .Cell.WorkCell.BackgroundColor = ColWO
                                .Cell.NonWorkCell.BackgroundColor = ColWO
                            End If
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                    Case 5: 'Freitag
                        SpZe1 = TimeValue(Mid$(TmSpr, 98, 5))
                        SpZe2 = TimeValue(Mid$(TmSpr, 104, 5))
                        SpZe3 = TimeValue(Mid$(TmSpr, 110, 5))
                        SpZe4 = TimeValue(Mid$(TmSpr, 116, 5))
                        If Mid$(TmSpr, 97, 1) = "A" Then SpEn1 = True
                        If Mid$(TmSpr, 109, 1) = "A" Then SpEn2 = True
                        If SpEn1 = True Then
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        End If
                        If StaZe > SpZe2 And StaZe < SpZe3 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                        End If
                        If SpEn2 = False Then
                            If StaZe > SpZe3 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe > SpZe3 And StaZe < SpZe4 Then
                                .Cell.WorkCell.BackgroundColor = ColWO
                                .Cell.NonWorkCell.BackgroundColor = ColWO
                            End If
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                    Case 6: 'Samstag
                        SpZe1 = TimeValue(Mid$(TmSpr, 122, 5))
                        SpZe2 = TimeValue(Mid$(TmSpr, 128, 5))
                        SpZe3 = TimeValue(Mid$(TmSpr, 134, 5))
                        SpZe4 = TimeValue(Mid$(TmSpr, 140, 5))
                        If Mid$(TmSpr, 121, 1) = "A" Then SpEn1 = True
                        If Mid$(TmSpr, 133, 1) = "A" Then SpEn2 = True
                        If SpEn1 = True Then
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        End If
                        If StaZe > SpZe2 And StaZe < SpZe3 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                        End If
                        If SpEn2 = False Then
                            If StaZe > SpZe3 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe > SpZe3 And StaZe < SpZe4 Then
                                .Cell.WorkCell.BackgroundColor = ColWO
                                .Cell.NonWorkCell.BackgroundColor = ColWO
                            End If
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                    Case 7: 'Sonntag
                        SpZe1 = TimeValue(Mid$(TmSpr, 146, 5))
                        SpZe2 = TimeValue(Mid$(TmSpr, 152, 5))
                        SpZe3 = TimeValue(Mid$(TmSpr, 158, 5))
                        SpZe4 = TimeValue(Mid$(TmSpr, 164, 5))
                        If Mid$(TmSpr, 145, 1) = "A" Then SpEn1 = True
                        If Mid$(TmSpr, 157, 1) = "A" Then SpEn2 = True
                        If SpEn1 = True Then
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe < SpZe1 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            ElseIf StaZe < SpZe2 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        End If
                        If StaZe > SpZe2 And StaZe < SpZe3 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                        End If
                        If SpEn2 = False Then
                            If StaZe > SpZe3 Then
                                .Cell.WorkCell.BackgroundColor = ColNW
                                .Cell.NonWorkCell.BackgroundColor = ColNW
                            End If
                        Else
                            If StaZe > SpZe3 And StaZe < SpZe4 Then
                                .Cell.WorkCell.BackgroundColor = ColWO
                                .Cell.NonWorkCell.BackgroundColor = ColWO
                            End If
                        End If
                        If StaZe > SpZe4 Then
                            .Cell.WorkCell.BackgroundColor = ColNW
                            .Cell.NonWorkCell.BackgroundColor = ColNW
                        End If
                    End Select
                End If
            End With
            
        End If
    End If
End If

End Sub
Private Sub calCont1_BeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, bCancelOperation As Boolean)
On Error Resume Next

If GlTVe = False Then 'Terminverschiebung zulassen
    If OpParams.Operation = xtpCalendarEO_DragCopy Or xtpCalendarEO_DragMove Or xtpCalendarEO_DragResizeBegin Or xtpCalendarEO_DragResizeEnd Then
        bCancelOperation = True
    End If
End If
If OpParams.Operation = xtpCalendarEO_DeleteSelectedEvents Or xtpCalendarEO_Unknown Then
    bCancelOperation = True
End If

End Sub
Private Sub calCont1_PrePopulate(ByVal ViewGroup As XtremeCalendarControl.CalendarViewGroup, ByVal Events As XtremeCalendarControl.CalendarEvents)
On Error Resume Next

Dim CaEvt As XtremeCalendarControl.CalendarEvent

For Each CaEvt In Events
    If CaEvt.Reminder = True Then CaEvt.CustomIcons.Add xtpCalendarEventIconIDReminder
    If CaEvt.PrivateFlag = True Then CaEvt.CustomIcons.Add xtpCalendarEventIconIDPrivate
    If CaEvt.MeetingFlag = True Then CaEvt.CustomIcons.Add xtpCalendarEventIconIDMeeting
    If CaEvt.RecurrenceState = xtpCalendarRecurrenceOccurrence Then CaEvt.CustomIcons.Add xtpCalendarEventIconIDOccurrence
    If CaEvt.RecurrenceState = xtpCalendarRecurrenceException Then CaEvt.CustomIcons.Add xtpCalendarEventIconIDException
Next CaEvt

End Sub

Private Sub calCont1_SelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
    If GlTZe = True Then 'Terminzeitanpassung
        TeAkt
    End If
End Sub
Private Sub calCont1_ViewChanged()
On Error Resume Next

Dim DatSt As Date
Dim DatEn As Date
Dim AnzTa As Integer

Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents

Set CaCol = FM.calCont1

Set ViEvs = CaCol.ActiveView.GetSelectedEvents

If GlAkt = False Then
    If ViEvs.Count > 0 Then
        For Each ViEvt In ViEvs
            If ViEvt.Selected = True Then
                ViEvt.Selected = False
            End If
        Next ViEvt
    End If
    
    With CaCol
        AnzTa = .ActiveView.DaysCount - 1 'Anzahl angezeigter Tage
        Select Case .ViewType
        Case xtpCalendarDayView:
            DatSt = CDate(.ActiveView.Days(0).Date)
            DatEn = CDate(.ActiveView.Days(0).Date)
        Case xtpCalendarFullWeekView:
            DatSt = CDate(.ActiveView.Days(0).Date)
            DatEn = CDate(.ActiveView.Days(AnzTa).Date)
        Case xtpCalendarWorkWeekView:
            DatSt = CDate(.ActiveView.Days(0).Date)
            DatEn = CDate(.ActiveView.Days(AnzTa).Date)
        Case xtpCalendarWeekView:
            DatSt = CDate(.ActiveView.Days(0).Date)
            DatEn = CDate(.ActiveView.Days(AnzTa).Date)
        Case xtpCalendarMonthView:
            DatSt = CDate(.ActiveView.Days(0).Date)
            DatEn = CDate(.ActiveView.Days(AnzTa).Date)
        Case Else:
            DatSt = CDate(GlDFi)
            DatEn = CDate(GlDLa)
        End Select
    End With
    
    GlDFi = DatSt
    GlDLa = DatEn

    Screen.MousePointer = vbHourglass
    S_TeLi
    Screen.MousePointer = vbNormal
End If

End Sub
Private Sub chkExpa2_Click()
    If GlAkt = False Then STrSa 8
End Sub

Private Sub chkExpa3_Click()
    If GlAkt = False Then STrSa 8
End Sub


Private Sub chkExpan_Click()
    If GlAkt = False Then STrSa 8
End Sub
Private Sub chkJahre_Click()
    FDiZe
End Sub
Private Sub chkStor2_Click()
    If GlAkt = False Then
        If chkStor2.Value = xtpChecked Then
            GlSuT.SuIdx = 20 'zeige Stornierte
            SUpTe
        Else
            FSuAu
        End If
    End If
End Sub
Private Sub chkStorn_Click()
    If GlAkt = False Then
        STrSa 9
    End If
End Sub

Private Sub cmbJahre_Click()
    Set OpJah = Me.optZeit3
    OpJah.Value = True
    FDiZe
End Sub
Private Sub cmbMonat_Click()
    Set OpMon = Me.optZeit1
    OpMon.Value = True
    FDiZe
End Sub

Private Sub cmbQurta_Click()
    Set OpQua = Me.optZeit2
    OpQua.Value = True
    FDiZe
End Sub
Private Sub cmbVgJah_Click()
    FDiZe
End Sub
Private Sub CmSta_SliderPaneClick(ByVal Pane As XtremeCommandBars.StatusBarSliderPane, ByVal Command As XtremeCommandBars.XTPSliderCommand, ByVal Pos As Long)
On Error Resume Next

Dim FrZom As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set TxCoN = FM.TexCont1
Set CmBrs = FM.comBar01
Set CmSta = CmBrs.StatusBar

FrZom = GlZoT

If GlSta = False Then
    Select Case Command
    Case XTP_SB_LEFT: FrZom = 0
    Case XTP_SB_RIGHT: FrZom = 200
    Case XTP_SB_LINELEFT: FrZom = WinMax((Int(FrZom / 10) - 1) * 10, 0)
    Case XTP_SB_LINERIGHT: FrZom = WinMin((Int(FrZom / 10) + 1) * 10, 200)
    Case XTP_SB_THUMBTRACK: FrZom = Pos
    Case XTP_SB_PAGELEFT: FrZom = WinMax(FrZom - 20, 0)
    Case XTP_SB_PAGERIGHT: FrZom = WinMin(FrZom + 20, 200)
    End Select
    If (FrZom <> GlZoT) Then
        If FrZom > 10 Then
            GlZoT = FrZom
            Pane.Value = FrZom
            TxCoN.ZoomFactor = GlZoT
            CmSta.FindPane(Tex_Pa_ZoPan).Text = "Zoom: " & Format$(FrZom, "000") & "%"
            IniSetVal "GUI", "ZomTex", GlZoT
        End If
    End If
End If

End Sub
Private Sub CmSta_SwitchPaneClick(ByVal Pane As XtremeCommandBars.StatusBarSwitchPane, ByVal Switch As Long)
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSwi As XtremeCommandBars.StatusBarSwitchPane

Set FM = frmMain
Set TxCoN = FM.TexCont1
Set TxRu1 = FM.TexRule1
Set TxRu2 = FM.TexRule2
Set CmBrs = FM.comBar01
Set CmSta = CmBrs.StatusBar

If GlSta = False Then
    With CmSta
        Set CmSwi = .FindPane(Pane.id)
        Select Case Pane.id
        Case Tex_Pa_Layou:
            Select Case Switch
            Case IC16_AnsNor:
                GlViT = 0
                CmSwi.Checked = IC16_AnsNor
            Case IC16_AnsBre:
                GlViT = 2
                CmSwi.Checked = IC16_AnsBre
            Case IC16_AnsFli:
                GlViT = 3
                CmSwi.Checked = IC16_AnsFli
            End Select
            TxCoN.ViewMode = GlViT 'ViewMode Textverarbeitung
            IniSetVal "System", "ViewTx", "C" & GlViT
        Case Tex_Linial:
            GlLiT = Not GlLiT 'Lineal Textverarbeitung
            If GlLiT = True Then
                CmSwi.Checked = IC16_Ruler
            Else
                CmSwi.Checked = 0
            End If
            TxRu1.Visible = GlLiT
            TxRu2.Visible = GlLiT
            IniSetVal "Layout", "LinTex", GlLiT
            SPosi
        Case Tex_Pa_Progr:
            SRemo
        End Select
    End With
End If

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SwitchPaneClick " & Err.Number
Resume Next

End Sub
Private Sub dtpDatu3_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
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
Private Sub dtpDatu3_MonthChanged()

Dim DayFi As Date
Dim DayLa As Date

Set MoKal = Me.dtpDatu3

With MoKal
    DayFi = .FirstVisibleDay
    DayLa = .LastVisibleDay
End With

If GlAkt = False Then
    'S_AbTe DayFi, DayLa
End If

Set MoKal = Nothing

End Sub
Private Sub dtpDatu6_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSunday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu6_SelectionChanged()
    If GlAkt = False Then
        FDat6
    End If
End Sub

Private Sub dtpDatu7_SelectionChanged()
    FKaSe
End Sub

Private Sub filView1_Click()
    If GlAkt = False Then
        SBild
    End If
End Sub
Private Sub filView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo 1
        End If
    End If
End Sub
Private Sub filView1_OnAfterItemAdd(ByVal Item As FileViewControl.IListItem)
On Error Resume Next

Dim DaNam As String
Dim NeNam As String
Dim Lange As Integer
Dim Posit As Integer
 
Set FM = frmMain
Set LiFi1 = FM.filView1

DaNam = Item.DisplayName
Lange = Len(DaNam)

Posit = InStrRev(DaNam, "#_", Lange, 1)
If Posit > 0 Then
    NeNam = Mid$(DaNam, Posit + 2, Lange - (Posit + 1))
    Item.Text = NeNam
ElseIf Mid$(DaNam, 8, 1) = "_" Then
    NeNam = "Bilddokument"
    Item.Text = NeNam
Else
    NeNam = DaNam
End If

Item.SetColumnText "Dateiname", 3, NeNam

End Sub

Private Sub filView1_OnBeforeShellCommandExecute(ByVal CommandStr As String, Cancel As Boolean, ByVal Command As FileViewControl.CmdTypes)
On Error Resume Next

Dim FiNam As String
Dim AusZa As Integer

Set FM = frmMain
Set LiFi1 = FM.filView1

If LiFi1.SelectedCount > 0 Then
    Set LiFit = LiFi1.FirstSelectedItem
    FiNam = LiFit.DisplayName
            
    Select Case LCase(Right$(FiNam, 3))
    Case "ini":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "mdb":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "ldb":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "dbx":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "dbv":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "dax":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "blg":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "crd":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "lst":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "lsv":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "lsp":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "bat":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case Else:
        If GlRDP = True Then
            For AusZa = 1 To UBound(GlAus)
                If LCase(Right$(FiNam, 3)) = LCase(GlAus(0, AusZa)) Then
                    SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
                    Cancel = True
                    Exit For
                End If
            Next AusZa
        End If
    End Select
End If

End Sub
Private Sub filView1_OnBeforeItemDelete(ByVal Item As FileViewControl.IListItem)
On Error Resume Next

Dim FiNam As String

If GlAkt = False Then
    If Item.Selected = True Then
        FiNam = Item.DisplayName
        If Item.Attributes(Folder) And Folder Then
            Select Case FiNam
            Case "Praxisdaten": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht kopiert werden", IC48_Forbidden
            Case "Abrechnung": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht kopiert werden", IC48_Forbidden
            Case "Backup": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht kopiert werden", IC48_Forbidden
            End Select
        Else
            Select Case LCase(Right$(FiNam, 3))
            Case "mdb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "ldb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "dbx": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "dbv": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "dax": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            End Select
        End If
    End If
End If

End Sub
Private Sub filView1_OnEndItemRename(ByVal Item As FileViewControl.IListItem, NewName As String, Cancel As Boolean)
On Error Resume Next

Dim Posit As Integer

Set LiFi1 = FM.filView1

Select Case LCase(Right$(NewName, 3))
Case "sys": Cancel = True
Case "ini": Cancel = True
Case "bat": Cancel = True
Case "vbs": Cancel = True
Case "exe": Cancel = True
Case "com": Cancel = True
Case "hta": Cancel = True
Case "pif": Cancel = True
Case "scf": Cancel = True
Case "scr": Cancel = True
Case "cmd": Cancel = True
End Select

Posit = InStrRev(NewName, ".", -1, 1)

If Posit <= 0 Then Cancel = True

LiFi1.RefreshViewFast

End Sub
Private Sub filView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo 1
        End If
    End If
End Sub

Private Sub filView2_OnBeforeItemDelete(ByVal Item As FileViewControl.IListItem)
On Error Resume Next

Dim FiNam As String

If GlAkt = False Then
    If Item.Selected = True Then
        FiNam = Item.DisplayName
        If Item.Attributes(Folder) And Folder Then
            Select Case FiNam
            Case "Praxisdaten": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht kopiert werden", IC48_Forbidden
            Case "Abrechnung": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht kopiert werden", IC48_Forbidden
            Case "Backup": SPopu "Systemordner", "Der Ordner: " & FiNam & " darf nicht kopiert werden", IC48_Forbidden
            End Select
        Else
            Select Case LCase(Right$(FiNam, 3))
            Case "mdb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "ldb": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "dbx": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "dbv": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            Case "dax": SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht gelöscht werden", IC48_Forbidden
            End Select
        End If
    End If
End If

End Sub
Private Sub filView2_OnBeforeShellCommandExecute(ByVal CommandStr As String, Cancel As Boolean, ByVal Command As FileViewControl.CmdTypes)
On Error Resume Next

Dim FiNam As String
Dim AusZa As Integer

Set FM = frmMain
Set LiFld = FM.fldView1
Set LiFi2 = FM.filView2

If LiFi2.SelectedCount > 0 Then
    Set LiFit = LiFi2.FirstSelectedItem
    FiNam = LiFit.DisplayName
            
    Select Case LCase(Right$(FiNam, 3))
    Case "ini":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "mdb":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "ldb":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "dbx":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "dbv":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "dax":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "blg":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "crd":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "lst":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "lsv":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "lsp":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case "bat":
        SPopu "Systemdatei", "Die Datei: " & FiNam & " darf nicht geöffnet werden", IC48_Forbidden
        Cancel = True
    Case Else:
        If GlRDP = True Then
            Cancel = True
            KDatei 11
        End If
    End Select
End If

End Sub
Private Sub filView2_OnEndItemRename(ByVal Item As FileViewControl.IListItem, NewName As String, Cancel As Boolean)
On Error Resume Next

Dim Posit As Integer

Set LiFi2 = FM.filView2

Select Case LCase(Right$(NewName, 3))
Case "sys": Cancel = True
Case "ini": Cancel = True
Case "bat": Cancel = True
Case "vbs": Cancel = True
Case "exe": Cancel = True
Case "com": Cancel = True
Case "hta": Cancel = True
Case "pif": Cancel = True
Case "scf": Cancel = True
Case "scr": Cancel = True
Case "cmd": Cancel = True
End Select

Posit = InStrRev(NewName, ".", -1, 1)

If Posit <= 0 Then Cancel = True

LiFi2.RefreshViewFast

End Sub
Private Sub fldView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo 2
        End If
    End If
End Sub

Private Sub fldView1_OnAfterNodeRename(ByVal Node As FolderViewControl.ITreeNode, NewName As String, Cancel As Boolean)
On Error Resume Next

Dim NoNam As String

NoNam = Node.DisplayName

If GlAkt = False Then
    Select Case NoNam
    Case "Praxisdaten":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Abrechnung":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Backup":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Dokumente":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Bilder":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Vorlagen":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Formulare":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    End Select
    Node.Selected = True
End If
    
End Sub

Private Sub fldView1_OnBeforeNodeDelete(ByVal Node As FolderViewControl.ITreeNode)
On Error Resume Next

Dim NoNam As String

NoNam = Node.DisplayName

If GlAkt = False Then
    Select Case NoNam
    Case "Praxisdaten":
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Abrechnung":
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Backup":
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Dokumente":
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Bilder":
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Vorlagen":
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    Case "Formulare":
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
    End Select
End If
    
End Sub
Private Sub fldView1_OnBeforeNodeRename(ByVal Node As FolderViewControl.ITreeNode, Cancel As Boolean)
On Error Resume Next

Dim NoNam As String

NoNam = Node.DisplayName

If GlAkt = False Then
    Select Case NoNam
    Case "Praxisdaten":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Abrechnung":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Backup":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Dokumente":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Bilder":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Vorlagen":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    Case "Formulare":
        Cancel = True
        SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
    End Select
    Node.Selected = True
End If

End Sub

Private Sub fldView1_OnBeforeShellCommandExecute(ByVal CommandStr As String, Cancel As Boolean, ByVal Command As FolderViewControl.CmdTypes)
On Error Resume Next

Dim NoNam As String

Set FM = frmMain
Set LiFld = FM.fldView1

Set LiNod = LiFld.SelectedNode

NoNam = LiNod.DisplayName

If GlAkt = False Then
    Select Case Command
    Case cmdDelete:
        Select Case NoNam
        Case "Praxisdaten":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
        Case "Abrechnung":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
        Case "Backup":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
        Case "Dokumente":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
        Case "Bilder":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
        Case "Vorlagen":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
        Case "Formulare":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht gelöscht werden", IC48_Forbidden
        End Select
    Case cmdRename:
        Select Case NoNam
        Case "Praxisdaten":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Abrechnung":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Backup":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Dokumente":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Bilder":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Vorlagen":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Formulare":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        End Select
    Case cmdProperties:
        Select Case NoNam
        Case "Praxisdaten":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Abrechnung":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Backup":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Dokumente":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Bilder":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Vorlagen":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        Case "Formulare":
            Cancel = True
            SPopu "Systemordner", "Der Ordner: " & NoNam & " darf nicht verändert werden", IC48_Forbidden
        End Select
    End Select
End If

End Sub

Private Sub Form_Activate()
    SPosi
End Sub
Private Sub Form_Load()
On Error GoTo WiErr

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMinHeight = 10500
    .ClientMinWidth = 15000
End With

FStat

Set FrmEx = Nothing
    
GlAkt = True

WindowClas2 Me.hwnd 'Subclassing fuer WM_DISPLAYCHANGE aktivieren (Grid-Fix nach RDP Reconnect)

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " frmMain " & Err.Number
Resume Next

End Sub
Private Sub calCont1_ContextMenu(ByVal x As Single, ByVal y As Single)
On Error Resume Next
    
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo

Set FM = frmMain
Set CaCol = FM.calCont1

Set CaHit = CaCol.ActiveView.HitTest

If Not CaHit.HitCode = xtpCalendarHitTestUnknown Then
    If Not CaHit.HitCode = xtpCalendarHitTestDayViewTimeScale Then
        SMePo 1
    Else
        SMePo 2
    End If
End If

Set CaCol = Nothing
    
End Sub

Private Sub calCont1_DblClick()
    FKaMo
    DoEvents
    STerm False, True
End Sub

Private Sub calCont1_EventAddedEx(ByVal pEvent As XtremeCalendarControl.CalendarEvent)
On Error Resume Next

Set FM = frmMain

If GlAkt = False Then
    If GlSeF = False Then 'Formular wird geladen
        Screen.MousePointer = vbHourglass

        If GlOTS = False Then 'Online-Terminbuchungs Sytem
            S_TeSa pEvent, True
        Else
            If GlOTK = True Then 'Online-Terminbuchungs System autom. Aktualisierung
                S_TeSa pEvent, True
            End If
        End If

        If GlTrL = False Then 'Termine werden geladen
            S_TeLi
        End If
        
        Screen.MousePointer = vbNormal
    End If
End If

Set CaCol = Nothing

End Sub
Private Sub calCont1_EventChangedEx(ByVal pEvent As XtremeCalendarControl.CalendarEvent)
On Error Resume Next

Set FM = frmMain

If GlAkt = False Then
    If GlSeF = False Then 'Formular wird geladen
        Screen.MousePointer = vbHourglass

        If GlOTS = False Then 'Online-Terminbuchungs Sytem
            S_TeSa pEvent
        Else
            If GlOTK = True Then 'Online-Terminbuchungs System autom. Aktualisierung
                S_TeSa pEvent
            End If
        End If
        
        If GlTrL = False Then 'Termine werden geladen
            S_TeLi
        End If

        Screen.MousePointer = vbNormal
    End If
End If

Set CaCol = Nothing

End Sub
Private Sub calCont1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GlTrB = True
End Sub

Private Sub calCont1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo
    
Set FM = frmMain
Set CaCol = FM.calCont1

Set CaHit = CaCol.ActiveView.HitTest

If GlAkt = False Then
    If Button = vbLeftButton Then
        GlTrM = True 'Terminbearbeitung direkt im Kalender
        If CaHit.HitCode = 8193 Then
            Select Case GlBut
            Case RibTab_Ter_Raeume: GlCaR = CaHit.ViewGroup.GroupIndex 'Kalendarresourceindex
            Case RibTab_Ter_Mitarb: GlCaR = CaHit.ViewGroup.GroupIndex 'Kalendarresourceindex
            Case Else: GlCaR = 0
            End Select
            DoEvents
        Else
            GlCaR = 0
        End If
        FKaMo
    Else
        FTeSe
    End If
End If

End Sub
Private Sub calCont1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbLeftButton Then
            STeDe
        End If
    End If
End Sub
Private Sub calCont1_OnReminders(ByVal Action As XtremeCalendarControl.CalendarRemindersAction, ByVal Reminder As XtremeCalendarControl.CalendarReminder)
    If Action = xtpCalendarRemindersFire Then
        TimInit 1, 1
    End If
End Sub
Private Sub chkAbsen_Click()
    If GlAkt = False Then FRzAb
End Sub

Private Sub comBar01_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlAkt = False Then
        If Control.id = XTP_ID_RIBBONCONTROLTAB Then
            FTabu
        Else
            Select Case Control.id
            Case Tex_FaVor1: FTool Control.id, Control.Color
            Case Tex_FaHin1: FTool Control.id, Control.Color
            Case Tex_FntAu1: FTool Control.id, , Control.Text
            Case Tex_FntAu2: FTool Control.id, , Control.Text
            Case Tex_FntAu3: FTool Control.id, , Control.Text
            Case Tex_FntAu4: FTool Control.id, , Control.Text
            Case Tex_FntGr1: FTool Control.id, , Control.Text
            Case Tex_FntGr2: FTool Control.id, , Control.Text
            Case Tex_FntGr3: FTool Control.id, , Control.Text
            Case Tex_FntGr4: FTool Control.id, , Control.Text
            Case Tex_DaFeAd: FTool Control.id, , Control.Text
            Case Else: FTool Control.id
            End Select
        End If
    End If
End Sub
Private Sub comBar01_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlSta = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    SPosi
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub

Private Sub dcpDoc01_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)

If GlAkt = False Then
    Select Case Action
    Case PaneActionActivated:
        GlAcP = Pane.id
    Case PaneActionDeactivated:
        GlAcP = 0
    Case PaneActionPinned:
        Select Case GlBut
        Case RibTab_Krankenbla:
            GlP06 = False
            IniSetVal "GUI", "Pane06", GlP06
        Case RibTab_Abrechnung:
            GlP02 = False
            IniSetVal "GUI", "Pane02", GlP02
        Case RibTab_Rezeptmodul:
            GlP03 = False
            IniSetVal "GUI", "Pane03", GlP03
        Case RibTab_Belegmodul:
            GlP03 = False
            IniSetVal "GUI", "Pane03", GlP03
        Case RibTab_Rechnungen:
            GlP09 = False
            IniSetVal "GUI", "Pane09", GlP09
        Case RibTab_Buchungen:
            GlP08 = False
            IniSetVal "GUI", "Pane08", GlP08
        Case RibTab_HomeBanki:
            GlP01 = False
            IniSetVal "GUI", "Pane01", GlP01
        Case RibTab_LabBericht:
            GlP04 = False
            IniSetVal "GUI", "Pane04", GlP04
        Case RibTab_LabAuftrag:
            GlP04 = False
            IniSetVal "GUI", "Pane04", GlP04
        Case RibTab_Ter_Kalend:
            GlP05 = False
            IniSetVal "GUI", "Pane05", GlP05
        Case RibTab_Ter_Raeume:
            GlP05 = False
            IniSetVal "GUI", "Pane05", GlP05
        Case RibTab_Ter_Mitarb:
            GlP05 = False
            IniSetVal "GUI", "Pane05", GlP05
        Case RibTab_Bildmodul:
            GlP07 = False
            IniSetVal "GUI", "Pane07", GlP07
        Case RibTab_Tex_Dokumt:
            GlP10 = False
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_Vorlag:
            GlP10 = False
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_Rezept:
            GlP10 = False
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_NewsLe:
            GlP10 = False
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_Email:
            GlP11 = False
            IniSetVal "GUI", "Pane11", GlP11
        End Select
        SSpSav
    Case PaneActionUnpinned:
        Select Case GlBut
        Case RibTab_Krankenbla:
            GlP06 = True
            IniSetVal "GUI", "Pane06", GlP06
        Case RibTab_Abrechnung:
            GlP02 = True
            IniSetVal "GUI", "Pane02", GlP02
        Case RibTab_Rezeptmodul:
            GlP03 = True
            IniSetVal "GUI", "Pane03", GlP03
        Case RibTab_Belegmodul:
            GlP03 = True
            IniSetVal "GUI", "Pane03", GlP03
        Case RibTab_Rechnungen:
            GlP09 = True
            IniSetVal "GUI", "Pane09", GlP09
        Case RibTab_Buchungen:
            GlP08 = True
            IniSetVal "GUI", "Pane08", GlP08
        Case RibTab_HomeBanki:
            GlP01 = True
            IniSetVal "GUI", "Pane01", GlP01
        Case RibTab_LabBericht:
            GlP04 = True
            IniSetVal "GUI", "Pane04", GlP04
        Case RibTab_LabAuftrag:
            GlP04 = True
            IniSetVal "GUI", "Pane04", GlP04
        Case RibTab_Ter_Kalend:
            GlP05 = True
            IniSetVal "GUI", "Pane05", GlP05
        Case RibTab_Ter_Raeume:
            GlP05 = True
            IniSetVal "GUI", "Pane05", GlP05
        Case RibTab_Ter_Mitarb:
            GlP05 = True
            IniSetVal "GUI", "Pane05", GlP05
        Case RibTab_Bildmodul:
            GlP07 = True
            IniSetVal "GUI", "Pane07", GlP07
        Case RibTab_Tex_Dokumt:
            GlP10 = True
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_Vorlag:
            GlP10 = True
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_Rezept:
            GlP10 = True
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_NewsLe:
            GlP10 = True
            IniSetVal "GUI", "Pane10", GlP10
        Case RibTab_Tex_Email:
            GlP11 = True
            IniSetVal "GUI", "Pane11", GlP11
        End Select
        SSpSav
    Case PaneActionExpanded:
        If GlBut = RibTab_Abrechnung Then
            GlPan = Pane.id
        End If
    Case PaneActionClosed:
    Case PaneActionSplitterResized:
        Set DocPa = Me.dcpDoc01
        IniSetVal "DocPa3", "PanLay", DocPa.SaveStateToString
        Set DocPa = Nothing
    Case PaneActionDocked:
        Set DocPa = Me.dcpDoc01
        Select Case GlBut
        Case RibTab_Krankenbla: IniSetVal "DocPa3", "DocSi1", Pane.Position
        Case RibTab_Bildmodul:
        End Select
        IniSetVal "DocPa3", "PanLay", DocPa.SaveStateToString
        Set DocPa = Nothing
    End Select
End If

End Sub
Private Sub FDrop(Optional ByVal DrDat As String, Optional ByVal KrSor As Long, Optional ByVal Drubr As Boolean)
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim IdxNr As Long
Dim DroDa As Date
Dim KetNa As String
Dim KetKu As String
Dim DocPa As XtremeDockingPane.DockingPane
Dim DcTab As XtremeDockingPane.TabPaintManager
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set DocPa = FM.dcpDoc01
Set RpCo3 = FM.repCont3
Set RpCo5 = FM.repCont5

If DrDat <> vbNullString Then
    If IsDate(DrDat) = True Then
        DroDa = DrDat
    Else
        DroDa = GlTag(1)
    End If
Else
    DroDa = GlTag(1)
End If

Select Case GlBut
Case RibTab_LabBericht:
        Set RpSel = RpCo5.SelectedRows
        If RpSel.Count > 0 Then
            Set FM = frmKatPE
            Set CmBrs = FM.comBar02
            Set RbBar = CmBrs.Item(1)
            Set RbTab = RbBar.SelectedTab
            Select Case RbTab.id
            Case RibTab_Kat_EinLaP: K_Kat2 "LaPa", , 1
            Case RibTab_Kat_KetLaP: K_Kat2 "LaPa", True, 1
            End Select
        Else
            SUpLa
        End If
Case RibTab_LabAuftrag:
        Set RpSel = RpCo5.SelectedRows
        If RpSel.Count > 0 Then
            Set FM = frmKatPE
            Set CmBrs = FM.comBar02
            Set RbBar = CmBrs.Item(1)
            Set RbTab = RbBar.SelectedTab
            Select Case RbTab.id
            Case RibTab_Kat_EinLaP: K_Kat2 "LaPa", , 1
            Case RibTab_Kat_KetLaP: K_Kat2 "LaPa", True, 1
            End Select
        Else
            SUpAu
        End If
Case RibTab_Rezeptmodul:
        Set RpSel = RpCo5.SelectedRows
        If RpSel.Count > 0 Then
            Set FM = frmKatRE
            Set CmBrs = FM.comBar02
            Set RbBar = CmBrs.Item(1)
            Set RbTab = RbBar.SelectedTab
            Select Case RbTab.id
            Case RibTab_Kat_EinRez: K_RzEi "ReEi", , 1
            Case RibTab_Kat_KetRez: K_RzEi "ReEi", True, 1
            End Select
        Else
            SUpRe
        End If
Case RibTab_Krankenbla:
        Select Case GlAcP 'Active Pane
        Case PA_DP_KaAE:
                Set FM = frmKatAE
                K_Kat1 "AnEi", , 1, DroDa, KrSor, Drubr
        Case PA_DP_KaKD:
                Set FM = frmKatKD
                Set CmBrs = FM.comBar02
                Set RbBar = CmBrs.Item(1)
                Set RbTab = RbBar.SelectedTab
                Select Case RbTab.id
                Case RibTab_Kat_EinDiK: K_Kat1 "KrDi", , 1, DroDa, KrSor, Drubr
                Case RibTab_Kat_KetDiK: K_Kat1 "KrDi", True, 1, DroDa, KrSor, Drubr
                End Select
        Case PA_DP_KaKM:
                Set FM = frmKatKM
                Set CmBrs = FM.comBar02
                Set RbBar = CmBrs.Item(1)
                Set RbTab = RbBar.SelectedTab
                Select Case RbTab.id
                Case RibTab_Kat_EinMeK: K_Kat1 "KrMe", , 1, DroDa, KrSor, Drubr
                Case RibTab_Kat_KetMeK: K_Kat1 "KrMe", True, 1, DroDa, KrSor, Drubr
                End Select
        End Select
Case Else:
    Set RpSel = RpCo3.SelectedRows
    If RpSel.Count > 0 Then
        Select Case GlAcP 'Active Pane
        Case PA_DP_KaDE:
                Set FM = frmKatDE
                Set RpCo7 = FM.repCont7
                Set CmBrs = FM.comBar02
                Set RbBar = CmBrs.Item(1)
                Set RbTab = RbBar.SelectedTab
                Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
                Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
                Select Case RbTab.id
                Case RibTab_Kat_EinDia: K_Kat1 "DiEi", , 1, DroDa, KrSor, Drubr
                Case RibTab_Kat_KetDia:
                    Set RpCls = RpCo7.Columns
                    Set RpSel = RpCo7.SelectedRows
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        If RpRow.GroupRow = False Then
                            Set RpCol = RpCls.Find(Kat_ID0)
                            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                            Set RpCol = RpCls.Find(Kat_GOID)
                            KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            Set RpCol = RpCls.Find(Kat_IDKurz)
                            KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            GlNod = "F1"
                            GlKSt = "DiEi"
                            GlKeE = True
                            EMain IdxNr, KetNa, KetKu
                        End If
                    End If
                End Select
        Case PA_DP_KaGE:
                Set FM = frmKatGE
                Set RpCo7 = FM.repCont7
                Set CmBrs = FM.comBar02
                Set RbBar = CmBrs.Item(1)
                Set RbTab = RbBar.SelectedTab
                Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
                Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
                Select Case RbTab.id
                Case RibTab_Kat_EinGeb: K_Kat1 "GbEi", , 1, DroDa, KrSor, Drubr
                Case RibTab_Kat_KetGeb:
                    Set RpCls = RpCo7.Columns
                    Set RpSel = RpCo7.SelectedRows
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        If RpRow.GroupRow = False Then
                            Set RpCol = RpCls.Find(Kat_ID0)
                            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                            Set RpCol = RpCls.Find(Kat_GOID)
                            KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            Set RpCol = RpCls.Find(Kat_IDKurz)
                            KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            GlNod = "D" & GlGKa(CmSu2.ListIndex, 0)
                            GlKSt = "GbEi"
                            GlKeE = True
                            EMain IdxNr, KetNa, KetKu
                        End If
                    End If
                End Select
        Case PA_DP_KaME:
                Set FM = frmKatME
                Set RpCo7 = FM.repCont7
                Set CmBrs = FM.comBar02
                Set RbBar = CmBrs.Item(1)
                Set RbTab = RbBar.SelectedTab
                Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
                Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
                Select Case RbTab.id
                Case RibTab_Kat_EinMed: K_Kat1 "MeEi", , 1, DroDa, KrSor, Drubr
                Case RibTab_Kat_KetMed:
                    Set RpCls = RpCo7.Columns
                    Set RpSel = RpCo7.SelectedRows
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        If RpRow.GroupRow = False Then
                            Set RpCol = RpCls.Find(Kat_ID0)
                            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                            Set RpCol = RpCls.Find(Kat_GOID)
                            KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            Set RpCol = RpCls.Find(Kat_IDKurz)
                            KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            GlNod = "J1"
                            GlKSt = "MeEi"
                            GlKeE = True
                            EMain IdxNr, KetNa, KetKu
                        End If
                    End If
                End Select
        Case PA_DP_KaAR:
                Set FM = frmKatAR
                Set RpCo7 = FM.repCont7
                Set CmBrs = FM.comBar02
                Set RbBar = CmBrs.Item(1)
                Set RbTab = RbBar.SelectedTab
                Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
                Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
                Select Case RbTab.id
                Case RibTab_Kat_EinMed: K_Kat1 "ArLi", , 1, DroDa, KrSor, Drubr
                Case RibTab_Kat_KetMed:
                    Set RpCls = RpCo7.Columns
                    Set RpSel = RpCo7.SelectedRows
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        If RpRow.GroupRow = False Then
                            Set RpCol = RpCls.Find(Kat_ID0)
                            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                            Set RpCol = RpCls.Find(Kat_GOID)
                            KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            Set RpCol = RpCls.Find(Kat_IDKurz)
                            KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            GlNod = "P1"
                            GlKSt = "ArLi"
                            GlKeE = True
                            EMain IdxNr, KetNa, KetKu
                        End If
                    End If
                End Select
        Case PA_DP_KaLE:
                Set FM = frmKatLE
                Set RpCo7 = FM.repCont7
                Set CmBrs = FM.comBar02
                Set RbBar = CmBrs.Item(1)
                Set RbTab = RbBar.SelectedTab
                Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
                Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
                Select Case RbTab.id
                Case RibTab_Kat_EinLab: K_Kat1 "LaEi", , 1, DroDa, KrSor, Drubr
                Case RibTab_Kat_KetLab:
                    Set RpCls = RpCo7.Columns
                    Set RpSel = RpCo7.SelectedRows
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        If RpRow.GroupRow = False Then
                            Set RpCol = RpCls.Find(Kat_ID0)
                            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                            Set RpCol = RpCls.Find(Kat_GOID)
                            KetKu = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            Set RpCol = RpCls.Find(Kat_IDKurz)
                            KetNa = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                            GlNod = "H" & GlGKa(CmSu2.ListIndex, 0)
                            GlKSt = "LaEi"
                            GlKeE = True
                            EMain IdxNr, KetNa, KetKu
                        End If
                    End If
                End Select
        Case PA_DP_KaBE:
                Set FM = frmKatBE
                K_Kat1 "BeEi", , 1, DroDa, KrSor, Drubr
        Case PA_DP_KaRE:
                Set FM = frmKatRE
        Case PA_DP_KaLP:
                Set FM = frmKatPE
        End Select
    Else
        SUpAb
    End If
End Select

Set RpCo3 = Nothing
Set RpCo5 = Nothing
Set RpCo7 = Nothing
Set DocPa = Nothing
Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDrop " & Err.Number
Resume Next

End Sub
Private Sub FEmal()
On Error GoTo LaErr

Dim PatNr As Long
Dim EmAdr As String
Dim EmTex As String
Dim EmBrf As String

If GlAdr > 0 Then
    PatNr = GlAdr
    S_AdDe PatNr 'Adressendetails
    With GlADt
        EmAdr = .AdTe5
        EmBrf = .AdBrf
    End With
    EmTex = EmBrf & vbCrLf & vbCrLf
    SMaNe PatNr, EmAdr, , EmTex
End If

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDrop " & Err.Number
Resume Next

End Sub
Private Sub FEinf(ByVal Flag As Integer)
On Error Resume Next

Dim DayFi As Date
Dim DayLa As Date
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim MoKal As XtremeCalendarControl.DatePicker

Select Case Flag
Case 1: Set FM = frmKatDE
Case 2: Set FM = frmKatKD
Case 3: Set FM = frmKatKM
End Select

Set MoKal = FM.dtpDatu1
Set CmBrs = FM.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

With MoKal
    DayFi = .FirstDayOfWeek
    DayLa = .LastVisibleDay
End With

Select Case Flag
Case 1:
    Select Case RbTab.id
    Case RibTab_Kat_EinDia: K_Kat1 "DiEi", False, 2, GlTag(1)
    Case RibTab_Kat_KetDia: K_Kat1 "DiEi", True, 2, GlTag(1)
    End Select
Case 2:
    Select Case RbTab.id
    Case RibTab_Kat_EinDiK: K_Kat1 "KrDi", False, 2, GlTag(1)
    Case RibTab_Kat_KetDiK: K_Kat1 "KrDi", True, 2, GlTag(1)
    End Select
Case 3:
    Select Case RbTab.id
    Case RibTab_Kat_EinMeK: K_Kat1 "KrMe", False, 2, GlTag(1)
    Case RibTab_Kat_KetMeK: K_Kat1 "KrMe", True, 2, GlTag(1)
    End Select
End Select

S_AbTe DayFi, DayLa

Set MoKal = Nothing
Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

End Sub
Private Sub FFrAb()
On Error GoTo OrErr
'Fragebogen Anfordern

Dim TmStr As String

TmStr = InputBox("Wie lautet die Einreichungs-ID des Fragebogens, der zum Abruf angefordert werden soll?", "Fragebogen Anfordern", vbNullString)

If TmStr <> vbNullString Then
    TmStr = Trim$(TmStr)
    S_AnBoU TmStr
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FFrAb " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case GlBut
Case RibTab_Startseite:
    TeTit = IniGetOpt("Hilfe", 50001)
    TeMai = IniGetOpt("Hilfe", 50002)
    TeInh = IniGetOpt("Hilfe", 50003)
    TeFus = IniGetOpt("Hilfe", 50004)
Case RibTab_Adressen:
    TeTit = IniGetOpt("Hilfe", 50921)
    TeMai = IniGetOpt("Hilfe", 50922)
    TeInh = IniGetOpt("Hilfe", 50923)
    TeFus = IniGetOpt("Hilfe", 50924)
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
Case RibTab_Fragebogen:
    TeTit = IniGetOpt("Hilfe", 51051)
    TeMai = IniGetOpt("Hilfe", 51052)
    TeInh = IniGetOpt("Hilfe", 51053)
    TeFus = IniGetOpt("Hilfe", 51054)
Case RibTab_Krankenbla:
    TeTit = IniGetOpt("Hilfe", 51061)
    TeMai = IniGetOpt("Hilfe", 51062)
    TeInh = IniGetOpt("Hilfe", 51063)
    TeFus = IniGetOpt("Hilfe", 51064)
Case RibTab_Rezeptmodul:
    TeTit = IniGetOpt("Hilfe", 51071)
    TeMai = IniGetOpt("Hilfe", 51072)
    TeInh = IniGetOpt("Hilfe", 51073)
    TeFus = IniGetOpt("Hilfe", 51074)
Case RibTab_Belegmodul:
    TeTit = IniGetOpt("Hilfe", 51081)
    TeMai = IniGetOpt("Hilfe", 51082)
    TeInh = IniGetOpt("Hilfe", 51083)
    TeFus = IniGetOpt("Hilfe", 51084)
Case RibTab_Bildmodul:
    TeTit = IniGetOpt("Hilfe", 51091)
    TeMai = IniGetOpt("Hilfe", 51092)
    TeInh = IniGetOpt("Hilfe", 51093)
    TeFus = IniGetOpt("Hilfe", 51094)
Case RibTab_Abrechnung:

Case RibTab_Tagesproto:

Case RibTab_Vorbereit:

Case RibTab_Rechnungen:

Case RibTab_Mahnwesen:

Case RibTab_Buchungen:

Case RibTab_Statistik:

Case RibTab_HomeBanki:

Case RibTab_Ter_Kalend:

Case RibTab_Ter_Raeume:

Case RibTab_Ter_Mitarb:

Case RibTab_Ter_Listen:

Case RibTab_Ter_Akont:

Case RibTab_Ter_Warte:

Case RibTab_LabBericht:

Case RibTab_LabBerichte:

Case RibTab_LabAuftrag:

Case RibTab_LabAuftrage:

Case RibTab_Tex_Dokumt:

Case RibTab_Tex_Vorlag:

Case RibTab_Kat_Eintrg:

Case RibTab_Kat_Ketten:

Case RibTab_Ket_Edit:

Case RibTab_Ket_Anwe:

Case RibTab_Kra_View:

Case RibTab_Kra_Noti:

Case RibTab_Wart_Noti:

Case RibTab_Wart_Wied:

Case RibTab_Wart_Beha:

Case RibTab_Kat_Explor:

Case RibTab_Kat_Frage:

Case RibTab_Tex_Email:

Case RibTab_Tex_Rezept:

Case RibTab_Tex_NewsLe:

Case RibTab_Kat_EinDia:

Case RibTab_Kat_KetDia:

Case RibTab_Kat_EinGeb:

Case RibTab_Kat_KetGeb:

Case RibTab_Kat_EinMed:

Case RibTab_Kat_KetMed:

Case RibTab_Kat_EinBeg:

Case RibTab_Kat_KetBeg:

Case RibTab_Kat_EinLab:

Case RibTab_Kat_KetLab:

Case RibTab_Kat_EinTer:

Case RibTab_Kat_KetTer:

Case RibTab_Kat_EinLaP:

Case RibTab_Kat_KetLaP:

Case RibTab_Kat_EinDiK:

Case RibTab_Kat_KetDiK:

Case RibTab_Kat_EinMeK:

Case RibTab_Kat_KetMeK:

Case RibTab_Kat_EinBuc:

Case RibTab_Kat_KetBuc:

Case RibTab_Kat_EinTex:

Case RibTab_Kat_KetTex:

Case RibTab_Kat_EinAna:

Case RibTab_Kat_KetAna:

Case RibTab_Kat_EinRez:

Case RibTab_Kat_KetRez:

Case RibTab_Kat_EinBan:

Case RibTab_Kat_KetBan:

Case RibTab_Kat_EinRec:

Case RibTab_Kat_KetRec:

End Select

If TeTit <> vbNullString Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
End If

End Sub
Private Sub FKaHe()
On Error GoTo OrErr
'Wechselt in die Tagesansicht auf den heutigen Tag

Dim AkDat As Date
Dim ZelNr As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CaCol = FM.calCont1
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

If GlKFo = True Then 'Kalenderfocus
    AkDat = TimeValue(Now)
Else
    AkDat = TimeValue(IniGetVal("TerSys", "StaZei"))
End If

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

CmAcs(SY_TE_Termin_ArWoche).Checked = False
CmAcs(SY_TE_Termin_ErWoche).Checked = False
CmAcs(SY_TE_Termin_Woche).Checked = False
CmAcs(SY_TE_Termin_Monat).Checked = False
CmAcs(SY_TE_Termin_Tag).Checked = True

With CaCol
    .UseMultiColumnWeekMode = True
    .ViewType = xtpCalendarDayView
    .ActiveView.ShowDay Date, False
    ZelNr = .DayView.GetCellNumber(TimeValue(AkDat))
    If ZelNr < 1 Then ZelNr = 1
    .DayView.ScrollV ZelNr
End With

GlCal = 1 'Kalenderanzeige

IniSetVal "Layout", "KalAnz", 1

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmBrs = Nothing
Set CaCol = Nothing

Set clFen = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaHe " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu2 As Date
Dim Datu3 As Date

Set FM = frmMain
Set TxDa2 = FM.txtDatu2
Set TxDa3 = FM.txtDatu3
Set DaPi6 = FM.dtpDatu6
Set Rahm5 = FM.frmRahm5

Select Case KalWa
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

With DaPi6
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    Select Case KalWa
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left + Rahm5.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 3: .Top = TxDa3.Top + TxDa3.Height
            .Left = TxDa3.Left + Rahm5.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa3.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    End Select
End With

Datu2 = TxDa2.Text
Datu3 = TxDa3.Text

If Datu3 < Datu2 Then TxDa2.Text = Datu3

Set DaPi6 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKaMo()
On Error GoTo OrErr
'Mouse click on Calender

Dim DayFi As Date
Dim DayLa As Date
Dim TmDat As String
Dim MitId As Integer
Dim DayGa As Boolean
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo

Set FM = frmMain
Set CaCol = FM.calCont1

Set ViEvs = CaCol.ActiveView.GetSelectedEvents
If ViEvs.Count > 0 Then
    For Each ViEvt In ViEvs
        If ViEvt.Selected = True Then
            Set CaEvt = ViEvt.Event
            TmDat = CaEvt.StartTime
            DayGa = CaEvt.AllDayEvent
            Exit For
        End If
    Next ViEvt
End If

Set CaHit = CaCol.ActiveView.HitTest

If CaHit.ViewEvent Is Nothing Then
    If TmDat = vbNullString Then
        CaCol.ActiveView.GetSelection DayFi, DayLa, DayGa
        TmDat = DayFi
    End If
    GlTem = 0
Else
    Set ViEvt = CaHit.ViewEvent
    GlTem = ViEvt.Event.id
    MitId = ViEvt.Event.ScheduleID
End If

If CaHit.HitCode = 8193 Or 16385 Then
    Select Case GlBut
    Case RibTab_Ter_Kalend:
                    GlTRx = GlRau - 1
                    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
                        If GlOTS = True Then 'Online-Terminbuchungs Sytem
                            GlTBx = GlMiT(GlSMo, 0) - 1
                        Else
                            GlTBx = GlMiA(GlSmI, 0) - 1
                        End If
                    Else
                        GlTBx = GlMan(GlSMa, 0) - 1
                    End If
    Case RibTab_Ter_Raeume:
                    GlTRx = CaHit.ViewGroup.GroupIndex 'Recourcegruppenindex Raum
                    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
                        If GlOTS = True Then 'Online-Terminbuchungs Sytem
                            GlTBx = GlMiT(GlSMo, 0) - 1
                        Else
                            GlTBx = GlMiA(GlSmI, 0) - 1
                        End If
                    Else
                        GlTBx = GlMan(GlSMa, 0) - 1
                    End If
    Case RibTab_Ter_Mitarb:
                    GlTRx = GlRau - 1
                    GlTBx = CaHit.ViewGroup.GroupIndex 'Recourcegruppenindex Mandant / Mitarbeiter
    End Select
Else
    GlTBx = MitId - 1
End If

If TmDat = vbNullString Then
    TmDat = CaHit.HitDateTime 'Ausgewählter Tag im Kalender
End If

TmDat = RTrim$(TmDat)

If Len(TmDat) < 12 Then
    If Mid$(TmDat, 3, 1) = "." Then
        TmDat = TmDat & " " & "08:00:00"
    ElseIf Mid$(TmDat, 3, 1) = ":" Then
        TmDat = DayFi
    End If
End If

GlDay = TmDat

Set CaCol = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaMo " & Err.Number
Resume Next

End Sub

Private Sub FKaSe()
On Error GoTo OrErr
'Click onto left MonthCalendar

Dim DaSta As Date
Dim DaEnd As Date
Dim AkDat As Date
Dim ZelNr As Integer

Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents

Set FM = frmMain
Set DaPi7 = FM.dtpDatu7
Set CaCol = FM.calCont1

Set ViEvs = CaCol.ActiveView.GetSelectedEvents

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

GlAkt = True

If GlKFo = True Then 'Kalenderfocus
    AkDat = TimeValue(Now)
Else
    AkDat = TimeValue(IniGetVal("TerSys", "StaZei"))
End If

If ViEvs.Count > 0 Then
    For Each ViEvt In ViEvs
        If ViEvt.Selected = True Then
            ViEvt.Selected = False
        End If
    Next ViEvt
End If

DaSta = DaPi7.Selection.Blocks(0).DateBegin

Select Case GlCal 'Kalenderanzeige
Case 1:
    DaEnd = DaSta
Case 2:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 4, DaSta)
Case 3:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 6, DaSta)
Case 4:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 6, DaSta)
Case 5:
    If Weekday(DaSta, vbMonday) > 1 Then
        DaSta = DateAdd("d", -(Weekday(DaSta, vbMonday) - 1), DaSta)
    End If
    DaEnd = DateAdd("d", 29, DaSta)
End Select

GlDFi = DaSta
GlDLa = DaEnd

With CaCol
    Select Case GlCal 'Kalenderanzeige
    Case 1:
        .ActiveView.ShowDay DaSta, True
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayAllWeek
        .ViewType = xtpCalendarDayView
    Case 2:
        .ActiveView.ShowDay DaSta, True
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayMo_Fr
        .ViewType = xtpCalendarWorkWeekView
    Case 3:
        .ActiveView.ShowDay DaSta, True
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayAllWeek
        .ViewType = xtpCalendarFullWeekView
    Case 4:
        .ActiveView.ShowDay DaSta, True
        .UseMultiColumnWeekMode = False
        .ViewType = xtpCalendarWeekView
    Case 5:
        .ActiveView.ShowDay DaSta, True
        .UseMultiColumnWeekMode = False
        .ViewType = xtpCalendarMonthView
    End Select
    ZelNr = .DayView.GetCellNumber(TimeValue(AkDat))
    If ZelNr < 1 Then ZelNr = 1
    .DayView.ScrollV ZelNr
End With

DoEvents
S_TeLi

GlAkt = False

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set CaCol = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaSe " & Err.Number
Resume Next

End Sub
Private Sub FKaSu()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim DaPi3 As XtremeCalendarControl.DatePicker
Dim CmDat As XtremeCommandBars.CommandBarEdit

Set FM = frmMain
Set CmBrs = FM.comBar01
Set DaPi3 = FM.dtpDatu3

Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)

If IsDate(CmDat.Text) Then
    NeuDa = CmDat.Text
Else
    NeuDa = Date
End If

With DaPi3
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Left = 6550
    .Top = 2550
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            CmDat.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set DaPi3 = Nothing

FSuch

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaSu " & Err.Number
Resume Next

End Sub
Private Sub FKran(Optional ByVal CoIdx As Long)
On Error GoTo PoErr
'Änderungen im Krankneblatt

Dim RetWe As Long
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl

Set RpCo6 = Me.repCont6
Set RpCoK = Me.repContK

Select Case GlBut
Case RibTab_Krankenbla:
        Select Case CoIdx
        Case Kra_Ziffer: K_Inp
        Case Kra_Anz: SKrAk
        Case Kra_Faktor: SKrAk
        Case Kra_Betrag: SKrAk
        End Select
Case RibTab_Abrechnung:
        Select Case CoIdx
        Case Kra_Ziffer: K_Inp
        Case Kra_Anz: SKrAk
        Case Kra_Faktor: SKrAk
        Case Kra_Betrag: SKrAk
        Case Kra_IDD: SKrDi
        End Select
Case RibTab_LabBericht:
        Select Case CoIdx
        Case Lbl_Ergebniswert: S_LaEin
        Case Lbl_Gruppe:
        End Select
End Select

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case GlBut
Case RibTab_Krankenbla:
        GlSav = True
        S_KrSa
Case RibTab_Abrechnung:
        GlSav = True
        S_KrSa
Case RibTab_LabBericht:
        GlSav = True
        S_LaSa
Case RibTab_LabAuftrag:
        GlSav = True
        S_LaSa
End Select

If GlKrE = False Then
    Select Case GlBut
    Case RibTab_Krankenbla: RetWe = RpCoK.EnableDragDrop("Katalog", xtpReportAllowDrag + xtpReportAllowDrop)
    Case Else: RetWe = RpCo6.EnableDragDrop("Katalog", xtpReportAllowDrag + xtpReportAllowDrop)
    End Select
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set RpCo6 = Nothing
Set RpCoK = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKran " & Err.Number
Resume Next

End Sub
Private Sub FKrDu()
On Error GoTo PoErr

Set FM = frmMain

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

SClip
S_KrEi , True

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKrDu " & Err.Number
Resume Next

End Sub


Private Sub FMark(Optional ByVal OptMa As Integer = 0)
On Error GoTo AnErr
'Markiert alle Einträge

Dim NeuDa As Date
Dim Lange As Long
Dim GeSum As Double
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmPa2 As XtremeCommandBars.StatusBarPane
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set RpCo6 = FM.repCont6
Set RpCoK = FM.repContK
Set RpCo8 = FM.repCont8
Set TxCoN = FM.TexCont1
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmPa2 = CmSta.FindPane(Tex_Pa_Labl2)

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case GlBut
Case RibTab_Adressen:
        With RpCo2
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
        CmPa2.Text = "Patienten markiert : " & RpRws.Count
Case RibTab_Mandanten:
        With RpCo2
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
        CmPa2.Text = "Mandanten markiert : " & RpRws.Count
Case RibTab_Verordner:
        With RpCo2
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
        CmPa2.Text = "Verordner markiert : " & RpRws.Count
Case RibTab_Mitarbeit:
        With RpCo2
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
        CmPa2.Text = "Mitarbeiter markiert : " & RpRws.Count
Case RibTab_Fragebogen:
        With RpCo3
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Krankenbla:
        Select Case OptMa
        Case 1:
            With RpCoK
                Set RpRws = .Rows
                Set RpCls = .Columns
                Set RpSel = .SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        Set RpCol = RpCls.Find(Kra_Datum)
                        NeuDa = RpRow.Record(RpCol.ItemIndex).Value
                        For Each RpRow In RpRws
                            If RpRow.GroupRow = False Then
                                If RpRow.Record(Kra_Datum).Value = NeuDa Then
                                    RpRow.Selected = True
                                End If
                            End If
                        Next RpRow
                    End If
                End If
            End With
        Case Else:
            If GlKrO = True Then 'Krankenblattdokument
                With TxCoN
                    Lange = Len(.Text)
                    .SelStart = 0
                    .SelLength = Lange
                End With
            Else
                With RpCoK
                    If .Rows.Count > 0 Then
                        Set RpRws = .Rows
                        If GlFoc = True Then
                            Set .FocusedRow = RpRws.Row(1)
                        End If
                        For Each RpRow In RpRws
                            If RpRow.GroupRow = False Then
                                RpRow.Selected = True
                            End If
                        Next RpRow
                    End If
                End With
            End If
        End Select
Case RibTab_Abrechnung:
        Select Case OptMa
        Case 1:
            With RpCo6
                Set RpRws = .Rows
                Set RpCls = .Columns
                Set RpSel = .SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        Set RpCol = RpCls.Find(Kra_Datum)
                        NeuDa = RpRow.Record(RpCol.ItemIndex).Value
                        For Each RpRow In RpRws
                            If RpRow.GroupRow = False Then
                                If RpRow.Record(Kra_Datum).Value = NeuDa Then
                                    RpRow.Selected = True
                                End If
                            End If
                        Next RpRow
                    End If
                End If
            End With
        Case 2:
            With RpCo6
                Set RpRws = .Rows
                Set RpCls = .Columns
                Set RpSel = .SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = True Then
                        For Each RpRow In RpRow.Childs
                            If RpRow.GroupRow = False Then
                                RpRow.Selected = True
                            End If
                        Next RpRow
                    End If
                End If
            End With
        Case Else:
            With RpCo6
                If .Rows.Count > 0 Then
                    Set RpRws = .Rows
                    If GlFoc = True Then
                        Set .FocusedRow = RpRws.Row(1)
                    End If
                    For Each RpRow In RpRws
                        If RpRow.GroupRow = False Then
                            RpRow.Selected = True
                        End If
                    Next RpRow
                End If
            End With
        End Select
Case RibTab_Vorbereit:
        With RpCo6
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Belegmodul:
        With RpCo3
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Rezeptmodul:
        With RpCo3
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Bildmodul:
            
Case RibTab_Rechnungen:
        With RpCo4
            Set RpCls = .Columns
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                        Set RpCol = RpCls.Find(Rec_Type)
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            If RpRow.Record(RpCol.ItemIndex).Value <> "U" Then
                                If RpRow.Record(RpCol.ItemIndex).Value <> "V" Then
                                    Set RpCol = RpCls.Find(Rec_Betrag)
                                    GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                                End If
                            End If
                        End If
                    End If
                Next RpRow
                CmPa2.Text = "Rechnungen markiert : " & RpRws.Count & " - Summe : " & Format$(GeSum, GlWa1)
            End If
        End With
Case RibTab_Mahnwesen:
        With RpCo1
            Set RpCls = .Columns
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                        Set RpCol = RpCls.Find(OPo_OffBetrag)
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                        End If
                    End If
                Next RpRow
                CmPa2.Text = "Posten markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            End If
        End With
Case RibTab_Buchungen:
        With RpCo1
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_HomeBanki:
        With RpCo1
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Statistik:
            
Case RibTab_Ter_Kalend:
            
Case RibTab_Ter_Listen:
        With RpCo1
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Ter_Akont:
        With RpCo1
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Ter_Warte:
        With RpCo1
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Ter_Raeume:
            
Case RibTab_Ter_Mitarb:
            
Case RibTab_LabBericht:
        With RpCo5
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_LabAuftrag:
        With RpCo5
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_LabBerichte:
        With RpCo1
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_LabAuftrage:
        With RpCo1
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Kat_Eintrg:
        With RpCo8
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Kat_Ketten:
        With RpCo8
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Kat_Frage:
        With RpCo8
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Tex_Email:
        With RpCo8
            If .Rows.Count > 0 Then
                Set RpRws = .Rows
                If GlFoc = True Then
                    Set .FocusedRow = RpRws.Row(1)
                End If
                For Each RpRow In RpRws
                    If RpRow.GroupRow = False Then
                        RpRow.Selected = True
                    End If
                Next RpRow
            End If
        End With
Case RibTab_Kat_Explor:
        LiFi2.SelectItems FIVSTAll
        For AktZa = 1 To LiFi2.SelectedCount
            Set LiFit = LiFi2.ListItem(AktZa)
            If LiFit.Attributes(Folder) And Folder Then
                LiFit.Selected = False
            End If
        Next AktZa
Case RibTab_Tex_Dokumt:
        With TxCoN
            Lange = Len(.Text)
            .SelStart = 0
            .SelLength = Lange
        End With
Case RibTab_Tex_Vorlag:
        With TxCoN
            Lange = Len(.Text)
            .SelStart = 0
            .SelLength = Lange
        End With
Case RibTab_Tex_Rezept:
        With TxCoN
            Lange = Len(.Text)
            .SelStart = 0
            .SelLength = Lange
        End With
Case RibTab_Tex_NewsLe:
        With TxCoN
            Lange = Len(.Text)
            .SelStart = 0
            .SelLength = Lange
        End With
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo6 = Nothing
Set RpCoK = Nothing
Set RpCo8 = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMark " & Err.Number
Resume Next

End Sub
Private Sub FMenu()
On Error GoTo AnErr
'Blendet im Krankenblatt Ribbongroups ein und aus

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups

Set FM = frmMain
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set CmAcs = CmBrs.Actions
Set RbTab = RbBar.FindTab(RibTab_Krankenbla)

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If GlMen = True Then 'Krankenblattdokument
    RbTab.Groups(1).Visible = False 'RibGrp_Kra_Eintrag
    RbTab.Groups(2).Visible = False 'RibGrp_Kra_Bearbeit
    RbTab.Groups(3).Visible = False 'RibGrp_Kra_Ansicht
    RbTab.Groups(4).Visible = True 'RibGrp_Tex_Clipboard
    RbTab.Groups(5).Visible = True 'RibGrp_Tex_Absatz
    RbTab.Groups(6).Visible = True 'RibGrp_Tex_Schrift
    CmAcs(SY_KB_KraBla_Expor).Enabled = True
    CmAcs(SY_KB_KraBla_Nachr).Enabled = True
    CmAcs(SY_KB_KraBla_DoLnk).Enabled = True
Else
    RbTab.Groups(1).Visible = True 'RibGrp_Kra_Eintrag
    RbTab.Groups(2).Visible = True 'RibGrp_Kra_Bearbeit
    RbTab.Groups(3).Visible = True 'RibGrp_Kra_Ansicht
    RbTab.Groups(4).Visible = False 'RibGrp_Tex_Clipboard
    RbTab.Groups(5).Visible = False 'RibGrp_Tex_Absatz
    RbTab.Groups(6).Visible = False 'RibGrp_Tex_Schrift
    CmAcs(SY_KB_KraBla_Expor).Enabled = False
    CmAcs(SY_KB_KraBla_Nachr).Enabled = False
    CmAcs(SY_KB_KraBla_DoLnk).Enabled = False
End If

RbBar.RedrawBar

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMona " & Err.Number
Resume Next

End Sub
Private Sub FMona(Optional ByVal SetFi As Boolean)
On Error GoTo AnErr
'Monat wechselt

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CaCol = FM.calCont1

Set CmCom = CmBrs.FindControl(CmCom, SY_TE_Termin_FiltIdx, , True)

If SetFi = True Then
    GlCaS = CmCom.ListIndex 'Kalenderfilterinhalt
    IniSetVal "TerSys", "KaGrSu", GlCaS
End If

Screen.MousePointer = vbHourglass

With CaCol
    Select Case GlCal 'Kalenderanzeige
    Case 1:
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayAllWeek
        .ViewType = xtpCalendarDayView
    Case 2:
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayMo_Fr
        .ViewType = xtpCalendarWorkWeekView
    Case 3:
        .UseMultiColumnWeekMode = True
        .Options.WorkWeekMask = xtpCalendarDayAllWeek
        .ViewType = xtpCalendarFullWeekView
    Case 4:
        .UseMultiColumnWeekMode = False
        .ViewType = xtpCalendarWeekView
    Case 5:
        .UseMultiColumnWeekMode = False
        .ViewType = xtpCalendarMonthView
    End Select
End With
DoEvents

S_TeLi
DoEvents

Screen.MousePointer = vbNormal

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMona " & Err.Number
Resume Next

End Sub
Private Sub FObje(ByVal TxFun As Integer)
On Error GoTo PoErr

Dim ObjNr As Long
Dim DrhDc As Long
Dim TxBre As Long
Dim TxHoh As Long
Dim ObjBr As Long
Dim ObjHo As Long
Dim SclBr As Long
Dim SclHo As Long
Dim ColID As Long
Dim PaStr As String
Dim ExVer As String
Dim SuFix As String
Dim FiNam As String
Dim SuStr As String
Dim DaNam As String
Dim DaPfa As String
Dim TmGui As String
Dim NeuNa As String
Dim DaNaO As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim Posit As Integer
Dim LadVa As Integer
Dim GesZa As Integer
Dim Frage As Integer
Dim AktZa As Integer
Dim ReDru As Integer
Dim Lange As Integer
Dim SeiZa As Variant
Dim RetWe As Boolean
Dim TxFnt As New StdFont
Dim Mld1, Tit1 As String
Dim Mld2, Tit2 As String
Dim Mld3, Tit3 As String

Set FM = frmMain
Set TxCoN = FM.TexCont1
Set CoDia = FM.comDialo
Set LiVw4 = FM.lstView4
Set TabCo = FM.TabCont1
Set LiIts = LiVw4.ListItems

Set clFil = New clsFile

Tit1 = "Dokument Überschreiben"
Tit2 = "Dokument Entfernen"
Tit3 = "Dokument Kopieren"
Mld1 = "Soll das markierte Dokument wirklich überschrieben werden?"
Mld2 = "Soll das markierte Dokument wirklich entfernt werden?"
Mld3 = "Soll das markierte Dokument wirklich kopiert werden?"

Select Case TxFun
Case Tex_DatSpe:
    RetWe = STxSa()
Case Tex_EinGr1:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.jpg"
        .DialogTitle = "Bitte Name und Ordner der Datei angeben"
        .Filter = "Joint Photographic Experts Group (.jpg)|*.jpg|Windows Bitmap (.bmp)|*.bmp|Portable Network Graphics (.png)|*.png|Tagged Image Format (.tif)|*.tif|Windows-Meta-File (.wmf)|*.wmf|Alle Dateien (*.*)|*.*"
        .InitDir = GlIPf
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
        TxBre = TxCoN.Width - ((TxCoN.PageMarginL + TxCoN.PageMarginR) * 2)
        ObjNr = TxCoN.ImageInsert(FiNam, -1, 1, 0, 0, 100, 100, 3, 100, 100, 100, 100)
        TxCoN.ObjectCurrent = ObjNr
        TxCoN.ObjectSizeMode = 3 'WICHTIG!
        TxCoN.ObjectInsertionMode = 2 'WICHTIG!
        TxCoN.ObjectTextflow = 1 'WICHTIG!
        TxCoN.ObjectTransparency = 0
        TxCoN.ImageDisplayMode = 0
        TxCoN.ImageSaveMode = 1
        ObjBr = TxCoN.ObjectWidth
        ObjHo = TxCoN.ObjectHeight
        If ObjBr > TxBre Then
            SclBr = Fix((TxBre / ObjBr) * 100) 'Abrunden
            TxCoN.ObjectScaleX = SclBr
            TxCoN.ObjectScaleY = SclBr
        End If
    End If
Case Tex_EinGr2:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.jpg"
        .DialogTitle = "Bitte Name und Ordner der Datei angeben"
        .Filter = "Joint Photographic Experts Group (.jpg)|*.jpg|Windows Bitmap (.bmp)|*.bmp|Portable Network Graphics (.png)|*.png|Tagged Image Format (.tif)|*.tif|Windows-Meta-File (.wmf)|*.wmf|Alle Dateien (*.*)|*.*"
        .InitDir = GlIPf
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
        TxBre = TxCoN.Width - ((TxCoN.PageMarginL + TxCoN.PageMarginR) * 2)
        ObjNr = TxCoN.ImageInsertAsChar(FiNam, -1, 100, 100)
        TxCoN.ObjectCurrent = ObjNr
        TxCoN.ObjectSizeMode = 3 'WICHTIG!
        TxCoN.ObjectInsertionMode = 1 'WICHTIG!
        TxCoN.ObjectTextflow = 0 'WICHTIG!
        TxCoN.ObjectTransparency = 0
        TxCoN.ImageDisplayMode = 0
        TxCoN.ImageSaveMode = 1
        ObjBr = TxCoN.ObjectWidth
        ObjHo = TxCoN.ObjectHeight
        If ObjBr > TxBre Then
            SclBr = Fix((TxBre / ObjBr) * 100) 'Abrunden
            TxCoN.ObjectScaleX = SclBr
            TxCoN.ObjectScaleY = SclBr
        End If
    End If
Case Tex_EinGr3:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.jpg"
        .DialogTitle = "Bitte Name und Ordner der Datei angeben"
        .Filter = "Joint Photographic Experts Group (.jpg)|*.jpg|Windows Bitmap (.bmp)|*.bmp|Portable Network Graphics (.png)|*.png|Tagged Image Format (.tif)|*.tif|Windows-Meta-File (.wmf)|*.wmf|Alle Dateien (*.*)|*.*"
        .InitDir = GlIPf
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
        TxBre = TxCoN.Width - ((TxCoN.PageMarginL + TxCoN.PageMarginR) * 2)
        SeiZa = TxCoN.CurrentInputPosition
        ObjNr = TxCoN.ImageInsertFixed(FiNam, SeiZa(0), 0, 0, 100, 100, 3, 100, 100, 100, 100)
        TxCoN.ObjectCurrent = ObjNr
        TxCoN.ObjectSizeMode = 1 'WICHTIG!
        TxCoN.ObjectInsertionMode = 3 'WICHTIG!
        TxCoN.ObjectTextflow = 3 'WICHTIG!
        TxCoN.ObjectTransparency = 0
        TxCoN.ImageDisplayMode = 0
        TxCoN.ImageSaveMode = 1
        ObjBr = TxCoN.ObjectWidth
        ObjHo = TxCoN.ObjectHeight
        If ObjBr > TxBre Then
            SclBr = Fix((TxBre / ObjBr) * 100) 'Abrunden
            TxCoN.ObjectScaleX = SclBr
            TxCoN.ObjectScaleY = SclBr
        End If
    End If
Case Tex_EinMar:
    With TxCoN
        ObjNr = .TextFrameInsert(-1, 0, -1000, 4350, 500, 500, 3, 0, 0, 0, 0)
        .ObjectCurrent = ObjNr
        .ObjectTextflow = 1
        .ObjectTransparency = 50
        .TextFrameBorderWidth = 0
        .TextFrameInternalMargin(1) = 0
        .TextFrameInternalMargin(2) = 0
        .TextFrameInternalMargin(3) = 0
        .TextFrameInternalMargin(4) = 0
        .TextFrameSelect ObjNr
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .SelText = "___"
        .TextFrameSelect 0
    End With
Case Tex_EinTex:
        TxCoN.ObjectAttrDialog
Case Tex_EinTab:
        TxCoN.TabDialog
Case Tex_EinObj:
        ObjNr = TxCoN.ObjectInsert(1, 0, -1, 0, 0, 0, 100, 100, 3, 100, 100, 100, 100)
        TxCoN.ObjectCurrent = ObjNr
        TxCoN.ObjectSizeMode = 3
        TxCoN.ObjectInsertionMode = 2
        TxCoN.ObjectTextflow = 1
        TxCoN.ObjectTransparency = 50
Case IC16_FarVor:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.ForeColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.ForeColor = ColID
Case IC16_FarHin:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.TextBkColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.TextBkColor = ColID
Case Tex_FaVor6:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.ForeColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.ForeColor = ColID
Case Tex_FaHin6:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .Color = TxCoN.TextBkColor
        .ShowColor
        ColID = .Color
    End With
    TxCoN.TextBkColor = ColID
Case Tex_TexRah:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    ObjNr = TxCoN.TextFrameInsert(-1, 1, 0, 0, 0, 0, 3, 200, 200, 200, 200)
Case Tex_ClpEin:
    With TxCoN
        If .CanPaste = True Then
            .Paste 2
        End If
    End With
Case Tex_ClpInh:
    With TxCoN
        If .CanPaste = True Then
            SuStr = Clipboard.GetText
            Clipboard.Clear
            Clipboard.SetText SuStr
            .Paste 5
        End If
    End With
Case Tex_DatLoa:
    With clFil
        .hwnd = FM.hwnd
        .StaVe = GlVor
        .DaTit = "Bitte Name und Ordner der Datei angeben"
        Select Case GlBut
        Case RibTab_Tex_Dokumt:
                .DaExt = "txm"
                .DaStr = "Textverarbeitung (.txm)|*.txm|Microsoft Word 2002/2003 (.doc)" & Chr(0) & "*.doc" & Chr(0) & "Rich Text Format (.rtf)" & Chr(0) & "*.rtf" & Chr(0) & "ANSI-Textdatei (.txt)" & Chr(0) & "*.txt" & Chr(0) & "Hypertext Markup Language (.htm)" & Chr(0) & "*.htm" & Chr(0) & "Extensible Markup Language (.xml)" & Chr(0) & "*.xml" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        Case RibTab_Tex_Vorlag:
                .DaExt = "txm"
                .DaStr = "Textverarbeitung (.txm)" & Chr(0) & "*.txm" & Chr(0) & "Microsoft Word 2002/2003 (.doc)" & Chr(0) & "*.doc" & Chr(0) & "Rich Text Format (.rtf)" & Chr(0) & "*.rtf" & Chr(0) & "ANSI-Textdatei (.txt)" & Chr(0) & "*.txt" & Chr(0) & "Hypertext Markup Language (.htm)" & Chr(0) & "*.htm" & Chr(0) & "Extensible Markup Language (.xml)" & Chr(0) & "*.xml" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        Case RibTab_Tex_Rezept:
                .DaExt = "txr"
                .DaStr = "Langrezepte (.txr)" & Chr(0) & "*.txr" & Chr(0) & "Microsoft Word 2002/2003 (.doc)" & Chr(0) & "*.doc" & Chr(0) & "Rich Text Format (.rtf)" & Chr(0) & "*.rtf" & Chr(0) & "ANSI-Textdatei (.txt)" & Chr(0) & "*.txt" & Chr(0) & "Hypertext Markup Language (.htm)" & Chr(0) & "*.htm" & Chr(0) & "Extensible Markup Language (.xml)" & Chr(0) & "*.xml" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        Case RibTab_Tex_NewsLe:
                .DaExt = "txn"
                .DaStr = "Newslettervorlage (.txn)" & Chr(0) & "*.txn" & Chr(0) & "Microsoft Word 2002/2003 (.doc)" & Chr(0) & "*.doc" & Chr(0) & "Rich Text Format (.rtf)" & Chr(0) & "*.rtf" & Chr(0) & "ANSI-Textdatei (.txt)" & Chr(0) & "*.txt" & Chr(0) & "Hypertext Markup Language (.htm)" & Chr(0) & "*.htm" & Chr(0) & "Extensible Markup Language (.xml)" & Chr(0) & "*.xml" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*"
        End Select
        FiNam = .FilOpn
    End With
    If FiNam = vbNullString Then
        Set CoDia = Nothing
        Set clFil = Nothing
        Exit Sub
    End If

    If Not IsNull(FiNam) And Not FiNam = vbNullString Then
        Posit = InStrRev(FiNam, ".", Len(FiNam), 1)
        If Posit > 0 Then
            SuFix = Mid$(FiNam, Posit + 1, Len(FiNam) - Posit)
        Else
            SuFix = vbNullString
        End If

        With clFil
            .FilPfa FiNam
            GlTDa = .DaNam
        End With

        GlTxF = FiNam 'Filname für Textcontrol Error
        GlTxU = LCase(SuFix)

        Select Case LCase(SuFix)
        Case "txm":
            With TxCoN
                .ResetContents
                .Load FiNam, , 3
            End With
        Case "txn":
            With TxCoN
                .ResetContents
                .Load FiNam, , 3
            End With
        Case "txr":
            With TxCoN
                .ResetContents
                .Load FiNam, , 3
            End With
        Case "htm":
            With TxCoN
                .ResetContents
                .Load FiNam, , 4
            End With
        Case "xml":
            With TxCoN
                .ResetContents
                .Load FiNam, , 10
            End With
        Case "css":
            With TxCoN
                .ResetContents
                .Load FiNam, , 11
            End With
        Case "doc":
            With TxCoN
                .ResetContents
                .Load FiNam, , 9
            End With
        Case "docx":
            With TxCoN
                .ResetContents
                .Load FiNam, , 13
            End With
        Case Else:
            With TxCoN
                .ResetContents
               .Text = vbNullString
            End With
        End Select
        DoEvents
        GlTxN = False 'Kein neues Dokument mehr
        GlTxS = False
        GlTSV = False 'Speichern Textverarbeitung
    End If
Case Tex_DatSpV:
    With clFil
        .hwnd = FM.hwnd
        .StaVe = GlVor
        .DaNam = vbNullString
        .DaTit = "Bitte Name und Ordner der Datei angeben"
        Select Case GlBut
        Case RibTab_Tex_Dokumt:
                .DaExt = "txm"
                .DaStr = "Textverarbeitung (*.txm)" & Chr(0) & "*.txm" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        Case RibTab_Tex_Vorlag:
                .DaExt = "txm"
                .DaStr = "Textverarbeitung (*.txm)" & Chr(0) & "*.txm" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        Case RibTab_Tex_Rezept:
                .DaExt = "txr"
                .DaStr = "Langrezepte (*.txr)" & Chr(0) & "*.txr" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        Case RibTab_Tex_NewsLe:
                .DaExt = "txn"
                .DaStr = "Newslettervorlage (*.txn)" & Chr(0) & "*.txn" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        End Select
         FiNam = .FilSav
    End With
    If FiNam = vbNullString Then
        Set CoDia = Nothing
        Set clFil = Nothing
        Exit Sub
    End If

    Select Case GlBut
    Case RibTab_Tex_Dokumt:
            If Right$(FiNam, 4) <> ".txm" Then FiNam = FiNam & ".txm"
    Case RibTab_Tex_Vorlag:
            If Right$(FiNam, 4) <> ".txm" Then FiNam = FiNam & ".txm"
    Case RibTab_Tex_Rezept:
            If Right$(FiNam, 4) <> ".txr" Then FiNam = FiNam & ".txr"
    Case RibTab_Tex_NewsLe:
            If Right$(FiNam, 4) <> ".txn" Then FiNam = FiNam & ".txn"
    End Select

    With clFil
        If .FilVor(FiNam) = True Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                .DaLoe = FiNam & vbNullChar
                .FilLoe
            End If
        Else
            Frage = 6
        End If
    End With
    
    If Frage = 6 Then
        TxCoN.Save FiNam, 0, 3
        DoEvents
    End If

    GlTSV = False 'Speichern Textverarbeitung
    GlTxN = False 'Kein neues Dokument mehr
    GlTxS = False
Case Tex_DatSav:
    DaNam = SDaNa()
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Bitte Name und Ordner der Datei angeben"
        Select Case GlBut
        Case RibTab_Tex_Dokumt:
                .DefaultExt = "*.txm"
                .Filter = "Textverarbeitung (*.txm)|*.txm|Alle Dateien (*.*)|*.*"
        Case RibTab_Tex_Vorlag:
                .DefaultExt = "*.txm"
                .Filter = "Textverarbeitung (*.txm)|*.txm|Alle Dateien (*.*)|*.*"
        Case RibTab_Tex_Rezept:
                .DefaultExt = "*.txr"
                .Filter = "Langrezepte (*.txr)|*.txr|Alle Dateien (*.*)|*.*"
        Case RibTab_Tex_NewsLe:
                .DefaultExt = "*.txn"
                .Filter = "Newslettervorlage (*.txn)|*.txn|Alle Dateien (*.*)|*.*"
        End Select
        .FileName = GlVor & DaNam
        .InitDir = GlVor
        .ShowSave
        FiNam = .FileName
        If .FileTitle = vbNullString Then
            Set CoDia = Nothing
            Set clFil = Nothing
            Exit Sub
        End If
    End With
    
    Select Case GlBut
    Case RibTab_Tex_Dokumt:
            If Right$(FiNam, 4) <> ".txm" Then FiNam = FiNam & ".txm"
    Case RibTab_Tex_Vorlag:
            If Right$(FiNam, 4) <> ".txm" Then FiNam = FiNam & ".txm"
    Case RibTab_Tex_Rezept:
            If Right$(FiNam, 4) <> ".txr" Then FiNam = FiNam & ".txr"
    Case RibTab_Tex_NewsLe:
            If Right$(FiNam, 4) <> ".txn" Then FiNam = FiNam & ".txn"
    End Select
    
    With clFil
        If .FilVor(FiNam) = True Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                .DaLoe = FiNam & vbNullChar
                .FilLoe
            End If
        Else
            Frage = 6
        End If
    End With

    If Frage = 6 Then
        TxCoN.Save FiNam, 0, 3
        DoEvents
    End If

    GlTSV = False 'Speichern Textverarbeitung
    GlTxN = False 'Kein neues Dokument mehr
    GlTxS = False
Case Tex_DaFeVe:
    GlAkt = True
    
    S_TxEin
    DoEvents 'Laden der Patientendaten in Array GlSer()
    
    STxV2 'Verbinden der Textfelder mit GlSer()
    DoEvents
    GlTSV = True 'Speichern Textverarbeitung

    GlAkt = False
Case Tex_DatLoe:
    If GlRch(0, 19) = 0 Then
        WindowMess "Sie besitzen keine Berechtigung für diesen Vorgang", Dial3, "Entfernen", FM.hwnd
        Exit Sub
    End If
    Frage = WindowMess(Mld2, Dial1, Tit2, FM.hwnd)
    If Frage = 6 Then
        If LiIts.Count > 0 Then
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    DaNam = LiItm.Tag
                    FiNam = GlDox & DaNam
                    Lange = Len(DaNam)
                    DaNaO = Left$(DaNam, Lange - 3) & "bmp"
                    If Mid$(DaNam, 9, 1) = "_" And Lange > 30 Then
                        TmGui = Mid$(DaNam, 10, 32)
                    End If
                    LiIts.Remove LiItm.Index
                    Exit For
                End If
            Next LiItm
            
            If clFil.FilVor(FiNam) = True Then
                clFil.DaLoe = FiNam & vbNullChar
                clFil.FilLoe
                DoEvents
            End If

            If clFil.FilVor(GlDox & DaNaO) = True Then
                clFil.DaLoe = GlDox & DaNaO & vbNullChar
                clFil.FilLoe
                DoEvents
            End If

            If TmGui <> vbNullString Then
                DBCmEx1 "qrySimAbLoG", "@IdStr", TmGui
                DoEvents
            End If
            
            If LiIts.Count > 0 Then
                LiIts(1).Selected = True
                STxLa
            Else
                STxNe
            End If
            DoEvents
            GlSav = False
        End If
    End If
Case Tex_DatKop:
    TeTit = "Dokument Kopieren"
    TeMai = "Soll das markierte Dokument wirklich kopiert werden?"
    TeInh = "Bei diesem Vorgang wird ein Duplikat des markierten Dokuments angefertigt und zur Bearbeitung geöffnet."
    TeFus = "Für das kopierte Dokument wird automatisch ein Verweis in das Krankenblatt des Patienten hinzugefügt."
    If GlDoK = True Then 'Dokument Kopieren Nachfrage
        SMeFr TeTit, TeMai, TeInh, TeInh, True, 0, False, FM.hwnd
        If GlMso = True Then 'Mesage Dialog Option
            IniSetVal "System", "DocKop", 0
            GlDoK = False
            GlMso = False
        End If
    Else
        GlMes = 33565
    End If

    If GlMes = 33565 Then 'Ja
        If LiIts.Count > 0 Then
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    FiNam = GlDox & LiItm.Tag
                    LiIts.Remove LiItm.Index
                    Exit For
                End If
            Next LiItm
            If clFil.FilVor(FiNam) = True Then
                clFil.FilPfa FiNam
                DaPfa = clFil.DaPfa & "\"
                
                TmGui = CreateID("D")
                PaStr = Format$(GlAdr, "000000")

                SuStr = InputBox("Bitte geben Sie einen Kommentar ein:", "Kommentar", "Textdokument")
                If SuStr <> vbNullString Then
                    SuStr = SUmw(SuStr, True, False, True, True)
                    DaNam = "TD" & PaStr & "_" & TmGui & "_" & SuStr & ".txm"
                    NeuNa = DaPfa & DaNam

                    clFil.DaCop = FiNam & ";" & NeuNa & vbNullChar
                    If clFil.FilCop(1) = True Then
                        DoEvents
                        If DaNam <> vbNullString Then
                            GlNeK = GlKoX
                            With GlNeK
                                .PatNr = GlAdr
                                .IdxNr = 0
                                .EiDat = Format$(Date, "dd.mm.yyyy")
                                .EiZei = TimeValue(Now)
                                .EiTyp = 24 'Textdokument
                                .KoStr = DaNam
                                .KoGui = TmGui
                                .TeStr = SuStr
                                .NeuEi = True
                                .Mitar = GlMiA(GlSmI, 2)
                            End With
                            K_Einf
                        End If
                        DoEvents
                        S_KrLa
                    End If
                End If
            End If
            DoEvents
            GlSav = False
        End If
    End If
Case Tex_Eigens:
    If LiIts.Count > 0 Then
        For Each LiItm In LiIts
            If LiItm.Selected = True Then
                DaNam = LiItm.Tag
                FiNam = GlDox & DaNam
                Exit For
            End If
        Next LiItm
        If clFil.FilVor(FiNam) = True Then
            clFil.FilInfo FM.hwnd, FiNam
        End If
    End If
Case Tex_DocVor:

Case Tex_DocDru:
    If GlTxM = False Then 'Serienbriefmodus
        TxCoN.PrintDialog GlTDa
    Else
        If GlTxS = True Then 'Seriendokument
            TxCoN.PrintDialog GlTDa
        Else
            frmTxStat.Show vbModal
        End If
    End If
Case Tex_NweSen:
    frmTxStat.EmSen = True
    frmTxStat.EmTes = False
    frmTxStat.Show vbModal
Case Tex_NweVor:
    frmTxStat.EmSen = True
    frmTxStat.EmTes = True
    frmTxStat.Show vbModal
End Select

Set CoDia = Nothing

Set clFil = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FObje " & Err.Number
Resume Next

End Sub
Private Sub FPrnt()
On Error GoTo AnErr

Select Case GlBut
Case RibTab_Adressen:
            SDrLis 1
Case RibTab_Mandanten:
            SDrLis 1
Case RibTab_Verordner:
            SDrLis 1
Case RibTab_Mitarbeit:
            SDrLis 1
Case RibTab_Fragebogen:
            SDruck "AnBog", True
Case RibTab_Krankenbla:
            SKrDr
Case RibTab_Abrechnung:
            SDrLis 2
Case RibTab_Tagesproto:
Case RibTab_Vorbereit:
Case RibTab_Rezeptmodul:
            SDrLis 6
Case RibTab_Belegmodul:
            SDrLis 6
Case RibTab_Bildmodul:

Case RibTab_Rechnungen:
            SDrLis 2
Case RibTab_Mahnwesen:
            SDrLis 3
Case RibTab_Buchungen:
            frmZeitraum.Show vbModal
Case RibTab_HomeBanki:
            
Case RibTab_Statistik:
            
Case RibTab_Ter_Kalend:
            STePr
Case RibTab_Ter_Listen:
            SDrLis 1
Case RibTab_Ter_Akont:
            SDrLis 1
Case RibTab_Ter_Warte:
            SDrLis 1
Case RibTab_Ter_Raeume:
            SDrLis 1
Case RibTab_Ter_Mitarb:
            SDrLis 1
Case RibTab_LabBericht:
            SDruck "LabKom", True
Case RibTab_LabAuftrag:
            SDruck "LabAuf", True
Case RibTab_LabBerichte:
            SDruck "LabKom", True
Case RibTab_LabAuftrage:
            SDruck "LabAuf", True
Case RibTab_Tex_Dokumt:
            FObje Tex_DocDru
Case RibTab_Tex_Vorlag:
            FObje Tex_DocDru
Case RibTab_Tex_Rezept:
            FObje Tex_DocDru
Case RibTab_Tex_NewsLe:
            FObje Tex_DocDru
Case RibTab_Tex_Email:
            S_MaAbr
End Select

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPrnt " & Err.Number
Resume Next

End Sub
Private Sub FRpPr()
On Error GoTo OrErr
'ReportControl Drucken

Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCo9 As XtremeReportControl.ReportControl

Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set RpCo6 = FM.repCont6
Set RpCoK = FM.repContK
Set RpCo8 = FM.repCont8
Set RpCo9 = FM.repCont9

Select Case GlBut
Case RibTab_Adressen: RpCo2.PrintPreview True
Case RibTab_Mandanten: RpCo2.PrintPreview True
Case RibTab_Verordner: RpCo2.PrintPreview True
Case RibTab_Mitarbeit: RpCo2.PrintPreview True
Case RibTab_Fragebogen: RpCo5.PrintPreview True
Case RibTab_Krankenbla: RpCoK.PrintPreview True
Case RibTab_Abrechnung: RpCo6.PrintPreview True
Case RibTab_Vorbereit: RpCo6.PrintPreview True
Case RibTab_Tagesproto: RpCo6.PrintPreview True
Case RibTab_Rezeptmodul: RpCo3.PrintPreview True
Case RibTab_Belegmodul: RpCo3.PrintPreview True
Case RibTab_Rechnungen: RpCo4.PrintPreview True
Case RibTab_Mahnwesen: RpCo1.PrintPreview True
Case RibTab_Buchungen: RpCo1.PrintPreview True
Case RibTab_HomeBanki: RpCo1.PrintPreview True
Case RibTab_Ter_Listen: RpCo1.PrintPreview True
Case RibTab_Ter_Akont: RpCo1.PrintPreview True
Case RibTab_Ter_Warte: RpCo1.PrintPreview True
Case RibTab_LabBericht: RpCo5.PrintPreview True
Case RibTab_LabAuftrag: RpCo5.PrintPreview True
Case RibTab_LabBerichte: RpCo1.PrintPreview True
Case RibTab_LabAuftrage: RpCo1.PrintPreview True
Case RibTab_Kat_Eintrg: RpCo8.PrintPreview True
Case RibTab_Kat_Ketten: RpCo8.PrintPreview True
Case RibTab_Kat_Frage: RpCo8.PrintPreview True
Case RibTab_Tex_Email: RpCo8.PrintPreview True
Case RibTab_Tex_Dokumt:
Case RibTab_Tex_Vorlag:
Case RibTab_Tex_Rezept:
Case RibTab_Tex_NewsLe:
Case RibTab_Kat_Explor:
End Select

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRpPr " & Err.Number
Resume Next

End Sub
Private Sub FRzAb()
On Error GoTo OrErr
'Blendet den Tezeptabsender ein / aus

Dim ChAbs As XtremeSuiteControls.CheckBox
Dim TxAb1 As XtremeSuiteControls.FlatEdit
Dim TxAb2 As XtremeSuiteControls.FlatEdit

Set FM = frmMain
Set ChAbs = FM.chkAbsen
Set TxAb1 = FM.txtAbsNa
Set TxAb2 = FM.txtAbsAd

If ChAbs.Value = 1 Then
    TxAb1.Visible = False
    TxAb2.Visible = False
    IniSetVal "System", "RzAbse", -1
Else
    TxAb1.Visible = True
    TxAb2.Visible = True
    IniSetVal "System", "RzAbse", 0
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRzAb " & Err.Number
Resume Next

End Sub

Private Sub FRzBe()
On Error GoTo OrErr
'Lädt den gewünschten Beleg

Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case GlBut
Case RibTab_Rezeptmodul: Set CmCom = CmBrs.FindControl(CmCom, SY_RZ_Rezept_Vorlage, , True)
Case RibTab_Belegmodul: Set CmCom = CmBrs.FindControl(CmCom, SY_RZ_Beleg_Vorlage, , True)
Case RibTab_Tex_Rezept: Set CmCom = CmBrs.FindControl(CmCom, Tex_DatVor, , True)
End Select

LiIdx = CmCom.ListIndex

Select Case GlBut
Case RibTab_Rezeptmodul: IniSetVal "System", "RzForm", "O" & Format$(LiIdx, "00")
Case RibTab_Belegmodul: IniSetVal "System", "BlForm", "O" & Format$(LiIdx, "00")
Case RibTab_Tex_Rezept: IniSetVal "System", "LrForm", "O" & Format$(LiIdx, "00")
End Select

Select Case GlBut
Case RibTab_Rezeptmodul:
        SRzBe
        S_RzSav
Case RibTab_Belegmodul:
        SRzBe
        S_RzSav
Case RibTab_Tex_Rezept:
        STxRz LiIdx
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRzBe " & Err.Number
Resume Next

End Sub
Private Sub FSam()
On Error GoTo InErr
'Sämtliche Einträge zeigen (F6)

Select Case GlBut
Case RibTab_Adressen:
        FSuAu
Case RibTab_Mandanten:
Case RibTab_Verordner:
Case RibTab_Mitarbeit:
Case RibTab_Fragebogen:
Case RibTab_Krankenbla:
        KrMain 26
Case RibTab_Abrechnung:
        SReDi 1
Case RibTab_Tagesproto:
        FSuAu
Case RibTab_Vorbereit:
        SVoAb
Case RibTab_Belegmodul:

Case RibTab_Rezeptmodul:
Case RibTab_Bildmodul:
        SBiIo
Case RibTab_Rechnungen:
        FSuAu
Case RibTab_Mahnwesen:
        FSuAu
Case RibTab_Buchungen:
        FSuAu
Case RibTab_HomeBanki:
        FSuAu
Case RibTab_Statistik:
        
Case RibTab_Ter_Kalend:
    
Case RibTab_Ter_Listen:
        FSuAu
Case RibTab_Ter_Akont:
        FSuAu
Case RibTab_Ter_Warte:
        FSuAu
Case RibTab_Ter_Raeume:
Case RibTab_Ter_Mitarb:
Case RibTab_LabBerichte:
        FSuAu
Case RibTab_LabAuftrage:
        FSuAu
Case RibTab_Kat_Eintrg:
        FSuAu
Case RibTab_Kat_Ketten:
        FSuAu
Case RibTab_Kat_Frage:
        FSuAu
Case RibTab_Tex_Email:
        FSuAu
Case RibTab_Kat_Explor:

End Select

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSam & Err.Number"
Resume Next

End Sub
Private Sub FSeAr(ByVal SelOp As Boolean)
On Error GoTo OrErr

Dim GesZa As Long
Dim AktZa As Long
Dim RpCo9 As XtremeReportControl.ReportControl

Set RpCo9 = Me.repCont9
Set RpRcs = RpCo9.Records

GesZa = RpRcs.Count

If GesZa > 0 Then
    For AktZa = 0 To GesZa - 1
        SeAry(4, AktZa) = SelOp
    Next AktZa
End If

RpCo9.Populate

Set RpRcs = Nothing
Set RpCo9 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRzBe " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub FSuAu()
On Error GoTo OrErr
'Aufheben des Suchergebnissen

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo0 = FM.repCont0
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo4 = FM.repCont4
Set RpCo8 = FM.repCont8
Set RpCo5 = FM.repCont5
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Select Case GlBut
Case RibTab_Adressen:
        Set CmCom = CmBrs.FindControl(CmCom, SY_AD_Adresse_SortFeld, , True)
        GlSuP = GlSuX
        CmCom.ListIndex = 1
        SSuAu
Case RibTab_Mandanten:
        GlSuP = GlSuX
        SSuAu
Case RibTab_Verordner:
        GlSuP = GlSuX
        SSuAu
Case RibTab_Mitarbeit:
        GlSuP = GlSuX
        SSuAu
Case RibTab_Tagesproto:

Case RibTab_Rechnungen:
        GlSuR = GlSuX
Case RibTab_Mahnwesen:
        GlSuM = GlSuX
Case RibTab_Buchungen:
        GlSuB = GlSuX
Case RibTab_HomeBanki:
        GlSuH = GlSuX
Case RibTab_Ter_Listen:
        GlSuT = GlSuX
Case RibTab_Ter_Akont:
        GlSuT = GlSuX
Case RibTab_Ter_Warte:
        GlSuT = GlSuX
Case RibTab_LabBerichte:
        GlSuL = GlSuX
Case RibTab_LabAuftrage:
        GlSuU = GlSuX
Case RibTab_Kat_Eintrg:
        GlSuE = GlSuX
        SSuAu
Case RibTab_Kat_Ketten:
        GlSuN = GlSuX
        SSuAu
Case RibTab_Kat_Frage:
        GlSuG = GlSuX
        SSuAu
Case RibTab_Tex_Email:
        GlSuI = GlSuX
Case RibTab_Kat_Explor:

End Select

DoEvents
SSuch
DoEvents

Select Case GlBut
Case RibTab_Adressen: RpCo2.SetFocus
Case RibTab_Mandanten: RpCo2.SetFocus
Case RibTab_Verordner: RpCo2.SetFocus
Case RibTab_Mitarbeit: RpCo2.SetFocus
Case RibTab_Rechnungen: RpCo4.SetFocus
Case RibTab_Mahnwesen: RpCo1.SetFocus
Case RibTab_Buchungen: RpCo1.SetFocus
Case RibTab_Ter_Listen: RpCo1.SetFocus
Case RibTab_Ter_Akont: RpCo1.SetFocus
Case RibTab_Ter_Warte: RpCo1.SetFocus
Case RibTab_LabBerichte: RpCo1.SetFocus
Case RibTab_LabAuftrage: RpCo1.SetFocus
Case RibTab_Kat_Eintrg: RpCo8.SetFocus
Case RibTab_Kat_Ketten: RpCo8.SetFocus
Case RibTab_Kat_Frage: RpCo8.SetFocus
Case RibTab_Tex_Email: RpCo0.SetFocus
Case RibTab_Kat_Explor:
End Select

SPopu "Alle Einträge anzeigen", "HINWEIS! Es werden wieder alle Einträge angezeigt", IC48_Information

Set RpCo0 = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo8 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuAu " & Err.Number
Resume Next

End Sub
Private Sub FSuch(Optional ByVal GldKo As Boolean = False)
On Error GoTo OrErr
'Sucheingabe

Dim ManNr As Long
Dim MitNr As Long
Dim RmuNr As Long
Dim NeuDa As Date
Dim SuBet As Double
Dim BuJah As String
Dim SuStr As String
Dim TyStr As String
Dim SuMon As Integer
Dim SuWek As Integer
Dim SuAbg As Integer
Dim SuSta As Integer
Dim LiIdx As Integer
Dim IdZug As Integer
Dim TyIdx As Integer
Dim KaIdx As Integer
Dim AktZa As Integer
Dim SuAbs As Boolean
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmDat As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmKat As XtremeCommandBars.CommandBarComboBox
Dim CmTyp As XtremeCommandBars.CommandBarComboBox
Dim CmAbg As XtremeCommandBars.CommandBarComboBox
Dim CmSta As XtremeCommandBars.CommandBarComboBox
Dim CmTSt As XtremeCommandBars.CommandBarComboBox
Dim CmZug As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set RpCo0 = FM.repCont0
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set RpCo8 = FM.repCont8
Set DaPi4 = FM.dtpDatu4
Set CmBrs = FM.comBar01

GlReF = vbNullString 'Rechnungsfilter
Set CmCom = CmBrs.FindControl(CmCom, SY_SuJah, , True)
BuJah = CmCom.Text

Set CmCom = CmBrs.FindControl(CmCom, SY_SuMan, , True)
ManNr = CmCom.ItemData(CmCom.ListIndex)

Set CmCom = CmBrs.FindControl(CmCom, SY_SuMit, , True)
MitNr = CmCom.ItemData(CmCom.ListIndex)

Set CmCom = CmBrs.FindControl(CmCom, SY_SuRau, , True)
If GlRaV = True Then
    RmuNr = CmCom.ItemData(CmCom.ListIndex)
Else
    RmuNr = 0
End If

Set CmEdi = CmBrs.FindControl(CmEdi, SY_SuTex, , True)
Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)
If CmDat.Text <> vbNullString Then
    If IsDate(CmDat.Text) Then
        NeuDa = CmDat.Text
    Else
        CmDat.Text = Date
        NeuDa = Date
    End If
Else
    CmDat.Text = Date
    NeuDa = Date
End If

With DaPi4
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

SSuAu 'Hebt die markierten Suchbuchstaben wieder auf
DoEvents

If CmEdi.Visible = True Then
    If CmEdi.Text <> vbNullString Then
        SuStr = CmEdi.Text
        SuStr = Trim$(SuStr)
    Else
        Select Case GlBut
        Case RibTab_Adressen: GlSuP = GlSuX
        Case RibTab_Mandanten: GlSuP = GlSuX
        Case RibTab_Verordner: GlSuP = GlSuX
        Case RibTab_Mitarbeit: GlSuP = GlSuX
        Case RibTab_Rechnungen: GlSuR = GlSuX
        Case RibTab_Mahnwesen: GlSuM = GlSuX
        Case RibTab_Buchungen: GlSuB = GlSuX
        Case RibTab_HomeBanki: GlSuH = GlSuX
        Case RibTab_Ter_Listen: GlSuT = GlSuX
        Case RibTab_Ter_Akont: GlSuT = GlSuX
        Case RibTab_Ter_Warte: GlSuT = GlSuX
        Case RibTab_LabBerichte: GlSuL = GlSuX
        Case RibTab_LabAuftrage: GlSuU = GlSuX
        Case RibTab_Kat_Eintrg: GlSuE = GlSuX
        Case RibTab_Kat_Ketten: GlSuN = GlSuX
        Case RibTab_Kat_Frage: GlSuG = GlSuX
        Case RibTab_Tex_Email:  GlSuI = GlSuX
        Case RibTab_Kat_Explor:
        End Select
        SSuch
        Select Case GlBut
        Case RibTab_Adressen: RpCo2.SetFocus
        Case RibTab_Mandanten: RpCo2.SetFocus
        Case RibTab_Verordner: RpCo2.SetFocus
        Case RibTab_Mitarbeit: RpCo2.SetFocus
        Case RibTab_Rechnungen: RpCo4.SetFocus
        Case RibTab_Mahnwesen: RpCo1.SetFocus
        Case RibTab_Buchungen: RpCo1.SetFocus
        Case RibTab_Statistik:
        Case RibTab_Ter_Listen: RpCo1.SetFocus
        Case RibTab_Ter_Akont: RpCo1.SetFocus
        Case RibTab_Ter_Warte: RpCo1.SetFocus
        Case RibTab_LabBerichte: RpCo1.SetFocus
        Case RibTab_LabAuftrage: RpCo1.SetFocus
        Case RibTab_Kat_Eintrg: RpCo8.SetFocus
        Case RibTab_Kat_Ketten: RpCo8.SetFocus
        Case RibTab_Kat_Frage: RpCo8.SetFocus
        Case RibTab_Tex_Email: RpCo0.SetFocus
        Case RibTab_Kat_Explor:
        End Select
        SPopu "Alle Einträge anzeigen", "HINWEIS! Es werden wieder alle Einträge angezeigt", IC48_Information
        Exit Sub
    End If
End If

Select Case GlBut
Case RibTab_Adressen:
        Set CmCom = CmBrs.FindControl(CmCom, SY_AD_Adresse_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuP
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            With GlSuP
                .SuIdx = 4
                If IsNumeric(SuStr) Then
                    .SuNum = CLng(SuStr)
                Else
                    .SuNum = 0
                End If
            End With
        Case 3:
            With GlSuP
                .SuIdx = 2
                .SuStr = SuStr
            End With
        Case 4:
            With GlSuP
                .SuIdx = 3
                .SuStr = SuStr
            End With
        Case 5:
            With GlSuP
                .SuIdx = 10
                .SuDat = NeuDa
            End With
        Case 6: 'Mandanten
            With GlSuP
                .SuIdx = 6
                .SuMan = ManNr
            End With
        Case 7:
            With GlSuP
                .SuIdx = 12
                .SuStr = SuStr
            End With
        Case 8:
            With GlSuP
                .SuIdx = 13
                .SuStr = SuStr
            End With
        Case 9:
            With GlSuP
                .SuIdx = 14
                .SuStr = SuStr
            End With
        Case 10:
            With GlSuP
                .SuIdx = 15
                .SuStr = SuStr
            End With
        Case 11:
            With GlSuP
                .SuIdx = 16
                .SuStr = SuStr
            End With
        Case 12:
            With GlSuP
                .SuIdx = 17
                .SuStr = SuStr
            End With
        Case 13:
            With GlSuP
                .SuIdx = 18
                .SuStr = SuStr
            End With
        End Select
Case RibTab_Mandanten:
        Set CmCom = CmBrs.FindControl(CmCom, SY_MA_Mandant_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuP
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            With GlSuP
                .SuIdx = 2
                .SuStr = SuStr
            End With
        End Select
Case RibTab_Verordner:
        Set CmCom = CmBrs.FindControl(CmCom, SY_VE_Verord_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuP
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            With GlSuP
                .SuIdx = 2
                .SuStr = SuStr
            End With
        End Select
Case RibTab_Mitarbeit:
        Set CmCom = CmBrs.FindControl(CmCom, SY_MI_Mitarb_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuP
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            With GlSuP
                .SuIdx = 2
                .SuStr = SuStr
            End With
        End Select
Case RibTab_Rechnungen:
        Set CmCom = CmBrs.FindControl(CmCom, SY_RE_Rechnung_SuchCombp, , True)
        LiIdx = CmCom.ListIndex
        Set CmTyp = CmBrs.FindControl(CmTyp, SY_RE_Rechnung_Belegtyp, , True)
        TyIdx = CmTyp.ListIndex
        Select Case TyIdx
        Case 1: TyStr = "R"
        Case 2: TyStr = "V"
        Case 3: TyStr = "L"
        Case 4: TyStr = "A"
        Case 5: TyStr = "U"
        Case 6: TyStr = "M"
        Case 7: TyStr = "G"
        Case 8: TyStr = "I"
        Case 9: TyStr = "Y"
        Case 10: TyStr = "X"
        End Select
        Select Case LiIdx
        Case 1:
            With GlSuR
                .SuIdx = 1
                .SuStr = SuStr
                .SuTyp = TyStr
            End With
        Case 2:
            With GlSuR
                .SuIdx = 2
                .SuStr = SuStr
                .SuTyp = TyStr
            End With
        Case 3:
            With GlSuR
                .SuIdx = 3
                .SuDat = NeuDa
                .SuJah = BuJah
                .SuTyp = TyStr
            End With
        Case 4:
            With GlSuR
                .SuIdx = 4
                .SuStr = SuStr
                .SuJah = BuJah
                .SuTyp = TyStr
            End With
        Case 5:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuMon, , True)
            SuMon = CmCom.ItemData(CmCom.ListIndex)
            With GlSuR
                .SuIdx = 5
                .SuMon = SuMon
                .SuJah = BuJah
                .SuTyp = TyStr
            End With
        Case 6:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuWek, , True)
            SuWek = CmCom.ItemData(CmCom.ListIndex)
            With GlSuR
                .SuIdx = 6
                .SuWek = SuWek
                .SuJah = BuJah
                .SuTyp = TyStr
            End With
        Case 7: 'Mandanten
            With GlSuR
                .SuIdx = 7
                .SuJah = BuJah
                .SuMan = ManNr
                .SuTyp = TyStr
            End With
        Case 8: 'Mitarbeiter
            With GlSuR
                .SuIdx = 8
                .SuJah = BuJah
                .SuMit = MitNr
                .SuTyp = TyStr
            End With
        Case 9:
            With GlSuR
                .SuIdx = 9
                .SuStr = SuStr
                .SuTyp = TyStr
            End With
        Case 10:
            With GlSuR
                .SuIdx = 10
                .SuStr = SuStr
                .SuTyp = TyStr
            End With
        Case 11:
            If IsNumeric(SuStr) Then
                SuBet = CDbl(SuStr)
            Else
                SuBet = 0
            End If
            With GlSuR
                .SuIdx = 11
                .SuBet = SuBet
                .SuTyp = TyStr
            End With
        End Select
Case RibTab_Mahnwesen:
        Set CmCom = CmBrs.FindControl(CmCom, SY_PO_Posten_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuM
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            With GlSuM
                .SuIdx = 2
                .SuStr = SuStr
            End With
        Case 3:
            With GlSuM
                .SuIdx = 3
                .SuDat = NeuDa
                .SuJah = BuJah
            End With
        Case 4: 'Lastschriften
            With GlSuM
                .SuIdx = 4
                .SuStr = vbNullString
            End With
        Case 5:
            If IsNumeric(SuStr) Then
                SuBet = CDbl(SuStr)
            Else
                SuBet = 0
            End If
            With GlSuM
                .SuIdx = 5
                .SuBet = SuBet
            End With
        Case 6: 'Mandanten
            With GlSuM
                .SuIdx = 6
                .SuMan = ManNr
            End With
        Case 7: 'Mitarbeiter
            With GlSuM
                .SuIdx = 7
                .SuMit = MitNr
            End With
        End Select
Case RibTab_Buchungen:
        Set CmCom = CmBrs.FindControl(CmCom, SY_BU_Buchung_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuB
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            With GlSuB
                .SuIdx = 2
                .SuStr = SuStr
            End With
        Case 3:
            With GlSuB
                .SuIdx = 3
                .SuStr = SuStr
            End With
        Case 4:
            With GlSuB
                .SuIdx = 4
                .SuDat = NeuDa
                .SuJah = BuJah
            End With
        Case 5:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuMon, , True)
            SuMon = CmCom.ItemData(CmCom.ListIndex)
            With GlSuB
                .SuIdx = 5
                .SuMon = SuMon
                .SuJah = BuJah
            End With
        Case 6:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuWek, , True)
            SuWek = CmCom.ItemData(CmCom.ListIndex)
            With GlSuB
                .SuIdx = 6
                .SuWek = SuWek
                .SuJah = BuJah
            End With
        Case 7: 'Mandanten
            With GlSuB
                .SuIdx = 7
                .SuJah = BuJah
                .SuMan = ManNr
            End With
        Case 8: 'Mitarbeiter
            With GlSuB
                .SuIdx = 8
                .SuJah = BuJah
                .SuMit = MitNr
            End With
        Case 9:
            With GlSuB
                .SuIdx = 9
                .SuStr = SuStr
            End With
        Case 10:
            If IsNumeric(SuStr) Then
                SuBet = CDbl(SuStr)
            Else
                SuBet = 0
            End If
            With GlSuB
                .SuIdx = 10
                .SuBet = SuBet
            End With
        Case 11:
            If IsNumeric(SuStr) Then
                SuBet = CDbl(SuStr)
            Else
                SuBet = 0
            End If
            With GlSuB
                .SuIdx = 11
                .SuBet = SuBet
            End With
        End Select
        DoEvents
        If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
            If GlBuc = False Then 'einfache Buchhaltung verwenden
                If LiIdx = 7 Then
                    If GldKo = True Then
                        S_BaCm True
                    End If
                End If
            End If
        End If
Case RibTab_HomeBanki:
        Set CmCom = CmBrs.FindControl(CmCom, SY_BA_Banking_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuH
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            If IsNumeric(SuStr) Then
                SuBet = CDbl(SuStr)
            Else
                SuBet = 0
            End If
            With GlSuH
                .SuIdx = 2
                .SuBet = SuBet
            End With
        Case 3:
            With GlSuH
                .SuIdx = 3
                .SuDat = NeuDa
                .SuJah = BuJah
            End With
        Case 4:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuMon, , True)
            SuMon = CmCom.ItemData(CmCom.ListIndex)
            With GlSuH
                .SuIdx = 4
                .SuMon = SuMon
                .SuJah = BuJah
            End With
        Case 5:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuWek, , True)
            SuWek = CmCom.ItemData(CmCom.ListIndex)
            With GlSuH
                .SuIdx = 5
                .SuWek = SuWek
                .SuJah = BuJah
            End With
        Case 6: 'Mandanten
            With GlSuH
                .SuIdx = 6
                .SuJah = BuJah
                .SuMan = ManNr
            End With
        Case 7: 'Zugeordnet
            Set CmZug = CmBrs.FindControl(CmZug, SY_SuZug, , True)
            IdZug = CmZug.ListIndex
            With GlSuH
                .SuIdx = 7
                .SuJah = BuJah
                .SuAbg = IdZug
                Select Case IdZug
                Case 1: .SuAbs = False
                Case 2: .SuAbs = True
                Case 3: .SuAbs = False
                Case 4: .SuAbs = True
                End Select
            End With
        End Select
Case RibTab_Ter_Listen:
        Set CmCom = CmBrs.FindControl(CmCom, SY_TL_Terminliste_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuT
                .SuIdx = 1
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 2:
            With GlSuT
                .SuIdx = 2
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 3:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuMon, , True)
            SuMon = CmCom.ItemData(CmCom.ListIndex)
            With GlSuT
                .SuIdx = 3
                .SuMon = SuMon
                .SuJah = BuJah
            End With
        Case 4:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuWek, , True)
            SuWek = CmCom.ItemData(CmCom.ListIndex)
            With GlSuT
                .SuIdx = 4
                .SuWek = SuWek
                .SuJah = BuJah
            End With
        Case 5:
            With GlSuT
                .SuIdx = 5
                .SuDat = NeuDa
                .SuJah = BuJah
            End With
        Case 6: 'Mandanten
            With GlSuT
                .SuIdx = 6
                .SuMan = ManNr
                .SuJah = BuJah
            End With
        Case 7: 'Räume
            With GlSuT
                .SuIdx = 7
                .SuRmu = RmuNr
                .SuJah = BuJah
            End With
        Case 8:
            With GlSuT
                .SuIdx = 8
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 9: 'Mitarbeiter
            With GlSuT
                .SuIdx = 9
                .SuMit = MitNr
                .SuJah = BuJah
            End With
        Case 10:
            With GlSuT
                .SuIdx = 10
                .SuDat = NeuDa
                .SuJah = BuJah
            End With
        Case 11:
            With GlSuT
                .SuIdx = 11
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 12:
            Set CmAbg = CmBrs.FindControl(CmAbg, SY_SuAbg, , True)
            SuAbg = CmAbg.ItemData(CmAbg.ListIndex)
            With GlSuT
                .SuIdx = 12
                .SuAbg = SuAbg
                .SuJah = BuJah
            End With
        Case 13: 'Abgehakt
            Set CmSta = CmBrs.FindControl(CmSta, SY_SuSta, , True)
            SuAbs = CBool(CmSta.ItemData(CmSta.ListIndex))
            With GlSuT
                .SuIdx = 13
                .SuAbs = SuAbs
                .SuJah = BuJah
            End With
        Case 14: 'Terminstatus
            Set CmTSt = CmBrs.FindControl(CmTSt, SY_SuTSt, , True)
            SuSta = CInt(CmTSt.ItemData(CmTSt.ListIndex))
            With GlSuT
                .SuIdx = 14
                .SuSta = SuSta
                .SuJah = BuJah
            End With
        End Select
Case RibTab_Ter_Akont:
        Set CmCom = CmBrs.FindControl(CmCom, SY_TL_Terminliste_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuT
                .SuIdx = 1
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 2:
            With GlSuT
                .SuIdx = 2
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 3:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuMon, , True)
            SuMon = CmCom.ItemData(CmCom.ListIndex)
            With GlSuT
                .SuIdx = 3
                .SuMon = SuMon
                .SuJah = BuJah
            End With
        Case 4:
            Set CmCom = CmBrs.FindControl(CmCom, SY_SuWek, , True)
            SuWek = CmCom.ItemData(CmCom.ListIndex)
            With GlSuT
                .SuIdx = 4
                .SuWek = SuWek
                .SuJah = BuJah
            End With
        Case 5:
            With GlSuT
                .SuIdx = 5
                .SuDat = NeuDa
                .SuJah = BuJah
            End With
        Case 6: 'Mandnaten
            With GlSuT
                .SuIdx = 6
                .SuMan = ManNr
                .SuJah = BuJah
            End With
        Case 7:
            With GlSuT
                .SuIdx = 7
                .SuRmu = RmuNr
                .SuJah = BuJah
            End With
        Case 8:
            With GlSuT
                .SuIdx = 8
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 9: 'Mitarbeiter
            With GlSuT
                .SuIdx = 9
                .SuMit = MitNr
                .SuJah = BuJah
            End With
        Case 10:
            With GlSuT
                .SuIdx = 10
                .SuDat = NeuDa
                .SuJah = BuJah
            End With
        Case 11:
            With GlSuT
                .SuIdx = 11
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 12:
            Set CmAbg = CmBrs.FindControl(CmAbg, SY_SuAbg, , True)
            SuAbg = CmAbg.ItemData(CmAbg.ListIndex)
            With GlSuT
                .SuIdx = 12
                .SuAbg = SuAbg
                .SuJah = BuJah
            End With
        Case 13: 'Terminstatus
            Set CmSta = CmBrs.FindControl(CmSta, SY_SuSta, , True)
            SuAbs = CBool(CmSta.ItemData(CmSta.ListIndex))
            With GlSuT
                .SuIdx = 13
                .SuAbs = SuAbs
                .SuJah = BuJah
            End With
        Case 14: 'Terminstatus
            Set CmTSt = CmBrs.FindControl(CmTSt, SY_SuTSt, , True)
            SuSta = CInt(CmTSt.ItemData(CmTSt.ListIndex))
            With GlSuT
                .SuIdx = 14
                .SuSta = SuSta
                .SuJah = BuJah
            End With
        End Select
Case RibTab_LabBerichte:
        Set CmCom = CmBrs.FindControl(CmCom, SY_LB_Labor_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuL
                .SuIdx = 1
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 2:
            With GlSuL
                .SuIdx = 2
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 3: 'Mandnaten
            With GlSuL
                .SuIdx = 3
                .SuMan = ManNr
            End With
        Case 4:
            With GlSuL
                .SuIdx = 4
                .SuDat = CDate(SuStr)
            End With
        Case 5:
            With GlSuL
                .SuIdx = 5
                .SuNum = CLng(SuStr)
            End With
        End Select
Case RibTab_LabAuftrage:
        Set CmCom = CmBrs.FindControl(CmCom, SY_LA_Auftrag_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuU
                .SuIdx = 1
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 2:
            With GlSuU
                .SuIdx = 2
                .SuStr = SuStr
                .SuJah = BuJah
            End With
        Case 3: 'Mandanten
            With GlSuU
                .SuIdx = 3
                .SuMan = ManNr
            End With
        End Select
Case RibTab_Kat_Eintrg:
        Set CmCom = CmBrs.FindControl(CmCom, KA_Eint_SuchCombo, , True)
        Set CmKat = CmBrs.FindControl(CmKat, KA_Eint_KatCombo, , True)
        LiIdx = CmCom.ListIndex
        KaIdx = CmKat.ListIndex
        With GlSuE
            .SuIdx = LiIdx
            .SuStr = SuStr
            .SuKat = KaIdx
        End With
Case RibTab_Kat_Ketten:
        Set CmCom = CmBrs.FindControl(CmCom, KA_Kett_SuchCombo, , True)
        Set CmKat = CmBrs.FindControl(CmKat, KA_Kett_KatCombo, , True)
        LiIdx = CmCom.ListIndex
        KaIdx = CmKat.ListIndex
        With GlSuN
            .SuIdx = LiIdx
            .SuStr = SuStr
            .SuKat = KaIdx
        End With
Case RibTab_Kat_Frage:
        Set CmCom = CmBrs.FindControl(CmCom, KA_Frage_SuchCombo, , True)
        LiIdx = CmCom.ListIndex
        With GlSuG
            .SuIdx = LiIdx
            .SuStr = SuStr
            .SuKat = KaIdx
        End With
Case RibTab_Tex_Email:
        Set CmCom = CmBrs.FindControl(CmCom, TX_Mail_SucCombo, , True)
        LiIdx = CmCom.ListIndex
        Select Case LiIdx
        Case 1:
            With GlSuI
                .SuIdx = 1
                .SuStr = SuStr
            End With
        Case 2:
            With GlSuI
                .SuIdx = 2
                .SuStr = SuStr
            End With
        Case 3:
            With GlSuI
                .SuIdx = 3
                .SuStr = SuStr
            End With
        Case 4:
            With GlSuI
                .SuIdx = 4
                .SuDat = NeuDa
            End With
        Case 5:
            With GlSuI
                .SuIdx = 5
                .SuStr = SuStr
            End With
        Case 6:
            With GlSuI
                .SuIdx = 6
                .SuStr = SuStr
            End With
        End Select
End Select

DoEvents
SSuch
DoEvents

Select Case GlBut
Case RibTab_Adressen:
        If GlAdA > 0 Then 'Anzahl gefundener Adressen
            RpCo2.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Mandanten:
        If GlAdA > 0 Then 'Anzahl gefundener Adressen
            RpCo2.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Verordner:
        If GlAdA > 0 Then 'Anzahl gefundener Adressen
            RpCo2.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Mitarbeit:
        If GlAdA > 0 Then 'Anzahl gefundener Adressen
            RpCo2.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Rechnungen:
        Set RpCo4 = FM.repCont4
        If RpCo4.Records.Count = 0 Then
            SPopu "Rechnung nicht gefunden", "Die von Ihnen gesuchte Rechnung, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo4.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Mahnwesen:
        Set RpCo1 = FM.repCont1
        If RpCo1.Records.Count = 0 Then
            SPopu "Offener Posten nicht gefunden", "Der von Ihnen gesuchte Offene Posten, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo1.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Buchungen:
        Set RpCo1 = FM.repCont1
        If RpCo1.Records.Count = 0 Then
            SPopu "Buchung nicht gefunden", "Die von Ihnen gesuchte Buchung, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo1.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_HomeBanki:
        Set RpCo1 = FM.repCont1
        If RpCo1.Records.Count = 0 Then
            SPopu "Buchung nicht gefunden", "Die von Ihnen gesuchte Buchung, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo1.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Ter_Listen:
        Set RpCo1 = FM.repCont1
        If RpCo1.Records.Count = 0 Then
            SPopu "Termin nicht gefunden", "Der von Ihnen gesuchte Termin, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo1.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Ter_Akont:
        Set RpCo1 = FM.repCont1
        If RpCo1.Records.Count = 0 Then
            SPopu "Termin nicht gefunden", "Der von Ihnen gesuchte Termin, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo1.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Ter_Warte:
        Set RpCo1 = FM.repCont1
        If RpCo1.Records.Count = 0 Then
            SPopu "Termin nicht gefunden", "Der von Ihnen gesuchte Termin, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo1.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_LabBerichte:
        CmEdi.Text = vbNullString
Case RibTab_LabBerichte:
        CmEdi.Text = vbNullString
Case RibTab_Kat_Eintrg:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Kat_Ketten:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Kette nicht gefunden", "Die von Ihnen gesuchte Kette, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Kat_Frage:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Tex_Email:
        If GlEmV > 0 Then
            RpCo0.SetFocus
        End If
        CmEdi.Text = vbNullString
Case RibTab_Kat_Explor:

End Select

Set CmBrs = Nothing
Set RpCo0 = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo8 = Nothing
Set DaPi4 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuch " & Err.Number
Resume Next

End Sub
Private Sub FSuFa()
On Error GoTo OrErr
'Favoriten Knopf

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo8 As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

If GlFav = False Then
    CmAcs(KA_Eint_Favoriten).Checked = True
    GlFav = True
    Select Case GlBut
    Case RibTab_Kat_Eintrg:
            With GlSuE
                .SuIdx = 5
            End With
    Case RibTab_Kat_Ketten:
            With GlSuN
                .SuIdx = 5
            End With
    Case RibTab_Kat_Frage:
            With GlSuG
                .SuIdx = 5
            End With
    End Select
Else
    CmAcs(KA_Eint_Favoriten).Checked = False
    GlFav = False
    Select Case GlBut
    Case RibTab_Kat_Eintrg:
            With GlSuE
                .SuIdx = 0
            End With
    Case RibTab_Kat_Ketten:
            With GlSuN
                .SuIdx = 0
            End With
    Case RibTab_Kat_Frage:
            With GlSuG
                .SuIdx = 0
            End With
    End Select
End If

DoEvents
SSuch
DoEvents

Select Case GlBut
Case RibTab_Kat_Eintrg:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
Case RibTab_Kat_Ketten:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Kette nicht gefunden", "Die von Ihnen gesuchte Kette, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
Case RibTab_Kat_Frage:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
Case RibTab_Tex_Email:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Email nicht gefunden", "Die von Ihnen gesuchte Email, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
End Select

Set CmBrs = Nothing
Set RpCo8 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFa " & Err.Number
Resume Next

End Sub

Private Sub FSuFe(Optional ByVal KeyPr As Boolean = False)
On Error GoTo OrErr
'Suchleiste einblenden oder Suchformular anzeigen

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmEdi As XtremeCommandBars.CommandBarEdit

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Set CmEdi = CmBrs.FindControl(CmEdi, SY_SuTex, , True)

Select Case GlBut
Case RibTab_Startseite:
            GlAdU = 1
            frmAdrSuch.Show vbModal
Case RibTab_Adressen:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSAd = False Then
                    SSuLe
                End If
            End If
            If GlSAd = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_AD_Adresse_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Mandanten:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSMn = False Then
                    SSuLe
                End If
            End If
            If GlSMn = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_MA_Mandant_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Verordner:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSVe = False Then
                    SSuLe
                End If
            End If
            If GlSVe = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_VE_Verord_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Mitarbeit:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSMr = False Then
                    SSuLe
                End If
            End If
            If GlSMr = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_MI_Mitarb_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Fragebogen:
            frmAdrSuch.Show vbModal
Case RibTab_Krankenbla:
            frmAdrSuch.Show vbModal
Case RibTab_Abrechnung:
            frmAdrSuch.Show vbModal
Case RibTab_Tagesproto:
            FTaSu
Case RibTab_Rezeptmodul:
            frmAdrSuch.Show vbModal
Case RibTab_Belegmodul:
            frmAdrSuch.Show vbModal
Case RibTab_Bildmodul:
            frmAdrSuch.Show vbModal
Case RibTab_Tex_Dokumt:
            frmAdrSuch.Show vbModal
Case RibTab_Tex_Vorlag:
            frmAdrSuch.Show vbModal
Case RibTab_Tex_Rezept:
            frmAdrSuch.Show vbModal
Case RibTab_Rechnungen:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSRe = False Then 'Suchleiste Rechnungen
                    SSuLe
                End If
            End If
            If GlSRe = True Then 'Suchleiste Rechnungen
                Set CmCom = CmBrs.FindControl(CmCom, SY_RE_Rechnung_SuchCombp, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Mahnwesen:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSOp = False Then
                    SSuLe
                End If
            End If
            If GlSOp = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_PO_Posten_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Buchungen:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSBu = False Then
                    SSuLe
                End If
            End If
            If GlSBu = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_BU_Buchung_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_HomeBanki:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSBa = False Then
                    SSuLe
                End If
            End If
            If GlSBa = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_BA_Banking_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Ter_Listen:
            If KeyPr = False Then
                SSuLe
            Else
                If GlStT = False Then
                    SSuLe
                End If
            End If
            If GlStT = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_TL_Terminliste_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Ter_Akont:
            If KeyPr = False Then
                SSuLe
            Else
                If GlStT = False Then
                    SSuLe
                End If
            End If
            If GlStT = True Then
                Set CmCom = CmBrs.FindControl(CmCom, SY_TL_Terminliste_SuchCombo, , True)
                If CmEdi.Visible = True Then
                    With CmEdi
                        .SetFocus
                        .Execute
                    End With
                Else
                    With CmCom
                        .SetFocus
                        .Execute
                    End With
                End If
            End If
Case RibTab_Ter_Warte:

Case RibTab_LabBericht:
            frmAdrSuch.Show vbModal
Case RibTab_LabBerichte:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSLb = False Then
                    SSuLe
                End If
            End If
            If GlSLb = True Then
                With CmEdi
                    .SetFocus
                    .Execute
                End With
            End If
Case RibTab_LabAuftrag:
            frmAdrSuch.Show vbModal
Case RibTab_LabAuftrage:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSLa = False Then
                    SSuLe
                End If
            End If
            If GlSLa = True Then
                With CmEdi
                    .SetFocus
                    .Execute
                End With
            End If
Case RibTab_Kat_Eintrg:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSei = False Then
                    SSuLe
                End If
            End If
            If GlSei = True Then
                With CmEdi
                    .SetFocus
                    .Execute
                End With
            End If
Case RibTab_Kat_Ketten:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSKe = False Then
                    SSuLe
                End If
            End If
            If GlSKe = True Then
                With CmEdi
                    .SetFocus
                    .Execute
                End With
            End If
Case RibTab_Kat_Frage:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSFr = False Then
                    SSuLe
                End If
            End If
            If GlSFr = True Then
                With CmEdi
                    .SetFocus
                    .Execute
                End With
            End If
Case RibTab_Tex_Email:
            If KeyPr = False Then
                SSuLe
            Else
                If GlSMl = False Then
                    SSuLe
                End If
            End If
            If GlSMl = True Then
                With CmEdi
                    .SetFocus
                    .Execute
                End With
            End If
End Select

Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFe " & Err.Number
Resume Next

End Sub
Private Sub FSuKr()
On Error GoTo OrErr
'In das Suchfeld springen

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Set CmEdi = CmBrs.FindControl(CmEdi, SY_KB_KraBla_Suchen, , True)

With CmEdi
    .SetFocus
    .Execute
End With

Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuKr " & Err.Number
Resume Next

End Sub
Private Sub FSuJa()
On Error GoTo OrErr
'Stellt ein neues Buchungsjahr ein

Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo4 = FM.repCont4
Set RpCo8 = FM.repCont8
Set RpCo5 = FM.repCont5

Select Case GlBut
Case RibTab_Rechnungen:
            GlSuR = GlSuX
            SSuch
            Set RpCo4 = FM.repCont4
            RpCo4.SetFocus
Case RibTab_Buchungen:
            GlSuB = GlSuX
            SSuch
            RpCo1.SetFocus
Case RibTab_HomeBanki:
            GlSuH = GlSuX
            SSuch
            RpCo1.SetFocus
Case RibTab_Ter_Listen:
            GlSuT = GlSuX
            SSuch
            RpCo1.SetFocus
Case RibTab_Ter_Akont:
            GlSuT = GlSuX
            SSuch
            RpCo1.SetFocus
Case RibTab_LabBerichte:
            GlSuL = GlSuX
            SSuch
            RpCo1.SetFocus
Case RibTab_LabAuftrage:
            GlSuU = GlSuX
            SSuch
            RpCo1.SetFocus
End Select

Set RpCo1 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo8 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuJa " & Err.Number
Resume Next

End Sub
Private Sub FSuLe(ByVal SuStr As String, ByVal TolId As Long)
On Error GoTo OrErr
'ABC Leiste

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo2 = FM.repCont2
Set RpCo8 = FM.repCont8
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

SSuAu 'Hebt die markierten Suchbuchstaben wieder auf
DoEvents

CmAcs(TolId).Checked = True

Select Case GlBut
Case RibTab_Adressen:
        With GlSuP
            .SuIdx = 7
            .SuStr = SuStr
        End With
Case RibTab_Mandanten:
        With GlSuP
            .SuIdx = 7
            .SuStr = SuStr
        End With
Case RibTab_Verordner:
        With GlSuP
            .SuIdx = 7
            .SuStr = SuStr
        End With
Case RibTab_Mitarbeit:
        With GlSuP
            .SuIdx = 7
            .SuStr = SuStr
        End With
Case RibTab_Kat_Eintrg:
        With GlSuE
            .SuIdx = 4
            .SuStr = SuStr
        End With
Case RibTab_Kat_Ketten:
        With GlSuN
            .SuIdx = 4
            .SuStr = SuStr
        End With
Case RibTab_Kat_Frage:
        With GlSuG
            .SuIdx = 4
            .SuStr = SuStr
        End With
End Select

DoEvents
SSuch
DoEvents

Select Case GlBut
Case RibTab_Adressen:
        Set RpCo2 = FM.repCont2
        If RpCo2.Records.Count = 0 Then
            SPopu "Patient nicht gefunden", "Der von Ihnen gesuchte Patient, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo2.SetFocus
        End If
Case RibTab_Mandanten:
        Set RpCo2 = FM.repCont2
        If RpCo2.Records.Count = 0 Then
            SPopu "Mandant nicht gefunden", "Der von Ihnen gesuchte Mandant, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo2.SetFocus
        End If
Case RibTab_Verordner:
        Set RpCo2 = FM.repCont2
        If RpCo2.Records.Count = 0 Then
            SPopu "Verordner nicht gefunden", "Der von Ihnen gesuchte Verordner, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo2.SetFocus
        End If
Case RibTab_Mitarbeit:
        Set RpCo2 = FM.repCont2
        If RpCo2.Records.Count = 0 Then
            SPopu "Mitarbeiter nicht gefunden", "Der von Ihnen gesuchte Mitarbeiter, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo2.SetFocus
        End If
Case RibTab_Kat_Eintrg:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
Case RibTab_Kat_Ketten:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Kette nicht gefunden", "Die von Ihnen gesuchte Kette, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
Case RibTab_Kat_Frage:
        Set RpCo8 = FM.repCont8
        If RpCo8.Records.Count = 0 Then
            SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
        Else
            RpCo8.SetFocus
        End If
End Select

Set CmBrs = Nothing
Set RpCo2 = Nothing
Set RpCo8 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuLe " & Err.Number
Resume Next

End Sub
Private Sub FStat()
On Error GoTo AnErr

Dim RetWe As Long
Dim GrIco As Long
Dim KlIco As Long
Dim NetNa As String
Dim BenNa As String
Dim GesNa As String
Dim PfNa3 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSli As XtremeCommandBars.StatusBarSliderPane
Dim CmSwi As XtremeCommandBars.StatusBarSwitchPane
Dim CmPrg As XtremeCommandBars.StatusBarProgressPane
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag

Set clNet = New clsNetz
Set clFil = New clsFile

NetNa = clNet.NetNam
BenNa = clNet.NetBen

If NetNa = vbNullString Then
    NetNa = "XenServer"
End If

If BenNa = vbNullString Then
    BenNa = "User"
End If

GesNa = BenNa & " - " & NetNa

If GlRDP = False Then
    PfNa3 = App.Path & "\Fernwartung14.exe"
End If

With CmSta
    .removeAll
    .EnableCustomization False
    .Font.Name = GlTFt.Name
    .Font.SIZE = 8
    .DrawDisabledText = True
    .EnableMarkup = False
    .IdleText = vbNullString
    .ShowSizeGripper = True
    Set CmPan = .AddPane(Tex_Pa_Plac1)
    With CmPan
        .Enabled = False
        .BeginGroup = False
        .Width = 40
    End With
    If GlRDP = False Then
        If clFil.FilVor(PfNa3) = True Then
            RetWe = ExtractIconEx(PfNa3, 0, GrIco, KlIco, 1)
            ImMan.Icons.AddIcon KlIco, Prg_Icn3, xtpImageNormal
            Set CmSwi = .AddSwitchPane(Tex_Pa_Progr)
        
            CmSwi.AddSwitch Prg_Icn3, vbNullString
            Set CmPan = .AddPane(Tex_Pa_Labl1)
            With CmPan
                .Width = 220
                .BeginGroup = False
                .Alignment = xtpAlignmentLeft
                .Style = SBPS_NOBORDERS
                .Text = " Fernsteuerung"
            End With
        Else
            Set CmPan = .AddPane(Tex_Pa_Labl1)
            With CmPan
                .Width = 220
                .Text = vbNullString
                .BeginGroup = False
                .Alignment = xtpAlignmentLeft
                .Style = SBPS_NOBORDERS
            End With
        End If
    Else
        Set CmPan = .AddPane(Tex_Pa_Progr)
        With CmPan
            .Width = 120
            .BeginGroup = False
            .Alignment = xtpAlignmentLeft
            .Style = SBPS_NOBORDERS
            .Text = vbNullString
        End With
        Set CmPan = .AddPane(Tex_Pa_Labl1)
        With CmPan
            .Width = 320
            .BeginGroup = False
            .Alignment = xtpAlignmentLeft
            .Style = SBPS_NOBORDERS
            .Text = GesNa
        End With
    End If
    Set CmPan = .AddPane(Tex_Pa_Plac4)
    With CmPan
        .Enabled = False
        .BeginGroup = False
        .Width = 10
    End With
    Set CmPan = .AddPane(Tex_Pa_Labl2)
    With CmPan
        .Width = 300
        .BeginGroup = True
        .Alignment = xtpAlignmentLeft
        .Style = SBPS_NOBORDERS
        .Text = vbNullString
    End With
    Set CmPan = .AddPane(Tex_Pa_Labl3)
    With CmPan
        .Style = SBPS_STRETCH + SBPS_NOBORDERS
        .Alignment = xtpAlignmentLeft
        .Text = vbNullString
    End With
    Set CmPan = .AddPane(Tex_Pa_Labl4)
    With CmPan
        .Width = 400
        .Alignment = xtpAlignmentLeft
        .Text = vbNullString
    End With
    Set CmPan = .AddPane(Tex_Pa_Plac5)
    With CmPan
        .Style = SBPS_STRETCH
        .Text = vbNullString
    End With
    Set CmPan = .AddPane(Tex_Pa_Seite)
    With CmPan
        .Alignment = xtpAlignmentLeft
        .BeginGroup = True
        .Text = " Seiten:"
        .Width = 100
    End With
    Set CmSwi = .AddSwitchPane(Tex_Pa_Layou)
    CmSwi.AddSwitch IC16_AnsNor, "Normalansicht"
    CmSwi.AddSwitch IC16_AnsBre, "Seitenansicht"
    CmSwi.AddSwitch IC16_AnsFli, "Fließansicht"
    Select Case GlViT 'ViewMode Textverarbeitung
    Case 0: CmSwi.Checked = IC16_AnsNor
    Case 2: CmSwi.Checked = IC16_AnsBre
    Case 3: CmSwi.Checked = IC16_AnsFli
    End Select
    Set CmPan = .AddPane(Tex_Pa_Linia)
    With CmPan
        .Alignment = xtpAlignmentRight
        .BeginGroup = True
        .Text = "Liniale:"
    End With
    Set CmSwi = .AddSwitchPane(Tex_Linial)
    CmSwi.AddSwitch IC16_Ruler, "Liniale"
    If GlLiT = True Then 'Lineal Textverarbeitung
        CmSwi.Checked = IC16_Ruler
    End If
    Set CmPan = .AddPane(Tex_Pa_ZoPan)
    With CmPan
        .SetPadding 15, 0, 15, 0
        .Text = "Zoom: " & GlZoT & "%"
        .ToolTip = "Zeigt die Zoomgröße des Dokuments"
    End With
    Set CmPan = .AddPane(Tex_Pa_Plac2)
    With CmPan
        .Enabled = False
        .Style = SBPS_NOBORDERS
        .BeginGroup = False
        .Width = 15
    End With
    Set CmSli = .AddSliderPane(Tex_Pa_ZoSli)
    With CmSli
        .BeginGroup = False
        .Min = 10
        .Max = 200
        .SetTicks 100
        .SetTooltipPart XTP_SB_LINELEFT, "Zoom verringern"
        .SetTooltipPart XTP_SB_LINERIGHT, "Zoom vergrößern"
        .ToolTip = "Zoomgröße"
        .Value = GlZoT
        .Width = 120
    End With
    Set CmPan = .AddPane(Tex_Pa_Plac3)
    With CmPan
        .Enabled = False
        .Style = SBPS_NOBORDERS
        .BeginGroup = False
        .Width = 50
    End With
    .Visible = True
End With

CmBrs.PaintManager.RefreshMetrics
DoEvents
CmBrs.RecalcLayout
DoEvents

Set clNet = Nothing
Set clFil = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStat " & Err.Number
Resume Next

End Sub

Private Sub FTabu()
On Error GoTo AnErr

Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Tit1 = "Dokument Speichern"
Mld1 = "Soll das aktuelle Dokument gespeichert werden?"

Select Case GlBut
Case RibTab_Tex_Dokumt:
    If GlTSV = True Then 'Speichern Textverarbeitung
        If RbTab.id <> RibTab_Tex_Vorlag Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                STxSa
            Else
                GlTSV = False
            End If
        End If
    End If
Case RibTab_Tex_Vorlag:
    If GlTSV = True Then 'Speichern Textverarbeitung
        If RbTab.id <> RibTab_Tex_Dokumt Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                STxSa
            Else
                GlTSV = False
            End If
        End If
    End If
Case RibTab_Tex_Rezept:
    If GlTSV = True Then 'Speichern Textverarbeitung
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            STxSa
        Else
            GlTSV = False
        End If
    End If
Case RibTab_Krankenbla:
    If GlMen = True Then 'Krankenblattdokument
        If GlTSV = True Then 'Speichern Textverarbeitung
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage = 6 Then
                SKrTx True
            Else
                GlTSV = False
            End If
        End If
    End If
End Select

Set FM = frmMain
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

GlAkt = True

GlBut = RbTab.id

Select Case GlSHt
Case ShoCut_Start:
        GlBu0 = RbTab.id
        'IniSetVal  "System", "LasTa0", GlBu0
Case ShoCut_Adresse:
        GlBu1 = RbTab.id
        'IniSetVal  "System", "LasTa1", GlBu1
Case ShoCut_Kranken:
        GlBu2 = RbTab.id
        IniSetVal "System", "LasTa2", GlBu2
Case ShoCut_Finanz:
        GlBu3 = RbTab.id
        IniSetVal "System", "LasTa3", GlBu3
Case ShoCut_Termin:
        GlBu4 = RbTab.id
        IniSetVal "System", "LasTa4", GlBu4
Case ShoCut_Labor:
        GlBu5 = RbTab.id
        IniSetVal "System", "LasTa5", GlBu5
Case ShoCut_Texte:
        GlBu6 = RbTab.id
        IniSetVal "System", "LasTa6", GlBu6
Case ShoCut_Katalog:
        GlBu7 = RbTab.id
        'IniSetVal  "System", "LasTa7", GlBu7
Case ShoCut_Abrechn:
        GlBu8 = RbTab.id
        'IniSetVal  "System", "LasTa8", GlBu8
End Select

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
DoEvents
If GlBiA = False Then 'Bildschirmaktualisierung
    clFen.FenDsk 1
Else
    clFen.FenDsk 2
End If

DoEvents
SButD 'Docpains ein- und ausblenden
DoEvents
SSuAu 'Hebt die markierten Suchbuchstaben wieder auf
DoEvents
SButt
DoEvents
SPosi
DoEvents
SBuLa
DoEvents

clFen.FenDsk 3
DoEvents
Screen.MousePointer = vbNormal

GlAlB = GlBut

GlAkt = False

Set clFen = Nothing

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTeAb()
On Error GoTo AdErr
'Ändert den Terminstatus Abhaken

Dim TerNr As Long
Dim RowNr As Long
Dim PrGui As String
Dim MitNa As String
Dim PrStr As String
Dim TeGui As String
Dim TeAbg As Boolean
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpSel As XtremeReportControl.ReportSelectedRows
Dim RpRow As XtremeReportControl.ReportRow
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents

Set FM = frmMain
Set CaCol = FM.calCont1
Set RpCo1 = FM.repCont1
Set RpCo6 = FM.repCont6

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If GlBut = RibTab_Ter_Listen Or GlBut = RibTab_Ter_Akont Or GlBut = RibTab_Ter_Warte Then
    If GlWaZ = True Then 'Wartezimmer Kontextmenü
        Set RpCls = RpCo6.Columns
        Set RpSel = RpCo6.SelectedRows
    Else
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
    End If
    If RpSel.Count > 0 Then
        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then
                RowNr = RpRow.Index
                Set RpCol = RpCls.Find(Ter_ID2)
                TerNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Ter_GuiID)
                TeGui = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Ter_Abgehakt)
                TeAbg = Not CBool(RpRow.Record(RpCol.ItemIndex).Checked)
                DBCmEx2 "qryTerAbh1", "@IdAbh", "@IdxNr", TeAbg, TerNr '[Abhaken] ändern
                DoEvents
                If TeAbg = True Then
                    PrStr = "Termin abgehakt"
                Else
                    PrStr = "Termin nicht abgehakt"
                End If
                PrGui = CreateID("P")
                MitNa = GlMiA(GlSmI, 1)
                DBCmEx7 "qryTerPrAd", "@IdxNr", "@GuiId", "@TerId", "@IdDat", "@IdZei", "@IdStr", "@IdKom", TerNr, PrGui, TeGui, Now, Now, PrStr, MitNa
            End If
        Next RpRow
        DoEvents
        SUpTe RowNr
    End If
    GlWaZ = False 'WICHTIG!
Else
    Set ViEvs = CaCol.ActiveView.GetSelectedEvents
    If ViEvs.Count > 0 Then
        For Each ViEvt In ViEvs
            If ViEvt.Selected = True Then
                Set CaEvt = ViEvt.Event
                TerNr = CaEvt.id
                If CaEvt.MeetingFlag = True Then
                    CaEvt.MeetingFlag = False 'wird später bei RedrawControl wieder hinzugefügt
                End If
                If CaEvt.PrivateFlag = True Then
                    CaEvt.PrivateFlag = False 'wird später bei RedrawControl wieder hinzugefügt
                End If
                If CaEvt.CustomIcons.Find(IC16_Sign_Check) >= 0 Then
                    CaEvt.CustomIcons.RemoveID IC16_Sign_Check
                    DBCmEx2 "qryTerAbh1", "@IdAbh", "@IdxNr", 0, TerNr '[Abhaken] ändern
                    PrStr = "Termin nicht abgehakt"
                Else
                    CaEvt.CustomIcons.RemoveID IC16_Sign_Check
                    CaEvt.CustomIcons.Add IC16_Sign_Check
                    DBCmEx2 "qryTerAbh1", "@IdAbh", "@IdxNr", -1, TerNr '[Abhaken] ändern
                    PrStr = "Termin abgehakt"
                End If
                DoEvents
                
                S_TeDe TerNr
                With GlTDt
                    TeGui = .TeGui
                End With
                DoEvents

                PrGui = CreateID("P")
                MitNa = GlMiA(GlSmI, 1)
                DBCmEx7 "qryTerPrAd", "@IdxNr", "@GuiId", "@TerId", "@IdDat", "@IdZei", "@IdStr", "@IdKom", TerNr, PrGui, TeGui, Now, Now, PrStr, MitNa
                DoEvents
                
                Exit For
            End If
        Next ViEvt
    End If
    With CaCol
        .DataProvider.ChangeEvent CaEvt
        .Populate
        .RedrawControl
    End With
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CaCol = Nothing
Set RpCo1 = Nothing
Set RpCo6 = Nothing

Set clFen = Nothing

Exit Sub

AdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeAb " & Err.Number
Resume Next

End Sub
Private Sub FTeAn(ByVal AnzOp As Integer)
On Error GoTo SaErr
'Stellt bestimmt Anzeigeoptionen im Kalender dar

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCoT As XtremeReportControl.ReportControl

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CaCol = FM.calCont1
Set RpCoT = FM.repContT
Set CmAcs = CmBrs.Actions

Screen.MousePointer = vbHourglass

Select Case AnzOp
Case 1:
    GlMZe = Not GlMZe
    CmAcs(SY_TE_Termin_GlMZe).Checked = GlMZe
    IniSetVal "TerSys", "ManZei", GlMZe
    CaCol.RedrawControl
Case 2:
    GlMFa = Not GlMFa
    CmAcs(SY_TE_Termin_GlMFa).Checked = GlMFa
    IniSetVal "TerSys", "ManFar", GlMFa
    CaCol.RedrawControl
Case 3:
    GlTGs = Not GlTGs
    CmAcs(SY_TE_Termin_GlTGs).Checked = GlTGs
    IniSetVal "TerSys", "TerGes", GlTGs
    DoEvents
    S_TeLi
    DoEvents
    SUpTe
Case 4:
    GlTeD = Not GlTeD 'Termindetails zeigen
    CmAcs(SY_TE_Termin_GlTeD).Checked = GlTeD
    IniSetVal "TerSys", "TerDet", GlTeD
Case 5:
    GlTVe = Not GlTVe 'Terminverschiebung zulassen
    CmAcs(SY_TE_Termin_GlTVe).Checked = GlTVe
    IniSetVal "TerSys", "TerVer", GlTVe
Case 6:
    GlSSt = Not GlSSt 'Starre Termintaktung
    CmAcs(SY_TE_Termin_GlTSt).Checked = GlSSt
    IniSetVal "TerSys", "StaRas", GlSSt
    S_SeSe 29, , , , GlSSt
Case 7:
    GlTDe = Not GlTDe 'Erweiterter Terminbetreff
    CmAcs(SY_TE_Termin_GlTDe).Checked = GlTDe
    IniSetVal "TerSys", "TerErw", GlTDe
    DoEvents
    S_TeLi
Case 8:
    GlTTe = Not GlTTe 'Telefonnummer Terminbetreff
    CmAcs(SY_TE_Termin_GlTTe).Checked = GlTTe
    IniSetVal "TerSys", "ZeiBet", GlTTe
    DoEvents
    S_TeLi
Case 9:
    GlMiW = Not GlMiW
    CmAcs(SY_TE_Termin_GlMiW).Checked = GlMiW
    IniSetVal "TerSys", "MitAus", GlMiW
    DoEvents
    RpCoT.Visible = GlMiW
    DoEvents
    SPosi
Case 10:
    GlTZe = Not GlTZe 'Terminzeitanpassung
    CmAcs(SY_TE_Termin_GlTZe).Checked = GlTZe
    IniSetVal "TerSys", "TeZeAn", GlTZe
Case 11:
    GlTKo = Not GlTKo 'Kalenderkopf
    CmAcs(SY_TE_Termin_GlTKo).Checked = GlTKo
    CmAcs(SY_TE_Termin_GlTKl).Enabled = GlTKo
    CaCol.ShowCaptionBar = GlTKo
    CaCol.Populate
    IniSetVal "TerSys", "TimKop", GlTKo
Case 12:
    GlTKl = Not GlTKl
    CmAcs(SY_TE_Termin_GlTKl).Checked = GlTKl
    CaCol.OneLineCaptionBar = GlTKl
    CaCol.Populate
    IniSetVal "TerSys", "TimKoK", GlTKl
Case 13:
    STeAn
Case 14:
    GlDeT = Not GlDeT 'Telefonnummer Terminbetreff
    CmAcs(SY_TE_Termin_GlDeT).Checked = GlDeT
    IniSetVal "TerSys", "ZeStTe", GlTTe
    S_SeSe 93, , , , GlDeT
    DoEvents
    STeDe
Case 15:
    GlTrD = Not GlTrD 'Termindetails mit Mitarbeitername
    CmAcs(SY_TE_Termin_GlTrD).Checked = GlTrD
    IniSetVal "TerSys", "TeDeMi", GlTrD
    S_SeSe 110, , , , GlTrD
    DoEvents
    STeDe
End Select

Screen.MousePointer = vbNormal

Set CmBrs = Nothing
Set CaCol = Nothing
Set RpCoT = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeAn " & Err.Number
Resume Next

End Sub
Private Sub FTeDo()
On Error GoTo AdErr
'Ändert den Terminstatus OnlTer

Dim TerNr As Long
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents

Set FM = frmMain
Set CaCol = FM.calCont1

Set ViEvs = CaCol.ActiveView.GetSelectedEvents
If ViEvs.Count > 0 Then
    For Each ViEvt In ViEvs
        If ViEvt.Selected = True Then
            Set CaEvt = ViEvt.Event
            TerNr = CaEvt.id
            If CaEvt.MeetingFlag = True Then
                CaEvt.MeetingFlag = False 'wird später bei RedrawControl wieder hinzugefügt
            End If
            If CaEvt.PrivateFlag = True Then
                CaEvt.PrivateFlag = False 'wird später bei RedrawControl wieder hinzugefügt
            End If
            If CaEvt.CustomIcons.Find(IC16_Phone_Mobil) >= 0 Then
                CaEvt.CustomIcons.RemoveID IC16_Phone_Mobil
                DBCmEx2 "qryTerOnTe", "@OnlTe", "@IdxNr", 0, TerNr '[OnlTer]
            Else
                CaEvt.CustomIcons.RemoveID IC16_Phone_Mobil
                CaEvt.CustomIcons.Add IC16_Phone_Mobil
                DBCmEx2 "qryTerOnTe", "@OnlTe", "@IdxNr", -1, TerNr '[OnlTer]
            End If
            Exit For
        End If
    Next ViEvt
End If

With CaCol
    .DataProvider.ChangeEvent CaEvt
    .Populate
    .RedrawControl
End With

Set CaCol = Nothing

Exit Sub

AdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeDo " & Err.Number
Resume Next

End Sub

Private Sub FTaEd(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
On Error Resume Next

Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo6 = FM.repCont6
Set RpCoK = FM.repContK
Select Case GlBut
Case RibTab_Krankenbla:
        Set RpSel = RpCoK.SelectedRows
        Set RpCls = RpCoK.Columns
Case RibTab_Abrechnung:
        Set RpSel = RpCo6.SelectedRows
        Set RpCls = RpCo6.Columns
Case RibTab_Tagesproto:
        Set RpSel = RpCo6.SelectedRows
        Set RpCls = RpCo6.Columns
Case Else:
        Exit Sub
End Select

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If Shift = 0 Then
            Select Case KeyCode
            Case vbKeyF2:
                    Select Case GlBut
                    Case RibTab_Krankenbla: RpCoK.Navigator.BeginEdit
                    Case RibTab_Abrechnung: RpCo6.Navigator.BeginEdit
                    Case RibTab_Tagesproto: RpCo6.Navigator.BeginEdit
                    End Select
            Case vbKeyTab:
            Case vbKeyReturn: SMark
            Case vbKeyDown: SMark
            Case vbKeyUp: SMark
            Case vbKeyPageDown: SMark
            Case vbKeyPageUp: SMark
            End Select
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo6 = Nothing
Set RpCoK = Nothing

End Sub
Private Sub FTast(ByVal Flag As Integer)
On Error GoTo AnErr
'Funktionstasten

Dim RetWe As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set CmBrs = FM.comBar01
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set RpCo8 = FM.repCont8
Set TxCoN = FM.TexCont1
Set CmAcs = CmBrs.Actions

Select Case Flag
Case KY_F2:
    Select Case GlBut
    Case RibTab_Adressen:
                SAdre 2
    Case RibTab_Mandanten:
                SAdre 2, True
    Case RibTab_Verordner:
                SAdre 2, True
    Case RibTab_Mitarbeit:
                SAdre 2, True
    Case RibTab_Fragebogen:
                SAnEd
    Case RibTab_Krankenbla:
                SKoEd
    Case RibTab_Abrechnung:
                frmReEdit.Show vbModal
    Case RibTab_Belegmodul:
                frmRzEdit.Show vbModal
    Case RibTab_Rezeptmodul:
                frmRzEdit.Show vbModal
    Case RibTab_Bildmodul:
                SBiUm
    Case RibTab_Rechnungen:
                frmReEdit.Show vbModal
    Case RibTab_Mahnwesen:
                frmOPEdit.Show vbModal
    Case RibTab_Buchungen:
                frmBuEdit.Show
    Case RibTab_HomeBanki:
                frmBaEdit.Show vbModal
    Case RibTab_Statistik:
                
    Case RibTab_Ter_Kalend:
                STerm
    Case RibTab_Ter_Listen:
                STerm
    Case RibTab_Ter_Akont:
                STerm
    Case RibTab_Ter_Warte:
                STerm
    Case RibTab_Ter_Raeume:
                STerm
    Case RibTab_Ter_Mitarb:
                STerm
    Case RibTab_LabBericht:
                frmLaBear.Show vbModal
    Case RibTab_LabAuftrag:
                frmLaBear.Show vbModal
    Case RibTab_LabBerichte:
                frmLaBear.Show vbModal
    Case RibTab_LabAuftrage:
                frmLaBear.Show vbModal
    Case RibTab_Kat_Eintrg:
                KaEdi
    Case RibTab_Kat_Ketten:
                KaEdi
    Case RibTab_Kat_Frage:
                KaEdi
    Case RibTab_Tex_Dokumt:
                SAdre 2
    Case RibTab_Tex_Vorlag:
                SAdre 2
    Case RibTab_Tex_Rezept:
                SAdre 2
    Case RibTab_Tex_NewsLe:

    Case RibTab_Tex_Email:
                SMaVie
    Case RibTab_Kat_Explor:
                KDatei 7
    End Select
Case KY_F3:
    Select Case GlBut
    Case RibTab_Startseite:
        SAdre 1
    Case RibTab_Adressen:
        SAdre 1
    Case RibTab_Mandanten:
        SAdre 1, True
    Case RibTab_Verordner:
        SAdre 1, True
    Case RibTab_Mitarbeit:
        SAdre 1, True
    Case RibTab_Fragebogen:
        SAnEd True
    Case RibTab_Krankenbla:
        KrMain -2
    Case RibTab_Abrechnung:
        Set FM = frmMain
        FM.cmbZiffe.Text = vbNullString
        FM.cmbBezei.Text = vbNullString
        FM.cmbTypen.SetFocus
        RetWe = SendMessage(FM.cmbTypen.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
        RetWe = SendMessage(FM.cmbTypen.hwnd, CB_SETCURSEL, 1, ByVal 0&)
    Case RibTab_Tagesproto:
    Case RibTab_Vorbereit:
    Case RibTab_Rezeptmodul:
        SReDi 3
    Case RibTab_Belegmodul:
        SReDi 3
    Case RibTab_Bildmodul:
        SBiEi
    Case RibTab_Rechnungen:
        SReDi 1
    Case RibTab_Mahnwesen:
        GlNeP = True
        frmOPEdit.Show vbModal
    Case RibTab_Buchungen:
        GlNeB = True 'neue Buchung
        frmBuEdit.Show
    Case RibTab_HomeBanki:
        SBuAb
    Case RibTab_Statistik:
    
    Case RibTab_Ter_Kalend:
        STerm True
    Case RibTab_Ter_Listen:
        STerm True
    Case RibTab_Ter_Akont:
        STerm True
    Case RibTab_Ter_Warte:
        STerm True
    Case RibTab_Ter_Raeume:
        STerm True
    Case RibTab_Ter_Mitarb:
        STerm True
    Case RibTab_LabBericht:
        STran 2
    Case RibTab_LabAuftrag:
        frmNeuAuf.Show vbModal
    Case RibTab_LabBerichte:
        STran 2
    Case RibTab_LabAuftrage:
        frmNeuAuf.Show vbModal
    Case RibTab_Kat_Eintrg:
        KaNeu
    Case RibTab_Kat_Ketten:
        KaNeu
    Case RibTab_Kat_Frage:
        FaNeu 1
    Case RibTab_Tex_Dokumt:
        FTxNe
    Case RibTab_Tex_Vorlag:
        FObje Tex_DatLoa
    Case RibTab_Tex_Rezept:
        FTxNe
    Case RibTab_Tex_NewsLe:
        FTxNe
    Case RibTab_Tex_Email:
        SMaNe
    Case RibTab_Kat_Explor:
        KDatei 11
    End Select
Case KY_F4:
    Select Case GlBut
    Case RibTab_Adressen:
                AdFMa
    Case RibTab_Mandanten:
                SAdre 2, True
    Case RibTab_Verordner:
                SAdre 2, True
    Case RibTab_Mitarbeit:
                SAdre 2, True
    Case RibTab_Fragebogen:
    Case RibTab_Krankenbla:
                SKrDa
    Case RibTab_Abrechnung:
                If CmAcs(SY_AB_Abrech_Zahlung).Enabled = True Then frmAnzahl.Show vbModal
    Case RibTab_Tagesproto:
    Case RibTab_Vorbereit:
    Case RibTab_Belegmodul:
    Case RibTab_Rezeptmodul:
    Case RibTab_Bildmodul:
                SBiSc
    Case RibTab_Rechnungen:
                frmReFilt.Show vbModal
    Case RibTab_Mahnwesen:
                frmOPAusg.Show vbModal
    Case RibTab_Buchungen:
                frmBuKont.Show vbModal
    Case RibTab_HomeBanki:
                S_BaBuc
    Case RibTab_Statistik:
    
    Case RibTab_Ter_Kalend:
                STerm False, True
    Case RibTab_Ter_Listen:
                STerm
    Case RibTab_Ter_Akont:
                STerm
    Case RibTab_Ter_Warte:
                STerm
    Case RibTab_Ter_Raeume:
                STerm False, True
    Case RibTab_Ter_Mitarb:
                STerm False, True
    Case RibTab_LabBericht:
    Case RibTab_LabAuftrag:
    Case RibTab_LabBerichte:
    Case RibTab_LabAuftrage:
    Case RibTab_Kat_Eintrg:
    Case RibTab_Kat_Ketten:
    Case RibTab_Kat_Frage:
    Case RibTab_Tex_Email:
                MaSav 5
    Case RibTab_Kat_Explor:
    End Select
Case KY_F7:
    Select Case GlBut
    Case RibTab_Adressen:
                SAdre 2
    Case RibTab_Mandanten:
                SAdre 2, True
    Case RibTab_Verordner:
                SAdre 2, True
    Case RibTab_Mitarbeit:
                SAdre 2, True
    Case RibTab_Fragebogen:
                SAdre 3
    Case RibTab_Krankenbla:
                SAdre 3
    Case RibTab_Abrechnung:
                SAdre 3
    Case RibTab_Tagesproto:
    Case RibTab_Vorbereit:
    Case RibTab_Belegmodul:
                SAdre 3
    Case RibTab_Rezeptmodul:
                SAdre 3
    Case RibTab_Bildmodul:
                SAdre 3
    Case RibTab_Rechnungen:
                SReZe
    Case RibTab_Mahnwesen:
    Case RibTab_Buchungen:
    Case RibTab_HomeBanki:
    Case RibTab_Statistik:
    Case RibTab_Ter_Kalend:
    Case RibTab_Ter_Listen:
    Case RibTab_Ter_Akont:
    Case RibTab_Ter_Warte:
    Case RibTab_Ter_Raeume:
    Case RibTab_Ter_Mitarb:
    Case RibTab_LabBericht:
    Case RibTab_LabAuftrag:
    Case RibTab_LabBerichte:
    Case RibTab_LabAuftrage:
    Case RibTab_Kat_Eintrg:
    Case RibTab_Kat_Ketten:
    Case RibTab_Kat_Frage:
    Case RibTab_Tex_Dokumt:
                SAdre 3
    Case RibTab_Tex_Vorlag:
                SAdre 3
    Case RibTab_Tex_Rezept:
                SAdre 3
    Case RibTab_Tex_NewsLe:

    Case RibTab_Tex_Email:
    Case RibTab_Kat_Explor:
    End Select
Case KY_DEL:
    Select Case GlBut
    Case RibTab_Adressen:
            SLoHa
    Case RibTab_Mandanten:
            SLoHa
    Case RibTab_Verordner:
            SLoHa
    Case RibTab_Mitarbeit:
            SLoHa
    Case RibTab_Fragebogen:
            SLoHa
    Case RibTab_Krankenbla:
            S_KrLo
    Case RibTab_Abrechnung:
            S_KrLo
    Case RibTab_Vorbereit:
    Case RibTab_Tagesproto:
    Case RibTab_Belegmodul:
            SLoHa
    Case RibTab_Rezeptmodul:
            SLoHa
    Case RibTab_Bildmodul:
    Case RibTab_Rechnungen:
            SLoHa
    Case RibTab_Mahnwesen:
            SLoHa
    Case RibTab_Buchungen:
            SLoHa
    Case RibTab_HomeBanki:
            SLoHa
    Case RibTab_Statistik:
    Case RibTab_Ter_Kalend:
    Case RibTab_Ter_Listen:
            SLoHa
    Case RibTab_Ter_Akont:
            SLoHa
    Case RibTab_Ter_Warte:
            SLoHa
    Case RibTab_Ter_Raeume:
    Case RibTab_Ter_Mitarb:
    Case RibTab_LabBericht:
            SLoHa
    Case RibTab_LabAuftrag:
            SLoHa
    Case RibTab_LabBerichte:
            SLoHa
    Case RibTab_LabAuftrage:
            SLoHa
    Case RibTab_Kat_Eintrg:
            SLoHa
    Case RibTab_Kat_Ketten:
            SLoHa
    Case RibTab_Tex_Dokumt:
            TxCoN.SelText = vbNullString
    Case RibTab_Tex_Vorlag:
            TxCoN.SelText = vbNullString
    Case RibTab_Tex_Rezept:
            TxCoN.SelText = vbNullString
    Case RibTab_Tex_NewsLe:
            TxCoN.SelText = vbNullString
    Case RibTab_Kat_Frage:
            SLoHa
    Case RibTab_Tex_Email:
            SLoHa
    Case RibTab_Kat_Explor:
    End Select
End Select

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTast " & Err.Number
Resume Next

End Sub
Private Sub FTaSu()
On Error GoTo SaErr

Dim SuStr As String
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

SuStr = InputBox("Bitte geben Sie den gewünschten Suchbegriff ein:", "Suchbegriff", vbNullString)

If SuStr <> vbNullString Then
    S_TaPo SuStr
    CmAcs(SY_TP_TagPro_Vollst).Enabled = True
End If

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTaSu " & Err.Number
Resume Next

End Sub
Private Sub FTeFi()
On Error GoTo SaErr
'Stellt einen anderen Resourcen Filter für die Kalender ein

Dim AktZa As Integer

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmFiT As XtremeCommandBars.CommandBarComboBox
Dim CmFil As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

If GlAkt = False Then
    Set clFen = New clsFenster
    clFen.hwnd = FM.hwnd
            
    Screen.MousePointer = vbHourglass
    clFen.FenDsk 2
    
    Set CmFiT = CmBrs.FindControl(CmFiT, SY_TE_Termin_FiltTyp, , True)
    Set CmFil = CmBrs.FindControl(CmFil, SY_TE_Termin_FiltIdx, , True)
    
    GlCaF = CmFiT.ListIndex 'Kalenderfilterauswahl
    
    Select Case GlCaF
    Case 1:
        CmFil.Clear
        CmAcs(SY_TE_Termin_FiltIdx).Enabled = False
    Case 2:
        If GlRaV = True Then 'Räume vorhanden
            CmFil.Clear
            For AktZa = 1 To UBound(GlRmu)
                CmFil.AddItem GlRmu(AktZa, 1)
                CmFil.ItemData(AktZa) = GlRmu(AktZa, 2)
            Next AktZa
            CmFil.ListIndex = GlCaS 'Kalenderfilterinhalt
        End If
        CmAcs(SY_TE_Termin_FiltIdx).Enabled = True
    Case 3:
        CmFil.Clear
        If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
            For AktZa = 1 To UBound(GlMiT)
                CmFil.AddItem GlMiT(AktZa, 1)
                CmFil.ItemData(AktZa) = GlMiT(AktZa, 2)
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMaT)
                CmFil.AddItem GlMaT(AktZa, 1)
                CmFil.ItemData(AktZa) = GlMaT(AktZa, 2)
            Next AktZa
        End If
        CmFil.ListIndex = GlCaS 'Kalenderfilterinhalt
        CmAcs(SY_TE_Termin_FiltIdx).Enabled = True
    End Select
    
    IniSetVal "TerSys", "KaGrId", GlCaF 'Kalenderfilterauswahl
    DoEvents

    With CaCol
        Select Case GlCal 'Kalenderanzeige
        Case 1:
            .UseMultiColumnWeekMode = True
            .Options.WorkWeekMask = xtpCalendarDayAllWeek
            .ViewType = xtpCalendarDayView
        Case 2:
            .UseMultiColumnWeekMode = True
            .Options.WorkWeekMask = xtpCalendarDayMo_Fr
            .ViewType = xtpCalendarWorkWeekView
        Case 3:
            .UseMultiColumnWeekMode = True
            .Options.WorkWeekMask = xtpCalendarDayAllWeek
            .ViewType = xtpCalendarFullWeekView
        Case 4:
            .UseMultiColumnWeekMode = False
            .ViewType = xtpCalendarWeekView
        Case 5:
            .UseMultiColumnWeekMode = False
            .ViewType = xtpCalendarMonthView
        End Select
    End With
    DoEvents

    STeAk
    
    clFen.FenDsk 3
    Screen.MousePointer = vbNormal
End If

Set CmBrs = Nothing
Set CmAcs = Nothing

Set clFen = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeFi " & Err.Number
Resume Next

End Sub
Private Sub FTeSe()
On Error GoTo AdErr
'Termin selketieren

Dim TerNr As Long
Dim DayGa As Boolean
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents

Set FM = frmMain
Set CaCol = FM.calCont1

GlTrM = False 'Terminbearbeitung direkt im Kalender
GlTrB = False 'Terminbearbeitung direkt im Kalender

Set ViEvs = CaCol.ActiveView.GetSelectedEvents
For Each ViEvt In ViEvs
    If ViEvt.Selected = True Then
        ViEvt.Selected = False
    End If
Next ViEvt

Set CaHit = CaCol.ActiveView.HitTest
If Not CaHit.HitCode = xtpCalendarHitTestDayViewTimeScale Then
    If Not CaHit.HitCode = xtpCalendarHitTestUnknown Then
        If Not CaHit.ViewEvent Is Nothing Then
            Set CaEvt = CaHit.ViewEvent.Event
            TerNr = CaEvt.id
        Else
            TerNr = 0
        End If
    Else
        TerNr = 0
    End If
Else
    TerNr = 0
End If

If TerNr > 0 Then
    If CaEvt.CustomIcons.Find(IC16_Phone_Mobil) >= 0 Then
        GlTNa = True 'Terminnachricht gesendet
    Else
        GlTNa = False
    End If
Else
    GlTNa = False
End If

If TerNr > 0 Then
    If CaEvt.CustomIcons.Find(IC16_Sign_Check) >= 0 Then
        GlTeA = True
    Else
        GlTeA = False
    End If
Else
    GlTeA = False
End If

If TerNr > 0 Then
    GlTem = TerNr
    GlAkS = CaEvt.BusyStatus + 1
    CaHit.ViewEvent.Selected = True
    CaCol.Populate
End If

Set CaCol = Nothing

Exit Sub

AdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeSe " & Err.Number
Resume Next

End Sub
Private Sub FTeSt(ByVal IdSta As Integer)
On Error GoTo AdErr
'Ändert den Terminstatus

Dim TerNr As Long
Dim PrGui As String
Dim MitNa As String
Dim PrStr As String
Dim TeGui As String
Dim NotSt As Integer
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents

Set FM = frmMain
Set CaCol = FM.calCont1

Tit1 = "E-Mail-Erinnerung deaktivieren"
Mld1 = "Soll die E-Mail-Erinnerung zu diesem Termin deaktviert werden?"

Set ViEvs = CaCol.ActiveView.GetSelectedEvents
If ViEvs.Count > 0 Then
    For Each ViEvt In ViEvs
        If ViEvt.Selected = True Then
            Set CaEvt = ViEvt.Event
            If CaEvt.CustomIcons.Count > 0 Then
                CaEvt.CustomIcons.RemoveID IC16_Pin_Green
                CaEvt.CustomIcons.RemoveID IC16_Pin_Gray
                CaEvt.CustomIcons.RemoveID IC16_Pin_Red
                If CaEvt.MeetingFlag = True Then
                    CaEvt.MeetingFlag = False 'wird später bei RedrawControl wieder hinzugefügt
                End If
                If CaEvt.PrivateFlag = True Then
                    CaEvt.PrivateFlag = False 'wird später bei RedrawControl wieder hinzugefügt
                End If
            End If
            TerNr = CaEvt.id
            CaEvt.BusyStatus = IdSta - 1
            Select Case IdSta
            Case 1: CaEvt.CustomIcons.Add IC16_Pin_Green
            Case 2: CaEvt.CustomIcons.Add IC16_Pin_Gray
            Case 3:
            Case 4: CaEvt.CustomIcons.Add IC16_Pin_Red
            End Select

            S_TeDe TerNr
            With GlTDt
                TeGui = .TeGui
                NotSt = .NotSt
            End With
            DoEvents

            DBCmEx2 "qryTerStat", "@IdSta", "@IdxNr", IdSta, TerNr 'Status ändern
            DoEvents

            If NotSt > 1 Then
                If IdSta = 1 Or IdSta = 4 Then
                    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
                    If Frage = 6 Then
                        DBCmEx2 "qryTerNotStat", "@IdSet", "@IdxNr", 1, TerNr 'E-Mail-Erinnerung
                        PrGui = CreateID("P")
                        MitNa = GlMiA(GlSmI, 1)
                        PrStr = "Emailerinnerung deaktiviert"
                        DBCmEx7 "qryTerPrAd", "@IdxNr", "@GuiId", "@TerId", "@IdDat", "@IdZei", "@IdStr", "@IdKom", TerNr, PrGui, TeGui, Now, Now, PrStr, MitNa
                        DoEvents
                    End If
                End If
            End If

            PrGui = CreateID("P")
            MitNa = GlMiA(GlSmI, 1)
            PrStr = "Terminstatusänderung: " & GlTeS(IdSta, 1)
            DBCmEx7 "qryTerPrAd", "@IdxNr", "@GuiId", "@TerId", "@IdDat", "@IdZei", "@IdStr", "@IdKom", TerNr, PrGui, TeGui, Now, Now, PrStr, MitNa
            DoEvents

            GlNeK = GlKoX 'Protokolleintrag
            With GlNeK
                .PatNr = GlMan(GlSMa, 2)
                .IdxNr = 0
                .EiDat = Format$(Date, "dd.mm.yyyy")
                .EiZei = TimeValue(Now)
                .EiTyp = 104
                .TeStr = PrStr & Space$(1) & TeGui
                .ZiStr = Format$(Now, "hh:mm") & " Uhr"
                .NeuEi = True
                .KeiAk = True
                .Mitar = GlMiA(GlSmI, 2)
            End With
            S_Prot

            Exit For
        End If
    Next ViEvt
End If

With CaCol
    .DataProvider.ChangeEvent CaEvt
    .Populate
    .RedrawControl
End With

Set CaCol = Nothing

Exit Sub

AdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeSt " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long, Optional ByVal ColID As Long, Optional ByVal CoTex As String)
On Error Resume Next

Set FM = frmMain

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F2: FTast KY_F2
Case KY_F3: FTast KY_F3
Case KY_F4: FTast KY_F4
Case KY_F5: FSuFe True
Case KY_F7: FTast KY_F7
Case KY_F6: FSam
Case KY_F8: SSave
Case KY_F9: WaMain RibTab_Wart_Wied
Case KY_F10: FPrnt
Case KY_F11: Unload FM
Case KY_F12: SAbme
Case KY_DEL: FTast KY_DEL
Case KY_CT_A: FMark
Case KY_CT_M: FMark 1
Case KY_CT_0: STaZg 0
Case KY_CT_1: If GlRch(0, 0) = 1 Then STaZg 1
Case KY_CT_2: If GlRch(0, 5) = 1 Then STaZg 2
Case KY_CT_3: If GlRch(0, 11) = 1 Then STaZg 3
Case KY_CT_4: If GlRch(0, 15) = 1 Then STaZg 4
Case KY_CT_5: If GlRch(0, 16) = 1 Then STaZg 5
Case KY_CT_6: If GlRch(0, 17) = 1 Then STaZg 6 'Textverarbeitung
Case KY_CT_7: If GlRch(0, 18) = 1 Then STaZg 7
Case KY_CT_8: If GlRch(0, 9) = 1 Then STaZg 8
Case KY_CT_9: If GlRch(0, 6) = 1 Then STaZg 9
Case KY_CT_Z: FTxRu
Case KY_CT_Y: FTxRu True
Case KY_CT_AL_D: FKrDu
Case KY_CT_AL_L: S_KrLo True

Case ME_Adresse_Hinzufuegen: SAdre 1
Case ME_Adresse_Bearbeiten: SAdre 2
Case ME_Adresse_Kopieren: S_AdKo
Case ME_Adresse_Loeschen: SLoHa
Case ME_Adresse_Filtern: AdFMa
Case ME_Gruppe_Hinzufuegen: SAGrNe
Case ME_Gruppe_Loeschen: GrAd_Loe
Case ME_Gruppe_Umbenennen: SGrUm 1
Case ME_Patient_Drucken: SDrLis 1
Case ME_SamRech_Drucken: frmReSam.Show vbModal
Case ME_Adresse_Email: SEmNe
Case ME_Adresse_SMS: SSMSd
Case ME_Adresse_Markieren: FMark
Case ME_Adresse_Termin: SAdre 4
Case ME_Adresse_TerLis: STeZe
Case ME_Adresse_Wiedervor: SAdre 5
Case ME_Adresse_ZusaFass: S_AdZu
Case ME_Adresse_Anpassen: frmAdrAnpa.Show vbModal
Case ME_Adresse_Exchange: S_AdEs

Case ME_Mandant_Hinzufuegen: SAdre 1, True
Case ME_Mandant_Bearbeiten: SAdre 2, True
Case ME_Mandant_Kopieren: S_AdKo
Case ME_Mandant_Loeschen: SLoHa
Case ME_Mandant_Filtern: AdFMa
Case ME_Mandant_Markieren: FMark
Case ME_Mandant_Email: SEmNe
Case ME_Mandant_Anpassen: frmAdrAnpa.Show vbModal

Case ME_Verord_Hinzufuegen: SAdre 1, True
Case ME_Verord_Bearbeiten: SAdre 2, True
Case ME_Verord_Kopieren: S_AdKo
Case ME_Verord_Loeschen: SLoHa
Case ME_Verord_Filtern: AdFMa
Case ME_Verord_Markieren: FMark
Case ME_Verord_Email: SEmNe
Case ME_Verord_Anpassen: frmAdrAnpa.Show vbModal

Case ME_Mitarb_Hinzufuegen: SAdre 1, True
Case ME_Mitarb_Bearbeiten: SAdre 2, True
Case ME_Mitarb_Kopieren: S_AdKo
Case ME_Mitarb_Loeschen: SLoHa
Case ME_Mitarb_Filtern: AdFMa
Case ME_Mitarb_Markieren: FMark
Case ME_Mitarb_Email: SEmNe
Case ME_Mitarb_Anpassen: frmAdrAnpa.Show vbModal

Case ME_Krankeneintrag_DatEmail: SKrDo True
Case ME_Krankeneintrag_DatExport: SKrDo False, True
Case ME_Krankeneintrag_DatUpload: SKrDo False, False, True
Case ME_Krankeneintrag_DatHinzu: SKrDa
Case ME_Kranke_Hinzufuegen: FTast KY_F3
Case ME_Krankenblatt_Loeschen: S_KrLo
Case ME_Krankenblatt_Assistent: frmKraKop.Show vbModal
Case ME_Krankeneintrag_Speichern: SSave
Case ME_Krankenblatt_Ausschn: SClip True
Case ME_Krankenblatt_Kopieren: SClip
Case ME_Krankenblatt_Einfugen: S_KrEi
Case ME_Krankenblatt_Drucken: SKrDr
Case ME_Krankenblatt_Leerzeile: S_KrZe
Case ME_Krankenblatt_Bearbeiten: SKoEd
Case ME_Krankeneintrag_Kopieren: SClip
Case ME_Krankeneintrag_Einfugen: S_KrEi
Case ME_Krankeneintrag_Duplizie: FKrDu
Case ME_Krankenblatt_Zuordnung: S_AbDi
Case ME_Krankenblatt_Anpassen: frmKraDa.Show vbModal
Case ME_Krankenblatt_Markieren: FMark 1

Case ME_Diagn_Bearbeiten: SGrUm 3
Case ME_Diagn_Loeschen: Dia_Lo
Case ME_Diagn_LokaKop: Dia_Ko
Case ME_Diagn_Ausschne: Dia_Cli True
Case ME_Diagn_Kopieren: Dia_Cli
Case ME_Diagn_Einfuege: Dia_Ein
Case ME_Diagn_Anpassen: FDiDa

Case ME_TagPro_AlCheck: FTrMa
Case ME_TagPro_UnCheck: FTrMa False, True
Case ME_TagPro_Export: SExFo 3, 2, 0

Case ME_Anamnesebo_Hinzufueg: SAnEd True
Case ME_Anamnesebo_Loeschen: SLoHa
Case ME_Anamnesebo_Bearbeiten: SAnEd
Case ME_Anamnesebo_Speichern: SSave
Case ME_Anamnesebo_Dokument: STxDi
Case ME_Anamnesebo_DruAnaAus: SDruck "AnBeBo", True
Case ME_Anamnesebo_DruAnaFra: SDruck "AnFrBo", True

Case ME_Beleg_Hinzufuegen: FTast KY_F3
Case ME_Beleg_Kopieren: SReDi 4
Case ME_Beleg_Loeschen: SLoHa
Case ME_Beleg_Speichern: SSave
Case ME_Beleg_Drucken: SDrLis 6
Case ME_Beleg_Bearbeiten: frmRzEdit.Show vbModal
Case ME_Beleg_DiagLoe: S_RzEn
Case ME_Beleg_DiagZur: S_RzZu
Case ME_Beleg_Stornierte: SAbSp 33

Case ME_Bild_Hinzufuegen: SBiEi
Case ME_Bild_Oeffnen: SBiBe
Case ME_Bild_Loeschen: SBiLo
Case ME_Bild_Grabben: SBiVi
Case ME_Bild_Scannen: SBiSc
Case ME_Bild_Umbenennen: SBiUm

Case ME_Berichte_Vergleich: LVMain
Case ME_Berichte_Importieren: STran 2
Case ME_Berichte_Exportieren: SExFo 4, 0, 0
Case ME_Berichte_Zuordnen: SLaZu
Case ME_Berichte_Bearbeiten: frmLaBear.Show vbModal
Case ME_Berichte_Loeschen: SLoHa
Case ME_Berichte_Drucken: SDruck "LabKom", True
Case ME_Berichte_Zusammen: S_LaZu
Case ME_Berichte_Parameter: S_LaLo
Case ME_Berichte_Status: S_LaSta
Case ME_Berichte_Ausschneiden: SClip True
Case ME_Berichte_Kopieren: SClip
Case ME_Berichte_Einfuegen: S_LaEi
Case ME_Berichte_ParaVerg: S_LaPV
Case ME_Berichte_Anpassen: frmKraDa.Show vbModal

Case ME_Auftrag_Hinzufuegen: frmNeuAuf.Show vbModal
Case ME_Auftrag_Bearbeiten: frmLaBear.Show vbModal
Case ME_Auftrag_Loeschen: SLoHa
Case ME_Auftrag_Drucken: SDruck "LabAuf", True
Case ME_Auftrag_Status: S_LaSta
Case ME_Auftrag_Rechnung: S_LaReA
Case ME_Auftrag_Parameter: S_LaLo
Case ME_Auftrag_Ausschneiden: SClip True
Case ME_Auftrag_Kopieren: SClip
Case ME_Auftrag_Einfuegen: S_LaEi

Case ME_Rechnung_Hinzufuegen: SReDi 1
Case ME_Rechnung_Kopieren: SReDi 2
Case ME_Rechnung_Loeschen: SLoHa
Case ME_Rechnung_Stornieren: SLoHa
Case ME_Rechnung_Bearbeiten: frmReEdit.Show vbModal
Case ME_Rechnung_Abschliessen: frmReAbs.Show vbModal
Case ME_Rechnung_Filtern: frmReFilt.Show vbModal
Case ME_Rechnung_Markieren: FMark
Case ME_Rechnung_Exportieren: frmReExpo.Show vbModal
Case ME_Rechnung_Importieren: frmImport.Show vbModal
Case ME_Rechnung_Oeffnen: SReStu
Case ME_Rechnung_Anpassen: frmReAnd.Show vbModal
Case ME_Rechnung_Serieplan: TeVoMa

Case ME_Posten_Hinzufuegen: FTast KY_F3
Case ME_Posten_Bearbeiten: frmOPEdit.Show vbModal
Case ME_Posten_Stornieren: SLoHa
Case ME_Posten_Ausgleichen: frmOPAusg.Show vbModal
Case ME_Posten_Zurueckholen: S_OPZu
Case ME_Posten_Rechnung: SReZe 0, True
Case ME_Posten_Aufheben: S_OPBe
Case ME_Posten_Exportieren: SExFo 8, 0, 0
Case ME_Posten_Drucken: SDrLis 3

Case ME_Buchung_Hinzufuegen: FTast KY_F3
Case ME_Buchung_Bearbeiten: frmBuEdit.Show
Case ME_Buchung_Zuordnen: frmBuZuo.Show vbModal
Case ME_Buchung_Ableiten: S_BuAb
Case ME_Buchung_Kopieren: S_BuKo
Case ME_Buchung_Loeschen: SLoHa
Case Me_Buchung_Bereinigen: S_BuBer
Case ME_Buchung_Stornieren: SLoHa
Case Me_Buchung_Wiederher: S_OPZu
Case ME_Buchung_Exportieren: frmBuExp.Show vbModal
Case ME_Buchung_Importieren: S_BuIm
Case ME_Buchung_Anpassen: frmBuAnp.Show vbModal
Case ME_Buchung_Dokument: SBuDo
Case ME_Buchung_Drucken: frmZeitraum.Show vbModal
Case ME_Buchung_Kontenpruf1: S_BuKt
Case ME_Buchung_Kontenpruf2: S_BuKs

Case ME_Banking_Importieren: frmImport.Show vbModal
Case ME_Banking_Exportieren:
Case ME_Banking_Bearbeiten: frmBaEdit.Show vbModal
Case ME_Banking_ZuorAufh: S_BaAuf
Case ME_Banking_Kennzei: S_BaMar
Case ME_Banking_Loeschen: SLoHa
Case ME_Banking_Ableiten: SBuAb
Case ME_Banking_SplitBuch: SBuAb True
Case ME_Banking_Vorlagen: frmBaAbl.Show vbModal
Case ME_Banking_Buchen: S_BaBuc
Case ME_Banking_Nachricht: SBaNa
Case ME_Banking_Posten: S_BaOpo
Case ME_Banking_Gutschrift:

Case ME_KassenBuchung_Hinzufuegen: FTast KY_F3
Case ME_KassenBuchung_Bearbeiten: frmBuEdit.Show
Case ME_KassenBuchung_Kopieren: S_BuKo
Case ME_KassenBuchung_Loeschen: SLoHa
Case ME_KassenBuchung_Stornieren: SLoHa
Case Me_KassenBuchung_Wieder: S_OPZu
Case ME_KassenBuchungen_Exportieren: frmBuExp.Show vbModal
Case ME_KassenBuchungen_Anpassen: frmBuAnp.Show vbModal
Case ME_KassenAuswertungen_Drucken: frmZeitraum.Show vbModal

Case ME_Termin_Hinzufuegen: STerm True
Case ME_Termin_Bearbeiten: STerm
Case ME_Termin_Ausschneiden: S_TeKo False, True
Case ME_Termin_Kopieren: S_TeKo False
Case ME_Termin_Einfuegen: S_TeKo True
Case ME_Termin_Loeschen: STeLo
Case ME_Termine_Kranken: STePa
Case ME_Termine_Drucken: STePr
Case ME_Termine_Anrufen: STeTe
Case ME_Termine_Adresse: STeAd
Case ME_Termine_PatTerm: STePl
Case ME_Termine_TerSerie: STePl True
Case ME_Termine_Rechnung: frmReErs.Show vbModal
Case ME_Termine_Buchung: S_TeBu
Case ME_Termine_MailSend: S_TeNa
Case ME_Termine_Status1: FTeSt 1
Case ME_Termine_Status2: FTeSt 2
Case ME_Termine_Status3: FTeSt 3
Case ME_Termine_Status4: FTeSt 4
Case ME_Termine_Abschlus: FTeAb
Case ME_Termine_Dokument: FTeDo

Case ME_Terminliste_Hinzufuegen: STerm True
Case ME_Terminliste_Bearbeiten: STerm
Case ME_Terminliste_Kopieren: STeKo
Case ME_Terminliste_Loeschen: SLoHa
Case ME_Terminliste_Rechnung: frmReErs.Show vbModal
Case ME_Terminliste_Buchung:
Case ME_Termineliste_Drucken: SDrLis 1
Case ME_Termineliste_Anrufen: STeTe
Case ME_Terminliste_Zahlung: STeZa
Case ME_Terminliste_Storno: STeSt
Case ME_Terminliste_Anpassen: frmTerAnp.Show vbModal
Case ME_Terminliste_Nachricht: S_TeMa
Case ME_Terminliste_Durchnummer:  S_TeNu

Case ME_Raumtermin_Hinzufuegen: STerm True
Case ME_Raumtermin_Bearbeiten: STerm
Case ME_Raumtermin_Kopieren: STeKo
Case ME_Raumtermin_Loeschen: STeLo
Case ME_Raumtermine_Drucken: STePr

Case KM_Eint_Hinzufuegen: KaNeu
Case KM_Eint_Bearbeiten: KaEdi
Case KM_Eint_Kopieren: frmKaKop.Show vbModal
Case KM_Eint_Loeschen: K_Loe
Case KM_Eint_Diagnose: TrMain 1
Case KM_Eint_Anpassen: frmKaAnp.Show vbModal

Case KM_Kett_Hinzufuegen: KaNeu
Case KM_Kett_Bearbeiten: KaEdi
Case KM_Kett_Kopieren: frmKaKop.Show vbModal
Case KM_Kett_Loeschen: K_Loe
Case KM_Kett_Aktualisieren: K_Bere

Case KM_Frage_Hinzufuegen: FaNeu 1
Case KM_Frage_Bearbeiten: KaEdi
Case KM_Frage_Kopieren: frmKaKop.Show vbModal
Case KM_Frage_Loeschen: F_Loe
Case KM_Frage_Anpassen:

Case Tex_Mail_Hinzufuegen: SMaNe
Case Tex_Mail_Loeschen: SLoHa
Case Tex_Mail_Antworten: SMaAn 3
Case Tex_Mail_Verschieben: frmKaKop.Show vbModal
Case Tex_Mail_Archivieren: S_MaArc
Case Tex_Mail_Ungelesen: S_MaMa 1
Case Tex_Mail_Markieren: S_MaMa 2
Case Tex_Mail_Junkmail: S_MaMa 3
Case Tex_Mail_Rechnung: MaSav 8
Case Tex_Mail_Kommentar: S_MaKom
Case Tex_Mail_Thread: MaThr
Case Tex_Mail_PatSuch: frmAdrSuch.Show vbModal
Case Tex_Mail_PatEdit: MaAdr
Case Tex_Mail_PatPruf: S_MaAdr
Case Tex_Gruppe_Hinzufuegen: SMGrNe
Case Tex_Gruppe_Loeschen: GrMa_Loe
Case Tex_Gruppe_Umbenennen: SMGrU 1

Case KM_Katalog_Hinzufuegen: KGrNe
Case KM_Katalog_Loeschen: K_GrLo
Case KM_Katalog_Umbenennen: KGrUm
Case KM_Katalog_Exportieren: K_ExIm 1
Case KM_Katalog_Importieren: K_ExIm 2
Case KM_Katalog_Anfuegen: K_ExIm 3
Case KM_Katalog_WebLink: KFrKo
Case KM_Katalog_WebDele: S_AnBoL
Case KM_Katalog_EmaLink: KFrKo True

Case ME_PaBi_Grabben: SVido
Case ME_PaBi_Hinzufuegen: SPhot
Case ME_PaBi_Loeschen: SPhot True
Case ME_PaBi_Autoimport: SBiIo True
Case ME_PaBi_Bearbeiten: SPhoE

Case RibCon_Zeilenumbruch: If GlAkt = False Then SGrLa "GrdZei"
Case RibCon_Zeilenmarker: If GlAkt = False Then SGrLa "GrdMkr"
Case RibCon_Gitternetzlin: If GlAkt = False Then SGrLa "GrdGrl"
Case RibCon_Vorschauzeile: If GlAkt = False Then SGrLa "GrdPrv"
Case RibCon_Tooltips: If GlAkt = False Then SGrLa "GrdTip"
Case RibCon_SaveLayout: If GlAkt = False Then SSpSav
Case RibCon_Schnelldruck: If GlAkt = False Then PrtMain 3
Case RibCon_Layout: frmLayout.Show vbModal

Case ME_Termin_Minuten_01: SKaSc 1
Case ME_Termin_Minuten_05: SKaSc 5
Case ME_Termin_Minuten_10: SKaSc 10
Case ME_Termin_Minuten_15: SKaSc 15
Case ME_Termin_Minuten_20: SKaSc 20
Case ME_Termin_Minuten_30: SKaSc 30
Case ME_Termin_Minuten_60: SKaSc 60
Case ME_Termin_Minuten_120: SKaSc 120

Case RibCon_Dbd_Oeffnen: SReb 1, True
Case RobCon_Dbd_Sichern: SReb 2, True
Case RibCon_Dbd_Zurueck: SReb 3, True
Case RibCon_Dbd_Wechsel: frmSQLCon.Show vbModal
Case RibCon_Dbd_Hinzuf: SDaNe
Case RibCon_Dbd_GoBDExp: SGoBD True
Case RibCon_Dbd_TAR_Exp: TSEExp
Case RibCon_Dbd_GoBAbsh: SGoBD
Case RibCon_Dbd_Pruefen: DBWart
Case RibCon_Outlookabg: OuMain
Case RibCon_Abmelden: SAbme
Case RibCon_Aufgaben: WaMain RibTab_Wart_Wied
Case RibCon_Benutzer: SMaZe
Case RibCon_Lizenz: SLize 3
Case RibCon_PrgInfo:  SLize 1
Case RibCon_Formulare: frmFormular.Show vbModal
Case RibCon_Optionen: frmOptions.Show
Case RibCon_Hilfe: FHilfe
Case RibCon_Ifap: SIfap
Case RibCon_SmartCard: frmChipcard.Show vbModal
Case RibCon_Beenden: Unload FM
Case RibCon_Neustart: SNeSt
Case RibCon_Reset: SRest
Case RibCon_Wegamed: SWega
Case RibCon_GDT_Appli: SGDT
Case RibCon_Remote: SRemo
Case RibCon_Refresh: SBuLa True

Case SY_AD_Adresse_Hinzufueg: SAdre 1
Case SY_AD_Adresse_Bearbeiten: SAdre 2
Case SY_AD_Adresse_Kopieren: S_AdKo
Case SY_AD_Adresse_Loeschen: SLoHa
Case SY_AD_Adresse_Vollst: FSuAu
Case SY_MA_Mandant_Vollst: FSuAu
Case SY_VE_Verord_Vollst: FSuAu
Case SY_MI_Mitarb_Vollst: FSuAu
Case SY_AD_Adresse_Filtern: AdFMa
Case SY_AD_Adresse_Serien: STxSe False
Case SY_AD_Adresse_Einzel: STxPa 2, 2
Case SY_AD_Adresse_Wiederv: SAdre 5
Case SY_AD_Adressen_Drucken: SDrLis 1
Case SY_AD_Adressen_Mark: FMark
Case SY_AD_Adressen_ExpImp: SExFo 1, 2, 0
Case SY_AD_Adressen_Export: SExFo 1, 2, 0
Case SY_AD_Adressen_Import: frmImport.Show vbModal
Case SY_AD_Adressen_Email: SEmNe
Case SY_AD_Adresse_Suchen: FSuFe
Case SY_AD_Adresse_Eamil: SEmNe
Case SY_AD_Adresse_SMS: SSMSd
Case SY_AD_Adresse_SerMa: SEmNe True
Case SY_AD_Adresse_Termin: SAdre 4
Case SY_AD_Adresse_TerLis: STeZe

Case SY_MA_Mandant_Hinzufueg: SAdre 1, True
Case SY_MA_Mandant_Bearbeiten: SAdre 2, True
Case SY_MA_Mandant_Kopieren: S_AdKo
Case SY_MA_Mandant_Loeschen: SLoHa
Case SY_MA_Mandant_TxBr: STxPa 2, 2
Case SY_MA_Mandant_TxVo: STxVo
Case SY_MA_Mandant_Drucken: SDrLis 1
Case SY_MA_Mandant_Mark: FMark
Case SY_MA_Mandant_Export: SExFo 1, 2, 0
Case SY_MA_Mandant_Email: SEmNe
Case SY_MA_Mandant_Suchen: FSuFe

Case SY_MI_Mitarb_Hinzufueg: SAdre 1, True
Case SY_MI_Mitarb_Bearbeiten: SAdre 2, True
Case SY_MI_Mitarb_Kopieren: S_AdKo
Case SY_MI_Mitarb_Loeschen: SLoHa
Case SY_MI_Mitarb_TxBr: STxPa 2, 2
Case SY_MI_Mitarb_TxVo: STxVo
Case SY_MI_Mitarb_Drucken: SDrLis 1
Case SY_MI_Mitarb_Mark: FMark
Case SY_MI_Mitarb_Export: SExFo 1, 2, 0
Case SY_MI_Mitarb_Email: SEmNe
Case SY_MI_Mitarb_Suchen: FSuFe

Case SY_VE_Verord_Hinzufueg: SAdre 1, True
Case SY_VE_Verord_Bearbeiten: SAdre 2, True
Case SY_VE_Verord_Kopieren: S_AdKo
Case SY_VE_Verord_Loeschen: SLoHa
Case SY_VE_Verord_TxBr: STxPa 2, 2
Case SY_VE_Verord_TxVo: STxVo
Case SY_VE_Verord_Drucken: SDrLis 1
Case SY_VE_Verord_Mark: FMark
Case SY_VE_Verord_Export: SExFo 1, 2, 0
Case SY_VE_Verord_Import: frmImport.Show vbModal
Case SY_VE_Verord_Email: SEmNe
Case SY_VE_Verord_Suchen: FSuFe

Case SY_KB_KraBla_Datei: SKrDa
Case SY_KB_KraBla_Hinzufueg: KrMain -2
Case SY_KB_KraBla_Suchen: S_KrSu
Case SY_KB_KraBla_Vollst: SAbSo
Case SY_KB_KraBla_Loeschen: S_KrLo
Case SY_KB_KraBla_Umschalt: SUmsa
Case SY_KB_KraBla_Speichern: SSave
Case SY_KB_KraBla_Import: SBiIm
Case SY_KB_KraBla_Downl: SMDDo
Case SY_KB_KraBla_AdrSuch: FSuFe
Case SY_KB_KraBla_AdrHinz:  SAdre 1
Case SY_KB_KraBla_AdrBear:  SAdre 3
Case SY_KB_KraBla_Dokument: STxDi
Case SY_KB_KraBla_Befund: KrMain 103
Case SY_KB_KraBla_Drucken: SKrDr
Case SY_KB_KraBla_Email: SKrDr 1
Case SY_KB_KraBla_Auswa: If GlAkt = False Then SSuFe
Case SY_KB_KraBla_Typen: SUmsa
Case SY_KB_KraBla_Grupp: SKrGr
Case SY_KB_KraBla_Sorti: SAbSo
Case SY_KB_KraBla_Jahre: If GlAkt = False Then SSuFe
Case SY_KB_KraBla_Quart: If GlAkt = False Then SSuFe
Case SY_KB_KraBla_Monat: If GlAkt = False Then SSuFe
Case SY_KB_KraBla_Woche: If GlAkt = False Then SSuFe
Case SY_KB_KraBla_Datum: If GlAkt = False Then SSuFe
Case SY_KB_KraBla_Expor: STxMa True, False 'Dokument Exportieren
Case SY_KB_KraBla_Nachr: STxMa False, True 'Dokument als Nachricht
Case SY_KB_KraBla_DoLnk: STxMa True, False, 1 'Dokument Downloadlink
Case RibCon_KB_Ansicht:

Case SY_AN_AnaBog_AdrSuch: FSuFe
Case SY_AN_AnaBog_AdrHinz: SAdre 1
Case SY_AN_AnaBog_AdrBear: SAdre 3
Case SY_AN_AnaBog_Hinzufueg: SAnEd True
Case SY_AN_AnaBog_Bearbeiten: SAnEd
Case SY_AN_AnaBog_Loeschen: SLoHa
Case SY_AN_AnaBog_Speichern: SSave
Case SY_AN_AnaBog_Dokument: STxDi
Case SY_AN_AnaBog_TexNPa: STxPa 2
Case SY_AN_AnaBog_TexNAr: STxPa 4
Case SY_AN_AnaBog_TexNVe: STxPa 3
Case SY_AN_AnaBog_TexNPL: STxPa 5
Case SY_AN_AnaBog_TexVor: STxVo
Case SY_AN_AnaBog_DocSe2: FTxMa 2 'Dokument Digitalunterschrift
Case SY_AN_AnaBog_DocSe3: SNfSe
Case SY_AN_AnaBog_TexPla: STxEi 5
Case SY_AN_AnaBog_TexSMS: SNaSe
Case SY_AN_AnaBog_GDTExp: S_AdGDT
Case SY_AN_AnaBog_GDTImp: S_GDT True
Case SY_AN_AnaBog_TexEml: FEmal
Case SY_AN_AnaBog_Termin: SAdTe
Case SY_AN_AnaBog_Wieder: SWied
Case SY_AN_AnaBog_Warte: frmWaKom.Show vbModal
Case SY_AN_AnaBog_ExpFra: SExFo 2, 4, 1
Case SY_AN_AnaBog_ExpBer: SExFo 2, 4, 2
Case SY_AN_AnaBog_DruFra: SDruck "AnFrBo", True
Case SY_AN_AnaBog_EmaAus: SDruck "AnBeBo", True, , , , , 1
Case SY_AN_AnaBog_DruAus: SDruck "AnBeBo", True
Case SY_AN_AnaBog_EmaFra: SDruck "AnFrBo", True, , , , , 1
Case SY_AN_AnaBog_Import: S_AnBoC
Case SY_AN_AnaBog_Unlock: FFrAb

Case SY_TB_TagPro_AlCheck: FTrMa
Case SY_TP_TagPro_UnCheck: FTrMa False, True
Case SY_TP_TagPro_Suchen: FTaSu
Case SY_TP_TagPro_Vollst: SSuch
Case SY_TP_TagPro_Auswahl: If GlAkt = False Then SSuFe
Case SY_TP_TagPro_Jahre:  If GlAkt = False Then SSuFe
Case SY_TP_TagPro_Quart: If GlAkt = False Then SSuFe
Case SY_TP_TagPro_Monat: If GlAkt = False Then SSuFe
Case SY_TP_TagPro_Woche: If GlAkt = False Then SSuFe
Case SY_TP_TagPro_Datum: If GlAkt = False Then SSuFe
Case SY_TP_TagPro_Expan: STaEx
Case SY_TP_TagPro_Export: SExFo 3, 2, 0
Case SY_TP_TagPro_Drucken: STaDr
Case SY_TP_TagPro_Typen: SUmsa
Case SY_TP_TagPro_Grupp: STaGr
Case SY_TP_TagPro_Sorti: SAbSo
Case SY_VB_Vorbe_AdrBear: SVoAd 1
Case SY_VB_Vorbe_AdrKrab: SVoAd 2
Case SY_VB_Vorbe_AdrMar: FMark
Case SY_VB_Vorbe_Abges: SVoAb
Case SY_VB_Vorbe_Typ01: SUmsa 1
Case SY_VB_Vorbe_Typ02: SUmsa 2
Case SY_VB_Vorbe_Typ03: SUmsa 3
Case SY_VB_Vorbe_Mitar: SUmsa 4
Case SY_VB_Vorbe_Auswahl: If GlAkt = False Then SSuFe
Case SY_VB_Vorbe_Jahre: If GlAkt = False Then SSuFe
Case SY_VB_Vorbe_Monat: If GlAkt = False Then SSuFe
Case SY_VB_Vorbe_Datum: If GlAkt = False Then SSuFe
Case SY_VB_Vorbe_Woche: If GlAkt = False Then SSuFe
Case SY_VB_Vorbe_Quart: If GlAkt = False Then SSuFe
Case SY_VB_Vorbe_Erstel: frmReGen.Show vbModal
Case SY_VB_Vorbe_Email: SVoDr True
Case SY_VB_Vorbe_Export: SVoDr True, True
Case SY_VB_Vorbe_Liste: SAbDr
Case SY_VB_Vorbe_Druck: SDrLis 1
Case SY_AB_Abrech_ReHinzufueg: SReDi 1
Case SY_AB_Abrech_ReBearbeit: frmReEdit.Show vbModal
Case SY_AB_Abrech_RechTyp1: S_ReTyp "R"
Case SY_AB_Abrech_RechTyp2: S_ReTyp "V"
Case SY_AB_Abrech_RechTyp3: S_ReTyp "L"
Case SY_AB_Abrech_RechTyp4: S_ReTyp "A"
Case SY_AB_Abrech_RechTyp5: S_ReTyp "U"
Case SY_AB_Abrech_RechTyp6: S_ReTyp "M"
Case SY_AB_Abrech_RechTyp7: S_ReTyp "G"
Case SY_AB_Abrech_RechTyp8: S_ReTyp "I"
Case SY_AB_Abrech_ReLoeschen: SLoHa
Case SY_AB_Abrech_ReStornieren: SLoHa
Case SY_AB_Abrech_ReKopieren: SReDi 2
Case SY_AB_Abrech_Speichern: SSave
Case SY_AB_Abrech_AdrSuch: FSuFe
Case SY_AB_Abrech_AdrBear:  SAdre 3
Case SY_AB_Abrech_AdrHinz:  SAdre 1
Case SY_AB_Abrech_Zahlung: frmAnzahl.Show vbModal
Case SY_AB_Abrech_Assistent: frmKraKop.Show vbModal
Case SY_AB_Abrech_Loeschen: S_KrLo
Case SY_AB_Abrech_Drucken: SDrLis 2
Case SY_AB_Abrech_Email: SDrLis 2, 1
Case SY_AB_Abrech_Downl: SDrLis 2, 6
Case SY_AB_Abrech_Splitten: S_ReTei
Case SY_AB_Datum_Aendern: frmKraDa.Show vbModal
Case SY_AB_Ausschneiden: SClip True
Case SY_AB_Kopieren: SClip
Case SY_AB_Einfuegen: S_KrEi
Case SY_AB_Einf_Datum: S_KrEi True
Case SY_AB_Abrech_ReHinSta: SReDi 1, "R"
Case SY_AB_Abrech_ReHinAng: SReDi 1, "V"
Case SY_AB_Abrech_ReHinLab: SReDi 1, "L"
Case SY_AB_Abrech_ReHinGut: SReDi 1, "U"
Case SY_AB_Abrech_ReHinAuf: SReDi 1, "M"
Case SY_AB_Abrech_ReHinGew: SReDi 1, "G"
Case RibCon_AB_Ansicht:
Case SY_AB_Abrech_Expan: SAbEx
Case SY_AB_Abrech_Grupp: SAbGr
Case SY_AB_Abrech_Sorti: SAbSo

Case SY_RZ_Rezept_Hinzufueg: FTast KY_F3
Case SY_RZ_Rezept_Bearbeiten: frmRzEdit.Show vbModal
Case SY_RZ_Rezept_Kopieren: SReDi 4
Case SY_RZ_Rezept_Loeschen: SLoHa
Case SY_RZ_Rezept_Speichern: SSave
Case SY_RZ_Rezept_AdrSuch: FSuFe
Case SY_RZ_Rezept_AdrBear:  SAdre 3
Case SY_RZ_Rezept_AdrHinz:  SAdre 1
Case SY_RZ_Rezept_Ifap: SIfap
Case SY_RZ_Rezept_Drucken: SDrLis 6
Case SY_RZ_Rezept_Email: SDrLis 6, 1
Case SY_RZ_Rezept_Downl: SDrLis 6, 6
Case SY_RZ_Rezept_Vorlage: FRzBe

Case SY_RZ_Beleg_Hinzufueg: FTast KY_F3
Case SY_RZ_Beleg_Bearbeiten: frmRzEdit.Show vbModal
Case SY_RZ_Beleg_Kopieren: SReDi 4
Case SY_RZ_Beleg_Loeschen: SLoHa
Case SY_RZ_Beleg_Speichern: SSave
Case SY_RZ_Beleg_AdrSuch: FSuFe
Case SY_RZ_Beleg_AdrBear:  SAdre 3
Case SY_RZ_Beleg_AdrHinz:  SAdre 1
Case SY_RZ_Beleg_Ifap: SIfap
Case SY_RZ_Beleg_Drucken: SDrLis 6
Case SY_RZ_Beleg_Email: SDrLis 6, 1
Case SY_RZ_Beleg_Downl: SDrLis 6, 6
Case SY_RZ_Beleg_Vorlage: FRzBe

Case SY_BI_Bild_Einfuegen: SBiEi
Case SY_BI_Bild_Oeffnen: SBiBe
Case SY_BI_Bild_Loeschen: SBiLo
Case SY_BI_Bild_Umbenenn: SBiUm
Case SY_BI_Bild_Grabben: SBiVi
Case SY_BI_Bild_AdrSuch: FSuFe
Case SY_BI_Bild_AdrBear: SAdre 3
Case SY_BI_Bild_AdrHinz: SAdre 1
Case SY_BI_Bild_Scannen: SBiSc
Case SY_BI_Bild_Drucken: SBiBe
Case SY_BI_Bild_Einlesen: SBiIm
Case SY_BI_Bild_Ansicht: KDatei 22
Case SY_BI_Bild_Sorter: KDatei 24
Case SY_BI_Bild_Thumbna: KDatei 26

Case SY_RE_Rechnung_Hinzufueg: SReDi 1
Case SY_RE_Rechnung_Bearbeiten: frmReEdit.Show vbModal
Case SY_RE_Rechnung_Kopieren: SReDi 2
Case SY_RE_Rechnung_Loeschen: SLoHa
Case SY_RE_Rechnung_Stornieren: SLoHa
Case SY_RE_Rechnung_Filtern: frmReFilt.Show vbModal
Case SY_RE_Rechnung_Vollst: FSuAu
Case SY_RE_Rechnung_Exportieren: frmReExpo.Show vbModal
Case SY_RE_Rechnung_Importieren: frmImport.Show vbModal
Case SY_RE_Rechnung_Drucken: SDrLis 2
Case SY_RE_Rechnung_Email: SDrLis 2, 1
Case SY_RE_Rechnung_Server: SDrLis 2, 6
Case SY_RE_Rechnung_Export: SDrLis 2, 5
Case SY_RE_Rechnung_Listendruck: SDrLis 4
Case SY_RE_Rechnung_ListEmail: SDrLis 4, 1
Case SY_RE_Rechnung_Abschliessen: frmReAbs.Show vbModal
Case SY_RE_Rechnung_Status: SReStu
Case SY_RE_Rechnung_Splitten: S_ReTei
Case SY_RE_Rechnung_Anpassen: frmReAnd.Show vbModal
Case SY_RE_Rechnung_Suchen: FSuFe
Case SY_RE_Rechnung_HinSta: SReDi 1, "R"
Case SY_RE_Rechnung_HinAng: SReDi 1, "V"
Case SY_RE_Rechnung_HinLab: SReDi 1, "L"
Case SY_RE_Rechnung_HinGut: SReDi 1, "U"
Case SY_RE_Rechnung_HinGew: SReDi 1, "G"
Case SY_RE_Rechnung_HinAuf: SReDi 1, "M"

Case SY_PO_Posten_Hinzufueg: FTast KY_F3
Case SY_PO_Posten_Bearbeiten: SReZe 0, True
Case SY_PO_Posten_Ausgleichen: frmOPAusg.Show vbModal
Case SY_PO_Posten_Stornieren: SLoHa
Case SY_PO_Posten_Exportieren: SExFo 8, 0, 0
Case SY_PO_Posten_Vollst: FSuAu
Case SY_PO_Mahnst_Raufsetzen: S_OPMa True, True
Case SY_PO_Mahnst_Runtersetzen: S_OPMa False, True
Case SY_PO_Posten_Listen: SDrLis 5
Case SY_PO_Posten_ListEmail: SDrLis 5, 1
Case SY_PO_Posten_Drucken: SDrLis 3
Case SY_PO_Posten_Email: SDrLis 3, 1
Case SY_PO_Posten_Export: SDrLis 3, 5
Case SY_PO_Mahnst_Zurueck: S_OPZu
Case SY_PO_Posten_Suchen: FSuFe
Case SY_PO_Posten_Saldo1:
Case SY_PO_Posten_Saldo2:
Case SY_PO_Posten_Saldo3:

Case SY_BU_Buchung_Hinzufuegen: FTast KY_F3
Case SY_BU_Buchung_Zuordnen: frmBuZuo.Show vbModal
Case SY_BU_Buchung_Ableiten: S_BuAb
Case SY_BU_Buchung_Loeschen: SLoHa
Case SY_BU_Buchung_Stornieren: SLoHa
Case SY_BU_Buchung_Wiederhers: S_OPZu
Case SY_BU_Buchung_Kontenplan: frmBuKont.Show vbModal
Case SY_BU_Buchung_Suchen: FSuFe
Case SY_BU_Buchung_Vollst: FSuAu
Case SY_BU_Buchung_Exportieren: frmBuExp.Show vbModal
Case SY_BU_Buchung_Anfang: frmBuAnf.Show
Case SY_BU_Buchung_Drucken: frmZeitraum.Show vbModal
Case SY_BU_Buchung_Sort1: SBuSo 1
Case SY_BU_Buchung_Sort2: SBuSo 2
Case SY_BU_Buchung_Saldo:
Case SY_BU_Buchung_Zaehlung: frmZaehlung.Show vbModal
Case SY_BU_Buchung_BuchEinfach: FDoBu 1
Case SY_BU_Buchung_BuchDoppelt: FDoBu 2
Case SY_BU_Buchung_Einrichtung:

Case SY_BA_Banking_Ableiten: SBuAb
Case SY_BA_Banking_Buchen: S_BaBuc
Case SY_BA_Banking_Bearbeiten: frmBaEdit.Show vbModal
Case SY_BA_Banking_Assistent: S_BaSuc
Case SY_BA_Banking_ZuorAufh: S_BaAuf
Case SY_BA_Banking_Anpassen: frmBaAnp.Show vbModal
Case SY_BA_Banking_Löschen: SLoHa
Case SY_BA_Banking_Suchen: FSuFe
Case SY_BA_Banking_Vollst: FSuAu
Case SY_BA_Banking_Import: frmImport.Show vbModal
Case SY_BA_Banking_Export: SExFo 12, 0, 0
Case SY_BA_Banking_Drucken: SDrLis 7
Case SY_BA_Banking_Sort1: SBaSo 1
Case SY_BA_Banking_Sort2: SBaSo 2
Case SY_BA_Banking_Saldo:

Case SY_TE_Termin_Hinzufu: STerm True
Case SY_TE_Termin_Bearbeiten: STerm
Case SY_TE_Termin_Duplizieren: STeKo
Case SY_TE_Termin_Loeschen: STeLo
Case SY_TE_Termin_Tag: SKale 1
Case SY_TE_Termin_ArWoche: SKale 2
Case SY_TE_Termin_ErWoche: SKale 3
Case SY_TE_Termin_Woche: SKale 4
Case SY_TE_Termin_Monat: SKale 5
Case SY_TE_Termin_Heute: FKaHe
Case SY_TE_Termin_ZeiSto: STeAn
Case SY_TE_Termin_Drucken: SDrLis 1
Case SY_TE_Termin_Email: SDrLis 1, 1
Case SY_TE_Termin_Assist:
Case SY_TE_Termin_Urlaub: TerAs True
Case SY_TE_Termin_KopAss: TerAs
Case SY_TE_Termin_SerTer: STeVo
Case SY_TE_Termin_FiltTyp: FTeFi
Case SY_TE_Termin_FiltIdx: FMona True
Case SY_TE_Termin_Kopiere: S_TeKo False
Case SY_TE_Termin_Einfueg: S_TeKo True
Case SY_TE_Termin_PatZei: STePa
Case SY_TE_Termin_PatNeu: SAdre 1
Case SY_TE_Termin_Aktual: FMona
Case SY_TE_Termin_Ausschn: STeAu
Case SY_TE_Termin_Docume:
Case SY_TE_Termin_DocVor: STeNa 0
Case SY_TE_Termin_DocBse: STeNa 1
Case SY_TE_Termin_DocBes: STeNa 2
Case SY_TE_Termin_DocEri: STeNa 3
Case SY_TE_Termin_DocAbs: STeNa 4
Case SY_TE_Termin_DocVrs: STeNa 5
Case SY_TE_Termin_EmlBes: STeNa 6
Case SY_TE_Termin_EmlEri: STeNa 7
Case SY_TE_Termin_EmlAbs: STeNa 8
Case SY_TE_Termin_EmlVrs: STeNa 9
Case SY_TE_Termin_SMSBes: STeNa 10
Case SY_TE_Termin_SMSEri: STeNa 11
Case SY_TE_Termin_SMSAbs: STeNa 12
Case SY_TE_Termin_SMSVrs: STeNa 13
Case SY_TE_Termin_DocSto: STeNa 14
Case SY_TE_Termin_EmlSto: STeNa 15
Case SY_TE_Termin_SMSSto: STeNa 16

Case SY_TL_Terminliste_Hinzufuegen: STerm True
Case SY_TL_Terminliste_Bearbeiten: STerm
Case SY_TL_Terminliste_Kopieren: STeKo
Case SY_TL_Terminliste_Loeschen: SLoHa
Case SY_TL_Terminliste_Rechnungen: frmReErs.Show vbModal
Case SY_TL_Terminliste_Vollst: SUmsa
Case SY_TL_Terminliste_Exportieren: SExFo 7, 0, 0
Case SY_TL_Terminliste_Importieren: frmImport.Show vbModal
Case SY_TL_Terminliste_TermReset:
Case SY_TL_Terminliste_SyncReset: S_TeEx
Case SY_TL_Terminliste_OnTeReset: S_TeOn
Case SY_TL_Terminliste_Terminfar: S_TeFa
Case SY_TL_Terminliste_Document:
Case SY_TL_Terminliste_Brief_Vor: STeNa 0
Case SY_TL_Terminliste_Brief_Pat: STeNa 20
Case SY_TL_Terminliste_Brief_Bse: STeNa 1
Case SY_TL_Terminliste_Brief_Bes: STeNa 2
Case SY_TL_Terminliste_Brief_Eri: STeNa 3
Case SY_TL_Terminliste_Brief_Vrs: STeNa 4
Case SY_TL_Terminliste_Brief_Abs: STeNa 5
Case SY_TL_Terminliste_Email_Bes: STeNa 6
Case SY_TL_Terminliste_Email_Eri: STeNa 7
Case SY_TL_Terminliste_Email_Abs: STeNa 8
Case SY_TL_Terminliste_Email_Vrs: STeNa 9
Case SY_TL_Terminliste_SMS_Bes: STeNa 10
Case SY_TL_Terminliste_SMS_Eri: STeNa 11
Case SY_TL_Terminliste_SMS_Abs: STeNa 12
Case SY_TL_Terminliste_SMS_Vrs: STeNa 13
Case SY_TL_Terminliste_Brief_Sto: STeNa 14
Case SY_TL_Terminliste_Email_Sto: STeNa 15
Case SY_TL_Terminliste_SMS_Sto: STeNa 16
Case SY_TL_Terminliste_Listendruck: SDrLis 1
Case SY_TL_Terminliste_Zahlung: STeZa
Case SY_TL_Terminliste_Skonto: STeSt
Case SY_TL_Terminliste_Suchen: FSuFe
Case SY_TL_Terminliste_Filt1: SUmsa 1
Case SY_TL_Terminliste_Filt2: SUmsa 2
Case SY_TL_Terminliste_Filt3: SUmsa 3
Case SY_TL_Terminliste_Geburtstage: Adr_Geb
Case SY_TL_Terminliste_WarteSta: SWaSe 1
Case SY_TL_Terminliste_WarteInBe: SWaSe 2
Case SY_TL_Terminliste_WarteAuBe: SWaSe 3
Case SY_TL_Terminliste_WarteEnd: SWaSe 4

Case SY_LB_Labor_AdrHinz: SAdre 1
Case SY_LB_Labor_AdrBear: SAdre 3
Case SY_LB_Labor_AdrSuch: FSuFe
Case SY_LB_Labor_Importieren: frmLaImp.Show vbModal
Case SY_LB_Labor_Bearbeiten: frmLaBear.Show vbModal
Case SY_LB_Labor_Loeschen: SLoHa
Case SY_LB_Labor_Speichern: SSave
Case SY_LB_Labor_Suchen: FSuFe
Case SY_LB_Labor_Vollst: FSuAu
Case SY_LB_Labor_Zuordnen: SLaZu
Case SY_LB_Labor_Vergleich: LVMain
Case SY_LB_Labor_Status: S_LaSta
Case SY_LB_Param_Loeschen: S_LaLo
Case SY_LB_Labor_Rechn_Neu: frmReErs.Show vbModal
Case SY_LB_Labor_Drucken: SDruck "LabKom", True
Case SY_LB_Labor_Export: SExFo 4, 0, 0
Case SY_LB_Labor_Zusammen: S_LaZu
Case SY_LB_Labor_ParVerg: S_LaPV
Case SY_LB_Abrech_Expan: SLeEx
Case SY_LB_Abrech_Grupp: SLeGr
Case SY_LB_Abrech_GrZei: SLeZe

Case SY_LA_Auftrag_AdrHinz: SAdre 1
Case SY_LA_Auftrag_AdrBear: SAdre 3
Case SY_LA_Auftrag_AdrSuch: FSuFe
Case SY_LA_Auftrag_Hinzufuegen: frmNeuAuf.Show vbModal
Case SY_LA_Auftrag_Bearbeiten: frmLaBear.Show vbModal
Case SY_LA_Auftrag_Loeschen: SLoHa
Case SY_LA_Auftrag_Speichern: SSave
Case SY_LA_Auftrag_Suchen: FSuFe
Case SY_LA_Auftrag_Vollst: FSuAu
Case SY_LA_Auftrag_Drucken: SDruck "LabAuf", True
Case SY_LA_Auftrag_Status: S_LaSta
Case SY_LA_Auftrag_Rechnung: S_LaReA
Case SY_LA_Auftrag_Bericht: S_LaBrt
Case SY_LA_Param_Loeschen: S_LaLo
Case SY_LA_Auftrag_Exportieren: SExFo 5, 0, 0

Case KA_Eint_Hinzufuegen: KaNeu
Case KA_Eint_Kopieren: frmKaKop.Show vbModal
Case KA_Eint_Loeschen: K_Loe
Case KA_Eint_Bearbeiten: KaEdi
Case KA_Eint_Suchen: FSuFe
Case KA_Eint_Drucken: KPrint "KatLis", False
Case KA_Eint_Druck01: KPrint "BesLis", False
Case KA_Eint_Druck02: KPrint "InvLis", False
Case KA_Eint_Vollst: FSuAu
Case KA_Eint_Favoriten: FSuFa
Case KA_Kat_Hinzufuegen: KGrNe
Case KA_Kat_Loeschen: K_GrLo
Case KA_Kat_Umbenennen: KGrUm
Case KA_Eint_Import: K_ExIm 2
Case KA_Eint_Export: K_ExIm 1
Case KA_Eint_Anfuegen: K_ExIm 3

Case KA_Kett_Hinzufuegen: KaNeu
Case KA_Kett_Kopieren: frmKaKop.Show vbModal
Case KA_Kett_Loeschen: K_Loe
Case KA_Kett_Bearbeiten: KaEdi
Case KA_Kett_Suchen: FSuFe
Case KA_Kett_Vollst: FSuAu
Case KA_Kett_Drucken: KPrint "KetLis", False

Case KA_Frage_Hinzufuegen: FaNeu 1
Case KA_Frage_Kopieren: frmKaKop.Show vbModal
Case KA_Frage_Loeschen: F_Loe
Case KA_Frage_Bearbeiten: KaEdi
Case KA_Frage_Suchen: FSuFe
Case KA_Frage_Vollst: FSuAu
Case KA_Frage_ExpoNorm: FPubl
Case KA_Frage_Drucken:

Case TX_Mail_Hinzufuegen: SMaNe
Case TX_Mail_Antworten: SMaAn 3
Case TX_Mail_Weiterleiten: SMaAn 4
Case TX_Mail_Loeschen: SLoHa
Case TX_Mail_Suchen: FSuFe
Case TX_Mail_Vollst: FSuAu
Case TX_Mail_Erneut: SMaAn 5
Case TX_Mail_Ungelesen: S_MaMa 1
Case TX_Mail_Markieren: S_MaMa 2
Case TX_Mail_Junkmail: S_MaMa 3
Case TX_Mail_Empfangen: S_MaAbr
Case TX_Mail_Konten: frmMailKont.Show vbModal
Case TX_Mail_PatSuch: frmAdrSuch.Show vbModal
Case TX_Mail_PatEdit: MaAdr
Case TX_Mail_AttOpen: MaSav 5
Case TX_Mail_AttSave: MaSav 6
Case TX_Mail_AttExpo: MaSav 7
Case TX_Mail_Rechnun: MaSav 8
Case TX_Mail_AttImpo: MaSav 9

Case SY_AD_Adresse_SortFeld: If GlAkt = False Then SSuFe
Case SY_AD_Adresse_SuchCombo: If GlAkt = False Then SSuFe
Case SY_MA_Mandant_SuchCombo: If GlAkt = False Then SSuFe
Case SY_VE_Verord_SuchCombo: If GlAkt = False Then SSuFe
Case SY_MI_Mitarb_SuchCombo: If GlAkt = False Then SSuFe
Case SY_RE_Rechnung_SuchCombp: If GlAkt = False Then SSuFe
Case SY_RE_Rechnung_Belegtyp: If GlAkt = False Then SSuFe
Case SY_PO_Posten_SuchCombo: If GlAkt = False Then SSuFe
Case SY_BU_Buchung_SuchCombo: If GlAkt = False Then SSuFe
Case SY_BA_Banking_SuchCombo: If GlAkt = False Then SSuFe
Case SY_TL_Terminliste_SuchCombo: If GlAkt = False Then SSuFe
Case SY_TL_Terminliste_SortFeld: STeSo
Case SY_LB_Labor_SuchCombo: If GlAkt = False Then SSuFe
Case SY_LA_Auftrag_SuchCombo: If GlAkt = False Then SSuFe
Case KA_Eint_SuchCombo: If GlAkt = False Then SSuFe
Case KA_Kett_SuchCombo: If GlAkt = False Then SSuFe
Case KA_Frage_SuchCombo: If GlAkt = False Then SSuFe
Case TX_Mail_SucCombo: If GlAkt = False Then SSuFe
Case KA_Eint_KatCombo: KButt
Case KA_Kett_KatCombo: KButt
Case KA_Mail_KatCombo: MaKat
Case KA_Mail_SorCombo: SAbSo
Case KA_Mail_TexCombo: MaTex

Case Sta_Auswa: FDiSt
Case Sta_CmMan: SStaSt
Case Sta_ChCol: FDiTy 1
Case Sta_ChBar: FDiTy 2
Case Sta_ChPie: FDiTy 3
Case Sta_ChLin: FDiTy 4
Case Sta_ChAre: FDiTy 5
Case Sta_ChDon: FDiTy 6
Case Sta_Expor: FDiEx
Case Sta_ChAus: frmZeitraum.Show vbModal
Case Sta_ChDru: FDiPr
Case Sta_ChOp1: FDiOp 1
Case Sta_ChOp2: FDiOp 2
Case Sta_ChOp3: FDiOp 3
Case Sta_Capt1:
Case Sta_Capt2:

Case SY_SuTex: FSuch
Case SY_SuMan: FSuch True
Case SY_SuMit: FSuch
Case SY_SuRau: FSuch
Case SY_SuDat: FSuch
Case SY_SuWek: FSuch
Case SY_SuMon: FSuch
Case SY_SuJah: FSuJa
Case SY_SuBuh: FSuch
Case SY_SuBut: FKaSu
Case SY_SuAbg: FSuch
Case SY_SuSta: FSuch
Case SY_SuTSt: FSuch
Case SY_SuZug: FSuch

Case SY_TE_Termin_Fonts: STeFo
Case SY_TE_Termin_Spalte: STeSp
Case SY_TE_Termin_GlMZe: FTeAn 1
Case SY_TE_Termin_GlMFa: FTeAn 2
Case SY_TE_Termin_GlTGs: FTeAn 3
Case SY_TE_Termin_GlTeD: FTeAn 4
Case SY_TE_Termin_GlTVe: FTeAn 5
Case SY_TE_Termin_GlTSt: FTeAn 6
Case SY_TE_Termin_GlTDe: FTeAn 7
Case SY_TE_Termin_GlTTe: FTeAn 8
Case SY_TE_Termin_GlMiW: FTeAn 9
Case SY_TE_Termin_GlTZe: FTeAn 10
Case SY_TE_Termin_GlTKo: FTeAn 11
Case SY_TE_Termin_GlTKl: FTeAn 12
Case SY_TE_Termin_GlTeS: FTeAn 13
Case SY_TE_Termin_GlDeT: FTeAn 14
Case SY_TE_Termin_GlTrD: FTeAn 15
Case SY_AB_Spa_Multip: SAbSp 1
Case SY_AB_Spa_Zeierf: SAbSp 2
Case SY_AB_Spa_Mitarb: SAbSp 3
Case SY_AB_Spa_Einhei: SAbSp 4
Case SY_AB_Spa_Steuer: SAbSp 5
Case SY_AB_Spa_TabMod: SAbSp 6
Case SY_AB_Spa_Diagno: SAbSp 7
Case SY_AB_Spa_EigDia: SAbSp 8
Case SY_AB_Spa_Vorsch: SAbSp 9
Case SY_KB_Spa_Multip: SAbSp 1
Case SY_KB_Spa_Zeierf: SAbSp 2
Case SY_KB_Spa_Mitarb: SAbSp 3
Case SY_KB_Spa_Einhei: SAbSp 4
Case SY_KB_Spa_Steuer: SAbSp 5
Case SY_TP_Spa_Multip: SAbSp 1
Case SY_TP_Spa_Zeierf: SAbSp 2
Case SY_TP_Spa_Mitarb: SAbSp 3
Case SY_TP_Spa_Einhei: SAbSp 4
Case SY_TP_Spa_Steuer: SAbSp 5
Case SY_TP_Spa_Diagno: SAbSp 7
Case SY_AB_Dia_DatuZe: SAbSp 10
Case SY_AB_Dia_ICDZei: SAbSp 11
Case SY_KB_Spa_Ziffer: SAbSp 12
Case SY_AB_Kra_Antliz: SAbSp 13
Case SY_KB_Spa_PZNCod: SAbSp 14
Case SY_AB_Kra_Restri: SAbSp 15
Case SY_AB_Spa_Analog: SAbSp 16
Case SY_AB_Kra_AufMed: SAbSp 17
Case SY_KB_Spa_Betrag: SAbSp 18
Case SY_AB_Spa_StoRec: SAbSp 19
Case SY_KB_Kra_RecDet: SAbSp 20
Case SY_AB_Kra_Vorsch: SAbSp 21
Case SY_KB_Kra_DirBea: SAbSp 22
Case SY_AB_Spa_LaBetr: SAbSp 23
Case SY_AB_Kra_ZeiUmb: SAbSp 24
Case SY_AB_Zei_Toltip: SAbSp 25
Case SY_AB_Spa_Katalo: SAbSp 26
Case SY_KB_Kra_FliTex: SAbSp 27
Case SY_AB_Kra_AufDia: SAbSp 28
Case SY_KB_Kra_KraBla: SAbSp 29
Case SY_AB_Kra_Sorter: SAbSp 30
Case SY_AB_Spa_EinTyp: SAbSp 31
Case SY_AB_Kra_StoEin: SAbSp 32

Case SY_EI_Gruppierung: KGrLa
Case SY_AN_AnaBog_Sorti: SAbSo
Case SY_AN_AnaBog_Grupp: SAnGr
Case SY_AN_AnaBog_Expan: SAnEx
Case SY_AN_AnaBog_MarSel: S_AnSel 1, True
Case SY_AN_AnaBog_UnmSel: S_AnSel 1, False
Case SY_AN_AnaBog_MarAll: S_AnSel 2, True
Case SY_AN_AnaBog_UnmKei: S_AnSel 2, False

Case SY_EX_Kopieren: KDatei 9
Case SY_EX_Einfügen: KDatei 10
Case SY_EX_Ausschne: KDatei 8
Case SY_EX_Datei_Eig: KDatei 12
Case SY_EX_Datei_Bea: KDatei 11
Case SY_EX_Datei_Umb: KDatei 7
Case SY_EX_Datei_Del: KDatei 6
Case SY_EX_Datei_Ans: KDatei 21
Case SY_EX_Datei_Sor: KDatei 23
Case SY_EX_Datei_Thm: KDatei 25
Case SY_EX_Datei_Neu: KDatei 14
Case SY_EX_Datei_Fil: KDatei 28
Case SY_EX_Ordner_Neu: KDatei 2
Case SY_EX_Ordner_Umb: KDatei 3
Case SY_EX_Ordner_Del: KDatei 6
Case SY_EX_Ordner_Eig: KDatei 5
Case SY_EX_Ordner_Akt: KDatei 13
Case SY_EX_Ordner_Hom: KDatei 15
Case SY_EX_Ordner_Eml: KDatei 16
Case SY_EX_Ordner_Imp: KDatei 17
Case SY_EX_Ordner_Exp: KDatei 18
Case SY_EX_Ordner_Doc: KDatei 19
Case SY_EX_Ordner_Frm: KDatei 20
Case SY_EX_Datei_Impo: KDatei 27
Case SY_EX_Datei_Emai: KDatei 29

Case Tex_PaSuch: FSuFe
Case Tex_PaBear: SAdre 3
Case Tex_PaFilt: AdFMa
Case Tex_PaAlle: S_TxLis True
Case Tex_NePaFi: AdFMa
Case Tex_NePaAl: S_TxLis True
Case Tex_EdUndo: FTxCo TolId
Case Tex_EdRedo: FTxCo TolId
Case Tex_FntAu1: FTxFA TolId, CoTex
Case Tex_FntAu2: FTxFA TolId, CoTex
Case Tex_FntAu3: FTxFA TolId, CoTex
Case Tex_FntAu4: FTxFA TolId, CoTex
Case Tex_FntGr1: FTxFA TolId, CoTex
Case Tex_FntGr2: FTxFA TolId, CoTex
Case Tex_FntGr3: FTxFA TolId, CoTex
Case Tex_FntGr4: FTxFA TolId, CoTex
Case Tex_AusrLi: FTxFA TolId
Case Tex_AusrRe: FTxFA TolId
Case Tex_AusrZe: FTxFA TolId
Case Tex_AusrBl: FTxFA TolId
Case Tex_EinzLi: FTxFF TolId
Case Tex_EinzRe: FTxFF TolId
Case Tex_Zeiche: FTxCo TolId
Case Tex_Absatz: FTxCo TolId
Case Tex_Aufzah: FTxCo TolId
Case Tex_Numeri: FTxCo TolId
Case Tex_ForFet: FTxFF TolId
Case Tex_ForKur: FTxFF TolId
Case Tex_ForUnt: FTxFF TolId
Case Tex_ForDur: FTxFF TolId
Case Tex_TexMar: FTxFA TolId
Case Tex_FntKle: FTxFF TolId
Case Tex_FntGro: FTxFF TolId
Case Tex_FntHoh: FTxFF TolId
Case Tex_FntTif: FTxFF TolId
Case Tex_KopFus: FTxCo TolId
Case Tex_ForSpl: FTxCo TolId
Case Tex_KopZei: FTxFF TolId
Case Tex_FusZei: FTxFF TolId
Case Tex_TabEin: FTxCo TolId
Case Tex_TabAtr: FTxCo TolId
Case Tex_SpEiRe: FTxCo TolId
Case Tex_SpEiLi: FTxCo TolId
Case Tex_ZeEiUn: FTxCo TolId
Case Tex_ZeEiOb: FTxCo TolId
Case Tex_SpalLo: FTxCo TolId
Case Tex_ZeilLo: FTxCo TolId
Case Tex_TxUndo: FTxCo TolId
Case Tex_TxRedo: FTxCo TolId
Case Tex_ForSty: FTxCo TolId
Case Tex_ForVor: FTxCo TolId
Case Tex_TexRah: FObje TolId
Case Tex_EinGr1: FObje TolId
Case Tex_EinGr2: FObje TolId
Case Tex_EinGr3: FObje TolId
Case Tex_EinMar: FObje TolId
Case Tex_EinTex: FObje TolId
Case Tex_EinTab: FObje TolId
Case Tex_Tabell: 'Tabellenmenü
Case Tex_EinLnk: 'Hyperlink
Case Tex_EinObj: FObje TolId
Case Tex_Suchen: FTxCo TolId
Case Tex_Ersetz: FTxCo TolId
Case Tex_TexCut: FTxCo TolId
Case Tex_TexCop: FTxCo TolId
Case Tex_TexEin: FTxCo TolId
Case Tex_DatNeu: STxEi 0 'STxNe
Case Tex_DaNePa: STxEi 2
Case Tex_DaNeVe: STxEi 3
Case Tex_DaNeAr: STxEi 4
Case Tex_DaNePl: STxEi 5
Case Tex_DaNeRz: STxVo True
Case Tex_DaNeNe: STxVo True
Case Tex_DaNeN1: STxNe
Case Tex_DaNeN2: STxNw 1
Case Tex_DaNeN3: STxNw 2
Case Tex_DaNR01: STxRz 1
Case Tex_DaNR02: STxRz 2
Case Tex_DaNR03: STxRz 3
Case Tex_DaNR04: STxRz 4
Case Tex_DaNR05: STxRz 5
Case Tex_DaNR06: STxRz 6
Case Tex_DaNR07: STxRz 7
Case Tex_DaNR08: STxRz 8
Case Tex_DaNR09: STxRz 9
Case Tex_DaNR10: STxRz 10
Case Tex_DatSVo: STxVo True
Case Tex_DatVor: FRzBe
Case Tex_DaNeVo: STxVo
Case Tex_DaNeSe: STxSe False
Case Tex_DaNeAs: STxDi
Case Tex_DaFeAd: FTxDa CoTex
Case Tex_DaFeLo: FTxFF TolId
Case Tex_DaFeVe: FObje TolId
Case Tex_KopDat: FTxFF TolId
Case Tex_FusZal: FTxFF TolId
Case Tex_DatLoa: FObje TolId
Case Tex_DatSpe: FObje TolId
Case Tex_DatSpV: FObje TolId
Case Tex_DatSav: FObje TolId
Case Tex_DatLoe: FObje TolId
Case Tex_DatKop: FObje TolId
Case Tex_DocDru: FObje TolId
Case Tex_DocVor: FObje TolId
Case Tex_Eigens: FObje TolId
Case Tex_DocSig:
Case Tex_EtiDru: SDrLis 1
Case Tex_DocExp: STxMa True, False 'Dokument Exportieren
Case Tex_DocMa1: STxMa True, True 'Dokument als Emailanlage
Case Tex_DocMa2: STxMa False, True 'Dokument als Nachricht
Case Tex_DocSe1: STxMa True, False, 1 'Dokument Downloadlink
Case Tex_DocSe2: STxMa True, False, 2 'Dokument Digitalunterschrift
Case Tex_DocSe3: STxMa True, False, 3 'Dokument Neuaufnahmeformular
Case Tex_DigSig: FTxMa 2 'Dokument Digitalunterschrift
Case Tex_NeuAuf: FTxMa 3 'Dokument Neuaufnahmeformular
Case Tex_ZeiAb1: FTxFZ TolId
Case Tex_ZeiAb2: FTxFZ TolId
Case Tex_ZeiAb3: FTxFZ TolId
Case Tex_ZeiAb4: FTxFZ TolId
Case Tex_ZeiAb5: FTxFZ TolId
Case Tex_ZeiAb6: FTxFZ TolId
Case Tex_ZeiAb7: FTxFZ TolId
Case Tex_ClpEin: FObje TolId
Case Tex_ClpInh: FObje TolId
Case Tex_FaVor1: FTxCo TolId, ColID
Case Tex_FaVor2: FTxCo TolId 'RibTab_Tex_Dokumt
Case Tex_FaVor3: FTxCo TolId 'RibTab_Tex_Rezept
Case Tex_FaVor4: FTxCo TolId 'RibTab_Tex_NewsLe
Case Tex_FaVor5: FTxCo TolId 'RibTab_Krankenbla
Case Tex_FaVor6: FObje TolId
Case Tex_FaHin1: FTxCo TolId, ColID
Case Tex_FaHin2: FTxCo TolId 'RibTab_Tex_Dokumt
Case Tex_FaHin3: FTxCo TolId 'RibTab_Tex_Rezept
Case Tex_FaHin4: FTxCo TolId 'RibTab_Tex_NewsLe
Case Tex_FaHin5: FTxCo TolId 'RibTab_Krankenbla
Case Tex_FaHin6: FObje TolId
Case Tex_NweSen: FObje TolId
Case Tex_NweVor: FObje TolId
Case IC16_FarVor: FObje TolId
Case IC16_FarHin: FObje TolId
Case Tex_AdrSel: FSeAr False
Case Tex_AdrDes: FSeAr True
Case Tex_KraSpe: SKrTx True
Case Tex_KraKon: S_KrKn
Case Tex_KraDok: FTxKr

Case ShoCut_Start: STaSe ShoCut_Start, 0
Case ShoCut_Adresse: STaSe ShoCut_Adresse, 0
Case ShoCut_Kranken: STaSe ShoCut_Kranken, 0
Case ShoCut_Finanz: STaSe ShoCut_Finanz, 0
Case ShoCut_Termin: STaSe ShoCut_Termin, 0
Case ShoCut_Labor: STaSe ShoCut_Labor, 0
Case ShoCut_Texte: STaSe ShoCut_Texte, 0
Case ShoCut_Katalog: STaSe ShoCut_Katalog, 0
Case ShoCut_Abrechn: STaSe ShoCut_Abrechn, 0
Case ShoCut_Termin1: STaSe ShoCut_Termin1, 0
Case ShoCut_Termin2: STaSe ShoCut_Termin2, 0
Case ShoCut_Termin3: STaSe ShoCut_Termin3, 0

Case FaLei01: TeFarb 1, 3
Case FaLei02: TeFarb 2, 3
Case FaLei03: TeFarb 3, 3
Case FaLei04: TeFarb 4, 3
Case FaLei05: TeFarb 5, 3
Case FaLei06: TeFarb 6, 3
Case FaLei07: TeFarb 7, 3
Case FaLei08: TeFarb 8, 3
Case FaLei09: TeFarb 9, 3
Case FaLei10: TeFarb 10, 3
Case FaLei11: TeFarb 11, 3
Case FaLei12: TeFarb 12, 3
Case FaLei13: TeFarb 13, 3
Case FaLei14: TeFarb 14, 3
Case FaLei15: TeFarb 15, 3
Case FaLei16: TeFarb 16, 3
Case FaLei17: TeFarb 17, 3
Case FaLei18: TeFarb 18, 3
Case FaLei19: TeFarb 19, 3
Case FaLei20: TeFarb 20, 3

Case SuLei_1: FSuAu
Case SuLei_A: FSuLe "A", TolId
Case SuLei_B: FSuLe "B", TolId
Case SuLei_C: FSuLe "C", TolId
Case SuLei_D: FSuLe "D", TolId
Case SuLei_E: FSuLe "E", TolId
Case SuLei_F: FSuLe "F", TolId
Case SuLei_G: FSuLe "G", TolId
Case SuLei_H: FSuLe "H", TolId
Case SuLei_I: FSuLe "I", TolId
Case SuLei_J: FSuLe "J", TolId
Case SuLei_K: FSuLe "K", TolId
Case SuLei_L: FSuLe "L", TolId
Case SuLei_M: FSuLe "M", TolId
Case SuLei_N: FSuLe "N", TolId
Case SuLei_O: FSuLe "O", TolId
Case SuLei_P: FSuLe "P", TolId
Case SuLei_Q: FSuLe "Q", TolId
Case SuLei_R: FSuLe "R", TolId
Case SuLei_S: FSuLe "S", TolId
Case SuLei_T: FSuLe "T", TolId
Case SuLei_U: FSuLe "U", TolId
Case SuLei_V: FSuLe "V", TolId
Case SuLei_W: FSuLe "W", TolId
Case SuLei_X: FSuLe "X", TolId
Case SuLei_Y: FSuLe "Y", TolId
Case SuLei_Z: FSuLe "Z", TolId
Case SuLei_Ä: FSuLe "Ä", TolId
Case SuLei_Ö: FSuLe "Ö", TolId
Case SuLei_Ü: FSuLe "Ü", TolId

Case Else:
    If TolId > 0 Then
        If TolId < 1200 Then
            KrMain TolId - 1000
        Else
            FaNeu TolId - 1200
        End If
    Else
        SPaSe TolId
    End If
End Select

GlToo = False

End Sub
Private Sub FTrMa(Optional ByVal DriMa As Boolean, Optional ByVal DeMar As Boolean)
On Error GoTo SaErr
'Markiert oder Demarkiert die Krankenblatttypen in TreeView

Set FM = frmMain
Set TrLi5 = FM.trvList5

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

GlAkt = True

If DriMa = True Then
    If TrLi5.Nodes(1).Checked = False Then
        For Each Knote In TrLi5.Nodes
            If Knote.Index > 1 Then
                Knote.Checked = False
            End If
        Next Knote
    Else
        For Each Knote In TrLi5.Nodes
            If Knote.Index > 1 Then
                Knote.Checked = True
            End If
        Next Knote
    End If
Else
    GlAkt = True
    TrLi5.Nodes(1).Checked = False
    If DeMar = True Then
        For Each Knote In TrLi5.Nodes
            If Knote.Index > 1 Then
                Knote.Checked = False
            End If
        Next Knote
    Else
        For Each Knote In TrLi5.Nodes
            If Knote.Index > 1 Then
                Knote.Checked = True
            End If
        Next Knote
    End If
    GlAkt = False
End If

GlAkt = False

DoEvents
SSuFe

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTrMa " & Err.Number
Resume Next

End Sub
Private Sub FTxFA(ByVal TxFun As Integer, Optional ByVal TxStr As String)
On Error GoTo PoErr

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set TxCoN = FM.TexCont1
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Select Case TxFun
Case Tex_AusrLi: TxCoN.Alignment = 0
Case Tex_AusrRe: TxCoN.Alignment = 1
Case Tex_AusrZe: TxCoN.Alignment = 2
Case Tex_AusrBl: TxCoN.Alignment = 3
Case Tex_FntAu1: TxCoN.FontName = TxStr
Case Tex_FntAu2: TxCoN.FontName = TxStr
Case Tex_FntAu3: TxCoN.FontName = TxStr
Case Tex_FntAu4: TxCoN.FontName = TxStr
Case Tex_FntGr1: TxCoN.FontSize = CLng(TxStr)
Case Tex_FntGr2: TxCoN.FontSize = CLng(TxStr)
Case Tex_FntGr3: TxCoN.FontSize = CLng(TxStr)
Case Tex_FntGr4: TxCoN.FontSize = CLng(TxStr)
Case Tex_TexMar:
    If TxCoN.ControlChars = False Then
        TxCoN.ControlChars = True
        CmAcs(Tex_TexMar).Checked = True
    Else
        TxCoN.ControlChars = False
        CmAcs(Tex_TexMar).Checked = False
    End If
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing

STxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxFA " & Err.Number
Resume Next

End Sub
Private Sub FTxCo(ByVal TxFun As Integer, Optional ByVal ColID As Long)
On Error GoTo PoErr

Dim ImVer As String
Dim AktZa As Integer

Set FM = frmMain
Set TxCoN = FM.TexCont1

Select Case TxFun
Case Tex_TexCut: TxCoN.Clip 1
Case Tex_TexCop: TxCoN.Clip 2
Case Tex_TexEin: If TxCoN.CanPaste = True Then TxCoN.Paste 5
Case Tex_Suchen: TxCoN.FindReplace 1
Case Tex_Ersetz: TxCoN.FindReplace 2
Case Tex_TxUndo: TxCoN.Undo
Case Tex_TxRedo: TxCoN.Redo
Case Tex_Zeiche: TxCoN.FontDialog
Case Tex_Absatz: TxCoN.ParagraphDialog
Case Tex_Aufzah: TxCoN.ListAttrDialog
Case Tex_Numeri: TxCoN.ListAttrDialog
Case Tex_TabEin: If TxCoN.TableCanInsert = True Then TxCoN.TableInsertDialog
Case Tex_TabAtr: If TxCoN.TableCanChangeAttr = True Then TxCoN.TableAttrDialog
Case Tex_SpEiRe: If TxCoN.TableCanInsertColumn = True Then TxCoN.TableInsertColumn txTableInsertAfter
Case Tex_SpEiLi: If TxCoN.TableCanInsertColumn = True Then TxCoN.TableInsertColumn txTableInsertInFront
Case Tex_ZeEiUn: If TxCoN.TableCanInsertLines = True Then TxCoN.TableInsertLines txTableInsertAfter, 1
Case Tex_ZeEiOb: If TxCoN.TableCanInsertLines = True Then TxCoN.TableInsertLines txTableInsertInFront, 1
Case Tex_SpalLo: If TxCoN.TableCanDeleteColumn = True Then TxCoN.TableDeleteColumn
Case Tex_ZeilLo: If TxCoN.TableCanDeleteLines = True Then TxCoN.TableDeleteLines
Case Tex_FaVor1: TxCoN.ForeColor = ColID
Case Tex_FaVor2: TxCoN.ForeColor = vbBlack 'RibTab_Tex_Dokumt
Case Tex_FaVor3: TxCoN.ForeColor = vbBlack 'RibTab_Tex_Rezept
Case Tex_FaVor4: TxCoN.ForeColor = vbBlack 'RibTab_Tex_NewsLe
Case Tex_FaVor5: TxCoN.ForeColor = vbBlack 'RibTab_Krankenbla
Case Tex_FaHin1: TxCoN.TextBkColor = ColID
Case Tex_FaHin2: TxCoN.TextBkColor = vbWhite 'RibTab_Tex_Dokumt
Case Tex_FaHin3: TxCoN.TextBkColor = vbWhite 'RibTab_Tex_Rezept
Case Tex_FaHin4: TxCoN.TextBkColor = vbWhite 'RibTab_Tex_NewsLe
Case Tex_FaHin5: TxCoN.TextBkColor = vbWhite 'RibTab_Krankenbla
Case Tex_ForVor: TxCoN.SectionFormatDialog 0
Case Tex_KopFus: TxCoN.SectionFormatDialog 1
Case Tex_ForSpl: TxCoN.SectionFormatDialog 2
Case Tex_ForSty: TxCoN.StyleDialog
Case Tex_EdUndo: TxCoN.Undo
Case Tex_EdRedo: TxCoN.Redo
End Select

STxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxCO " & Err.Number
Resume Next

End Sub
Private Sub FTxDa(ByVal CoTex As String)
On Error GoTo PoErr

Dim FlIdx As Long
Dim FeSta As Long
Dim FeEnd As Long
Dim FlDat As String
Dim AktZa As Integer

Set FM = frmMain
Set TxCoN = FM.TexCont1

For AktZa = 1 To UBound(GlSer)
    If CoTex = GlSer(AktZa, 0) Then
        FlDat = GlSer(AktZa, 1)
        Exit For
    End If
Next AktZa

With TxCoN
    .FieldInsert "<" & CoTex & ">"
    FlIdx = .FieldCurrent
    If FlIdx > 0 Then
        .FieldType(FlIdx) = txFieldStandard
        .FieldData(FlIdx) = FlDat
        .FieldEditAttr(FlIdx) = &H2 + &H10
        .FieldChangeable = False
        .FieldDeleteable = True
        FeSta = .FieldStart
        FeEnd = .FieldEnd
        .SelStart = FeEnd
        .SelLength = 1
        .SelText = Space$(1)
    End If
End With

STxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxDa " & Err.Number
Resume Next

End Sub
Private Sub FTxFF(ByVal TxFun As Integer)
On Error GoTo PoErr

Dim FeIdx As Long
Dim Lange As Long
Dim FeSta As Long
Dim FeEnd As Long
Dim TxGro As Integer
Dim TxIde As Integer
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim TxFnt As New StdFont
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set TxCoN = FM.TexCont1
Set CoDia = FM.comDialo
Set CmAcs = CmBrs.Actions

Tit1 = "Datenfeld Entfernen"
Mld1 = "Möchten Sie das markierte Datenfeld wirklich löschen?"

Select Case TxFun
Case Tex_ForFet:
    If TxCoN.FontBold = 0 Then
        TxCoN.FontBold = 1
    Else
        TxCoN.FontBold = 0
    End If
Case Tex_ForKur:
    If TxCoN.FontItalic = 0 Then
        TxCoN.FontItalic = 1
    Else
        TxCoN.FontItalic = 0
    End If
Case Tex_ForUnt:
    If TxCoN.FontUnderline = 0 Then
        TxCoN.FontUnderline = 1
    Else
        TxCoN.FontUnderline = 0
    End If
Case Tex_ForDur:
    If TxCoN.FontStrikethru = 0 Then
        TxCoN.FontStrikethru = 1
    Else
        TxCoN.FontStrikethru = 0
    End If
Case Tex_FntHoh:
    If TxCoN.BaseLine = 0 Then
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = 100
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 3) * 2)
        Else
            TxCoN.FontSize = (CInt(TxGro / 3) * 2) - 1
        End If
    Else
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = 0
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 2) * 3)
        Else
            TxCoN.FontSize = (CInt(TxGro / 2) * 3) - 1
        End If
    End If
Case Tex_FntTif:
    If TxCoN.BaseLine = 0 Then
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = -100
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 3) * 2)
        Else
            TxCoN.FontSize = (CInt(TxGro / 3) * 2) - 1
        End If
    Else
        TxGro = TxCoN.FontSize
        TxCoN.BaseLine = 0
        If TxGro Mod 2 = 0 Then
            TxCoN.FontSize = (CInt(TxGro / 2) * 3)
        Else
            TxCoN.FontSize = (CInt(TxGro / 2) * 3) - 1
        End If
    End If
Case Tex_FntKle:
    TxGro = TxCoN.FontSize
    TxCoN.FontSize = TxGro - 2
Case Tex_FntGro:
    TxGro = TxCoN.FontSize
    TxCoN.FontSize = TxGro + 2
Case Tex_EinzLi:
    TxIde = TxCoN.IndentL
    TxCoN.IndentL = TxIde + 400
Case Tex_EinzRe:
    TxIde = TxCoN.IndentL
    TxCoN.IndentL = TxIde - 400
Case Tex_KopZei:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstHeader Then
            .HeaderFooterActivate txFirstHeader
            GlHeA = txFirstHeader
        Else
            .HeaderFooterActivate txHeader
            GlHeA = txHeader
        End If
        .HeaderFooterSelect txHeader
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = TxFnt.SIZE
        .SelText = vbNullString
        .HeaderFooterSelect 0
    End With
Case Tex_FusZei:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstFooter Then
            .HeaderFooterActivate txFirstFooter
            GlHeA = txFirstFooter
        Else
            .HeaderFooterActivate txFooter
            GlHeA = txFooter
        End If
        .HeaderFooterSelect txFooter
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = TxFnt.SIZE
        .SelText = vbNullString
        .HeaderFooterSelect 0
    End With
Case Tex_KopDat:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        .HeaderFooterActivate txMainText
        GlHeA = 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstHeader Then
            .HeaderFooterActivate txFirstHeader
            GlHeA = txFirstHeader
        Else
            .HeaderFooterActivate txHeader
            GlHeA = txHeader
        End If
        .HeaderFooterSelect txHeader
        .Alignment = 1
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = TxFnt.SIZE
        .FieldInsert "<" & GlSer(74, 0) & ">" 'Tagesdatum
        FeIdx = .FieldCurrent
        .FieldType(FeIdx) = txFieldStandard
        .FieldData(FeIdx) = GlSer(74, 1)
        .FieldEditAttr(FeIdx) = &H2 + &H10
        FeSta = .FieldStart
        FeEnd = .FieldEnd
        .SelStart = FeEnd
        .SelLength = 1
        .SelText = Space$(1)
        .FieldChangeable = False
        .FieldDeleteable = True
        .HeaderFooterSelect 0
    End With
    CmAcs(Tex_KopDat).Enabled = False
Case Tex_FusZal:
    TxFnt.Name = GlXFt.Name
    TxFnt.SIZE = GlXFt.SIZE
    With TxCoN
        .TextFrameSelect 0
        .HeaderFooterActivate txMainText
        GlHeA = 0
        If .HeaderFooter = 0 Then
            .HeaderFooter = txHeader + txFooter
        End If
        If .HeaderFooter And txFirstFooter Then
            .HeaderFooterActivate txFirstFooter
            GlHeA = txFirstFooter
        Else
            .HeaderFooterActivate txFooter
            GlHeA = txFooter
        End If
        .HeaderFooterSelect txFooter
        .Alignment = 1
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        .FontName = TxFnt.Name
        .FontSize = 8
        .SelText = "Seite: "
        .SelLength = 0
        .SelStart = 7
        .FieldInsert vbNullString
        FeIdx = .FieldCurrent
        .FieldType(FeIdx) = txFieldPageNumber
        .FieldEditAttr(FeIdx) = &H2 + &H10
        FeSta = .FieldStart
        FeEnd = .FieldEnd
        .SelStart = FeEnd
        .SelLength = 1
        .SelText = Space$(1)
        .FieldChangeable = False
        .FieldDeleteable = True
        .HeaderFooterSelect 0
        .PageNumberDialog FeIdx
    End With
    CmAcs(Tex_FusZal).Enabled = False
Case Tex_DaFeLo:
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
        FeIdx = TxCoN.FieldAtInputPos
        If FeIdx > 0 Then
            TxCoN.FieldCurrent = FeIdx
            TxCoN.FieldDelete True
        End If
    End If
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing
Set CoDia = Nothing

STxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxFF " & Err.Number
Resume Next

End Sub
Private Sub FTxFZ(ByVal TxFun As Integer)
On Error GoTo PoErr
'Zeilenabstand

Set FM = frmMain
Set TxCoN = FM.TexCont1

Select Case TxFun
Case Tex_ZeiAb1: TxCoN.LineSpacing = 100
Case Tex_ZeiAb2: TxCoN.LineSpacing = 120
Case Tex_ZeiAb3: TxCoN.LineSpacing = 130
Case Tex_ZeiAb4: TxCoN.LineSpacing = 150
Case Tex_ZeiAb5: TxCoN.LineSpacing = 200
Case Tex_ZeiAb6: TxCoN.LineSpacing = 250
Case Tex_ZeiAb7: TxCoN.LineSpacing = 300
End Select

STxFo

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxFZ " & Err.Number
Resume Next

End Sub
Private Sub FTxKr()
On Error GoTo PoErr
'Texcoontrol Neu

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = "Krankenblattdokument zurücksetzen"
TeMai = "Soll das Krankenblattdokument jetzt zurückgesetzt werden?"
TeInh = "Beim Zurücksetzen des Krankenblattdokumentes werden alle bisher eingefügten oder übernommenen Inhalte entfernt, danach kann eine neue Übernahme der Krankenblatttabelle durchgeführt werden."
TeFus = "Beim Zurücksetzen des Krankenblattdokumentes wird diese Änderung aus Sicherheitsgründen nicht automatisch gespeichert, damit diese Änderung falls notwendig wieder rückgängig gemacht werden kann. "

Set FM = frmMain
Set TxCoN = FM.TexCont1

SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
If GlMes = 33565 Then
    STxNe
    GlTSV = True 'Speichern Textverarbeitung
End If

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxRu " & Err.Number
Resume Next

End Sub
Private Sub FTxMa(ByVal AuTyp As Integer)
On Error GoTo PoErr

STxVo
DoEvents

Select Case AuTyp
Case 2: STxMa True, False, 2 'Dokument Digitalunterschrift
Case 3: STxMa True, False, 3 'Dokument Neuaufnahmeformular
End Select

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxMa " & Err.Number
Resume Next

End Sub
Private Sub FTxNe()
On Error GoTo PoErr

Dim Frage As Integer
Dim Mld1, Tit1 As String

Set FM = frmMain

Tit1 = "Neues Dokument"
Mld1 = "Soll ein neues Dokument hinzugefügt werden?"

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    If GlTSV = True Then 'Speichern Textverarbeitung
        STxSa
    End If
    STxVo True
End If

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxNe " & Err.Number
Resume Next

End Sub

Private Sub FTxPh(ByVal TxStr As String)
On Error GoTo AnErr
'Textphrasensuche

Dim AktZa As Integer
Dim AkZei As Integer

Set FM = frmMain
Set TxCoN = FM.TexCont1

If TxStr <> vbNullString Then
    If Len(TxPhr) < 8 Then
        TxPhr = TxPhr & TxStr
        For AktZa = 1 To UBound(GlTxP) 'Textphrasen
            If LCase(TxPhr) = LCase(GlTxP(AktZa, 1)) Then
                AkZei = TxCoN.SelStart

                If AkZei - (Len(TxPhr) + 1) < 0 Then
                    TxCoN.SelStart = 0
                Else
                    TxCoN.SelStart = AkZei - (Len(TxPhr) + 1)
                End If
                TxCoN.SelLength = Len(TxPhr)
                'TxCoN.SelText = GlTxP(AktZa, 2)
                TxPhr = vbNullString
                Exit For
            End If
            AktZa = AktZa + 1
        Next AktZa
    Else
        TxPhr = vbNullString
    End If
End If

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxPh " & Err.Number
Resume Next

End Sub
Private Sub FTxRu(Optional ByVal xReDo As Boolean = False)
On Error GoTo PoErr
'Texcoontrol Undo / redo

Set FM = frmMain
Set TxCoN = FM.TexCont1

If xReDo = True Then
    Select Case GlBut
    Case RibTab_Krankenbla: TxCoN.Redo
    Case RibTab_Tex_Dokumt: TxCoN.Redo
    Case RibTab_Tex_Vorlag: TxCoN.Redo
    Case RibTab_Tex_Rezept: TxCoN.Redo
    Case RibTab_Tex_NewsLe: TxCoN.Redo
    End Select
Else
    Select Case GlBut
    Case RibTab_Krankenbla: TxCoN.Undo
    Case RibTab_Tex_Dokumt: TxCoN.Undo
    Case RibTab_Tex_Vorlag: TxCoN.Undo
    Case RibTab_Tex_Rezept: TxCoN.Undo
    Case RibTab_Tex_NewsLe: TxCoN.Undo
    End Select
End If

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxRu " & Err.Number
Resume Next

End Sub
Private Sub FTxTb()
On Error GoTo PoErr
'Tabulator Click Textverarbeitung

Dim RetWe As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set LiVw4 = FM.lstView4
Set CmBrs = FM.comBar01
Set TabCo = FM.TabCont1
Set TxCoN = FM.TexCont1
Set CmAcs = CmBrs.Actions
Set LiIts = LiVw4.ListItems

Select Case TabCo.Selected.Index
Case 0:
    CmAcs(Tex_PaSuch).Visible = True
    CmAcs(Tex_PaBear).Visible = True
    CmAcs(Tex_PaFilt).Visible = False
    CmAcs(Tex_PaAlle).Visible = False
    CmAcs(Tex_DatSpe).Enabled = True
    CmAcs(Tex_DaNePa).Enabled = True
    CmAcs(Tex_DaNeAr).Enabled = True
    CmAcs(Tex_DaNeVe).Enabled = True
    CmAcs(Tex_DaNeAs).Enabled = True
    CmAcs(Tex_DaNeVo).Enabled = True
    CmAcs(Tex_DaNePl).Enabled = True
    CmAcs(Tex_DocVor).Enabled = True
    CmAcs(Tex_DocExp).Enabled = True
    CmAcs(Tex_DocMa1).Enabled = True
    CmAcs(Tex_DocMa2).Enabled = True
    CmAcs(Tex_DocSe1).Enabled = True
    CmAcs(Tex_DocSe2).Enabled = True
    CmAcs(Tex_DocSe3).Enabled = True
    CmAcs(Tex_EtiDru).Enabled = True
    CmAcs(Tex_Eigens).Enabled = True
Case 1:
    CmAcs(Tex_PaSuch).Visible = False
    CmAcs(Tex_PaBear).Visible = False
    CmAcs(Tex_PaFilt).Visible = True
    CmAcs(Tex_PaAlle).Visible = True
    CmAcs(Tex_DatSpe).Enabled = False
    CmAcs(Tex_DaNePa).Enabled = False
    CmAcs(Tex_DaNeAr).Enabled = False
    CmAcs(Tex_DaNeVe).Enabled = False
    CmAcs(Tex_DaNeAs).Enabled = False
    CmAcs(Tex_DaNePl).Enabled = False
    CmAcs(Tex_DaNeVo).Enabled = True
    CmAcs(Tex_DocVor).Enabled = False
    CmAcs(Tex_DocExp).Enabled = False
    CmAcs(Tex_DocMa1).Enabled = False
    CmAcs(Tex_DocMa2).Enabled = False
    CmAcs(Tex_DocSe1).Enabled = False
    CmAcs(Tex_DocSe2).Enabled = False
    CmAcs(Tex_DocSe3).Enabled = False
    CmAcs(Tex_EtiDru).Enabled = False
    CmAcs(Tex_Eigens).Enabled = False
End Select

If TabCo.Selected.Index = 0 Then
    GlTxM = False 'Serienbriefmodus
Else
    GlTxM = True
End If
DoEvents

If TabCo.Selected.Index = 1 Then
    DoEvents
    S_TxLis
Else
    If LiIts.Count > 0 Then
        STxLa
    End If
End If

GlTSV = False 'Speichern Textverarbeitung

Set TabCo = Nothing
Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxTb " & Err.Number
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If GlRes = False Then 'Reset der Einstellungen
    FClos
Else
    Cancel = False
End If

For Each FM In VB.Forms
    If FM.Name <> "frmMain" Then
        Unload FM
        Set FM = Nothing
    End If
Next FM

TimEnde 1
TimEnde 2
TimEnde 3
TimEnde 4
TimEnde 5
TimEnde 6
TimEnde 7
DoEvents

WindowEnSu Me.hwnd
DoEvents

WindowClose
DoEvents

Unload frmMain

End Sub

Private Sub lstView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If GlAkt = False Then SGrUm 4, NewString
End Sub
Private Sub lstView1_DblClick()
    If GlAkt = False Then SGrUm 3
End Sub
Private Sub lstView1_GotFocus()
On Error Resume Next

Dim DocPa As XtremeDockingPane.DockingPane

Set FM = frmMain
Set DocPa = FM.dcpDoc01

DocPa.FindPane(PA_DP_Top1).Selected = True

End Sub

Private Sub lstView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
        Case vbKeyF12: SGrUm 3
        Case vbKeyDelete: Dia_Lo
        Case vbKeyBack: Dia_Lo
        Case 119:
        Case 17:
        End Select
    End If
End Sub
Private Sub lstView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo 3
        End If
    End If
End Sub
Private Sub lstView1_OLEDragDrop(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    FEinf 1
End Sub
Private Sub lstView2_GotFocus()
    GlKrT = 1
End Sub
Private Sub lstView2_OLEDragDrop(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    FEinf 2
End Sub
Private Sub lstView3_AfterLabelEdit(Cancel As Integer, NewString As String)
    If GlAkt = False Then SGrUm 4, NewString
End Sub
Private Sub lstView3_DblClick()
    If GlAkt = False Then SGrUm 3
End Sub
Private Sub lstView3_GotFocus()
    GlKrT = 2
End Sub
Private Sub lstView3_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
        Case vbKeyF12: SGrUm 3
        Case vbKeyDelete: Dia_Lo
        Case vbKeyBack: Dia_Lo
        Case 119:
        Case 17:
        End Select
    End If
End Sub
Private Sub lstView3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo 3
        End If
    End If
End Sub
Private Sub lstView3_OLEDragDrop(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    FEinf 3
End Sub

Private Sub lstView4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Dim Frage As Integer
Dim Mld1, Tit1 As String

Set LiVw4 = Me.lstView4
    
Tit1 = "Dokument Speichern"
Mld1 = "Soll das aktuelle Dokument gespeichert werden?"

If GlAkt = False Then
    If LiVw4.ListItems.Count > 0 Then
        If Button = vbRightButton Then
            SMePo 1
        Else
            If Not LiVw4.HitTest(x, y) Is Nothing Then
                If LiVw4.HitTest(x, y).Selected = True Then
                    If GlTSV = True Then 'Speichern Textverarbeitung
                        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
                        If Frage = 6 Then
                            STxSa
                            DoEvents
                        Else
                            GlTSV = False
                        End If
                    End If
                    STxLa LiVw4.HitTest(x, y).Tag
                End If
            End If
        End If
    End If
End If
    
End Sub

Private Sub optGrup1_Click()
    If GlAkt = False Then
        STrSa 1
    End If
End Sub

Private Sub optGrup2_Click()
    If GlAkt = False Then
        STrSa 2
    End If
End Sub


Private Sub optGrup3_Click()
    If GlAkt = False Then
        STrSa 3
    End If
End Sub


Private Sub optGrup4_Click()
    If GlAkt = False Then
        STrSa 4
    End If
End Sub


Private Sub optGrup5_Click()
    If GlAkt = False Then
        STrSa 5
    End If
End Sub


Private Sub optGrup6_Click()
    If GlAkt = False Then
        STrSa 6
    End If
End Sub


Private Sub optLaGr1_Click()
    If GlAkt = False Then
        STrSa 1
    End If
End Sub

Private Sub optLaGr2_Click()
    If GlAkt = False Then
        STrSa 2
    End If
End Sub

Private Sub optLaGr3_Click()
    If GlAkt = False Then
        STrSa 3
    End If
End Sub

Private Sub optLaGr4_Click()
    If GlAkt = False Then
        STrSa 4
    End If
End Sub


Private Sub optLaGr5_Click()
    If GlAkt = False Then STrSa 5
End Sub


Private Sub optLaGr6_Click()
    If GlAkt = False Then STrSa 6
End Sub


Private Sub optTeGr1_Click()
    If GlAkt = False Then STrSa 1
End Sub

Private Sub optTeGr2_Click()
    If GlAkt = False Then STrSa 2
End Sub


Private Sub optTeGr3_Click()
    If GlAkt = False Then STrSa 3
End Sub


Private Sub optTeGr4_Click()
    If GlAkt = False Then STrSa 4
End Sub


Private Sub optTeGr5_Click()
    If GlAkt = False Then STrSa 5
End Sub


Private Sub optTeGr6_Click()
    If GlAkt = False Then STrSa 6
End Sub


Private Sub optZeit1_Click()
    FDiZe
End Sub

Private Sub optZeit2_Click()
    FDiZe
End Sub

Private Sub optZeit3_Click()
    FDiZe
End Sub

Private Sub optZeit4_Click()
    FDiZe
End Sub

Private Sub picBild5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo
        End If
    End If
End Sub

Private Sub picRah14_Resize()
On Error Resume Next

Set FM = frmMain
Set PiR14 = FM.picRah14
Set Labl9 = FM.lblDeta9

Labl9.Move 200, 100, PiR14.Width - 400, PiR14.Height - 200

End Sub


Private Sub popCont1_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.id = 2 Then popCont1.Close
End Sub

Private Sub popCont2_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.id = 2 Then popCont2.Close
End Sub

Private Sub popCont3_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.id = 2 Then popCont3.Close
End Sub
Private Sub prpGrid1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = vbRightButton Then
        SMePo 2
    End If
End Sub

Private Sub prpGrid1_ValueChanged(ByVal Item As XtremePropertyGrid.IPropertyGridItem)
On Error Resume Next

Dim TmFar As Long    'Defaultfarbwert
Dim TmBoD As Boolean 'Deufault Boolean Wert
Dim TmBoV As Boolean 'Value Boolean Wert
Dim AryIt() As String

Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrSub As XtremePropertyGrid.PropertyGridItem
Dim PrIts As XtremePropertyGrid.PropertyGridItems
Dim PrBol As XtremePropertyGrid.PropertyGridItemBool
Dim PrDat As XtremePropertyGrid.PropertyGridItemDate
Dim PrFnt As XtremePropertyGrid.PropertyGridItemFont
Dim PrOpt As XtremePropertyGrid.PropertyGridItemOption

Set PrGr1 = Me.prpGrid1
Set PrIts = PrGr1.Categories

For Each PrKat In PrIts
    For Each PrItm In PrKat.Childs
        If PrItm.Selected = True Then
            Select Case PrItm.Type
            Case PropertyItemCategory:
            
            Case PropertyItemString:
                If PrItm.defaultValue <> PrItm.Value Then
                    S_AnSav PrItm.Value, PrItm.id
                End If
            Case PropertyItemNumber:
                If PrItm.defaultValue <> PrItm.Value Then
                    S_AnSav PrItm.Value, PrItm.id
                End If
            Case PropertyItemBool:
                Set PrBol = PrItm
                TmBoD = PrBol.defaultValue
                TmBoV = PrBol.Value
                If TmBoD <> TmBoV Then
                    S_AnSav PrItm.Value, PrItm.id
                End If
                For Each PrSub In PrKat.Childs
                    If PrBol.Tag = PrSub.Tag Then
                        If CBool(PrBol.Value) = True Then
                            PrSub.Hidden = False
                        Else
                            PrSub.Hidden = True
                            PrBol.Hidden = False
                        End If
                    End If
                Next PrSub
            Case PropertyItemOption:
                Set PrOpt = PrItm
                If PrOpt.defaultValue <> PrOpt.Value Then
                    S_AnSav PrOpt.Value, PrOpt.id
                End If
            Case PropertyItemColor:
                AryIt = Split(PrItm.defaultValue, ";")
                TmFar = RGB(AryIt(0), AryIt(1), AryIt(2))
                If TmFar <> PrItm.Value Then
                    S_AnSav PrItm.Value, PrItm.id
                End If
            Case PropertyItemFont:
                Set PrFnt = PrItm
                'If NeFn1 = True Then FRgSv PrFnt.id
                'If NeFn2 = True Then FRgSv PrFnt.id
                'If NeFn3 = True Then FRgSv PrFnt.id
            Case PropertyItemDate:
                If PrItm.defaultValue <> PrItm.Value Then
                    S_AnSav PrItm.Value, PrItm.id
                End If
            End Select
        End If
    Next PrItm
Next PrKat

End Sub
Private Sub repCont0_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If GlSta = False Then
    If GlEmV > 0 Then 'Anzahl E-Mails
        If GlEmV > Row.Index Then
            If MaAry(Item.Index, Row.Index) <> vbNullString Then
                Metrics.Text = MaAry(Item.Index, Row.Index)
                If MaAry(Mai_Gelesen, Row.Index) <> vbNullString Then
                    If CBool(MaAry(Mai_Gelesen, Row.Index)) = False Then
                        Metrics.Font.Bold = True
                    End If
                Else
                    Metrics.Font.Bold = True
                End If
                If MaAry(Mai_Marker, Row.Index) <> vbNullString Then
                    If CBool(MaAry(Mai_Marker, Row.Index)) = True Then
                        Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
                    End If
                End If
                If GlMKa = 1 Then 'Mailkatalog (1=Posteingang 2=Postaisgang)
                    If CBool(MaAry(Mai_Spammail, Row.Index)) = True Then
                        Metrics.ForeColor = 8421504
                        Metrics.Font.Strikethrough = True
                    End If
                End If
                If GlMKa = 2 Then 'Mailkatalog (1=Posteingang 2=Postaisgang)
                    If Left$(MaAry(Mai_Mailsize, Row.Index), 4) = "1 KB" Then
                        Metrics.ForeColor = 8421504
                        Metrics.Font.Strikethrough = True
                    End If
                End If
                If Item.Index = Mai_Priority Then
                    Metrics.Text = vbNullString
                    If Val(MaAry(Mai_Priority, Row.Index)) = 1 Then 'High
                        Metrics.ItemIcon = IC16_Sign_Info
                    ElseIf Val(MaAry(Mai_Priority, Row.Index)) = 5 Then 'Low
                        Metrics.ItemIcon = IC16_Sign_Check
                        Metrics.ForeColor = vbGreen
                    End If
                End If
                If Item.Index = Mai_Attachment Then
                    Metrics.Text = vbNullString
                    If CBool(MaAry(Mai_Attachment, Row.Index)) = True Then
                        Metrics.ItemIcon = IC16_Paperclip
                    End If
                End If
                If Item.Index = Mai_Marker Then
                    Metrics.Text = vbNullString
                    If CBool(MaAry(Mai_Marker, Row.Index)) = True Then
                        Metrics.ItemIcon = IC16_Pin_Norm
                    ElseIf MaAry(Mai_ID0, Row.Index) = 0 Then
                        Metrics.ItemIcon = IC16_Sign_Help
                    End If
                End If
                If Item.Index = Mai_SenderName Then
                    If MaAry(Mai_Gelesen, Row.Index) <> vbNullString Then
                        If CBool(MaAry(Mai_Gelesen, Row.Index)) = True Then
                            If MaAry(Mai_Sensitivity, Row.Index) <> vbNullString Then
                                If Val(MaAry(Mai_Sensitivity, Row.Index)) = 1 Then
                                    Metrics.ItemIcon = IC16_Mail_Read
                                ElseIf Val(MaAry(Mai_Sensitivity, Row.Index)) = 2 Then
                                    Metrics.ItemIcon = IC16_Mail_Import
                                End If
                            Else
                                Metrics.ItemIcon = IC16_Mail_Open
                            End If
                        Else
                            Metrics.ItemIcon = IC16_Mail_Close
                        End If
                    Else
                        Metrics.ItemIcon = IC16_Mail_Close
                    End If
                End If
            End If
        End If
    End If
End If

End Sub
Private Sub repCont0_ColumnOrderChangedEx(ByVal Column As XtremeReportControl.IReportColumn, ByVal Reason As XtremeReportControl.XTPReportColumnOrderChangedReason)
    If GlAkt = False Then SSpSav
End Sub


Private Sub repCont0_ColumnWidthChanged(ByVal Column As XtremeReportControl.IReportColumn, ByVal PrevWidth As Long, ByVal NewWidth As Long)
    If GlAkt = False Then SSpSav
End Sub
Private Sub repCont0_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            If GlAkt = False Then
                
                
            End If
        End If
    End If
End Sub
Private Sub repCont0_KeyUp(KeyCode As Integer, Shift As Integer)

Dim RpCo0 As XtremeReportControl.ReportControl

Set RpCo0 = Me.repCont0

If GlAkt = False Then
    If RpCo0.Records.Count > 0 Then
        Set RpSel = RpCo0.SelectedRows
        If RpSel.Count > 0 Then
            If Shift = 0 Then
                If KeyCode >= 65 And KeyCode <= 90 Then
                    With GlSuI
                        .SuIdx = 3
                        .SuStr = Chr$(KeyCode)
                    End With
                    SSuch
                Else
                    SMark
                End If
            End If
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo0 = Nothing

End Sub
Private Sub repCont0_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo LaErr

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo0 = Me.repCont0
Set RpRws = RpCo0.Rows
Set HiTes = RpCo0.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

If RpRws.Count > 0 Then
    Select Case HiTes.ht
    Case xtpHitTestGroupBox:
    Case xtpHitTestHeader:
    Case xtpHitTestReportArea:
            SMark
            If Button = vbRightButton Then
                SMePo 2
            Else
                MaDet
            End If
    Case xtpHitTestUnknown:
    End Select
End If

Set RpRws = Nothing
Set RpCo0 = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub repCont0_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        If Row.GroupRow = False Then
            If WindowLoad("frmMaiView") = True Then
                Unload frmMaiView
                DoEvents
            End If
            GlMaY = 0 'Emailflyoutfenster Mailindex
            GlNaT = 1 'Mailtyp (1=View 2=Neu 3=Antwort)
            MaMain 'MaAry(Mai_IDA, Row.Index)
        End If
    End If
End Sub
Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long
Dim FarRo As Long
Dim FarGr As Long
Dim FarBl As Long

Dim AktZa As Integer
Dim FaRGB As tRGB

If GlSta = False Then
    If Row.GroupRow = False Then
        Select Case GlBut
        Case RibTab_Mahnwesen:
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
        Case RibTab_Buchungen:
            If Row.Record.ItemCount > 27 Then
                For AktZa = 1 To UBound(GlGeK)
                    If GlGeK(AktZa, 0) = Row.Record(Buh_IDB).Value Then
                        If CBool(GlGeK(AktZa, 5)) = True Then
                            Metrics.ForeColor = 16711680
                        End If
                        Exit For
                    End If
                Next AktZa
                If IsNull(Row.Record(Buh_Storniert).Value) = False Then
                    If Row.Record(Buh_Storniert).Value <> vbNullString Then
                        If CBool(Row.Record(Buh_Storniert).Value) = True Then
                            Metrics.Font.Strikethrough = True
                            Metrics.ForeColor = 8421504
                        End If
                    End If
                End If
            End If
        Case RibTab_HomeBanki:
            If Row.Record.ItemCount > 16 Then
                If LCase(Row.Record(Ban_Selekt).Value) = "ja" Then 'Zugeordnet oder gekennzeichnet
                    If CDbl(Row.Record(Ban_KoBetrag).Value) < 0 Then
                        Metrics.ForeColor = 210 'rot
                    ElseIf CDbl(Row.Record(Ban_KoBetrag).Value) > 0 Then
                        If CDbl(Row.Record(Ban_GeBetrag).Value) > 0 Then
                            If CDbl(Row.Record(Ban_KoBetrag).Value) > CDbl(Row.Record(Ban_GeBetrag).Value) Then
                                Metrics.ForeColor = 33023 'orange
                            ElseIf CDbl(Row.Record(Ban_KoBetrag).Value) < CDbl(Row.Record(Ban_GeBetrag).Value) Then
                                Metrics.ForeColor = 10519290 'rosa
                            ElseIf CDbl(Row.Record(Ban_KoBetrag).Value) = CDbl(Row.Record(Ban_GeBetrag).Value) Then
                                Metrics.ForeColor = 44800 'grün
                            End If
                        Else
                            Metrics.ForeColor = 54528 'helgrün
                        End If
                    End If
                    If LCase(Row.Record(Ban_Bezahlt).Value) = "nein" Then 'Bezahlt
                        Metrics.Font.Bold = True
                    End If
                Else
                    If LCase(Row.Record(Ban_Bezahlt).Value) = "ja" Then 'Bezahlt
                        If CDbl(Row.Record(Ban_KoBetrag).Value) < 0 Then
                            Metrics.ForeColor = 210 'rot
                        Else
                            Metrics.ForeColor = 54528 'helgrün
                        End If
                    End If
                    If Row.Record.ItemCount > 37 Then
                        If Row.Record(Ban_IDZ).Value <> vbNullString Then
                            If IsNumeric(Row.Record(Ban_IDZ).Value) = True Then
                                If CInt(Row.Record(Ban_IDZ).Value) < 0 Then
                                    Metrics.ForeColor = 8421504 'grau
                                End If
                            End If
                        End If
                    End If
                End If
                If Row.Record(Ban_GeBetrag).Value = vbNullString Then
                    Row.Record(Ban_GeBetrag).Value = 0
                End If
            End If
        Case RibTab_Ter_Listen:
            If IsNumeric(Row.Record(Ter_Farbe).Value) Then
                FrbZa = Row.Record(Ter_Farbe).Value
                If FrbZa > 1 And FrbZa <= 20 Then
                    Metrics.BackColor = GlTmF(FrbZa, 1)
                End If
            End If
            If IsDate(Row.Record(Ter_VonDat).Value) = True Then
                If CDate(Row.Record(Ter_VonDat).Value) >= Date Then
                    Metrics.Font.Bold = True
                End If
            End If
            If Row.Record.ItemCount >= Ter_Passiv Then
                If Row.Record(Ter_Passiv).Value <> vbNullString Then
                    If CBool(Row.Record(Ter_Passiv).Value) = True Then
                        Metrics.Font.Strikethrough = True
                        Metrics.ForeColor = 8421504
                    Else
                        If Row.Selected = True Then
                            FaRGB = WindowRGB(Metrics.BackColor)
                            If (FaRGB.rot - 50) > 0 Then
                                FarRo = FaRGB.rot - 50
                            Else
                                FarRo = FaRGB.rot
                            End If
                            If (FaRGB.grün - 50) > 0 Then
                                FarGr = FaRGB.grün - 50
                            Else
                                FarGr = FaRGB.grün
                            End If
                            If (FaRGB.blau - 50) > 0 Then
                                FarBl = FaRGB.blau - 50
                            Else
                                FarBl = FaRGB.blau
                            End If
                            Metrics.BackColor = RGB(FarRo, FarGr, FarBl)
                        End If
                    End If
                End If
            End If
        Case RibTab_Ter_Akont:
            If IsNumeric(Row.Record(Ter_Farbe).Value) Then
                FrbZa = Row.Record(Ter_Farbe).Value
                If FrbZa > 1 And FrbZa <= 20 Then
                    Metrics.BackColor = GlTmF(FrbZa, 1)
                End If
            End If
            If Row.Record(Ter_Fallig1).Value <> vbNullString Then
                If CDate(Row.Record(Ter_Fallig1).Value) <= Date Then
                    If Row.Record(Ter_SerBet).Value > Row.Record(Ter_BezBet).Value Then
                        Metrics.Font.Bold = True
                        Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnunge
                    End If
                End If
            End If
        Case RibTab_Ter_Warte:
            If IsNumeric(Row.Record(Ter_Farbe).Value) Then
                FrbZa = Row.Record(Ter_Farbe).Value
                If FrbZa > 1 And FrbZa <= 20 Then
                    Metrics.BackColor = GlTmF(FrbZa, 1)
                End If
            End If
            If IsDate(Row.Record(Ter_VonDat).Value) = True Then
                If CDate(Row.Record(Ter_VonDat).Value) >= Date Then
                    Metrics.Font.Bold = True
                End If
            End If
            If Row.Record.ItemCount >= Ter_Passiv Then
                If Row.Record(Ter_Passiv).Value <> vbNullString Then
                    If CBool(Row.Record(Ter_Passiv).Value) = True Then
                        Metrics.Font.Strikethrough = True
                        Metrics.ForeColor = 8421504
                    Else
                        If Row.Selected = True Then
                            FaRGB = WindowRGB(Metrics.BackColor)
                            If (FaRGB.rot - 50) > 0 Then
                                FarRo = FaRGB.rot - 50
                            Else
                                FarRo = FaRGB.rot
                            End If
                            If (FaRGB.grün - 50) > 0 Then
                                FarGr = FaRGB.grün - 50
                            Else
                                FarGr = FaRGB.grün
                            End If
                            If (FaRGB.blau - 50) > 0 Then
                                FarBl = FaRGB.blau - 50
                            Else
                                FarBl = FaRGB.blau
                            End If
                            Metrics.BackColor = RGB(FarRo, FarGr, FarBl)
                        End If
                    End If
                End If
            End If
        End Select
    End If
End If

End Sub
Private Sub repCont1_ColumnOrderChanged()
    If GlAkt = False Then SSpSav
End Sub

Private Sub repCont1_ColumnOrderChangedEx(ByVal Column As XtremeReportControl.IReportColumn, ByVal Reason As XtremeReportControl.XTPReportColumnOrderChangedReason)
    If GlAkt = False Then SSpSav
End Sub
Private Sub repCont1_ColumnWidthChanged(ByVal Column As XtremeReportControl.IReportColumn, ByVal PrevWidth As Long, ByVal NewWidth As Long)
    If GlAkt = False Then SSpSav
End Sub

Private Sub repCont1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            If GlAkt = False Then
                Select Case GlBut
                Case RibTab_Mahnwesen: frmOPEdit.Show vbModal
                Case RibTab_Buchungen: frmBuEdit.Show
                Case RibTab_HomeBanki: frmBaEdit.Show vbModal
                Case RibTab_Ter_Listen: STerm
                Case RibTab_Ter_Akont: STerm
                Case RibTab_Ter_Warte: STerm
                End Select
            End If
        End If
    End If
End Sub

Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        If Shift = 0 Then
            SMark
        End If
    End If
End Sub

Private Sub repCont1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo LaErr

If GlAkt = False Then
    Dim RpCo1 As XtremeReportControl.ReportControl
    Set RpCo1 = Me.repCont1
    If RpCo1.Records.Count > 0 Then
        Select Case RpCo1.HitTest(x, y).ht
        Case xtpHitTestGroupBox:
        Case xtpHitTestHeader:
        Case xtpHitTestReportArea: SMark
        Case xtpHitTestUnknown:
        End Select
    End If
    Set RpCo1 = Nothing
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub repCont1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo1 = Me.repCont1
Set RpRws = RpCo1.Rows
Set HiTes = RpCo1.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

If RpRws.Count > 0 Then
    Select Case HiTes.ht
    Case xtpHitTestGroupBox:
    Case xtpHitTestHeader:
    Case xtpHitTestReportArea:
            If Button = vbRightButton Then
                SMePo
            Else
                If GlBut = RibTab_Ter_Warte Then
                    STeDe
                End If
            End If
    Case xtpHitTestUnknown:
    End Select
End If

Set RpRws = Nothing
Set RpCo1 = Nothing

End Sub

Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        If Row.GroupRow = False Then
            SMark
            Select Case GlBut
            Case RibTab_Mahnwesen:
                    frmOPEdit.Show vbModal
            Case RibTab_Buchungen:
                    frmBuEdit.Show
            Case RibTab_HomeBanki:
                    frmBaEdit.Show vbModal
            Case RibTab_Ter_Listen:
                    STerm
            Case RibTab_Ter_Akont:
                    STerm
            Case RibTab_Ter_Warte:
                    STerm
            Case RibTab_LabBerichte:
                    Select Case GlAdO
                    Case 0: SLaZe
                    Case 1: SLaZe
                    Case 2: frmLaBear.Show vbModal
                    End Select
            Case RibTab_LabAuftrage:
                    Select Case GlAdO
                    Case 0: SLaZe
                    Case 1: SLaZe
                    Case 2: frmLaBear.Show vbModal
                    End Select
            End Select
        End If
    End If
End Sub
Private Sub repCont2_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim AktZa As Integer

If GlSta = False Then
    If GlAdA > 0 Then 'Anzahl gefundener Adressen
        If GlAdA > Row.Index Then
            If AdAry(Item.Index, Row.Index) <> vbNullString Then
                Metrics.Text = AdAry(Item.Index, Row.Index)
                If Item.Index = Adr_IDKurz Then
                    Metrics.ItemIcon = IC16_IDCard_Norm
                End If
            End If
            If Row.GroupRow = False Then
                Select Case GlBut
                Case RibTab_Mandanten:
                        Metrics.ForeColor = 44800
                Case RibTab_Verordner:
                        Metrics.ForeColor = 16711680
                Case RibTab_Mitarbeit:
                        Metrics.ForeColor = 2162853
                Case Else:
                    For AktZa = 1 To UBound(GlGKa)
                        If GlGKa(AktZa, 0) = AdAry(Adr_ID3, Row.Index) Then
                            Metrics.ForeColor = GlGKa(AktZa, 3)
                            Exit For
                        End If
                    Next AktZa
                End Select
                If CBool(AdAry(Adr_Passiv, Row.Index)) = True Then
                    Metrics.Font.Strikethrough = True
                    Metrics.ForeColor = 8421504
                End If
                If IsNumeric(AdAry(Adr_Mandant, Row.Index)) = True Then
                    AdAry(Adr_Mandant, Row.Index) = Format$(AdAry(Adr_Mandant, Row.Index), "000000")
                End If
                If AdAry(Adr_Versand, Row.Index) <> vbNullString Then
                    If IsNumeric(AdAry(Adr_Versand, Row.Index)) = True Then
                        Select Case CInt(AdAry(Adr_Versand, Row.Index))
                        Case 0: AdAry(Adr_Versand, Row.Index) = vbNullString
                        Case 1: AdAry(Adr_Versand, Row.Index) = "M"
                        Case 2: AdAry(Adr_Versand, Row.Index) = "D"
                        End Select
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub repCont2_ColumnOrderChanged()
    If GlAkt = False Then SSpSav
End Sub
Private Sub repCont2_ColumnOrderChangedEx(ByVal Column As XtremeReportControl.IReportColumn, ByVal Reason As XtremeReportControl.XTPReportColumnOrderChangedReason)
    If GlAkt = False Then
        SSpSav
    End If
End Sub
Private Sub repCont2_ColumnWidthChanged(ByVal Column As XtremeReportControl.IReportColumn, ByVal PrevWidth As Long, ByVal NewWidth As Long)
    If GlAkt = False Then
        SSpSav
    End If
End Sub
Private Sub repCont2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            If GlAkt = False Then
                Select Case GlAdO
                Case 0:
                    Select Case GlBut
                    Case RibTab_Mandanten: SAdre 2, True
                    Case RibTab_Verordner: SAdre 2, True
                    Case RibTab_Mitarbeit: SAdre 2, True
                    Case Else:
                            SAbZe
                            If GlDaK = False Then
                                FDaK5 Date
                            End If
                    End Select
                Case 1:
                    Select Case GlBut
                    Case RibTab_Mandanten: SAdre 2, True
                    Case RibTab_Verordner: SAdre 2, True
                    Case RibTab_Mitarbeit: SAdre 2, True
                    Case Else:
                            SKrZe
                            If GlDaK = False Then
                                FDaK5 Date
                            End If
                    End Select
                Case 2:
                    Select Case GlBut
                    Case RibTab_Mandanten: SAdre 2, True
                    Case RibTab_Verordner: SAdre 2, True
                    Case RibTab_Mitarbeit: SAdre 2, True
                    Case Else: SAdre 2
                    End Select
                End Select
            End If
        End If
    End If
End Sub
Private Sub repCont2_KeyUp(KeyCode As Integer, Shift As Integer)

Dim RpCo2 As XtremeReportControl.ReportControl

Set RpCo2 = Me.repCont2

If GlAkt = False Then
    If RpCo2.Records.Count > 0 Then
        Set RpSel = RpCo2.SelectedRows
        If RpSel.Count > 0 Then
            If Shift = 0 Then
                If KeyCode >= 65 And KeyCode <= 90 Then
                    With GlSuP
                        .SuIdx = 7
                        .SuStr = Chr$(KeyCode)
                    End With
                    SSuch
                Else
                    SMark
                End If
            End If
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo2 = Nothing

End Sub

Private Sub repCont2_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo LaErr

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo2 = Me.repCont2
Set RpRws = RpCo2.Rows
Set HiTes = RpCo2.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

If RpRws.Count > 0 Then
    Select Case HiTes.ht
    Case xtpHitTestGroupBox:
    Case xtpHitTestHeader:
    Case xtpHitTestReportArea:
            SMark
            If Button = vbRightButton Then
                SMePo 2
            End If
    Case xtpHitTestUnknown:
    End Select
End If

Set RpRws = Nothing
Set RpCo2 = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub repCont2_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        If Row.GroupRow = False Then
            Select Case GlAdO
            Case 0:
                Select Case GlBut
                Case RibTab_Mandanten: MMain AdAry(Adr_ID0, Row.Index)
                Case RibTab_Verordner: MMain AdAry(Adr_ID0, Row.Index)
                Case RibTab_Mitarbeit: MMain AdAry(Adr_ID0, Row.Index)
                Case Else:
                        SAbZe
                        If GlDaK = False Then
                            FDaK5 Date
                        End If
                End Select
            Case 1:
                Select Case GlBut
                Case RibTab_Mandanten: MMain AdAry(Adr_ID0, Row.Index)
                Case RibTab_Verordner: MMain AdAry(Adr_ID0, Row.Index)
                Case RibTab_Mitarbeit: MMain AdAry(Adr_ID0, Row.Index)
                Case Else:
                        SKrZe
                        If GlDaK = False Then
                            FDaK5 Date
                        End If
                End Select
            Case 2:
                Select Case GlBut
                Case RibTab_Mandanten: MMain AdAry(Adr_ID0, Row.Index)
                Case RibTab_Verordner: MMain AdAry(Adr_ID0, Row.Index)
                Case RibTab_Mitarbeit: MMain AdAry(Adr_ID0, Row.Index)
                Case Else: AMain AdAry(Adr_ID0, Row.Index)
                End Select
            End Select
        End If
    End If
End Sub

Private Sub repCont4_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If GlSta = False Then
    If Row.GroupRow = False Then
        Select Case Row.Record(Rec_Type).Value
        Case "M": Metrics.ForeColor = 16744448
        Case "L": Metrics.ForeColor = 33023
        Case "V": Metrics.ForeColor = 8421631
        Case "I": Metrics.ForeColor = 13138080
        Case "U": Metrics.ForeColor = 6604830
        Case Else:
            If Row.Record(Rec_Selekt).Value = "Nein" Then
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnunge
            End If
        End Select
        If Row.Record(Rec_Storniert).Value = True Then
            Metrics.Font.Strikethrough = True
            Metrics.ForeColor = 8421504
        End If
        If Row.Record(Rec_Selekt).Value = "Nein" Then
            Metrics.Font.Bold = True
        End If
    End If
End If

End Sub

Private Sub repCont4_ColumnOrderChanged()
    If GlAkt = False Then SSpSav
End Sub

Private Sub repCont4_ColumnOrderChangedEx(ByVal Column As XtremeReportControl.IReportColumn, ByVal Reason As XtremeReportControl.XTPReportColumnOrderChangedReason)
    If GlAkt = False Then SSpSav
End Sub
Private Sub repCont4_ColumnWidthChanged(ByVal Column As XtremeReportControl.IReportColumn, ByVal PrevWidth As Long, ByVal NewWidth As Long)
    If GlAkt = False Then SSpSav
End Sub
Private Sub repCont4_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            If GlAkt = False Then
                Select Case GlAdO
                Case 0: SReZe 0, True
                Case 1: SReZe 0, True
                Case 2: frmReEdit.Show vbModal
                End Select
            End If
        End If
    End If
End Sub

Private Sub repCont4_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        If Shift = 0 Then
            SMark
        End If
    End If
End Sub

Private Sub repCont4_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo LaErr

If GlAkt = False Then
    Dim RpCo4 As XtremeReportControl.ReportControl
    Set RpCo4 = Me.repCont4
    If RpCo4.Records.Count > 0 Then
        Select Case RpCo4.HitTest(x, y).ht
        Case xtpHitTestGroupBox:
        Case xtpHitTestHeader:
        Case xtpHitTestReportArea: SMark
        Case xtpHitTestUnknown:
        End Select
    End If
    Set RpCo4 = Nothing
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub repCont4_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo4 = Me.repCont4
Set RpRws = RpCo4.Rows
Set HiTes = RpCo4.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

If RpRws.Count > 0 Then
    Select Case HiTes.ht
    Case xtpHitTestGroupBox:
    Case xtpHitTestHeader:
    Case xtpHitTestReportArea:
            If Button = vbRightButton Then
                SMePo 2
            End If
    Case xtpHitTestUnknown:
    End Select
End If

Set RpRws = Nothing
Set RpCo4 = Nothing
    
End Sub
Private Sub repCont4_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        If Row.GroupRow = False Then
            Select Case GlAdO
            Case 0: SReZe 0, True
            Case 1: SReZe 0, True
            Case 2: frmReEdit.Show vbModal
            End Select
        End If
    End If
End Sub

Private Sub repCont5_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Select Case GlBut
Case RibTab_Rezeptmodul:
    If Item.Record(Rzp_Storniert).Value <> vbNullString Then
        If CBool(Item.Record(Rzp_Storniert).Value) = True Then
            Metrics.Font.Strikethrough = True
            Metrics.ForeColor = 8421504
        End If
    End If
Case RibTab_Belegmodul:
    If Item.Record(Rzp_Storniert).Value <> vbNullString Then
        If CBool(Item.Record(Rzp_Storniert).Value) = True Then
            Metrics.Font.Strikethrough = True
            Metrics.ForeColor = 8421504
        End If
    End If
End Select

End Sub
Private Sub repCont6_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim TmPfa As Long
Dim FrbZa As Long
Dim FarRo As Long
Dim FarGr As Long
Dim FarBl As Long
Dim TmpFo As String
Dim AktZa As Integer
Dim KrTyp As Integer
Dim FaRGB As tRGB

If GlSta = False Then
    If Row.GroupRow = False Then
        Select Case GlBut
        Case RibTab_Fragebogen:
                If Item.Record(Ana_Marke).Value <> vbNullString Then
                    If CBool(Item.Record(Ana_Marke).Value) = True Then
                        Metrics.ForeColor = 32768
                    End If
                End If
        Case RibTab_Tagesproto:
                If Item.Record(Kra_KrTyp).Value <> vbNullString Then
                    If IsNumeric(Item.Record(Kra_KrTyp).Value) = True Then
                        KrTyp = Item.Record(Kra_KrTyp).Value
                        For AktZa = 1 To UBound(GlKrA)
                            If KrTyp = GlKrA(AktZa, 0) Then
                                Metrics.ForeColor = GlKrA(AktZa, 3)
                                Exit For
                            End If
                        Next AktZa
                    End If
                End If
        Case RibTab_Abrechnung:
                If Item.Record(Kra_KrTyp).Value <> vbNullString Then
                    If IsNumeric(Item.Record(Kra_KrTyp).Value) = True Then
                        KrTyp = Item.Record(Kra_KrTyp).Value
                        For AktZa = 1 To UBound(GlKrA)
                            If KrTyp = GlKrA(AktZa, 0) Then
                                Metrics.ForeColor = GlKrA(AktZa, 3)
                                Exit For
                            End If
                        Next AktZa
                    End If
                End If
        Case RibTab_Vorbereit:
                If GlAzA = True Then
                    Metrics.ForeColor = 8421504
                End If
        Case RibTab_LabBericht:
                If Item.Record(Lbl_Grenzwert).Value <> vbNullString Then
                    Select Case LTrim$(Item.Record(Lbl_Grenzwert).Value)
                    Case "(*)":
                    Case "(+)": Metrics.ForeColor = GlFaL(2)
                    Case "(-)": Metrics.ForeColor = GlFaL(2)
                    Case "+": Metrics.ForeColor = GlFaL(1)
                    Case "-": Metrics.ForeColor = GlFaL(1)
                    Case "++": Metrics.ForeColor = GlFaL(1)
                    Case "--": Metrics.ForeColor = GlFaL(1)
                    End Select
                End If
        Case RibTab_Ter_Warte:
                If IsNumeric(Row.Record(War_Farbe).Value) = True Then
                    FrbZa = Row.Record(War_Farbe).Value
                    If FrbZa > 1 And FrbZa <= 20 Then
                        Metrics.BackColor = GlTmF(FrbZa, 1)
                    End If
                End If
                If Row.Selected = True Then
                    FaRGB = WindowRGB(Metrics.BackColor)
                    If (FaRGB.rot - 50) > 0 Then
                        FarRo = FaRGB.rot - 50
                    Else
                        FarRo = FaRGB.rot
                    End If
                    If (FaRGB.grün - 50) > 0 Then
                        FarGr = FaRGB.grün - 50
                    Else
                        FarGr = FaRGB.grün
                    End If
                    If (FaRGB.blau - 50) > 0 Then
                        FarBl = FaRGB.blau - 50
                    Else
                        FarBl = FaRGB.blau
                    End If
                    Metrics.BackColor = RGB(FarRo, FarGr, FarBl)
                End If
        End Select
    End If
End If

End Sub
Private Sub repCont6_BeginEdit(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim RetWe As Long
Dim RpCo6 As XtremeReportControl.ReportControl

Set RpCo6 = Me.repCont6

If GlKrE = False Then
    RetWe = RpCo6.EnableDragDrop("Katalog", xtpReportAllowDrop)
End If

End Sub

Private Sub repCont8_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim TreKy As String
Dim KrTyp As Integer
Dim AktZa As Integer

TreKy = Left$(GlNod, 1)

If GlSta = False Then
    If GlBut = RibTab_Kat_Eintrg Then
        If TreKy = "A" Then 'Gebühren
            If Row.GroupRow = False Then
                If Item.Record(Kat_Typ).Value <> vbNullString Then
                    If IsNumeric(Item.Record(Kat_Typ).Value) = True Then
                        KrTyp = Item.Record(Kat_Typ).Value
                        For AktZa = 1 To UBound(GlKrA)
                            If KrTyp = GlKrA(AktZa, 0) Then
                                Metrics.ForeColor = GlKrA(AktZa, 3)
                                Exit For
                            End If
                        Next AktZa
                    End If
                End If
            End If
        End If
    End If
End If

End Sub
Private Sub repCont9_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo9 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo9 = Me.repCont9
Set RpRws = RpCo9.Rows
Set HiTes = RpCo9.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

If RpRws.Count > 0 Then
    Select Case HiTes.ht
    Case xtpHitTestGroupBox:
    Case xtpHitTestHeader:
    Case xtpHitTestReportArea:
            If Button = vbRightButton Then
                SMePo 2
            End If
    Case xtpHitTestUnknown:
    End Select
End If

Set RpRws = Nothing
Set RpCo9 = Nothing

End Sub
Private Sub repContK_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim TmPfa As Long
Dim FaRGB As tRGB
Dim TmpFo As String
Dim KrTyp As Integer
Dim AktZa As Integer

If GlSta = False Then
    If Row.GroupRow = False Then
        If Item.Record(Kra_KrTyp).Value <> vbNullString Then
            If IsNumeric(Item.Record(Kra_KrTyp).Value) Then
                KrTyp = Item.Record(Kra_KrTyp).Value
                For AktZa = 1 To UBound(GlKrA)
                    If KrTyp = GlKrA(AktZa, 0) Then
                        TmPfa = GlKrA(AktZa, 3)
                        Exit For
                    End If
                Next AktZa
                Select Case KrTyp
                Case 24: Item.Editable = False
                Case 101: Item.Editable = False
                Case 102: Item.Editable = False
                Case 104: Item.Editable = False
                Case 105: Item.Editable = False
                End Select
            End If
        End If
        If GlFaU = True Then 'Farbunterscheidung Krankenblatt
            If Item.Record(Kra_Provision).Value <> vbNullString Then
                If Len(Item.Record(Kra_Provision).Value) > 5 Then
                    TmpFo = Item.Record(Kra_Provision).Value
                    If IsNumeric(Mid$(TmpFo, 6, 8)) = True Then
                        Metrics.ForeColor = CLng(Mid$(TmpFo, 6, 8))
                    End If
                    If IsNumeric(Mid$(TmpFo, 14, 8)) = True Then
                        Metrics.BackColor = CLng(Mid$(TmpFo, 14, 8))
                    End If
                    If Item.Index = Kra_Bezeichnung Then
                        If Mid$(TmpFo, 1, 1) = "1" Then Metrics.Font.Bold = True
                        If Mid$(TmpFo, 2, 1) = "1" Then Metrics.Font.Italic = True
                        If Mid$(TmpFo, 3, 1) = "1" Then Metrics.Font.Underline = True
                        If Mid$(TmpFo, 4, 1) = "1" Then Metrics.Font.Strikethrough = True
                        If Mid$(TmpFo, 22, 2) <> vbNullString Then Metrics.Font.SIZE = CLng(Mid$(TmpFo, 22, 2))
                        If (Len(TmpFo) - 23) > 0 Then Metrics.Font.Name = Mid$(TmpFo, 24, Len(TmpFo) - 23)
                    End If
                    If Row.Selected = True Then
                        FaRGB = WindowRGB(Metrics.BackColor)
                        Metrics.BackColor = RGB(IIf((FaRGB.rot - 50) > 0, FaRGB.rot - 50, FaRGB.rot), IIf((FaRGB.grün - 50) > 0, FaRGB.grün - 50, FaRGB.grün), IIf((FaRGB.blau - 50) > 0, FaRGB.blau - 50, FaRGB.blau))
                    End If
                Else
                    Metrics.ForeColor = TmPfa
                End If
            Else
                Metrics.ForeColor = TmPfa
            End If
        End If
        If CBool(Item.Record(Kra_Storniert).Value) = True Then
            Metrics.Font.Strikethrough = True
            Metrics.ForeColor = 8421504
        End If
    End If
End If

End Sub
Private Sub repContK_BeginEdit(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)

Dim RetWe As Long
Dim RpCoK As XtremeReportControl.ReportControl

Set RpCoK = Me.repContK

If GlKrE = False Then
    RetWe = RpCoK.EnableDragDrop("Katalog", xtpReportAllowDrop)
End If

End Sub

Private Sub repCont6_EditCanceled(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim RetWe As Long
Dim RpCo6 As XtremeReportControl.ReportControl

Set RpCo6 = Me.repCont6

If GlKrE = False Then
    RetWe = RpCo6.EnableDragDrop("Katalog", xtpReportAllowDrag + xtpReportAllowDrop)
End If

End Sub
Private Sub repContK_EditCanceled(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)

Dim RetWe As Long
Dim RpCoK As XtremeReportControl.ReportControl

Set RpCoK = Me.repContK

If GlDiB = False Then
    RetWe = RpCoK.EnableDragDrop("Katalog", xtpReportAllowDrag + xtpReportAllowDrop)
End If

End Sub

Private Sub repCont6_InplaceButtonDown(ByVal Button As XtremeReportControl.IReportInplaceButton)
On Error Resume Next

If GlAkt = False Then
    If GlBut = RibTab_Abrechnung Then
        If Button.Column.ItemIndex = Kra_Datum Then
            SDaKa Button
        End If
    End If
End If

End Sub
Private Sub repCont6_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String
Dim TmpTg As String

If GlAkt = False Then
    TmpTg = Item.Tag
    TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
    Item.Tag = "@" & TmTag
    FKran
    Item.Tag = TmpTg
End If

End Sub
Private Sub repCont6_KeyUp(KeyCode As Integer, Shift As Integer)
    FTaEd KeyCode, Shift
End Sub
Private Sub repCont6_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo6 = Me.repCont6
Set RpRws = RpCo6.Rows
Set RpSel = RpCo6.SelectedRows
Set HiTes = RpCo6.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

Select Case HiTes.ht
Case xtpHitTestGroupBox:
Case xtpHitTestHeader:
Case xtpHitTestReportArea:
        If Button = vbRightButton Then
            If GlBut = RibTab_Ter_Warte Then
                GlWaZ = True
            End If
            SMePo 2
        Else
            If GlBut = RibTab_Vorbereit Then
                SMark
            ElseIf GlBut = RibTab_Ter_Warte Then
                SWaDe
            Else
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        SAbDa
                    Else
                        FMark 2
                    End If
                End If
            End If
        End If
Case xtpHitTestUnknown:
End Select

Set RpRws = Nothing
Set RpCo6 = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub repContK_GotFocus()
    If GlAkt = False Then
        If GlBut = RibTab_Krankenbla Then
            GlMen = False
            FMenu
        End If
    End If
End Sub
Private Sub repContK_KeyUp(KeyCode As Integer, Shift As Integer)
    FTaEd KeyCode, Shift
End Sub
Private Sub repCont6_OLEDragDrop(ByVal data As XtremeReportControl.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next

Dim AktZa As Integer
Dim DroDa As XtremeReportControl.DataObjectFiles

If data.GetFormat(vbCFFiles) Then
    Set DroDa = data.Files
    If Not (DroDa Is Nothing) Then
        ReDim Preserve GlDro(DroDa.Count)
        For AktZa = 1 To UBound(GlDro)
            GlDro(AktZa) = DroDa(AktZa)
        Next AktZa
        SFilDr
    End If
End If

End Sub

Private Sub repCont6_RecordsDropped(ByVal TargetRecord As XtremeReportControl.IReportRecord, ByVal Records As XtremeReportControl.IReportRecords, ByVal Above As Boolean)
On Error GoTo LaErr

Dim NeDat As Date
Dim IdxNr As Long
Dim ReNum As Long
Dim RetWe As Long
Dim NeSor As Long
Dim AnzPo As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmKatGE
Set CmBrs = FM.comBar02
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)

Set RpCo6 = Me.repCont6
Set RpRws = RpCo6.Rows

AnzPo = RpCo6.Records.Count

If GlAkt = False Then
    If IsNumeric(Records(0).Item(Kra_ID3).Value) = True Then
        If Records(0).Item(Kra_ID3).Value > 0 Then
            If GlBut = RibTab_Abrechnung Then
                If Not TargetRecord Is Nothing Then
                    ReNum = TargetRecord(Kra_IDR).Value
                    NeSor = TargetRecord(Kra_ID3).Value
                    NeDat = TargetRecord(Kra_Datum).Value
                    IdxNr = Records(0).Item(Kra_ID2).Value
                    S_KrSor ReNum, NeDat, IdxNr, NeSor, Above
                    If GlKrE = False Then 'Krankenblatteditmodus
                        DoEvents
                        RetWe = RpCo6.EnableDragDrop("Katalog", xtpReportAllowDrag + xtpReportAllowDrop)
                    End If
                    For Each RpRow In RpRws
                        If RpRow.Selected = True Then
                            RpRow.Selected = False
                        End If
                    Next RpRow
                End If
            End If
        Else
            If AnzPo = 1 Then
                FDrop
            Else
                If Not TargetRecord Is Nothing Then
                    NeSor = TargetRecord(Kra_ID3).Value
                    NeDat = TargetRecord(Kra_Datum).Value
                Else
                    NeSor = 0
                    NeDat = CDate(CmEdt.Text)
                End If
                Select Case GlBut
                Case RibTab_LabBericht: FDrop
                Case RibTab_LabAuftrag: FDrop
                Case Else: FDrop NeDat, NeSor, Above
                End Select
            End If
        End If
    Else
        If AnzPo = 1 Then
            FDrop
        Else
            If Not TargetRecord Is Nothing Then
                NeSor = TargetRecord(Kra_ID3).Value
                NeDat = TargetRecord(Kra_Datum).Value
            Else
                NeSor = 0
                NeDat = CDate(CmEdt.Text)
            End If
            Select Case GlBut
            Case RibTab_LabBericht: FDrop
            Case RibTab_LabAuftrag: FDrop
            Case RibTab_Tagesproto:
            Case RibTab_Abrechnung: FDrop NeDat, NeSor, Above
            End Select
        End If
    End If
End If

Set RpRws = Nothing
Set RpCo6 = Nothing
Set CmBrs = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub repContK_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCoK = Me.repContK
Set RpRws = RpCoK.Rows
Set RpSel = RpCoK.SelectedRows
Set HiTes = RpCoK.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

Select Case HiTes.ht
Case xtpHitTestGroupBox:
Case xtpHitTestHeader:
Case xtpHitTestReportArea:
        If Button = vbRightButton Then
            SMePo 2
        Else
            If GlBut = RibTab_Vorbereit Then
                SMark
            Else
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        If MauDo = False Then
                            SKrVo
                        End If
                    Else
                        FMark 2
                    End If
                End If
            End If
        End If
Case xtpHitTestUnknown:
End Select

Set RpRws = Nothing
Set RpCoK = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub repContK_OLEDragDrop(ByVal data As XtremeReportControl.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error Resume Next

Dim AktZa As Integer
Dim DroDa As XtremeReportControl.DataObjectFiles

If data.GetFormat(vbCFFiles) Then
    Set DroDa = data.Files
    If Not (DroDa Is Nothing) Then
        ReDim Preserve GlDro(DroDa.Count)
        For AktZa = 1 To UBound(GlDro)
            GlDro(AktZa) = DroDa(AktZa)
        Next AktZa
        SFilDr
    End If
End If

End Sub
Private Sub repContK_RecordsDropped(ByVal TargetRecord As XtremeReportControl.IReportRecord, ByVal Records As XtremeReportControl.IReportRecords, ByVal Above As Boolean)
On Error GoTo LaErr

Dim RetWe As Long
Dim AnzPo As Integer
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCoK = Me.repContK
Set RpRws = RpCoK.Rows

AnzPo = RpCoK.Records.Count

If GlAkt = False Then
    If IsNumeric(Records(0).Item(3).Value) Then
        If AnzPo = 1 Then
            FDrop
        Else
            If Not TargetRecord Is Nothing Then
                FDrop TargetRecord(Kra_Datum).Value, TargetRecord(Kra_ID3).Value, Above
            Else
                FDrop
            End If
        End If
    Else
        If AnzPo = 1 Then
            FDrop
        Else
            If Not TargetRecord Is Nothing Then
                FDrop TargetRecord(Kra_Datum).Value, TargetRecord(Kra_ID3).Value, Above
            Else
                FDrop
            End If
        End If
    End If
End If

Set RpRws = Nothing
Set RpCoK = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub repCont6_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        Select Case GlBut
        Case RibTab_Abrechnung: SKrDo
        Case RibTab_Vorbereit: SKrDo
        Case RibTab_Ter_Warte: SWaSe 3
        End Select
    End If
End Sub
Private Sub repCont6_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String
Dim TmpTg As String
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo6 = FM.repCont6

If GlAkt = False Then
    If Item.Tag <> vbNullString Then
        TmpTg = Item.Tag
        TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
        Item.Tag = "@" & TmTag
        FKran Column.ItemIndex
        Item.Tag = TmpTg
    End If
End If

End Sub
Private Sub repCont6_ValueChanging(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem, NewValue As Variant, Cancel As Boolean)
On Error Resume Next

If GlAkt = False Then
    If GlBut = RibTab_Abrechnung Then
        If Row.GroupRow = False Then
            Select Case Column.Index
            Case Kra_Ziffer:
                If NewValue = vbNullString Then
                    Cancel = True
                    SPopu "Keine Ziffer vorhanden", "Die Spalte Ziffer darf nicht leer sein.", IC48_Warning
                End If
            Case Kra_Anz:
                If NewValue <> vbNullString Then
                    If IsNumeric(NewValue) = True Then
                        If NewValue >= 1 Then
                            If K_RePr(CInt(Item.Value), CInt(NewValue)) = False Then 'Regelprüfung
                                Cancel = True
                            End If
                            If K_BePr(CSng(NewValue)) < 0 Then 'Lagerbestandsprüfung
                                Cancel = True
                            End If
                            If S_KrPr(Row.Index, Column.Index, NewValue) = True Then 'Anzahlungsprüfung
                                Cancel = True
                                SPopu "Rechnungsbetrag zu gering", "Der Rechnungsbetrag ist geringer als die Summe aller Zahlungen.", IC48_Warning
                            End If
                        Else
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Else
                    Cancel = True
                End If
            Case Kra_Faktor:
                If NewValue <> vbNullString Then
                    If IsNumeric(NewValue) = True Then
                        If NewValue >= 0.25 Then
                            If S_KrPr(Row.Index, Column.Index, NewValue) = True Then 'Anzahlungsprüfung
                                Cancel = True
                                SPopu "Rechnungsbetrag zu gering", "Der Rechnungsbetrag ist geringer als die Summe aller Zahlungen.", IC48_Warning
                            End If
                        Else
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Else
                    Cancel = True
                End If
            Case Kra_Betrag:
                If NewValue <> vbNullString Then
                    If IsNumeric(NewValue) = True Then
                        If NewValue >= 0 Then
                            If S_KrPr(Row.Index, Column.Index, NewValue) = True Then 'Anzahlungsprüfung
                                Cancel = True
                                SPopu "Rechnungsbetrag zu gering", "Der Rechnungsbetrag ist geringer als die Summe aller Zahlungen.", IC48_Warning
                            End If
                        Else
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Else
                    Cancel = True
                End If
            End Select
        End If
    End If
End If

End Sub
Private Sub repCont8_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            If GlAkt = False Then
                Select Case GlBut
                Case RibTab_Kat_Eintrg: KaEdi
                Case RibTab_Kat_Ketten: KaEdi
                Case RibTab_Kat_Frage: KaEdi
                End Select
            End If
        End If
    End If
End Sub

Private Sub repCont8_KeyUp(KeyCode As Integer, Shift As Integer)
    
Dim RpCo8 As XtremeReportControl.ReportControl

Set RpCo8 = Me.repCont8

If GlAkt = False Then
    If Shift = 0 Then
        If RpCo8.Records.Count > 0 Then
            Set RpSel = RpCo8.SelectedRows
            Select Case KeyCode
            Case 65: FSuLe "A", SuLei_A
            Case 66: FSuLe "B", SuLei_B
            Case 67: FSuLe "C", SuLei_C
            Case 68: FSuLe "D", SuLei_D
            Case 69: FSuLe "E", SuLei_E
            Case 70: FSuLe "F", SuLei_F
            Case 71: FSuLe "G", SuLei_G
            Case 72: FSuLe "H", SuLei_H
            Case 73: FSuLe "I", SuLei_I
            Case 74: FSuLe "J", SuLei_J
            Case 75: FSuLe "K", SuLei_K
            Case 76: FSuLe "L", SuLei_L
            Case 77: FSuLe "M", SuLei_M
            Case 78: FSuLe "N", SuLei_N
            Case 79: FSuLe "O", SuLei_O
            Case 80: FSuLe "P", SuLei_P
            Case 81: FSuLe "Q", SuLei_Q
            Case 82: FSuLe "R", SuLei_R
            Case 83: FSuLe "S", SuLei_S
            Case 84: FSuLe "T", SuLei_T
            Case 85: FSuLe "U", SuLei_U
            Case 86: FSuLe "V", SuLei_V
            Case 87: FSuLe "W", SuLei_W
            Case 88: FSuLe "X", SuLei_X
            Case 89: FSuLe "Y", SuLei_Y
            Case 90: FSuLe "Z", SuLei_Z
            Case 132: FSuLe "Ä", SuLei_Ä
            Case 142: FSuLe "Ä", SuLei_Ä
            Case 129: FSuLe "Ü", SuLei_Ü
            Case 154: FSuLe "Ü", SuLei_Ü
            Case 148: FSuLe "Ö", SuLei_Ö
            Case 153: FSuLe "Ö", SuLei_Ö
            End Select
        Else
            SMark
        End If
    End If
End If
    
Set RpCo8 = Nothing
    
End Sub
Private Sub repCont8_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    SMark
End Sub

Private Sub repCont8_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim KeySt As String
Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set RpRws = RpCo8.Rows
Set HiTes = RpCo8.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

KeySt = Left$(GlNod, 1)

Select Case HiTes.ht
Case xtpHitTestGroupBox:
Case xtpHitTestHeader:
Case xtpHitTestReportArea:
        If Button = vbRightButton Then
            If KeySt = "A" Then
                CmAcs(KM_Eint_Diagnose).Enabled = True
            Else
                CmAcs(KM_Eint_Diagnose).Enabled = False
            End If
            SMePo 1
        End If
Case xtpHitTestUnknown:
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing
Set RpRws = Nothing
Set RpCo8 = Nothing

End Sub
Private Sub repCont8_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        If Row.GroupRow = False Then
            KaEdi
        End If
    End If
End Sub

Private Sub repCont9_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Select Case Item.Index
Case 0:
    Metrics.Text = SeAry(0, Row.Index)
Case 1:
    Metrics.Text = SeAry(1, Row.Index)
    Metrics.ItemIcon = IC16_IDCard_Norm
Case 2:
    If CBool(SeAry(4, Row.Index)) = False Then
        Metrics.ItemIcon = IC16_Check
    Else
        Metrics.ItemIcon = 0
    End If
End Select

End Sub

Private Sub repCont9_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo9 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo9 = Me.repCont9
Set RpRws = RpCo9.Rows
Set RpSel = RpCo9.SelectedRows
Set HiTes = RpCo9.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

If RpRws.Count > 0 Then
    If RpSel.Count > 0 Then
        Select Case HiTes.ht
        Case xtpHitTestGroupBox:
        Case xtpHitTestHeader:
        Case xtpHitTestReportArea:
                If Button = vbLeftButton Then
                    Set RpRow = RpCo9.HitTest(x, y).Row
                    If RpRow.GroupRow = False Then
                        If CBool(SeAry(4, RpRow.Index)) = True Then
                            SeAry(4, RpRow.Index) = False
                        Else
                            SeAry(4, RpRow.Index) = True
                        End If
                    End If
                End If
        Case xtpHitTestUnknown:
        End Select
    End If
End If

Set RpSel = Nothing
Set RpCo9 = Nothing

End Sub
Private Sub repContK_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        MauDo = True
        SKrDo
        MauDo = False
    End If
End Sub

Private Sub repContK_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String
Dim TmpTg As String
Dim RpCoK As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCoK = FM.repContK

If GlAkt = False Then
    If Item.Tag <> vbNullString Then
        TmpTg = Item.Tag
        TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
        Item.Tag = "@" & TmTag
        FKran Column.ItemIndex
        Item.Tag = TmpTg
    End If
End If

End Sub

Private Sub repContK_ValueChanging(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem, NewValue As Variant, Cancel As Boolean)
On Error Resume Next

If GlAkt = False Then
    If GlBut = RibTab_Abrechnung Then
        If Row.GroupRow = False Then
            If Column.Index = Kra_Anz Then
                If Item.Value <> vbNullString Then
                    If IsNumeric(Item.Value) = True Then
                        If NewValue <> vbNullString Then
                            If IsNumeric(NewValue) = True Then
                                If K_RePr(CInt(Item.Value), CInt(NewValue)) = False Then
                                    Cancel = True
                                End If
                                If K_BePr(CSng(NewValue)) < 0 Then
                                    Cancel = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub repContT_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim ManNr As Long
Dim IdGes As Boolean

ManNr = CLng(Item.Record(2).Value)
IdGes = Not CBool(Item.Checked)

Man_Set IdGes, ManNr

End Sub
Private Sub shtCut01_ClientSizeChanged()
On Error Resume Next

Dim ClRe As RECT

If GlAkt = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    SPosi
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub

Private Sub TabCont1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If GlSta = False Then
        FTxTb
    End If
End Sub

Private Sub TexCont1_Error(Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next

Set TxCoN = Me.TexCont1

If Number = 321 Then
    Select Case LCase(GlTxU)
    Case "doc": TxCoN.Load GlTxF, , 13 'Filname für Textcontrol Error
    Case "docx": TxCoN.Load GlTxF, , 9
    End Select
End If

CancelDisplay = True

End Sub
Private Sub TexCont1_FieldChanged(ByVal FieldId As Integer)
On Error Resume Next

Set TxCoN = Me.TexCont1

If GlAkt = False Then
    TxCoN.FieldCurrent = FieldId
    SPopu "Datenfeld geändert", "Das Datenfeld " & TxCoN.FieldText & " wurde verändert", IC48_Warning
End If

End Sub
Private Sub TexCont1_FieldClicked(ByVal FieldId As Integer)
On Error Resume Next

Set TxCoN = Me.TexCont1

If GlAkt = False Then
    TxCoN.FieldCurrent = FieldId
    If FieldId > 0 Then
        STxDa TxCoN.FieldText
    End If
End If

End Sub
Private Sub TexCont1_FieldDeleted(ByVal FieldId As Integer)
On Error Resume Next

Set TxCoN = Me.TexCont1

If GlAkt = False Then
    TxCoN.FieldCurrent = FieldId
    If FieldId > 0 Then
        TxCoN.FieldCurrent = FieldId
        SPopu "Datenfeld gelöscht", "Das Datenfeld " & TxCoN.FieldText & " wurde gelöscht", IC48_Warning
    End If
End If

End Sub
Private Sub TexCont1_FieldEntered(ByVal FieldId As Integer)
On Error Resume Next

Set TxCoN = Me.TexCont1

If GlAkt = False Then
    TxCoN.FieldCurrent = FieldId
    If FieldId > 0 Then
        STxDa TxCoN.FieldText
    End If
End If

End Sub
Private Sub TexCont1_FieldLeft(ByVal FieldId As Integer)
    STxDa vbNullString, True
End Sub
Private Sub TexCont1_FieldSetCursor(ByVal FieldId As Integer, MousePointer As Integer)
On Error Resume Next

Set TxCoN = Me.TexCont1

If GlAkt = False Then
    If FieldId > 1 Then
        TxCoN.FieldCurrent = FieldId
        STxDa TxCoN.FieldText
    End If
End If

End Sub

Private Sub TexCont1_GotFocus()
    If GlAkt = False Then
        If GlBut = RibTab_Krankenbla Then
            GlMen = True
            FMenu
        End If
    End If
End Sub
Private Sub TexCont1_HeaderFooterActivated(ByVal HeaderFooter As Integer)
    GlHeA = HeaderFooter
End Sub

Private Sub TexCont1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Set FM = frmMain
Set TxCoN = FM.TexCont1

If Shift = vbCtrlMask Then
    If KeyCode = vbKeyV Then
        KeyCode = 0
        TxCoN.Paste 5
    End If
End If

End Sub
Private Sub TexCont1_KeyPress(KeyAscii As Integer)
On Error Resume Next

Select Case KeyAscii
Case 9: TxPhr = vbNullString
Case 13: TxPhr = vbNullString
Case 32: TxPhr = vbNullString
Case 94: TxPhr = vbNullString
Case 96: TxPhr = vbNullString
Case 180: TxPhr = vbNullString
Case Else: 'FTxPh Chr$(KeyAscii)
End Select

End Sub

Private Sub TexCont1_ObjectClicked(ByVal ObjectId As Integer)
On Error Resume Next

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set TxCoN = FM.TexCont1
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

If ObjectId > 0 Then
    CmAcs(Tex_EinTex).Enabled = True
Else
    CmAcs(Tex_EinTex).Enabled = False
End If

Set CmAcs = Nothing
Set CmBrs = Nothing

End Sub
Private Sub TexCont1_PosChange()
    If GlAkt = False Then
        If Me.TexCont1.Text <> vbNullString Then
            GlTSV = True 'Speichern Textverarbeitung
            STxFo
        End If
    End If
End Sub

Private Sub trvList3_AfterLabelEdit(Cancel As Integer, NewString As String)
    If GlAkt = False Then
        SMGrU 2, NewString
    End If
End Sub
Private Sub trvList5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo LaErr

Set TrLi5 = Me.trvList5

Screen.MousePointer = vbHourglass
DoEvents

If GlAkt = False Then
    If Button = vbRightButton Then
        Set TrLi5.SelectedItem = TrLi5.HitTest(x, y)
        SMePo 1
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub trvList5_NodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error GoTo LaErr

Set TrLi5 = Me.trvList5

If GlAkt = False Then
    If Node.Index = 1 Then
        FTrMa True
    Else
        SSuFe
    End If
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub tskDialo_ButtonClicked(ByVal id As Long, CloseDialog As Variant)
    GlMes = id
End Sub
Private Sub tskDialo_RadioButtonClicked(ByVal id As Long)
    GlMso = True
End Sub

Private Sub tskDialo_VerificationClicked(ByVal Checked As Boolean)
    GlMso = Checked
End Sub
Private Sub txtAnzal_GotFocus()
    Me.txtAnzal.SelStart = 0
    Me.txtAnzal.SelLength = Len(Me.txtAnzal.Text)
End Sub
Private Sub txtAnzal_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.txtMulti.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        If GlAkt = False Then
            SKrEi
            K_Eing 3
        End If
    End If
End Sub

Private Sub txtBiKom_KeyPress(KeyAscii As Integer)
    GlSav = True
End Sub
Private Sub txtDaBis_LostFocus()
On Error Resume Next

Dim NeuDa As Date
Dim TxDaB As XtremeSuiteControls.FlatEdit

Set TxDaB = Me.txtDaBis

If IsDate(TxDaB.Text) Then
    NeuDa = TxDaB.Text
    TxDaB.Text = NeuDa
End If

End Sub

Private Sub txtDaNeu_LostFocus()
On Error Resume Next

Dim NeuDa As Date
Dim TxDaN As XtremeSuiteControls.FlatEdit

Set TxDaN = Me.txtDaNeu

If IsDate(TxDaN.Text) Then
    NeuDa = TxDaN.Text
    TxDaN.Text = NeuDa
End If

End Sub

Private Sub txtDatu2_LostFocus()
    If GlAkt = False Then
        KalWa = 2
        FDaK6
    End If
End Sub

Private Sub txtDatu3_LostFocus()
    If GlAkt = False Then
        KalWa = 3
        FDaK6
    End If
End Sub
Private Sub txtDaVon_LostFocus()
On Error Resume Next

Dim NeuDa As Date
Dim TxDaV As XtremeSuiteControls.FlatEdit

Set TxDaV = Me.txtDaVon

If IsDate(TxDaV.Text) = True Then
    NeuDa = TxDaV.Text
    TxDaV.Text = NeuDa
End If

End Sub
Private Sub txtDeta0_KeyPress(KeyAscii As Integer)
    GlSav = True
End Sub
Private Sub txtDeta1_KeyPress(KeyAscii As Integer)
    GlSav = True
End Sub
Private Sub txtDeta2_KeyPress(KeyAscii As Integer)
    GlSav = True
End Sub
Private Sub txtDeta3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtDeta6_KeyPress(KeyAscii As Integer)
    GlSav = True
End Sub
Private Sub txtDeta6_LostFocus()
On Error Resume Next

Dim Frage As Integer
Dim Mld1, Tit1 As String

Tit1 = "Diagnose Speichern"
Mld1 = "Soll die geänderte Diagnose jetzt gespeichert werden?"

If GlAkt = False Then
    If GlSav = True Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            S_KrDs
        End If
    End If
End If

End Sub

Private Sub txtDeta8_KeyPress(KeyAscii As Integer)
    GlSav = True
End Sub

Private Sub txtDeta8_LostFocus()
On Error Resume Next

Dim Frage As Integer
Dim Mld1, Tit1 As String

Tit1 = "Einträge Speichern"
Mld1 = "Soll die geänderten Einträge jetzt gespeichert werden?"

If GlAkt = False Then
    If GlSav = True Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            S_KrSv
        End If
    End If
End If

End Sub

Private Sub txtEinze_GotFocus()
    Me.txtEinze.SelStart = 0
    Me.txtEinze.SelLength = Len(Me.txtEinze.Text)
End Sub

Private Sub txtEinze_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        If KeyCode = vbKeyTab Then
            Me.cmbMitar.SetFocus
        ElseIf KeyCode = vbKeyReturn Then
            SKrEi
            K_Eing 3
        End If
    End If
End Sub

Private Sub txtGesBr_GotFocus()
    Me.txtGesBr.SelStart = 0
    Me.txtGesBr.SelLength = Len(Me.txtGesBr.Text)
End Sub

Private Sub txtGesBr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtGesBr_LostFocus()
On Error Resume Next

Dim SteSa As Single

If Me.txtGesBr.Text <> vbNullString Then
    If IsNumeric(Me.txtGesBr.Text) Then
        SteSa = CSng(Me.txtGesBr.Text)
        If SteSa > 25 Then
            SteSa = 25
        End If
        Me.txtGesBr.Text = Format$(SteSa, GlWa1)
    Else
        Me.txtGesBr.Text = GlWa2
    End If
Else
    Me.txtGesBr.Text = GlWa2
End If

End Sub
Private Sub txtMulti_GotFocus()
    Me.txtMulti.SelStart = 0
    Me.txtMulti.SelLength = Len(Me.txtMulti.Text)
End Sub
Private Sub txtMulti_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        If KeyCode = vbKeyTab Then
            Me.txtEinze.SetFocus
        ElseIf KeyCode = vbKeyReturn Then
            SKrEi
            K_Eing 3
        End If
    End If
End Sub

Private Sub txtReTex_KeyPress(KeyAscii As Integer)
On Error Resume Next

Dim RetWe As Long

RetWe = SendMessage(Me.txtReTex.hwnd, EM_GETLINECOUNT, -1, 0&)

If KeyAscii = 13 Then
    If RetWe >= GlMxZ Then
        KeyAscii = 0
    End If
End If

End Sub
Private Sub txtReTex_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        FDrop
    End If
End Sub

Private Sub txtSumme_Change()
On Error Resume Next

If GlAkt = False Then
    If IsNumeric(Me.txtSumme.Text) = True Then
        If CDbl(Me.txtSumme.Text) < 2147483647 Then
            Me.txtBetra.Text = WinZaUm(CLng(Me.txtSumme.Text))
        End If
    End If
End If

End Sub
Private Sub txtSumme_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtSumme_LostFocus()
On Error Resume Next

Dim GeSum As Single

If Me.txtSumme.Text <> vbNullString Then
    If IsNumeric(Me.txtSumme.Text) Then
        GeSum = CSng(Me.txtSumme.Text)
        If GeSum > 250 Then
            GeSum = 250
            SPopu "Betragsobergrenze überschritten", "Die Betragsobergrenze wurde überschritten und automatisch korrigiert", IC48_Information
        End If
        Me.txtSumme.Text = Format$(GeSum, GlWa1)
    Else
        Me.txtSumme.Text = GlWa2
    End If
Else
    Me.txtSumme.Text = GlWa2
End If

End Sub

Private Sub FDat4()
On Error GoTo OrErr

Dim NeuDa As Date
Dim DaChk As Date

Set DaPi4 = Me.dtpDatu4

DaChk = DateAdd("yyyy", -10, Date)

If DaPi4.Selection.BlocksCount > 0 Then
    If IsDate(DaPi4.Selection.Blocks(0).DateBegin) = True Then
        NeuDa = DaPi4.Selection.Blocks(0).DateBegin
        If NeuDa < DaChk Then
            NeuDa = Date
            With DaPi4
                .EnsureVisible NeuDa - 30
                .Select NeuDa
                .SelectRange NeuDa, NeuDa
            End With
        End If
        ReDim Preserve GlTag(1)
        GlTag(1) = NeuDa
        If GlPop = True Then
            S_AbTa
            S_AbDo
        End If
    End If
End If

Set DaPi4 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDat4 " & Err.Number
Resume Next

End Sub
Private Sub FDat6()
On Error GoTo OrErr

Dim NeuDa As Date
Dim DaChk As Date

Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
Set DaPi6 = Me.dtpDatu6
Set OpZei = Me.optZeit4

DaChk = DateAdd("yyyy", -10, Date)

NeuDa = DaPi6.Selection.Blocks(0).DateBegin

If NeuDa < DaChk Then
    NeuDa = Date
    With DaPi6
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
End If

Select Case KalWa
Case 2: TxDa2.Text = NeuDa
        TxDa3.Text = NeuDa
        TxDa2.SetFocus
Case 3: TxDa3.Text = NeuDa
        TxDa3.SetFocus
End Select

OpZei.Value = True

DoEvents
FDiZe

Set DaPi6 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDat6 " & Err.Number
Resume Next

End Sub
Private Sub FDiDa()
On Error GoTo OrErr
'Diagnosedatum Ändern

GlDDa = True

frmKraDa.Show vbModal

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDiDa " & Err.Number
Resume Next

End Sub
Private Sub FDiEx()
On Error GoTo OrErr
'Statistik Exportieren

Dim FiNam As String

Set FM = frmMain
Set CoDia = FM.comDialo
Set ChCon = FM.chrCont1
Set ChCnt = ChCon.Content

Set clFil = New clsFile

FiNam = SUmw(ChCnt.Titles.Item(0).Text)
FiNam = Replace(FiNam, Chr$(32), "_", 1)
FiNam = Replace(FiNam, "___", "_", 1)
FiNam = FiNam & ".png"
FiNam = Replace(FiNam, Chr$(58), 1, 1)

With CoDia
    .CancelError = True
    .DialogStyle = 1
    .DialogTitle = "Bitte Name und Ordner der Datei angeben"
    .DefaultExt = "*.png"
    .FileName = GlEPf & FiNam
    .Filter = "Portable Network Graphics (.png)|*.png|Alle Dateien (*.*)|*.*"
    .InitDir = GlEPf
    .ShowSave
    FiNam = .FileName
    If .FileTitle = vbNullString Then
        Set ChCnt = Nothing
        Set CoDia = Nothing
        Set clFil = Nothing
        Exit Sub
    End If
End With

If Not IsNull(FiNam) And Not FiNam = vbNullString Then
    If Right$(FiNam, 4) <> ".png" Then
        FiNam = FiNam & ".png"
    End If
    ChCon.SaveAsImage FiNam, ChCon.Width / Screen.TwipsPerPixelX, ChCon.Height / Screen.TwipsPerPixelY
End If

Set ChCnt = Nothing
Set CoDia = Nothing
Set clFil = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDiEx " & Err.Number
Resume Next

End Sub
Private Sub FDiOp(ByVal Flag As Integer)
On Error GoTo OrErr
'Statistik Optionen

Dim ZeVer As Boolean
Dim ChJah As XtremeSuiteControls.CheckBox
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCoS As XtremeCommandBars.CommandBarComboBox
Dim CmCoT As XtremeCommandBars.CommandBarComboBox
Dim CmCoC As XtremeCommandBars.CommandBarComboBox
Dim CmCoL As XtremeCommandBars.CommandBarComboBox
Dim CmCoH As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01
Set ChCon = FM.chrCont1
Set ChJah = FM.chkJahre
Set ChCnt = ChCon.Content

Set CmCoS = CmBrs.FindControl(CmCoS, Sta_Auswa, , True)
Set CmCoT = CmBrs.FindControl(CmCoT, Sta_CmMan, , True)
Set CmCoC = CmBrs.FindControl(CmCoC, Sta_ChOp1, , True)
Set CmCoL = CmBrs.FindControl(CmCoL, Sta_ChOp2, , True)
Set CmCoH = CmBrs.FindControl(CmCoH, Sta_ChOp3, , True)

If ChJah.Enabled = True Then
    If ChJah.Value = xtpChecked Then
        ZeVer = True
    Else
        ZeVer = False
    End If
Else
    ZeVer = False
End If

If GlAkt = False Then
    Select Case Flag
    Case 1:
        ChCnt.Appearance.SetPalette GlPal(CmCoC.ListIndex)
        IniSetVal "Layout", "StaFar", CmCoC.ListIndex
    Case 2:
        For Each ChSrs In ChCnt.Series
            Select Case GlDiT
            Case 1:
                Set ChBar = ChSrs.Style
                With ChBar
                    .Border.Visible = False
                    .SideBySide = True
                    Select Case GlDiS
                    Case 1: .ColorEach = Not ZeVer
                    Case 1: .ColorEach = Not ZeVer
                    Case 8: .ColorEach = Not ZeVer
                    Case 11: .ColorEach = Not ZeVer
                    Case 13: .ColorEach = Not ZeVer
                    Case 14: .ColorEach = Not ZeVer
                    End Select
                End With
                With ChBar.Label
                    .Font.StdFont.Name = GlTFt.Name
                    .Font.StdFont.Bold = False
                    If CmCoL.ListIndex = 3 Then
                        .Visible = False
                    Else
                        .Visible = True
                        .Position = CmCoL.ListIndex - 1
                    End If
                End With
                ChSrs.Style.Label.Antialiasing = False
            Case 2:
                Set ChBar = ChSrs.Style
                With ChBar
                    .Border.Visible = False
                    .SideBySide = True
                    Select Case GlDiS
                    Case 1: .ColorEach = Not ZeVer
                    Case 1: .ColorEach = Not ZeVer
                    Case 8: .ColorEach = Not ZeVer
                    Case 11: .ColorEach = Not ZeVer
                    Case 13: .ColorEach = Not ZeVer
                    Case 14: .ColorEach = Not ZeVer
                    End Select
                End With
                With ChBar.Label
                    .Font.StdFont.Name = GlTFt.Name
                    .Font.StdFont.Bold = False
                    If CmCoL.ListIndex = 3 Then
                        .Visible = False
                    Else
                        .Visible = True
                        .Position = CmCoL.ListIndex - 1
                    End If
                End With
                With ChSrs
                    .Diagram.Rotated = True
                    .Style.Label.Antialiasing = False
                End With
            Case 3:
                Set ChPie = ChSrs.Style
                With ChPie
                    .ColorEach = True
                    .ExplodedDistancePercent = 20
                    .HolePercent = 0
                    .Label.Antialiasing = False
                    .Label.Font.StdFont.Name = GlTFt.Name
                    .Label.Font.StdFont.Bold = False
                    .Rotation = 0
                End With
                ChSrs.Points(0).Special = True 'Hervorhebung
            Case 4:
                Set ChLin = ChSrs.Style
                With ChLin
                    .Label.Font.StdFont.Name = GlTFt.Name
                    .Label.Font.StdFont.Bold = False
                End With
                If CmCoL.ListIndex = 3 Then
                    ChSrs.Style.Label.Visible = False
                Else
                    ChSrs.Style.Label.Visible = True
                End If
                With ChSrs
                    .Style.Marker.SIZE = 10
                    .Style.Label.Antialiasing = False
                End With
            Case 5:
                Set ChAre = ChSrs.Style
                With ChAre
                    .Label.Font.StdFont.Name = GlTFt.Name
                    .Label.Font.StdFont.Bold = False
                End With
                If CmCoL.ListIndex = 3 Then
                    ChSrs.Style.Label.Visible = False
                Else
                    ChSrs.Style.Label.Visible = True
                End If
                With ChSrs
                    .Style.Marker.SIZE = 10
                    .Style.Label.Antialiasing = False
                End With
            Case 6:
                Set ChPyr = ChSrs.Style
                With ChPyr
                    .ColorEach = True
                    .PointDistance = 5
                    .Transparency = 230
                    .Label.Antialiasing = False
                    .Label.Font.StdFont.Name = GlTFt.Name
                    .Label.Font.StdFont.Bold = False
                End With
            End Select
        Next ChSrs
                
        IniSetVal "Layout", "StaCap", CmCoL.ListIndex
    Case 3:
        ChCnt.Appearance.SetAppearance GlHin(CmCoH.ListIndex)
        IniSetVal "Layout", "StaHin", CmCoH.ListIndex
    End Select

    Set CmCoS = Nothing
    Set CmCoT = Nothing
    Set CmCoC = Nothing
    Set CmCoL = Nothing
    Set CmCoH = Nothing
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDiOp " & Err.Number
Resume Next

End Sub
Private Sub FClos()
On Error GoTo OrErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim DocPa As XtremeDockingPane.DockingPane

If GlApp = True Then 'AppMode
    Exit Sub
End If

Set FM = frmMain
Set DocPa = FM.dcpDoc01
Set LiFld = FM.fldView1
Set LiFi2 = FM.filView2
Set LiFi1 = FM.filView1

DocPa.FindPane(PA_DP_Task).GetClientRect ClLin, ClObn, ClBre, ClHoh

DocPa.FindPane(PA_DP_Star).Closed = True
DocPa.FindPane(PA_DP_View).Closed = True
DocPa.FindPane(PA_DP_Unte).Closed = True
DocPa.FindPane(PA_DP_Task).Closed = True
DocPa.FindPane(PA_DP_Top1).Closed = True
DocPa.FindPane(PA_DP_Top2).Closed = True
DocPa.FindPane(PA_DP_Top3).Closed = True
DocPa.FindPane(PA_DP_Text).Closed = True
DocPa.FindPane(PA_DP_KaDE).Closed = True
DocPa.FindPane(PA_DP_KaGE).Closed = True
DocPa.FindPane(PA_DP_KaME).Closed = True
DocPa.FindPane(PA_DP_KaBE).Closed = True
DocPa.FindPane(PA_DP_KaAR).Closed = True
DocPa.FindPane(PA_DP_KaLE).Closed = True
DocPa.FindPane(PA_DP_KaTE).Closed = True
DocPa.FindPane(PA_DP_KaTD).Closed = True
DocPa.FindPane(PA_DP_KaAE).Closed = True
DocPa.FindPane(PA_DP_KaRE).Closed = True
DocPa.FindPane(PA_DP_KaLP).Closed = True
DocPa.FindPane(PA_DP_KaKD).Closed = True
DocPa.FindPane(PA_DP_Bild).Closed = True
DocPa.FindPane(PA_DP_KaDT).Closed = True
DocPa.FindPane(PA_DP_KaKM).Closed = True
DocPa.FindPane(PA_DP_KaTW).Closed = True
DocPa.FindPane(PA_DP_Buch).Closed = True
DocPa.FindPane(PA_DP_Rech).Closed = True
DocPa.FindPane(PA_DP_Bank).Closed = True
DocPa.FindPane(PA_DP_BaVo).Closed = True
DoEvents

If GlRDP = False Then
    Set clFen = New clsFenster
    clFen.hwnd = Me.hwnd

    If GlRes = False Then 'Reset der Einstellungen
        If GlIdi = False Then 'Idiotenmodus
            clFen.FenSav
            If clFen.FeSta <> 2 Then
                IniSetVal "GUI", "FenMax", clFen.FeSta
            End If
            If clFen.FeBre > 640 Then
                If clFen.FeHoh > 450 Then
                    If clFen.FeSta = 0 Then
                        IniSetVal "GUI", "FenLin", clFen.FeLin
                        IniSetVal "GUI", "FenObe", clFen.FeObn
                        IniSetVal "GUI", "FenBre", clFen.FeBre
                        IniSetVal "GUI", "FenHoh", clFen.FeHoh
                    End If
                End If
            End If
        End If
    End If
    
    Set clFen = Nothing
End If

Set LiFld = Nothing
Set LiFi2 = Nothing
Set LiFi1 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FDaK3()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim SetDa As Date
Dim DayFi As Date
Dim DayLa As Date
Dim DaChk As Date
Dim AnzTa As Integer

Set FM = frmMain
Set TxDa1 = FM.txtDatu1
Set DaPi4 = FM.dtpDatu4

DaChk = DateAdd("yyyy", -10, Date)

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

If NeuDa < DaChk Then
    NeuDa = Date
End If

DayFi = NeuDa - 30
DayLa = NeuDa + 90

With DaPi4
    .RedrawControl
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    S_AbTe DayFi, DayLa
    .Left = TxDa1.Left
    .Top = TxDa1.Top + TxDa1.Height
    If .ShowModal(2, 2) Then
        AnzTa = .Selection.BlocksCount
        SetDa = .Selection.Blocks(0).DateBegin
        If AnzTa > 0 Then
            TxDa1.Text = SetDa
        End If
    End If
End With

Set DaPi4 = Nothing

If SetDa = "00:00:00" Then
    SetDa = Date
End If

FDaK5 SetDa

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaK3 " & Err.Number
Resume Next

End Sub

Private Sub FDaK4()
On Error GoTo LaErr
'Kontrolliert und formatiert das Eingabedatum neu

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmBr1 As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit

Set FM = frmMain
Set TxDa1 = FM.txtDatu1
Set DaPi4 = Me.dtpDatu4

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) Then
        NeuDa = CDate(TxDa1.Text)
        TxDa1.Text = NeuDa
        With DaPi4
            .EnsureVisible NeuDa - 30
            .Select NeuDa
            .SelectRange NeuDa, NeuDa
        End With
        GlTag(1) = NeuDa
        
        FDaK5 NeuDa
        DoEvents

        If NeuDa > Date Then
            SPopu NeuDa & " liegt in der Zukunft!", "Der Behandlungstag " & NeuDa & " liegt in der Zukunft", IC48_Information
        End If
    End If
End If

Set DaPi4 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaK4 " & Err.Number
Resume Next

End Sub
Private Sub FDaK5(ByVal NeuDa As Date)
On Error GoTo LaErr
'Kontrolliert und formatiert das Eingabedatum neu

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmBr1 As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit

Set FM = frmMain
Set TxDa1 = FM.txtDatu1
Set DaPi4 = FM.dtpDatu4

TxDa1.Text = NeuDa
With DaPi4
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .EnsureVisible NeuDa - 30
End With

Set FM = frmKatBE
Set DaPi4 = FM.dtpDatu1
With DaPi4
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .EnsureVisible NeuDa - 30
End With
Set CmBrs = FM.comBar02
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)
CmEdt.Text = NeuDa

Set FM = frmKatDE
Set DaPi4 = FM.dtpDatu1
With DaPi4
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .EnsureVisible NeuDa - 30
End With
Set CmBrs = FM.comBar02
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)
CmEdt.Text = NeuDa

Set FM = frmKatGE
Set DaPi4 = FM.dtpDatu1
With DaPi4
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .EnsureVisible NeuDa - 30
End With
Set CmBrs = FM.comBar02
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)
CmEdt.Text = NeuDa

Set FM = frmKatLE
Set DaPi4 = FM.dtpDatu1
With DaPi4
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .EnsureVisible NeuDa - 30
End With
Set CmBrs = FM.comBar02
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)
CmEdt.Text = NeuDa

Set FM = frmKatME
Set DaPi4 = FM.dtpDatu1
With DaPi4
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .EnsureVisible NeuDa - 30
End With
Set CmBrs = FM.comBar02
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)
CmEdt.Text = NeuDa

Set FM = frmKatAR
Set DaPi4 = FM.dtpDatu1
With DaPi4
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .EnsureVisible NeuDa - 30
End With
Set CmBrs = FM.comBar02
Set CmEdt = CmBrs.FindControl(CmEdt, KA_Kalen, , True)
CmEdt.Text = NeuDa

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaK5 " & Err.Number
Resume Next

End Sub
Private Sub FDaK6()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
Set DaPi6 = Me.dtpDatu6

Select Case KalWa
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

With DaPi6
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaK6 " & Err.Number
Resume Next

End Sub
Private Sub FDiPr()
On Error GoTo OrErr
'Statistik Drucken

PrtMain 1

Exit Sub

Set FM = frmMain
Set ChCon = FM.chrCont1

With ChCon
    .PrintPreview
End With

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDiPr " & Err.Number
Resume Next

End Sub

Private Sub FDiSt()
On Error GoTo OrErr
'Statistik wählen

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCoS As XtremeCommandBars.CommandBarComboBox
Dim CmCoT As XtremeCommandBars.CommandBarComboBox
Dim CmCoC As XtremeCommandBars.CommandBarComboBox
Dim CmCoL As XtremeCommandBars.CommandBarComboBox
Dim CmCoH As XtremeCommandBars.CommandBarComboBox
Dim ChJah As XtremeSuiteControls.CheckBox

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmbJv = FM.cmbVgJah
Set ChJah = FM.chkJahre

Set CmCoS = CmBrs.FindControl(CmCoS, Sta_Auswa, , True)
Set CmCoT = CmBrs.FindControl(CmCoT, Sta_CmMan, , True)
Set CmCoC = CmBrs.FindControl(CmCoC, Sta_ChOp1, , True)
Set CmCoL = CmBrs.FindControl(CmCoL, Sta_ChOp2, , True)
Set CmCoH = CmBrs.FindControl(CmCoH, Sta_ChOp3, , True)

If GlAkt = False Then
    GlDiS = CmCoS.ListIndex 'Statistikauswahl

    IniSetVal "System", "StaMod", "S" & Format$(GlDiS, "00")
    DoEvents

    SStaSt
End If

Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDiSt " & Err.Number
Resume Next

End Sub

Private Sub FDiTy(ByVal DiaTy As Integer)
On Error GoTo OrErr
'Diagrammtyp wählen

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

If GlAkt = False Then
    Screen.MousePointer = vbHourglass

    CmAcs(Sta_ChCol).Checked = False
    CmAcs(Sta_ChBar).Checked = False
    CmAcs(Sta_ChPie).Checked = False
    CmAcs(Sta_ChLin).Checked = False
    CmAcs(Sta_ChAre).Checked = False
    CmAcs(Sta_ChDon).Checked = False

    GlDiT = DiaTy
    
    Select Case DiaTy
    Case 1: CmAcs(Sta_ChCol).Checked = True
    Case 2: CmAcs(Sta_ChBar).Checked = True
    Case 3: CmAcs(Sta_ChPie).Checked = True
    Case 4: CmAcs(Sta_ChLin).Checked = True
    Case 5: CmAcs(Sta_ChAre).Checked = True
    Case 6: CmAcs(Sta_ChDon).Checked = True
    End Select
    
    DoEvents
    Sta_Lad
    
    Screen.MousePointer = vbNormal
End If

Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDiTy " & Err.Number
Resume Next

End Sub

Private Sub FDiZe()
On Error GoTo OrErr
'Diagramm Zeitraumwechsel

If GlAkt = False Then
    Screen.MousePointer = vbHourglass

    Sta_Lad
    
    Screen.MousePointer = vbNormal
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDiZe " & Err.Number
Resume Next

End Sub
Private Sub FDoBu(ByVal EiBuc As Integer)
On Error GoTo OrErr
'einfache Buchhaltung verwenden

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Screen.MousePointer = vbHourglass

Select Case EiBuc
Case 1:
    GlBuc = True
    CmAcs(SY_BU_Buchung_BuchEinfach).Checked = True
    CmAcs(SY_BU_Buchung_BuchDoppelt).Checked = False
Case 2:
    GlBuc = False
    CmAcs(SY_BU_Buchung_BuchEinfach).Checked = False
    CmAcs(SY_BU_Buchung_BuchDoppelt).Checked = True
End Select

IniSetVal "System", "EiBuch", GlBuc
S_SeSe 77, , , , GlBuc
DoEvents

STrNo
SSpLa
S_List

Screen.MousePointer = vbNormal

Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDoBu " & Err.Number
Resume Next

End Sub

Private Sub FKrTy()
On Error GoTo OrErr

Dim CoIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmTyp = FM.cmbTypen
Set CmZif = FM.cmbZiffe
Set CmBez = FM.cmbBezei
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

CoIdx = CmTyp.ListIndex + 1

If GlEin <> CoIdx Then 'Eingabetyp
    GlEin = CoIdx
    CmZif.Clear
    CmBez.Clear
    Select Case CoIdx
    Case 1:
        S_KrCo
        CmBez.SetFocus
    Case 6:
        If CmAcs(SY_AB_Abrech_Zahlung).Enabled = True Then frmAnzahl.Show vbModal
    Case Else:
        S_KrCo
        CmZif.SetFocus
    End Select
Else
    Select Case CoIdx
    Case 1:
        CmBez.SetFocus
    Case 6:
        If CmAcs(SY_AB_Abrech_Zahlung).Enabled = True Then frmAnzahl.Show vbModal
    Case Else:
        CmZif.SetFocus
    End Select
End If

Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKrTy " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    If GlAkt = False Then FDaK3
End Sub
Private Sub cmbBezei_GotFocus()
    If GlAkt = False Then Me.cmbBezei.Text = vbNullString
End Sub
Private Sub cmbBezei_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        If KeyCode = vbKeyReturn Then
            K_Eing 2
        ElseIf KeyCode = vbKeyTab Then
            Me.txtAnzal.SetFocus
        End If
    End If
End Sub
Private Sub cmbTypen_Click()
    If GlSta = False Then
        If GlAkt = False Then
            FKrTy
        End If
     End If
End Sub
Private Sub cmbTypen_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
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
        Case vbKeyTab: Me.cmbZiffe.SetFocus
        End Select
    End If
End Sub
Private Sub cmbZiffe_GotFocus()
    If GlAkt = False Then Me.cmbZiffe.Text = vbNullString
    If GlAkt = False Then Me.cmbBezei.Text = vbNullString
End Sub
Private Sub cmbZiffe_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        If KeyCode = vbKeyReturn Then
            K_Eing 1
        ElseIf KeyCode = vbKeyTab Then
            Me.cmbBezei.SetFocus
        End If
    End If
End Sub
Private Sub dtpDatu4_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
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
Private Sub dtpDatu4_MonthChanged()

Dim DayFi As Date
Dim DayLa As Date

Set DaPi4 = Me.dtpDatu4

With DaPi4
    DayFi = .FirstVisibleDay
    DayLa = .LastVisibleDay
End With

If GlAkt = False Then
    'S_AbTe DayFi, DayLa
End If

Set DaPi4 = Nothing

End Sub
Private Sub dtpDatu4_SelectionChanged()
    If GlAkt = False Then FDat4
End Sub

Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu2_GotFocus()
    Me.txtDatu2.SelStart = 0
    Me.txtDatu2.SelLength = Len(Me.txtDatu2.Text)
End Sub

Private Sub txtDatu3_GotFocus()
    Me.txtDatu3.SelStart = 0
    Me.txtDatu3.SelLength = Len(Me.txtDatu3.Text)
End Sub

Private Sub txtDatu1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        Me.cmbTypen.SetFocus
    End If
End Sub
Private Sub txtDatu1_LostFocus()
    If GlAkt = False Then FDaK4
End Sub
Private Sub UpDown1_DownClick()

Dim AltDa As Date
Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text
NeuDa = DateAdd("d", -1, AltDa)

TxDa1.Text = NeuDa
DoEvents
FDaK5 NeuDa

End Sub
Private Sub UpDown1_UpClick()

Dim AltDa As Date
Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1

AltDa = TxDa1.Text
NeuDa = DateAdd("d", 1, AltDa)

TxDa1.Text = NeuDa
DoEvents
FDaK5 NeuDa

End Sub

Private Sub repCont5_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        Select Case GlBut
        Case RibTab_LabBericht: SList
        Case RibTab_LabAuftrag: SList
        Case RibTab_Rezeptmodul: SList
        Case RibTab_Belegmodul: SList
        Case RibTab_Fragebogen: SList
        End Select
        SMark
    End If
End Sub
Private Sub repCont5_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo LaErr

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo5 = Me.repCont5
Set RpRws = RpCo5.Rows
Set HiTes = RpCo5.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

Select Case HiTes.ht
Case xtpHitTestGroupBox:
Case xtpHitTestHeader:
Case xtpHitTestReportArea:
        Select Case GlBut
        Case RibTab_LabBericht:
            SList
            SMark
            If Button = vbRightButton Then
                SMePo 1
            End If
        Case RibTab_LabAuftrag:
            SList
            SMark
            If Button = vbRightButton Then
                SMePo 1
            End If
        Case RibTab_Rezeptmodul:
            SList
            SMark
            If Button = vbRightButton Then
                SMePo 1
            End If
        Case RibTab_Belegmodul:
            SList
            SMark
            If Button = vbRightButton Then
                SMePo 1
            End If
        Case RibTab_Fragebogen:
            SList
            SMark
            If Button = vbRightButton Then
                SMePo 1
            End If
        End Select
Case xtpHitTestUnknown:

End Select

Set RpRws = Nothing
Set RpCo5 = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub repCont3_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

If GlSta = False Then
    If Row.GroupRow = False Then
        Select Case Row.Record(Rec_Type).Value
        Case "M": Metrics.ForeColor = 16744448
        Case "L": Metrics.ForeColor = 33023
        Case "V": Metrics.ForeColor = 8421631
        Case "I": Metrics.ForeColor = 13138080
        Case "U": Metrics.ForeColor = 6604830
        Case Else:
            If CBool(Row.Record(Rec_Selekt).Value) = False Then
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnunge
            End If
        End Select
        If CBool(Row.Record(Rec_Selekt).Value) = False Then
            Metrics.Font.Bold = True
        End If
        If Row.Record(Rec_Storniert).Value = True Then
            Metrics.Font.Strikethrough = True
            Metrics.ForeColor = 8421504
        End If
    End If
End If

End Sub
Private Sub repCont3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            If GlAkt = False Then
                Select Case GlAdO
                Case 0: SReZe
                Case 1: SReZe
                Case 2: frmReEdit.Show vbModal
                End Select
            End If
        End If
    End If
End Sub
Private Sub repCont3_KeyUp(KeyCode As Integer, Shift As Integer)
    If GlAkt = False Then
        SList
        SMark
    End If
End Sub
Private Sub repCont3_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo LaErr

Dim HiRow As XtremeReportControl.ReportRow
Dim HiCol As XtremeReportControl.ReportColumn
Dim HiItm As XtremeReportControl.ReportRecordItem
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows

Set RpCo3 = Me.repCont3
Set RpRws = RpCo3.Rows
Set HiTes = RpCo3.HitTest(x, y)
Set HiRow = HiTes.Row
Set HiItm = HiTes.Item
Set HiCol = HiTes.Column

If RpRws.Count > 0 Then
    Select Case HiTes.ht
    Case xtpHitTestGroupBox:
    Case xtpHitTestHeader:
    Case xtpHitTestReportArea:
            SList
            SMark
            If Button = vbRightButton Then
                SMePo 1
            End If
    Case xtpHitTestUnknown:
    End Select
End If

Set RpRws = Nothing
Set RpCo3 = Nothing
    
LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub repCont3_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then
        If Row.GroupRow = False Then
            frmReEdit.Show vbModal
        End If
    End If
End Sub
Private Sub txtDeta6_GotFocus()
On Error Resume Next

Dim DocPa As XtremeDockingPane.DockingPane

Set FM = frmMain
Set DocPa = FM.dcpDoc01

DocPa.FindPane(PA_DP_Top1).Selected = True

End Sub
Private Sub shtCut01_SelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)
    If GlAkt = False Then
        STaSe Item.id, 0
    End If
End Sub
Private Sub trvList1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If GlAkt = False Then SGrUm 2, NewString
End Sub

Private Sub trvList1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo LaErr

Set TrLi1 = Me.trvList1

Screen.MousePointer = vbHourglass
DoEvents

If GlAkt = False Then
    If TrLi1.Nodes.Count > 0 Then
        If Button = vbRightButton Then
            GlPaK = TrLi1.HitTest(x, y).Key
            Set TrLi1.SelectedItem = TrLi1.HitTest(x, y)
            DoEvents
            SMePo 1
        Else
            If Not TrLi1.HitTest(x, y) Is Nothing Then
                If TrLi1.HitTest(x, y).Selected = True Then
                    If TrLi1.HitTest(x, y).Key <> GlPaK Then
                        SAdNo
                    End If
                End If
            End If
        End If
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

Exit Sub

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub trvList1_OLEDragDrop(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error GoTo LaErr

Set TrLi1 = Me.trvList1

If TrLi1.Nodes.Count > 0 Then
    If Not TrLi1.HitTest(x, y) Is Nothing Then
        With TrLi1
            .HitTest(x, y).Selected = True
            .DropHighlight = Nothing
        End With
    End If
End If

For Each Knote In TrLi1.Nodes
    If Knote.Selected = True Then
        GlPaK = Knote.Key
        GlSuP = GlSuX
        GrAd_Ord GlPaK, Effect
        DoEvents
        SAdNo
        Exit For
    End If
Next Knote

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub trvList1_OLEDragOver(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal State As Integer)
On Error GoTo LaErr

Set TrLi1 = Me.trvList1

If TrLi1.Nodes.Count > 0 Then
    Set TrLi1.DropHighlight = TrLi1.HitTest(x, y)
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub trvList2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo LaErr

Dim GrpNr As Long
Dim KeySt As String
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set TrLi2 = FM.trvList2
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

Screen.MousePointer = vbHourglass
DoEvents

If GlAkt = False Then
    If Button = vbRightButton Then
        If TrLi2.Nodes.Count > 0 Then
            GlNod = TrLi2.HitTest(x, y).Key
            GlNoT = TrLi2.HitTest(x, y).Text
            Set TrLi2.SelectedItem = TrLi2.HitTest(x, y)
            KeySt = Left$(GlNod, 1)
            GrpNr = Mid$(GlNod, 2, Len(GlNod) - 1)
        End If
        If GrpNr > 0 Then
            CmAcs(KA_Eint_Hinzufuegen).Enabled = True
            CmAcs(KA_Eint_Bearbeiten).Enabled = True
            CmAcs(KA_Eint_Kopieren).Enabled = True
            CmAcs(KA_Eint_Loeschen).Enabled = True
            'CmAcs(KA_Eint_Suchen).Enabled = True
            CmAcs(KA_Eint_Drucken).Enabled = True
            CmAcs(KA_Kett_Hinzufuegen).Enabled = True
            CmAcs(KA_Kett_Bearbeiten).Enabled = True
            CmAcs(KA_Kett_Kopieren).Enabled = True
            CmAcs(KA_Kett_Loeschen).Enabled = True
            CmAcs(KA_Kett_Suchen).Enabled = True
            CmAcs(KA_Kett_Drucken).Enabled = True
            CmAcs(KA_Kat_Hinzufuegen).Enabled = True
            CmAcs(KA_Kat_Loeschen).Enabled = True
            CmAcs(KA_Kat_Umbenennen).Enabled = True
            CmAcs(KA_Eint_Favoriten).Enabled = True
            CmAcs(KA_Eint_Export).Enabled = True
            CmAcs(KA_Eint_Import).Enabled = True
            SMePo 2
        Else
            CmAcs(KA_Eint_Hinzufuegen).Enabled = False
            CmAcs(KA_Eint_Bearbeiten).Enabled = False
            CmAcs(KA_Eint_Kopieren).Enabled = False
            CmAcs(KA_Eint_Loeschen).Enabled = False
            'CmAcs(KA_Eint_Suchen).Enabled = False
            CmAcs(KA_Eint_Drucken).Enabled = False
            CmAcs(KA_Kett_Hinzufuegen).Enabled = False
            CmAcs(KA_Kett_Bearbeiten).Enabled = False
            CmAcs(KA_Kett_Kopieren).Enabled = False
            CmAcs(KA_Kett_Loeschen).Enabled = False
            CmAcs(KA_Kett_Suchen).Enabled = False
            CmAcs(KA_Kett_Drucken).Enabled = False
            CmAcs(KA_Kat_Hinzufuegen).Enabled = True
            CmAcs(KA_Kat_Loeschen).Enabled = False
            CmAcs(KA_Kat_Umbenennen).Enabled = False
            CmAcs(KA_Eint_Favoriten).Enabled = False
            CmAcs(KA_Eint_Export).Enabled = False
            CmAcs(KA_Eint_Import).Enabled = False
            SMePo 3
        End If
    Else
        If TrLi2.Nodes.Count > 0 Then
            If Not TrLi2.HitTest(x, y) Is Nothing Then
                If TrLi2.HitTest(x, y).Selected = True Then
                    GlNod = TrLi2.SelectedItem.Key
                    GlNoT = TrLi2.SelectedItem.Text
                    KeySt = Left$(GlNod, 1)
                    GrpNr = Mid$(GlNod, 2, Len(GlNod) - 1)
                    If GrpNr > 0 Then
                        CmAcs(KA_Eint_Hinzufuegen).Enabled = True
                        CmAcs(KA_Eint_Bearbeiten).Enabled = True
                        CmAcs(KA_Eint_Kopieren).Enabled = True
                        CmAcs(KA_Eint_Loeschen).Enabled = True
                        'CmAcs(KA_Eint_Suchen).Enabled = True
                        CmAcs(KA_Eint_Drucken).Enabled = True
                        CmAcs(KA_Kett_Hinzufuegen).Enabled = True
                        CmAcs(KA_Kett_Bearbeiten).Enabled = True
                        CmAcs(KA_Kett_Kopieren).Enabled = True
                        CmAcs(KA_Kett_Loeschen).Enabled = True
                        CmAcs(KA_Kett_Suchen).Enabled = True
                        CmAcs(KA_Kett_Drucken).Enabled = True
                        CmAcs(KA_Kat_Hinzufuegen).Enabled = True
                        CmAcs(KA_Kat_Loeschen).Enabled = True
                        CmAcs(KA_Kat_Umbenennen).Enabled = True
                        CmAcs(KA_Eint_Favoriten).Enabled = True
                        CmAcs(KA_Eint_Export).Enabled = True
                        CmAcs(KA_Eint_Import).Enabled = True
                    Else
                        CmAcs(KA_Eint_Hinzufuegen).Enabled = False
                        CmAcs(KA_Eint_Bearbeiten).Enabled = False
                        CmAcs(KA_Eint_Kopieren).Enabled = False
                        CmAcs(KA_Eint_Loeschen).Enabled = False
                        'CmAcs(KA_Eint_Suchen).Enabled = False
                        CmAcs(KA_Eint_Drucken).Enabled = False
                        CmAcs(KA_Kett_Hinzufuegen).Enabled = False
                        CmAcs(KA_Kett_Bearbeiten).Enabled = False
                        CmAcs(KA_Kett_Kopieren).Enabled = False
                        CmAcs(KA_Kett_Loeschen).Enabled = False
                        CmAcs(KA_Kett_Suchen).Enabled = False
                        CmAcs(KA_Kett_Drucken).Enabled = False
                        CmAcs(KA_Kat_Hinzufuegen).Enabled = True
                        CmAcs(KA_Kat_Loeschen).Enabled = False
                        CmAcs(KA_Kat_Umbenennen).Enabled = False
                        CmAcs(KA_Eint_Favoriten).Enabled = False
                        CmAcs(KA_Eint_Export).Enabled = False
                        CmAcs(KA_Eint_Import).Enabled = False
                    End If
                    DoEvents
                    KList
                End If
            End If
        End If
    End If
Else
    TrLi2.SelectedItem.Image = IC16_Folder_Open
End If

DoEvents
Screen.MousePointer = vbNormal

Set CmBrs = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub trvList2_OLEDragDrop(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error GoTo AdErr
    
Dim Katal As String

Set TrLi2 = Me.trvList2
    
Katal = Left$(GlNod, 1)

If TrLi2.Nodes.Count > 0 Then
    Select Case Katal
    Case "C":
        With TrLi2
            .HitTest(x, y).Selected = True
            .DropHighlight = Nothing
        End With
        K_Kop2
        GlNod = TrLi2.HitTest(x, y).Key
        KList
    Case "I":
        With TrLi2
            .HitTest(x, y).Selected = True
            .DropHighlight = Nothing
        End With
        K_Kop2
        GlNod = TrLi2.HitTest(x, y).Key
        KList
    End Select
End If
    
Exit Sub

AdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "DragDrop " & Err.Number
Resume Next

End Sub
Private Sub trvList2_OLEDragOver(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal State As Integer)
On Error GoTo LaErr

Dim Katal As String

Set TrLi2 = Me.trvList2
    
Katal = Left$(GlNod, 1)

If TrLi2.Nodes.Count > 0 Then
    Select Case Katal
    Case "C":
        Set TrLi2.DropHighlight = TrLi2.HitTest(x, y)
        Set TrLi2.SelectedItem = TrLi2.HitTest(x, y)
    Case "I":
        Set TrLi2.DropHighlight = TrLi2.HitTest(x, y)
        Set TrLi2.SelectedItem = TrLi2.HitTest(x, y)
    End Select
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub

Private Sub dtpDatu7_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
On Error Resume Next

Dim AktTa As Long
Dim AktZa As Integer

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

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

End Sub
Private Sub picBild1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo 5
        End If
    End If
End Sub


Private Sub picBild1_DblClick()
    SPhot
End Sub

Private Sub txtDeta2_LostFocus()
On Error Resume Next

Dim Frage As Integer
Dim Mld1, Tit1 As String

Tit1 = "Eintrag Speichern"
Mld1 = "Soll der geänderte Text jetzt gespeichert werden?"

If GlAkt = False Then
    If GlSav = True Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            If GlBut = RibTab_Abrechnung Then
                S_KrDs
            Else
                S_KrSv
            End If
        End If
    End If
End If

End Sub
Private Sub txtSumme_GotFocus()
    Me.txtSumme.SelStart = 0
    Me.txtSumme.SelLength = Len(Me.txtSumme.Text)
End Sub
Private Sub lstView2_AfterLabelEdit(Cancel As Integer, NewString As String)
    If GlAkt = False Then SGrUm 4, NewString
End Sub
Private Sub lstView2_DblClick()
    If GlAkt = False Then SGrUm 3
End Sub
Private Sub lstView2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
        Case vbKeyF12: SGrUm 3
        Case vbKeyDelete: Dia_Lo
        Case vbKeyBack: Dia_Lo
        Case 119:
        Case 17:
        End Select
    End If
End Sub
Private Sub lstView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GlAkt = False Then
        If Button = vbRightButton Then
            SMePo 3
        End If
    End If
End Sub
Private Sub txtDeta0_LostFocus()
On Error Resume Next

Dim Frage As Integer
Dim Mld1, Tit1 As String

Tit1 = "Arzneimittel Speichern"
Mld1 = "Sollen die geänderten Arzneimittel jetzt gespeichert werden?"

If GlAkt = False Then
    If GlSav = True Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            S_KrSv
        End If
    End If
End If

End Sub
Private Sub txtDeta1_LostFocus()
On Error Resume Next

Dim Frage As Integer
Dim Mld1, Tit1 As String

Tit1 = "Einträge Speichern"
Mld1 = "Sollen die geänderten Einträge jetzt gespeichert werden?"

If GlAkt = False Then
    If GlSav = True Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            S_KrSv
        End If
    End If
End If

End Sub
Private Sub trvList3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo LaErr

Set TrLi3 = Me.trvList3

Screen.MousePointer = vbHourglass
DoEvents

If GlAkt = False Then
    If TrLi3.Nodes.Count > 0 Then
        If Button = vbRightButton Then
            GlMaK = TrLi3.HitTest(x, y).Key
            Set TrLi3.SelectedItem = TrLi3.HitTest(x, y)
            DoEvents
            SMePo 1
        Else
            Set Knote = TrLi3.HitTest(x, y)
            If Not Knote Is Nothing Then
                If Knote.Selected = True Then
                    If Knote.Key <> GlMaK Then 'Aktueller Email-Nodekey
                        SMaNo
                    End If
                End If
            End If
        End If
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

Exit Sub

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub


Private Sub trvList3_OLEDragDrop(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error GoTo LaErr
'Ordnet den Emails eine Gruppe per OLEDrag zu

Dim RowNr As Long
Dim NeGru As String

Set TrLi3 = Me.trvList3

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
DoEvents

clFen.FenDsk 2
DoEvents

If TrLi3.Nodes.Count > 0 Then
    With TrLi3
        If Not .HitTest(x, y) Is Nothing Then
            .HitTest(x, y).Selected = True
            .DropHighlight = Nothing
        End If
    End With
End If
DoEvents

For Each Knote In TrLi3.Nodes
    If Knote.Key <> "P800" Then
        If Knote.Selected = True Then
            NeGru = Knote.Key
            Exit For
        End If
    End If
Next Knote
DoEvents

If NeGru <> vbNullString Then
    GlSuI = GlSuX 'Suchkriterien Emails
    RowNr = GrMa_Ord(NeGru)
    DoEvents
    SUpMa RowNr
End If

DoEvents
clFen.FenDsk 3

DoEvents
Screen.MousePointer = vbNormal

Set clFen = Nothing

Set FM = Nothing

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub trvList3_OLEDragOver(ByVal data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal State As Integer)
On Error GoTo LaErr

Set TrLi3 = Me.trvList3

If TrLi3.Nodes.Count > 0 Then
    Set TrLi3.DropHighlight = TrLi3.HitTest(x, y)
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
