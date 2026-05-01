VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Object = "{5B44EC52-B95B-45CF-98FF-A49DFEED5A92}#16.3#0"; "Codejock.PropertyGrid.v16.3.1.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Einstellungen"
   ClientHeight    =   9900
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9285
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   9285
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   4095
      Left            =   480
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   7695
      _Version        =   1048579
      _ExtentX        =   13573
      _ExtentY        =   7223
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont3 
         Height          =   1575
         Left            =   1560
         TabIndex        =   14
         Top             =   2160
         Width           =   4575
         _Version        =   1048579
         _ExtentX        =   8070
         _ExtentY        =   2778
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFeOrt 
         Height          =   350
         Left            =   3010
         TabIndex        =   15
         Top             =   1600
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtFePLZ 
         Height          =   350
         Left            =   1560
         TabIndex        =   16
         Top             =   1600
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":74F2
         Height          =   800
         Left            =   700
         TabIndex        =   18
         Top             =   400
         Width           =   5500
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PLZ / Ort :"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   1650
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm5 
      Height          =   5415
      Left            =   8400
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   6855
      _Version        =   1048579
      _ExtentX        =   12091
      _ExtentY        =   9551
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremePropertyGrid.PropertyGrid prpGrid2 
         Height          =   3375
         Left            =   400
         TabIndex        =   19
         Top             =   350
         Width           =   2535
         _Version        =   1048579
         _ExtentX        =   4471
         _ExtentY        =   5953
         _StockProps     =   68
         ToolBarVisible  =   0   'False
         HelpVisible     =   -1  'True
         PropertySort    =   0
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4545
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   6555
      _Version        =   1048579
      _ExtentX        =   11562
      _ExtentY        =   8017
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont4 
         Height          =   1695
         Left            =   1680
         TabIndex        =   9
         Top             =   2280
         Width           =   4455
         _Version        =   1048579
         _ExtentX        =   7858
         _ExtentY        =   2990
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFeBnk 
         Height          =   310
         Left            =   3010
         TabIndex        =   10
         Top             =   1600
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFeBLZ 
         Height          =   310
         Left            =   1560
         TabIndex        =   11
         Top             =   1600
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "BLZ / Bank :"
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   1650
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":75EA
         Height          =   800
         Left            =   700
         TabIndex        =   12
         Top             =   400
         Width           =   5500
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   2055
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
      _Version        =   1048579
      _ExtentX        =   7011
      _ExtentY        =   3625
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont2 
         Height          =   1095
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _Version        =   1048579
         _ExtentX        =   2566
         _ExtentY        =   1931
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   1455
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
         _Version        =   1048579
         _ExtentX        =   2990
         _ExtentY        =   2566
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         HideSelection   =   0   'False
         BackColor       =   16777215
         ForeColor       =   4473924
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3255
      _Version        =   1048579
      _ExtentX        =   5741
      _ExtentY        =   7435
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremePropertyGrid.PropertyGrid prpGrid1 
         Height          =   3375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2535
         _Version        =   1048579
         _ExtentX        =   4471
         _ExtentY        =   5953
         _StockProps     =   68
         ToolBarVisible  =   0   'False
         HelpVisible     =   -1  'True
         PropertySort    =   0
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   14000
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
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   480
      Top             =   0
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
Attribute VB_Name = "frmOptions"
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
Private TxPLZ As XtremeSuiteControls.FlatEdit
Private TxOrt As XtremeSuiteControls.FlatEdit
Private TxBLZ As XtremeSuiteControls.FlatEdit
Private TxBnk As XtremeSuiteControls.FlatEdit
Private txVorn As XtremeSuiteControls.FlatEdit
Private txName As XtremeSuiteControls.FlatEdit
Private txStra As XtremeSuiteControls.FlatEdit
Private txPost As XtremeSuiteControls.FlatEdit
Private txOrte As XtremeSuiteControls.FlatEdit
Private txTele As XtremeSuiteControls.FlatEdit
Private txFaxe As XtremeSuiteControls.FlatEdit
Private txBank As XtremeSuiteControls.FlatEdit
Private txBaLZ As XtremeSuiteControls.FlatEdit
Private txKont As XtremeSuiteControls.FlatEdit
Private txSteu As XtremeSuiteControls.FlatEdit
Private txIKNr As XtremeSuiteControls.FlatEdit
Private txBeru As XtremeSuiteControls.FlatEdit
Private txTite As XtremeSuiteControls.FlatEdit
Private cmFach As XtremeSuiteControls.ComboBox
Private txtFa01 As XtremeSuiteControls.FlatEdit
Private txtFa02 As XtremeSuiteControls.FlatEdit
Private txtFa03 As XtremeSuiteControls.FlatEdit
Private txtFa04 As XtremeSuiteControls.FlatEdit
Private txtFa05 As XtremeSuiteControls.FlatEdit
Private txtFa06 As XtremeSuiteControls.FlatEdit
Private txtFa07 As XtremeSuiteControls.FlatEdit
Private txtFa08 As XtremeSuiteControls.FlatEdit
Private txtFa09 As XtremeSuiteControls.FlatEdit
Private txtFa10 As XtremeSuiteControls.FlatEdit
Private txtFa11 As XtremeSuiteControls.FlatEdit
Private txtFa12 As XtremeSuiteControls.FlatEdit
Private txtFa13 As XtremeSuiteControls.FlatEdit
Private txtFa14 As XtremeSuiteControls.FlatEdit
Private txtFa15 As XtremeSuiteControls.FlatEdit
Private txtFa16 As XtremeSuiteControls.FlatEdit
Private txtFa17 As XtremeSuiteControls.FlatEdit
Private txtFa18 As XtremeSuiteControls.FlatEdit
Private txtFa19 As XtremeSuiteControls.FlatEdit
Private txtFa20 As XtremeSuiteControls.FlatEdit
Private CoDia As XtremeSuiteControls.CommonDialog
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmBat As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmTol As XtremeCommandBars.TabToolBar
Private CmTab As XtremeCommandBars.TabControlItem
Private PrGr1 As XtremePropertyGrid.PropertyGrid
Private PrGr2 As XtremePropertyGrid.PropertyGrid
Private PrKat As XtremePropertyGrid.PropertyGridItem
Private PrItm As XtremePropertyGrid.PropertyGridItem
Private PrIts As XtremePropertyGrid.PropertyGridItems
Private PrBol As XtremePropertyGrid.PropertyGridItemBool
Private PrFnt As XtremePropertyGrid.PropertyGridItemFont
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
Private WithEvents TbBar As XtremeCommandBars.TabToolBar
Attribute TbBar.VB_VarHelpID = -1

Private TmTag As String
Private NoStr As String
Private LiTyp As Integer
Private NeFn1 As Boolean 'Fontänderung
Private NeFn2 As Boolean 'Fontänderung
Private NeFn3 As Boolean 'Fontänderung
Private NeFn4 As Boolean 'Fontänderung
Private NeFn5 As Boolean 'Fontänderung
Private DrNam() As String
Private DrGef As Boolean
Private OpLad As Boolean
Private TaIdx As Integer

Private Const KEYEVENTF_KEYUP = &H2

Private clFil As clsFile
Private clWor As clsWord
Private clFen As clsFenster
Private WithEvents clDru As clsDruck
Attribute clDru.VB_VarHelpID = -1

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
Private Sub FCol()
On Error GoTo ReErr

Dim Farbe As Long
Dim Knots As XtremeSuiteControls.TreeViewNodes
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set TrLi1 = Me.trvList1
Set RpCon = Me.repCont2
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows
Set CoDia = frmMain.comDialo

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If RpRow.Record(3).Value <> vbNullString Then
            Farbe = RpRow.Record(3).Value
        Else
            Farbe = vbBlack
        End If
        With CoDia
            .CancelError = True
            .DialogStyle = 1
            .Color = Farbe
            .DialogTitle = "Bitte wählen Sie die gewünschte Farbe"
            .ShowColor
            RpRow.Record(3).Value = .Color
            RpRow.Record(3).Tag = "@Farbe"
            Opt_Sav TrLi1.SelectedItem.Key
        End With
    End If
End If

Set TrLi1 = Nothing
Set CoDia = Nothing
Set RpSel = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FCol " & Err.Number
Exit Sub

End Sub
Private Function FDial(ByVal TolId As Long) As String
On Error GoTo MeErr

Dim RetWe As Long
Dim FiNam As String

Set CoDia = frmMain.comDialo

Set clFil = New clsFile

Select Case TolId
Case 1201:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        Select Case GlTyp
        Case 2:
            .DefaultExt = "*.dbx"
            .Filter = GlPrg & " Datenbanken (*.dbx)|*.dbx"
        Case 3:
            .DefaultExt = "*.dbv"
            .Filter = GlPrg & " Datenbanken (*.dbv)|*.dbv"
        End Select
        .DialogTitle = "Wo befindet sich die richtige Datenbankdatei?"
        .InitDir = IniGetVal("SysPfa", "DatPfa")
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
Case 2637:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.exe"
        .Filter = "Programm-Dateien (*.exe)|*.exe|Alle Dateien (*.*)|*.*"
        .DialogTitle = "Wo befindet sich die richtige Programmdatei?"
        .InitDir = IniGetVal("SysPfa", "WegPfa")
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
Case 2863:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.exe"
        .Filter = "Programm-Dateien (*.exe)|*.exe|Alle Dateien (*.*)|*.*"
        .DialogTitle = "Wo befindet sich die richtige Programmdatei?"
        .InitDir = IniGetVal("SysPfa", "GDTPrg")
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
Case 2890:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.txm"
        .Filter = "Textverarbeitung (.txm)|*.txm|Alle Dateien (*.*)|*.*"
        .DialogTitle = "Wo befindet sich die richtige Dokumentenvorlage?"
        .InitDir = GlVor
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
    With clFil
        .FilPfa FiNam
        FiNam = .DaNam
    End With
Case 2891:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.txr"
        .Filter = "Langrezeptvorlagen (.txr)|*.txr|Alle Dateien (*.*)|*.*"
        .DialogTitle = "Wo befindet sich die richtige Langrezeptvorlage?"
        .InitDir = GlVor
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
    With clFil
        .FilPfa FiNam
        FiNam = .DaNam
    End With
Case 2892:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.txn"
        .Filter = "Newslettervorlage (.txn)|*.txn|Alle Dateien (*.*)|*.*"
        .DialogTitle = "Wo befindet sich die richtige Newslettervorlage?"
        .InitDir = GlVor
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
    With clFil
        .FilPfa FiNam
        FiNam = .DaNam
    End With
Case 2643:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.exe"
        .Filter = "Programm-Dateien (*.exe)|*.exe|Alle Dateien (*.*)|*.*"
        .DialogTitle = "Wo befindet sich die richtige Programmdatei?"
        .InitDir = IniGetVal("SysPfa", "DatPfa")
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
Case 2642:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DefaultExt = "*.*"
        .Filter = "Alle Dateien (*.*)|*.*"
        .DialogTitle = "Wo befindet sich das richtige Startdokument?"
        .InitDir = IniGetVal("SysPfa", "DatPfa")
        .FileName = vbNullString
        .ShowOpen
        FiNam = .FileName
        If .FileTitle = vbNullString Then Exit Function
    End With
Case 1203:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "DatPfa")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1204:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = GlEig
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1210:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = GlEig
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1205:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "DatPfa")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1206:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "DatPfa")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1207:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "DatPfa")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1208:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "DatPfa")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1213:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "FiltPf")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1214:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "TmpPfa")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 1215:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Ordner?"
        .FileName = IniGetVal("SysPfa", "ForPfa")
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
Case 2968:
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet der TSE Stick?"
        .FileName = CStr(GlSet(1, 94))
        RetWe = .ShowBrowseFolder
        FiNam = .FileName
        If RetWe = 0 Then Exit Function
    End With
End Select

If FiNam <> vbNullString Then
    Select Case TolId
    Case 1201:
    Case 1202:
    Case 2637:
    Case 2863:
    Case 2890:
    Case 2891:
    Case 2892:
    Case 2642:
    Case 2643:
    Case Else: If Right$(FiNam, 1) <> "\" Then FiNam = FiNam & "\"
    End Select
End If

FDial = FiNam

Set clFil = Nothing
Set CoDia = Nothing

Exit Function

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDial " & Err.Number
Exit Function

End Function
Private Sub FFarb()
On Error GoTo ReErr

Dim AktZa As Integer
Dim KatZa As Integer

Set FM = frmOptions
Set PrGr2 = FM.prpGrid2

Set PrKat = PrGr2.AddCategory("Farbenbeschriftungen")
PrKat.id = 1100
PrKat.Expandable = True
PrKat.Expanded = True

For AktZa = 1 To UBound(GlTmF)
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Farbenbeschriftung " & Format$(AktZa, "00"), GlTmF(AktZa, 0))
    PrItm.id = GlTmF(AktZa, 2)
Next AktZa

Set PrKat = PrGr2.AddCategory("Terminfarben")
PrKat.id = 1200
PrKat.Expandable = True
PrKat.Expanded = True

For AktZa = 1 To UBound(GlTmF)
    Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Terminfarbe " & Format$(AktZa, "00"), GlTmF(AktZa, 1))
    PrItm.id = GlTmF(AktZa, 2)
Next AktZa

For KatZa = 1 To UBound(GlTmH)
    Set PrKat = PrGr2.AddCategory("Kalenderhintergrund " & Format$(KatZa, "00"))
    PrKat.id = 1300 + GlTmH(KatZa, 0)
    PrKat.Expandable = True
    PrKat.Expanded = True
    Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Farbe Arbeitsbereich", GlTmH(KatZa, 1))
    PrItm.id = 1
    Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Farbe Freizeitbereich", GlTmH(KatZa, 2))
    PrItm.id = 2
    Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Farbe Kalenderkopf", GlTmH(KatZa, 3))
    PrItm.id = 3
    Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Farbe Kalenderlinie 1", GlTmH(KatZa, 4))
    PrItm.id = 4
    Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Farbe Kalenderlinie 2", GlTmH(KatZa, 5))
    PrItm.id = 5
Next KatZa

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FFarb " & Err.Number
Exit Sub

End Sub
Private Sub FHelp()
On Error Resume Next

Dim TmStr As String
Dim TemTx As String
Dim Knots As XtremeSuiteControls.TreeViewNodes

Set TrLi1 = Me.trvList1
Set PrGr1 = Me.prpGrid1
Set Knots = TrLi1.Nodes
Set PrIts = PrGr1.Categories

TmStr = "klicken Sie bitte einmal mit der linken Maustaste oben links auf den Systembutton und wählen: 'Einstellungen'. Im folgenden Einstellungendialog "

Select Case TaIdx
Case 0:
    For Each PrKat In PrIts
        For Each PrItm In PrKat.Childs
            If PrItm.Selected = True Then
                TemTx = PrItm.Caption
                TmStr = TmStr & "erweiterten Sie bitte den Abschnitt: '" & PrKat.Caption & "' und klicken auf die Option: '" & TemTx & "'."
                Clipboard.Clear
                Clipboard.SetText TmStr
                Exit For
            End If
        Next PrItm
    Next PrKat
Case 1:
    For Each Knote In Knots
        If Knote.Selected = True Then
            TemTx = Knote.Text
            TmStr = TmStr & "klicken Sie bitte oben auf das Register: 'Systemtabellen' und dann links auf: '" & TemTx & "'."
            Clipboard.Clear
            Clipboard.SetText TmStr
            Exit For
        End If
    Next Knote


End Select


End Sub

Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = ""
TeMai = ""
TeInh = ""
TeFus = ""

'SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FInit()
On Error GoTo MeErr

Dim TmFnt As New StdFont
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl
Dim RpPLZ As XtremeReportControl.ReportControl
Dim RpBLZ As XtremeReportControl.ReportControl

Set FM = frmOptions
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5
Set TrLi1 = FM.trvList1
Set RpCon = FM.repCont2
Set RpPLZ = FM.repCont3
Set RpBLZ = FM.repCont4
Set ImMan = frmMain.imgManag

TmFnt.Name = GlTFt.Name
TmFnt.SIZE = GlTFt.SIZE

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = GlKrE
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = True
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False 'Gruppierungen
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
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = TmFnt.Name
    .PaintManager.TextFont.SIZE = TmFnt.SIZE
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
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.FixedRowHeight = False
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = TmFnt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = TmFnt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpPLZ
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = True
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False 'Gruppierungen
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
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = TmFnt.Name
    .PaintManager.TextFont.SIZE = TmFnt.SIZE
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
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.FixedRowHeight = False
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = TmFnt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = TmFnt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpBLZ
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = True
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False 'Gruppierungen
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
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = TmFnt.Name
    .PaintManager.TextFont.SIZE = TmFnt.SIZE
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
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.FixedRowHeight = False
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = TmFnt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = TmFnt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With TrLi1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .Checkboxes = False
    .Font.SIZE = GlTFt.SIZE
    .Font.Name = TmFnt.Name
    .ForeColor = -2147483641
    .FullRowSelect = False
    .HideSelection = False
    .HotTracking = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpTreeViewLabelManual
    .Scroll = True
    .ShowBorder = True
    .ShowLines = xtpTreeViewShowLines
    .ShowPlusMinus = True
    .SingleSel = False
End With

With TrLi1
    Set Knote = .Nodes.Add(, , "K00", "Systemtabellen", IC16_Folder_View)
    Set Knote = .Nodes.Add("K00", 4, "K01", "Adressgruppen", IC16_Folder_Open)
    Set Knote = .Nodes.Add("K00", 4, "K02", "Gebührenkataloge", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K05", "Zahlungsziele", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K06", "Währungen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K03", "Geldkonten", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K07", "Mahnstufen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K08", "Arzneigruppen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K09", "Buchungstexte", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K10", "Steuersätze", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K11", "Terminbetreffs", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K12", "Raumzuordnung", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K13", "Eigene Felder", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K14", "Krankenblatttypen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K15", "Tarifinformationen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K16", "Rechnungskommentare", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K17", "Länder", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K18", "Kalendermarker", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K19", "Fragebogengruppen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K20", "Behinderungsgrad", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K21", "Terminstatus", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K22", "Diagnosegruppen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K23", "Emailgruppen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K24", "OTS-Betreffs", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K25", "OTS-Eigenschaften", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K26", "OTS-Emailbestätigung", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K27", "Terminnachrichtentexte", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K28", "Terminnachrichtenbetreff", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K29", "Emailtextvorlagen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K30", "Emailbetreffvorlagen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K31", "Artikelgruppen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K32", "Kataloggruppen", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K33", "TSE-Klienten", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K34", "Zahlungstexte", IC16_Folder_Close)
    Set Knote = .Nodes.Add("K00", 4, "K35", "Laborgruppen", IC16_Folder_Close)
    .Nodes("K00").Expanded = True
    .Nodes("K01").Selected = True
End With

Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
FM.BackColor = GlBak

Set RpCon = Nothing
Set RpPLZ = Nothing
Set RpBLZ = Nothing
Set ImMan = Nothing

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Function FKon(ByVal KoStr As String, ByVal Flag As Integer) As Variant
On Error Resume Next
'Konvertierung der Registryschlüssel

Select Case Flag
Case 1:
    Select Case KoStr
    Case "0": FKon = False
    Case "-1": FKon = True
    Case "8202": FKon = "Laborgemeinschaft"
    Case "8201": FKon = "Labor-Facharzt-Bericht"
    Case "S0": FKon = "Microsoft SQL Native Client"
    Case "S1": FKon = "Microsoft SQL Server OLEDB"
    Case "S2": FKon = "DBX Datenbank"
    Case "S3": FKon = "DBV Datenbank"
    Case "R1": FKon = "NT-Authentifizierung"
    Case "R2": FKon = "Benutzerrechte"
    Case "X2": FKon = "DOS-ASCII"
    Case "X3": FKon = "Windows-OEM"
    Case "T1": FKon = "Arzt (GOÄ)"
    Case "T2": FKon = "Heilpraktiker (GebüH)"
    Case "T3": FKon = "Heilhilfsberufe"
    Case "F2": FKon = "Jahr-Monat-000000"
    Case "F3": FKon = "JahrMonat-000000"
    Case "F4": FKon = "Jahr-000000"
    Case "F5": FKon = "JahrMonat-0000"
    Case "Z0": FKon = "Emailversand an einen Patienten"
    Case "Z1": FKon = "Emailversand an den Mandanten"
    Case "Z2": FKon = "Emailversand an alle Patienten"
    Case "K3": FKon = "Absteigend"
    Case "K4": FKon = "Aufsteigend"
    Case "W1": FKon = "Word 97/2000"
    Case "W2": FKon = "Word XP/2003"
    Case "W3": FKon = "Word 2007/2010"
    Case "L0": FKon = "Office 2000"
    Case "L1": FKon = "Office XP"
    Case "L2": FKon = "Office 2003"
    Case "L3": FKon = "Windows XP"
    Case "L4": FKon = "Windows Whidbey"
    Case "L5": FKon = "Office 2007 Einfach"
    Case "L6": FKon = "Office 2007 Ribbon"
    Case "P1": FKon = "Office2007Blue"
    Case "P2": FKon = "Office2007Black"
    Case "P3": FKon = "Office2007Silver"
    Case "P4": FKon = "Office2007Aqua"
    Case "V1": FKon = "Interner Viewer"
    Case "V3": FKon = "Externer Viewer"
    Case "U1": FKon = "Windows XP Design"
    Case "U2": FKon = "Office 2000 Design"
    Case "U3": FKon = "Office XP Design"
    Case "U4": FKon = "Office 2003 Design"
    Case "B1": FKon = "Keine"
    Case "B2": FKon = "Abwechselnd"
    Case "B3": FKon = "Verblassen"
    Case "B4": FKon = "Schieben"
    Case "B5": FKon = "Ausbreiten"
    Case "B6": FKon = "Windows Standard"
    Case "D0": FKon = "Abrechnung"
    Case "D1": FKon = "Dokumentation"
    Case "D2": FKon = "Adressenmaske"
    Case "O1": FKon = GlKoR(1, 0)
    Case "O2": FKon = GlKoR(2, 0)
    Case "O3": FKon = GlKoR(3, 0)
    Case "O4": FKon = GlKoR(4, 0)
    Case "O5": FKon = GlKoR(5, 0)
    Case "O6": FKon = GlKoR(6, 0)
    Case "C0": FKon = "Normalansicht"
    Case "C2": FKon = "Seitenansicht"
    Case "C3": FKon = "Fließtext"
    Case "N1": FKon = "Geburtstagsliste"
    Case "N2": FKon = "Terminliste"
    Case "N3": FKon = "Offene-Postenliste"
    Case "N4": FKon = "Aufgabenliste"
    Case "N5": FKon = "Akontoliste"
    Case "J1": FKon = "Standardmandant aus Optionsdialog"
    Case "J2": FKon = "Mandant aus Adresseneingabemaske"
    Case "J3": FKon = "Mandant aus Mitarbeitereingabemaske"
    Case "G00": FKon = GlKvk(0, 0)
    Case "G01": FKon = GlKvk(1, 0)
    Case "G02": FKon = GlKvk(2, 0)
    Case "G03": FKon = GlKvk(3, 0)
    Case "G04": FKon = GlKvk(4, 0)
    Case "G05": FKon = GlKvk(5, 0)
    Case "G06": FKon = GlKvk(6, 0)
    Case "G07": FKon = GlKvk(7, 0)
    Case "G08": FKon = GlKvk(8, 0)
    Case "G09": FKon = GlKvk(9, 0)
    Case "G10": FKon = GlKvk(10, 0)
    Case "G11": FKon = GlKvk(11, 0)
    Case "Q0": FKon = "(Firma), Name, Vorname, Geburtsdatum"
    Case "Q1": FKon = "(Firma), Name, Vorname, Nummer"
    Case "Q2": FKon = "(Firma), Name, Vorname, Ort"
    Case 9101: FKon = "Wordvorlagen"
    Case 9107: FKon = "Worddokumente"
    Case 9108: FKon = "Alle Adressen"
    Case 9109: FKon = "Outlookabgleich"
    Case 9106: FKon = "Serienmailadressen"
    Case "E0": FKon = "Komma"
    Case "E1": FKon = "Punkt"
    Case "I0": FKon = "Privat Inland"
    Case "I1": FKon = "Privat Europa"
    Case "I2": FKon = "Privat Ausland"
    Case "I3": FKon = "Gewerb. Inland"
    Case "I4": FKon = "Gewerb. Europa"
    Case "I5": FKon = "Gewerb. Ausland"
    Case "H1": FKon = "Einfache Druckvorschau"
    Case "H2": FKon = "Erweiterte Druckvorschau"
    Case "H3": FKon = "Externe Druckvorschau"
    Case "Y1": FKon = "Druckervorgabe"
    Case "Y2": FKon = "DIN A4"
    Case "Y3": FKon = "DIN A5"
    Case "Y4": FKon = "DIN A6"
    Case "K1": FKon = "Nachfragen Konvertierung"
    Case "K2": FKon = "Automatische Konvertierung"
    Case "K3": FKon = "Keine Konvertierung"
    Case "A1": FKon = "Variable"
    Case "A2": FKon = "Maximiert"
    Case "A3": FKon = "Vorgegeben"
    Case "M0": FKon = "keine TSE"
    Case "M1": FKon = "SwissBit TSE"
    Case "M2": FKon = "fiskaly TSE"
    Case "M5": FKon = "Diagnosecodeprüfung"
    Case "M6": FKon = "Diagnosetextprüfung"
    Case "M7": FKon = "Keine Diagnoseprüfung"
    Case "#1": FKon = "Postversand"
    Case "#2": FKon = "Emailversand"
    Case "#3": FKon = "Downloadlink"
    Case "R": FKon = "R - Standardrechnung"
    Case "L": FKon = "L - Laborrechnung"
    Case "A": FKon = "A - Abrechnungsstelle"
    Case "M": FKon = "M - Rechnungsauftrag"
    Case "G": FKon = "G - Gewerberechnung"
    Case "I": FKon = "I - Importrechnung"
    End Select
Case 2:
    Select Case KoStr
    Case False: FKon = 0
    Case True: FKon = -1
    Case "Laborgemeinschaft": FKon = "8202"
    Case "Labor-Facharzt-Bericht": FKon = "8201"
    Case "Microsoft SQL Native Client": FKon = "S0"
    Case "Microsoft SQL Server OLEDB": FKon = "S1"
    Case "DBX Datenbank": FKon = "S2"
    Case "DBV Datenbank": FKon = "S3"
    Case "NT-Authentifizierung": FKon = "R1"
    Case "Benutzerrechte": FKon = "R2"
    Case "DOS-ASCII": FKon = "X2"
    Case "Windows-OEM": FKon = "X3"
    Case "Arzt (GOÄ)": FKon = "T1"
    Case "Heilpraktiker (GebüH)": FKon = "T2"
    Case "Heilhilfsberufe": FKon = "T3"
    Case "Jahr-Monat-000000": FKon = "F2"
    Case "JahrMonat-000000": FKon = "F3"
    Case "Jahr-000000": FKon = "F4"
    Case "JahrMonat-0000": FKon = "F5"
    Case "Emailversand an einen Patienten": FKon = "Z0"
    Case "Emailversand an den Mandanten": FKon = "Z1"
    Case "Emailversand an alle Patienten": FKon = "Z2"
    Case "Absteigend": FKon = "K3"
    Case "Aufsteigend": FKon = "K4"
    Case "Word 97/2000": FKon = "W1"
    Case "Word XP/2003": FKon = "W2"
    Case "Word 2007/2010": FKon = "W3"
    Case "Office 2000": FKon = "L0"
    Case "Office XP": FKon = "L1"
    Case "Office 2003": FKon = "L2"
    Case "Windows XP": FKon = "L3"
    Case "Windows Whidbey": FKon = "L4"
    Case "Office 2007 Einfach": FKon = "L5"
    Case "Office 2007 Ribbon": FKon = "L6"
    Case "Office2007Blue": FKon = "P1"
    Case "Office2007Black": FKon = "P2"
    Case "Office2007Silver": FKon = "P3"
    Case "Office2007Aqua": FKon = "P4"
    Case "Interner Viewer": FKon = "V1"
    Case "Externer Viewer": FKon = "V3"
    Case "Windows XP Design": FKon = "U1"
    Case "Office 2000 Design": FKon = "U2"
    Case "Office XP Design": FKon = "U3"
    Case "Office 2003 Design": FKon = "U4"
    Case "Keine": FKon = "B1"
    Case "Abwechselnd": FKon = "B2"
    Case "Verblassen": FKon = "B3"
    Case "Schieben": FKon = "B4"
    Case "Ausbreiten": FKon = "B5"
    Case "Windows Standard": FKon = "B6"
    Case "Abrechnung": FKon = "D0"
    Case "Dokumentation": FKon = "D1"
    Case "Adressenmaske": FKon = "D2"
    Case GlKvk(0, 0): FKon = "G00"
    Case GlKvk(1, 0): FKon = "G01"
    Case GlKvk(2, 0): FKon = "G02"
    Case GlKvk(3, 0): FKon = "G03"
    Case GlKvk(4, 0): FKon = "G04"
    Case GlKvk(5, 0): FKon = "G05"
    Case GlKvk(6, 0): FKon = "G06"
    Case GlKvk(7, 0): FKon = "G07"
    Case GlKvk(8, 0): FKon = "G08"
    Case GlKvk(9, 0): FKon = "G09"
    Case GlKvk(10, 0): FKon = "G10"
    Case GlKvk(11, 0): FKon = "G11"
    Case "(Firma), Name, Vorname, Geburtsdatum": FKon = "Q0"
    Case "(Firma), Name, Vorname, Nummer": FKon = "Q1"
    Case "(Firma), Name, Vorname, Ort": FKon = "Q2"
    Case "Thumbnails": FKon = "V4"
    Case GlKoR(1, 0): FKon = "O1"
    Case GlKoR(2, 0): FKon = "O2"
    Case GlKoR(3, 0): FKon = "O3"
    Case GlKoR(4, 0): FKon = "O4"
    Case GlKoR(5, 0): FKon = "O5"
    Case GlKoR(6, 0): FKon = "O6"
    Case "Normalansicht": FKon = "C0"
    Case "Seitenansicht": FKon = "C2"
    Case "Fließtext": FKon = "C3"
    Case "Geburtstagsliste": FKon = "N1"
    Case "Terminliste": FKon = "N2"
    Case "Offene-Postenliste": FKon = "N3"
    Case "Aufgabenliste": FKon = "N4"
    Case "Akontoliste": FKon = "N5"
    Case "Standardmandant aus Optionsdialog": FKon = "J1"
    Case "Mandant aus Adresseneingabemaske": FKon = "J2"
    Case "Mandant aus Mitarbeitereingabemaske": FKon = "J3"
    Case "Wordvorlagen": FKon = 9101
    Case "Worddokumente": FKon = 9107
    Case "Alle Adressen": FKon = 9108
    Case "Outlookabgleich": FKon = 9109
    Case "Serienmailadressen": FKon = 9106
    Case "Komma": FKon = "E0"
    Case "Punkt": FKon = "E1"
    Case "Privat Inland": FKon = "I0"
    Case "Privat Europa": FKon = "I1"
    Case "Privat Ausland": FKon = "I2"
    Case "Gewerb. Inland": FKon = "I3"
    Case "Gewerb. Europa": FKon = "I4"
    Case "Gewerb. Ausland": FKon = "I5"
    Case "Einfache Druckvorschau": FKon = "H1"
    Case "Erweiterte Druckvorschau": FKon = "H2"
    Case "Externe Druckvorschau": FKon = "H3"
    Case "Druckervorgabe": FKon = "Y1"
    Case "DIN A4": FKon = "Y2"
    Case "DIN A5": FKon = "Y3"
    Case "DIN A6": FKon = "Y4"
    Case "Nachfragen Konvertierung": FKon = "K1"
    Case "Automatische Konvertierung": FKon = "K2"
    Case "Keine Konvertierung": FKon = "K3"
    Case "Variable": FKon = "A1"
    Case "Maximiert": FKon = "A2"
    Case "Vorgegeben": FKon = "A3"
    Case "keine TSE": FKon = "M0"
    Case "SwissBit TSE": FKon = "M1"
    Case "fiskaly TSE": FKon = "M2"
    Case "Diagnosecodeprüfung": FKon = "M5"
    Case "Diagnosetextprüfung": FKon = "M6"
    Case "Keine Diagnoseprüfung": FKon = "M7"
    Case "Postversand": FKon = "#1"
    Case "Emailversand": FKon = "#2"
    Case "Downloadlink": FKon = "#3"
    Case "R - Standardrechnung": FKon = "R"
    Case "L - Laborrechnung": FKon = "L"
    Case "A - Abrechnungsstelle": FKon = "A"
    Case "M - Rechnungsauftrag": FKon = "M"
    Case "G - Gewerberechnung": FKon = "G"
    Case "I - Importrechnung": FKon = "I"
    End Select
End Select

End Function
Private Sub FLoe()
On Error GoTo MeErr

Set FM = frmAdress
Set TxPLZ = Me.txtFePLZ
Set TxOrt = Me.txtFeOrt
Set TxBLZ = Me.txtFeBLZ
Set TxBnk = Me.txtFeBnk

Select Case LiTyp
Case 1:
Case 2:
    Opt_Loe
Case 3:
    If TxBLZ.Text <> vbNullString Then
        Opt_BLo TxBLZ.Text, 1
    ElseIf TxBnk.Text <> vbNullString Then
        Opt_BLo TxBnk.Text, 2
    End If
Case 4:
    If TxPLZ.Text <> vbNullString Then
        Opt_PLo TxPLZ.Text, 1
    ElseIf TxOrt.Text <> vbNullString Then
        Opt_PLo TxOrt.Text, 2
    End If
End Select

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoe " & Err.Number
Resume Next

End Sub
Private Sub FMenu()
On Error GoTo MeErr

Dim TmFnt As New StdFont
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmTab As XtremeCommandBars.TabControlItem
Dim TaPai As XtremeCommandBars.TabPaintManager
Dim ToTab As XtremeCommandBars.TabControlItem
Dim CmAcs As XtremeCommandBars.CommandBarActions

Set FM = frmOptions
Set PrGr1 = FM.prpGrid1
Set PrGr2 = FM.prpGrid2
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag

TmFnt.Name = GlTFt.Name
TmFnt.SIZE = GlTFt.SIZE

With PrGr1
    Select Case GlSty
    Case 8:
        .VisualTheme = xtpGridThemeOffice2013
        .PaintManager.LineColor = GlMoB
        .PaintManager.HighlightBackColor = GlMoB
        .PaintManager.HighlightForeColor = -2147483640
    Case 7:
        .VisualTheme = xtpGridThemeOffice2013
        .PaintManager.LineColor = GlMoB
        .PaintManager.HighlightBackColor = GlMoB
        .PaintManager.HighlightForeColor = -2147483640
    Case Else:
        .VisualTheme = xtpGridThemeResource
        .PaintManager.LineColor = GlMoB
        .PaintManager.HighlightBackColor = GlMoB
        .PaintManager.HighlightForeColor = -2147483640
    End Select
    .Font.Name = TmFnt.Name
    .Font.SIZE = 8
    .BorderStyle = xtpGridBorderNone
    .HelpBackColor = -2147483643
    .HelpForeColor = -2147483640
    .HighlightChangedItems = True
    .HideSelection = False
    .HelpVisible = True
    .LockRedraw = False
    .NavigateItems = True
    .PropertySort = NoSort
    .ShowInplaceButtonsAlways = False
    .ToolBarVisible = False
    .VariableSplitterPos = True
    .ViewBackColor = -2147483643
    .ViewCategoryForeColor = -2147483640
    .ViewForeColor = -2147483640
    .ViewReadOnlyForeColor = 8421504
    .Verbs.Clear
End With

With PrGr2
    Select Case GlSty
    Case 8:
        .VisualTheme = xtpGridThemeOffice2013
        .PaintManager.LineColor = GlMoB
        .PaintManager.HighlightBackColor = GlMoB
        .PaintManager.HighlightForeColor = -2147483640
    Case 7:
        .VisualTheme = xtpGridThemeOffice2013
        .PaintManager.LineColor = GlMoB
        .PaintManager.HighlightBackColor = GlMoB
        .PaintManager.HighlightForeColor = -2147483640
    Case Else:
        .VisualTheme = xtpGridThemeResource
        .PaintManager.LineColor = GlMoB
        .PaintManager.HighlightBackColor = GlMoB
        .PaintManager.HighlightForeColor = -2147483640
    End Select
    .Font.Name = TmFnt.Name
    .Font.SIZE = 8
    .BorderStyle = xtpGridBorderNone
    .HelpBackColor = -2147483643
    .HelpForeColor = -2147483640
    .HighlightChangedItems = True
    .HideSelection = False
    .HelpVisible = False
    .LockRedraw = False
    .NavigateItems = True
    .PropertySort = NoSort
    .ShowInplaceButtonsAlways = False
    .ToolBarVisible = False
    .VariableSplitterPos = True
    .ViewBackColor = -2147483643
    .ViewCategoryForeColor = -2147483640
    .ViewForeColor = -2147483640
    .ViewReadOnlyForeColor = 8421504
    .Verbs.Clear
End With

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmSta
    .Font.SIZE = 8
    Set CmPan = .AddPane(1)
    CmPan.Width = 200
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    .Visible = True
End With

With CmAcs
    Set CmAct = .Add(SY_OP_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_OP_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_OP_Reset, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_OP_Speichern, vbNullString, vbNullString, vbNullString, vbNullString)
End With

Set TbBar = CmBrs.AddTabToolBar("TabBar")

Set ToTab = TbBar.InsertCategory(RibTab_Opti1, "Einstellungen")
With ToTab
    .ToolTip = "Zeigt die Einstellungen des Programms"
    .Visible = True
    .Selected = True
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Einstellungen"
        .IconId = IC24_Doc_Add
        .ToolTipText = "Legt einen neuen Eintrag an"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Einstellungen"
        .IconId = IC24_Doc_Del
        .ToolTipText = "Löscht einen Eintrag"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Einstellungen"
        .BeginGroup = True
        .IconId = IC24_Disk_Norm
        .ToolTipText = "Speichert die Einstellungen"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Standardpfade")
    With CmCon
        .Category = "Einstellungen"
        .IconId = IC24_Folder_Paper
        .BeginGroup = True
        .ToolTipText = "Zurücksetzen der Pfadangaben"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Einstellungen"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Einstellungen"
        .BeginGroup = True
        .IconId = IC24_Exit
        .ToolTipText = "Schließt den Optionsdialog"
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti2, "Systemtabellen")
With ToTab
    .ToolTip = "Zeigt die Grunddaten in Tabellen des Programms"
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Systemtabellen"
        .IconId = IC24_Doc_Add
        .ToolTipText = "Legt einen neuen Eintrag an"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Systemtabellen"
        .IconId = IC24_Doc_Del
        .ToolTipText = "Löscht einen Eintrag"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Systemtabellen"
        .BeginGroup = True
        .IconId = IC24_Disk_Norm
        .ToolTipText = "Speichert die Einstellungen"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Standardpfade")
    With CmCon
        .Category = "Systemtabellen"
        .IconId = IC24_Folder_Paper
        .BeginGroup = True
        .ToolTipText = "Zurücksetzen der Pfadangaben"
        .Enabled = False
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Systemtabellen"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Systemtabellen"
        .BeginGroup = True
        .IconId = IC24_Exit
        .ToolTipText = "Schließt den Optionsdialog"
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti3, "Bankleitzahlen")
With ToTab
    .ToolTip = "Zeigt die Daten des Benutzers"
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Bankleitzahlen"
        .IconId = IC24_Doc_Add
        .ToolTipText = "Legt einen neuen Eintrag an"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Bankleitzahlen"
        .IconId = IC24_Doc_Del
        .ToolTipText = "Löscht einen Eintrag"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Bankleitzahlen"
        .BeginGroup = True
        .IconId = IC24_Disk_Norm
        .ToolTipText = "Speichert die Einstellungen"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Standardpfade")
    With CmCon
        .Category = "Bankleitzahlen"
        .IconId = IC24_Folder_Paper
        .BeginGroup = True
        .ToolTipText = "Zurücksetzen der Pfadangaben"
        .Enabled = False
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Bankleitzahlen"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Bankleitzahlen"
        .BeginGroup = True
        .IconId = IC24_Exit
        .ToolTipText = "Schließt den Optionsdialog"
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti4, "Postleitzahlen")
With ToTab
    .ToolTip = "Zeigt das Postleitzahlenverzeichnis"
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Postleitzahlen"
        .IconId = IC24_Doc_Add
        .ToolTipText = "Legt einen neuen Eintrag an"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Postleitzahlen"
        .IconId = IC24_Doc_Del
        .ToolTipText = "Löscht einen Eintrag"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Postleitzahlen"
        .BeginGroup = True
        .IconId = IC24_Disk_Norm
        .ToolTipText = "Speichert die Einstellungen"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Standardpfade")
    With CmCon
        .Category = "Postleitzahlen"
        .IconId = IC24_Folder_Paper
        .BeginGroup = True
        .ToolTipText = "Zurücksetzen der Pfadangaben"
        .Enabled = False
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Postleitzahlen"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Postleitzahlen"
        .BeginGroup = True
        .IconId = IC24_Exit
        .ToolTipText = "Schließt den Optionsdialog"
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti5, "Kalenderfarben")
With ToTab
    .ToolTip = "Zeigt die Terminfarben und die Hintergrundfarben"
    .Visible = True
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Kalenderfarben"
        .IconId = IC24_Doc_Add
        .ToolTipText = "Legt einen neuen Eintrag an"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Kalenderfarben"
        .IconId = IC24_Doc_Del
        .ToolTipText = "Löscht einen Eintrag"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Kalenderfarben"
        .BeginGroup = True
        .IconId = IC24_Disk_Norm
        .ToolTipText = "Speichert die Einstellungen"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Reset, "Standardpfade")
    With CmCon
        .Category = "Kalenderfarben"
        .IconId = IC24_Folder_Paper
        .BeginGroup = True
        .ToolTipText = "Zurücksetzen der Pfadangaben"
        .Enabled = False
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Kalenderfarben"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Kalenderfarben"
        .BeginGroup = True
        .IconId = IC24_Exit
        .ToolTipText = "Schließt den Optionsdialog"
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
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F11, KY_F11
    .KeyBindings.Add FCONTROL, Asc("R"), KY_CT_AL_R
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

CmAcs(SY_OP_Hinzufuegen).Enabled = False
CmAcs(SY_OP_Loeschen).Enabled = False

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FNode(ByVal NoKey As String)
On Error GoTo NoErr

Dim CmBrs As XtremeCommandBars.CommandBars

Set TrLi1 = Me.trvList1
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Opt_Spl NoKey, True
Opt_Lad NoKey

CmAcs(SY_OP_Hinzufuegen).Enabled = True

Select Case NoKey
Case "K01": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K02": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K03": CmAcs(SY_OP_Loeschen).Enabled = False
Case "K04": CmAcs(SY_OP_Loeschen).Enabled = False
Case "K05": CmAcs(SY_OP_Loeschen).Enabled = True
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K06": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K07": CmAcs(SY_OP_Loeschen).Enabled = False
Case "K08": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K09": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K10": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K11": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K12": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K13": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K14": CmAcs(SY_OP_Loeschen).Enabled = False
Case "K15": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K16": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K17": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K18": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K19": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K20": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K21": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K22": CmAcs(SY_OP_Loeschen).Enabled = False
Case "K23": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K24": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K25": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K26": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K27": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K28": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K29": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K30": CmAcs(SY_OP_Loeschen).Enabled = False
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K31": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K32": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K33": CmAcs(SY_OP_Loeschen).Enabled = True
            CmAcs(SY_OP_Hinzufuegen).Enabled = False
Case "K34": CmAcs(SY_OP_Loeschen).Enabled = True
Case "K35": CmAcs(SY_OP_Loeschen).Enabled = True
End Select

CmSta.Pane(0).Text = TrLi1.SelectedItem.Text

clFen.FenDsk 3
Screen.MousePointer = vbNormal
        
Set clFen = Nothing

Set CmAct = Nothing
Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

NoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNode " & Err.Number
Resume Next

End Sub
Private Sub FPfad()
On Error GoTo ReErr

Dim DaPfa As String
Dim OrdNa As String
Dim Frage As Integer
Dim Mld1, Tit1 As String

Set PrGr1 = Me.prpGrid1
Set PrIts = PrGr1.Categories

Set clFil = New clsFile

Tit1 = "Pfadeinstellungen"
Mld1 = "Sollen die Pfadeinstellungen jetzt auf die Standardwerte zurückgesetzt werden?"

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    DaPfa = IniGetVal("SysPfa", "DatPfa")

    With clFil
        .FilPfa DaPfa
        OrdNa = .DaPfa
    End With
    
    IniSetVal "System", "ForPfa", LCase(OrdNa & "\Formulare\")
    IniSetVal "System", "ImpPfa", LCase(OrdNa & "\Import\")
    IniSetVal "System", "ImpOrd", LCase(OrdNa & "\Import\")
    If GlRDP = True Then
        IniSetVal "System", "ExpPfa", LCase(OrdNa & "\Import\")
        IniSetVal "System", "ExpOrd", LCase(OrdNa & "\Import\")
    Else
        IniSetVal "System", "ExpPfa", LCase(OrdNa & "\Export\")
        IniSetVal "System", "ExpOrd", LCase(OrdNa & "\Export\")
    End If
    IniSetVal "System", "DockPf", LCase(OrdNa & "\Dokumente\")
    IniSetVal "System", "DocPfa", LCase(OrdNa & "\Vorlagen\")
    IniSetVal "System", "TermPf", LCase(OrdNa & "\Termine\")
    IniSetVal "System", "BilPfa", LCase(OrdNa & "\Bilder\")
    IniSetVal "System", "EmalPf", LCase(OrdNa & "\Emails\")
    IniSetVal "System", "BackPf", LCase(OrdNa & "\Backup\")
    IniSetVal "System", "FiltPf", LCase(OrdNa & "\Filter\")
    IniSetVal "System", "TmpPfa", LCase(OrdNa & "\Temp\")

    Set PrItm = PrGr1.FindItem(1215)
    PrItm.Value = LCase(OrdNa & "\Formulare\")
    Set PrItm = PrGr1.FindItem(1204)
    PrItm.Value = LCase(OrdNa & "\Import\")
    Set PrItm = PrGr1.FindItem(1210)
    PrItm.Value = LCase(OrdNa & "\Export\")
    Set PrItm = PrGr1.FindItem(1206)
    PrItm.Value = LCase(OrdNa & "\Dokumente\")
    Set PrItm = PrGr1.FindItem(1208)
    PrItm.Value = LCase(OrdNa & "\Vorlagen\")
    Set PrItm = PrGr1.FindItem(1212)
    PrItm.Value = LCase(OrdNa & "\Termine\")
    Set PrItm = PrGr1.FindItem(1205)
    PrItm.Value = LCase(OrdNa & "\Bilder\")
    Set PrItm = PrGr1.FindItem(1211)
    PrItm.Value = LCase(OrdNa & "\Emails\")
    Set PrItm = PrGr1.FindItem(1203)
    PrItm.Value = LCase(OrdNa & "\Backup\")
    Set PrItm = PrGr1.FindItem(1213)
    PrItm.Value = LCase(OrdNa & "\Filter\")
    Set PrItm = PrGr1.FindItem(1214)
    PrItm.Value = LCase(OrdNa & "\Temp\")

    GlFrO = LCase(OrdNa & "\Formulare\")
    GlIPf = LCase(OrdNa & "\Import\")
    GlImO = LCase(OrdNa & "\Import\")
    If GlRDP = True Then
        GlEPf = LCase(OrdNa & "\Import\")
        GlExO = LCase(OrdNa & "\Import\")
    Else
        GlEPf = LCase(OrdNa & "\Export\")
        GlExO = LCase(OrdNa & "\Export\")
    End If
    GlDox = LCase(OrdNa & "\Dokumente\")
    GlVor = LCase(OrdNa & "\Vorlagen\")
    GlTEx = LCase(OrdNa & "\Termine\")
    GlBPf = LCase(OrdNa & "\Bilder\")
    GlFPf = LCase(OrdNa & "\Filter\")
    GlTmp = LCase(OrdNa & "\Temp\")
End If

Set clFil = Nothing

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPfad " & Err.Number
Exit Sub

End Sub
Private Sub FPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl
Dim RpPLZ As XtremeReportControl.ReportControl
Dim RpBLZ As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set TrLi1 = Me.trvList1
Set PrGr1 = Me.prpGrid1
Set PrGr2 = Me.prpGrid2
Set RpCon = Me.repCont2
Set RpPLZ = Me.repCont3
Set RpBLZ = Me.repCont4
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set Rahm5 = Me.frmRahm5

If Me.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    Rahm1.Move 0, ClObn, ClBre, ClHoh
    Rahm2.Move 0, ClObn, ClBre, ClHoh
    Rahm3.Move 0, ClObn, ClBre, ClHoh
    Rahm4.Move 0, ClObn, ClBre, ClHoh
    Rahm5.Move 0, ClObn, ClBre, ClHoh
    
    PrGr1.Move 10, 10, ClBre - 30, ClHoh - 40
    PrGr2.Move 10, 10, ClBre - 30, ClHoh - 40
    TrLi1.Move 10, 10, 2700, ClHoh - 40
    RpCon.Move 2750, 10, ClBre - 2770, ClHoh - 40
    RpBLZ.Move 10, 2200, ClBre - 30, ClHoh - 2240
    RpPLZ.Move 10, 2200, ClBre - 30, ClHoh - 2240
End If

Set RpCon = Nothing
Set RpPLZ = Nothing
Set RpBLZ = Nothing
Set CmBrs = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FRgLa()
On Error GoTo MeErr

Dim AktZa As Long
Dim TmStr As String
Dim TmKon As String
Dim KeyNa As String
Dim EiKo1 As Boolean
Dim EiKo2 As Boolean
Dim StKo1 As Boolean
Dim TmKo1 As Boolean
Dim TmKo2 As Boolean
Dim StMan As Integer
Dim StMit As Integer
Dim StMio As Integer
Dim StLab As Integer
Dim StRau As Integer
Dim StBri As Integer
Dim TmGe1 As Integer
Dim TmGe2 As Integer
Dim TmGe3 As Integer
Dim TmFnt As New StdFont

Set PrGr1 = Me.prpGrid1
KeyNa = "Optionentexte"

StLab = CInt(GlSet(2, 18))
StMan = CInt(IniGetVal("Vorgabe", "StaMan"))
StMit = CInt(IniGetVal("Vorgabe", "StaMit"))
StMio = CInt(IniGetVal("Vorgabe", "StaMio"))

If StMan < 1 Or StMan > UBound(GlMaA) Then StMan = 1
If StMit < 1 Or StMit > UBound(GlMiA) Then StMit = 1
If StMio < 0 Or StMio > UBound(GlMiA) Then StMio = 0
If StLab < 1 Or StLab > UBound(GlLab) Then StLab = 1

'------------------------------

Set PrKat = PrGr1.AddCategory("Allgemein")
PrKat.id = 1100
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)
PrKat.Expanded = True

If GlStK < 1 Or GlStK > UBound(GlGKa) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Gebührenkatalog*", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Gebührenkatalog*", GlGKa(GlStK, 1))
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlGKa)
    PrItm.Constraints.Add GlGKa(AktZa, 1)
Next AktZa
PrItm.id = 1102
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

If GlSet(2, 1) > UBound(GlKet) Or GlKe1 < 1 Or GlKe1 > UBound(GlKet) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Gebührenkette 1*", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Gebührenkette 1*", GlKet(GlKe1, 2))
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlKet)
    PrItm.Constraints.Add GlKet(AktZa, 2)
Next AktZa
PrItm.id = 1106
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

If GlSet(2, 74) > UBound(GlKet) Or GlKe2 < 1 Or GlKe2 > UBound(GlKet) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Gebührenkette 2*", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Gebührenkette 2*", GlKet(GlKe2, 2))
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlKet)
    PrItm.Constraints.Add GlKet(AktZa, 2)
Next AktZa
PrItm.id = 1714
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

If GlSet(2, 21) < 1 Or GlSet(2, 21) > UBound(GlZah) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Zahlungsziel*", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Zahlungsziel*", GlZah(GlSet(2, 21), 1))
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlZah)
    PrItm.Constraints.Add GlZah(AktZa, 1)
Next AktZa
PrItm.id = 1103
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

If GlSet(2, 64) < 1 Or GlSet(2, 64) > UBound(GlStu) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Steuersatz*", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Steuersatz*", GlStu(GlSet(2, 64), 2))
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlStu)
    PrItm.Constraints.Add GlStu(AktZa, 2)
Next AktZa
PrItm.id = 1104
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

If GlSet(2, 31) < 1 Or GlSet(2, 31) > UBound(GlWar) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Währung*", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Währung*", GlWar(GlSet(2, 31), 1))
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlWar)
    PrItm.Constraints.Add GlWar(AktZa, 1)
Next AktZa
PrItm.id = 1105
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Empfängertyp", FKon(IniGetVal("Vorgabe", "StaEmp"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Privat Inland"
PrItm.Constraints.Add "Privat Europa"
PrItm.Constraints.Add "Privat Ausland"
PrItm.Constraints.Add "Gewerb. Inland"
PrItm.Constraints.Add "Gewerb. Europa"
PrItm.Constraints.Add "Gewerb. Ausland"
PrItm.id = 1107
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

StBri = IniGetVal("Vorgabe", "StaBrf")
If StBri < 1 Or StBri > UBound(GlBri) Then
    StBri = 1
    If UBound(GlBri) < 1 Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Briefanrede", vbNullString)
        GoTo SkipBri
    End If
End If
Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Briefanrede", GlBri(StBri, 0))
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlBri)
    PrItm.Constraints.Add GlBri(AktZa, 0), GlBri(AktZa, 1)
Next AktZa
PrItm.id = 2987
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
SkipBri:

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Dezimaltrennung*", FKon(GlSet(1, 19), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Komma"
PrItm.Constraints.Add "Punkt"
PrItm.id = 2894
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

StRau = IniGetVal("Vorgabe", "StaRau")
If StRau < 1 Or StRau > UBound(GlRmu) Then
    StRau = 1
    If UBound(GlRmu) < 1 Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Behandlungsraum", vbNullString)
        GoTo SkipRau
    End If
End If
Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Behandlungsraum", GlRmu(StRau, 1))
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlRmu)
    PrItm.Constraints.Add GlRmu(AktZa, 1)
Next AktZa
PrItm.id = 1110
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
SkipRau:

'------------------------------

Set PrKat = PrGr1.AddCategory("Pfadangaben")
PrKat.id = 1200
PrKat.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Datenbankdatei", IniGetVal("SysPfa", "DatPfa"))
PrItm.id = 1201
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
If GlTyp < 2 Then PrItm.ReadOnly = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Backupordner", IniGetVal("System", "BackPf"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1203
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Importordner", IniGetVal("System", "ImpPfa"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1204
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Exportordner", IniGetVal("System", "ExpPfa"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1210
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Bilderordner", IniGetVal("System", "BilPfa"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1205
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Emailordner", IniGetVal("System", "EmalPf"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1211
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Termineordner", IniGetVal("System", "TermPf"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1212
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Dokumentenordner", IniGetVal("System", "DockPf"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1206
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Dokumentenvorlagen", IniGetVal("System", "DocPfa"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1208
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Filterordner", IniGetVal("System", "FiltPf"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1213
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Temporärordner", IniGetVal("System", "TmpPfa"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1214
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Formulareordner", IniGetVal("System", "ForPfa"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 1215
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "WEGAMED Programmdatei", IniGetVal("System", "WegPfa"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 2637
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "GDT Programmdatei", IniGetVal("System", "GDTPrg"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 2863
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Startseitendokument", IniGetVal("System", "StaWeb"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 2642
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlRDP

'------------------------------

Set PrKat = PrGr1.AddCategory("Abrechnungsstelle")
PrKat.id = 1400
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Rechnungsexport mit Zeilenumbruch", FKon(IniGetVal("System", "PADZei"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1403
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Rechnungsexport mit Umlautkonvertierung", FKon(IniGetVal("System", "PADUml"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1404
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Rechnungsexport mit benanntem Gebührenkatalog*", CBool(GlSet(4, 43)))
PrBol.CheckBoxStyle = True
PrBol.id = 2825
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Keine Preisberechnung bei IGeL*", CBool(GlSet(4, 44)))
PrBol.CheckBoxStyle = True
PrBol.id = 2876
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Keine Positionskennzeichen bei Medikamenten und Begründungen*", CBool(GlSet(4, 45)))
PrBol.CheckBoxStyle = True
PrBol.id = 2927
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Adressenverwaltung")
PrKat.id = 2700
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Microsoft-Word Version", FKon(IniGetVal("System", "WorTre"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Word 97/2000"
PrItm.Constraints.Add "Word XP/2003"
PrItm.Constraints.Add "Word 2007/2010"
PrItm.id = 1703
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.Hidden = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Aktion bei Doppelklicken einer Adresse", FKon(IniGetVal("System", "AdDoKl"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Abrechnung"
PrItm.Constraints.Add "Dokumentation"
PrItm.Constraints.Add "Adressenmaske"
PrItm.id = 2822
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Format des Adressenverkehrsnames*", FKon(GlSet(1, 8), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "(Firma), Name, Vorname, Geburtsdatum"
PrItm.Constraints.Add "(Firma), Name, Vorname, Nummer"
PrItm.Constraints.Add "(Firma), Name, Vorname, Ort"
PrItm.id = 2850
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Adresseneingabemaske immer im Vordergrund", FKon(IniGetVal("Layout", "AdrVor"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1708
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Outlookabgleich nur gekennzeichneter Adressen", FKon(IniGetVal("System", "OutAdr"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1712
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Outlookkontakte mit Geburtsdatum abgleichen", FKon(IniGetVal("System", "OutGeb"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2619
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Adressgruppen beim Starten expandieren", FKon(IniGetVal("System", "TreExp"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1710
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Telefonnummer mit Landesvorwahl einfügen", FKon(IniGetVal("System", "IntTel"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1715
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Anschrift ohne Anrede erzeugen", FKon(IniGetVal("System", "AnsAnr"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1716
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Bei PLZ Eingabe den Ort automatisch zuweisen", FKon(IniGetVal("System", "OrtAbf"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1719
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Bei PLZ Eingabe den Ortsteil zum einfügen", FKon(IniGetVal("System", "TeiOrt"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1718
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Bei BLZ Eingabe den Banknamen einfügen", FKon(IniGetVal("System", "KurBnk"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1720
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Einzelbriefübergabe mit geschätzten Formularfeldern", FKon(IniGetVal("System", "WorUbe"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1721
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Bei Anschrift Firma in zweiter Zeile anzeigen", FKon(IniGetVal("System", "FirZei"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2623
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Geburtstagliste nur mit gekennzeichneten Adressen", FKon(IniGetVal("System", "AdrGeb"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1723
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Geburtstage zusätzlich in Aufgabenliste anzeigen", FKon(IniGetVal("System", "WieGeb"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2808
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Erweiterter BDT Datenimport", FKon(IniGetVal("System", "BDTImp"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2852
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "TAPI-Rufnummer Übergabe mit Klammern", FKon(IniGetVal("System", "TAPIKl"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1724
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mitarbeiternummer als GDT-Dateiname*", CBool(GlSet(4, 71)))
PrBol.CheckBoxStyle = True
PrBol.id = 2943
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "GDT-Speicherung ohne Speichern-Dialog*", CBool(GlSet(4, 72)))
PrBol.CheckBoxStyle = True
PrBol.id = 2944
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Überschrift des Bemerkungsfeldes", IniGetVal("Layout", "AdTit1"))
PrItm.id = 2624
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Überschrift des Notizenfeldes", IniGetVal("Layout", "AdTit2"))
PrItm.id = 2625
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Angezeigter Name des GDT Programms", IniGetVal("System", "GDTApp"))
PrItm.id = 2865
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Dateiname der GDT Exportdatei*", CStr(GlSet(1, 73)))
PrItm.id = 2866
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Anwenderdaten")
PrKat.id = 1900
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Benutzeranmeldung beim Start*", CBool(GlSet(4, 9)))
PrBol.CheckBoxStyle = True
PrBol.id = 1903
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Separierter Mandanten Rechnungsnummernkreis*", CBool(GlSet(4, 10)))
PrBol.CheckBoxStyle = True
PrBol.id = 2824
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Separierter Mandanten Buchungsnummernkreis*", CBool(GlSet(4, 11)))
PrBol.CheckBoxStyle = True
PrBol.id = 2842
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

If GlRst = True Then  'Mandantenbezogene Datenbegrenzung
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Mandant neue(s) Rechnung/Rezept*", CStr(FKon("J3", 1)))
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Mandant neue(s) Rechnung/Rezept*", CStr(FKon(GlSet(1, 50), 1)))
End If
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Standardmandant aus Optionsdialog"
PrItm.Constraints.Add "Mandant aus Adresseneingabemaske"
PrItm.Constraints.Add "Mandant aus Mitarbeitereingabemaske"
PrItm.id = 2926
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mandantenspalte im Abrechnungsmodul", FKon(IniGetVal("Layout", "ManSpa"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2875
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mandantenbezogene Vorgabenbenutzung*", CBool(GlSet(4, 34)))
PrBol.CheckBoxStyle = True
PrBol.id = 2918
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mandantenbezogene Datenbegrenzung*", CBool(GlSet(4, 65)))
PrBol.CheckBoxStyle = True
PrBol.id = 2940
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

If StMan >= 1 And StMan <= UBound(GlMaA) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Mandant", GlMaA(StMan, 1))
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Mandant", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlMaA)
    PrItm.Constraints.Add GlMaA(AktZa, 1)
Next AktZa
PrItm.id = 2626
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

If StMit >= 1 And StMit <= UBound(GlMiA) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Mitarbeiter", GlMiA(StMit, 1))
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Mitarbeiter", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlMiA)
    PrItm.Constraints.Add GlMiA(AktZa, 1)
Next AktZa
PrItm.id = 2647
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

TmFnt.Name = IniGetVal("Layout", "FnNaUn")
TmFnt.SIZE = IniGetVal("Layout", "FnGrUn")
Set PrFnt = PrKat.AddChildItem(PropertyItemFont, "Schriftart für autom. Unterschrift des Mandanten", TmFnt)
PrFnt.Color = IniGetVal("Layout", "FnFaUn")
PrFnt.id = 2826
PrFnt.Description = IniGetOpt(KeyNa, PrFnt.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Buchhaltung")
PrKat.id = 1500
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Kontenrahmen*", CStr(FKon(GlSet(1, 23), 1)))
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlKoR) 'Standardkontenrahmen
    PrItm.Constraints.Add GlKoR(AktZa, 0)
Next AktZa
PrItm.id = 1505
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

TmKon = CLng(GlSet(1, 78)) 'Standardsteuerkonto
If UBound(GlSaU) >= 1 Then
    For AktZa = 1 To UBound(GlSaU) 'Sachkonten mit Steuerkontenzuordnung
        If CLng(GlSaU(AktZa, 2)) = TmKon Then
            StKo1 = True
            Exit For
        End If
    Next AktZa
    If StKo1 = True And AktZa >= 1 And AktZa <= UBound(GlSaU) Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardsteuerkonto*", GlSaU(AktZa, 3))
    Else
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardsteuerkonto*", GlSaU(1, 3))
    End If
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardsteuerkonto*", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlSaU)
    PrItm.Constraints.Add GlSaU(AktZa, 3)
Next AktZa
PrItm.id = 2950
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

TmKon = CLng(GlSet(1, 25)) 'Standarderlöskonto Bank
If UBound(GlErK) >= 1 Then
    For AktZa = 1 To UBound(GlErK) 'Erlöskonten
        If GlErK(AktZa, 0) = TmKon Then
            EiKo1 = True
            Exit For
        End If
    Next AktZa
    If EiKo1 = True And AktZa >= 1 And AktZa <= UBound(GlErK) Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standarderlöskonto (Bank)*", GlErK(AktZa, 1))
    Else
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standarderlöskonto (Bank)*", GlErK(1, 1))
    End If
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standarderlöskonto (Bank)*", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlErK)
    PrItm.Constraints.Add GlErK(AktZa, 1)
Next AktZa
PrItm.id = 1507
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

TmKon = CLng(GlSet(1, 24)) 'Standarderlöskonto Kasse
If UBound(GlErK) >= 1 Then
    For AktZa = 1 To UBound(GlErK) 'Erlöskonten
        If GlErK(AktZa, 0) = TmKon Then
            EiKo2 = True
            Exit For
        End If
    Next AktZa
    If EiKo2 = True And AktZa >= 1 And AktZa <= UBound(GlErK) Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standarderlöskonto (Kasse)*", GlErK(AktZa, 1))
    Else
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standarderlöskonto (Kasse)*", GlErK(1, 1))
    End If
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standarderlöskonto (Kasse)*", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlErK)
    PrItm.Constraints.Add GlErK(AktZa, 1)
Next AktZa
PrItm.id = 1501
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

TmGe1 = CInt(GlSet(2, 26)) 'Standardgeldkonto (Bank)
If UBound(GlGeK) >= 1 Then
    For AktZa = 1 To UBound(GlGeK) 'Geldkonten
        If GlGeK(AktZa, 0) = TmGe1 Then
            TmKo1 = True
            Exit For
        End If
    Next AktZa
    If TmKo1 = True And AktZa >= 1 And AktZa <= UBound(GlGeK) Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardgeldkonto (Bank)*", GlGeK(AktZa, 4))
    Else
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardgeldkonto (Bank)*", GlGeK(1, 4))
    End If
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardgeldkonto (Bank)*", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlGeK) 'Geldkonten
    PrItm.Constraints.Add GlGeK(AktZa, 4), GlGeK(AktZa, 0)
Next AktZa
PrItm.id = 2845
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

TmGe2 = CInt(GlSet(2, 27)) 'Standardgeldkonto (Kasse)
If UBound(GlGeK) >= 1 Then
    For AktZa = 1 To UBound(GlGeK) 'Geldkonten
        If GlGeK(AktZa, 0) = TmGe2 Then
            TmKo2 = True
            Exit For
        End If
    Next AktZa
    If TmKo2 = True And AktZa >= 1 And AktZa <= UBound(GlGeK) Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardgeldkonto (Kasse)*", GlGeK(AktZa, 4))
    ElseIf UBound(GlGeK) >= 2 Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardgeldkonto (Kasse)*", GlGeK(2, 4))
    Else
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardgeldkonto (Kasse)*", GlGeK(1, 4))
    End If
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardgeldkonto (Kasse)*", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlGeK) 'Geldkonten
    PrItm.Constraints.Add GlGeK(AktZa, 4), GlGeK(AktZa, 0)
Next AktZa
PrItm.id = 1311
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

If GlVar = "PS3" Then
    If GlRDP = False Then
        Set PrItm = PrKat.AddChildItem(PropertyItemString, "TSE Verfahren*", FKon(GlSet(1, 96), 1))
        PrItm.flags = ItemHasComboButton
        PrItm.Constraints.Add "keine TSE"
        PrItm.Constraints.Add "SwissBit TSE"
        PrItm.Constraints.Add "fiskaly TSE"
        PrItm.id = 2970
        PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
        
        If GlSet(1, 96) = "M1" Then
            Set PrItm = PrKat.AddChildItem(PropertyItemString, "TSE Laufwerk*", CStr(GlSet(1, 94)))
            PrItm.flags = ItemHasExpandButton
            PrItm.id = 2968
            PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
        End If
    End If
End If

Set PrItm = PrKat.AddChildItem(PropertyItemString, "DATEV Beraternummer*", CLng(GlSet(2, 48)))
PrItm.id = 1508
PrItm.EditStyle = EditStyleNumber
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "DATEV Mandantennummer*", CLng(GlSet(2, 49)))
PrItm.id = 1509
PrItm.EditStyle = EditStyleNumber
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "DATEV Schnittstelle SOLL und HABEN Tausch*", CBool(GlSet(4, 6)))
PrBol.CheckBoxStyle = True
PrBol.id = 1606
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "DATEV Schnittstelle Sachkonten vierstellig*", CBool(GlSet(4, 47)))
PrBol.CheckBoxStyle = True
PrBol.id = 1608
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "DATEV Sachkontenformatierung*", CBool(GlSet(4, 32)))
PrBol.CheckBoxStyle = True
PrBol.id = 2925
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "DATEV Debitorennamen exportieren", FKon(IniGetVal("System", "BuExPa"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2946
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "DATEV Debitorennummer exportieren", FKon(IniGetVal("System", "DebNum"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2993
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "DATEV Debitorennummer ersetzt Sachkontennummer", FKon(IniGetVal("System", "DebRep"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2994
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Separierter Geldkonten Buchungsnummernkreis*", CBool(GlSet(4, 7)))
PrBol.CheckBoxStyle = True
PrBol.id = 2218
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "GoBD Festschreibung bei Buchungsexport", FKon(IniGetVal("System", "BuExGo"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2988
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Stapelbuchen aktivieren", FKon(IniGetVal("System", "StapBu"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1502
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Umsatzsteuer Splittbuchungen*", CBool(GlSet(4, 46)))
PrBol.CheckBoxStyle = True
PrBol.id = 2879
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Einfache Buchführung verwenden*", CBool(GlSet(4, 76)))
PrBol.CheckBoxStyle = True
PrBol.id = 2948
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)
PrBol.ReadOnly = CBool(GlSet(4, 76))

'------------------------------

Set PrKat = PrGr1.AddCategory("Rechnungen")
PrKat.id = 1300
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Rechnungsbelegtyp*", FKon(GlSet(1, 30), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "R - Standardrechnung"
PrItm.Constraints.Add "L - Laborrechnung"
PrItm.Constraints.Add "A - Abrechnungsstelle"
PrItm.Constraints.Add "M - Rechnungsauftrag"
PrItm.Constraints.Add "G - Gewerberechnung"
PrItm.Constraints.Add "I - Importrechnung"
PrItm.id = 1109
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ReadOnly = GlMVo 'mandantenbezogene Vorgaben verwenden

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Belegversandweg*", FKon(GlSet(1, 101), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Posversand"
PrItm.Constraints.Add "Emailversand"
PrItm.Constraints.Add "Downloadlink"
PrItm.id = 2974
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Rechnungsobergrenze", Format$(IniGetVal("System", "ReObGr"), GlWa1))
PrItm.id = 1301
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Rechnungsuntergrenze", Format$(IniGetVal("System", "ReUnGr"), GlWa1))
PrItm.id = 1302
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Steigerungsfaktor Begründungsaufforderung", Format$(IniGetVal("System", "AbrFak"), GlWa1))
PrItm.id = 2640
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Anzahl der Rechnungsdrucke (Kopien)*", CInt(GlSet(2, 2)))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "1"
PrItm.Constraints.Add "2"
PrItm.Constraints.Add "3"
PrItm.Constraints.Add "4"
PrItm.id = 1303
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Format der Rechnungsnummer*", FKon(GlSet(1, 3), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Jahr-Monat-000000"
PrItm.Constraints.Add "JahrMonat-000000"
PrItm.Constraints.Add "Jahr-000000"
PrItm.Constraints.Add "JahrMonat-0000"
PrItm.id = 1304
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Belegversand an gekennzeichnete E-Mail-Adresse", FKon(IniGetVal("System", "RecEma"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2986
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Warnung bei Überschreiten der Rechnungssumme", FKon(IniGetVal("System", "SumWar"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1306
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Rechnungsnummern immer sofort erzeugen*", CBool(GlSet(4, 4)))
PrBol.CheckBoxStyle = True
PrBol.id = 1307
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Neustart der Rechnungsnummer am Jahresanfang*", CBool(GlSet(4, 5)))
PrBol.CheckBoxStyle = True
PrBol.id = 2823
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Dauerdiagnose mit Doppelklick einfügen", FKon(IniGetVal("System", "DauDia"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1309
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Hinweis bei Überschreiten des Steigerungsfaktors", FKon(IniGetVal("System", "FakHin"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2641
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Eigenes Diagnosetextfeld anzeigen", FKon(IniGetVal("Layout", "EigDia"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2829
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Behandlungsdatum zur Diagnose anzeigen", FKon(IniGetVal("System", "DiaDat"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2833
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Verzögerte Rechnungsübersicht", FKon(IniGetVal("System", "ReVerz"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2848
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Rechnungsvermerk im Krankenblatt*", CBool(GlSet(4, 103)))
PrBol.CheckBoxStyle = True
PrBol.id = 2976
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Farbe unabgeschlossener Rechnungen", CLng(IniGetVal("Layout", "ReAbFa")))
PrItm.id = 2638
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Laboroptionen")
PrKat.id = 1600
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

If StLab >= 1 And StLab <= UBound(GlLab) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Laborkatalog*", GlLab(StLab, 1))
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard-Laborkatalog*", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlLab)
    PrItm.Constraints.Add GlLab(AktZa, 1)
Next AktZa
PrItm.id = 1108
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "LDT Exporttyp", FKon(IniGetVal("System", "ExpTyp"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Laborgemeinschaft"
PrItm.Constraints.Add "Labor-Facharzt-Bericht"
PrItm.id = 1601
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "LDT Import-Zeichensatz*", CStr(FKon(GlSet(1, 66), 1)))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "DOS-ASCII"
PrItm.Constraints.Add "Windows-OEM"
PrItm.id = 1602
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "LDT Export-Zeichensatz", FKon(IniGetVal("System", "ExpFor"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "DOS-ASCII"
PrItm.Constraints.Add "Windows-OEM"
PrItm.id = 1603
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Steigerungsfaktor Laborparameter*", Format$(CSng(GlSet(3, 67)), GlWa1))
PrItm.id = 2650
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "THEDEX Nachrichtendateien generieren", FKon(IniGetVal("System", "TheDex"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1604
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

'------------------------------


Set PrKat = PrGr1.AddCategory("Layouteinstellungen")
PrKat.id = 2100
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

TmFnt.Name = IniGetVal("Layout", "FntNam")
TmFnt.SIZE = IniGetVal("Layout", "FntGro")
Set PrFnt = PrKat.AddChildItem(PropertyItemFont, "Allgemeine System-Schriftart", TmFnt)
PrFnt.Color = IniGetVal("Layout", "FntFar")
PrFnt.id = 1801
PrFnt.Description = IniGetOpt(KeyNa, PrFnt.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Allgemeine Zeilenmarker Farbe für Tabellen", CLng(IniGetVal("Layout", "FarZei")))
PrItm.id = 2101
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Allgemeine Zeilenmarker Farbe für Krankenblatt", CLng(IniGetVal("Layout", "ZeiFar")))
PrItm.id = 2102
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Tabellen-Zeilenmarker anzeigen", FKon(IniGetVal("Layout", "GrdMkr"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1805
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Tabellen-Spaltenköpfe anzeigen", FKon(IniGetVal("Layout", "SpaKop"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1804
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Sortierfunktion mit Tabellen-Spaltenkopf", FKon(IniGetVal("Layout", "SpaSor"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2645
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Tabellen-Gitternetzlinien anzeigen", FKon(IniGetVal("Layout", "GrdGrl"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1806
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Tabelen-Vorschauzeile anzeigen", FKon(IniGetVal("Layout", "GrdPrv"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1807
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Max. Anzahl der Vorschauzeilen", CStr(IniGetVal("Layout", "KraZei")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 1808
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Max. Anzahl Kalenderwahl", CStr(IniGetVal("Layout", "MaxKal")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2884
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Designfensterrahmen einschalten", FKon(IniGetVal("GUI", "Rahmen"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2120
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Menüanimation einschalten", FKon(IniGetVal("Layout", "MenAni"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2897
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Popup-Kalenderfeld", FKon(IniGetVal("Layout", "PopKal"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2649
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "ClearType Textqualität einschalten", FKon(IniGetVal("Layout", "CleTyp"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2123
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Flyoutfenster-Positionen automatisch Speichern", FKon(IniGetVal("Layout", "DocLay"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2124
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Infotabellen auf Startbildschirm zeigen", FKon(IniGetVal("Layout", "StaTab"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2871
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Autom. Erweitern des Mail-Mitarbeiterordners", FKon(IniGetVal("Layout", "MaMiOr"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2873
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Autom. Speichern von Layout Einstellungen verhindern", FKon(IniGetVal("System", "IdiMod"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2830
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)
PrBol.Expanded = Not GlRDP

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Runder Systembutton", FKon(IniGetVal("Layout", "SysBut"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2886
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Steuersatzspalte*", CBool(GlSet(4, 22)))
PrBol.CheckBoxStyle = True
PrBol.id = 2905
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Farbige Register", FKon(IniGetVal("Layout", "FarReg"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2887
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Farbige Modulkennzeichnung", FKon(IniGetVal("Layout", "FarMod"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2888
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Bildschirm-Aktualisierung", FKon(IniGetVal("Layout", "BilAkt"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2898
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Krankenblattdialog im Vordergrund", FKon(IniGetVal("GUI", "KrFoVo"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2985
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Splashscreen deaktivieren", FKon(IniGetVal("Layout", "NoSpla"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2126
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Fenstergröße beim Programmstart", FKon(IniGetVal("Layout", "FenGro"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Variable"
PrItm.Constraints.Add "Maximiert"
PrItm.Constraints.Add "Vorgegeben"
PrItm.id = 2857
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Vorgegebene Fensterbreite", CStr(IniGetVal("Layout", "FeVoBr")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2914
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Vorgegebene Fensterhöhe", CStr(IniGetVal("Layout", "FeVoHo")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2915
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Symbolleisten Schriftgröße", CStr(IniGetVal("Layout", "TolFoH")))
PrItm.EditStyle = EditStyleNumber
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add 10
PrItm.Constraints.Add 11
PrItm.Constraints.Add 12
PrItm.Constraints.Add 13
PrItm.id = 2899
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Krankenblatteinstellungen")
PrKat.id = 2200
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Katalogdiagnosensortierung", FKon(IniGetVal("System", "DiaSor"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2125
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Zahlung Eintrag Gesamtbetrag", FKon(IniGetVal("System", "KraZah"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2952
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Auf Nachfrage neue Rechnung anlegen", FKon(IniGetVal("System", "RechAu"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2883
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Dokumentiert Emails in Krankenblatt*", CBool(GlSet(4, 70)))
PrBol.CheckBoxStyle = True
PrBol.id = 2896
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Alle Einträge beim Einfügen einer Kette markieren", FKon(IniGetVal("System", "KetMar"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2859
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Krankenblattdiagnosen autom. in Abrechnung übernehmen", FKon(IniGetVal("System", "KatDia"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2847
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Autom. setzen des Behandlungsdatums in Eingabezeile", FKon(IniGetVal("System", "AktDat"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2201
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Spaltenköpfe in Krankenblatt anzeigen", FKon(IniGetVal("Layout", "SpalUb"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2202
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Zeilenmarker in Krankenblatt anzeigen", FKon(IniGetVal("Layout", "ZeiMar"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2203
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Gitternetzlinien in Krankenblatt anzeigen", FKon(IniGetVal("Layout", "LinTyp"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2205
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Farbdifferenzierung der Eintragstypen im Krankenblatt", FKon(IniGetVal("Layout", "RowCol"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2206
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Autom. Fokuswechsel zum zuletzt eingefügten Eintrag", FKon(IniGetVal("System", "KraFoc"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2216
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Autom. Krankenblatteintrag für ausgehende Dokumente", FKon(IniGetVal("System", "DocPro"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2820
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Gebührendiagnosezuordnungen auslassen", FKon(IniGetVal("System", "GeDiZu"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2923
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Konstante Krankenblattsortierung*", CBool(GlSet(4, 86)))
PrBol.CheckBoxStyle = True
PrBol.id = 2959
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Cave Infotext zeigen", FKon(IniGetVal("System", "CavTex"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2984
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Krankenblatt Dokumentenimport*", CBool(GlSet(4, 85)))
PrBol.CheckBoxStyle = True
PrBol.id = 2958
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

'Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Nur die eigenen Krankenblatteinträge sichtbar machen", FKon(IniGetVal("System", "MaVoKr"), 1))
'PrBol.CheckBoxStyle = True
'PrBol.id = 2962
'PrBol.Description = IniGetOpt(KeyNa, PrBol.id)
'PrBol.ReadOnly = Not GlRst 'Mandantenbezogene Datenbegrenzung

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Cursorplatzierung bei Bearbeiten eines Eintrags", FKon(IniGetVal("System", "EinCur"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2992
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Prüfung bereits vorhandener Diagnosen*", FKon(GlSet(1, 87), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Diagnosecodeprüfung"
PrItm.Constraints.Add "Diagnosetextprüfung"
PrItm.Constraints.Add "Keine Diagnoseprüfung"
PrItm.id = 2960
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Internen Bild- und PDF Viewer verwenden", CStr(FKon(IniGetVal("System", "BldVie"), 1)))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Interner Viewer"
PrItm.Constraints.Add "Externer Viewer"
PrItm.id = 1605
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Eingabesortierung des Abrechnungskrankenblattes", FKon(IniGetVal("Layout", "ZifAuf"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Absteigend"
PrItm.Constraints.Add "Aufsteigend"
PrItm.id = 2210
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

TmGe3 = CInt(IniGetVal("System", "EinTyp"))
If TmGe3 >= 1 And TmGe3 <= UBound(GlKrA) Then
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Vorgabe für einen neuen Krankenblatteintrag", GlKrA(TmGe3, 2))
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Vorgabe für einen neuen Krankenblatteintrag", vbNullString)
End If
PrItm.flags = ItemHasComboButton
For AktZa = 1 To UBound(GlKrA)
    If GlKrA(AktZa, 0) > 9 Then
        Select Case GlKrA(AktZa, 0)
        Case 24:    'Textdokumente
        Case 101:   'Beleg / Rezept
        Case 102:   'Datei
        Case 104:   'Protokoll
        Case 105:   'Bilddatei
        Case Else:
            PrItm.Constraints.Add GlKrA(AktZa, 2), GlKrA(AktZa, 0)
        End Select
    End If
Next AktZa
PrItm.id = 2212
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

TmFnt.Name = IniGetVal("Layout", "KraFnt")
TmFnt.SIZE = IniGetVal("Layout", "KraGro")
Set PrFnt = PrKat.AddChildItem(PropertyItemFont, "Krankenblattschriftart", TmFnt)
PrFnt.id = 2217
PrFnt.Description = IniGetOpt(KeyNa, PrFnt.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Sonstige Optionen")
PrKat.id = 1700
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Chipkartenlesegerät", FKon(IniGetVal("System", "ChpCrd"), 1))
PrItm.flags = ItemHasComboButton
PrItm.DropDownItemCount = 12
If UBound(GlKvk) >= 1 Then
    For AktZa = 0 To UBound(GlKvk) - 1
        PrItm.Constraints.Add GlKvk(AktZa, 0)
    Next AktZa
End If
PrItm.id = 1704
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Port des Chipkartenlesegerätes", CStr(IniGetVal("System", "SmaPor")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2877
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Verweildauer der Downloadlinks*", CInt(GlSet(2, 100)))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2973
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Popup Benachrichtigungen anzeigen", FKon(IniGetVal("System", "PopFen"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1701
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Popup Verweildauer (Sek.)", CStr(IniGetVal("System", "PopTim")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 1702
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Regelprüfung in Abrechnung aktivieren", FKon(IniGetVal("System", "ResPru"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2846
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "PZN bei Rezept einfügen*", CBool(GlSet(4, 33)))
PrBol.CheckBoxStyle = True
PrBol.id = 1705
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Arzneibezeichnung bei Rezept einfügen", FKon(IniGetVal("System", "RezBet"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2628
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

TmFnt.Name = IniGetVal("Layout", "FnNaRz")
TmFnt.SIZE = IniGetVal("Layout", "FnGrRz")
Set PrFnt = PrKat.AddChildItem(PropertyItemFont, "Rezeptschriftart", TmFnt)
PrFnt.Color = IniGetVal("Layout", "FnFaRz")
PrFnt.id = 1706
PrFnt.Description = IniGetOpt(KeyNa, PrFnt.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Maximale Rezept Zeilenlänge", CStr(IniGetVal("Layout", "MaxZei")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2861
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Programmstart mit Startseite", FKon(IniGetVal("System", "StaSei"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1722
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Startlogo beim Start einblenden", FKon(IniGetVal("System", "StaFor"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2817
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Wartezimmerliste Schließen", FKon(IniGetVal("System", "WaLiCl"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1312
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "IPC3 Nutzung", FKon(IniGetVal("System", "IPCAkt"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1313
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "E-Mail Beschleunigung*", CBool(GlSet(4, 102)))
PrBol.CheckBoxStyle = True
PrBol.id = 2975
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Startseitenübersicht Links", CStr(FKon(IniGetVal("Layout", "StaUb1"), 1)))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Geburtstagsliste"
PrItm.Constraints.Add "Terminliste"
PrItm.Constraints.Add "Offene-Postenliste"
PrItm.Constraints.Add "Aufgabenliste"
PrItm.Constraints.Add "Akontoliste"
PrItm.id = 2812
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Startseitenübersicht Mitte", CStr(FKon(IniGetVal("Layout", "StaUb2"), 1)))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Geburtstagsliste"
PrItm.Constraints.Add "Terminliste"
PrItm.Constraints.Add "Offene-Postenliste"
PrItm.Constraints.Add "Aufgabenliste"
PrItm.Constraints.Add "Akontoliste"
PrItm.id = 2813
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Startseitenübersicht Rechts", CStr(FKon(IniGetVal("Layout", "StaUb3"), 1)))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Geburtstagsliste"
PrItm.Constraints.Add "Terminliste"
PrItm.Constraints.Add "Offene-Postenliste"
PrItm.Constraints.Add "Aufgabenliste"
PrItm.Constraints.Add "Akontoliste"
PrItm.id = 2814
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Terminkalendereinstellungen")
PrKat.id = 2600
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Der Arbeitstag im Kalender beginnt um", IniGetVal("TerSys", "StaZei"))
PrItm.flags = ItemHasComboButton
For AktZa = 0 To 23
    PrItm.Constraints.Add Format$(AktZa, "00") & ":00"
    PrItm.Constraints.Add Format$(AktZa, "00") & ":30"
Next AktZa
PrItm.id = 2601
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Der Arbeitstag in Kalender endet um", IniGetVal("TerSys", "EndZei"))
PrItm.flags = ItemHasComboButton
For AktZa = 0 To 23
    PrItm.Constraints.Add Format$(AktZa, "00") & ":00"
    PrItm.Constraints.Add Format$(AktZa, "00") & ":30"
Next AktZa
PrItm.id = 2602
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Die Skalierung im Kalender beginnt um", IniGetVal("TerSys", "AnsSta"))
PrItm.flags = ItemHasComboButton
For AktZa = 0 To 23
    PrItm.Constraints.Add Format$(AktZa, "00") & ":00"
    PrItm.Constraints.Add Format$(AktZa, "00") & ":30"
Next AktZa
PrItm.id = 2629
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Die Skalierung im Kalender endet um", IniGetVal("TerSys", "AnsEnd"))
PrItm.flags = ItemHasComboButton
For AktZa = 0 To 23
    PrItm.Constraints.Add Format$(AktZa, "00") & ":00"
    PrItm.Constraints.Add Format$(AktZa, "00") & ":30"
Next AktZa
PrItm.id = 2630
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Der Druckbereich beginnt um", IniGetVal("TerSys", "StaPri"))
PrItm.flags = ItemHasComboButton
For AktZa = 0 To 23
    PrItm.Constraints.Add Format$(AktZa, "00") & ":00"
Next AktZa
PrItm.id = 2607
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Der Druckbereich endet um", IniGetVal("TerSys", "EndPri"))
PrItm.flags = ItemHasComboButton
For AktZa = 0 To 23
    PrItm.Constraints.Add Format$(AktZa, "00") & ":00"
Next AktZa
PrItm.id = 2608
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Kalender Zeitscala (Min.)", IniGetVal("TerSys", "TimSca"))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "01"
PrItm.Constraints.Add "05"
PrItm.Constraints.Add "10"
PrItm.Constraints.Add "15"
PrItm.Constraints.Add "20"
PrItm.Constraints.Add "30"
PrItm.Constraints.Add "45"
PrItm.Constraints.Add "60"
PrItm.id = 2603
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Kalenderfocus beim Tages- oder Wochenwechsel", FKon(IniGetVal("TerSys", "TeKaFo"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2983
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminzeit aus dem Terminbetreff verwenden*", CBool(GlSet(4, 51)))
PrBol.CheckBoxStyle = True
PrBol.id = 2627
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminnachricht mit ICS Dateiversand*", CBool(GlSet(4, 68)))
PrBol.CheckBoxStyle = True
PrBol.id = 2917
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminnachricht auch an BCC*", CBool(GlSet(4, 80)))
PrBol.CheckBoxStyle = True
PrBol.id = 2945
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Starre Termintaktung verwenden*", CBool(GlSet(4, 28)))
PrBol.CheckBoxStyle = True
PrBol.id = 1607
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Anzeige stornierter Termine in den Termindetails*", CBool(GlSet(4, 92)))
PrBol.CheckBoxStyle = True
PrBol.id = 2966
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminbelegungswarnung", FKon(IniGetVal("TerSys", "TerWar"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2942
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Zeitscala mit Minutenmarkierung", FKon(IniGetVal("TerSys", "ScaMin"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2609
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Die Terminzeit in Monatsansicht als Uhr anzeigen", FKon(IniGetVal("TerSys", "TimClo"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2604
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Erweiterte Termininformationen anzeigen", FKon(IniGetVal("TerSys", "TerErw"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2617
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Anzeige der Endzeit in Monatsansicht", FKon(IniGetVal("TerSys", "EndTim"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2605
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Keine Farbunterscheidung in Raumbelegung", FKon(IniGetVal("TerSys", "KeFaRa"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2872
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Die Raumzuordnung numerisch sortiert anzeigen*", CBool(GlSet(4, 35)))
PrBol.CheckBoxStyle = True
PrBol.id = 2840
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Outlook Font Glyphs verwenden", FKon(IniGetVal("TerSys", "FntGly"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2606
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Serienterminplanung ohne Ganztagstermine", FKon(IniGetVal("TerSys", "GanzBe"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2810
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminplanung mit Sprechzeitenprüfung", FKon(IniGetVal("TerSys", "SpreBe"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2821
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)


Set PrItm = PrKat.AddChildItem(PropertyItemString, "Mindestbreite der Zeitspalten im Kalender in Pixel", CStr(IniGetVal("TerSys", "TiSpBr")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2612
PrBol.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Kalenderdruck mit horizontaler Druckausrichtung", FKon(IniGetVal("TerSys", "TiPrHo"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2613
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Scroll Leiste für Monatsansicht anzeigen", FKon(IniGetVal("TerSys", "KalScl"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2614
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Privaten Outloktermine nicht synchronisieren", FKon(IniGetVal("TerSys", "OutPri"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2811
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

'Set PrItm = PrKat.AddChildItem(PropertyItemString, "Intervall Kalenderaktualisierung in Sek.", CStr(IniGetVal( "TerSys", "TimInt")))
'PrItm.EditStyle = EditStyleNumber
'PrItm.flags = ItemHasComboButton
'PrItm.Constraints.Add "120"
'PrItm.Constraints.Add "300"
'PrItm.Constraints.Add "600"
'PrItm.Constraints.Add "900"
'PrItm.id = 2616
'PrItm.Description = IniGetopt( KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Betreff Trennzeichen für Terminsplitt", CStr(IniGetVal("TerSys", "TreZei")))
PrItm.id = 2631
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mitarbeitername in Terminort speichern*", CBool(GlSet(4, 69)))
PrBol.CheckBoxStyle = True
PrBol.id = 2654
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Telefonnummer in Terminbetreff anzeigen", FKon(IniGetVal("TerSys", "ZeiBet"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2632
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminerinnerung im Kalender anzeigen", FKon(IniGetVal("TerSys", "TerErn"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2636
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Einträge in Terminliste erst nach Sucheingabe", FKon(IniGetVal("TerSys", "TerUbe"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2639
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Abfrage E-Mail-Terminerinnerung bei Terminverschiebung", FKon(IniGetVal("TerSys", "TeErSp"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2991
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Virtuelle Endlostermine ermöglichen", FKon(IniGetVal("TerSys", "EndSer"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2646
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Anzeigen der Sprechzeiten im Kalender", FKon(IniGetVal("TerSys", "ManZei"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2834
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mandanten mit unterschiedlicher Farbkennzeichnung", FKon(IniGetVal("TerSys", "ManFar"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2836
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Behandlungsräume mit unterschiedlicher Farbkennzeichnung", FKon(IniGetVal("TerSys", "RauFar"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2837
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Geschlechter im Termin darstellen", FKon(IniGetVal("TerSys", "TerGes"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2838
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminbetreff in Fettdruck darstellen", FKon(IniGetVal("TerSys", "TeFnBl"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2839
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Patienten-CAVE ins Kommentarfeld", FKon(IniGetVal("TerSys", "TerCav"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2951
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminmarkierung bei Terminzetteldruck", FKon(IniGetVal("TerSys", "AuMark"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2919
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Standard-Gebührenkette mit Terminlänge multiplizieren", FKon(IniGetVal("TerSys", "LeiAnz"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2648
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Alternative Berechnung der Serienfälligkeit", FKon(IniGetVal("TerSys", "SerBer"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2816
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Berücksichtigen von Aktontozahlungen bei Terminleistungen", FKon(IniGetVal("TerSys", "TeLeAk"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2880
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mitarbeiter / Mandanten in Raumbelegung aktivieren", FKon(IniGetVal("TerSys", "RauMan"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2921
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Verkehrsnamen in Mitarbeiterplan", FKon(IniGetVal("TerSys", "KalMit"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2989
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminbetreff-Kompatibilitätsmodus", FKon(IniGetVal("TerSys", "TerBet"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2920
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminleistungen nur für passene Referenzrechnung", FKon(IniGetVal("TerSys", "ReTeSe"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2924
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Name inaktiver Mitarbeiter in den Termindetails", FKon(IniGetVal("TerSys", "TeInMi"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2990
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

If CBool(GlSet(4, 29)) = False Then
    Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Mitarbeiterplan anstelle Mandantenplan*", CBool(GlSet(4, 29)))
    PrBol.CheckBoxStyle = True
    PrBol.id = 2869
    PrBol.Description = IniGetOpt(KeyNa, PrBol.id)
End If

'------------------------------

Set PrKat = PrGr1.AddCategory("Textverarbeitung")
PrKat.id = 2800
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

TmFnt.Name = IniGetVal("TexVer", "FntNam")
TmFnt.SIZE = IniGetVal("TexVer", "FntGro")
Set PrFnt = PrKat.AddChildItem(PropertyItemFont, "Standardschriftart", TmFnt)
PrFnt.Color = IniGetVal("TexVer", "FntFar")
PrFnt.id = 2801
PrFnt.Description = IniGetOpt(KeyNa, PrFnt.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Linker Rand (mm)", CStr(IniGetVal("TexVer", "RnLink")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2802
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Rechter Rand (mm)", CStr(IniGetVal("TexVer", "RnRech")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2803
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Oberer Rand (mm)", CStr(IniGetVal("TexVer", "RnOben")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2804
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Unterer Rand (mm)", CStr(IniGetVal("TexVer", "RnUnte")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2805
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard Seitenhöhe (mm)", CStr(IniGetVal("TexVer", "SeiHoh")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2806
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standard Seitenbreite (mm)", CStr(IniGetVal("TexVer", "SeiBre")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2807
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Automatische Microsoft-Word Konvertierung", FKon(IniGetVal("TexVer", "WorKon"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Nachfragen Konvertierung"
PrItm.Constraints.Add "Automatische Konvertierung"
PrItm.Constraints.Add "Keine Konvertierung"
PrItm.id = 2651
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Krankenblattpositionen mit Datum", FKon(IniGetVal("TexVer", "KrTyDa"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2864
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standarddokumentenvorlage", IniGetVal("TexVer", "StaDoc"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 2890
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardrezeptvorlage", IniGetVal("TexVer", "StaRez"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 2891
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Standardnewslettervorlage", IniGetVal("TexVer", "StaNew"))
PrItm.flags = ItemHasExpandButton
PrItm.id = 2892
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Datenbankeinstellungen")
PrKat.id = 2300
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)
PrKat.Expandable = Not GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Verwendeter Datenbanktyp", FKon(IniGetVal("System", "DBaTyp"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Microsoft SQL Native Client"
PrItm.Constraints.Add "Microsoft SQL Server OLEDB"
PrItm.Constraints.Add "DBX Datenbank"
PrItm.Constraints.Add "DBV Datenbank"
PrItm.id = 2301
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Datenbanktimeout (Sek.)", CStr(IniGetVal("System", "DatAkt")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2302
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Name", IniGetVal("System", "DatSer"))
PrItm.id = 2303
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Datenbankname", IniGetVal("System", "DatTab"))
PrItm.id = 2304
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Authentifizierung", FKon(IniGetVal("System", "DatVer"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "NT-Authentifizierung"
PrItm.Constraints.Add "Benutzerrechte"
PrItm.id = 2305
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Benutzername", IniGetVal("System", "DatUsr"))
PrItm.id = 2309
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

If GlVrs = True Then 'Kennwort verschlüsselt
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Benutzerpasswort", SCrypt(IniGetVal("System", "DaUsPa"), False))
    PrItm.PasswordMask = True
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Benutzerpasswort", IniGetVal("System", "DaUsPa"))
End If
PrItm.id = 2310
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "SQL-Server Benutzerpasswort verschlüsseln", FKon(IniGetVal("System", "DaVers"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2308
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Asynchrone Verbindung verwenden", FKon(IniGetVal("System", "DatLad"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2306
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Jet Sofortspeicherung verwenden", FKon(IniGetVal("System", "JetSav"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2307
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

'------------------------------

Set PrKat = PrGr1.AddCategory("Arzneidatenbank")
PrKat.id = 2400
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)
PrKat.Expandable = Not GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Verwendeter Datenbanktyp", FKon(IniGetVal("System", "MatTyp"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Microsoft SQL Native Client"
PrItm.Constraints.Add "Microsoft SQL Server OLEDB"
PrItm.Constraints.Add "DBX Datenbank"
PrItm.Constraints.Add "DBV Datenbank"
PrItm.id = 2311
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Datenbanktimeout (Sek.)", CStr(IniGetVal("System", "MatAkt")))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2312
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Name", IniGetVal("System", "MatSer"))
PrItm.id = 2313
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Datenbankname", IniGetVal("System", "MatTab"))
PrItm.id = 2314
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Authentifizierung", FKon(IniGetVal("System", "MatVer"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "NT-Authentifizierung"
PrItm.Constraints.Add "Benutzerrechte"
PrItm.id = 2315
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Benutzername", IniGetVal("System", "MatUsr"))
PrItm.id = 2318
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

If GlVrs = True Then 'Kennwort verschlüsselt
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Benutzerpasswort", SCrypt(IniGetVal("System", "MaUsPa"), False))
    PrItm.PasswordMask = True
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "SQL-Server Benutzerpasswort", IniGetVal("System", "MaUsPa"))
End If
PrItm.id = 2319
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "SQL-Server Benutzerpasswort verschlüsseln", FKon(IniGetVal("System", "MaVers"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2317
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Asynchrone Verbindung verwenden", FKon(IniGetVal("System", "MatLad"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2316
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Exit Sub

MeErr:
SLogi "FRgLa: FEHLER - " & Err.Description & " (Nr. " & Err.Number & ")"
DoEvents
If GlDbg = True Then MsgBox Err.Description, 48, "FRgLa " & Err.Number
Resume Next

End Sub
Private Sub FRgLo()
On Error GoTo MeErr

Dim KeyNa As String
Dim AktZa As Integer
Dim TmFnt As New StdFont

Set PrGr1 = Me.prpGrid1
KeyNa = "Optionentexte"

Set PrKat = PrGr1.AddCategory("Druckeinstellungen")
PrKat.id = 2900
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Vor dem Drucken immer Druckvorschau", FKon(IniGetVal("System", "DruVor"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 1305
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Art der Druckvorschau", FKon(IniGetVal("System", "VorTyp"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Einfache Druckvorschau"
PrItm.Constraints.Add "Erweiterte Druckvorschau"
PrItm.Constraints.Add "Externe Druckvorschau"
PrItm.id = 2895
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.Expandable = Not GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Seitengröße aktivieren", FKon(IniGetVal("System", "SeiGro"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Druckervorgabe"
PrItm.Constraints.Add "DIN A4"
PrItm.Constraints.Add "DIN A5"
PrItm.Constraints.Add "DIN A6"
PrItm.id = 2903
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Druckauftrags-Separierung aktivieren", FKon(IniGetVal("System", "DruSep"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2904
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Dublexdruck aktivieren", FKon(IniGetVal("System", "DubDru"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2901
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Netzwerkdrucker auflisten", FKon(IniGetVal("System", "NetDru"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2868
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrKat = PrGr1.AddCategory("Onlinedienste")
PrKat.id = 2500
PrKat.Description = IniGetOpt(KeyNa, PrKat.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs Benutzername*", CStr(GlSet(1, 14)))
PrItm.id = 2652
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs Sytem Passwort*", SCrypt(CStr(GlSet(1, 15)), False))
PrItm.PasswordMask = True
PrItm.id = 2653
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs Sytem Aktivieren*", CBool(GlSet(4, 12)))
PrBol.CheckBoxStyle = True
PrBol.id = 2858
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs Sytem Mitarbeiterwahl*", CBool(GlSet(4, 13)))
PrBol.CheckBoxStyle = True
PrBol.id = 2916
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs Sytem Adressenerfassung*", CBool(GlSet(4, 79)))
PrBol.CheckBoxStyle = True
PrBol.id = 2953
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs Sytem PIN Dialog*", CBool(GlSet(4, 40)))
PrBol.CheckBoxStyle = True
PrBol.id = 2922
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs Sytem Storno-Dialog*", CBool(GlSet(4, 20)))
PrBol.CheckBoxStyle = True
PrBol.id = 2941
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs Sytem Warteliste-Dialog*", CBool(GlSet(4, 89)))
PrBol.CheckBoxStyle = True
PrBol.id = 2963
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs System belegte Buchungszeiten*", CBool(GlSet(4, 52)))
PrBol.CheckBoxStyle = True
PrBol.id = 2928
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs System ICS bei Emailbestätigung*", CBool(GlSet(4, 53)))
PrBol.CheckBoxStyle = True
PrBol.id = 2929
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs Sytem Stornierte Entfernen*", CBool(GlSet(4, 108)))
PrBol.CheckBoxStyle = True
PrBol.id = 2981
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Online-Terminbuchungs System autom. Aktualisierung*", CBool(GlSet(4, 91)))
PrBol.CheckBoxStyle = True
PrBol.id = 2965
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs System Schriftart*", CStr(GlSet(1, 54)))
PrItm.flags = ItemHasComboButton
For AktZa = 0 To UBound(GlOTF) - 1
    PrItm.Constraints.Add GlOTF(AktZa)
Next AktZa
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.id = 2930

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Online-Terminbuchungs System allg. Textfarbe*", CLng(GlSet(2, 55)))
PrItm.id = 2931
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Online-Terminbuchungs System allg. Hintergrundfarbe*", CLng(GlSet(2, 56)))
PrItm.id = 2932
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs System allg. Textgröße*", CLng(GlSet(2, 57)))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "10"
PrItm.Constraints.Add "11"
PrItm.Constraints.Add "12"
PrItm.Constraints.Add "13"
PrItm.Constraints.Add "14"
PrItm.Constraints.Add "15"
PrItm.id = 2933
PrItm.EditStyle = EditStyleNumber
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Online-Terminbuchungs System Button Hintergrundfarbe*", CLng(GlSet(2, 58)))
PrItm.id = 2934
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Online-Terminbuchungs System Button Textfarbe*", CLng(GlSet(2, 59)))
PrItm.id = 2935
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Online-Terminbuchungs System Button Hooverfarbe*", CLng(GlSet(2, 60)))
PrItm.id = 2936
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemColor, "Online-Terminbuchungs System Button Deaktiviertfarbe*", CLng(GlSet(2, 61)))
PrItm.id = 2937
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs System Link Anschlussseite*", CStr(GlSet(1, 62)))
PrItm.id = 2938
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ValueMetrics.MaxLength = 250

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs System Link Datenschutzerklärung*", CStr(GlSet(1, 63)))
PrItm.id = 2939
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ValueMetrics.MaxLength = 250

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs System Link für Impressum*", CStr(GlSet(1, 38)))
PrItm.id = 2912
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ValueMetrics.MaxLength = 250

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Online-Terminbuchungs Provider*", CStr(GlSet(1, 16)))
PrItm.id = 2853
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

TmFnt.Name = IniGetVal("Layout", "MaiFnt")
TmFnt.SIZE = IniGetVal("Layout", "MaiGro")
Set PrFnt = PrKat.AddChildItem(PropertyItemFont, "E-Mail Schriftart", TmFnt)
PrFnt.Color = IniGetVal("Layout", "MaiFar")
PrFnt.id = 2860
PrFnt.Description = IniGetOpt(KeyNa, PrFnt.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "E-Mail Terminerinnerung aktivieren*", CBool(GlSet(4, 88)))
PrBol.CheckBoxStyle = True
PrBol.id = 2961
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "E-Mail Sortierung in Gruppen", FKon(IniGetVal("System", "EmlSor"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2867
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "E-Mail Antwort Einleitungssatz", FKon(IniGetVal("System", "EmlDia"), 1))
PrBol.CheckBoxStyle = True
PrBol.id = 2827
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "E-Mail Nachricht standardmäßig an", FKon(IniGetVal("System", "AutArt"), 1))
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "Emailversand an einen Patienten"
PrItm.Constraints.Add "Emailversand an den Mandanten"
PrItm.Constraints.Add "Emailversand an alle Patienten"
PrItm.id = 2832
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Neuaufnahmeformular-Webadresse*", CStr(GlSet(1, 82)))
PrItm.id = 2955
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "SMS Terminerinnerung aktivieren*", CBool(GlSet(4, 90)))
PrBol.CheckBoxStyle = True
PrBol.id = 2964
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SMS Produkt Token*", CStr(GlSet(1, 39)))
PrItm.id = 2913
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ValueMetrics.MaxLength = 40

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SMS Account-ID*", CStr(GlSet(1, 36)))
PrItm.id = 2910
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ValueMetrics.MaxLength = 40

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SMS Absenderkennung*", CStr(GlSet(1, 37)))
PrItm.id = 2911
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.ValueMetrics.MaxLength = 20

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "CalDAV / CardDAV / Exchange Synchronisation*", CBool(GlSet(4, 17)))
PrBol.CheckBoxStyle = True
PrBol.id = 2893
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Terminland WebCAL (ICS) aktivieren*", CBool(GlSet(4, 81)))
PrBol.CheckBoxStyle = True
PrBol.id = 2954
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Terminland Benutzername*", CStr(GlSet(1, 83)))
PrItm.id = 2956
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Terminland Passwort*", SCrypt(CStr(GlSet(1, 84)), False))
PrItm.PasswordMask = True
PrItm.id = 2957
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Proxyserver Verwenden*", CBool(GlSet(4, 41)))
PrBol.CheckBoxStyle = True
PrBol.id = 2881
PrBol.Description = IniGetOpt(KeyNa, PrBol.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Proxyserver Name*", CStr(GlSet(1, 42)))
PrItm.id = 2882
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)
PrItm.Expandable = Not GlRDP

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SMTP Praxisname*", CStr(GlSet(1, 104)))
PrItm.id = 2977
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SMTP IP-Adresse*", CStr(GlSet(1, 105)))
PrItm.id = 2978
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SMTP SocksProxyServer*", CStr(GlSet(1, 106)))
PrItm.id = 2979
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Set PrItm = PrKat.AddChildItem(PropertyItemString, "SMTP SocksProxyPort*", CInt(GlSet(2, 107)))
PrItm.EditStyle = EditStyleNumber
PrItm.id = 2980
PrItm.Description = IniGetOpt(KeyNa, PrItm.id)

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRgLo " & Err.Number
Resume Next

End Sub
Private Sub FRgSv(ByVal TolId As Long)
On Error GoTo MeErr

Dim AktZa As Long
Dim TmStr As String
Dim TmFnt As New StdFont

Set PrGr1 = Me.prpGrid1

Select Case TolId
Case 1102: 'Standard-Gebührenkatalog
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlGKa)
            If GlGKa(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaKat", AktZa
                S_SeSe 1, , AktZa
                Exit For
            End If
        Next AktZa
Case 1103: 'Standardzahlungsziel
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlZah)
            If GlZah(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaZil", AktZa
                S_SeSe 22, , AktZa
                Exit For
            End If
        Next AktZa
Case 1104: 'Standardsteuersatz
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlStu)
            If GlStu(AktZa, 2) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaStu", AktZa
                S_SeSe 65, , AktZa
                Exit For
            End If
        Next AktZa
Case 1105: 'Standwährung
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlWar)
            If GlWar(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaWar", AktZa
                S_SeSe 32, , AktZa
                Exit For
            End If
        Next AktZa
Case 1106: 'Standard-Gebührenkette 1
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlKet)
            If GlKet(AktZa, 2) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaKet", AktZa
                S_SeSe 2, , AktZa
                Exit For
            End If
        Next AktZa
Case 1714: 'Standard-Gebührenkette 2
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlKet)
            If GlKet(AktZa, 2) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaKe2", AktZa
                S_SeSe 75, , AktZa
                Exit For
            End If
        Next AktZa
Case 1107: 'Standard-Empfängertyp
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Vorgabe", "StaEmp", FKon(PrItm.Value, 2)
Case 1108: 'Standard-Laborkatalog
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlLab)
            If GlLab(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaLab", AktZa
                S_SeSe 19, , AktZa
                Exit For
            End If
        Next AktZa
Case 2960: 'Prüfung bereits vorhandener Diagnosen
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DiaPru", FKon(PrItm.Value, 2)
        S_SeSe 88, FKon(PrItm.Value, 2)
Case 2894: 'Standard-Dezimaltrennzeichen
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Vorgabe", "StaDez", FKon(PrItm.Value, 2)
        S_SeSe 20, FKon(PrItm.Value, 2)
Case 1109: 'Standard-Belegtyp
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Vorgabe", "StaRet", FKon(PrItm.Value, 2)
        S_SeSe 31, FKon(PrItm.Value, 2)
Case 2974: 'Standard-Rechnungsversandweg
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "RecVer", FKon(PrItm.Value, 2)
        S_SeSe 102, FKon(PrItm.Value, 2)
Case 1110: 'Standard-Behandlungsraum
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlRmu)
            If GlRmu(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaRau", AktZa
                Exit For
            End If
        Next AktZa
Case 2987: 'Standard-Briefanrede
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlBri)
            If GlBri(AktZa, 0) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaBrf", AktZa
                Exit For
            End If
        Next AktZa
Case 2626: 'Standard-Mandant
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlMaA)
            If GlMaA(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaMan", AktZa
                Exit For
            End If
        Next AktZa
Case 2647: 'Standard-Mitarbeiter
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlMiA)
            If GlMiA(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaMit", AktZa
                Exit For
            End If
        Next AktZa
Case 2856: 'Standard-Mitarbeiter Online-Terminbuchungs Sytem
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlMiT)
            If GlMiT(AktZa, 1) = PrItm.Value Then
                IniSetVal "Vorgabe", "StaMio", AktZa
                Exit For
            End If
        Next AktZa
Case 2826: 'Schriftart Unterschrift
        Set PrFnt = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FnNaUn", PrFnt.FontFaceName
        IniSetVal "Layout", "FnGrUn", PrFnt.FontSize
        IniSetVal "Layout", "FnFaUn", PrFnt.Color
Case 1201: 'Datenbankdatei
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "SysPfa", "DatPfa", PrItm.Value
Case 1203: 'Backupordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "BackPf", PrItm.Value
Case 1204: 'Importordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ImpPfa", PrItm.Value
Case 1210: 'Exportordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ExpPfa", PrItm.Value
Case 1205: 'Bilderordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "BilPfa", PrItm.Value
Case 1211: 'Emailordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "EmalPf", PrItm.Value
Case 1212: 'Termineordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "TermPf", PrItm.Value
Case 1206: 'Dokumentenordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DockPf", PrItm.Value
Case 1208: 'Worddokumente
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DocPfa", PrItm.Value
Case 1213: 'Filterordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "FiltPf", PrItm.Value
Case 1214: 'Temporärordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "TmpPfa", PrItm.Value
Case 1215: 'Formulareordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ForPfa", PrItm.Value
Case 2637: 'WEGAMED Programmordner
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "WegPfa", PrItm.Value
Case 2863: 'GDT Programmdatei
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "GDTPrg", PrItm.Value
Case 2890: 'Standarddokumentenvorlage
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "StaDoc", PrItm.Value
Case 2891: 'Standardrezeptvorlage
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "StaRez", PrItm.Value
Case 2892: 'Standardnewslettervorlage
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "StaNew", PrItm.Value
Case 2642: 'Startseitendokument
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "StaWeb", PrItm.Value
Case 1301: 'Rechnungsobergrenze
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ReObGr", PrItm.Value
Case 1302: 'Rechnungsuntergrenze
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ReUnGr", PrItm.Value
Case 2640: 'Steiegrungsfaktor Begründung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "AbrFak", PrItm.Value
Case 2650: 'Steigerungsfaktor Laborparameter
        Set PrItm = PrGr1.FindItem(TolId)
        If PrItm.Value <= 0 Then PrItm.Value = 1
        IniSetVal "System", "LabFak", PrItm.Value
        S_SeSe 68, , , CSng(PrItm.Value)
Case 1303: 'Anzahl der Rechnungsdrucke (Kopien)
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DruKop", CInt(PrItm.Value)
        S_SeSe 3, , CInt(PrItm.Value)
Case 2973: 'Verweildauer der Downloadlinks
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "RecVer", CInt(PrItm.Value)
        S_SeSe 101, , CInt(PrItm.Value)
Case 1304: 'Format der Rechnungsnummer
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "NumFor", FKon(PrItm.Value, 2)
        S_SeSe 4, FKon(PrItm.Value, 2)
Case 1305: 'Druckvorschau
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DruVor", FKon(PrBol.Value, 2)
Case 2846: 'Regelprüfung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "ResPru", FKon(PrBol.Value, 2)
Case 2986: 'Belegversand an gekennzeichnete E-Mail-Adresse
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "RecEma", FKon(PrBol.Value, 2)
Case 1306: 'Warnung bei Rechnungssumme
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "SumWar", FKon(PrBol.Value, 2)
Case 2926: 'Mandant der neuen Rechnung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "RecMan", FKon(PrItm.Value, 2)
        S_SeSe 51, CStr(FKon(PrItm.Value, 2))
Case 1307: 'Rechnungsnummern sofort erzeugen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "ReNuEr", FKon(PrBol.Value, 2)
        S_SeSe 5, , , , CBool(PrBol.Value)
        DoEvents
        GlNeK = GlKoX 'Protokolleintrag
        With GlNeK
            .PatNr = GlMan(GlSMa, 2)
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = 104
            .ZiStr = Format$(Now, "hh:mm") & " Uhr"
            .NeuEi = True
            .KeiAk = True
            .Mitar = GlMiA(GlSmI, 2)
            .TeStr = "Einstellungen im Optionsdialog geändert: Rechnungsnummern sofort erzeugen"
        End With
        S_Prot
Case 2823: 'Neustart der Rechnungsnummer am Jahresanfang
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "ReNuSt", FKon(PrBol.Value, 2)
        S_SeSe 6, , , , CBool(PrBol.Value)
        DoEvents
        GlNeK = GlKoX 'Protokolleintrag
        With GlNeK
            .PatNr = GlMan(GlSMa, 2)
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = 104
            .ZiStr = Format$(Now, "hh:mm") & " Uhr"
            .NeuEi = True
            .KeiAk = True
            .Mitar = GlMiA(GlSmI, 2)
            .TeStr = "Einstellungen im Optionsdialog geändert: Neustart der Rechnungsnummer am Jahresanfang"
        End With
        S_Prot
Case 1309: 'Dauerdiagnose mit Doppelklick einfügen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DauDia", FKon(PrBol.Value, 2)
Case 2641: 'Hinweis Steigerungsfaktor
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "FakHin", FKon(PrBol.Value, 2)
Case 2829: 'Eigenes Diagnosetextfeld
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "EigDia", FKon(PrBol.Value, 2)
Case 2833: 'Behandlungsdatum zur Diagnose
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DiaDat", FKon(PrBol.Value, 2)
Case 2848: 'Verzögerte Rechnungsübersicht
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "ReVerz", FKon(PrBol.Value, 2)
Case 2638: 'Rechnungsfarbe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "ReAbFa", PrItm.Value
Case 2958: 'Krankenblatt Dokumentenimport
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DokEin", FKon(PrBol.Value, 2)
        S_SeSe 86, , , , CBool(PrBol.Value)
Case 2976: 'Rechnungsvermerk im Krankenblatt
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "RecKra", FKon(PrBol.Value, 2)
        S_SeSe 104, , , , CBool(PrBol.Value)
Case 2959: 'Konstante Krankenblattsortierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "KraSor", FKon(PrBol.Value, 2)
        S_SeSe 87, , , , CBool(PrBol.Value)
Case 2961: 'Emailerinnerung aktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "EriEma", FKon(PrBol.Value, 2)
        S_SeSe 89, , , , CBool(PrBol.Value)
Case 1401: 'PAD/MAD Kundennummer
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "PVSNum", PrItm.Value
Case 1402: 'Letzte Stapelnummer
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "PADVol", PrItm.Value
Case 1403: 'PAD-Zeilenumbruch
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "PADZei", FKon(PrBol.Value, 2)
Case 1404: 'PAD-Umlautekonvertierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "PADUml", FKon(PrBol.Value, 2)
Case 2825: 'Rechnungsexport mit benanntem Gebührenkatalog
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "PADKat", FKon(PrBol.Value, 2)
        S_SeSe 44, , , , CBool(PrBol.Value)
Case 2876: 'Keine Preisberechnung bei IGeL
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "PrBeIg", FKon(PrBol.Value, 2)
        S_SeSe 45, , , , CBool(PrBol.Value)
Case 2927: 'Keine Positionskennzeichen bei Medikamenten und Begründungen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "PvPoMe", FKon(PrBol.Value, 2)
        S_SeSe 46, , , , CBool(PrBol.Value)
Case 1501: 'Standarderlöskonto Kasse
        Set PrItm = PrGr1.FindItem(TolId)
        If GlKnF = False Then 'Sachkontenformatierung sechsstellig
            TmStr = Val(Left$(PrItm.Value, 6))
        Else
            TmStr = Val(Left$(PrItm.Value, 4))
        End If
        IniSetVal "System", "EinKto", SBuFo(CLng(TmStr))
        S_SeSe 25, SBuFo(CLng(TmStr))
Case 1507: 'Standarderlöskonto Bankkonto
        TmStr = Val(Left$(PrItm.Value, 6))
        IniSetVal "System", "EinKt2", SBuFo(CLng(TmStr))
        S_SeSe 26, SBuFo(CLng(TmStr))
Case 2950: 'Standardsteuerkonto
        TmStr = Val(Left$(PrItm.Value, 6))
        IniSetVal "System", "StStKt", SBuFo(CLng(TmStr))
        S_SeSe 79, SBuFo(CLng(TmStr))
Case 2845: 'Standardgeldkonto (Bankkonto)
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "StaGld", CLng(Left$(PrItm.Value, 2))
        S_SeSe 27, , CLng(Left$(PrItm.Value, 2))
Case 1311: 'Standardgeldkonto (Kasse)
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "StaGl2", CLng(Left$(PrItm.Value, 2))
        S_SeSe 28, , CLng(Left$(PrItm.Value, 2))
Case 1608: 'DATEV Schnittstelle Sachkonten vierstellig
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DaExKo", FKon(PrBol.Value, 2)
        S_SeSe 48, , , , CBool(PrBol.Value)
Case 2993: 'DATEV Debitorennummer exportieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DebNum", FKon(PrBol.Value, 2)
Case 2994: 'DATEV Debitorennummer ersetzt Sachkontennummer
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DebRep", FKon(PrBol.Value, 2)
Case 2925: 'DATEV Sachkontenformatierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "KntFor", FKon(PrBol.Value, 2)
        S_SeSe 33, , , , CBool(PrBol.Value)
Case 1502: 'Stapelbuchen in Buchhaltung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "StapBu", FKon(PrBol.Value, 2)
Case 2988: 'GoBD Festschreibung bei Buchungexport
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BuExGo", FKon(PrBol.Value, 2)
Case 2946: 'DATEV Debitorennamen exportieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BuExPa", FKon(PrBol.Value, 2)
Case 1505: 'Standardkontenrahmen
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "KntRam", FKon(PrItm.Value, 2)
        S_SeSe 24, FKon(PrItm.Value, 2)
Case 1508: 'DATEV Beraternummer
        Set PrItm = PrGr1.FindItem(TolId)
        If PrItm.Value <> vbNullString Then IniSetVal "System", "DATVBe", PrItm.Value
        S_SeSe 49, , CLng(PrItm.Value)
Case 1509: 'DATEV Mandantennummer
        Set PrItm = PrGr1.FindItem(TolId)
        If PrItm.Value <> vbNullString Then IniSetVal "System", "DATVMa", PrItm.Value
        S_SeSe 50, , CLng(PrItm.Value)
Case 2877: 'Port des Chipkartenlesegerätes
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "SmaPor", PrItm.Value
Case 1601: 'LDT Exporttyp
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ExNam", FKon(PrItm.Value, 2)
Case 1602: 'LDT Import-Zeichensatz
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ImpFor", FKon(PrItm.Value, 2)
        S_SeSe 67, CStr(FKon(PrItm.Value, 2))
Case 1603: 'LDT Export-Zeichensatz
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ExpFor", FKon(PrItm.Value, 2)
Case 1604: 'THEDEX nachrichten
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "TheDex", FKon(PrBol.Value, 2)
Case 1701: 'Popup Benachrichtigungen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "PopFen", FKon(PrBol.Value, 2)
Case 1702: 'Popup Verweildauer (Sek.)
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "PopTim", PrItm.Value
Case 2861: 'Maximale Rezept Zeilenlänge
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "MaxZei", PrItm.Value
Case 1703: 'Microsoft-Word Version
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "WorTre", FKon(PrItm.Value, 2)
Case 2822: 'Doppelklick in Adressenmaske
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "AdDoKl", FKon(PrItm.Value, 2)
Case 1704: 'Chipkartenlesegerät
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ChpCrd", FKon(PrItm.Value, 2)
Case 2850: 'Format der Adressenkurzbezeichnung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "AdrKur", FKon(PrItm.Value, 2)
        S_SeSe 9, FKon(PrItm.Value, 2)
Case 1713: 'ICD-10 Codes einfügen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "ICDEin", FKon(PrBol.Value, 2)
Case 1705: 'PZN bei Rezept einfügen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "RezPZN", FKon(PrBol.Value, 2)
        S_SeSe 34, , , , CBool(PrBol.Value)
Case 2628: 'Betreff bei Rezept einfügen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "RezBet", FKon(PrBol.Value, 2)
Case 1706: 'Rezeptschriftart
        Set PrFnt = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FnNaRz", PrFnt.FontFaceName
        IniSetVal "Layout", "FnGrRz", PrFnt.FontSize
        IniSetVal "Layout", "FnFaRz", PrFnt.Color
Case 1707: 'Doppelklick öffnet Adressmaske
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "AdrDop", FKon(PrBol.Value, 2)
Case 2809: 'Standard MAPI-Ordner
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "OutStM", FKon(PrBol.Value, 2)
Case 2817: 'Startlogo Einblenden
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "StaFor", FKon(PrBol.Value, 2)
Case 1312: 'Wartezimmerliste Schließen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "WaLiCl", FKon(PrBol.Value, 2)
Case 1313: 'IPC3 Nutzung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "IPCAkt", FKon(PrBol.Value, 2)
Case 2975: 'E-Mail Beschleunigung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "MaiBew", FKon(PrBol.Value, 2)
        S_SeSe 103, , , , CBool(PrBol.Value)
Case 2868: 'Netzwerkdrucker auflisten
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "NetDru", FKon(PrBol.Value, 2)
Case 2901: 'Dublexdruck aktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DubDru", FKon(PrBol.Value, 2)
Case 2922: 'Online-Terminbuchungs Sytem PIN Dialog
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsPIN", FKon(PrBol.Value, 2)
        S_SeSe 41, , , , CBool(PrBol.Value)
Case 2941: 'Online-Terminbuchungs Sytem Storno-Dialog
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsSto", FKon(PrBol.Value, 2)
        S_SeSe 21, , , , CBool(PrBol.Value)
Case 2964: 'SMS Terminerinnerung aktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "EriSMS", FKon(PrBol.Value, 2)
        S_SeSe 91, , , , CBool(PrBol.Value)
Case 2963: 'Online-Terminbuchungs Sytem Warteliste-Dialog
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsWar", FKon(PrBol.Value, 2)
        S_SeSe 90, , , , CBool(PrBol.Value)
Case 2928: 'Online-Terminbuchungs System zeige belegte Buchungszeiten
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsBel", FKon(PrBol.Value, 2)
        S_SeSe 53, , , , CBool(PrBol.Value)
Case 2929: 'Online-Terminbuchungs System ICS Datei bei Emailbestätigung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsICS", FKon(PrBol.Value, 2)
        S_SeSe 54, , , , CBool(PrBol.Value)
Case 2965: 'Online-Terminbuchungs System autom. Aktualisierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "OtsAut", FKon(PrBol.Value, 2)
        S_SeSe 92, , , , CBool(PrBol.Value)
Case 2930: 'Online-Terminbuchungs System Schriftart
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsGWF", PrItm.Value
        S_SeSe 55, CStr(PrItm.Value)
Case 2931: 'Online-Terminbuchungs System allgemeine Textfarbe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsFTe", PrItm.Value
        S_SeSe 56, , CLng(PrItm.Value)
Case 2932: 'Online-Terminbuchungs System allgemeine Hintergrundfarbe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsFBa", PrItm.Value
        S_SeSe 57, , CLng(PrItm.Value)
Case 2933: 'Online-Terminbuchungs System allgemeine Textgröße
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsFSi", PrItm.Value
        S_SeSe 58, , CLng(PrItm.Value)
Case 2934: 'Online-Terminbuchungs System Button Hintergrundfarbe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsBBa", PrItm.Value
        S_SeSe 59, , CLng(PrItm.Value)
Case 2935: 'Online-Terminbuchungs System Button Textfarbe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsBTe", PrItm.Value
        S_SeSe 60, , CLng(PrItm.Value)
Case 2936: 'Online-Terminbuchungs System Button Hooverfarbe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsBHo", PrItm.Value
        S_SeSe 61, , CLng(PrItm.Value)
Case 2937: 'Online-Terminbuchungs System Button Deaktiviertfarbe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsBDi", PrItm.Value
        S_SeSe 62, , CLng(PrItm.Value)
Case 2938: '"Online-Terminbuchungs System Link Anschlussseite
        TmStr = PrItm.Value
        If Len(TmStr) > 30 Then TmStr = Left$(TmStr, 30)
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsLas", TmStr
        S_SeSe 63, TmStr
Case 2939: 'Online-Terminbuchungs System Link Datenschutzerklärung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsLin", PrItm.Value
        S_SeSe 64, CStr(PrItm.Value)
Case 2904: 'Druckauftragsseparierung aktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DruSep", FKon(PrBol.Value, 2)
Case 2895: 'Art der Druckvorschau
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "VorTyp", FKon(PrItm.Value, 2)
Case 2903: 'Seitengröße aktivieren
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "SeiGro", FKon(PrItm.Value, 2)
Case 2812: 'Startseitenübersicht Links
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "StaUb1", FKon(PrItm.Value, 2)
Case 2813: 'Startseitenübersicht Mitte
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "StaUb2", FKon(PrItm.Value, 2)
Case 2814: 'Startseitenübersicht Rechts
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "StaUb3", FKon(PrItm.Value, 2)
Case 1710: 'Gliederungsexpansion
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "TreExp", FKon(PrBol.Value, 2)
Case 1715: 'Telefoneingabe mit Landesvorwahl
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "IntTel", FKon(PrBol.Value, 2)
Case 1716: 'Anschrift ohne Anrede
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "AnsAnr", FKon(PrBol.Value, 2)
Case 1718: 'Teilort mit Einfügen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "TeiOrt", FKon(PrBol.Value, 2)
Case 1719: 'Abfrage des Ortsteils
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "OrtAbf", FKon(PrBol.Value, 2)
Case 1720: 'Kurzbezeichnung des Bankkontos einfügen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "KurBnk", FKon(PrBol.Value, 2)
Case 1721: 'Einzelbriefübergabe mit geschützten Formularfeldern
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "WorUbe", FKon(PrBol.Value, 2)
Case 2623: 'Firma als zweite Adressenzeile
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "FirZei", FKon(PrBol.Value, 2)
Case 2633: 'Kommentar bei Einzelbriefübergabe
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BriKom", FKon(PrBol.Value, 2)
Case 1723: 'Geburtstagsliste Filtern
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "AdrGeb", FKon(PrBol.Value, 2)
Case 2808: 'Geburtstag Wiedervorlage
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "WieGeb", FKon(PrBol.Value, 2)
Case 2852: 'Erweiterter BDT Datenimport
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BDTImp", FKon(PrBol.Value, 2)
Case 1724: 'TAPI-Rufnummer übergabe mit Klammern
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "TAPIKl", FKon(PrBol.Value, 2)
Case 2943: 'Mitarbeiternummer als GDT-Dateiname
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "GDTMit", FKon(PrBol.Value, 2)
        S_SeSe 72, , , , CBool(PrBol.Value)
Case 2944: 'GDT-Speicherung ohne Speichern-Dialog
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "GDTDia", FKon(PrBol.Value, 2)
        S_SeSe 73, , , , CBool(PrBol.Value)
Case 2624: 'Überschrift Bemerkungsfeld
        Set PrItm = PrGr1.FindItem(TolId)
        If PrItm.Value <> vbNullString Then IniSetVal "Layout", "AdTit1", PrItm.Value
Case 2625: 'Überschrift Notizenfeld
        Set PrItm = PrGr1.FindItem(TolId)
        If PrItm.Value <> vbNullString Then IniSetVal "Layout", "AdTit2", PrItm.Value
Case 2865: 'Angezeigter Name des GDT Programms
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "GDTApp", PrItm.Value
Case 2866: 'Dateiname der GDT Exportdatei
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "GDTDat", PrItm.Value
        S_SeSe 74, CStr(PrItm.Value)
Case 1722: 'Starten mit Startseite
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "StaSei", FKon(PrBol.Value, 2)
Case 1711: 'Fehlermeldungen zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "GlDbug", FKon(PrBol.Value, 2)
Case 1708: 'Adresseneingabemaske immer im Vordergrund
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "AdrVor", FKon(PrBol.Value, 2)
Case 1712: 'Outlookabgleich nur bei gekennzeichnete Adressen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "OutAdr", FKon(PrBol.Value, 2)
Case 2619: 'Outlookkontakte mit Geburtsdatum
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "OutGeb", FKon(PrBol.Value, 2)
Case 1801: 'Tabellenschriftart
        Set PrFnt = PrGr1.FindItem(TolId)
        If PrFnt.FontSize > 13 Then PrFnt.FontSize = 13
        IniSetVal "Layout", "FntNam", PrFnt.FontFaceName
        IniSetVal "Layout", "FntGro", PrFnt.FontSize
        IniSetVal "Layout", "FntFar", PrFnt.Color
Case 2645: 'Spaltensortierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "SpaSor", FKon(PrBol.Value, 2)
Case 1802: 'Zeilenumbruch
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "GrdZei", FKon(PrBol.Value, 2)
Case 1803: 'Gruppenkopf
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "GrdGkp", FKon(PrBol.Value, 2)
Case 1804: 'Spaltenköpfe zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "SpaKop", FKon(PrBol.Value, 2)
Case 1805: 'Zeilenmarker
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "GrdMkr", FKon(PrBol.Value, 2)
Case 1806: 'Gitternetzlinien
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "GrdGrl", FKon(PrBol.Value, 2)
Case 1807: 'Vorschauzeile
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "GrdPrv", FKon(PrBol.Value, 2)
Case 1808: 'Anzahl der Vorschauzeile
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "KraZei", PrItm.Value
Case 1903: 'Benutzeranmeldung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BenAnm", FKon(PrBol.Value, 2)
        S_SeSe 10, , , , CBool(PrItm.Value)
Case 2824: 'Eigener Mandanten Rechnungsnummernkreis
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "ReMaKr", FKon(PrBol.Value, 2)
        S_SeSe 11, , , , CBool(PrBol.Value)
Case 2842: 'Eigener Mandanten Buchungsnummernkreis
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BuMaKr", FKon(PrBol.Value, 2)
        S_SeSe 12, , , , CBool(PrBol.Value)
Case 2879: 'Umsatzsteuer Splittbuchungen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "UmStSp", FKon(PrBol.Value, 2)
        S_SeSe 47, , , , CBool(PrBol.Value)
Case 2948: 'Einfache Buchführung verwenden
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "EiBuch", FKon(PrBol.Value, 2)
        S_SeSe 77, , , , CBool(PrBol.Value)
        DoEvents
        GlNeK = GlKoX 'Protokolleintrag
        With GlNeK
            .PatNr = GlMan(GlSMa, 2)
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = 104
            .ZiStr = Format$(Now, "hh:mm") & " Uhr"
            .NeuEi = True
            .KeiAk = True
            .Mitar = GlMiA(GlSmI, 2)
            .TeStr = "Einstellungen im Optionsdialog geändert: Einfache Buchführung verwenden"
        End With
        S_Prot
Case 1606: 'SOLL und HABEN Tausch in DATEV Schnittstelle
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BuTaus", FKon(PrBol.Value, 2)
        S_SeSe 7, , , , CBool(PrBol.Value)
Case 2218: 'Separierter Geldkonten Buchungsnummernkreis
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BuGeKr", FKon(PrBol.Value, 2)
        S_SeSe 8, , , , CBool(PrBol.Value)
Case 1906: 'Notizkürzel
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "KonBen", PrItm.Value
Case 2101: 'Allgemeiner Zeilenmarker
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FarZei", PrItm.Value
Case 2102: 'Krankenblatt Zeilenmarker
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "ZeiFar", PrItm.Value
Case 2120: 'Designfensterrahmen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "GUI", "Rahmen", FKon(PrBol.Value, 2)
Case 2121: 'Luna Farben
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "GUI", "LunCol", FKon(PrBol.Value, 2)
Case 2897: 'Menüanimation einschalten
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "MenAni", FKon(PrBol.Value, 2)
Case 2649: 'Popup-Kalenderfeld
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "PopKal", FKon(PrBol.Value, 2)
Case 2123: 'ClearType Textqualität
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "CleTyp", FKon(PrBol.Value, 2)
Case 2124: 'Fensterpositionen Speichern
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "DocLay", FKon(PrBol.Value, 2)
Case 2131: 'Multiselektion
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "MulMar", FKon(PrBol.Value, 2)
Case 2871: 'Infotabellen auf Startbildschirm zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "StaTab", FKon(PrBol.Value, 2)
Case 2873: 'Autom. Erweitern des Mail-Mitarbeiterordners
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "MaMiOr", FKon(PrBol.Value, 2)
Case 2830: 'Sicherer Konfigurationsmodus
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "IdiMod", FKon(PrBol.Value, 2)
Case 2886: 'Runder Systembutton
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "SysBut", FKon(PrBol.Value, 2)
Case 2905: 'Steuersatzspalte
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "KrSteu", FKon(PrBol.Value, 2)
        S_SeSe 23, , , , CBool(PrBol.Value)
Case 2887: 'Farbige Register
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FarReg", FKon(PrBol.Value, 2)
Case 2888: 'Farbige Modulkennzeichnung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FarMod", FKon(PrBol.Value, 2)
Case 2898: 'Bildschirm -Aktualisierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "BilAkt", FKon(PrBol.Value, 2)
Case 2985: 'Krankenblattdialog im Vordergrund
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "GUI", "KrFoVo", FKon(PrBol.Value, 2)
Case 2126: 'Splashscreen deaktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "NoSpla", FKon(PrBol.Value, 2)
Case 2857: 'Fenstergröße beim Progreammstart
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FenGro", FKon(PrItm.Value, 2)
Case 2899: 'Symbolleisten Schrifthöhe
        Set PrItm = PrGr1.FindItem(TolId)
        If PrItm.Value > 13 Then PrItm.Value = 13
        IniSetVal "Layout", "TolFoH", PrItm.Value
Case 2914: 'Fensterbreite
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FeVoBr", PrItm.Value
Case 2915: 'Fensterhöhe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FeVoHo", PrItm.Value
Case 2201: 'Dynamische Datumsanpassung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "AktDat", FKon(PrBol.Value, 2)
Case 2125: 'Katalogdiagnosensortierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DiaSor", FKon(PrBol.Value, 2)
Case 2952: 'Zahlung Eintrag Gesamtbetrag
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "KraZah", FKon(PrBol.Value, 2)
Case 2883: 'Auf Nachfrage neue Rechnung anlegen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "RechAu", FKon(PrBol.Value, 2)
Case 2896: 'Dokumentiert Emails in Krankenblatt
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "EmlKra", FKon(PrBol.Value, 2)
        S_SeSe 71, , , , CBool(PrBol.Value)
Case 2884: 'Max. Anzahl Kalenderwahl
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "MaxKal", PrItm.Value
Case 1605: 'Internen Bild- und PDF Viewer verwenden
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "BldVie", FKon(PrItm.Value, 2)
Case 2859: 'Alle Eintrüge beim Einfügen einer Kette markieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "KetMar", FKon(PrBol.Value, 2)
Case 2847: 'Krankenblattdiagnosen autom. in Abrechnung übernehmen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "KatDia", FKon(PrBol.Value, 2)
Case 2202: 'Spaltenköpfe Anzeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "SpalUb", FKon(PrBol.Value, 2)
Case 2203: 'Zeilenmarker zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "ZeiMar", FKon(PrBol.Value, 2)
Case 2205: 'Gitternetzlinien zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "LinTyp", FKon(PrBol.Value, 2)
Case 2206: 'Farbunterscheidung zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "RowCol", FKon(PrBol.Value, 2)
Case 2208: 'Steigerungsfaktor zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FakAnz", FKon(PrBol.Value, 2)
Case 2214: 'Zeitspalte zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "ZeiAnz", FKon(PrBol.Value, 2)
Case 2215: 'Bearbeitungsmodus bei Klicken
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "KraEdi", FKon(PrBol.Value, 2)
Case 2216: 'Fokusplazierung im Krankenblatt
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "KraFoc", FKon(PrBol.Value, 2)
Case 2820: 'Dokumenten Protokollierung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DocPro", FKon(PrBol.Value, 2)
Case 2923: 'Gebührendiagnosezuordnungen auslassen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "GeDiZu", FKon(PrBol.Value, 2)
Case 2984: 'Cave Infotext zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "CavTex", FKon(PrBol.Value, 2)
Case 2962: 'Krankenblatteinträge aller Mitarbeiter sichtbar machen
'        Set PrBol = PrGr1.FindItem(TolId)
'        IniSetVal "System", "MaVoKr", FKon(PrBol.Value, 2)
Case 2992: 'Cursorplatzierung bei Bearbeiten eines Eintrags
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "EinCur", FKon(PrBol.Value, 2)
Case 2851: 'Automatisches Umbenennen von Bildern und Dokumenten
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "BilUmb", FKon(PrBol.Value, 2)
Case 2875: 'Mandantenspalte im Abrechnungsmodul
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "ManSpa", FKon(PrBol.Value, 2)
Case 2918: 'Mandantenbezogene Vorgabenbenutzung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "ManVor", FKon(PrBol.Value, 2)
        S_SeSe 35, , , , CBool(PrBol.Value)
Case 2940: 'Mandantenbezogene Datenbegrenzung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Vorgabe", "StMaRe", FKon(PrBol.Value, 2)
        S_SeSe 66, , , , CBool(PrBol.Value)
Case 2210: 'Eingabesortierung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "ZifAuf", FKon(PrItm.Value, 2)
Case 2212: 'Vorgabe für Krankenblatteintrag
        Set PrItm = PrGr1.FindItem(TolId)
        For AktZa = 1 To UBound(GlKrA)
            If GlKrA(AktZa, 0) > 9 Then
                Select Case GlKrA(AktZa, 0)
                Case 24:    'Textdokumente
                Case 101:   'Beleg / Rezept
                Case 102:   'Datei
                Case 104:   'Protokoll
                Case 105:   'Bilddatei
                Case Else:
                    If GlKrA(AktZa, 2) = PrItm.Value Then
                        IniSetVal "System", "EinTyp", AktZa
                        Exit For
                    End If
                End Select
            End If
        Next AktZa
Case 2217: 'Krankenblattschriftart
        Set PrFnt = PrGr1.FindItem(TolId)
        If PrFnt.FontSize > 13 Then PrFnt.FontSize = 13
        IniSetVal "Layout", "KraFnt", PrFnt.FontFaceName
        IniSetVal "Layout", "KraGro", PrFnt.FontSize
Case 2860: 'Emailschriftart
        Set PrFnt = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "MaiFnt", PrFnt.FontFaceName
        IniSetVal "Layout", "MaiGro", PrFnt.FontSize
        IniSetVal "Layout", "MaiFar", PrFnt.Color
Case 2601: 'Arbeitstag Beginn
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "StaZei", PrItm.Value
Case 2602: 'Arbeitstag Ende
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "EndZei", PrItm.Value
Case 2629: 'Start der Skalierungsanzeige
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "AnsSta", PrItm.Value
Case 2630: 'Ende der Skalierungsanzeige
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "AnsEnd", PrItm.Value
Case 2603: 'Kalender Zeitscala
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TimSca", PrItm.Value
Case 2609: 'Zeitscala mit Minutenanzeige
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ScaMin", FKon(PrBol.Value, 2)
Case 2942: 'Terminbelegungswarnung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerWar", FKon(PrBol.Value, 2)
Case 2983: 'Kalenderfocus beim Tages- oder Wochenwechsel
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeKaFo", FKon(PrBol.Value, 2)
Case 2945: 'Terminnachricht auch an BCC
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerBCC", FKon(PrBol.Value, 2)
        S_SeSe 81, , , , CBool(PrBol.Value)
Case 1607: 'Starre Termintaktung verwenden
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "StaRas", FKon(PrBol.Value, 2)
        S_SeSe 29, , , , CBool(PrBol.Value)
Case 2966: 'Anzeige stornierter Termine in den Termindetails
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ZeStTe", FKon(PrBol.Value, 2)
        S_SeSe 93, , , , CBool(PrBol.Value)
Case 2967: 'TSE Kassenname
        Set PrItm = PrGr1.FindItem(TolId)
        TmStr = PrItm.Value
        If GlRDP = False Then
            IniSetVal "System", "TSEKas", TmStr
            S_SeSe 94, TmStr
        End If
Case 2968: 'TSE Laufwerk
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "TSEDrv", PrItm.Value
        S_SeSe 95, CStr(PrItm.Value)
Case 2969: 'TSE Aktiviert
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "TSEAkt", FKon(PrBol.Value, 2)
        S_SeSe 96, , , , CBool(PrBol.Value)
Case 2970: 'TSE Verfahren
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "ClouID", FKon(PrItm.Value, 2)
        S_SeSe 97, FKon(PrItm.Value, 2)
Case 2869: 'Mitarbeiterplan anstelle Mandantenplan
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "MitPla", FKon(PrBol.Value, 2)
        S_SeSe 30, , , , CBool(PrBol.Value)
Case 2604: 'Terminzeit als Uhr
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TimClo", FKon(PrBol.Value, 2)
Case 2605: 'TEndzeit anzeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "EndTim", FKon(PrBol.Value, 2)
Case 2872: 'Keine Farbunterscheidung in Raumbelegung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "KeFaRa", FKon(PrBol.Value, 2)
Case 2606: 'Font Glyphs
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "FntGly", FKon(PrBol.Value, 2)
Case 2627: 'Terminzeit aus dem Terminbetreff verwenden
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerZei", FKon(PrBol.Value, 2)
        S_SeSe 52, , , , CBool(PrBol.Value)
Case 2917: 'Terminnachricht mit ICS Dateiversand
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ICSFil", FKon(PrBol.Value, 2)
        S_SeSe 69, , , , CBool(PrBol.Value)
Case 2810: 'Ganztagstermine nicht berücksichtigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "GanzBe", FKon(PrBol.Value, 2)
Case 2821: 'Sprechzeiten nicht berücksichtigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "SpreBe", FKon(PrBol.Value, 2)
Case 2991: 'Abfrage E-Mail-Terminerinnerung bei Terminverschiebung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeErSp", FKon(PrBol.Value, 2)
Case 2612: 'Mindestbreite der Zeitspalten im Kalender in Pixel
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TiSpBr", PrItm.Value
Case 2613: 'Kalenderdruck mit horizontaler Druckausrichtung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TiPrHo", FKon(PrBol.Value, 2)
Case 2614: 'Scrolleisre für Monatsansicht
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "KalScl", FKon(PrBol.Value, 2)
Case 2811: 'Keine privaten Outloktermine
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OutPri", FKon(PrBol.Value, 2)
Case 2616: 'Intervall Kalenderaktualisierung in Sek.
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TimInt", PrItm.Value
Case 2617: 'Erweiterte Termininformationen
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerErw", FKon(PrBol.Value, 2)
Case 2607: 'Druckbereich Beginn
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "StaPri", PrItm.Value
Case 2608: 'Druckbereich Ende
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "EndPri", PrItm.Value
Case 2631: '"Betreff Trennzeichen
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TreZei", PrItm.Value
Case 2654: 'Mitarbeitername in Terminort speichern
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeOrMi", FKon(PrBol.Value, 2)
        S_SeSe 70, , , , CBool(PrBol.Value)
Case 2632: 'Telefonnummer in Terminbetreff
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ZeiBet", FKon(PrBol.Value, 2)
Case 2636: 'Terminerinnerung im Kalender
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerErn", FKon(PrBol.Value, 2)
Case 2639: 'Schnelle Terminliste zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerUbe", FKon(PrBol.Value, 2)
Case 2646: 'Virtuelle Endlostermine
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "EndSer", FKon(PrBol.Value, 2)
Case 2834: 'Anzeigen der Sprechzeiten
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ManZei", FKon(PrBol.Value, 2)
Case 2836: 'Mandanten Farbkennzeichnung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ManFar", FKon(PrBol.Value, 2)
Case 2837: 'Raum Farbkennzeichnung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "RauFar", FKon(PrBol.Value, 2)
Case 2838: 'Geschlecht Darstellen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerGes", FKon(PrBol.Value, 2)
Case 2839: 'Termintiteldarstellung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeFnBl", FKon(PrBol.Value, 2)
Case 2951: 'Patienten-CAVE ins Kommentarfeld
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerCav", FKon(PrBol.Value, 2)
Case 2919: 'Terminmarkierung bei Terminzetteldruck
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "AuMark", FKon(PrBol.Value, 2)
Case 2648: 'Standard-Gebührenkette mit Terminlänge multiplizieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "LeiAnz", FKon(PrBol.Value, 2)
Case 2840: 'Raumzuordnung numerisch sortieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "RmuSor", FKon(PrBol.Value, 2)
        S_SeSe 36, , , , CBool(PrBol.Value)
Case 2816: 'Alternative Berechnung der Serienfälligkeit
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "SerBer", FKon(PrBol.Value, 2)
Case 2880: 'Berücksichtigen von Aktontozahlungen bei Terminleistungen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeLeAk", FKon(PrBol.Value, 2)
Case 2920: 'Terminbetreff-Kompatibilitätsmodus
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerBet", FKon(PrBol.Value, 2)
Case 2989: 'Verkehrsnamen in Mitarbeiterplan
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "KalMit", FKon(PrBol.Value, 2)
Case 2924: 'Terminleistungen nur für passene Referenzrechnung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ReTeSe", FKon(PrBol.Value, 2)
Case 2990: 'Name inaktiver  Mitarbeiter in den Termindetails
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeInMi", FKon(PrBol.Value, 2)
Case 2921: 'Mitarbeiter / Mandanten in Raumbelegung aktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "RauMan", FKon(PrBol.Value, 2)
Case 2301: 'Verwendeter Datenbanktyp
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DBaTyp", FKon(PrItm.Value, 2)
Case 2302: 'Datenbanktimeout (Sek.)
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DatAkt", PrItm.Value
Case 2303: 'SQL-Server Name
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DatSer", PrItm.Value
Case 2304: 'SQL-Server Datenbank
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DatTab", PrItm.Value
Case 2305: 'SQL-Server Verbindung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DatVer", FKon(PrItm.Value, 2)
Case 2306: 'Asynchrone Verbindung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DatLad", FKon(PrBol.Value, 2)
Case 2307: 'Jet Sofortspeicherung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "JetSav", FKon(PrBol.Value, 2)
Case 2308: 'Kennwort verschlüsselt
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DaVers", FKon(PrBol.Value, 2)
Case 2309: 'SQL-Server Benutzername
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "DatUsr", PrItm.Value
Case 2310: '"SQL-Server Benutzerpasswort
        Set PrItm = PrGr1.FindItem(TolId)
        If GlVrs = True Then 'Kennwort verschlüsselt
            IniSetVal "System", "DaUsPa", SCrypt(PrItm.Value, True)
        Else
            IniSetVal "System", "DaUsPa", PrItm.Value
        End If
Case 2311: 'Verwendeter Datenbanktyp
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "MatTyp", FKon(PrItm.Value, 2)
Case 2312: 'Datenbanktimeout (Sek.)
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "MatAkt", PrItm.Value
Case 2313: 'SQL-Server Name
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "MatSer", PrItm.Value
Case 2314: 'SQL-Server Datenbank
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "MatTab", PrItm.Value
Case 2315: 'SQL-Server Verbindung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "MatVer", FKon(PrItm.Value, 2)
Case 2316: 'Asynchrone Verbindung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "DatLad", FKon(PrBol.Value, 2)
Case 2317: 'Kennwort verschlüsselt
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "MaVers", FKon(PrBol.Value, 2)
Case 2318: 'SQL-Server Benutzername
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "MatUsr", PrItm.Value
Case 2319: '"SQL-Server Benutzerpasswort
        Set PrItm = PrGr1.FindItem(TolId)
        If GlVrs = True Then 'Kennwort verschlüsselt
            IniSetVal "System", "MaUsPa", SCrypt(PrItm.Value, True)
        Else
            IniSetVal "System", "MaUsPa", PrItm.Value
        End If
Case 2401: 'HotTracking
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "HotTra", FKon(PrBol.Value, 2)
Case 2402: 'Systemordner zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "SpeFol", FKon(PrBol.Value, 2)
Case 2403: 'Versteckte Dateien zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "HidItm", FKon(PrBol.Value, 2)
Case 2404: 'Gitternetzlinien
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FilGrd", FKon(PrBol.Value, 2)
Case 2405: 'Ordner in Dateiansicht zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "OrdZei", FKon(PrBol.Value, 2)
Case 2406: 'Ordnerverbindungslinien zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "OrdLin", FKon(PrBol.Value, 2)
Case 2407: 'Dokumentenfilter
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "Layout", "FilFil", FKon(PrBol.Value, 2)
Case 2832: 'Authentifizierungstyp
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "AutArt", FKon(PrItm.Value, 2)
Case 2827: 'Emaildialog zeigen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "EmlDia", FKon(PrBol.Value, 2)
Case 2867: 'Absteigende Sortierung in Gruppen
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "System", "EmlSor", FKon(PrBol.Value, 2)
Case 2858: 'Online-Terminbuchungs Sytem Aktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TerSyn", FKon(PrBol.Value, 2)
        S_SeSe 13, , , , CBool(PrBol.Value)
Case 2916: 'Online-Terminbuchungs Sytem Mitarbeiterwahl
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OTSRei", FKon(PrBol.Value, 2)
        S_SeSe 14, , , , CBool(PrBol.Value)
Case 2953: 'Online-Terminbuchungs Sytem Adressenerfassung
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsAdr", FKon(PrBol.Value, 2)
        S_SeSe 80, , , , CBool(PrBol.Value)
Case 2652: 'Online-Terminbuchungs Benutzername
        Set PrItm = PrGr1.FindItem(TolId)
        If SNaFi(PrItm.Value) = PrItm.Value Then
            IniSetVal "TerSys", "OTSUse", PrItm.Value
            S_SeSe 15, CStr(PrItm.Value)
        Else
            WindowMess "Der von Ihnen geänderte Eintrag darf keine Umlaute oder Sonderzeichen enthalten!", Dial2, "Keine Sonderzeichen", Me.hwnd
        End If
Case 2653: 'Online-Terminbuchungs Sytem Passwort
        Set PrItm = PrGr1.FindItem(TolId)
        S_SeSe 16, SCrypt(PrItm.Value, True)
        If SNaFi(PrItm.Value) = PrItm.Value Then
            IniSetVal "TerSys", "OTSPas", SCrypt(PrItm.Value, True)
        Else
            WindowMess "Der von Ihnen geänderte Eintrag darf keine Umlaute oder Sonderzeichen enthalten!", Dial2, "Keine Sonderzeichen", Me.hwnd
        End If
Case 2853: 'Online-Terminbuchungs Provider
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "DomNam", PrItm.Value
        S_SeSe 17, CStr(PrItm.Value)
Case 2893: 'CalDAV / CardDAV / Exchange Synchronisation
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ExTeSy", FKon(PrBol.Value, 2)
        S_SeSe 18, , , , CBool(PrBol.Value)
Case 2881: 'Proxyserver Verwenden
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ProxVe", FKon(PrBol.Value, 2)
        S_SeSe 42, , , , CBool(PrBol.Value)
Case 2801: 'Standardschriftart
        Set PrFnt = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "FntNam", PrFnt.FontFaceName
        IniSetVal "TexVer", "FntGro", PrFnt.FontSize
        IniSetVal "TexVer", "FntFar", PrFnt.Color
Case 2802: 'Linker Rand
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "RnLink", PrItm.Value
Case 2803: 'Rechter Rand
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "RnRech", PrItm.Value
Case 2804: 'Oberer Rand
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "RnOben", PrItm.Value
Case 2805: 'Unterer Rand
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "RnUnte", PrItm.Value
Case 2806: 'Standard Seitenhöhe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "SeiHoh", PrItm.Value
Case 2807: 'Standard Seitenhöhe
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "SeiBre", PrItm.Value
Case 2651: 'Automatische Microsoft-Word Konvertierung"
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "WorKon", FKon(PrItm.Value, 2)
Case 2864: 'Krankenblattpositionen mit Datum
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TexVer", "KrTyDa", FKon(PrBol.Value, 2)
Case 2882: 'Proxyserver Name
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "ProxNa", PrItm.Value
        S_SeSe 43, CStr(PrItm.Value)
Case 2910: 'SMS Account-ID
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "SMSGat", PrItm.Value
        S_SeSe 37, CStr(PrItm.Value)
Case 2912: 'Online-Terminbuchungs System Link für Impressum
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsImp", PrItm.Value
        S_SeSe 39, CStr(PrItm.Value)
Case 2911: 'SMS Absenderkennung
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "SMSAbs", PrItm.Value
        S_SeSe 38, CStr(PrItm.Value)
Case 2913: 'SMS Identifikationskey
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "SMSKey", PrItm.Value
        S_SeSe 40, CStr(PrItm.Value)
Case 2954: 'Terminland WebCAL (ICS) aktivieren
        Set PrBol = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeLaAk", FKon(PrBol.Value, 2)
        S_SeSe 82, , , , CBool(PrBol.Value)
Case 2955: 'Neuaufnahmeformular-Webadresse
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "NeAuFo", PrItm.Value
        S_SeSe 83, CStr(PrItm.Value)
Case 2956: 'Terminland Benutzername
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeLaUs", PrItm.Value
        S_SeSe 84, CStr(PrItm.Value)
Case 2957: 'Terminland Passwort
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "TeLaPa", SCrypt(PrItm.Value, True)
        S_SeSe 85, SCrypt(PrItm.Value, True)
Case 2977: 'SMTP Praxisname
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "SmPrNa", SNaFi(PrItm.Value, True, True, True)
        S_SeSe 105, SNaFi(PrItm.Value, True, True, True)
Case 2978: 'SMTP IP-Adresse
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "SmIPAd", PrItm.Value
        S_SeSe 106, CStr(PrItm.Value)
Case 2979: 'SMTP SocksProxyServer
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "System", "SmPrSe", PrItm.Value
        S_SeSe 107, CStr(PrItm.Value)
Case 2980: 'SMTP SocksProxyPort
        Set PrItm = PrGr1.FindItem(TolId)
        If IsNumeric(PrItm.Value) = True Then
            IniSetVal "System", "SmPrPo", CInt(PrItm.Value)
            S_SeSe 108, , CInt(PrItm.Value)
        End If
Case 2981: 'Online-Terminbuchungs Sytem Stornierte Entfernen
        Set PrItm = PrGr1.FindItem(TolId)
        IniSetVal "TerSys", "OtsLoe", PrBol.Value
        S_SeSe 109, , , , CBool(PrBol.Value)
End Select

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRgSv " & Err.Number
Resume Next

End Sub
Private Sub FSave()
On Error GoTo MeErr

Dim TmFar As Long 'Defaultfarbwert
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim TmBoD As Boolean 'Deufault Boolean Wert
Dim TmBoV As Boolean 'Value Boolean Wert
Dim NeuSt As Boolean
Dim IdxZa As Integer
Dim AktZa As Integer
Dim Frage As Integer
Dim AryIt() As String
Dim Mld1, Tit1 As String

Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set TxPLZ = Me.txtFePLZ
Set TxOrt = Me.txtFeOrt
Set TxBLZ = Me.txtFeBLZ
Set TxBnk = Me.txtFeBnk
Set TrLi1 = Me.trvList1
Set PrGr1 = Me.prpGrid1
Set PrGr2 = Me.prpGrid2
Set RpCon = Me.repCont2
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows
Set PrIts = PrGr1.Categories

TeTit = "Programmneustart"
TeMai = "Möchten Sie das Programm jetzt neu starten?"
TeInh = "Damit die vorgenommenen Einstellungen wirksam werden, muss das Programm neu gestartet werden."
TeFus = "Bei einem Programmneustart werden alle Daten gespeichert und die Verdingung zur Datenbank neu aufgebaut."

Select Case LiTyp
Case 1: NeuSt = True
        For Each PrKat In PrIts
            For Each PrItm In PrKat.Childs
                Select Case PrItm.Type
                Case PropertyItemCategory:
                
                Case PropertyItemString:
                    If PrItm.defaultValue <> PrItm.Value Then
                        FRgSv PrItm.id
                    End If
                Case PropertyItemNumber:
                    If PrItm.defaultValue <> PrItm.Value Then
                        FRgSv PrItm.id
                    End If
                Case PropertyItemBool:
                    Set PrBol = PrItm
                    TmBoD = PrBol.defaultValue
                    TmBoV = PrBol.Value
                    If TmBoD <> TmBoV Then
                        FRgSv PrBol.id
                    End If
                Case PropertyItemColor:
                    AryIt = Split(PrItm.defaultValue, ";")
                    TmFar = RGB(AryIt(0), AryIt(1), AryIt(2))
                    If TmFar <> PrItm.Value Then
                        FRgSv PrItm.id
                    End If
                Case PropertyItemFont:
                    Set PrFnt = PrItm
                    If NeFn1 = True Then FRgSv PrFnt.id
                    If NeFn2 = True Then FRgSv PrFnt.id
                    If NeFn3 = True Then FRgSv PrFnt.id
                    If NeFn4 = True Then FRgSv PrFnt.id
                    If NeFn5 = True Then FRgSv PrFnt.id
                Case PropertyItemDate:
                
                End Select
            Next PrItm
        Next PrKat
Case 2: NeuSt = True
        Opt_Sav TrLi1.SelectedItem.Key
Case 3: NeuSt = False
        If TxBLZ.Text <> vbNullString Then
            Opt_BLs TxBLZ.Text, 1
        ElseIf TxBnk.Text <> vbNullString Then
            Opt_BLs TxBnk.Text, 2
        End If
Case 4: NeuSt = True
        If TxPLZ.Text <> vbNullString Then
            Opt_PLs TxPLZ.Text, 1
        ElseIf TxOrt.Text <> vbNullString Then
            Opt_PLs TxOrt.Text, 2
        End If
Case 5: NeuSt = True
        Opt_Grp
End Select

Set TrLi1 = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

If NeuSt = True Then
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
    If GlMes = 33565 Then
        GlRes = True 'Reset der Einstellungen
        Unload Me
        DoEvents
        Unload frmMain
    Else
        SNeSt True
        Unload Me
    End If
Else
    Unload Me
End If

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmAcs As XtremeCommandBars.CommandBarActions

Set FM = frmOptions
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

LiTyp = TaIdx + 1

If OpLad = False Then
    Screen.MousePointer = vbHourglass
    clFen.FenDsk 2
    
    Select Case TaIdx
    Case 0:
        Rahm1.Visible = True
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        CmAcs(SY_OP_Hinzufuegen).Enabled = False
        CmAcs(SY_OP_Loeschen).Enabled = False
        CmAcs(SY_OP_Speichern).Enabled = True
    Case 1:
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = False
        CmAcs(SY_OP_Hinzufuegen).Enabled = True
        CmAcs(SY_OP_Loeschen).Enabled = False
        CmAcs(SY_OP_Speichern).Enabled = True
    Case 2:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = True
        Rahm4.Visible = False
        Rahm5.Visible = False
        CmAcs(SY_OP_Hinzufuegen).Enabled = False
        CmAcs(SY_OP_Loeschen).Enabled = False
        CmAcs(SY_OP_Speichern).Enabled = True
    Case 3:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = True
        Rahm5.Visible = False
        CmAcs(SY_OP_Hinzufuegen).Enabled = False
        CmAcs(SY_OP_Loeschen).Enabled = False
        CmAcs(SY_OP_Speichern).Enabled = True
    Case 4:
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
        Rahm5.Visible = True
        CmAcs(SY_OP_Hinzufuegen).Enabled = False
        CmAcs(SY_OP_Loeschen).Enabled = False
        CmAcs(SY_OP_Speichern).Enabled = True
    End Select
    
    clFen.FenDsk 3
    Screen.MousePointer = vbNormal
End If

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

Select Case TolId
Case SY_OP_Hinzufuegen: Opt_Neu
Case SY_OP_Loeschen: FLoe
Case SY_OP_Speichern: FSave
Case SY_OP_Reset: FPfad
Case SY_OP_Abbruch: Unload Me
Case SY_OP_Hilfe: FHilfe
Case KY_F8: FSave
Case KY_F11: Unload FM
Case KY_F1: FHilfe
Case KY_CT_AL_R: FHelp
End Select

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
Private Sub comBar02_Resize()
    FPosi
End Sub

Private Sub prpGrid1_AfterEdit(ByVal Item As XtremePropertyGrid.IPropertyGridItem, NewValue As String, Cancel As Boolean)

If Item.Type = PropertyItemFont Then
    If NewValue <> Item.defaultValue Then
        Select Case Item.id
        Case 1706: NeFn1 = True
        Case 1801: NeFn2 = True
        Case 2217: NeFn3 = True
        Case 2860: NeFn4 = True
        Case 2801: NeFn5 = True
        End Select
    End If
End If

End Sub

Private Sub prpGrid1_InplaceButtonDown(ByVal Button As XtremePropertyGrid.IPropertyGridInplaceButton, Cancel As Variant)

Dim FiNam As Variant

FiNam = FDial(Button.Item.id)

If FiNam <> vbNullString Then
    Button.Item.Value = FiNam
End If

End Sub
Private Sub prpGrid1_ValueChanged(ByVal Item As XtremePropertyGrid.IPropertyGridItem)
On Error Resume Next

Dim TmFar As Long    'Defaultfarbwert
Dim TmBoD As Boolean 'Deufault Boolean Wert
Dim TmBoV As Boolean 'Value Boolean Wert
Dim AryIt() As String

Set PrGr1 = Me.prpGrid1
Set PrIts = PrGr1.Categories

For Each PrKat In PrIts
    For Each PrItm In PrKat.Childs
        Select Case PrItm.Type
        Case PropertyItemCategory:
        
        Case PropertyItemString:
            If PrItm.defaultValue <> PrItm.Value Then
                FRgSv PrItm.id
            End If
        Case PropertyItemNumber:
            If PrItm.defaultValue <> PrItm.Value Then
                FRgSv PrItm.id
            End If
        Case PropertyItemBool:
            Set PrBol = PrItm
            TmBoD = PrBol.defaultValue
            TmBoV = PrBol.Value
            If TmBoD <> TmBoV Then
                FRgSv PrBol.id
            End If
        Case PropertyItemColor:
            AryIt = Split(PrItm.defaultValue, ";")
            TmFar = RGB(AryIt(0), AryIt(1), AryIt(2))
            If TmFar <> PrItm.Value Then
                FRgSv PrItm.id
            End If
        Case PropertyItemFont:
            Set PrFnt = PrItm
            If NeFn1 = True Then FRgSv PrFnt.id
            If NeFn2 = True Then FRgSv PrFnt.id
            If NeFn3 = True Then FRgSv PrFnt.id
            If NeFn4 = True Then FRgSv PrFnt.id
            If NeFn5 = True Then FRgSv PrFnt.id
        Case PropertyItemDate:
        
        End Select
    Next PrItm
Next PrKat

End Sub

Private Sub repCont2_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Set TrLi1 = Me.trvList1

If TrLi1.SelectedItem.Key = "K02" Then
    If Item.Index = 4 Then
        If Row.Record.Item(3).Value <> vbNullString Then
            Item.BackColor = Row.Record.Item(3).Value
            Item.ForeColor = Row.Record.Item(3).Value
        End If
    End If
ElseIf TrLi1.SelectedItem.Key = "K11" Then
    If Item.Index = 4 Then
        If Row.Record.Item(3).Value <> vbNullString Then
            Item.BackColor = Row.Record.Item(3).Value
            Item.ForeColor = Row.Record.Item(3).Value
        End If
    End If
ElseIf TrLi1.SelectedItem.Key = "K14" Then
    If Item.Index = 4 Then
        If Row.Record.Item(3).Value <> vbNullString Then
            Item.BackColor = Row.Record.Item(3).Value
            Item.ForeColor = Row.Record.Item(3).Value
        End If
    End If
ElseIf TrLi1.SelectedItem.Key = "K18" Then
    If Item.Index = 4 Then
        If Row.Record.Item(3).Value <> vbNullString Then
            Item.BackColor = Row.Record.Item(3).Value
            Item.ForeColor = Row.Record.Item(3).Value
        End If
    End If
End If

End Sub
Private Sub repCont2_InplaceButtonDown(ByVal Button As XtremeReportControl.IReportInplaceButton)
    If Button.Column.ItemIndex = 4 Then
        FCol
    End If
End Sub

Private Sub repCont2_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String
Dim TreKy As String
Dim TmStr As String
Dim RowNr As Integer
Dim Knots As XtremeSuiteControls.TreeViewNodes
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmOptions
Set TrLi1 = FM.trvList1
Set RpCon = FM.repCont2
Set Knots = TrLi1.Nodes
Set RpRws = RpCon.Rows
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

For Each Knote In Knots
    If Knote.Selected = True Then
        TreKy = Knote.Key
        Exit For
    End If
Next Knote

If TreKy = "K25" Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        TmStr = vbNullString
        Set RpCol = RpCls.Find(5) 'Mo
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            TmStr = "1"
        Else
            TmStr = "2"
        End If
        Set RpCol = RpCls.Find(6) 'Di
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            TmStr = TmStr & "1"
        Else
            TmStr = TmStr & "2"
        End If
        Set RpCol = RpCls.Find(7) 'Ni
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            TmStr = TmStr & "1"
        Else
            TmStr = TmStr & "2"
        End If
        Set RpCol = RpCls.Find(8) 'Do
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            TmStr = TmStr & "1"
        Else
            TmStr = TmStr & "2"
        End If
        Set RpCol = RpCls.Find(9) 'Fr
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            TmStr = TmStr & "1"
        Else
            TmStr = TmStr & "2"
        End If
        Set RpCol = RpCls.Find(10) 'Sa
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            TmStr = TmStr & "1"
        Else
            TmStr = TmStr & "2"
        End If
        Set RpCol = RpCls.Find(11) 'So
        If RpRow.Record(RpCol.ItemIndex).Checked = True Then
            TmStr = TmStr & "1"
        Else
            TmStr = TmStr & "2"
        End If
        DoEvents
        RpRow.Record(4).Value = TmStr
        TmTag = RpRow.Record(4).Tag
        If TmTag <> vbNullString Then
            TmTag = Mid$(TmTag, 2, Len(TmTag) - 1)
            RpRow.Record(4).Tag = "@" & TmTag
        End If
    End If
ElseIf TreKy = "K33" Then 'Zahlungstexte
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        RowNr = RpRow.Index
        If RpRow.GroupRow = False Then
            For Each RpRow In RpRws
                RpRow.Record(3).Checked = False
                DoEvents
                TmTag = RpRow.Record(3).Tag
                If TmTag <> vbNullString Then
                    TmTag = Mid$(TmTag, 2, Len(TmTag) - 1)
                    RpRow.Record(3).Tag = "@" & TmTag
                End If
            Next RpRow
            DoEvents
            RpRws(RowNr).Record(3).Checked = True
            DoEvents
            Set RpRow = RpRws(RowNr)
            TmTag = RpRow.Record(3).Tag
            If TmTag <> vbNullString Then
                TmTag = Mid$(TmTag, 2, Len(TmTag) - 1)
                RpRow.Record(3).Tag = "@" & TmTag
            End If
        End If
    End If
Else
    TmTag = Item.Tag
    If TmTag <> vbNullString Then
        TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
        Item.Tag = "@" & TmTag
    End If
End If

Opt_Sav TrLi1.SelectedItem.Key

Set RpRws = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

End Sub
Private Sub repCont2_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String
Dim TmpTg As String
Dim TreKy As String
Dim AktZa As Integer
Dim ZaZil As Integer
Dim RowNr As Integer
Dim IdxZa As Integer
Dim Mahnb As Boolean
Dim Knots As XtremeSuiteControls.TreeViewNodes
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmOptions
Set TrLi1 = FM.trvList1
Set RpCon = FM.repCont2
Set Knots = TrLi1.Nodes
Set RpRws = RpCon.Rows
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

For Each Knote In Knots
    If Knote.Selected = True Then
        TreKy = Knote.Key
        Exit For
    End If
Next Knote

TmpTg = Item.Tag
TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
Item.Tag = "@" & TmTag
DoEvents

If TreKy = "K05" Then 'Zahlungsziele
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(1)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                ZaZil = CInt(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                ZaZil = 0
            End If
            Set RpCol = RpCls.Find(3)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                Mahnb = CBool(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                Mahnb = 0
            End If
            Select Case Column.Index
            Case 3:
                If Mahnb = True Then
                    If ZaZil <= 1 Then
                        RpRow.Record(1).Value = 3
                        TmTag = RpRow.Record(1).Tag
                        TmTag = Mid$(TmTag, 2, Len(TmTag) - 1)
                        RpRow.Record(1).Tag = "@" & TmTag
                    End If
                Else
                    If ZaZil > 0 Then
                        RpRow.Record(1).Value = 0
                        TmTag = RpRow.Record(1).Tag
                        TmTag = Mid$(TmTag, 2, Len(TmTag) - 1)
                        RpRow.Record(1).Tag = "@" & TmTag
                    End If
                End If
            Case 1:
                If ZaZil = 0 Then
                    RpRow.Record(3).Value = True
                    TmTag = RpRow.Record(1).Tag
                    TmTag = Mid$(TmTag, 2, Len(TmTag) - 1)
                    RpRow.Record(1).Tag = "@" & TmTag
                Else
                    RpRow.Record(3).Value = False
                    TmTag = RpRow.Record(1).Tag
                    TmTag = Mid$(TmTag, 2, Len(TmTag) - 1)
                    RpRow.Record(1).Tag = "@" & TmTag
                End If
            End Select
        End If
    End If
ElseIf TreKy = "K24" Then 'OTS-Betreffs
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        RowNr = RpRow.Index
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(5)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                IdxZa = CInt(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                IdxZa = 0
            End If
            If Opt_Prf(IdxZa) = True Then
                FNode "K24"
                RpRws.Row(0).Selected = False
                RpRws.Row(RowNr).EnsureVisible
                RpRws.Row(RowNr).Selected = True
            End If
        End If
    End If
End If

Opt_Sav TrLi1.SelectedItem.Key

Item.Tag = TmpTg

RpCon.Populate

Set RpRws = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

End Sub
Private Sub repCont3_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String

Set TxPLZ = Me.txtFePLZ
Set TxOrt = Me.txtFeOrt

TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

Item.Tag = "@" & TmTag

If TxPLZ.Text <> vbNullString Then
    Opt_PLs TxPLZ.Text, 1
ElseIf TxOrt.Text <> vbNullString Then
    Opt_PLs TxOrt.Text, 2
End If

End Sub

Private Sub repCont4_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim TmTag As String

Set TxBLZ = Me.txtFeBLZ
Set TxBnk = Me.txtFeBnk

TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

Item.Tag = "@" & TmTag

If TxBLZ.Text <> vbNullString Then
    Opt_BLs TxBLZ.Text, 1
ElseIf TxBnk.Text <> vbNullString Then
    Opt_BLs TxBnk.Text, 2
End If

End Sub
Private Sub clDru_DruGef(ByVal DruNam As String, ByVal PorNam As String, ByVal DrvInf As String, ByVal Kommen As String, ByVal DruZa As Integer)
On Error Resume Next

ReDim Preserve DrNam(DruZa)

DrNam(DruZa) = DruNam

DrGef = True

End Sub
Private Sub Form_Load()
On Error Resume Next

Set clDru = New clsDruck

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 12000
    .ClientMaxWidth = 12000
    .ClientMinHeight = 8000
    .ClientMinWidth = 7800
End With

OpLad = True

LiTyp = 1

With clDru
    If GlRDP = True Then
        If GlNeD = True Then 'Netzwerkdrucker auflisten
            .DruLst
        Else
            .DruErm 1
        End If
    Else
        .DruLst
    End If
End With

FInit
AFont Me
FMenu
FRgLa
FRgLo
FFarb
Opt_PLc
Opt_BLc
Opt_Spl "K01"
Opt_Lad "K01"

OpLad = False

Set clDru = Nothing
Set FrmEx = Nothing

If GlRah = True Then
    SFrame 1, Me.hwnd
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub
Private Sub TbBar_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    TaIdx = Item.Index
    FTabu
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
Private Sub trvList1_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)

Dim TreKy As String

Set TrLi1 = Me.trvList1

If OpLad = False Then
    TreKy = Node.Key
    If TreKy <> "K00" Then
        FNode TreKy
    End If
    
    For Each Knote In TrLi1.Nodes
        Knote.Image = IC16_Folder_Close
    Next Knote
    
    Node.Image = IC16_Folder_Open
    TrLi1.Nodes(1).Image = IC16_Folder_View
End If

End Sub
Private Sub txtFeBLZ_GotFocus()
    Me.txtFeBnk.Text = vbNullString
End Sub
Private Sub txtFeBLZ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.txtFeBLZ.Text <> vbNullString Then
            Opt_BLZ Me.txtFeBLZ.Text, 1
        End If
    End If
End Sub
Private Sub txtFeBnk_GotFocus()
    Me.txtFeBLZ.Text = vbNullString
End Sub
Private Sub txtFeBnk_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.txtFeBnk.Text <> vbNullString Then
            Opt_BLZ Me.txtFeBnk.Text, 2
        End If
    End If
End Sub
Private Sub txtFeOrt_GotFocus()
    Me.txtFePLZ.Text = vbNullString
End Sub
Private Sub txtFeOrt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.txtFeOrt.Text <> vbNullString Then
            Opt_PLZ Me.txtFeOrt.Text, 2
        End If
    End If
End Sub
Private Sub txtFePLZ_GotFocus()
    Me.txtFeOrt.Text = vbNullString
End Sub
Private Sub txtFePLZ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.txtFePLZ.Text <> vbNullString Then
            Opt_PLZ Me.txtFePLZ.Text, 1
        End If
    End If
End Sub
