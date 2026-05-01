VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmTerAla 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Erinnerung"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "frmTerAla.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1940
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1160
      Width           =   6570
      _Version        =   1048579
      _ExtentX        =   11589
      _ExtentY        =   3422
      _StockProps     =   64
      BorderStyle     =   3
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdButt3 
      Height          =   400
      Left            =   5340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3400
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "S&chließen"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdButt2 
      Height          =   400
      Left            =   3600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3400
      Width           =   1600
      _Version        =   1048579
      _ExtentX        =   2822
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "Termin Bearbeiten"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdButt1 
      Height          =   400
      Left            =   160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3400
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Alle Schließen"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   80
   End
   Begin VB.Label lblLab02 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   440
      TabIndex        =   6
      Top             =   560
      Width           =   6000
   End
   Begin VB.Label lblLab01 
      BackStyle       =   0  'Transparent
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
      Left            =   440
      TabIndex        =   5
      Top             =   180
      Width           =   6000
   End
End
Attribute VB_Name = "frmTerAla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Lbl01 As VB.Label
Private Lbl02 As VB.Label
Private CaCol As XtremeCalendarControl.CalendarControl
Private CaEvt As XtremeCalendarControl.CalendarEvent
Private CaEvs As XtremeCalendarControl.CalendarEvents
Private CaRem As XtremeCalendarControl.CalendarReminder
Private CaRms As XtremeCalendarControl.CalendarReminders
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ImMan As XtremeCommandBars.ImageManager

Private clFen As clsFenster

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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FLiAd()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim IdxNr As Long
Dim MinAn As Long
Dim TeBet As String
Dim TiStr As String
Dim AnzPo As Integer
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerAla
Set RpCon = FM.repCont1
Set CaCol = frmMain.calCont1
Set CaRms = CaCol.Reminders
Set RpRcs = RpCon.Records

AnzPo = RpCon.Records.Count

For Each CaRem In CaRms
    Set CaEvt = CaRem.Event
    IdxNr = CaEvt.id
    GlTem = IdxNr
    TeBet = CaEvt.Subject
    MinAn = Abs(DateDiff("n", Now, CaEvt.StartTime))
    If MinAn > 0 Then
        TiStr = TeFor(MinAn, True)
    Else
        TiStr = "Seit " & TeFor(-1 * MinAn, True) & " überfällig"
    End If
    If AnzPo > 0 Then
        For Each RpRec In RpRcs
            IdxNr = RpRec.Item(0).Value
            If IdxNr <> GlTem Then
                Set RpRec = RpRcs.Add()
                Set RpItm = RpRec.AddItem(GlTem)
                Set RpItm = RpRec.AddItem(TeBet)
                RpItm.Icon = IC16_Calendar_Year
                Set RpItm = RpRec.AddItem(TiStr)
                Exit For
            End If
        Next RpRec
    Else
        Set RpRec = RpRcs.Add()
        Set RpItm = RpRec.AddItem(GlTem)
        Set RpItm = RpRec.AddItem(TeBet)
        RpItm.Icon = IC16_Calendar_Year
        Set RpItm = RpRec.AddItem(TiStr)
    End If
Next CaRem

RpCon.Populate

Set CaRms = Nothing
Set RpRcs = Nothing
Set RpCon = Nothing
Set CaCol = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLiAd " & Err.Number
Resume Next

End Sub
Private Sub FDisA()
On Error GoTo InErr

Dim IdxNr As Long
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerAla
Set RpCon = FM.repCont1
Set CaCol = frmMain.calCont1
Set CaRms = CaCol.Reminders

For Each CaRem In CaRms
    Set CaEvt = CaRem.Event
    IdxNr = CaEvt.id
    S_TeAl IdxNr
Next CaRem

CaRms.DismissAll

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    .Populate
End With

FLiSe
     
Set CaCol = Nothing
Set RpCon = Nothing
     
Unload Me

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDisA " & Err.Number
Resume Next

End Sub
Private Sub FDism()
On Error GoTo InErr

Dim GlTem As Long
Dim IdxNr As Long
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerAla
Set RpCon = FM.repCont1
Set RpSel = RpCon.SelectedRows
Set RpRcs = RpCon.Records
Set CaCol = frmMain.calCont1
Set CaRms = CaCol.Reminders

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        IdxNr = RpRow.Record.Index
        
        If CaRms.Count > 0 Then
            Set CaRem = CaRms(IdxNr)
            
            Set CaEvt = CaRem.Event
    
            IdxNr = CaEvt.id
            
            GlTem = IdxNr
            
            S_TeAl GlTem
            
            If CaRem.Dismiss() = True Then
                RpRcs.RemoveAt IdxNr
                RpCon.Populate
                FLiSe
            End If
            
            Set RpRcs = RpCon.Records
            If RpRcs.Count = 0 Then Unload Me
        Else
            Unload Me
        End If
    End If
End If

Set CaCol = Nothing
Set RpRcs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDism " & Err.Number
Resume Next

End Sub
Private Sub FErne()
On Error GoTo InErr
'Snooze

Dim AnMin As Long
Dim IdxNr As Long
Dim GlTem As Long
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerAla
Set RpCon = FM.repCont1
Set RpSel = RpCon.SelectedRows
Set RpRcs = RpCon.Records
Set CaCol = frmMain.calCont1
Set CaRms = CaCol.Reminders

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    
    IdxNr = RpRow.Record.Index
    
    Set CaRem = CaRms(IdxNr)
    
    Set CaEvt = CaRem.Event

    IdxNr = CaEvt.id
    
    GlTem = IdxNr

    S_TeAl GlTem, AnMin
    
    If CaRem.Snooze(AnMin) = True Then
        RpRcs.RemoveAt IdxNr
        RpCon.Populate
        FLiSe
    End If
End If

Set RpRcs = RpCon.Records
If RpRcs.Count = 0 Then Unload Me

Set CaCol = Nothing
Set RpRcs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FErne " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim RetWe As Long
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerAla
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set ImMan = frmMain.imgManag

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

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
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderClientEdge
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = True
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Leistungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Leistungen vorhanden"
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
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCls
    Set RpCol = .Add(0, "ID0", 0, False)
    Set RpCol = .Add(1, "Betreff", 200, True)
    Set RpCol = .Add(2, "Fällig in", 100, True)
    RpCol.AutoSize = True
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Groupable = True
    RpCol.Sortable = False
    RpCol.Alignment = xtpAlignmentIconLeft
Next RpCol

Me.BackColor = GlBak

clFen.FenVor

Set clFen = Nothing

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FLiSe()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim IdxNr As Long
Dim RetWe As Long
Dim BuEna As Boolean
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerAla
Set RpCon = FM.repCont1
Set Lbl01 = FM.lblLab01
Set Lbl02 = FM.lblLab02
Set RpSel = RpCon.SelectedRows
Set RpRcs = RpCon.Records
Set CaCol = frmMain.calCont1
Set CaRms = CaCol.Reminders

BuEna = False

If RpSel.Count = 0 Then
    Lbl01.Caption = vbNullString
    If RpRcs.Count > 0 Then
        Lbl02.Caption = "Es sind keine Erinnerungen selektiert"
    Else
        Lbl02.Caption = "Es sind keine Erinnerungen vorhanden"
    End If
Else
    Set RpRow = RpSel(0)
    BuEna = True
End If

cmdButt1.Enabled = BuEna
cmdButt2.Enabled = BuEna
cmdButt3.Enabled = BuEna

If BuEna = True Then
    IdxNr = RpRow.Record.Index

    If CaRms.Count > 0 Then
        Set CaRem = CaRms(IdxNr)
        Lbl01.Caption = CaRem.Event.Subject
        Lbl02.Caption = "Startzeit:  " & Format$(CaRem.Event.StartTime, "dd.mm.yyyy") & " - " & Format$(CaRem.Event.StartTime, "hh:mm") & " Uhr"
    End If
End If

FM.Caption = RpRcs.Count & " Erinnerung" & IIf(RpRcs.Count > 1, "en", "")

Set RpSel = Nothing
Set RpRcs = Nothing
Set CaCol = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLiSe " & Err.Number
Resume Next

End Sub
Private Sub FTeOp()
On Error GoTo InErr

Dim IdxNr As Long
Dim GlTem As Long
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmTerAla
Set RpCon = FM.repCont1
Set RpRcs = RpCon.Records
Set RpSel = RpCon.SelectedRows
Set CaCol = frmMain.calCont1
Set CaRms = CaCol.Reminders

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        IdxNr = RpRow.Record.Index
        
        Set CaRem = CaRms(IdxNr)

        Set CaEvt = CaRem.Event
        
        Set CaCol = Nothing
        
        Unload Me
        
        IdxNr = CaEvt.id
    
        GlTem = IdxNr
        
        TeMain GlTem
    End If
Else
    Exit Sub
End If

Set RpSel = Nothing
Set RpRcs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTeOp " & Err.Number
Resume Next

End Sub

Private Sub cmdButt1_Click()
    FDisA
End Sub
Private Sub cmdButt2_Click()
    FTeOp
End Sub

Private Sub cmdButt3_Click()
    FDism
End Sub
Private Sub Form_Load()
On Error Resume Next

FKonf
AFont Me
FLiAd
FLiSe
SFrame 1, Me.hwnd
WinSound 7

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTerAla = Nothing
End Sub

Private Sub repCont1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    FLiSe
End Sub
