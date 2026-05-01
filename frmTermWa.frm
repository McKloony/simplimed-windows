VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmTermWa 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Wartezimmer"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   7
      Top             =   4400
      Width           =   5400
      _Version        =   1048579
      _ExtentX        =   9525
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   3400
         TabIndex        =   10
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
         Left            =   2000
         TabIndex        =   9
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
         Left            =   700
         TabIndex        =   8
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
      Height          =   4400
      Left            =   800
      TabIndex        =   1
      Top             =   0
      Width           =   3800
      _Version        =   1048579
      _ExtentX        =   6703
      _ExtentY        =   7761
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbRaum1 
         Height          =   310
         Left            =   300
         TabIndex        =   4
         Top             =   2030
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   310
         Left            =   300
         TabIndex        =   5
         Top             =   2730
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox5"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   310
         Left            =   300
         TabIndex        =   6
         Top             =   3430
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox5"
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   350
         Left            =   1220
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1330
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtVonZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtVonZe 
         Height          =   350
         Left            =   300
         TabIndex        =   2
         Tag             =   "0ZeiVon"
         Top             =   1330
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
      Begin VB.Label lblLab58 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   210
         Left            =   305
         TabIndex        =   15
         Top             =   2500
         Width           =   1200
      End
      Begin VB.Label lblLab56 
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   210
         Left            =   305
         TabIndex        =   14
         Top             =   3200
         Width           =   1200
      End
      Begin VB.Label lblLab49 
         BackStyle       =   0  'Transparent
         Caption         =   "Raum :"
         Height          =   210
         Left            =   305
         TabIndex        =   13
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblLab50 
         BackStyle       =   0  'Transparent
         Caption         =   "Uhrzeit :"
         Height          =   210
         Left            =   305
         TabIndex        =   12
         Top             =   1100
         Width           =   1200
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   730
         Left            =   100
         TabIndex        =   11
         Top             =   200
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   1288
         _StockProps     =   79
         Caption         =   $"frmTermWa.frx":0000
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txoDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6000
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
End
Attribute VB_Name = "frmTermWa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private TxVon As XtremeSuiteControls.FlatEdit
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmRau As XtremeSuiteControls.ComboBox
Private UpCo2 As XtremeSuiteControls.UpDown

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

Public WaSet As Integer

Private clFen As clsFenster

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub btnHilfe_Click()
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
Private Sub btnSchließ_Click()
    Unload Me
End Sub

Private Sub btnWieter_Click()
    FAbs
    Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub FAbs()
On Error GoTo SuErr

Dim TerNr As Long
Dim RowNr As Long
Dim WarNr As Long
Dim SuStr As String
Dim IdWar As Integer
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpSel As XtremeReportControl.ReportSelectedRows
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo6 = FM.repCont6

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If WaSet > 1 Then
    Set RpCls = RpCo6.Columns
    Set RpSel = RpCo6.SelectedRows
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            WarNr = RpRow.Index
            Set RpCol = RpCls.Find(War_ID2)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                TerNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                TerNr = 0
            End If
            Set RpCol = RpCls.Find(War_GuiID)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                SuStr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                SuStr = vbNullString
            End If
        End If
    End If
    S_WaWa TerNr, WaSet
Else
    Set RpCls = RpCo1.Columns
    Set RpSel = RpCo1.SelectedRows
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            Set RpCol = RpCls.Find(Ter_ID2)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                TerNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                TerNr = 0
            End If
            Set RpCol = RpCls.Find(Ter_GuiID)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                SuStr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                SuStr = vbNullString
            End If
            Set RpCol = RpCls.Find(Ter_WartZim)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                IdWar = RpRow.Record(RpCol.ItemIndex).Value
            Else
                IdWar = 0
            End If
        End If
    End If
    If IdWar = 0 Then
        S_WaWa TerNr, WaSet
    Else
        SPopu "Patientenaufnahme", "Dieser Termin wurde der Wartezimmerliste bereits hinzugefügt", IC48_Forbidden
    End If
End If

DoEvents
S_WaPo WarNr, SuStr
DoEvents
SUpTe RowNr

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RpCo1 = Nothing
Set RpCo6 = Nothing
Set RpCls = Nothing
Set RpSel = Nothing

Set clFen = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAbs " & Err.Number
Resume Next

End Sub


Private Sub FInit()
On Error GoTo SuErr

Dim ZeiAn As Date
Dim ManNr As Long
Dim MitNr As Long
Dim RauNr As Integer
Dim AktZa As Integer
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpSel As XtremeReportControl.ReportSelectedRows
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmTermWa
Set TxVon = FM.txtVonZe
Set CmMan = FM.cmbManda
Set CmMit = FM.cmbMitar
Set CmRau = FM.cmbRaum1
Set UpCo2 = FM.updCont2
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set RpCo1 = frmMain.repCont1
Set RpCo6 = frmMain.repCont6

Select Case WaSet
Case 1:
    Set RpCls = RpCo1.Columns
    Set RpSel = RpCo1.SelectedRows
Case 2:
    Set RpCls = RpCo6.Columns
    Set RpSel = RpCo6.SelectedRows
End Select

If GlRaV = True Then 'Räume
    For AktZa = 1 To UBound(GlRmu)
        CmRau.AddItem GlRmu(AktZa, 1)
        CmRau.ItemData(AktZa - 1) = GlRmu(AktZa, 2)
    Next AktZa
    CmRau.ListIndex = 0
End If

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa
CmMan.ListIndex = GlSMa - 1

For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
    CmMit.AddItem GlMiA(AktZa, 1)
    CmMit.ItemData(AktZa - 1) = GlMiA(AktZa, 2)
Next AktZa
CmMit.ListIndex = GlSmI - 1

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Select Case WaSet
        Case 1: Set RpCol = RpCls.Find(Ter_IDP)
        Case 2: Set RpCol = RpCls.Find(War_IDP)
        End Select
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            ManNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            ManNr = 0
        End If
        Select Case WaSet
        Case 1: Set RpCol = RpCls.Find(Ter_IDM)
        Case 2: Set RpCol = RpCls.Find(War_IDM)
        End Select
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            MitNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            MitNr = 0
        End If
        Select Case WaSet
        Case 1: Set RpCol = RpCls.Find(Ter_IDR)
        Case 2: Set RpCol = RpCls.Find(War_IDR)
        End Select
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            RauNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            RauNr = 0
        End If

        For AktZa = 1 To UBound(GlMan)
            If ManNr = GlMan(AktZa, 2) Then
                CmMan.ListIndex = AktZa - 1
                Exit For
            End If
        Next AktZa
        
        For AktZa = 1 To UBound(GlMiA)
            If MitNr = GlMiA(AktZa, 2) Then
                CmMit.ListIndex = AktZa - 1
                Exit For
            End If
        Next AktZa
        
        For AktZa = 1 To UBound(GlRmu)
            If RauNr = GlRmu(AktZa, 2) Then
                CmRau.ListIndex = AktZa - 1
                Exit For
            End If
        Next AktZa
    End If
End If

With TxVon
    .SetMask "00:00:00", "__:__"
    .Text = Format$(TimeValue(Now), "hh:mm")
End With

If WaSet = 2 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(War_ZeiAN)
        If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
            ZeiAn = TimeValue(RpRow.Record(RpCol.ItemIndex).Value)
        Else
            ZeiAn = TimeValue(Now)
        End If
    End If
    With TxVon
        .Enabled = False
        .Text = Format$(ZeiAn, "hh:mm")
    End With
    UpCo2.Enabled = False
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak

Select Case WaSet
Case 0: FM.Caption = "Patient Entlassen"
Case 1: FM.Caption = "Patient Aufnehmen"
Case 2: FM.Caption = "Patient in Behandlung"
End Select

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing
Set RpCo6 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTermWa = Nothing
End Sub

Private Sub updCont2_DownClick()
On Error Resume Next

Dim AkTim As Date

Set TxVon = Me.txtVonZe

If IsDate(TxVon) = True Then
    AkTim = TimeValue(TxVon.Text)
    AkTim = DateAdd("n", -1, AkTim)
    TxVon.Text = Format$(AkTim, "hh:mm")
End If

End Sub
Private Sub updCont2_UpClick()
On Error Resume Next

Dim AkTim As Date

Set TxVon = Me.txtVonZe

If IsDate(TxVon) = True Then
    AkTim = TimeValue(TxVon.Text)
    AkTim = DateAdd("n", 1, AkTim)
    TxVon.Text = Format$(AkTim, "hh:mm")
End If

End Sub


