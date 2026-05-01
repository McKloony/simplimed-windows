VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmZuord 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Berichtzuordnung"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   2500
      Left            =   100
      TabIndex        =   1
      Top             =   800
      Width           =   11700
      _Version        =   1048579
      _ExtentX        =   20637
      _ExtentY        =   4410
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont2 
      Height          =   2500
      Left            =   100
      TabIndex        =   2
      Top             =   3600
      Width           =   11700
      _Version        =   1048579
      _ExtentX        =   20637
      _ExtentY        =   4410
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkAdrVe 
      Height          =   240
      Left            =   10000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   1700
      _Version        =   1048579
      _ExtentX        =   2999
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Adressenvergleich"
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   5
      Top             =   6700
      Width           =   12000
      _Version        =   1048579
      _ExtentX        =   21167
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   10000
         TabIndex        =   9
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
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   8600
         TabIndex        =   8
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
      Begin XtremeSuiteControls.PushButton btnFunkt 
         Height          =   400
         Left            =   7200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Löschen"
         UseVisualStyle  =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   5900
         TabIndex        =   6
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
   Begin XtremeSuiteControls.ProgressBar prbStat1 
      Height          =   350
      Left            =   100
      TabIndex        =   3
      Top             =   6300
      Visible         =   0   'False
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   93
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   80
   End
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   495
      Left            =   210
      TabIndex        =   10
      Top             =   160
      Width           =   11200
      _Version        =   1048579
      _ExtentX        =   19756
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   $"frmZuord.frx":0000
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmZuord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PuBu1 As XtremeSuiteControls.PushButton
Private Rahm0 As XtremeSuiteControls.GroupBox
Private RpRow As XtremeReportControl.ReportRow
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRcs As XtremeReportControl.ReportRecords
Private ChAdr As XtremeSuiteControls.CheckBox

Public FoRST As ADODB.Recordset
Private Function FePru() As Boolean
On Error GoTo SuErr
'Prüft ob alle Fragebögen einem Patienten zugeordnet sind

Dim PatNr As Long
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo1 As XtremeReportControl.ReportControl

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpRws = RpCo1.Rows

For Each RpRow In RpRws
    If RpRow.GroupRow = False Then
        If RpRow.Record(19).CheckboxState = 0 Then 'Wenn Fragebogen nicht gelöscht werden soll
            If RpRow.Record(17).Value <> vbNullString Then
                PatNr = RpRow.Record(17).Value '[ID0]
                If PatNr = 0 Then
                    FePru = True
                    SPopu "Unvollständige Patientenzuordnung", "Nicht alle Fragebögen wurden zugeordnet!", IC48_Forbidden
                    Exit For
                End If
            End If
        End If
    End If
Next RpRow

Set RpRws = Nothing
Set RpCo1 = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FePru " & Err.Number
Resume Next

End Function

Private Sub FeLoe()
On Error GoTo InErr

Dim IdxNr As Long
Dim AnzPo As Long
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

AnzPo = RpSel.Count

Tit1 = "Einträge Entfernen"
If AnzPo > 1 Then
    Mld1 = "Möchten Sie die " & AnzPo & " markierten Einträge wirklich entfernen?"
Else
    Mld1 = "Möchten Sie den aktuellen Eintrag wirklich entfernen?"
End If

If AnzPo > 0 Then
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
        For Each RpRow In RpSel
            If RpRow.GroupRow = False Then
                IdxNr = RpRow.Record(0).Value
                DBCmEx1 "qryLabBeLo", "@IdxNr", IdxNr
            End If
        Next RpRow
        T_FeB
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FeLoe " & Err.Number
Resume Next

End Sub

Private Sub btnFunkt_Click()
    Select Case GlBut
    Case RibTab_LabBericht: FeLoe
    Case RibTab_LabBerichte: FeLoe
    Case Else: F_An3
    End Select
End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case GlBut
Case RibTab_LabBericht:
    TeTit = IniGetOpt("Hilfe", 50671)
    TeMai = IniGetOpt("Hilfe", 50672)
    TeInh = IniGetOpt("Hilfe", 50673)
    TeFus = IniGetOpt("Hilfe", 50674)
Case RibTab_LabBerichte:
    TeTit = IniGetOpt("Hilfe", 50671)
    TeMai = IniGetOpt("Hilfe", 50672)
    TeInh = IniGetOpt("Hilfe", 50673)
    TeFus = IniGetOpt("Hilfe", 50674)
Case Else:
    TeTit = IniGetOpt("Hilfe", 50681)
    TeMai = IniGetOpt("Hilfe", 50682)
    TeInh = IniGetOpt("Hilfe", 50683)
    TeFus = IniGetOpt("Hilfe", 50684)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    
Screen.MousePointer = vbHourglass
DoEvents

Select Case GlBut
Case RibTab_LabBericht:
            T_FeA
            DoEvents
            S_LaPV
            DoEvents
            SUpLa
            DoEvents
            Unload Me
Case RibTab_LabBerichte:
            T_FeA
            DoEvents
            S_LaPV
            DoEvents
            SUpLa
            DoEvents
            Unload Me
Case Else:
            If FePru = False Then 'Zuordnungsprüfung
                F_An4 'Adressenvergleich
                DoEvents
                S_AnBoH 'Fügt einen zugeordneten Fragebogen hinzu
                DoEvents
                S_AnBoP 'Befüllen des PropertyGrid
                DoEvents
                SUpAn
                DoEvents
                SAnLo
                DoEvents
                Unload Me
            End If
End Select

DoEvents
Screen.MousePointer = vbNormal

End Sub

Private Sub chkAdrVe_Click()
On Error Resume Next

Dim AdrVe As Boolean

Set ChAdr = Me.chkAdrVe

If ChAdr.Value = xtpChecked Then
    AdrVe = True
End If

Screen.MousePointer = vbHourglass
DoEvents
    
Select Case GlBut
Case RibTab_LabBericht: T_FeP
Case RibTab_LabBerichte: T_FeP
Case Else:
        FaVe1
        If AdrVe = True Then
            FaVe2
        Else
            F_An2
        End If
End Select

DoEvents
Screen.MousePointer = vbNormal

End Sub
Private Sub Form_Load()
On Error Resume Next

Set ChAdr = Me.chkAdrVe

Me.BackColor = GlBak
AFont Me
SFrame 1, Me.hwnd

Select Case GlBut
Case RibTab_LabBericht:
Case RibTab_LabBerichte:
Case Else: ChAdr.Visible = True
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case GlBut
    Case RibTab_LabBericht:
    Case RibTab_LabBerichte:
    Case Else:
        If FoRST.State = adStateOpen Then
            FoRST.Close
            Set FoRST = Nothing
        End If
    End Select
    Set frmZuord = Nothing
End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next
    
Select Case GlBut
Case RibTab_LabBericht:
Case RibTab_LabBerichte:
Case Else:
    If Row.Record(0).Value = vbNullString Then
        Metrics.ForeColor = 8421504
        Metrics.Font.Strikethrough = True
    End If
End Select

End Sub
Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim AdrVe As Boolean

Set ChAdr = Me.chkAdrVe

If ChAdr.Value = xtpChecked Then
    AdrVe = True
End If

Screen.MousePointer = vbHourglass
DoEvents
    
Select Case GlBut
Case RibTab_LabBericht: T_FeP
Case RibTab_LabBerichte: T_FeP
Case Else:
        FaVe1
        If AdrVe = True Then
            FaVe2
        Else
            F_An2
        End If
End Select

DoEvents
Screen.MousePointer = vbNormal

End Sub
Private Sub repCont1_MouseUp(Button As Integer, Shift As Integer, x As Long, Y As Long)
On Error Resume Next

Dim AdrVe As Boolean

Set ChAdr = Me.chkAdrVe

If ChAdr.Value = xtpChecked Then
    AdrVe = True
End If

Screen.MousePointer = vbHourglass
DoEvents
    
Select Case GlBut
Case RibTab_LabBericht: T_FeP
Case RibTab_LabBerichte: T_FeP
Case Else:
        FaVe1
        If AdrVe = True Then
            FaVe2
        Else
            F_An2
        End If
End Select

DoEvents
Screen.MousePointer = vbNormal

End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error GoTo SuErr

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow
Dim RpCo1 As XtremeReportControl.ReportControl

Set FM = frmZuord
Set RpCo1 = FM.repCont1
Set RpRws = RpCo1.Rows
Set RpSel = RpCo1.SelectedRows

TeTit = "Patientenzuordnung"
TeMai = "Soll die besthende Patientenzuordnung entfernt werden?"
TeInh = "Aufgrund der vom Patienten hinterlassenen E-Mail-Adresse wurde diesem Fragebogen ein Patient zugeordnet. Diese Zuordnung kann entfernt werden."
TeFus = "Nachdem die Zuordnung entfernt wurde, ist es möglich einen anderen Patienten zuzuordnen oder einen neuen Patienten hinzuzufügen."

Select Case GlBut
Case RibTab_LabBericht: frmAdrSuch.Show vbModal
Case RibTab_LabBerichte: frmAdrSuch.Show vbModal
Case Else:
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.Record(17).Value <> vbNullString Then
            If RpRow.Record(17).Value > 0 Then
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
                If GlMes = 33565 Then
                    RpRow.Record(16).Value = "?"
                    RpRow.Record(17).Value = 0
                End If
            End If
        End If
        frmAdrSuch.Show vbModal
    End If
End Select

Screen.MousePointer = vbNormal

Set RpSel = Nothing
Set RpRws = Nothing
Set RpCo1 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "repCont1_RowDblClick " & Err.Number
Resume Next

End Sub
Private Sub repCont2_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FaVe3
End Sub
