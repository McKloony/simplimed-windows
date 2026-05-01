VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmStart 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'Kein
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin XtremeReportControl.ReportControl repCon12 
      Height          =   660
      Left            =   2280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   795
      _Version        =   1048579
      _ExtentX        =   1402
      _ExtentY        =   1164
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCon11 
      Height          =   660
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   795
      _Version        =   1048579
      _ExtentX        =   1411
      _ExtentY        =   1164
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCon10 
      Height          =   660
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   795
      _Version        =   1048579
      _ExtentX        =   1411
      _ExtentY        =   1164
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnPict2 
      Height          =   555
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   555
      _Version        =   1048579
      _ExtentX        =   970
      _ExtentY        =   970
      _StockProps     =   79
      FlatStyle       =   -1  'True
      Appearance      =   6
      MultiLine       =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnPict3 
      Height          =   555
      Left            =   600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   555
      _Version        =   1048579
      _ExtentX        =   970
      _ExtentY        =   970
      _StockProps     =   79
      FlatStyle       =   -1  'True
      Appearance      =   6
      MultiLine       =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnPict4 
      Height          =   555
      Left            =   1200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   555
      _Version        =   1048579
      _ExtentX        =   970
      _ExtentY        =   970
      _StockProps     =   79
      FlatStyle       =   -1  'True
      Appearance      =   6
      MultiLine       =   0   'False
   End
   Begin XtremeSuiteControls.Label lblLab03 
      Height          =   300
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Wiedervorlagen"
   End
   Begin XtremeSuiteControls.Label lblLab02 
      Height          =   300
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Termine"
   End
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   300
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Geburtstage"
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Labl1 As XtremeSuiteControls.Label
Private Labl2 As XtremeSuiteControls.Label
Private Labl3 As XtremeSuiteControls.Label
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private PuBu4 As XtremeSuiteControls.PushButton
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private Sub FButt(ByVal RpCon As Integer)
On Error GoTo AnErr
'Filtert die Patientenrechnungen

Dim AnzPo As Long
Dim IdxNr As Long
Dim Mld1, Tit1 As String

If GlInF = False Then
    Exit Sub
End If

Select Case RpCon
Case 10:
    Select Case GlUb1
    Case "N1": SAdre 1
    Case "N2": STerm True
    Case "N3": S_StSt1
    Case "N4": FLad 1, True
    Case "N5": S_StSt1
    End Select
Case 11:
Select Case GlUb2
    Case "N1": SAdre 1
    Case "N2": STerm True
    Case "N3": S_StSt2
    Case "N4": FLad 1, True
    Case "N5": S_StSt2
    End Select
Case 12:
    Select Case GlUb3
    Case "N1": SAdre 1
    Case "N2": STerm True
    Case "N3": S_StSt3
    Case "N4": FLad 1, True
    Case "N5": S_StSt3
    End Select
End Select

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FButt " & Err.Number
Resume Next

End Sub
Private Sub FDopp(ByVal RpCon As Integer)
On Error GoTo AnErr
'Filtert die Patientenrechnungen

Dim AnzPo As Long
Dim IdxNr As Long
Dim Mld1, Tit1 As String
Dim Rpc10 As XtremeReportControl.ReportControl
Dim Rpc11 As XtremeReportControl.ReportControl
Dim Rpc12 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmStart
Set Rpc10 = FM.repCon10
Set Rpc11 = FM.repCon11
Set Rpc12 = FM.repCon12

Select Case RpCon
Case 10:
    Set RpCls = Rpc10.Columns
    Set RpSel = Rpc10.SelectedRows
    AnzPo = Rpc10.Records.Count
    If AnzPo > 0 Then
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Select Case GlUb1
                Case "N1":
                    Set RpCol = RpCls.Find(0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Select Case GlAdO
                    Case 0: SReZe IdxNr
                    Case 1: SKrZe IdxNr
                    Case 2: AMain IdxNr
                    End Select
                Case "N2":
                    Set RpCol = RpCls.Find(1)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    STerm False, False, IdxNr
                Case "N3":

                Case "N4":
                    Set RpCol = RpCls.Find(0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Select Case GlAdO
                    Case 0: SReZe IdxNr
                    Case 1: SKrZe IdxNr
                    Case 2: AMain IdxNr
                    End Select
                Case "N5":
                    Set RpCol = RpCls.Find(1)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    STerm False, False, IdxNr
                End Select
            End If
        End If
    End If
Case 11:
    Set RpCls = Rpc11.Columns
    Set RpSel = Rpc11.SelectedRows
    AnzPo = Rpc11.Records.Count
    If AnzPo > 0 Then
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Select Case GlUb2
                Case "N1":
                    Set RpCol = RpCls.Find(0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Select Case GlAdO
                    Case 0: SReZe IdxNr
                    Case 1: SKrZe IdxNr
                    Case 2: AMain IdxNr
                    End Select
                Case "N2":
                    Set RpCol = RpCls.Find(1)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    STerm False, False, IdxNr
                Case "N3":
                
                Case "N4":
                    Set RpCol = RpCls.Find(0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Select Case GlAdO
                    Case 0: SReZe IdxNr
                    Case 1: SKrZe IdxNr
                    Case 2: AMain IdxNr
                    End Select
                Case "N5":
                    Set RpCol = RpCls.Find(1)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    STerm False, False, IdxNr
                End Select
            End If
        End If
    End If
Case 12:
    Set RpCls = Rpc12.Columns
    Set RpSel = Rpc12.SelectedRows
    AnzPo = Rpc12.Records.Count
    If AnzPo > 0 Then
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Select Case GlUb3
                Case "N1":
                    Set RpCol = RpCls.Find(0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Select Case GlAdO
                    Case 0: SReZe IdxNr
                    Case 1: SKrZe IdxNr
                    Case 2: AMain IdxNr
                    End Select
                Case "N2":
                    Set RpCol = RpCls.Find(1)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    STerm False, False, IdxNr
                Case "N3":
                
                Case "N4":
                    Set RpCol = RpCls.Find(0)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    Select Case GlAdO
                    Case 0: SReZe IdxNr
                    Case 1: SKrZe IdxNr
                    Case 2: AMain IdxNr
                    End Select
                Case "N5":
                    Set RpCol = RpCls.Find(1)
                    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                    STerm False, False, IdxNr
                End Select
            End If
        End If
    End If
End Select

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set Rpc10 = Nothing
Set Rpc11 = Nothing
Set Rpc12 = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDopp " & Err.Number
Resume Next

End Sub
Private Sub FLad(ByVal LiTyp As Integer, Optional ByVal Flag As Boolean)
On Error GoTo OpErr
'Öffnet das Wiedervorlageformular

Dim AnzPo As Long
Dim IdxNr As Long
Dim Mld1, Tit1 As String
Dim CmMan As XtremeSuiteControls.ComboBox
Dim CmMit As XtremeSuiteControls.ComboBox
Dim PuBu2 As XtremeSuiteControls.PushButton
Dim PuBu3 As XtremeSuiteControls.PushButton
Dim Rpc11 As XtremeReportControl.ReportControl
Dim Rpc12 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmStart
Set Rpc11 = FM.repCon11
Set Rpc12 = FM.repCon12
Set ImMan = frmMain.imgManag

Set CmMan = frmWieder.cmbBehan
Set CmMit = frmWieder.cmbMitar

Select Case LiTyp
Case 1:
    Set RpCls = Rpc12.Columns
    Set RpSel = Rpc12.SelectedRows
    GlKoL = True
    GlKoS = False
    If Flag = True Then
        Load frmWieder
        GlKoN = True
        PuBu2.Icon = ImMan.Icons.GetImage(IC16_Doc_Edit, 16)
        PuBu3.Icon = ImMan.Icons.GetImage(IC16_Folder_Open, 16)
        CmMan.ListIndex = GlSMa - 1
        CmMit.ListIndex = GlSmI - 1
        CmMan.Tag = "1IDP"
        CmMit.Tag = "1IDM"
        frmWieder.Show vbModal
    Else
        GlKoN = False
        AnzPo = Rpc12.Records.Count
        If AnzPo > 0 Then
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Ter_ID2)
                IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                Wie_Lad IdxNr
                PuBu2.Icon = ImMan.Icons.GetImage(IC16_Doc_Edit, 16)
                PuBu3.Icon = ImMan.Icons.GetImage(IC16_Folder_Open, 16)
                frmWieder.Show vbModal
                GlKoS = False
            End If
        Else
            Mld1 = "Es ist kein Wiedervorlageeintrag vorhanden den Sie öffnen könnten"
            Tit1 = "Kein Wiedervorlageeintrag"
            SPopu Tit1, Mld1, IC48_Forbidden
        End If
    End If
    
    GlKoL = False
Case 2:
    Set RpCls = Rpc11.Columns
    Set RpSel = Rpc11.SelectedRows
    AnzPo = Rpc11.Records.Count
    If AnzPo > 0 Then
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            Set RpCol = RpCls.Find(Ter_ID2)
            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
            TeMain IdxNr
            Unload Me
        End If
    Else
        Mld1 = "Es ist kein Termineintrag vorhanden den Sie öffnen könnten"
        Tit1 = "Kein Termineintrag"
        SPopu Tit1, Mld1, IC48_Forbidden
    End If
End Select

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set Rpc11 = Nothing
Set Rpc12 = Nothing
Set ImMan = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLad " & Err.Number
Resume Next

End Sub
Private Sub btnPict2_Click()
    FButt 10
End Sub

Private Sub Form_Resize()
    If GlSta = False Then
        SStPo
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmStart = Nothing
End Sub
Private Sub repCon10_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

Select Case GlUb1
Case "N1":
    
Case "N2":
    If GlSta = False Then
        If Row.GroupRow = False Then
            If Row.Record(4).Value <> vbNullString Then
                FrbZa = Row.Record(4).Value
                If FrbZa > 1 And FrbZa <= 20 Then
                    Metrics.BackColor = GlTmF(FrbZa, 1)
                End If
            End If
        End If
    End If
Case "N3":
    If Row.Record(11).Value = 0 Then
        Metrics.Font.Strikethrough = False
        If Row.Record(9).Value = 0 Then
            Metrics.ForeColor = 12632256
        Else
            If Row.Record(8).Value < Date Then
                Metrics.ForeColor = 210
            Else
                Metrics.ForeColor = 44800
            End If
        End If
        If Row.Record(10).Value <> vbNullString Then
            If Row.Record(10).Value <> "0" Then
                Metrics.Font.Bold = True
            End If
        End If
    Else
        Metrics.ForeColor = 8421504
        Metrics.Font.Strikethrough = True
    End If
Case "N4":
    If CDate(Row.Record(2).Value) < Date Then
        Metrics.ForeColor = vbRed
    ElseIf CDate(Row.Record(2).Value) > Date Then
        Metrics.ForeColor = 8421504
    End If
Case "N5":
    If Row.Record(10).Value <> vbNullString Then
        FrbZa = Row.Record(10).Value
        If FrbZa > 1 And FrbZa <= 20 Then
            Metrics.BackColor = GlTmF(FrbZa, 1)
        End If
    End If
    If Row.Record(8).Value <= Date Then
        If Row.Record(4).Value > Row.Record(5).Value Then
            Metrics.Font.Bold = True
            Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
        End If
    End If
End Select

End Sub

Private Sub repCon10_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

If GlUb1 = "N4" Then
    Dim TmTag As String
    TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
    Item.Tag = "@" & TmTag
    S_WiSa
End If

End Sub
Private Sub repCon10_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then FDopp 10
End Sub

Private Sub repCon11_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

Select Case GlUb2
Case "N1":
    
Case "N2":
    If GlSta = False Then
        If Row.GroupRow = False Then
            If Row.Record(4).Value <> vbNullString Then
                FrbZa = Row.Record(4).Value
                If FrbZa > 1 And FrbZa <= 20 Then
                    Metrics.BackColor = GlTmF(FrbZa, 1)
                End If
            End If
        End If
    End If
Case "N3":
    If Row.Record(11).Value = 0 Then
        Metrics.Font.Strikethrough = False
        If Row.Record(9).Value = 0 Then
            Metrics.ForeColor = 12632256
        Else
            If Row.Record(8).Value < Date Then
                Metrics.ForeColor = 210
            Else
                Metrics.ForeColor = 44800
            End If
        End If
        If Row.Record(10).Value <> vbNullString Then
            If Row.Record(10).Value <> "0" Then
                Metrics.Font.Bold = True
            End If
        End If
    Else
        Metrics.ForeColor = 8421504
        Metrics.Font.Strikethrough = True
    End If
Case "N4":
    If CDate(Row.Record(2).Value) < Date Then
        Metrics.ForeColor = vbRed
    ElseIf CDate(Row.Record(2).Value) > Date Then
        Metrics.ForeColor = 8421504
    End If
Case "N5":
    If Row.Record(10).Value <> vbNullString Then
        FrbZa = Row.Record(10).Value
        If FrbZa > 1 And FrbZa <= 20 Then
            Metrics.BackColor = GlTmF(FrbZa, 1)
        End If
    End If
    If Row.Record(8).Value <= Date Then
        If Row.Record(4).Value > Row.Record(5).Value Then
            Metrics.Font.Bold = True
            Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
        End If
    End If
End Select

End Sub

Private Sub repCon11_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

If GlUb2 = "N4" Then
    Dim TmTag As String
    TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
    Item.Tag = "@" & TmTag
    S_WiSa
End If

End Sub
Private Sub repCon11_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then FDopp 11
End Sub
Private Sub repCon12_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

Select Case GlUb3
Case "N1":
    
Case "N2":
    If GlSta = False Then
        If Row.GroupRow = False Then
            If Row.Record(4).Value <> vbNullString Then
                FrbZa = Row.Record(4).Value
                If FrbZa > 1 And FrbZa <= 20 Then
                    Metrics.BackColor = GlTmF(FrbZa, 1)
                End If
            End If
        End If
    End If
Case "N3":
    If Row.Record(11).Value = 0 Then
        Metrics.Font.Strikethrough = False
        If Row.Record(9).Value = 0 Then
            Metrics.ForeColor = 12632256
        Else
            If Row.Record(8).Value < Date Then
                Metrics.ForeColor = 210
            Else
                Metrics.ForeColor = 44800
            End If
        End If
        If Row.Record(10).Value <> vbNullString Then
            If Row.Record(10).Value <> "0" Then
                Metrics.Font.Bold = True
            End If
        End If
    Else
        Metrics.ForeColor = 8421504
        Metrics.Font.Strikethrough = True
    End If
Case "N4":
    If CDate(Row.Record(2).Value) < Date Then
        Metrics.ForeColor = vbRed
    ElseIf CDate(Row.Record(2).Value) > Date Then
        Metrics.ForeColor = 8421504
    End If
Case "N5":
    If Row.Record(10).Value <> vbNullString Then
        FrbZa = Row.Record(10).Value
        If FrbZa > 1 And FrbZa <= 20 Then
            Metrics.BackColor = GlTmF(FrbZa, 1)
        End If
    End If
    If Row.Record(8).Value <= Date Then
        If Row.Record(4).Value > Row.Record(5).Value Then
            Metrics.Font.Bold = True
            Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
        End If
    End If
End Select

End Sub
Private Sub repCon12_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

If GlUb3 = "N4" Then
    Dim TmTag As String
    TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
    Item.Tag = "@" & TmTag
    S_WiSa
End If

End Sub
Private Sub repCon12_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlAkt = False Then FDopp 12
End Sub
Private Sub btnPict3_Click()
    FButt 11
End Sub
Private Sub btnPict4_Click()
    FButt 12
End Sub
