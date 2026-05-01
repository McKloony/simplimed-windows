VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatTX 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   2295
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   2415
      _Version        =   1048579
      _ExtentX        =   4260
      _ExtentY        =   4048
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   360
      Top             =   360
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKatTX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form

Private AktCo As VB.Control
Private TbBar As XtremeCommandBars.TabToolBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmSta As XtremeCommandBars.StatusBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCls As XtremeReportControl.ReportColumns
Private RpRow As XtremeReportControl.ReportRow
Private TxCoN As Tx4oleLib.TXTextControl
Private LiVw4 As XtremeSuiteControls.ListView
Private LiIts As XtremeSuiteControls.ListViewItems
Private LiItm As XtremeSuiteControls.ListViewItem

Private TabId As Integer
Private Sub FEinf()
On Error GoTo SuErr

Dim ObIdx As Long 'Objektnummer
Dim KaStr As String
Dim KoStr As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl

Set FM = frmMain
Set RpCo7 = Me.repCont7
Set TxCoN = FM.TexCont1
Set CmBrs = Me.comBar02
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

Select Case TabId
Case RibTab_Kat_EinTex:

    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then

            Set RpCol = RpCls.Find(0)
            KaStr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(1)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                KoStr = RpRow.Record(RpCol.ItemIndex).Value
            End If
                                
            If GlBut = RibTab_Tex_Rezept Then
                ObIdx = TxCoN.ObjectNext(0, &H20)
                If ObIdx > 0 Then
                    With TxCoN
                        .TextFrameSelect ObIdx
                        .SelText = KaStr & vbCrLf
                        .SelStart = Len(.Text)
                        .SelLength = 0
                        .SetFocus
                        .TextFrameSelect 0
                    End With
    
                    If KoStr <> vbNullString Then
                        ObIdx = TxCoN.ObjectNext(ObIdx, &H20)
                        With TxCoN
                            .TextFrameSelect ObIdx
                            .SelText = KoStr & vbCrLf
                            .SelStart = Len(.Text)
                            .SelLength = 0
                            .SetFocus
                            .TextFrameSelect 0
                        End With
                    End If
                Else
                
                    If KoStr <> vbNullString Then
                        KaStr = KaStr & vbCrLf & KoStr
                    End If
                
                    With TxCoN
                        .SelText = KaStr & vbCrLf
                        .SelStart = Len(.Text)
                        .SelLength = 0
                        .SetFocus
                    End With
                End If
            Else
                If KoStr <> vbNullString Then
                    KaStr = KaStr & vbCrLf & KoStr
                End If
            
                With TxCoN
                    .SelText = KaStr & vbCrLf
                    .SelStart = Len(.Text)
                    .SelLength = 0
                    .SetFocus
                End With
            End If
            
            GlTSV = True 'Speichern Textverarbeitung
        End If
    Next RpRow
    
Case RibTab_Kat_KetTex:

    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then
            
            Set RpCol = RpCls.Find(0)
            KaStr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(1)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                KoStr = RpRow.Record(RpCol.ItemIndex).Value
            End If

            With TxCoN
                .SelText = KoStr & vbCrLf
                .SelStart = Len(.Text)
                .SelLength = 0
                .SetFocus
            End With
            GlTSV = True 'Speichern Textverarbeitung
        End If
    Next RpRow
    
End Select

Set RpCo7 = Nothing
Set CmBrs = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEinf " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case TabId
Case RibTab_Kat_EinTex:
    TeTit = IniGetOpt("Hilfe", 50551)
    TeMai = IniGetOpt("Hilfe", 50552)
    TeInh = IniGetOpt("Hilfe", 50553)
    TeFus = IniGetOpt("Hilfe", 50554)
Case RibTab_Kat_KetTex:
    TeTit = IniGetOpt("Hilfe", 50561)
    TeMai = IniGetOpt("Hilfe", 50562)
    TeInh = IniGetOpt("Hilfe", 50563)
    TeFus = IniGetOpt("Hilfe", 50564)
End Select

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox

Set FM = frmKatTX
Set CmBrs = FM.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

TabId = RbTab.id

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

Select Case TabId
Case RibTab_Kat_EinTex: P_List "TxPh", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetTex: P_List "TxPh", 9, 2
End Select

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

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
Case KY_F2:
Case KY_F3:
Case KY_F4:
Case KY_F5: frmAdrSuch.Show vbModal
Case KY_F6:
Case KY_F7:
Case KY_F8: SSave
Case KY_F9:
Case KY_F10:
Case KY_F11: Unload frmMain
Case KA_Hilfe: FHilfe
Case KM_Zeilenumbruch: KGrKa "GrdZei"
Case KM_Zeilenmarker: KGrKa "GrdMkr"
Case KM_Gitternetz: KGrKa "GrdGrl"
Case KM_Multimarker: KGrKa "MulMar"
Case KA_Eint_Einfuegen: FEinf
Case KA_Kett_Einfuegen: FEinf
Case KA_Eint_Favoriten: FSuFa
Case KA_Eint_Vollst: FSuAu
Case KA_Kett_Vollst: FSuAu
Case KA_SuFe1: FSuch
Case KA_SuFe2: FSuch
Case KA_SuCo1: FSuGr
Case KA_SuCo2: FSuGr
Case 142: FSuLe "Ä", TolId
Case 153: FSuLe "Ö", TolId
Case 154: FSuLe "Ü", TolId
Case Else: If TolId >= 65 And TolId <= 90 Then FSuLe Chr$(TolId), TolId
End Select

GlToo = False

End Sub
Private Sub FSuLe(ByVal SuStr As String, ByVal TolId As Long)
On Error GoTo OrErr
'ABC Leiste

Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

KSuAu "TxPh"

CmAcs(TolId).Checked = True

Select Case TabId
Case RibTab_Kat_EinTex:
        With GlSuE
            .SuIdx = 4
            .SuStr = SuStr
        End With
Case RibTab_Kat_KetTex:
        With GlSuN
            .SuIdx = 4
            .SuStr = SuStr
        End With
End Select

DoEvents
Select Case TabId
Case RibTab_Kat_EinTex: KSuch "TxPh", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetTex: KSuch "TxPh", 9, 2
End Select
DoEvents

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuLe " & Err.Number
Resume Next

End Sub
Private Sub FSuGr()
On Error GoTo OrErr
'Favoriten Knopf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

KSuAu "TxPh"

Select Case TabId
Case RibTab_Kat_EinTex:
        With GlSuE
            .SuIdx = 0
        End With
Case RibTab_Kat_KetTex:
        With GlSuN
            .SuIdx = 0
        End With
End Select

Select Case TabId
Case RibTab_Kat_EinTex: KSuch "TxPh", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetTex: KSuch "TxPh", 9, 2
End Select

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuGr " & Err.Number
Resume Next

End Sub
Private Sub FSuFa()
On Error GoTo OrErr
'Favoriten Knopf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

KSuAu "TxPh"

If GlFTX = False Then
    CmAcs(KA_Eint_Favoriten).Checked = True
    GlFTX = True
    Select Case TabId
    Case RibTab_Kat_EinTex:
            With GlSuE
                .SuIdx = 5
            End With
    Case RibTab_Kat_KetTex:
            With GlSuN
                .SuIdx = 5
            End With
    End Select
Else
    CmAcs(KA_Eint_Favoriten).Checked = False
    GlFTX = False
    Select Case TabId
    Case RibTab_Kat_EinTex:
            With GlSuE
                .SuIdx = 0
            End With
    Case RibTab_Kat_KetTex:
            With GlSuN
                .SuIdx = 0
            End With
    End Select
End If

IniSetVal "Layout", "FavoTX", GlFTX

Select Case TabId
Case RibTab_Kat_EinTex: KSuch "TxPh", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetTex: KSuch "TxPh", 9, 2
End Select

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFa " & Err.Number
Resume Next

End Sub
Private Sub FSuch()
On Error GoTo OrErr
'Sucheingabe

Dim SuStr As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
Set CmEd1 = CmBrs.FindControl(CmEd1, KA_SuFe1, , True)
Set CmEd2 = CmBrs.FindControl(CmEd2, KA_SuFe2, , True)

KSuAu "TxPh"

Select Case TabId
Case RibTab_Kat_EinTex:
            SuStr = CmEd1.Text
            With GlSuE
                .SuIdx = 1
                .SuStr = SuStr
            End With
            KSuch "TxPh", GlMed(CmSu1.ListIndex, 0), 1
            DoEvents
            If RpCo7.Records.Count = 0 Then
                CmEd1.Text = vbNullString
                SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
            Else
                RpCo7.SetFocus
            End If
Case RibTab_Kat_KetTex:
            SuStr = CmEd2.Text
            With GlSuN
                .SuIdx = 1
                .SuStr = SuStr
            End With
            KSuch "TxPh", 9, 2
            DoEvents
            If RpCo7.Records.Count = 0 Then
                CmEd1.Text = vbNullString
                SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
            Else
                RpCo7.SetFocus
            End If
End Select

Set RbBar = Nothing
Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuch " & Err.Number
Resume Next

End Sub
Private Sub FSuAu()
On Error GoTo OrErr
'Hebt die markierten Suchbuchstaben wieder auf

Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RpCo7 = Me.repCont7

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

Select Case TabId
Case RibTab_Kat_EinTex: GlSuE = GlSuX
Case RibTab_Kat_KetTex: GlSuN = GlSuX
End Select

If GlFTX = True Then
    GlFTX = False
    IniSetVal "Layout", "FavoTX", GlFTX
End If

KSuAu "TxPh"
DoEvents

Select Case TabId
Case RibTab_Kat_EinTex: KSuch "TxPh", GlMed(CmSu1.ListIndex, 0), 1
Case RibTab_Kat_KetTex: KSuch "TxPh", 9, 2
End Select

DoEvents
RpCo7.SetFocus

Set RbBar = Nothing
Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuAu " & Err.Number
Resume Next

End Sub
Public Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCls
    Set RpCol = .Add(0, "Bezeichnung", 100, False)
    If RpCo7.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(1, "Kommentar", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Resizable = False
    RpCol.Sortable = True
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(0).AutoSize = True

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpal " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo OpErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long

Set CmBrs = Me.comBar02
Set RpCon = Me.repCont7

CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
ClHoh = ClHoh - ClObn

RpCon.Move ClLin, ClObn, ClBre - ClLin, ClHoh

Set RpCon = Nothing
Set CmBrs = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub Form_Load()
    TabId = 9236 'WICHTIG
    KMnRp "TxPh"
    FSpal
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatTX = Nothing
End Sub
Private Sub comBar02_Resize()
    If GlDcP = False Then
        FPosi
    End If
End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlAkt = False Then
        If Control.id = XTP_ID_RIBBONCONTROLTAB Then
            FTabu
        Else
            FTool Control.id
        End If
    End If
End Sub
Private Sub repCont7_BeginDrag(ByVal Records As XtremeReportControl.IReportRecords)
On Error Resume Next

Dim AktZa
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar

Set FM = frmMain
Set CmBrs = Me.comBar02
Set TxCoN = FM.TexCont1
Set RbBar = CmBrs.Item(1)

Select Case TabId
Case RibTab_Kat_EinTex:
    For AktZa = 0 To Records.Count - 1
        'Records(AktZa).Item(0).Value = vbNullString
    Next AktZa
Case RibTab_Kat_KetTex:
    For AktZa = 0 To Records.Count - 1
        'Records(AktZa).Item(0).Value = vbNullString
        Records(AktZa).Item(0).Value = Records(AktZa).Item(1).Value
    Next AktZa
End Select

TxCoN.SelStart = Len(TxCoN.Text)
TxCoN.SelLength = 0

End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlM11 = False Then
        FEinf
    End If
End Sub
Private Sub repCont7_DragDropCompleted(ByVal Records As XtremeReportControl.IReportRecords, ByVal dropEffect As Long)
On Error Resume Next

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMain
Set TxCoN = FM.TexCont1
Set RpCon = Me.repCont7
Set CmBrs = Me.comBar02

TxCoN.SetFocus
DoEvents
RpCon.Redraw
DoEvents
CmBrs.RecalcLayout

End Sub
