VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKatBA 
   BorderStyle     =   0  'Kein
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont7 
      Height          =   1815
      Left            =   1560
      TabIndex        =   0
      Top             =   2760
      Width           =   3015
      _Version        =   1048579
      _ExtentX        =   5318
      _ExtentY        =   3201
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
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
Attribute VB_Name = "frmKatBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form

Private AktCo As VB.Control
Private TxDa1 As XtremeSuiteControls.FlatEdit
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

Private TabId As Integer

Private Sub FEdit()
On Error GoTo AnErr

Dim IdxNr As Long
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns
Set RpSel = RpCo7.SelectedRows

If RpSel.Count > 0 Then
    Select Case TabId
    Case RibTab_Kat_EinBuc:
            Set RpRow = RpSel(0)
            Set RpCol = RpCls.Find(OPo_ID1)
            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
            'TeMain IdxNr 'This function must be created in future
    Case RibTab_Kat_KetBuc:
            
    End Select
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo7 = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEdit " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case TabId
Case RibTab_Kat_EinBan:
    TeTit = IniGetOpt("Hilfe", 50331)
    TeMai = IniGetOpt("Hilfe", 50332)
    TeInh = IniGetOpt("Hilfe", 50333)
    TeFus = IniGetOpt("Hilfe", 50334)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
Case RibTab_Kat_KetBan:
    If WindowLoad("frmKetten") = False Then
        TeTit = IniGetOpt("Hilfe", 50341)
        TeMai = IniGetOpt("Hilfe", 50342)
        TeInh = IniGetOpt("Hilfe", 50343)
        TeFus = IniGetOpt("Hilfe", 50344)
        SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
    End If
End Select

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

If ClHoh - 2200 > 0 Then
    RpCon.Move ClLin, ClObn, ClBre - ClLin, ClHoh
End If

Set RpCon = Nothing
Set CmBrs = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub

Public Sub FSpal()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

With RpCo7
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Select Case TabId:
Case RibTab_Kat_EinBan:
    With RpCls
        Set RpCol = .Add(OPo_ID1, "ID1", 0, False)
        Set RpCol = .Add(OPo_RechNr, "Rechnung", 0, True)
        Set RpCol = .Add(OPo_OffBetrag, "Offen", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Stufe, "M", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Patient, "Patient", 0, True)
        If RpCo7.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(OPo_ReBetrag, "Betrag", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Bezahlt, "Bezahlt", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Gebuehr, "Gebühr", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_W, "W", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Datum, "Datum", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Fällig, "Fällig", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Einzahlung, "Zahlung", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Mahnfrist, "Mahnfrist", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Groupable = False
        Set RpCol = .Add(OPo_IDW, "IDW", 0, False)
        Set RpCol = .Add(OPo_Mahnbar, "Mahnbar", 0, False)
        Set RpCol = .Add(OPo_Intervall, "Intervall", 0, False)
        Set RpCol = .Add(OPo_ID0, "ID0", 0, False)
        Set RpCol = .Add(OPo_Währung, "Währung", 0, False)
        Set RpCol = .Add(OPo_IDR, "IDR", 0, False)
        Set RpCol = .Add(OPo_Beleg, "Beleg", 0, False)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Selekt, "Selekt", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(OPo_IDP, "Mandant", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(OPo_Kommentar, "Kommentar", 0, False)
        Set RpCol = .Add(OPo_Berichtdatum, "Berichtdatum", 0, True)
        Set RpCol = .Add(OPo_Steuer, "Steuer", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(OPo_Mahnung1, "Mahnung01", 0, True)
        Set RpCol = .Add(OPo_Mahnung2, "Mahnung02", 0, True)
        Set RpCol = .Add(OPo_Mahnung3, "Mahnung03", 0, True)
        Set RpCol = .Add(OPo_Mahnung4, "Mahnung04", 0, True)
        Set RpCol = .Add(OPo_Mahnung5, "Mahnung05", 0, True)
        Set RpCol = .Add(OPo_Monat, "Monat", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            .EditOptions.Constraints.Add "Januar", 1
            .EditOptions.Constraints.Add "Februar", 2
            .EditOptions.Constraints.Add "März", 3
            .EditOptions.Constraints.Add "April", 4
            .EditOptions.Constraints.Add "Mai", 5
            .EditOptions.Constraints.Add "Juni", 6
            .EditOptions.Constraints.Add "Juli", 7
            .EditOptions.Constraints.Add "August", 8
            .EditOptions.Constraints.Add "September", 9
            .EditOptions.Constraints.Add "Oktober", 10
            .EditOptions.Constraints.Add "November", 11
            .EditOptions.Constraints.Add "Dezember", 12
        End With
        Set RpCol = .Add(OPo_Konto, "Konto", 0, False)
        Set RpCol = .Add(OPo_BLZ, "BLZ", 0, False)
        Set RpCol = .Add(OPo_IDT, "Mitarbeiter", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(OPo_IBAN, "IBAN", 0, False)
        Set RpCol = .Add(OPo_BIC, "BLC", 0, False)
        Set RpCol = .Add(OPo_GuiID, "GuiID", 0, False)
        Set RpCol = .Add(OPo_Versand, "V", 0, False)
        RpCol.HeaderAlignment = xtpAlignmentCenter
    End With
    
    For Each RpCol In RpCls
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = True
            .Sortable = True
            .AutoSize = False
            .AutoSortWhenGrouped = False
        End With
    Next RpCol
    
    If GlTFt.SIZE > 10 Then
        RpCls(OPo_RechNr).Width = 140
        RpCls(OPo_ReBetrag).Width = 80
        RpCls(OPo_Stufe).Width = 30
        RpCls(OPo_Patient).Width = 250
        RpCls(OPo_OffBetrag).Width = 80
        RpCls(OPo_Bezahlt).Width = 80
        RpCls(OPo_Gebuehr).Width = 80
        RpCls(OPo_W).Width = 30
        RpCls(OPo_Datum).Width = 110
        RpCls(OPo_Fällig).Width = 110
        RpCls(OPo_Einzahlung).Width = 110
        RpCls(OPo_Mahnfrist).Width = 110
        RpCls(OPo_IDP).Width = 0
        RpCls(OPo_Berichtdatum).Width = 110
        RpCls(OPo_Steuer).Width = 100
        RpCls(OPo_Mahnung1).Width = 110
        RpCls(OPo_Mahnung2).Width = 110
        RpCls(OPo_Mahnung3).Width = 110
        RpCls(OPo_Mahnung4).Width = 110
        RpCls(OPo_Mahnung5).Width = 110
        RpCls(OPo_IDT).Width = 0
        RpCls(OPo_Versand).Width = 20
    Else
        RpCls(OPo_RechNr).Width = 110
        RpCls(OPo_OffBetrag).Width = 70
        RpCls(OPo_Stufe).Width = 30
        RpCls(OPo_Patient).Width = 220
        RpCls(OPo_ReBetrag).Width = 70
        RpCls(OPo_Bezahlt).Width = 70
        RpCls(OPo_Gebuehr).Width = 70
        RpCls(OPo_W).Width = 30
        RpCls(OPo_Datum).Width = 80
        RpCls(OPo_Fällig).Width = 80
        RpCls(OPo_Einzahlung).Width = 80
        RpCls(OPo_Mahnfrist).Width = 80
        RpCls(OPo_IDP).Width = 0
        RpCls(OPo_Berichtdatum).Width = 80
        RpCls(OPo_Steuer).Width = 70
        RpCls(OPo_Mahnung1).Width = 80
        RpCls(OPo_Mahnung2).Width = 80
        RpCls(OPo_Mahnung3).Width = 80
        RpCls(OPo_Mahnung4).Width = 80
        RpCls(OPo_Mahnung5).Width = 80
        RpCls(OPo_IDT).Width = 0
        RpCls(OPo_Versand).Width = 20
    End If
Case RibTab_Kat_KetBan:
    With RpCls
        Set RpCol = .Add(Rec_ID1, "ID1", 0, False)
        Set RpCol = .Add(Rec_ID0, "ID0", 0, False)
        Set RpCol = .Add(Rec_RechNr, "Rechnung", 0, True)
        Set RpCol = .Add(Rec_Datum, "Datum", 0, True)
        RpCol.Groupable = False
        Set RpCol = .Add(Rec_Selekt, "Abgeschlossen", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Type, "T", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Versand, "V", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Betrag, "Betrag", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Bezahlt, "Bezahlt", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Differe, "Offen", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_IDKurz, "Patient", 0, True)
        Set RpCol = .Add(Rec_Offen, "B", 0, False)
        With RpCol
            .Alignment = xtpAlignmentIconCenter
            .HeaderAlignment = xtpAlignmentCenter
            .Tag = 1
        End With
        Set RpCol = .Add(Rec_Extrageb, "Extrageb.", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Rec_Fallig, "Fälligkeit", 0, True)
        Set RpCol = .Add(Rec_Wahrung, "Währung", 0, False)
        Set RpCol = .Add(Rec_IDR, "Zähler", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_ID3, "ID3", 0, False)
        Set RpCol = .Add(Rec_IDZ, "IDZ", 0, False)
        Set RpCol = .Add(Rec_Versicherer, "Katalog", 0, True)
        Set RpCol = .Add(Rec_Zahlziel, "Zahlungsziel", 0, True)
        Set RpCol = .Add(Rec_Drucken, "Drucken", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_IDW, "IDW", 0, False)
        Set RpCol = .Add(Rec_Symbol, "W", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Faktor, "Faktor", 0, False)
        Set RpCol = .Add(Rec_Ziel, "Ziel", 0, False)
        Set RpCol = .Add(Rec_Kommentar, "Kommentar", 0, False)
        Set RpCol = .Add(Rec_IDP, "Mandant", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Rec_Druckdatum, "Gedruckt", 0, True)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Kopie, "Kopie", 0, False)
        Set RpCol = .Add(Rec_Steuer, "Steuer", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
        End With
        Set RpCol = .Add(Rec_Monat, "Monat", 0, True)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            .EditOptions.Constraints.Add "Januar", 1
            .EditOptions.Constraints.Add "Februar", 2
            .EditOptions.Constraints.Add "März", 3
            .EditOptions.Constraints.Add "April", 4
            .EditOptions.Constraints.Add "Mai", 5
            .EditOptions.Constraints.Add "Juni", 6
            .EditOptions.Constraints.Add "Juli", 7
            .EditOptions.Constraints.Add "August", 8
            .EditOptions.Constraints.Add "September", 9
            .EditOptions.Constraints.Add "Oktober", 10
            .EditOptions.Constraints.Add "November", 11
            .EditOptions.Constraints.Add "Dezember", 12
        End With
        Set RpCol = .Add(Rec_Termin, "Termins.", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Storniert, "Storniert", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_PKU, "PKU", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Gruppe, "G", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Beendet, "E", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Rabatt, "Rabatt", 0, False)
        Set RpCol = .Add(Rec_IDM, "Mitarbeiter", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Rec_GuStr, "Gutschrift", 0, False)
        Set RpCol = .Add(Rec_GutNr, "GutNr", 0, False)
        Set RpCol = .Add(Rec_GuiID, "GuiID", 0, False)
        Set RpCol = .Add(Rec_AufNr, "AufNr", 0, False)
        Set RpCol = .Add(Rec_AuStr, "Auftrag", 0, False)
        Set RpCol = .Add(Rec_Formu, "Formular", 0, False)
        Set RpCol = .Add(Rec_OPLoe, "OPL", 0, False)
        RpCol.Alignment = xtpAlignmentIconLeft
        RpCol.Icon = IC16_Pin_Green
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Lock, "Lock", 0, False)
        RpCol.Alignment = xtpAlignmentIconLeft
        RpCol.Icon = IC16_Lock
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_IDO, "IDO", 0, False)
        Set RpCol = .Add(Rec_RzDat, "RzDat", 0, False)
        Set RpCol = .Add(Rec_RzNum, "RzNum", 0, False)
        Set RpCol = .Add(Rec_RzTex, "RzTex", 0, False)
        Set RpCol = .Add(Rec_Grund, "Grund", 0, False)
        Set RpCol = .Add(Rec_ForID, "FID", 0, False)
        RpCol.Tag = 1
    End With
    
    For Each RpCol In RpCls
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = True
            .Sortable = True
            .AutoSize = False
            .AutoSortWhenGrouped = False
        End With
    Next RpCol

    RpCls(Rec_ID1).Width = 0
    RpCls(Rec_ID0).Width = 0
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_RechNr).Width = 140
        RpCls(Rec_Datum).Width = 110
    Else
        RpCls(Rec_RechNr).Width = 110
        RpCls(Rec_Datum).Width = 80
    End If
    RpCls(Rec_Selekt).Width = 0
    RpCls(Rec_Type).Width = 20
    RpCls(Rec_Versand).Width = 0
    RpCls(Rec_Betrag).Width = 70
    RpCls(Rec_Bezahlt).Width = 70
    RpCls(Rec_Differe).Width = 70
    RpCls(Rec_IDKurz).Width = 220
    RpCls(Rec_Offen).Width = 0
    RpCls(Rec_Extrageb).Width = 70
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_Fallig).Width = 110
    Else
        RpCls(Rec_Fallig).Width = 80
    End If
    RpCls(Rec_Wahrung).Width = 0
    RpCls(Rec_IDR).Width = 60
    RpCls(Rec_ID3).Width = 0
    RpCls(Rec_IDZ).Width = 0
    RpCls(Rec_Versicherer).Width = 140
    RpCls(Rec_Zahlziel).Width = 140
    RpCls(Rec_Drucken).Width = 0
    RpCls(Rec_IDW).Width = 0
    RpCls(Rec_Symbol).Width = 30
    RpCls(Rec_Faktor).Width = 0
    RpCls(Rec_Ziel).Width = 0
    RpCls(Rec_Kommentar).Width = 0
    RpCls(Rec_IDP).Width = 0
    RpCls(Rec_Kopie).Width = 0
    RpCls(Rec_Steuer).Width = 60
    RpCls(Rec_Monat).Width = 0
    RpCls(Rec_Termin).Width = 75
    RpCls(Rec_Storniert).Width = 0
    RpCls(Rec_PKU).Width = 50
    RpCls(Rec_Beendet).Width = 0
    RpCls(Rec_Rabatt).Width = 0
    RpCls(Rec_IDM).Width = 0
    RpCls(Rec_GutNr).Width = 0
    RpCls(Rec_AufNr).Width = 0
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_AuStr).Width = 110
    Else
        RpCls(Rec_AuStr).Width = 80
    End If
    RpCls(Rec_Formu).Width = 120
    RpCls(Rec_OPLoe).Width = 18
    RpCls(Rec_Lock).Width = 18
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo7 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpal " & Err.Number
Resume Next

End Sub
Private Sub FSuAu()
On Error GoTo OrErr
'Hebt die markierten Suchbuchstaben wieder auf

Dim TyStr As String
Dim TyIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set CmBrs = Me.comBar02
Set RpCo7 = Me.repCont7

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

TyIdx = CmSu2.ListIndex

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

GlSuE = GlSuX
GlSuE.SuTyp = TyStr

Screen.MousePointer = vbHourglass

KSuAu "BaPo"
DoEvents

Select Case TabId
Case RibTab_Kat_EinBan: KSuch "BaPo", 3, 1
Case RibTab_Kat_KetBan: KSuch "BaPo", 3, 5, CmSu1.Text
End Select

Screen.MousePointer = vbNormal

DoEvents
RpCo7.SetFocus

Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuAu " & Err.Number
Resume Next

End Sub
Private Sub FTabu()
On Error GoTo AnErr

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCo7 = Me.repCont7
Set RpCls = RpCo7.Columns

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

TabId = RbTab.id

Screen.MousePointer = vbHourglass

FSpal
DoEvents

Select Case TabId
Case RibTab_Kat_EinBan: P_List "BaPo", 0, 1
Case RibTab_Kat_KetBan: P_List "BaPo", 0, 2, False, CmSu1.Text
End Select

Screen.MousePointer = vbNormal

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub

Private Sub repCont7_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Select Case TabId
Case RibTab_Kat_EinBan:
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
Case RibTab_Kat_KetBan:
        If Row.GroupRow = False Then
            Select Case Row.Record(Rec_Type).Value
            Case "M": Metrics.ForeColor = 16744448
            Case "L": Metrics.ForeColor = 33023
            Case "V": Metrics.ForeColor = 8421631
            Case "I": Metrics.ForeColor = 13138080
            Case "U": Metrics.ForeColor = 6604830
            Case Else:
                If CBool(Row.Record(Rec_Selekt).Value) = False Then
                    Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
                End If
            End Select
            If Row.Record(Rec_Storniert).Value = True Then
                Metrics.Font.Strikethrough = True
                Metrics.ForeColor = 8421504
            End If
            If CBool(Row.Record(Rec_Selekt).Value) = False Then
                Metrics.Font.Bold = True
            End If
        End If
End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKatBA = Nothing
End Sub
Private Sub Form_Load()
    KMnRp "BaPo"
    TabId = RibTab_Kat_EinBan
    FSpal
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
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KA_Hilfe: FHilfe
Case KM_Zeilenumbruch: KGrKa "GrdZei"
Case KM_Zeilenmarker: KGrKa "GrdMkr"
Case KM_Gitternetz: KGrKa "GrdGrl"
Case KM_Multimarker: KGrKa "MulMar"
Case KA_Eint_Einfuegen: S_BaSet False, False
Case KA_Eint_Vorschlag: S_BaSet True
Case KA_Eint_Loeschen: S_BaSet False, True
Case KA_Eint_Vollst: FSuAu
Case KA_Kett_Einfuegen: S_BaRch
Case KA_Kett_Vollst: FSuAu
Case KA_SuCo1: FSuGr
Case KA_SuCo2: FSuGr
Case SY_SuCm4: FSuGr
Case SY_SuTex: FSuch
End Select

GlToo = False

End Sub
Private Sub FSuGr()
On Error GoTo OrErr
'Favoriten Knopf

Dim TyStr As String
Dim TyIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo7 As XtremeReportControl.ReportControl

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02
Set CmAcs = CmBrs.Actions

Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, , True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)

TyIdx = CmSu2.ListIndex

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

KSuAu "BaPo"

Screen.MousePointer = vbHourglass

With GlSuE
    .SuIdx = 0
    .SuTyp = TyStr
End With

Select Case TabId
Case RibTab_Kat_EinBan: KSuch "BaPo", 1, 1
Case RibTab_Kat_KetBan: KSuch "BaPo", 1, 5, CmSu1.Text
End Select

Screen.MousePointer = vbNormal

If RpCo7.Records.Count = 0 Then
    SPopu "Eintrag nicht gefunden", "Die von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set RpCo7 = Nothing
Set CmBrs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuGr " & Err.Number
Resume Next

End Sub
Private Sub FSuch()
On Error GoTo OrErr
'Sucheingabe

Dim TyStr As String
Dim SuStr As String
Dim TyIdx As Integer
Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo7 As XtremeReportControl.ReportControl
Dim CmSu1 As XtremeCommandBars.CommandBarComboBox
Dim CmSu2 As XtremeCommandBars.CommandBarComboBox
Dim CmSu4 As XtremeCommandBars.CommandBarComboBox
Dim CmEd1 As XtremeCommandBars.CommandBarEdit

Set RpCo7 = Me.repCont7
Set CmBrs = Me.comBar02

Set CmSu4 = CmBrs.FindControl(CmSu4, SY_SuCm4, , True)
Set CmSu1 = CmBrs.FindControl(CmSu1, KA_SuCo1, True)
Set CmSu2 = CmBrs.FindControl(CmSu2, KA_SuCo2, , True)
Set CmEd1 = CmBrs.FindControl(CmEd1, SY_SuTex, , True)

LiIdx = CmSu4.ListIndex
TyIdx = CmSu2.ListIndex

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

Screen.MousePointer = vbHourglass

KSuAu "BaPo"

SuStr = CmEd1.Text
With GlSuE
    .SuIdx = LiIdx
    .SuTyp = TyStr
    .SuStr = SuStr
    If LiIdx = 3 Then
        If IsDate(SuStr) = True Then
            .SuDat = CDate(SuStr)
        End If
    End If
    If LiIdx = 4 Then
        If IsNumeric(SuStr) = True Then
            .SuBet = CSng(SuStr)
        End If
    End If
End With

Select Case TabId
Case RibTab_Kat_EinBan: KSuch "BaPo", 3, 1
Case RibTab_Kat_KetBan: KSuch "BaPo", 3, 5, CmSu1.Text
End Select
DoEvents

Screen.MousePointer = vbNormal

If RpCo7.Records.Count = 0 Then
    CmEd1.Text = vbNullString
    SPopu "Eintrag nicht gefunden", "Der von Ihnen gesuchte Eintrag, konnte nicht gefunden werden", IC48_Forbidden
Else
    RpCo7.SetFocus
End If

Set CmBrs = Nothing
Set RpCo7 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuch " & Err.Number
Resume Next

End Sub
Private Sub repCont7_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
On Error Resume Next

Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set CmBrs = Me.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Select Case RbTab.id
Case RibTab_Kat_EinBan: S_BaSet False
Case RibTab_Kat_KetBan:
End Select

End Sub

