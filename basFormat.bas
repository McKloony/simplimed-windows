Attribute VB_Name = "basFormat"
Option Explicit

Private FM As Form
Private PiBi1 As VB.PictureBox
Private PiBi2 As VB.PictureBox
Private PiBi3 As VB.PictureBox
Private PiBi4 As VB.PictureBox
Private PiBi5 As VB.PictureBox
Private PiBi6 As VB.PictureBox
Private PiR01 As VB.PictureBox
Private PiR02 As VB.PictureBox
Private PiR03 As VB.PictureBox
Private PiR04 As VB.PictureBox
Private PiR05 As VB.PictureBox
Private PiR06 As VB.PictureBox
Private PiR07 As VB.PictureBox
Private PiR08 As VB.PictureBox
Private PiR09 As VB.PictureBox
Private PiR10 As VB.PictureBox
Private PiR12 As VB.PictureBox
Private PiR13 As VB.PictureBox
Private Labl1 As XtremeSuiteControls.Label
Private Labl2 As XtremeSuiteControls.Label
Private Labl3 As XtremeSuiteControls.Label
Private Labl4 As XtremeSuiteControls.Label
Private Labl5 As XtremeSuiteControls.Label
Private Labl6 As XtremeSuiteControls.Label
Private Labl7 As XtremeSuiteControls.Label
Private Labl8 As XtremeSuiteControls.Label
Private Lbl01 As XtremeSuiteControls.Label
Private Lbl02 As XtremeSuiteControls.Label
Private Lbl03 As XtremeSuiteControls.Label
Private Lbl04 As XtremeSuiteControls.Label
Private Lbl05 As XtremeSuiteControls.Label
Private Lbl06 As XtremeSuiteControls.Label
Private Lbl07 As XtremeSuiteControls.Label
Private Lbl08 As XtremeSuiteControls.Label
Private Lbl09 As XtremeSuiteControls.Label
Private Lbl10 As XtremeSuiteControls.Label
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private TxDmy As XtremeSuiteControls.FlatEdit
Private TxDe0 As XtremeSuiteControls.FlatEdit
Private TxDe1 As XtremeSuiteControls.FlatEdit
Private TxDe2 As XtremeSuiteControls.FlatEdit
Private TxDe3 As XtremeSuiteControls.FlatEdit
Private TxDe4 As XtremeSuiteControls.FlatEdit
Private TxDe5 As XtremeSuiteControls.FlatEdit
Private TxDe6 As XtremeSuiteControls.FlatEdit
Private TxDe7 As XtremeSuiteControls.FlatEdit
Private TxDe8 As XtremeSuiteControls.FlatEdit
Private TxDe9 As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private TxAnz As XtremeSuiteControls.FlatEdit
Private TxMul As XtremeSuiteControls.FlatEdit
Private TxEin As XtremeSuiteControls.FlatEdit
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmZif As XtremeSuiteControls.ComboBox
Private CmBez As XtremeSuiteControls.ComboBox
Private CmbMo As XtremeSuiteControls.ComboBox
Private CmbQu As XtremeSuiteControls.ComboBox
Private CmbJa As XtremeSuiteControls.ComboBox
Private CmbJv As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private ChJah As XtremeSuiteControls.CheckBox
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private TsDia As XtremeSuiteControls.TaskDialog
Private CoDia As XtremeSuiteControls.CommonDialog
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private OpZei As XtremeSuiteControls.RadioButton
Private TabCo As XtremeSuiteControls.TabControl
Private LiVw1 As XtremeSuiteControls.ListView
Private LiVw2 As XtremeSuiteControls.ListView
Private LiVw3 As XtremeSuiteControls.ListView
Private LiVw4 As XtremeSuiteControls.ListView
Private LiVw5 As XtremeSuiteControls.ListView
Private LiItm As XtremeSuiteControls.ListViewItem
Private PrGr1 As XtremePropertyGrid.PropertyGrid
Private LiFld As FolderViewControl.FolderView
Private LiNod As FolderViewControl.TreeNode
Private LiFi1 As FileViewControl.FileView
Private LiFi2 As FileViewControl.FileView
Private LiFit As FileViewControl.ListItem
Private CmSta As XtremeCommandBars.StatusBar
Private TbBar As XtremeCommandBars.TabToolBar
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmSys As XtremeCommandBars.RibbonBarSystemButton
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmCSe As XtremeCommandBars.CommandBarControlColorSelector
Private CmBuT As XtremeCommandBars.CommandBarButton
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private PoItm As XtremeSuiteControls.PopupControlItem
Private DaPi3 As XtremeCalendarControl.DatePicker
Private DaPi4 As XtremeCalendarControl.DatePicker
Private DaPi6 As XtremeCalendarControl.DatePicker
Private DaPi7 As XtremeCalendarControl.DatePicker
Private CaCol As XtremeCalendarControl.CalendarControl
Private CaDay As XtremeCalendarControl.CalendarDayView
Private CaWek As XtremeCalendarControl.CalendarWeekView
Private CaMon As XtremeCalendarControl.CalendarMonthView
Private CaLab As XtremeCalendarControl.CalendarEventLabel
Private CaLbs As XtremeCalendarControl.CalendarEventLabels
Private CaLbl As XtremeCalendarControl.CalendarEventLabel
Private CaObj As XtremeCalendarControl.CalendarDialogs
Private CaThe As XtremeCalendarControl.CalendarThemeOffice2007
Private CaThI As XtremeCalendarControl.CalendarThemeImageList
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
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private PuBu4 As XtremeSuiteControls.PushButton
Private PuBu5 As XtremeSuiteControls.PushButton
Private PuBu6 As XtremeSuiteControls.PushButton
Private TrLi1 As XtremeSuiteControls.TreeView
Private TrLi2 As XtremeSuiteControls.TreeView
Private TrLi5 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private ChCon As XtremeChartControl.ChartControl
Private ShCut As XtremeShortcutBar.ShortcutBar
Private TxCoN As Tx4oleLib.TXTextControl

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E

Private clFil As clsFile
Private clWor As clsWord
Private clFen As clsFenster
Private clNet As clsNetz
Private clLis As clsLisLab
Private clReg As clsRegis

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Function SArSor(ByRef KeyAr As Variant) As Variant()
On Error GoTo LaErr
'Array sortieren

Dim z As Long
Dim i As Long
Dim strWert As Variant
 
For z = UBound(KeyAr) - 1 To LBound(KeyAr) Step -1
    For i = LBound(KeyAr) To z
        If LCase(KeyAr(i)) > LCase(KeyAr(i + 1)) Then
            strWert = KeyAr(i)
            KeyAr(i) = KeyAr(i + 1)
            KeyAr(i + 1) = strWert
        End If
    Next i
Next z
 
SArSor = KeyAr
     
Exit Function

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SArSor " & Err.Number
Resume Next
     
End Function
Public Sub SMark()
On Error GoTo LaErr
'Listet Anzahl markierter Einträge auf

Dim DayFi As Date
Dim DayLa As Date
Dim PatNr As Long
Dim MaNum As Long
Dim AnzTa As Long
Dim AktTa As Long
Dim AnzTe As Long
Dim IdxNr As Long
Dim GeSum As Double
Dim EiSum As Double
Dim AuSum As Double
Dim PaNam As String
Dim RecNr As String
Dim BuDat As String
Dim BuKto As String
Dim TeDat As String
Dim TeBet As String
Dim PaGeb As String
Dim PaVer As String
Dim BeNam As String
Dim PaAns As String
Dim PaTe1 As String
Dim PaTe2 As String
Dim PaTe3 As String
Dim PaFir As String
Dim GesZa As Integer
Dim AktPo As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmPa1 As XtremeCommandBars.StatusBarPane
Dim CmPa2 As XtremeCommandBars.StatusBarPane
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpRws As XtremeReportControl.ReportRows
Dim RpCls As XtremeReportControl.ReportColumns
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim CaCol As XtremeCalendarControl.CalendarControl

Set FM = frmMain
Set RpCo0 = FM.repCont0
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set RpCo6 = FM.repCont6
Set RpCoK = FM.repContK
Set RpCo8 = FM.repCont8
Set CaCol = FM.calCont1
Set Labl4 = FM.lblDeta4
Set Lbl10 = FM.lblDeta6
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set LiVw4 = FM.lstView4
Set LiFld = FM.fldView1
Set LiFi1 = FM.filView1
Set LiFi2 = FM.filView2

Set CmPa1 = CmSta.FindPane(Tex_Pa_Labl1)
Set CmPa2 = CmSta.FindPane(Tex_Pa_Labl2)

GeSum = 0

Select Case GlBut
Case RibTab_Adressen:
        Set RpCls = RpCo2.Columns
        Set RpSel = RpCo2.SelectedRows
        Set RpRws = RpCo2.Rows
Case RibTab_Mandanten:
        Set RpCls = RpCo2.Columns
        Set RpSel = RpCo2.SelectedRows
        Set RpRws = RpCo2.Rows
Case RibTab_Verordner:
        Set RpCls = RpCo2.Columns
        Set RpSel = RpCo2.SelectedRows
        Set RpRws = RpCo2.Rows
Case RibTab_Mitarbeit:
        Set RpCls = RpCo2.Columns
        Set RpSel = RpCo2.SelectedRows
        Set RpRws = RpCo2.Rows
Case RibTab_Krankenbla:
        Set RpCls = RpCoK.Columns
        Set RpSel = RpCoK.SelectedRows
Case RibTab_Abrechnung:
        Set RpCls = RpCo3.Columns
        Set RpSel = RpCo3.SelectedRows
Case RibTab_Vorbereit:
        Set RpCls = RpCo6.Columns
        Set RpSel = RpCo6.SelectedRows
Case RibTab_Tagesproto:
        Set RpCls = RpCo6.Columns
        Set RpSel = RpCo6.SelectedRows
Case RibTab_Bildmodul:
        Set RpCls = RpCo2.Columns
        Set RpSel = RpCo2.SelectedRows
Case RibTab_Rechnungen:
        Set RpCls = RpCo4.Columns
        Set RpSel = RpCo4.SelectedRows
Case RibTab_Fragebogen:
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_Rezeptmodul:
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_Belegmodul:
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_LabBericht:
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_LabAuftrag:
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_LabBerichte:
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
Case RibTab_LabAuftrage:
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
Case RibTab_Kat_Eintrg:
        Set RpCls = RpCo8.Columns
        Set RpSel = RpCo8.SelectedRows
Case RibTab_Kat_Ketten:
        Set RpCls = RpCo8.Columns
        Set RpSel = RpCo8.SelectedRows
Case RibTab_Kat_Frage:
        Set RpCls = RpCo8.Columns
        Set RpSel = RpCo8.SelectedRows
Case RibTab_Tex_Email:
        Set RpCls = RpCo0.Columns
        Set RpSel = RpCo0.SelectedRows
Case Else:
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
End Select

Select Case GlBut
Case RibTab_Adressen:
                If GlAdA > 0 Then 'Anzahl gefundener Adressen
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        If GlAdA > RpRow.Index Then
                            If AdAry(Adr_Mandant, RpRow.Index) <> vbNullString Then PatNr = AdAry(Adr_Mandant, RpRow.Index)
                            If AdAry(Adr_IDKurz, RpRow.Index) <> vbNullString Then PaNam = AdAry(Adr_IDKurz, RpRow.Index)
                            If AdAry(Adr_Geboren, RpRow.Index) <> vbNullString Then PaGeb = Format$(AdAry(Adr_Geboren, RpRow.Index), "ddd. dd.mm.yyyy") & " (" & AJahr(AdAry(Adr_Geboren, RpRow.Index)) & " Jahre)"
                            If AdAry(Adr_IDP, RpRow.Index) > 0 Then
                                MaNum = AdAry(Adr_IDP, RpRow.Index)
                                For AktPo = 1 To UBound(GlThe)
                                    If MaNum = GlThe(AktPo, 0) Then
                                        BeNam = GlThe(AktPo, 13)
                                        Exit For
                                    End If
                                Next AktPo
                            End If
                        End If
                    End If
                End If
                CmPa2.Text = "Patienten markiert : " & RpSel.Count
Case RibTab_Mandanten:
                If GlAdA > 0 Then 'Anzahl gefundener Adressen
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        If GlAdA >= RpRow.Index Then
                            PatNr = AdAry(Adr_ID0, RpRow.Index)
                            PaNam = AdAry(Adr_IDKurz, RpRow.Index)
                            If AdAry(Adr_Anschrift, RpRow.Index) <> vbNullString Then PaAns = AdAry(Adr_Anschrift, RpRow.Index)
                            If AdAry(Adr_Telefon2, RpRow.Index) <> vbNullString Then PaTe1 = AdAry(Adr_Telefon2, RpRow.Index)
                            If AdAry(Adr_Telefon3, RpRow.Index) <> vbNullString Then PaTe2 = AdAry(Adr_Telefon3, RpRow.Index)
                            If AdAry(Adr_Telefon5, RpRow.Index) <> vbNullString Then PaTe3 = AdAry(Adr_Telefon5, RpRow.Index)
                            If AdAry(Adr_Firma1, RpRow.Index) <> vbNullString Then PaFir = AdAry(Adr_Firma1, RpRow.Index)
                            If AdAry(Adr_Geboren, RpRow.Index) <> vbNullString Then PaGeb = Format$(AdAry(Adr_Geboren, RpRow.Index), "ddd. dd.mm.yyyy") & " (" & AJahr(AdAry(Adr_Geboren, RpRow.Index)) & " Jahre)"
                            If AdAry(Adr_IDP, RpRow.Index) > 0 Then
                                MaNum = AdAry(Adr_IDP, RpRow.Index)
                                For AktPo = 1 To UBound(GlThe)
                                    If MaNum = GlThe(AktPo, 0) Then
                                        BeNam = GlThe(AktPo, 13)
                                        Exit For
                                    End If
                                Next AktPo
                            End If
                        End If
                    End If
                End If
                CmPa2.Text = "Patienten markiert : " & RpSel.Count
                Labl4.Caption = "<TextBlock><Run FontSize='15'>" & PaNam & "</Run> <LineBreak/><Run FontSize='11'>" & PaGeb & "</Run> <LineBreak/><Bold> Telefon : " & PaTe1 & "</Bold> <LineBreak/> Mobil : " & PaTe2 & "<LineBreak/><Italic> EMail : " & PaTe3 & "</Italic><LineBreak/><Bold> Nummer: " & Format$(PatNr, "00000") & "</Bold> <LineBreak/><Run Foreground='DarkGray' FontSize='11'>" & PaFir & "</Run><LineBreak/></TextBlock>"
Case RibTab_Verordner:
                If GlAdA > 0 Then 'Anzahl gefundener Adressen
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        If GlAdA > RpRow.Index Then
                            PatNr = AdAry(Adr_ID0, RpRow.Index)
                            If AdAry(Adr_IDKurz, RpRow.Index) <> vbNullString Then PaNam = AdAry(Adr_IDKurz, RpRow.Index)
                            If AdAry(Adr_Anschrift, RpRow.Index) <> vbNullString Then PaAns = AdAry(Adr_Anschrift, RpRow.Index)
                            If AdAry(Adr_Telefon2, RpRow.Index) <> vbNullString Then PaTe1 = AdAry(Adr_Telefon2, RpRow.Index)
                            If AdAry(Adr_Telefon3, RpRow.Index) <> vbNullString Then PaTe2 = AdAry(Adr_Telefon3, RpRow.Index)
                            If AdAry(Adr_Telefon5, RpRow.Index) <> vbNullString Then PaTe3 = AdAry(Adr_Telefon5, RpRow.Index)
                            If AdAry(Adr_Geboren, RpRow.Index) <> vbNullString Then PaGeb = Format$(AdAry(Adr_Geboren, RpRow.Index), "ddd. dd.mm.yyyy") & " (" & AJahr(AdAry(Adr_Geboren, RpRow.Index)) & " Jahre)"
                            If AdAry(Adr_Firma1, RpRow.Index) <> vbNullString Then PaFir = AdAry(Adr_Firma1, RpRow.Index)
                            If AdAry(Adr_IDP, RpRow.Index) > 0 Then
                                MaNum = AdAry(Adr_IDP, RpRow.Index)
                                For AktPo = 1 To UBound(GlThe)
                                    If MaNum = GlThe(AktPo, 0) Then
                                        BeNam = GlThe(AktPo, 13)
                                        Exit For
                                    End If
                                Next AktPo
                            End If
                        End If
                    End If
                End If
                CmPa2.Text = "Patienten markiert : " & RpSel.Count
                Labl4.Caption = "<TextBlock><Run FontSize='15'>" & PaNam & "</Run> <LineBreak/><Run FontSize='11'>" & PaGeb & "</Run> <LineBreak/><Bold> Telefon : " & PaTe1 & "</Bold> <LineBreak/> Mobil : " & PaTe2 & "<LineBreak/><Italic> EMail : " & PaTe3 & "</Italic><LineBreak/><Bold> Nummer: " & Format$(PatNr, "00000") & "</Bold> <LineBreak/><Run Foreground='DarkGray' FontSize='11'>" & PaFir & "</Run><LineBreak/></TextBlock>"
Case RibTab_Mitarbeit:
                If GlAdA > 0 Then 'Anzahl gefundener Adressen
                    Set RpRow = RpSel(0)
                    If RpRow.GroupRow = False Then
                        If GlAdA > RpRow.Index Then
                            PatNr = AdAry(Adr_ID0, RpRow.Index)
                            PaNam = AdAry(Adr_IDKurz, RpRow.Index)
                            If AdAry(Adr_Anschrift, RpRow.Index) <> vbNullString Then PaAns = AdAry(Adr_Anschrift, RpRow.Index)
                            If AdAry(Adr_Telefon2, RpRow.Index) <> vbNullString Then PaTe1 = AdAry(Adr_Telefon2, RpRow.Index)
                            If AdAry(Adr_Telefon3, RpRow.Index) <> vbNullString Then PaTe2 = AdAry(Adr_Telefon3, RpRow.Index)
                            If AdAry(Adr_Telefon5, RpRow.Index) <> vbNullString Then PaTe3 = AdAry(Adr_Telefon5, RpRow.Index)
                            If AdAry(Adr_Geboren, RpRow.Index) <> vbNullString Then PaGeb = Format$(AdAry(Adr_Geboren, RpRow.Index), "ddd. dd.mm.yyyy") & " (" & AJahr(AdAry(Adr_Geboren, RpRow.Index)) & " Jahre)"
                            If AdAry(Adr_Firma1, RpRow.Index) <> vbNullString Then PaFir = AdAry(Adr_Firma1, RpRow.Index)
                            If AdAry(Adr_IDP, RpRow.Index) > 0 Then
                                MaNum = AdAry(Adr_IDP, RpRow.Index)
                                For AktPo = 1 To UBound(GlThe)
                                    If MaNum = GlThe(AktPo, 0) Then
                                        BeNam = GlThe(AktPo, 13)
                                        Exit For
                                    End If
                                Next AktPo
                            End If
                        End If
                    End If
                End If
                CmPa2.Text = "Patienten markiert : " & RpSel.Count
                Labl4.Caption = "<TextBlock><Run FontSize='15'>" & PaNam & "</Run> <LineBreak/><Run FontSize='11'>" & PaGeb & "</Run> <LineBreak/><Bold> Telefon : " & PaTe1 & "</Bold> <LineBreak/> Mobil : " & PaTe2 & "<LineBreak/><Italic> EMail : " & PaTe3 & "</Italic><LineBreak/><Bold> Nummer: " & Format$(PatNr, "00000") & "</Bold> <LineBreak/><Run Foreground='DarkGray' FontSize='11'>" & PaFir & "</Run><LineBreak/></TextBlock>"
Case RibTab_Fragebogen:
            CmPa2.Text = vbNullString
Case RibTab_Krankenbla:
            CmPa2.Text = vbNullString
Case RibTab_Abrechnung:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rec_ID1)
                GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                For Each RpRow In RpSel
                    If RpRow.GroupRow = False Then
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
                CmPa2.Text = "Rechnungen markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "Rechnungen markiert : " & RpSel.Count
            End If
Case RibTab_Vorbereit:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Adr_ID0)
                    PatNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Adr_IDKurz)
                    PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Adr_Anschrift)
                    PaAns = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Adr_Telefon1)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then PaTe1 = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Adr_Telefon2)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then PaTe2 = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Adr_Telefon5)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then PaTe3 = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Adr_Geboren)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then PaGeb = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Adr_IDP)
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MaNum = RpRow.Record(RpCol.ItemIndex).Value
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
                Set RpCol = RpCls.Find(30)
                For Each RpRow In RpSel
                    If RpRow.GroupRow = False Then
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                        End If
                    End If
                Next RpRow
                CmPa2.Text = "Patienten markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
                Lbl10.Caption = PaNam & vbCrLf & vbCrLf & PaAns & vbCrLf & vbCrLf & "Geboren: " & PaGeb & vbCrLf & vbCrLf & "Telefon :" & PaTe1 & vbCrLf & "Telefon :" & PaTe2 & vbCrLf & vbCrLf & "EMail :" & PaTe3 & vbCrLf & vbCrLf & BeNam
            End If
Case RibTab_Rezeptmodul:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rzp_ID1)
                GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rzp_Vorname)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    PaNam = RpRow.Record(RpCol.ItemIndex).Value
                End If
                Set RpCol = RpCls.Find(Rzp_Name)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    PaNam = PaNam & Chr$(32) & RpRow.Record(RpCol.ItemIndex).Value
                End If
                Set RpCol = RpCls.Find(Rzp_Geboren)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    PaGeb = RpRow.Record(RpCol.ItemIndex).Value
                End If
            End If
            CmPa2.Text = "Rezepte markiert : " & RpSel.Count
Case RibTab_Belegmodul:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Rzp_ID1)
                GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rzp_Vorname)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    PaNam = RpRow.Record(RpCol.ItemIndex).Value
                End If
                Set RpCol = RpCls.Find(Rzp_Name)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    PaNam = PaNam & Chr$(32) & RpRow.Record(RpCol.ItemIndex).Value
                End If
                Set RpCol = RpCls.Find(Rzp_Geboren)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    PaGeb = RpRow.Record(RpCol.ItemIndex).Value
                End If
            End If
            CmPa2.Text = "Rezepte markiert : " & RpSel.Count
Case RibTab_Bildmodul:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    IdxNr = AdAry(Adr_ID0, RpRow.Index)
                    PaNam = AdAry(Adr_IDKurz, RpRow.Index)
                    If AdAry(Adr_Geboren, RpRow.Index) <> vbNullString Then
                        PaGeb = AdAry(Adr_Geboren, RpRow.Index)
                    Else
                        PaGeb = vbNullString
                    End If
                    If AdAry(Adr_IDP, RpRow.Index) > 0 Then
                        MaNum = AdAry(Adr_IDP, RpRow.Index)
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
            End If
            CmPa2.Text = "Patienten markiert : " & RpSel.Count
Case RibTab_Rechnungen:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Rec_ID1)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Rec_RechNr)
                    RecNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Rec_IDKurz)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaNam = vbNullString
                    End If
                    Set RpCol = RpCls.Find(Rec_Versicherer)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaVer = RpRow.Record(RpCol.ItemIndex).Value
                    End If
                End If
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Rec_IDP)
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MaNum = RpRow.Record(RpCol.ItemIndex).Value
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
                For Each RpRow In RpSel
                    If RpRow.GroupRow = False Then
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
                CmPa2.Text = "Rechnungen markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "Rechnungen markiert : " & RpSel.Count
            End If
Case RibTab_Mahnwesen:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(OPo_ID1)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(OPo_RechNr)
                    RecNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(OPo_Patient)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaNam = vbNullString
                    End If
                End If
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(OPo_IDP)
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MaNum = RpRow.Record(RpCol.ItemIndex).Value
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
                Set RpCol = RpCls.Find(OPo_OffBetrag)
                For Each RpRow In RpSel
                    If RpRow.GroupRow = False Then
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                        End If
                    End If
                Next RpRow
                CmPa2.Text = "Posten markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "Posten markiert : " & RpSel.Count
            End If
Case RibTab_Buchungen:
            EiSum = 0
            AuSum = 0
            GeSum = 0
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Buh_ID0)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Buh_Datum)
                    BuDat = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Buh_Sachkontenbez)
                    BuKto = RpRow.Record(RpCol.ItemIndex).Value
                End If
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Buh_IDT)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                            If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                                MaNum = RpRow.Record(RpCol.ItemIndex).Value
                                For AktPo = 1 To UBound(GlThe)
                                    If MaNum = GlThe(AktPo, 0) Then
                                        BeNam = GlThe(AktPo, 13)
                                        Exit For
                                    End If
                                Next AktPo
                            End If
                        End If
                    End If
                End If
                For Each RpRow In RpSel
                    If RpRow.GroupRow = False Then
                        Set RpCol = RpCls.Find(Buh_Einnahme)
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                                If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                                    EiSum = EiSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                                End If
                            End If
                        End If
                        Set RpCol = RpCls.Find(Buh_Ausgabe)
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                                If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                                    AuSum = AuSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                                End If
                            End If
                        End If
                        GeSum = (EiSum - AuSum)
                    End If
                Next RpRow
                CmPa2.Text = "Buchungen markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "Buchungen markiert : " & RpSel.Count
            End If
Case RibTab_HomeBanki:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ban_ID2)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ban_RechNr1)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        RecNr = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        RecNr = vbNullString
                    End If
                    Set RpCol = RpCls.Find(Ban_Kommentar)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaNam = vbNullString
                    End If
                End If
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ban_IDP)
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MaNum = RpRow.Record(RpCol.ItemIndex).Value
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
                Set RpCol = RpCls.Find(Ban_KoBetrag)
                For Each RpRow In RpSel
                    If RpRow.GroupRow = False Then
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                        End If
                    End If
                Next RpRow
                CmPa2.Text = "Einträge markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "RpRow markiert : " & RpSel.Count
            End If
Case RibTab_Ter_Kalend:
            If GlTeV = True Then 'Termine vorhanden
                AnzTa = CaCol.ActiveView.DaysCount
                If AnzTa > 1 Then
                    DayFi = CaCol.ActiveView.Days(0).Date
                    DayLa = DateAdd("m", 6, DayFi)
                    For AktTa = 0 To AnzTa
                        Set CaEvs = CaCol.DataProvider.RetrieveDayEvents(DayFi + AktTa)
                        AnzTe = AnzTe + CaEvs.Count
                    Next AktTa
                    CmPa2.Text = "Anzahl Termine : " & AnzTe
                Else
                    If CaCol.ViewType <> xtpCalendarTimeLineView Then
                        DayFi = CaCol.ActiveView.Days(0).Date
                        Set CaEvs = CaCol.DataProvider.RetrieveDayEvents(DayFi)
                        CmPa2.Text = "Anzahl Termine : " & CaEvs.Count
                    End If
                End If
            End If
Case RibTab_Ter_Raeume:
            If GlTeV = True Then 'Termine vorhanden
                AnzTa = CaCol.ActiveView.DaysCount
                If AnzTa > 1 Then
                    DayFi = CaCol.ActiveView.Days(0).Date
                    DayLa = DateAdd("m", 6, DayFi)
                    For AktTa = 0 To AnzTa
                        Set CaEvs = CaCol.DataProvider.RetrieveDayEvents(DayFi + AktTa)
                        AnzTe = AnzTe + CaEvs.Count
                    Next AktTa
                    CmPa2.Text = "Anzahl Termine : " & AnzTe
                Else
                    DayFi = CaCol.ActiveView.Days(0).Date
                    Set CaEvs = CaCol.DataProvider.RetrieveDayEvents(DayFi)
                    CmPa2.Text = "Anzahl Termine : " & CaEvs.Count
                End If
            End If
Case RibTab_Ter_Mitarb:
            If GlTeV = True Then 'Termine vorhanden
                AnzTa = CaCol.ActiveView.DaysCount
                If AnzTa > 1 Then
                    DayFi = CaCol.ActiveView.Days(0).Date
                    DayLa = DateAdd("m", 6, DayFi)
                    For AktTa = 0 To AnzTa
                        Set CaEvs = CaCol.DataProvider.RetrieveDayEvents(DayFi + AktTa)
                        AnzTe = AnzTe + CaEvs.Count
                    Next AktTa
                    CmPa2.Text = "Anzahl Termine : " & AnzTe
                Else
                    DayFi = CaCol.ActiveView.Days(0).Date
                    Set CaEvs = CaCol.DataProvider.RetrieveDayEvents(DayFi)
                    CmPa2.Text = "Anzahl Termine : " & CaEvs.Count
                End If
            End If
Case RibTab_Ter_Listen:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ter_ID2)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_VonDat)
                    TeDat = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_IDKurz)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TeBet = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        TeBet = vbNullString
                    End If
                End If
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ter_IDP)
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MaNum = RpRow.Record(RpCol.ItemIndex).Value
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
            End If
            Set RpCol = RpCls.Find(Ter_TerBet)
            For Each RpRow In RpSel
                If RpRow.GroupRow = False Then
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                    End If
                End If
            Next RpRow
            If RpSel.Count > 0 Then
                CmPa2.Text = "Termine markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "Termine markiert : " & RpSel.Count
            End If
Case RibTab_Ter_Akont:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ter_ID2)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_VonDat)
                    TeDat = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_IDKurz)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TeBet = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        TeBet = vbNullString
                    End If
                End If
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ter_IDP)
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MaNum = RpRow.Record(RpCol.ItemIndex).Value
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
            End If
            Set RpCol = RpCls.Find(Ter_TerBet)
            For Each RpRow In RpSel
                If RpRow.GroupRow = False Then
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                    End If
                End If
            Next RpRow
            If RpSel.Count > 0 Then
                CmPa2.Text = "Termine markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "Termine markiert : " & RpSel.Count
            End If
Case RibTab_Ter_Warte:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ter_ID2)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_VonDat)
                    TeDat = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_IDKurz)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TeBet = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        TeBet = vbNullString
                    End If
                End If
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Ter_IDP)
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MaNum = RpRow.Record(RpCol.ItemIndex).Value
                        For AktPo = 1 To UBound(GlThe)
                            If MaNum = GlThe(AktPo, 0) Then
                                BeNam = GlThe(AktPo, 13)
                                Exit For
                            End If
                        Next AktPo
                    End If
                End If
            End If
            Set RpCol = RpCls.Find(Ter_TerBet)
            For Each RpRow In RpSel
                If RpRow.GroupRow = False Then
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        GeSum = GeSum + CDbl(RpRow.Record(RpCol.ItemIndex).Value)
                    End If
                End If
            Next RpRow
            If RpSel.Count > 0 Then
                CmPa2.Text = "Termine markiert : " & RpSel.Count & " - Summe : " & Format$(GeSum, GlWa1)
            Else
                CmPa2.Text = "Termine markiert : " & RpSel.Count
            End If
Case RibTab_LabBericht:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpRow = RpSel(0)
                    Set RpCol = RpCls.Find(Lab_ID0)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lab_Vorname)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaNam = vbNullString
                    End If
                    Set RpCol = RpCls.Find(Lab_Name)
                    PaNam = PaNam & Chr$(32) & RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lab_Geboren)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaGeb = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaGeb = vbNullString
                    End If
                End If
            End If
            CmPa2.Text = "Berichte markiert : " & RpSel.Count
Case RibTab_LabBerichte:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpRow = RpSel(0)
                    Set RpCol = RpCls.Find(Lab_ID0)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lab_Vorname)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaNam = vbNullString
                    End If
                    Set RpCol = RpCls.Find(Lab_Name)
                    PaNam = PaNam & Chr$(32) & RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lab_Geboren)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaGeb = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaGeb = vbNullString
                    End If
                End If
            End If
            CmPa2.Text = "Berichte markiert : " & RpSel.Count
Case RibTab_LabAuftrag:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpRow = RpSel(0)
                    Set RpCol = RpCls.Find(Lau_ID1)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lau_Vorname)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaNam = vbNullString
                    End If
                    Set RpCol = RpCls.Find(Lau_Name)
                    PaNam = PaNam & Chr$(32) & RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lau_Geboren)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaGeb = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaGeb = vbNullString
                    End If
                End If
            End If
            CmPa2.Text = "Aufträge markiert : " & RpSel.Count
Case RibTab_LabAuftrage:
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpRow = RpSel(0)
                    Set RpCol = RpCls.Find(Lau_ID1)
                    GlIdx = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lau_Vorname)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaNam = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaNam = vbNullString
                    End If
                    Set RpCol = RpCls.Find(Lau_Name)
                    PaNam = PaNam & Chr$(32) & RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Lau_Geboren)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PaGeb = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PaGeb = vbNullString
                    End If
                End If
            End If
            CmPa2.Text = "Aufträge markiert : " & RpSel.Count
Case RibTab_Kat_Eintrg:
        If RpSel.Count > 0 Then
            For Each RpRow In RpSel
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Kat_Preis1)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                            GeSum = GeSum + RpRow.Record(RpCol.ItemIndex).Value
                        End If
                    End If
                End If
            Next RpRow
            CmPa2.Text = "Einträge markiert : " & RpSel.Count
        End If
Case RibTab_Kat_Ketten:
        If RpSel.Count > 0 Then
            For Each RpRow In RpSel
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Kat_Preis1)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                            GeSum = GeSum + RpRow.Record(RpCol.ItemIndex).Value
                        End If
                    End If
                End If
            Next RpRow
            CmPa2.Text = "Einträge markiert : " & RpSel.Count
        End If
Case RibTab_Kat_Frage:
        CmPa2.Text = vbNullString
Case RibTab_Kat_Explor:
        CmPa2.Text = vbNullString
Case RibTab_Tex_Dokumt:
        CmPa2.Text = vbNullString
Case RibTab_Tex_Vorlag:
        CmPa2.Text = vbNullString
Case RibTab_Tex_Rezept:
        CmPa2.Text = vbNullString
Case RibTab_Tex_NewsLe:
        CmPa2.Text = vbNullString
Case RibTab_Tex_Email:
        CmPa2.Text = vbNullString
End Select

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo0 = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set RpCo6 = Nothing
Set RpCoK = Nothing
Set RpCo8 = Nothing
Set CaCol = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SMark " & Err.Number
Resume Next

End Sub
Public Function SMCUp(ByVal FiNam As String, Optional ByVal PatNr As Long = 0) As String
On Error GoTo InErr
'Cloudupload Downloadlink Rechnungsupload

Dim PaNum As String
Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim DaNam As String
Dim DaIni As String
Dim IniNa As String
Dim TmpSt As String
Dim TmZei As String
Dim DoLnk As String
Dim FoLnk As String
Dim DaStr As String
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim AryZe() As String

Set FM = frmMain

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smcloud.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smcloud.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

Screen.MousePointer = vbHourglass
DoEvents

If PatNr = 0 Then
    PatNr = GlAdr
End If

PaNum = Format$(PatNr, "000000")

If GlCID <> vbNullString Then 'Cloud-ID
    DaStr = Format(DateAdd("d", Now(), GlVrw), "yyyy-mm-dd")
    PrNam = Chr$(34) & PrNam & Chr$(34)
    IniNa = CreateID("U") & ".ini"
    DaIni = GlTmp & IniNa
    
    With clFil
        .FilPfa FiNam
        DaNam = .DaNam
    End With

    PaStr = "upload" & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & "CustomerFiles" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & PaNum & "/" & DaNam & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--expires=" & DaStr & Space$(1) & "--sharefolder=" & Chr$(34) & PaNum & Chr$(34)
    WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
    DoEvents

    With clFil
        If .FilVor(DaIni) = True Then
            .FilPfa DaIni
            TmpSt = .FilReSt
            DoEvents
            If TmpSt <> vbNullString Then
                AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                For AktZe = 0 To UBound(AryZe) - 1
                    If AryZe(AktZe) <> vbNullString Then
                        TmZei = AryZe(AktZe)
                        Lange = Len(TmZei)
                        Posit = InStr(1, TmZei, "=", 1)
                        If Posit > 0 Then
                            InTyp = LCase(Left$(TmZei, Posit - 1))
                            Select Case InTyp
                            Case "fileurl": DoLnk = Right$(TmZei, Lange - Posit)
                            Case "folderurl": FoLnk = Right$(TmZei, Lange - Posit)
                            Case "urlexpiresat":
                            Case "filename":
                            End Select
                        End If
                    End If
                Next AktZe
            End If
        End If
        DoEvents
        
        If DoLnk <> vbNullString Then
            SMCUp = DoLnk & ";" & FoLnk
        End If
        
        With clFil
            If GlLog = False Then 'General Logging
                .DaLoe = GlTmp & "*.ini" & vbNullChar
                .FilLoe
                If .FilVor(FiNam) = True Then
                    .DaLoe = FiNam & vbNullChar
                    .FilLoe
                End If
            Else
                Clipboard.Clear
                Clipboard.SetText PrNam & Space$(1) & PaStr
            End If
        End With
    End With
End If
        
DoEvents
Screen.MousePointer = vbNormal
        
Set clFil = Nothing

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SMCUp " & Err.Number
Resume Next

End Function
Public Sub SMDDo()
On Error GoTo SuErr
'Abfrufen eines signierten Dokumentes (Dokument Download)

Dim PatNr As Long
Dim Datu1 As Date
Dim Datu2 As Date
Dim PaNum As String
Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim PaNam As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim IniNa As String
Dim TmpSt As String
Dim TmZei As String
Dim RetSt As String
Dim ErrSt As String
Dim ZipNa As String
Dim FiNam As String
Dim ImOrd As String
Dim TmGui As String
Dim TyNam As String
Dim PfaNa As String
Dim DaNam As String
Dim DaExt As String
Dim DocNa As String
Dim NeuNa As String
Dim GuiID As String
Dim KoNam As String
Dim TmStr As String
Dim EiTyp As Integer
Dim AktZa As Integer
Dim AnzDa As Integer
Dim AnzIn As Integer
Dim AktZe As Integer
Dim AktFe As Integer
Dim DocZa As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim ZipOk As Boolean
Dim AryZe() As String
Dim DiNam() As String
Dim DaInf() As String

Set FM = frmMain

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Sub
    End If
End If

Screen.MousePointer = vbHourglass
DoEvents

TeTit = "Dokumente Abrufen"
TeMai = "Sollen die signierten Dokumente jetzt abgerufen werden?"
TeInh = "Beim Abruf der digital unterzeichneten Dokumente werden diese als PDF-Dateien heruntergeladen und automatisch den Patienten zugeordnet."
TeFus = "Die signierten Dokumente werden im Krankenblatt des jeweiligen Patienten hinterlegt und kann von dort aus angezeigt und ausgedruckt werden."

TmGui = CreateID("D")
ZipNa = GlTEx & TmGui & ".zip" 'Termineordner
IniNa = GlTmp & TmGui & ".ini"
ImOrd = GlTEx & "\Temp\"
PrNam = Chr$(34) & PrNam & Chr$(34)

If GlCID <> vbNullString Then 'Cloud-ID
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
    If GlMes = 33565 Then
                
        PaStr = "getsign" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & "--zip=" & Chr$(34) & ZipNa & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & IniNa & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        With clFil
            If .FilVor(IniNa) = True Then
                .FilPfa IniNa
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "documentcount": RetSt = Right$(TmZei, Lange - Posit)
                                Case "error": ErrSt = Right$(TmZei, Lange - Posit)
                                Case "signeddate":
                                Case "signedtime":
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If

            If RetSt <> vbNullString Then
                If IsNumeric(RetSt) = True Then
                    DocZa = CInt(RetSt) 'Anzahl Dokumente
                Else
                    DocZa = 0
                End If
                If DocZa > 0 Then
                    ReDim DaInf(3, DocZa)
                    For AktZe = 0 To UBound(AryZe) - 1
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Left$(TmZei, 1) = "[" Then
                                If Len(TmZei) > 10 Then
                                    AnzIn = AnzIn + 1
                                    DaInf(0, AnzIn - 1) = Mid$(TmZei, 2, Lange - 2)
                                End If
                            End If
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "signeddate":
                                        RetSt = Right$(TmZei, Lange - Posit)
                                        If IsDate(RetSt) = True Then
                                            DaInf(1, AnzIn - 1) = CDate(RetSt)
                                        Else
                                            DaInf(1, AnzIn - 1) = Date
                                        End If
                                Case "signedtime":
                                        RetSt = Right$(TmZei, Lange - Posit)
                                        If IsDate(RetSt) = True Then
                                            DaInf(2, AnzIn - 1) = TimeValue(RetSt)
                                        Else
                                            DaInf(2, AnzIn - 1) = TimeValue(Now)
                                        End If
                                End Select
                            End If
                        End If
                    Next AktZe

                    If .FilVor(ZipNa) = True Then
                        If .FilDir(ImOrd) = False Then
                            MkDir ImOrd
                            DoEvents
                        End If
                        If GlDbg = True Then
                            ZipOk = SZipp(ZipNa, ImOrd, False, True)
                        Else
                            ZipOk = SZipp(ZipNa, ImOrd, True, True)
                        End If
                        DoEvents
                        If ZipOk = True Then
                            If .FilVor(ImOrd & "*.*") = True Then
                                AnzDa = .FilLis(LCase(ImOrd), "*.*", DiNam)
                                DoEvents
                                If AnzDa > 0 Then
                                    For AktZa = 1 To UBound(DiNam)
                                        DaNam = DiNam(AktZa)
                                        Lange = Len(DaNam)
                                        FiNam = ImOrd & DaNam
                                        .FilPfa FiNam
                                        DaExt = .DaExt
                                        PaNum = Mid$(DaNam, 3, 6)
                                        GuiID = Mid$(DaNam, 10, 33)
                                        DocNa = Mid$(DaNam, 44, Lange - 47)
                                        If IsNumeric(PaNum) = True Then
                                            Posit = InStrRev(DaNam, "_", Len(DaNam), 1)
                                            If Posit > 0 Then
                                                KoNam = Mid$(DaNam, Posit + 1, Len(DaNam) - (Posit + 4))
                                            Else
                                                KoNam = "Textdokument"
                                            End If
                                            If LCase(DaExt) = "pdf" Then
                                                PatNr = CLng(PaNum)
                                                PaNam = S_AdIdx(PatNr, "IDKurz")
                                                DoEvents
                                                TyNam = "PDF-Dokument"
                                                PfaNa = GlBPf
                                                EiTyp = 105
                                                NeuNa = "SI" & Mid$(DaNam, 3, Len(DaNam) - 2)
                                                .DaCop = FiNam & ";" & PfaNa & NeuNa & vbNullChar
                                                If .FilCop(2) = True Then
                                                    DoEvents
                                                    Datu1 = Date
                                                    Datu2 = TimeValue(Now)
                                                    
                                                    For AktZe = 0 To UBound(DaInf)
                                                        If DaInf(0, AktZe) = GuiID Then
                                                            Datu1 = DaInf(1, AktZe)
                                                            Datu2 = DaInf(2, AktZe)
                                                            Exit For
                                                        End If
                                                    Next AktZe
                                                    
                                                    GlNeK = GlKoX
                                                    With GlNeK
                                                        .PatNr = PatNr
                                                        .IdxNr = 0
                                                        .EiDat = Datu1
                                                        .EiZei = Datu2
                                                        .EiTyp = EiTyp
                                                        .KoStr = NeuNa
                                                        .TeStr = TyNam
                                                        .TeStr = KoNam
                                                        .NeuEi = True
                                                        .Mitar = GlMiA(GlSmI, 2)
                                                    End With
                                                    K_Einf
                                                    DoEvents
                                                    If TmStr = vbNullString Then
                                                        TmStr = vbCrLf & Datu1 & " - " & Datu2 & " - " & PaNam & " - " & DocNa & vbCrLf
                                                    Else
                                                        TmStr = TmStr & vbCrLf & Datu1 & " - " & Datu2 & " - " & PaNam & " - " & DocNa & vbCrLf
                                                    End If
                                                End If
                                            Else
                                                SPopu "Falscher Dateiname", "Die heruntergeladene Datei hat den falschen Dateinamen", IC48_Warning
                                            End If
                                        Else
                                            SPopu "Dateiname nicht lesbar", "Der von Ihnen ausgewählte Dateiname kann nicht gelesen werden weil er ggf. Sonderzeichen enthält", IC48_Warning
                                        End If
                                    Next AktZa
                                End If
                            End If
                        End If
                        DoEvents
                        SUpKr
                        DoEvents
                        SKrVo
                        DoEvents
                        If DocZa > 1 Then
                            SPopu "Dokumentendownload", "Es wurden " & DocZa & " signierte Dokumente heruntergeladen und eingefügt", IC48_Information
                        Else
                            SPopu "Dokumentendownload", "Es wurden ein neues signiertes Dokument heruntergeladen und eingefügt", IC48_Information
                        End If
                        If TmStr <> vbNullString Then
                            frmTSEInit.Show
                            frmTSEInit.Caption = "Dokumentenabruf"
                            frmTSEInit.txtTSEIn.Text = vbCrLf & "Die folgenden Dokumente wurden signiert:" & vbCrLf & TmStr
                        End If
                    Else
                        If ErrSt = vbNullString Then
                            SPopu "Downloadfehler", "Unerwarteter Fehler, beim Herunterladen der Dokumente", IC48_Forbidden
                        Else
                            SPopu "Downloadfehler", ErrSt, IC48_Forbidden
                        End If
                    End If
                Else
                    SPopu "Dokumentenabruf", "Es liegen keine personalisierte und signierte Dokumente zum Download bereit.", IC48_Information
                End If
            Else
                Clipboard.Clear
                Clipboard.SetText PrNam & Space$(1) & PaStr
                If ErrSt = vbNullString Then
                    SPopu "Downloadfehler", "Unerwarteter Fehler, beim Herunterladen der Dokumente", IC48_Forbidden
                Else
                    SPopu "Downloadfehler", ErrSt, IC48_Forbidden
                End If
            End If

            If ErrSt = vbNullString Then
                If GlLog = False Then 'General Logging
                    .DaLoe = GlTmp & "*.ini" & vbNullChar
                    .FilLoe
                    If .FilVor(ImOrd & "*.*") = True Then
                        .DaLoe = ImOrd & "*.*" & vbNullChar
                        .FilLoe
                    End If
                    If .FilVor(ZipNa) = True Then
                        .DaLoe = ZipNa & vbNullChar
                        .FilLoe
                    End If
                Else
                    Clipboard.Clear
                    Clipboard.SetText PrNam & Space$(1) & PaStr
                End If
            End If
        End With
            
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then
    MsgBox Err.Description, 48, "SMDDo " & Err.Number
    SErLog Err.Description & " SMDDo " & Err.Number
End If
Resume Next

End Sub
Public Function SMDUp(ByVal FiNam As String, Optional ByVal PatNr As Long = 0, Optional ByVal MaEma As String, Optional ByVal MaBrf As String, Optional ByVal DaNam As String) As String
On Error GoTo InErr
'Upload eines Dokumemtes zur Signierung

Dim MitNr As Long
Dim ManNr As Long
Dim PaNum As String
Dim PatNa As String
Dim PatVo As String
Dim PatEm As String
Dim MaNam As String
Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim DaNaO As String
Dim DaIni As String
Dim IniNa As String
Dim TmpSt As String
Dim TmZei As String
Dim DoLnk As String
Dim SigId As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim KonZa As Integer
Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim AryZe() As String

Set FM = frmMain

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

If GlEKV = False Then 'Emailkonten vorhanden
    TeTit = "E-Mail-Versand"
    TeMai = "Es ist kein E-Mail-Konto vorhanden"
    TeInh = "Um eine E-Mail-Bestätigung für ein signiertes Dokument zu erhalten, ist es notwendig mind. ein E-Mail-Konto hinzuzufügen."
    TeFus = "Um ein E-Mail-Konto hinzuzufügen, wechseln Sie in das Modul: Textverarbeitung und dann oben auf Emails. Dort klicken Sie auf die Schaltfläche Emailkonten."
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, True, FM.hwnd
    Exit Function
End If

If PatNr = 0 Then
    PatNr = GlAdr
End If

PaNum = Format$(PatNr, "000000")

MitNr = GlMiA(GlSmI, 2)
KonZa = UBound(GlMkt)

If MaEma = vbNullString Then
    If KonZa > 0 Then
        For AktZa = 1 To KonZa
            If CLng(GlMkt(AktZa, 1)) = MitNr Then
                If CBool(GlMkt(AktZa, 20)) = True Then 'Standardemailkonto
                    If GlMiA(AktZa, 13) <> vbNullString Then
                        MaEma = GlMiA(AktZa, 13)
                        Exit For
                    End If
                End If
            End If
        Next AktZa
    End If
    
    For AktZa = 1 To UBound(GlMiA)
        If MitNr = GlMiA(AktZa, 2) Then
            ManNr = GlMiA(AktZa, 7)
            If MaEma = vbNullString Then
                If GlMiA(AktZa, 22) <> vbNullString Then
                    MaEma = GlMiA(AktZa, 22)
                End If
            End If
        End If
    Next AktZa
End If

For AktZa = 1 To UBound(GlThe) 'Mandanten
    If ManNr = GlThe(AktZa, 0) Then
        MaNam = GlThe(AktZa, 13)
        Exit For
    End If
Next AktZa

TeTit = "Digitale Unterschrift"
TeMai = "Soll dieses Dokument nun zur digitalen Unterschrift bereitgestellt werden?"
TeInh = "Durch die sichere Bereitstellung dieses Dokument im Internet, wird ein Link generiert, der dem Patienten gesendet werden kann."
TeFus = "Wurde dieses Dokument vom Patienten digital unterschrieben, wird eine Benachrichtigung an: " & MaEma & " gesendet, so dass es danach abgerufen und diesem Patienten automatisch im Krankenblatt eingefügt werden kann."

If Len(MaBrf) > 200 Then
    MaBrf = Left$(MaBrf, 200)
End If

If GlCID <> vbNullString Then 'Cloud-ID
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
    If GlMes = 33565 Then
        Screen.MousePointer = vbHourglass
        DoEvents
    
        PrNam = Chr$(34) & PrNam & Chr$(34)
        IniNa = CreateID("U") & ".ini"
        DaIni = GlTmp & IniNa
        
        With clFil
            .FilPfa FiNam
            DaNaO = .DaNaO
        End With
        
        S_AdDe PatNr 'Adressendetails
        With GlADt
            PatNa = .AdNam
            PatVo = .AdVor
            PatEm = .AdTe5
        End With
        DoEvents
    
        PaStr = "createsign" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaNam & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & PaNum & Chr$(34) & Space$(1) & Chr$(34) & PatVo & Chr$(34) & Space$(1) & Chr$(34) & PatNa & Chr$(34) & Space$(1) & Chr$(34) & PatEm & Chr$(34) & Space$(1) & Chr$(34) & DaNaO & Chr$(34) & Space$(1) & Chr$(34) & DaNam & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & FiNam & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--ocr=0" & Space$(1) & "--days=" & GlVrw
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents
        
        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "signurl": DoLnk = Right$(TmZei, Lange - Posit)
                                Case "signrequestId": SigId = Right$(TmZei, Lange - Posit)
                                Case "documentId":
                                Case "success":
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
            DoEvents
            
            If DoLnk <> vbNullString Then
                SMDUp = DoLnk
            End If
            
            With clFil
                If GlLog = False Then 'General Logging
                    .DaLoe = GlTmp & "*.ini" & vbNullChar
                    .FilLoe
                    If .FilVor(FiNam) = True Then
                        .DaLoe = FiNam & vbNullChar
                        .FilLoe
                    End If
                Else
                    Clipboard.Clear
                    Clipboard.SetText PrNam & Space$(1) & PaStr
                End If
            End With
        End With
        
        DoEvents
        Screen.MousePointer = vbNormal
    End If
End If
        
Set clFil = Nothing

Exit Function

InErr:
SuErr:
If GlDbg = True Then
    MsgBox Err.Description, 48, "SMDUp " & Err.Number
    SErLog Err.Description & " SMDUp " & Err.Number
End If

Resume Next

End Function
Public Sub SMeAc()
On Error GoTo LaErr
'Erstellt die CommandBar

Dim RetWe As Long
Dim PfNa1 As String
Dim PfNa2 As String
Dim PfNa3 As String
Dim PfNa4 As String
Dim AktZa As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim ImMan As XtremeCommandBars.ImageManager
Dim GrIco As Long
Dim KlIco As Long

Set FM = frmMain
Set CmBrs = FM.comBar01
Set ImMan = FM.imgManag
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

PfNa1 = IniGetVal("System", "WegPfa")
PfNa2 = IniGetVal("System", "MedPfa")
PfNa4 = IniGetVal("System", "GDTPrg")
If GlRDP = False Then
    PfNa3 = App.Path & "\Fernwartung14.exe"
End If
    
Select Case GlSkn
Case 1:
    With ImMan
        .Icons.LoadBitmap App.Path & "\Skins\skn16v.skn", Tol16, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn17v.skn", Tas16, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn24v.skn", Tol24, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn25v.skn", Tas24, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn32v.skn", Tol32, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn48v.skn", Tol48, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn64v.skn", Tol64, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn128v.skn", Tol128, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\Alphabet.skn", SuAry, xtpImageNormal
    End With
Case 2:
    With ImMan
        .Icons.LoadBitmap App.Path & "\Skins\skn16g.skn", Tol16, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn17g.skn", Tas16, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn24g.skn", Tol24, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn25g.skn", Tas24, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn32g.skn", Tol32, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn48g.skn", Tol48, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn64g.skn", Tol64, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn128g.skn", Tol128, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\Alphabet.skn", SuAry, xtpImageNormal
    End With
Case 3:
    With ImMan
        .Icons.LoadBitmap App.Path & "\Skins\skn16o.skn", Tol16, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn17o.skn", Tas16, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn24o.skn", Tol24, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn25o.skn", Tas24, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn32o.skn", Tol32, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn48o.skn", Tol48, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn64o.skn", Tol64, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\skn128o.skn", Tol128, xtpImageNormal
        .Icons.LoadBitmap App.Path & "\Skins\Alphabet.skn", SuAry, xtpImageNormal
    End With
End Select

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(RibCon_Ifap, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_SmartCard, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Aufgaben, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Abmelden, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Refresh, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Textvor, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Optionen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Formulare, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_KB_Ansicht, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_AB_Ansicht, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_TP_Ansicht, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_LA_Ansicht, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_TE_Ansicht, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_ST_Ansicht, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_LI_Ansicht, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Zeilenumbruch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Zeilenmarker, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Gitternetzlin, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_MenAnimantion, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Vorschauzeile, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Lizenz, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibCon_Tooltips, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Zeierf, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Diagno, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Multip, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Mitarb, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Einhei, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_LaBetr, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_EinTyp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Steuer, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_TabMod, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_Sorter, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Analog, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_StoRec, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_StoEin, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_Restri, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Spa_Vorsch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_Vorsch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_ZeiUmb, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_Zeierf, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_Multip, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_Mitarb, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_Einhei, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_Steuer, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_Spa_Zeierf, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_Spa_Diagno, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_Spa_Multip, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_Spa_Mitarb, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_Spa_Einhei, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_Spa_Steuer, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_Ziffer, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Dia_DatuZe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Dia_ICDZei, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_PZNCod, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Spa_Betrag, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Kra_DirBea, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Kra_RecDet, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_Antliz, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_AufMed, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Kra_AufDia, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Zei_Toltip, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_Kra_FliTex, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Adresse_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Adresse_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Beleg_Stornierte, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AD_Adresse_Hinzufueg, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AD_Adresse_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Verord_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Verord_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VE_Verord_Hinzufueg, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VE_Verord_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Mandant_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Mandant_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_MA_Mandant_Hinzufueg, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_MA_Mandant_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AN_AnaBog_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AN_AnaBog_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Hinzufueg, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Quart, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Monat, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Woche, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Datum, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Expor, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Nachr, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_DoLnk, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Email, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_KB_KraBla_Umschalt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Abrech_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Abrech_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Ausschneiden, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Abrech_Grupp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Abrech_Sorti, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Abrech_Expan, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LB_Abrech_Grupp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LB_Abrech_GrZei, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LB_Abrech_Expan, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_TagPro_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RZ_Rezept_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RZ_Rezept_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RZ_Beleg_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RZ_Beleg_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BI_Bild_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LB_Labor_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LA_Auftrag_AdrSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BI_Bild_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LB_Labor_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LA_Auftrag_AdrBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Kopieren, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Ausschneiden, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Terminliste_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Terminliste_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Terminliste_Kopieren, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Terminliste_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Raumtermin_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Raumtermin_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Raumtermin_Kopieren, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Raumtermin_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KM_Eint_Diagnose, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_Hinzufu, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AD_Adresse_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_MA_Mandant_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VE_Verord_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_MI_Mitarb_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RE_Rechnung_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_PO_Posten_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BU_Buchung_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BA_Banking_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BU_Buchung_BuchEinfach, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BU_Buchung_BuchDoppelt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AB_Abrech_Zahlung, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Druck01, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Druck02, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LB_Labor_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LA_Auftrag_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_TagPro_Export, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AD_Adresse_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_MA_Mandant_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VE_Verord_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RE_Rechnung_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_PO_Posten_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BU_Buchung_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BA_Banking_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Hinzufuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Bearbeiten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Filt1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Filt2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Filt3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_SyncReset, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_OnTeReset, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_Terminfar, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LB_Labor_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_LA_Auftrag_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_Tag, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_ArWoche, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_ErWoche, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_Woche, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_Monat, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_FiltTyp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_FiltIdx, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BU_Buchung_Saldo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Plac1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap01, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap02, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap04, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap06, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap07, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap10, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap11, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_Cap12, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AD_Adresse_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_MA_Mandant_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VE_Verord_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_MI_Mitarb_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RE_Rechnung_SuchCombp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_PO_Posten_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BU_Buchung_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_BA_Banking_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TL_Terminliste_SortFeld, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_RE_Rechnung_Belegtyp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AD_Adresse_SortFeld, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AN_AnaBog_Grupp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AN_AnaBog_Sorti, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_AN_AnaBog_Expan, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Frage_SuchCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_SucCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Mail_SorCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Mail_TexCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_KatCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_KatCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Mail_KatCombo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuTex, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuMan, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuMit, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuRau, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuDat, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuWek, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuMon, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuJah, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuBuh, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuBut, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuAbg, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuSta, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuTSt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuZug, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_TagPro_Quart, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_TagPro_Monat, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_TagPro_Woche, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TP_TagPro_Datum, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Abges, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Quart, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Monat, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Woche, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Datum, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Typ01, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Typ02, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Typ03, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_VB_Vorbe_Mitar, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_Spalte, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_Fonts, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTKo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTKl, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlMZe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlMFa, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTGs, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTeD, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTDe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTTe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlDeT, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTrD, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTVe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTSt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTZe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlTeS, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_TE_Termin_GlMiW, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termine_Abschlus, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termine_Dokument, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_EI_Gruppierung, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_01, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_05, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_10, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_15, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_20, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_30, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_60, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termin_Minuten_120, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termine_Status1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termine_Status2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termine_Status3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ME_Termine_Status4, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Erneut, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Empfangen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Antworten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Mail_Antworten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Mail_Rechnung, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Sta_ChCol, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Sta_ChBar, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Sta_ChPie, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Sta_ChLin, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Sta_ChAre, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Sta_ChDon, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ZeiAb1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ZeiAb2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ZeiAb3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ZeiAb4, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ZeiAb5, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ZeiAb6, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ZeiAb7, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_PaSuch, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_PaBear, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_PaFilt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_PaAlle, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForFet, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForKur, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForUnt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrLi, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrRe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrZe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrBl, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Aufzah, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Numeri, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FntHoh, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FntTif, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_EinzLi, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_EinzRe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Abstan, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_KopZei, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FusZei, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_KopDat, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FusZal, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_TexMar, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaNePa, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaNeAr, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaNeVe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaNePl, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaNeAs, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaNeVo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaNeSe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DatSpe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DatSav, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DaFeVe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DocVor, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DocExp, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DocMa1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DocMa2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DocSe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DocSe2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DocSe3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_EtiDru, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_EinTex, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Eigens, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Start, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Adresse, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Kranken, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Abrechn, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Finanz, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Termin, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Labor, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Texte, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(ShoCut_Katalog, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 13001 To 13020
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
    For AktZa = 13100 To 13129
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")

Set CmSys = RbBar.AddSystemButton()

Set CmCoS = CmSys.CommandBar.Controls
With CmCoS
    If GlTyp < 2 Then
        Set CmCon = .Add(xtpControlButton, RibCon_Dbd_Oeffnen, "Serverordner Wählen")
        CmCon.IconId = IC16_Data_View
        CmCon.Enabled = Not GlRDP
    Else
        Set CmCon = .Add(xtpControlButton, RibCon_Dbd_Oeffnen, "Datenbank Öffnen")
        CmCon.IconId = IC16_Data_View
        CmCon.Enabled = Not GlRDP
    End If
    If GlTyp < 2 Then
        Set CmCon = .Add(xtpControlButton, RibCon_Dbd_Wechsel, "Datenbank Wechseln")
        CmCon.IconId = IC16_Data_View
        CmCon.Visible = GlDBc 'SQL Datenbankwechsel erlauben
    End If
    Set CmCon = .Add(xtpControlButton, RibCon_Dbd_Hinzuf, "Datenbankdatei Hinzufügen")
    CmCon.IconId = IC16_Data_Add
    If GlTyp < 2 Then CmCon.Enabled = False
    Set CmCon = .Add(xtpControlButton, RibCon_Dbd_Pruefen, "Datenüberprüfungslauf")
    CmCon.IconId = IC16_Data_Check
    Set CmCon = .Add(xtpControlButton, RibCon_Dbd_GoBDExp, "GoBD Datenexport")
    CmCon.IconId = IC16_Data_Export
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, RibCon_Dbd_TAR_Exp, "TAR Datenexport")
    CmCon.IconId = IC16_Data_Export
    Set CmCon = .Add(xtpControlButton, RibCon_Dbd_GoBAbsh, "GoBD Festschreibung")
    CmCon.IconId = IC16_Lock
    Set CmCon = .Add(xtpControlButton, RobCon_Dbd_Sichern, "Datenausgabe Erstellen")
    CmCon.IconId = IC16_Data_Disk
    Set CmCon = .Add(xtpControlButton, RibCon_Dbd_Zurueck, "Datenausgabe Zurückspielen")
    If GlTyp < 2 Then CmCon.Enabled = False
    Set CmCon = .Add(xtpControlButton, RibCon_Outlookabg, "Outlookabgleich")
    CmCon.IconId = IC16_Organizer
    CmCon.BeginGroup = True
    If GlRDP = True Then
        CmCon.Enabled = False
    ElseIf GlESy = True Then 'CalDAV / CardDAV / Exchange Synchronisation
        CmCon.Enabled = False
    End If
    Set CmCon = .Add(xtpControlButton, RibCon_Formulare, "Formulardesigner")
    CmCon.IconId = IC16_Ruler
    CmCon.ShortcutText = "Strg+F"
    Set CmCon = .Add(xtpControlButton, RibCon_Optionen, "Einstellungen")
    CmCon.IconId = IC16_Hardware
    CmCon.ShortcutText = "Strg+O"
    Set CmCon = .Add(xtpControlButton, RibCon_Benutzer, "Benutzerdaten")
    CmCon.IconId = IC16_User_Norm
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, RibCon_Lizenz, "Lizenzierung")
    CmCon.IconId = IC16_Certificate
    CmCon.Enabled = Not GlRDP
    Set CmCon = .Add(xtpControlButton, RibCon_PrgInfo, "Info")
    CmCon.IconId = IC16_Sign_Info
    CmCon.Enabled = Not GlRDP
    If GlRDP = False Then
        If Dir$(PfNa3, vbNormal) <> vbNullString Then
            Set CmCon = CmBrs.CreateCommandBarControl("CXTPRibbonControlSystemPopupBarButton")
            CmCon.id = RibCon_Remote
            CmCon.Caption = "Fernsteuerung"
            CmCon.ToolTipText = "Öffnet das Fernsteuerungs Programm"
            .AddControl CmCon
        End If
    Else
        Set CmCon = CmBrs.CreateCommandBarControl("CXTPRibbonControlSystemPopupBarButton")
        CmCon.id = RibCon_Neustart
        CmCon.Caption = "Neustart"
        CmCon.ToolTipText = "Führt einen Neustart des Programms durch"
        .AddControl CmCon
    End If
    Set CmCon = CmBrs.CreateCommandBarControl("CXTPRibbonControlSystemPopupBarButton")
    CmCon.id = RibCon_Beenden
    CmCon.Caption = "Beenden"
    CmCon.ToolTipText = "Beendet das Programm"
    CmCon.ShortcutText = "F11"
    .AddControl CmCon
End With

Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Aufgaben, "Aufgabenliste")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Clipboard_Norm
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F9"
End With

Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_SmartCard, "Chipkarte")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_SmartCard
    .ToolTipText = "Liest die eGK / KVK ein"
    .Style = xtpButtonIconAndCaption
End With

Set CmPop = RbBar.Controls.Add(xtpControlPopup, RibCon_ST_Ansicht, "Ansicht")
With CmPop
    .flags = xtpFlagRightAlign
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, RibCon_SaveLayout, "Layout Speichern")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, RibCon_Reset, "Zurücksetzen")
    CmCon.Enabled = Not GlRDP
    Set CmCon = .Add(xtpControlButton, RibCon_Layout, "Layoutoptionen")
    CmCon.BeginGroup = True
End With

Set CmPop = RbBar.Controls.Add(xtpControlPopup, RibCon_LI_Ansicht, "Ansicht")
With CmPop
    .flags = xtpFlagRightAlign
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert die Tabellenansicht"
    .Style = xtpButtonIconAndCaption
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, RibCon_Zeilenmarker, "Zeilenmarker")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, RibCon_Gitternetzlin, "Gitternetzlinien")
    Set CmCon = .Add(xtpControlButton, RibCon_Vorschauzeile, "Vorschauzeile")
    Set CmCon = .Add(xtpControlButton, RibCon_Tooltips, "Zeilentooltips")
    Set CmCon = .Add(xtpControlButton, RibCon_SaveLayout, "Layout Speichern")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, RibCon_Reset, "Zurücksetzen")
    CmCon.Enabled = Not GlRDP
    Set CmCon = .Add(xtpControlButton, RibCon_Schnelldruck, "Schnelldruck")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, RibCon_Layout, "Layoutoptionen")
    CmCon.BeginGroup = True
End With

Set CmPop = RbBar.Controls.Add(xtpControlPopup, RibCon_KB_Ansicht, "Anzeigeopt.")
With CmPop
    .flags = xtpFlagRightAlign
    .IconId = IC16_DouChk
    .ToolTipText = "Blendet zusätzliche Spalten ein oder aus"
    .Style = xtpButtonIconAndCaption
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_AufDia, "Aufnahmediagnosen")
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_AufMed, "Aufnahmemedikamente")
    Set CmCon = .Add(xtpControlButton, SY_KB_Kra_KraBla, "Krankenblatttabelle")
    Set CmCon = .Add(xtpControlButton, SY_KB_Kra_FliTex, "Krankenblattdokument")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_EinTyp, "Eintragstypspalte")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, SY_KB_Spa_Ziffer, "Ziffernspalte")
    Set CmCon = .Add(xtpControlButton, SY_KB_Spa_Betrag, "Betragsspalte")
    If GlMsp = True Then 'Mandantenspalte anstelle von Mitarbeiterspalte in Abrechnung
        Set CmCon = .Add(xtpControlButton, SY_KB_Spa_Mitarb, "Mandantenspalte")
    Else
        Set CmCon = .Add(xtpControlButton, SY_KB_Spa_Mitarb, "Mitarbeiterspalte")
    End If
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_Antliz, "Antlitzbildanzeige")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_Vorsch, "Vorschauzeile")
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_ZeiUmb, "Zeilenumbruch")
    Set CmCon = .Add(xtpControlButton, SY_AB_Zei_Toltip, "Zeilentooltips")
    Set CmCon = .Add(xtpControlButton, SY_KB_Kra_DirBea, "Direktbearbeitung")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, SY_AB_Dia_DatuZe, "Diagnosedatum")
    Set CmCon = .Add(xtpControlButton, SY_KB_Kra_RecDet, "Rechnungsinhalt")
    Set CmCon = .Add(xtpControlButton, SY_AB_Dia_ICDZei, "ICD-10 Codes")
    Set CmCon = .Add(xtpControlButton, SY_KB_Spa_PZNCod, "PZN-Codes")
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_StoEin, "Stornierte Einträge")
    CmCon.BeginGroup = True
End With

Set CmPop = RbBar.Controls.Add(xtpControlPopup, RibCon_AB_Ansicht, "Anzeigeopt.")
With CmPop
    .flags = xtpFlagRightAlign
    .IconId = IC16_DouChk
    .ToolTipText = "Blendet zusätzliche Spalten ein oder aus"
    .Style = xtpButtonIconAndCaption
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_EinTyp, "Eintragstypspalte")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Multip, "Multiplikatorspalte")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Steuer, "Steuersatzspalte")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Diagno, "Diagnosespalte")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Analog, "Analogspalte")
    If GlMsp = True Then 'Mandantenspalte anstelle von Mitarbeiterspalte in Abrechnung
        Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Mitarb, "Mandantenspalte")
    Else
        Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Mitarb, "Mitarbeiterspalte")
    End If
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Zeierf, "Zeiterfassungsspalte")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Einhei, "Einheitenspalte")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Katalo, "Katalogspalte")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_LaBetr, "Einstandspreis")
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_Antliz, "Antlitzbild Zeigen")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Vorsch, "Vorschauzeile")
    Set CmCon = .Add(xtpControlButton, SY_AB_Zei_Toltip, "Zeilentooltips")
    Set CmCon = .Add(xtpControlButton, SY_AB_Dia_DatuZe, "Diagnosedatum")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, SY_AB_Dia_ICDZei, "ICD-10 Codes")
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_Sorter, "Eingabesortierung")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_EigDia, "Manuelle Diagnosen")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_TabMod, "Arbeitsblattmodus")
    CmCon.BeginGroup = True
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_StoRec, "Stornierte Rechnungen")
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_Restri, "Regelprüfung")
End With

Set CmPop = RbBar.Controls.Add(xtpControlPopup, RibCon_TP_Ansicht, "Anzeigeopt.")
With CmPop
    .flags = xtpFlagRightAlign
    .IconId = IC16_DouChk
    .ToolTipText = "Blendet zusätzliche Spalten ein oder aus"
    .Style = xtpButtonIconAndCaption
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    If GlMsp = True Then 'Mandantenspalte anstelle von Mitarbeiterspalte in Abrechnung
        Set CmCon = .Add(xtpControlButton, SY_TP_Spa_Mitarb, "Mandantenspalte")
    Else
        Set CmCon = .Add(xtpControlButton, SY_TP_Spa_Mitarb, "Mitarbeiterspalte")
    End If
    Set CmCon = .Add(xtpControlButton, RibCon_Schnelldruck, "Schnelldruck")
    CmCon.BeginGroup = True
End With

Set CmPop = RbBar.Controls.Add(xtpControlPopup, RibCon_LA_Ansicht, "Anzeigeopt.")
With CmPop
    .flags = xtpFlagRightAlign
    .IconId = IC16_DouChk
    .ToolTipText = "Blendet zusätzliche Spalten ein oder aus"
    .Style = xtpButtonIconAndCaption
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_AB_Kra_Antliz, "Antlitzbild Zeigen")
    Set CmCon = .Add(xtpControlButton, SY_AB_Zei_Toltip, "Zeilentooltips")
    Set CmCon = .Add(xtpControlButton, SY_AB_Spa_Vorsch, "Vorschauzeile")
    CmCon.BeginGroup = True
End With

Set CmPop = RbBar.Controls.Add(xtpControlPopup, RibCon_TE_Ansicht, "Anzeigeopt.")
With CmPop
    .flags = xtpFlagRightAlign
    .IconId = IC16_DouChk
    .ToolTipText = "Ändert bestimmte Einstellungen der Kalenderanzeige"
    .Style = xtpButtonIconAndCaption
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_Fonts, "Schriftartanpassung")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_Spalte, "Zeitspaltenbreite")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTKo, "Navigationsleiste")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTKl, "Reduzierte Leiste")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTeS, "Zeige entfernte Termine")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTrD, "Mitarbeitertermindetails")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlMZe, "Sprechzeitenanzeige")
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlMFa, "Farbunterscheidung")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTGs, "Geschlechteranzeige")
    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlMiW, "Mitarbeiterauswahl")
    Else
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlMiW, "Mandantenauswahl")
    End If
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTeD, "Termindetailsanzeige")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTVe, "Terminverschiebung")
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTZe, "Terminzeitanpassung")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTSt, "Starre Termintaktung")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTTe, "Telefo. Terminbetreff")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlTDe, "Erweit. Terminbetreff")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_GlDeT, "Erweit. Termindetails")
End With

If GlIPC = True Then
    Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Ifap, "ifap3")
    With CmBuT
        .flags = xtpFlagRightAlign
        .IconId = IC16_ifap3
        .ToolTipText = "Öffnet das ifap PraxisCenter"
        .Style = xtpButtonIconAndCaption
    End With
Else
    If GlAbd = True Then
        Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Ifap, "ABDATA")
        With CmBuT
            .flags = xtpFlagRightAlign
            .IconId = IC16_Pills
            .ToolTipText = "Öffnet das ifap PraxisCenter"
            .Style = xtpButtonIconAndCaption
        End With
    End If
End If

Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Refresh, "Aktualisieren")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Plugin
    .ToolTipText = "Aktualisiert die Einträge in einer Netzwerkumgebung"
    .Style = xtpButtonIconAndCaption
End With

If Dir$(PfNa1, vbNormal) <> vbNullString Then
    RetWe = ExtractIconEx(PfNa1, 0, GrIco, KlIco, 1)
    ImMan.Icons.AddIcon KlIco, Prg_Icn1, xtpImageNormal
    Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Wegamed, "WEGAMED")
    With CmBuT
        .flags = xtpFlagRightAlign
        .IconId = Prg_Icn1
        .ToolTipText = "Öffnet das WEGAMED Programm"
        .Style = xtpButtonIconAndCaption
    End With
    DestroyIcon RetWe
End If

If Dir$(PfNa4, vbNormal) <> vbNullString Then
    RetWe = ExtractIconEx(PfNa4, 0, GrIco, KlIco, 1)
    ImMan.Icons.AddIcon KlIco, Prg_Icn4, xtpImageNormal
    Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_GDT_Appli, "GDT-App")
    With CmBuT
        .flags = xtpFlagRightAlign
        .IconId = Prg_Icn4
        .ToolTipText = "Öffnet das GDT Programm"
        .Style = xtpButtonIconAndCaption
    End With
    DestroyIcon RetWe
End If

Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Hilfe, "Hilfe")
With CmBuT
    .ToolTipText = "Öffnet die Kurzhilfe"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
End With

Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Abmelden, "Abmelden")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .Style = xtpButtonIconAndCaption
    .ToolTipText = "Meldet den aktuellen Mitarbeiter ab"
    .ShortcutText = "F12"
End With

Set CmBuT = RbBar.Controls.Add(xtpControlButton, RibCon_Beenden, "Schließen")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .Style = xtpButtonIconAndCaption
    .ToolTipText = "Beendet das Programm"
    .ShortcutText = "F11"
End With

Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

If GlTxM = False Then 'Serienbriefmodus
    CmAcs(Tex_PaFilt).Visible = False
    CmAcs(Tex_PaAlle).Visible = False
End If

If CmAcs(Tex_DaNeSe).Enabled = False Then
    CmAcs(Tex_DaNeSe).Enabled = True
End If

Set CmSys = Nothing
Set RbBar = Nothing
Set CmAcs = Nothing
Set CmSta = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub
LaErr:
If GlDbg = True Then SErLog Err.Description & " SMeAc " & Err.Number
Resume Next

End Sub
Public Sub SMeMs(ByVal MsHed As String, ByVal MsMai As String, ByVal IcnId As Long)
On Error GoTo PoErr
'MessageBar

Dim MsBar As XtremeCommandBars.MessageBar
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set MsBar = CmBrs.MessageBar

With MsBar
    If .Visible = False Then
        .AddCloseButton "Schließt die Nachrichtenspalte"
        .Message = _
            "<StackPanel Orientation='Horizontal'>" & _
            "        <Image Source='" & IcnId & "'/>" & _
            "        <TextBlock Padding='3, 0, 0, 0' VerticalAlignment='Center'><Bold>" & MsHed & "</Bold></TextBlock>" & _
            "        <TextBlock Padding='10, 0, 0, 0' VerticalAlignment='Center'>" & MsMai & "</TextBlock></StackPanel>"
        .Visible = True
    End If
End With

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SMeMs " & Err.Number
Resume Next

End Sub
Public Sub SMePa()
On Error GoTo LaErr
'CommandBar Einstellungen

Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmMain
Set ImMan = FM.imgManag
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions
Set CmOpt = CmBrs.Options
Set CmSta = CmBrs.StatusBar
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

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
    .LargeIcons = True 'WICHTIG!
    .LunaColors = GlLun
    .MaxPopupWidth = 0.5
    .OfficeStyleDisabledIcons = True
    .SetIconSize True, 32, 32
    .ShowExpandButtonAlways = False
    .ShowFullAfterDelay = True
    .ShowPopupBarToolTips = False
    .ShowTextBelowIcons = True
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
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs '\\\
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case 8:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case Else:
        If GlRah = True Then 'Office EnableThemeframe
            .VisualTheme = xtpThemeRibbon
        Else
            If GlFRg = True Then 'farbige Register
                .VisualTheme = xtpThemeResource
            Else
                .VisualTheme = xtpThemeRibbon
            End If
        End If
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End Select
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = True
    .ToolTipContext.ShowOfficeBorder = True
    .ToolTipContext.ShowShadow = True
    .ToolTipContext.ShowTitleAndDescription False, xtpToolTipIconNone
    .ToolTipContext.Style = xtpToolTipResource
    .PaintManager.AutoResizeIcons = True 'WICHTIG!
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
    .KeyBindings.Add 0, VK_F12, KY_F12
    .KeyBindings.Add FCONTROL, VK_DELETE, KY_DEL
    .KeyBindings.Add FCONTROL, Asc("F"), RibCon_Formulare
    .KeyBindings.Add FCONTROL, Asc("O"), RibCon_Optionen
    .KeyBindings.Add FCONTROL, Asc("P"), KY_F10
    .KeyBindings.Add FCONTROL, Asc("N"), KY_F3
    .KeyBindings.Add FCONTROL, Asc("A"), KY_CT_A
    .KeyBindings.Add FCONTROL, Asc("M"), KY_CT_M
    .KeyBindings.Add FCONTROL, Asc("0"), KY_CT_0
    .KeyBindings.Add FCONTROL, Asc("1"), KY_CT_1
    .KeyBindings.Add FCONTROL, Asc("2"), KY_CT_2
    .KeyBindings.Add FCONTROL, Asc("3"), KY_CT_3
    .KeyBindings.Add FCONTROL, Asc("4"), KY_CT_4
    .KeyBindings.Add FCONTROL, Asc("5"), KY_CT_5
    .KeyBindings.Add FCONTROL, Asc("6"), KY_CT_6
    .KeyBindings.Add FCONTROL, Asc("7"), KY_CT_7
    .KeyBindings.Add FCONTROL, Asc("8"), KY_CT_8
    .KeyBindings.Add FCONTROL, Asc("9"), KY_CT_9
    .KeyBindings.Add FCONTROL, Asc("Z"), KY_CT_Z
    .KeyBindings.Add FCONTROL, Asc("Y"), KY_CT_Y
    .KeyBindings.Add FCONTROL + FALT, Asc("D"), KY_CT_AL_D
    .KeyBindings.Add FCONTROL + FALT, Asc("L"), KY_CT_AL_L
    Select Case GlSkn
    Case 1: .PaintManager.LoadFrameIcon App.hInstance, App.Path & "\Skins\doctor1.ico", 16, 16
    Case 2: .PaintManager.LoadFrameIcon App.hInstance, App.Path & "\Skins\doctor2.ico", 16, 16
    Case 3: .PaintManager.LoadFrameIcon App.hInstance, App.Path & "\Skins\doctor3.ico", 16, 16
    End Select
End With

Set RbBar = CmBrs.Item(1)
With RbBar
    .AllowMinimize = False
    .AllowQuickAccessCustomization = False
    .AllowQuickAccessDuplicates = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .EnableAnimation = GlMeA
    .FontHeight = GlToF 'WICHTIG!
    .GroupsVisible = True
    .MinimumVisibleWidth = 100
    .RibbonPaintManager.HotTrackingGroups = True
    .RibbonPaintManager.CaptionFont.SIZE = 8
    .RibbonPaintManager.CaptionFont.Name = GlTFt.Name
    .RibbonPaintManager.WindowCaptionFont.SIZE = 8
    .RibbonPaintManager.WindowCaptionFont.Name = GlTFt.Name
    .ShowQuickAccess = False
    .ShowQuickAccessBelowRibbon = False
    .ShowCaptionAlways = True
    .Position = xtpBarTop
    .SetIconSize 16, 16
    Select Case GlSty '\\\
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'farbige Register
            .TabPaintManager.Appearance = xtpTabAppearanceVisualStudio2005
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 6
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
        Else
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        End If
    End Select
    .TabPaintManager.Layout = xtpTabLayoutAutoSize
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.MinTabWidth = 100
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ClientFrame = xtpTabFrameNone
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = False
    .TabPaintManager.HotTracking = True
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.Font.SIZE = 8
    .TabPaintManager.Font.Name = GlTFt.Name
    If GlRDP = True Then
        .EnableFrameTheme
    Else
        If GlRah = True Then
            .EnableFrameTheme
        End If
    End If
End With

Set CmSys = RbBar.SystemButton
With CmSys
    .IconId = IC32_Doctor_Norm
    .Caption = "System"
    If GlBty = True Then
        .Style = xtpButtonAutomatic
    Else
        .Style = xtpButtonCaption
    End If
End With

Set CmSys = Nothing
Set RbBar = Nothing
Set CmAcs = Nothing
Set CmOpt = Nothing
Set CmSta = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub
LaErr:
If GlDbg = True Then SErLog Err.Description & " SMePa " & Err.Number
Resume Next

End Sub
Public Function SMFUp(ByVal FiNam As String, Optional ByVal MaEma As String, Optional ByVal MaBrf As String, Optional ByVal BogNa As String) As String
On Error GoTo SuErr
'Veröffentlicht (Publiziert) das Neuaufnahmeformular

Dim MitNr As Long
Dim ManNr As Long
Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim SuGui As String
Dim GuiKy As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim IniNa As String
Dim DaIni As String
Dim MaNam As String
Dim DaNaO As String
Dim TmpSt As String
Dim TmZei As String
Dim FmLnk As String
Dim FrmID As String
Dim ErrSt As String
Dim FrUpl As Integer
Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim RetWe As Boolean
Dim AryZe() As String

Const FrSet = 0
Const FrTer = 1
Const FrRed = 0
Const FrRec = 0

If GlFUp = True Then
    FrUpl = 1
End If

Set FM = frmMain

If GlEKV = False Then 'Emailkonten vorhanden
    TeTit = "E-Mail-Versand"
    TeMai = "Es ist kein E-Mail-Konto vorhanden"
    TeInh = "Um eine E-Mail-Bestätigung für ein ausgefüllten Aufnahmeformular zu erhalten, ist es notwendig mind. ein E-Mail-Konto hinzuzufügen."
    TeFus = "Um ein E-Mail-Konto hinzuzufügen, wechseln Sie in das Modul: Textverarbeitung und dann oben auf Emails. Dort klicken Sie auf die Schaltfläche Emailkonten."
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, True, FM.hwnd
    Exit Function
End If

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

PrNam = App.Path & "\smforms.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smforms.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

Screen.MousePointer = vbHourglass
DoEvents

MitNr = GlMiA(GlSmI, 2)

If GlCID <> vbNullString Then 'Cloud-ID
    GuiKy = "F" & Right$(GlCID, Len(GlCID) - 1)
Else
    GuiKy = CreateID("F")
End If

For AktZa = 1 To UBound(GlThe) 'Mandanten
    If ManNr = GlThe(AktZa, 0) Then
        MaNam = GlThe(AktZa, 13)
        Exit For
    End If
Next AktZa

TeTit = "Neuaufnahmeformular Publizieren"
TeMai = "Soll das Neuaufnahmeformular jetzt publiziert werden?"
TeInh = "Durch das Publizieren eines Neuaufnahmeformulars im Internet ist es möglich, Neuanmeldungen zu automatisieren."
TeFus = "Wurde dieses Anmeldeformular von einem Patienten bearbeitet, wird eine Benachrichtigung an: " & MaEma & " gesendet. Die Ergebnisse können dann abgerufen und automatisch verwertet werden."

If Len(MaBrf) > 200 Then
    MaBrf = Left$(MaBrf, 200)
End If

If FiNam <> vbNullString Then
    With clFil
        .FilPfa FiNam
        DaNaO = .DaNaO
    End With
    If GlTxK <> vbNullString Then
        DaNaO = GlTxK
    End If
    If Left$(DaNaO, 2) = "TD" Then
        Lange = Len(DaNaO)
        DaNaO = Right$(DaNaO, Lange - 43)
    End If
End If

If GlCID <> vbNullString Then 'Cloud-ID
    PrNam = Chr$(34) & PrNam & Chr$(34)
    IniNa = CreateID("U") & ".ini"
    DaIni = GlTmp & IniNa
    
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
    If GlMes = 33565 Then

        If FiNam <> vbNullString Then
            If GlOtL <> vbNullString Then 'Online-Terminbuchungs System Link für Datenschutzerklärung
                If GlOIm <> vbNullString Then 'Online-Terminbuchungs System Link für Impressum
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & FiNam & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & DaNaO & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & FiNam & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & DaNaO & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            Else
                If GlOIm <> vbNullString Then
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & FiNam & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & DaNaO & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--pdf=" & Chr$(34) & FiNam & Chr$(34) & Space$(1) & "--signTitle=" & Chr$(34) & DaNaO & Chr$(34) & Space$(1) & "--signDesc=" & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            End If
        Else
            If GlOtL <> vbNullString Then 'Online-Terminbuchungs System Link für Datenschutzerklärung
                If GlOIm <> vbNullString Then 'Online-Terminbuchungs System Link für Impressum
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--privacyUri=" & Chr$(34) & GlOtL & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            Else
                If GlOIm <> vbNullString Then
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--termsUri=" & Chr$(34) & GlOIm & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                Else
                    PaStr = "upload" & Space$(1) & Chr$(34) & GlCID & Chr$(34) & Space$(1) & Chr$(34) & MaBrf & Chr$(34) & Space$(1) & Chr$(34) & MaEma & Chr$(34) & Space$(1) & Chr$(34) & GuiKy & Chr$(34) & Space$(1) & Chr$(34) & BogNa & Chr$(34) & Space$(1) & Chr$(34) & Chr$(34) & Space$(1) & FrSet & Space$(1) & FrTer & Space$(1) & FrUpl & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--reducedMandatory=" & FrRed & Space$(1) & "--extendedForm=" & FrRec
                End If
            End If
        End If

        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If TmpSt <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "formurl": FrmID = Right$(TmZei, (Lange - Posit) - 1)
                                Case "completeformurl": FmLnk = Right$(TmZei, Lange - Posit)
                                Case "error": ErrSt = Right$(TmZei, Lange - Posit)
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
            DoEvents
            
            If FmLnk <> vbNullString Then
                SMFUp = FmLnk
            End If
            
            If ErrSt = vbNullString Then
                If GlLog = False Then 'General Logging
                    .DaLoe = GlTmp & "*.ini" & vbNullChar
                    .FilLoe
                    If FiNam <> vbNullString Then
                        If .FilVor(FiNam) = True Then
                            .DaLoe = FiNam & vbNullChar
                            .FilLoe
                        End If
                    End If
                Else
                    Clipboard.Clear
                    Clipboard.SetText PrNam & Space$(1) & PaStr
                End If
            End If
        End With
        DoEvents

        If FmLnk <> vbNullString Then
            frmTSEInit.Show
            frmTSEInit.Caption = "Neuaufnahme"
            frmTSEInit.txtTSEIn.Text = vbCrLf & "Die URL zu Ihrem Neuaufnahmeformular lautet:" & vbCrLf & vbCrLf & FmLnk & vbCrLf & vbCrLf & "Dieser Weblink befindet sich jetzt in Ihrer Zwischenablage, so dass dieser an anderer Stelle wieder eingefügt und verwendet werden kann."
            GlSet(1, 82) = FmLnk
            S_SeSe 83, FmLnk
            GlNaf = FmLnk
            DoEvents
            Clipboard.Clear
            If GlLog = False Then 'General Logging
                Clipboard.SetText FmLnk
            Else
                Clipboard.SetText PrNam & Space$(1) & PaStr
            End If
        Else
            Clipboard.Clear
            Clipboard.SetText PrNam & Space$(1) & PaStr
            If ErrSt = vbNullString Then
                SPopu "Uploadfehler", "Unerwarteter Fehler, beim Hochladen des Neuaufnahmeformulars", IC48_Forbidden
            Else
                SPopu "Uploadfehler", ErrSt, IC48_Forbidden
            End If
        End If
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Function

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SMFUp " & Err.Number
Resume Next

End Function

Public Function SMSGu() As String
On Error GoTo InErr
'SMS Versand

Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim DaIni As String
Dim IniNa As String
Dim TmpSt As String
Dim TmZei As String
Dim Gutha As String
Dim Warun As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim StaOk As Boolean
Dim AryZe() As String
Dim Mld1, Tit1 As String

Set clFil = New clsFile

PrNam = App.Path & "\smcm.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smcm.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

If GlTok = vbNullString Then
    TeTit = "Fehlender Produkttoken"
    TeMai = "Es wurde kein SMS-Produkttoken im eingetragen!"
    TeInh = "Um das Guthaben zum Versenden von SMS zu erfragen, ist es notwendig Ihren CM Telekom Produkt Token in den Einstellungendialog einzutragen."
    TeFus = "Bitte rufen die die Website von CM Telekom auf und entnehmen Sie dort den Produkt Token der Messaging API. Weitere Informationen hierzu entlehnen Sie bitte der Anleitung."
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Function
End If

If GlTok = vbNullString Then
    TeTit = "Fehlender Account ID"
    TeMai = "Es wurde keine SMS Account ID im eingetragen!"
    TeInh = "Um das Guthaben zum Versenden von SMS zu erfragen, ist es notwendig Ihre CM Telekom Account ID in den Einstellungendialog einzutragen."
    TeFus = "Bitte rufen die die Website von CM Telekom auf und entnehmen Sie die Account ID aus der URL der Webiste. Weitere Informationen hierzu entlehnen Sie bitte der Anleitung."
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Function
End If

PrNam = Chr$(34) & PrNam & Chr$(34)
IniNa = CreateID("S") & ".ini"
DaIni = GlTmp & IniNa

PaStr = "checkbalance" & Space$(1) & Chr$(34) & GlTok & Chr$(34) & Space$(1) & Chr$(34) & GlAcI & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34)
WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
DoEvents

With clFil
    If .FilVor(DaIni) = True Then
        .FilPfa DaIni
        TmpSt = .FilReSt
        DoEvents
        If TmpSt <> vbNullString Then
            AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
            For AktZe = 0 To UBound(AryZe) - 1
                If AryZe(AktZe) <> vbNullString Then
                    TmZei = AryZe(AktZe)
                    Lange = Len(TmZei)
                    Posit = InStr(1, TmZei, "=", 1)
                    If Posit > 0 Then
                        InTyp = LCase(Left$(TmZei, Posit - 1))
                        Select Case InTyp
                        Case "amount": Gutha = Right$(TmZei, Lange - Posit)
                        Case "currency": Warun = Right$(TmZei, Lange - Posit)
                        End Select
                    End If
                End If
            Next AktZe
        End If
    End If
    DoEvents
    
    If LCase(Gutha) <> vbNullString Then
        SMSGu = Warun & " " & Format$(Gutha, GlWa1)
    End If

    If GlLog = False Then 'General Logging
        With clFil
            .DaLoe = GlTmp & "*.ini" & vbNullChar
            .FilLoe
        End With
    Else
        Clipboard.Clear
        Clipboard.SetText PrNam & Space$(1) & PaStr
    End If
End With
        
Set clFil = Nothing

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SMSGu " & Err.Number
Resume Next

End Function

Public Function SMSSn(ByVal TelNr As String, ByVal NaTex As String) As Boolean
On Error GoTo InErr
'SMS Versand

Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim DaIni As String
Dim IniNa As String
Dim TmpSt As String
Dim TmZei As String
Dim StaCo As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim AktZa As Integer
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim StaOk As Boolean
Dim AryZe() As String

Set clFil = New clsFile

PrNam = App.Path & "\smcm.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smcm.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

If GlTok = vbNullString Then
    TeTit = "Fehlender Produkttoken"
    TeMai = "Es wurde kein SMS-Produkttoken im eingetragen!"
    TeInh = "Um das Guthaben zum Versenden von SMS zu erfragen, ist es notwendig Ihren CM Telekom Produkt Token in den Einstellungendialog einzutragen."
    TeFus = "Bitte rufen die die Website von CM Telekom auf und entnehmen Sie dort den Produkt Token der Messaging API. Weitere Informationen hierzu entlehnen Sie bitte der Anleitung."
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Function
End If

For AktZa = 1 To 20 'max. 20 x vorkommen
    NaTex = Replace(NaTex, vbCrLf, "$$", 1)
Next AktZa

If Len(NaTex) > 160 Then
    NaTex = Left$(NaTex, 160)
End If

PrNam = Chr$(34) & PrNam & Chr$(34)
IniNa = CreateID("S") & ".ini"
DaIni = GlTmp & IniNa

PaStr = "sendmessage" & Space$(1) & Chr$(34) & GlTok & Chr$(34) & Space$(1) & Chr$(34) & "SMS" & Chr$(34) & Space$(1) & Chr$(34) & GlAbs & Chr$(34) & Space$(1) & Chr$(34) & TelNr & Chr$(34) & Space$(1) & Chr$(34) & NaTex & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34)
WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
DoEvents

With clFil
    If .FilVor(DaIni) = True Then
        .FilPfa DaIni
        TmpSt = .FilReSt
        DoEvents
        If TmpSt <> vbNullString Then
            AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
            For AktZe = 0 To UBound(AryZe) - 1
                If AryZe(AktZe) <> vbNullString Then
                    TmZei = AryZe(AktZe)
                    Lange = Len(TmZei)
                    Posit = InStr(1, TmZei, "=", 1)
                    If Posit > 0 Then
                        InTyp = LCase(Left$(TmZei, Posit - 1))
                        Select Case InTyp
                        Case "statuscode": StaCo = Right$(TmZei, Lange - Posit)
                        End Select
                    End If
                End If
            Next AktZe
        End If
    End If
    DoEvents
    
    If LCase(StaCo) = "ok" Then
        SMSSn = True
    End If
    
    If GlLog = False Then 'General Logging
        With clFil
            .DaLoe = GlTmp & "*.ini" & vbNullChar
            .FilLoe
        End With
    Else
        Clipboard.Clear
        Clipboard.SetText PrNam & Space$(1) & PaStr
    End If
End With
        
Set clFil = Nothing

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SMSSn " & Err.Number
Resume Next

End Function
Public Function SMSTe(ByVal TelNr As String) As String
On Error GoTo InErr
'Testen des SMS Rufnummernformates

Dim InTyp As String
Dim PaStr As String
Dim PrNam As String
Dim DaIni As String
Dim IniNa As String
Dim TmpSt As String
Dim TmZei As String
Dim TeSta As String
Dim TeGes As String
Dim TeFrm As String
Dim TeReg As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim StaOk As Boolean
Dim AryZe() As String

Set clFil = New clsFile

PrNam = App.Path & "\smcm.exe"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smcm.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

If GlTok = vbNullString Then
    TeTit = "Fehlender Produkttoken"
    TeMai = "Es wurde kein SMS-Produkttoken im eingetragen!"
    TeInh = "Um das Guthaben zum Versenden von SMS zu erfragen, ist es notwendig Ihren CM Telekom Produkt Token in den Einstellungendialog einzutragen."
    TeFus = "Bitte rufen die die Website von CM Telekom auf und entnehmen Sie dort den Produkt Token der Messaging API. Weitere Informationen hierzu entlehnen Sie bitte der Anleitung."
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
    Exit Function
End If

PrNam = Chr$(34) & PrNam & Chr$(34)
IniNa = CreateID("S") & ".ini"
DaIni = GlTmp & IniNa

PaStr = "checkphone" & Space$(1) & Chr$(34) & GlTok & Chr$(34) & Space$(1) & Chr$(34) & TelNr & Chr$(34) & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34)
WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
DoEvents

With clFil
    If .FilVor(DaIni) = True Then
        .FilPfa DaIni
        TmpSt = .FilReSt
        DoEvents
        If TmpSt <> vbNullString Then
            AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
            For AktZe = 0 To UBound(AryZe) - 1
                If AryZe(AktZe) <> vbNullString Then
                    TmZei = AryZe(AktZe)
                    Lange = Len(TmZei)
                    Posit = InStr(1, TmZei, "=", 1)
                    If Posit > 0 Then
                        InTyp = LCase(Left$(TmZei, Posit - 1))
                        Select Case InTyp
                        Case "formatinternational": TeFrm = Right$(TmZei, Lange - Posit)
                        Case "isvalidnumber": TeSta = Right$(TmZei, Lange - Posit)
                        Case "carrier": TeGes = Right$(TmZei, Lange - Posit)
                        Case "region": TeReg = Right$(TmZei, Lange - Posit)
                        End Select
                    End If
                End If
            Next AktZe
        End If
    End If
    DoEvents
    
    If LCase(TeSta) = "true" Then
        SMSTe = TeFrm & " (" & TeGes & ")" & " " & TeReg
    End If
    
    If GlLog = False Then 'General Logging
        With clFil
            .DaLoe = GlTmp & "*.ini" & vbNullChar
            .FilLoe
        End With
    Else
        Clipboard.Clear
        Clipboard.SetText PrNam & Space$(1) & PaStr
    End If
End With
        
Set clFil = Nothing

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SMSTe " & Err.Number
Resume Next

End Function

Public Sub SSpLa()
On Error GoTo PoErr

Select Case GlBut
Case RibTab_Fragebogen: SSpLaB
Case RibTab_Abrechnung: SSpLa7
Case RibTab_Vorbereit: SSpLa7b
Case RibTab_Tagesproto: SSpLa7
Case RibTab_Mahnwesen: SSpLa3
Case RibTab_Buchungen: SSpLa4
Case RibTab_HomeBanki: SSpLa4a
Case RibTab_Rezeptmodul: SSpLa6
Case RibTab_Belegmodul: SSpLa6
Case RibTab_Ter_Listen: SSpLa9
Case RibTab_Ter_Akont: SSpLa9
Case RibTab_Ter_Warte: SSpLa9
                       SSpLa9a
Case RibTab_LabBericht: SSpLa5a
                        SSpLa5b
Case RibTab_LabAuftrag: SSpLa5c
                        SSpLa5d
Case RibTab_LabBerichte: SSpLa5e
Case RibTab_LabAuftrage: SSpLa5f
End Select

If GlSta = False Then
    Select Case GlBut
    Case RibTab_Fragebogen: S_AnSpl
    Case RibTab_Mahnwesen: S_OPSpl
    Case RibTab_Buchungen: S_BuSpl
    Case RibTab_HomeBanki: S_BaSpl
    Case RibTab_Rezeptmodul: S_RzSpl
    Case RibTab_Belegmodul: S_RzSpl
    Case RibTab_Ter_Listen: S_TeSpl
    Case RibTab_Ter_Akont: S_TeSpl
    Case RibTab_Ter_Warte: S_TeSpl
                           S_WaSpl
    Case RibTab_LabBericht: S_LbSp1
    Case RibTab_LabBerichte: S_LbSp2
    End Select
End If

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " SSpLa " & Err.Number
Resume Next

End Sub
Public Sub SSpLa1()
On Error GoTo SpErr

Dim SpStr As String
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo2 = FM.repCont2
Set RpCls = RpCo2.Columns

With RpCls
    Set RpCol = .Add(Adr_ID0, "ID0", 0, False)
    Set RpCol = .Add(Adr_ID3, "ID3", 0, False)
    Set RpCol = .Add(Adr_IDKurz, "Suchbegriff", 0, True)
    If RpCo2.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Adr_Geboren, "Geboren", 0, True)
    Set RpCol = .Add(Adr_Name, "Name", 0, True)
    Set RpCol = .Add(Adr_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Adr_Straße, "Straße", 0, True)
    Set RpCol = .Add(Adr_PLZ, "PLZ", 0, True)
    Set RpCol = .Add(Adr_Ort, "Ort", 0, True)
    Set RpCol = .Add(Adr_Firma1, "Firma/Info", 0, True)
    Set RpCol = .Add(Adr_Telefon1, "Privat", 0, True)
    Set RpCol = .Add(Adr_Telefon2, "Büro", 0, True)
    Set RpCol = .Add(Adr_Telefon3, "Telefax", 0, True)
    Set RpCol = .Add(Adr_Telefon4, "Mobil", 0, True)
    Set RpCol = .Add(Adr_Telefon5, "Email", 0, True)
    Set RpCol = .Add(Adr_Geschlecht, "Geschlecht", 0, True)
    Set RpCol = .Add(Adr_Datum, "Datun", 0, False)
    Set RpCol = .Add(Adr_Briefanrede, "Briefanrede", 0, False)
    Set RpCol = .Add(Adr_Anschrift, "Anschrift", 0, False)
    Set RpCol = .Add(Adr_TreKey, "TreKey", 0, False)
    Set RpCol = .Add(Adr_Grafik, "Grafik", 0, False)
    Set RpCol = .Add(Adr_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Adr_Objekt, "Objekt", 0, False)
    Set RpCol = .Add(Adr_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Adr_Mandant, "PIN", 0, True)
    Set RpCol = .Add(Adr_VIP, "VIP", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Adr_Titel, "Titel", 0, False)
    Set RpCol = .Add(Adr_Land, "Land", 0, False)
    Set RpCol = .Add(Adr_Behindert, "Behindert", 0, False)
    Set RpCol = .Add(Adr_Passiv, "Passiv", 0, False)
    Set RpCol = .Add(Adr_Gruppen, "Gruppen", 0, True)
    Set RpCol = .Add(Adr_Versand, "V", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
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

If GlIdi = True Then 'Idiotenmodus
    RpCls(Adr_ID0).Width = 0
    RpCls(Adr_ID3).Width = 0
    RpCls(Adr_IDKurz).Width = 220
    If GlTFt.SIZE > 10 Then
        RpCls(Adr_Geboren).Width = 110
    Else
        RpCls(Adr_Geboren).Width = 80
    End If
    RpCls(Adr_Name).Width = 100
    RpCls(Adr_Vorname).Width = 100
    RpCls(Adr_Straße).Width = 120
    RpCls(Adr_PLZ).Width = 60
    RpCls(Adr_Ort).Width = 100
    RpCls(Adr_Firma1).Width = 150
    RpCls(Adr_Telefon1).Width = 90
    RpCls(Adr_Telefon2).Width = 90
    RpCls(Adr_Telefon3).Width = 90
    RpCls(Adr_Telefon4).Width = 90
    RpCls(Adr_Telefon5).Width = 120
    RpCls(Adr_Geschlecht).Width = 80
    RpCls(Adr_Datum).Width = 0
    RpCls(Adr_Briefanrede).Width = 0
    RpCls(Adr_Anschrift).Width = 0
    RpCls(Adr_TreKey).Width = 0
    RpCls(Adr_Grafik).Width = 0
    RpCls(Adr_GuiID).Width = 0
    RpCls(Adr_Objekt).Width = 0
    RpCls(Adr_IDP).Width = 0
    RpCls(Adr_Mandant).Width = 50
    RpCls(Adr_VIP).Width = 0
    RpCls(Adr_Titel).Width = 0
    RpCls(Adr_Land).Width = 0
    RpCls(Adr_Behindert).Width = 0
    RpCls(Adr_Passiv).Width = 0
    RpCls(Adr_Gruppen).Width = 150
    RpCls(Adr_Versand).Width = 20
Else
    If IniGetSek(GlINI, "RpCnt1a") = False Then
        RpCls(Adr_ID0).Width = 0
        RpCls(Adr_ID3).Width = 0
        RpCls(Adr_IDKurz).Width = 220
        If GlTFt.SIZE > 10 Then
            RpCls(Adr_Geboren).Width = 110
        Else
            RpCls(Adr_Geboren).Width = 80
        End If
        RpCls(Adr_Name).Width = 100
        RpCls(Adr_Vorname).Width = 100
        RpCls(Adr_Straße).Width = 120
        RpCls(Adr_PLZ).Width = 60
        RpCls(Adr_Ort).Width = 100
        RpCls(Adr_Firma1).Width = 150
        RpCls(Adr_Telefon1).Width = 90
        RpCls(Adr_Telefon2).Width = 90
        RpCls(Adr_Telefon3).Width = 90
        RpCls(Adr_Telefon4).Width = 90
        RpCls(Adr_Telefon5).Width = 120
        RpCls(Adr_Geschlecht).Width = 80
        RpCls(Adr_Datum).Width = 0
        RpCls(Adr_Briefanrede).Width = 0
        RpCls(Adr_Anschrift).Width = 0
        RpCls(Adr_TreKey).Width = 0
        RpCls(Adr_Grafik).Width = 0
        RpCls(Adr_GuiID).Width = 0
        RpCls(Adr_Objekt).Width = 0
        RpCls(Adr_IDP).Width = 0
        RpCls(Adr_Mandant).Width = 50
        RpCls(Adr_VIP).Width = 0
        RpCls(Adr_Titel).Width = 0
        RpCls(Adr_Land).Width = 0
        RpCls(Adr_Behindert).Width = 0
        RpCls(Adr_Passiv).Width = 0
        RpCls(Adr_Gruppen).Width = 150
        RpCls(Adr_Versand).Width = 20
        IniSetSek "RpCnt1a"
        IniSetVal "RpCnt1a", "SSpLa1", RpCo2.SaveSettings
    Else
        SpStr = IniGetBig(GlINI, "RpCnt1a", "SSpLa1")
        RpCo2.LoadSettings SpStr
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo2 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLa1 " & Err.Number
Resume Next

End Sub
Public Sub SSpLa2()
On Error GoTo SpErr

Dim SpStr As String
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo4 = FM.repCont4
Set RpCls = RpCo4.Columns

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

If GlIdi = True Then 'Idiotenmodus
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
    RpCls(Rec_Versand).Width = 20
    RpCls(Rec_Betrag).Width = 75
    RpCls(Rec_Bezahlt).Width = 75
    RpCls(Rec_Differe).Width = 75
    RpCls(Rec_IDKurz).Width = 220
    RpCls(Rec_Offen).Width = 0
    RpCls(Rec_Extrageb).Width = 75
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
    RpCls(Rec_IDP).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_Druckdatum).Width = 110
    Else
        RpCls(Rec_Druckdatum).Width = 80
    End If
    RpCls(Rec_Kopie).Width = 0
    RpCls(Rec_Steuer).Width = 60
    RpCls(Rec_Monat).Width = 0
    RpCls(Rec_Termin).Width = 75
    RpCls(Rec_Storniert).Width = 0
    RpCls(Rec_PKU).Width = 50
    RpCls(Rec_Beendet).Width = 0
    RpCls(Rec_Rabatt).Width = 0
    RpCls(Rec_IDM).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_GuStr).Width = 110
    Else
        RpCls(Rec_GuStr).Width = 80
    End If
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
Else
    If IniGetSek(GlINI, "SplRp2") = False Then
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
        RpCls(Rec_Versand).Width = 20
        RpCls(Rec_Betrag).Width = 75
        RpCls(Rec_Bezahlt).Width = 75
        RpCls(Rec_Differe).Width = 75
        RpCls(Rec_IDKurz).Width = 220
        RpCls(Rec_Offen).Width = 0
        RpCls(Rec_Extrageb).Width = 75
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
        RpCls(Rec_IDP).Width = 180
        If GlTFt.SIZE > 10 Then
            RpCls(Rec_Druckdatum).Width = 110
        Else
            RpCls(Rec_Druckdatum).Width = 80
        End If
        RpCls(Rec_Kopie).Width = 0
        RpCls(Rec_Steuer).Width = 60
        RpCls(Rec_Monat).Width = 0
        RpCls(Rec_Termin).Width = 75
        RpCls(Rec_Storniert).Width = 0
        RpCls(Rec_PKU).Width = 50
        RpCls(Rec_Beendet).Width = 0
        RpCls(Rec_Rabatt).Width = 0
        RpCls(Rec_IDM).Width = 180
        If GlTFt.SIZE > 10 Then
            RpCls(Rec_GuStr).Width = 110
        Else
            RpCls(Rec_GuStr).Width = 80
        End If
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
        IniSetSek "SplRp2"
        IniSetVal "SplRp2", "SSpLa2", RpCo4.SaveSettings
    Else
        SpStr = IniGetBig(GlINI, "SplRp2", "SSpLa2")
        RpCo4.LoadSettings SpStr
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo4 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLa2 " & Err.Number
Resume Next

End Sub
Private Sub SSpLa3()
On Error GoTo SpErr

Dim SpStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

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
    If RpCo1.PaintManager.FixedRowHeight = False Then
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
    Set RpCol = .Add(OPo_Gebuehr, "Gebühr", 0, False)
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
    Set RpCol = .Add(OPo_Berichtdatum, "Bericht", 0, True)
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

If GlIdi = True Then 'Idiotenmodus
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
        RpCls(OPo_IDP).Width = 180
        RpCls(OPo_Berichtdatum).Width = 110
        RpCls(OPo_Steuer).Width = 100
        RpCls(OPo_Mahnung1).Width = 110
        RpCls(OPo_Mahnung2).Width = 110
        RpCls(OPo_Mahnung3).Width = 110
        RpCls(OPo_Mahnung4).Width = 110
        RpCls(OPo_Mahnung5).Width = 110
        RpCls(OPo_IDT).Width = 180
        RpCls(OPo_Versand).Width = 20
    Else
        RpCls(OPo_RechNr).Width = 110
        RpCls(OPo_OffBetrag).Width = 75
        RpCls(OPo_Stufe).Width = 30
        RpCls(OPo_Patient).Width = 220
        RpCls(OPo_ReBetrag).Width = 75
        RpCls(OPo_Bezahlt).Width = 75
        RpCls(OPo_Gebuehr).Width = 75
        RpCls(OPo_W).Width = 30
        RpCls(OPo_Datum).Width = 80
        RpCls(OPo_Fällig).Width = 80
        RpCls(OPo_Einzahlung).Width = 80
        RpCls(OPo_Mahnfrist).Width = 80
        RpCls(OPo_IDP).Width = 180
        RpCls(OPo_Berichtdatum).Width = 80
        RpCls(OPo_Steuer).Width = 75
        RpCls(OPo_Mahnung1).Width = 80
        RpCls(OPo_Mahnung2).Width = 80
        RpCls(OPo_Mahnung3).Width = 80
        RpCls(OPo_Mahnung4).Width = 80
        RpCls(OPo_Mahnung5).Width = 80
        RpCls(OPo_IDT).Width = 180
        RpCls(OPo_Versand).Width = 20
    End If
Else
    If IniGetSek(GlINI, "RpCnt3c") = False Then
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
            RpCls(OPo_IDP).Width = 180
            RpCls(OPo_Berichtdatum).Width = 110
            RpCls(OPo_Steuer).Width = 100
            RpCls(OPo_Mahnung1).Width = 110
            RpCls(OPo_Mahnung2).Width = 110
            RpCls(OPo_Mahnung3).Width = 110
            RpCls(OPo_Mahnung4).Width = 110
            RpCls(OPo_Mahnung5).Width = 110
            RpCls(OPo_IDT).Width = 180
            RpCls(OPo_Versand).Width = 20
        Else
            RpCls(OPo_RechNr).Width = 110
            RpCls(OPo_OffBetrag).Width = 75
            RpCls(OPo_Stufe).Width = 30
            RpCls(OPo_Patient).Width = 220
            RpCls(OPo_ReBetrag).Width = 75
            RpCls(OPo_Bezahlt).Width = 75
            RpCls(OPo_Gebuehr).Width = 75
            RpCls(OPo_W).Width = 30
            RpCls(OPo_Datum).Width = 80
            RpCls(OPo_Fällig).Width = 80
            RpCls(OPo_Einzahlung).Width = 80
            RpCls(OPo_Mahnfrist).Width = 80
            RpCls(OPo_IDP).Width = 180
            RpCls(OPo_Berichtdatum).Width = 80
            RpCls(OPo_Steuer).Width = 75
            RpCls(OPo_Mahnung1).Width = 80
            RpCls(OPo_Mahnung2).Width = 80
            RpCls(OPo_Mahnung3).Width = 80
            RpCls(OPo_Mahnung4).Width = 80
            RpCls(OPo_Mahnung5).Width = 80
            RpCls(OPo_IDT).Width = 180
            RpCls(OPo_Versand).Width = 20
        End If
        IniSetSek "RpCnt3c"
        IniSetVal "RpCnt3c", "SSpLa3", RpCo1.SaveSettings
    Else
        SpStr = IniGetBig(GlINI, "RpCnt3c", "SSpLa3")
        RpCo1.LoadSettings SpStr
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa3 " & Err.Number
Resume Next

End Sub
Private Sub SSpLa4()
On Error GoTo SpErr
'Formratieren der Spalten

Dim SpStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Buh_ID0, "ID0", 0, False)
    Set RpCol = .Add(Buh_Datum, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Buh_Buchtext, "Buchungstext", 0, True)
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        Set RpCol = .Add(Buh_Einnahme, "Einnahme", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Ausgabe, "Ausgabe", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Sachkonto, "Sachkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Geldkonto", 0, True)
    Else
        Set RpCol = .Add(Buh_Einnahme, "Betrag", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Ausgabe, "Brutto", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Sachkonto, "Sollkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Habenkonto", 0, True)
    End If
    Set RpCol = .Add(Buh_RechNr, "Belegzeichen", 0, True)
    Set RpCol = .Add(Buh_IDR, "IDR", 0, False)
    Set RpCol = .Add(Buh_Beleg, "Nummer", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Buh_Sachkontenbez, "Sachkontenbezeichnung", 0, True)
    Set RpCol = .Add(Buh_Geldkontenbez, "Geldkontenbezeichnung", 0, True)
    Set RpCol = .Add(Buh_Steuer, "Steuer", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Buh_W, "W", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Buh_Privat, "Privat", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Abziehbar, "Abziehbar", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDB, "IDB", 0, False)
    Set RpCol = .Add(Buh_IDA, "IDA", 0, False)
    Set RpCol = .Add(Buh_Währung, "Währung", 0, False)
    Set RpCol = .Add(Buh_Ermittlung, "KE", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Dokument, "DK", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_Paperclip
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_IDP, "IDP", 0, False)
    Set RpCol = .Add(Buh_IDArt, "IDArt", 0, False)
    Set RpCol = .Add(Buh_IDBank, "IDBank", 0, False)
    Set RpCol = .Add(Buh_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Buh_IDT, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Buh_Berichtdatum, "Bericht", 0, True)
    Set RpCol = .Add(Buh_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Buh_Monat, "Monat", 0, False)
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
    Set RpCol = .Add(Buh_Storniert, "Storniert", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Buh_Zuordnung, "ZU", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_User_Norm
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Lock, "Lock", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconLeft
        .Icon = IC16_Lock
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Datei, "Datei", 0, False)
    Set RpCol = .Add(Buh_Doppelt, "DO", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_TSELog, "TSELog", 0, False)
    Set RpCol = .Add(Buh_TSEZahl, "TSEZahl", 0, False)
    Set RpCol = .Add(Buh_TSESign, "TSESign", 0, False)
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
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 140
    RpCls(Buh_Buchtext).Width = 250
    RpCls(Buh_Einnahme).Width = 100
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        RpCls(Buh_Ausgabe).Width = 100
    Else
        RpCls(Buh_Ausgabe).Width = 0
    End If
Else
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 110
    RpCls(Buh_Buchtext).Width = 220
    RpCls(Buh_Einnahme).Width = 80
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        RpCls(Buh_Ausgabe).Width = 80
    Else
        RpCls(Buh_Ausgabe).Width = 0
    End If
End If
RpCls(Buh_Sachkonto).Width = 80
RpCls(Buh_Gegenkonto).Width = 80
RpCls(Buh_RechNr).Width = 90
RpCls(Buh_IDR).Width = 0
RpCls(Buh_Beleg).Width = 75
RpCls(Buh_Sachkontenbez).Width = 160
RpCls(Buh_Geldkontenbez).Width = 160
RpCls(Buh_Steuer).Width = 75
RpCls(Buh_W).Width = 40
RpCls(Buh_Privat).Width = 0
RpCls(Buh_Abziehbar).Width = 0
RpCls(Buh_IDB).Width = 0
RpCls(Buh_IDA).Width = 0
RpCls(Buh_Währung).Width = 0
RpCls(Buh_Ermittlung).Width = 25
RpCls(Buh_Dokument).Width = 25
RpCls(Buh_IDP).Width = 0
RpCls(Buh_IDArt).Width = 0
RpCls(Buh_IDBank).Width = 0
RpCls(Buh_Kommentar).Width = 0
RpCls(Buh_IDT).Width = 180
RpCls(Buh_Berichtdatum).Width = 80
RpCls(Buh_GuiID).Width = 0
RpCls(Buh_Monat).Width = 0
RpCls(Buh_Storniert).Width = 0
RpCls(Buh_IDM).Width = 180
RpCls(Buh_Zuordnung).Width = 18
RpCls(Buh_Lock).Width = 18
RpCls(Buh_Datei).Width = 0
RpCls(Buh_Doppelt).Width = 0
RpCls(Buh_TSELog).Width = 0
RpCls(Buh_TSEZahl).Width = 0
RpCls(Buh_TSESign).Width = 0

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa4 " & Err.Number
Resume Next

End Sub
Private Sub SSpLa4a()
On Error GoTo SpErr
'Homebanking

Dim SpStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Ban_ID2, "ID2", 0, False)
    Set RpCol = .Add(Ban_ID1, "ID1", 0, False)
    Set RpCol = .Add(Ban_IDB, "IDB", 0, False)
    Set RpCol = .Add(Ban_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Ban_IDR, "IDR", 0, False)
    Set RpCol = .Add(Ban_Datum, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Ban_IDKurz, "Umsatztext", 0, True)
    If RpCo1.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Ban_KoBetrag, "Betrag", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Ban_GeBetrag, "Offen", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_Kommentar, "Patient / Kommentar", 0, True)
    Set RpCol = .Add(Ban_Monat, "Monat", 0, False)
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
    Set RpCol = .Add(Ban_Selekt, "Zugeordnet", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Ban_Bezahlt, "Bezahlt", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Ban_Ausgabe, "Ausgabe", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Ban_Suche, "Suche", 0, False)
    Set RpCol = .Add(Ban_ID0, "ID0", 0, False)
    Set RpCol = .Add(Ban_ID12, "ID1_2", 0, False)
    Set RpCol = .Add(Ban_ID13, "ID1_3", 0, False)
    Set RpCol = .Add(Ban_ID14, "ID1_4", 0, False)
    Set RpCol = .Add(Ban_ID15, "ID1_5", 0, False)
    Set RpCol = .Add(Ban_IDR2, "IDR_2", 0, False)
    Set RpCol = .Add(Ban_IDR3, "IDR_3", 0, False)
    Set RpCol = .Add(Ban_IDR4, "IDR_4", 0, False)
    Set RpCol = .Add(Ban_IDR5, "IDR_5", 0, False)
    Set RpCol = .Add(Ban_RechNr1, "Rechnung1", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_RechNr2, "Rechnung2", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_RechNr3, "Rechnung3", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_RechNr4, "Rechnung4", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_RechNr5, "Rechnung5", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_OPBetrag1, "Offen1", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Ban_OPBetrag2, "Offen2", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Ban_OPBetrag3, "Offen3", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Ban_OPBetrag4, "Offen4", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Ban_OPBetrag5, "Offen5", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(Ban_Bank, "Geldkonto", 0, True)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_IDZ, "IDZ", 0, False)
    Set RpCol = .Add(Ban_IDI, "IDI", 0, False)
    Set RpCol = .Add(Ban_IDK, "Sachkonto", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ban_Buchtext, "Buchungstext", 0, False)
    Set RpCol = .Add(Ban_Konto, "Kontenbezeichnung", 0, False)
    Set RpCol = .Add(Ban_Steuer, "Steuer", 0, False)
    Set RpCol = .Add(Ban_Ermittlung, "KE", 0, False)
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

RpCls(Ban_ID2).Width = 0
If GlTFt.SIZE > 10 Then
    RpCls(Ban_Datum).Width = 140
    RpCls(Ban_IDKurz).Width = 510
    RpCls(Ban_KoBetrag).Width = 110
    RpCls(Ban_GeBetrag).Width = 110
    RpCls(Ban_Kommentar).Width = 230
    RpCls(Ban_RechNr1).Width = 120
    RpCls(Ban_RechNr2).Width = 120
    RpCls(Ban_RechNr3).Width = 120
    RpCls(Ban_RechNr4).Width = 120
    RpCls(Ban_RechNr5).Width = 120
    RpCls(Ban_OPBetrag1).Width = 100
    RpCls(Ban_OPBetrag2).Width = 100
    RpCls(Ban_OPBetrag3).Width = 100
    RpCls(Ban_OPBetrag4).Width = 100
    RpCls(Ban_OPBetrag5).Width = 100
    RpCls(Ban_IDK).Width = 90
    RpCls(Ban_Buchtext).Width = 280
    RpCls(Ban_Konto).Width = 200
    RpCls(Ban_Steuer).Width = 100
Else
    RpCls(Ban_Datum).Width = 110
    RpCls(Ban_IDKurz).Width = 480
    RpCls(Ban_KoBetrag).Width = 80
    RpCls(Ban_GeBetrag).Width = 80
    RpCls(Ban_Kommentar).Width = 200
    RpCls(Ban_RechNr1).Width = 90
    RpCls(Ban_RechNr2).Width = 90
    RpCls(Ban_RechNr3).Width = 90
    RpCls(Ban_RechNr4).Width = 90
    RpCls(Ban_RechNr5).Width = 90
    RpCls(Ban_OPBetrag1).Width = 75
    RpCls(Ban_OPBetrag2).Width = 75
    RpCls(Ban_OPBetrag3).Width = 75
    RpCls(Ban_OPBetrag4).Width = 75
    RpCls(Ban_OPBetrag5).Width = 75
    RpCls(Ban_IDK).Width = 80
    RpCls(Ban_Buchtext).Width = 250
    RpCls(Ban_Konto).Width = 180
    RpCls(Ban_Steuer).Width = 90
End If
RpCls(Ban_Monat).Width = 0
RpCls(Ban_Bank).Width = 180
RpCls(Ban_IDP).Width = 180
RpCls(Ban_IDM).Width = 180

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa4a " & Err.Number
Resume Next

End Sub
Private Sub SSpLa5a()
On Error GoTo SpErr
'Laborbericht

Dim SpStr As String
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns

With RpCo5
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Lab_ID0, "ID0", 0, False)
    Set RpCol = .Add(Lab_Auftrag, "Auftrag", 0, True)
    Set RpCol = .Add(Lab_F, "R", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Befundart, "Befundart", 0, True)
    Set RpCol = .Add(Lab_Patient, "Patient", 0, True)
    Set RpCol = .Add(Lab_Name, "Name", 0, True)
    Set RpCol = .Add(Lab_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Lab_Geboren, "Geboren", 0, True)
    Set RpCol = .Add(Lab_B, "B", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Report, "Report", 0, True)
    Set RpCol = .Add(Lab_Labor, "Labornummer", 0, True)
    Set RpCol = .Add(Lab_Eingang, "Eingang", 0, True)
    Set RpCol = .Add(Lab_Ausgang, "Ausgang", 0, True)
    Set RpCol = .Add(Lab_Arztnummer, "Arztnummer", 0, False)
    Set RpCol = .Add(Lab_Geschlecht, "G", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Beschreibung, "Beschreibung", 0, False)
    Set RpCol = .Add(Lab_Berichtsart, "Befundart", 0, True)
    Set RpCol = .Add(Lab_IDP, "IDP", 0, False)
    Set RpCol = .Add(Lab_Befundung, "B", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Selekt, "Selekt", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Lab_Importdatei, "Importdatei", 0, False)
    Set RpCol = .Add(Lab_Telefon5, "Telefon5", 0, False)
    Set RpCol = .Add(Lab_Telefax, "Telefax", 0, False)
    Set RpCol = .Add(Lab_Briefanrede, "Briefanrede", 0, False)
    Set RpCol = .Add(Lab_IDA, "IDA", 0, False)
    Set RpCol = .Add(Lab_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Lab_IP0, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Lab_Gruppe, "Kostenträgeruntergruppe", 0, False)
    Set RpCol = .Add(Lab_ID3, "Katalog", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
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

If GlIdi = True Then 'Idiotenmodus
    RpCls(Lab_ID0).Width = 0
    RpCls(Lab_Auftrag).Width = 100
    RpCls(Lab_F).Width = 19
    RpCls(Lab_Befundart).Width = 80
    RpCls(Lab_Patient).Width = 220
    RpCls(Lab_Name).Width = 120
    RpCls(Lab_Vorname).Width = 100
    RpCls(Lab_Geboren).Width = 80
    RpCls(Lab_B).Width = 19
    RpCls(Lab_Report).Width = 0
    RpCls(Lab_Labor).Width = 80
    RpCls(Lab_Eingang).Width = 80
    RpCls(Lab_Ausgang).Width = 80
    RpCls(Lab_Arztnummer).Width = 160
    RpCls(Lab_Geschlecht).Width = 50
    RpCls(Lab_Beschreibung).Width = 200
    RpCls(Lab_Berichtsart).Width = 80
    RpCls(Lab_IDP).Width = 0
    RpCls(Lab_Befundung).Width = 0
    RpCls(Lab_Selekt).Width = 0
    RpCls(Lab_Importdatei).Width = 0
    RpCls(Lab_Telefon5).Width = 0
    RpCls(Lab_Telefax).Width = 0
    RpCls(Lab_Briefanrede).Width = 0
    RpCls(Lab_IDA).Width = 0
    RpCls(Lab_Kommentar).Width = 0
    RpCls(Lab_IP0).Width = 180
    RpCls(Lab_Gruppe).Width = 0
    RpCls(Lab_ID3).Width = 150
Else
    If IniGetSek(GlINI, "RpCnt5") = False Then
        RpCls(Lab_ID0).Width = 0
        RpCls(Lab_Auftrag).Width = 100
        RpCls(Lab_F).Width = 19
        RpCls(Lab_Befundart).Width = 80
        RpCls(Lab_Patient).Width = 220
        RpCls(Lab_Name).Width = 120
        RpCls(Lab_Vorname).Width = 100
        RpCls(Lab_Geboren).Width = 80
        RpCls(Lab_B).Width = 19
        RpCls(Lab_Report).Width = 0
        RpCls(Lab_Labor).Width = 80
        RpCls(Lab_Eingang).Width = 80
        RpCls(Lab_Ausgang).Width = 80
        RpCls(Lab_Arztnummer).Width = 160
        RpCls(Lab_Geschlecht).Width = 50
        RpCls(Lab_Beschreibung).Width = 200
        RpCls(Lab_Berichtsart).Width = 80
        RpCls(Lab_IDP).Width = 0
        RpCls(Lab_Befundung).Width = 0
        RpCls(Lab_Selekt).Width = 0
        RpCls(Lab_Importdatei).Width = 0
        RpCls(Lab_Telefon5).Width = 0
        RpCls(Lab_Telefax).Width = 0
        RpCls(Lab_Briefanrede).Width = 0
        RpCls(Lab_IDA).Width = 0
        RpCls(Lab_Kommentar).Width = 0
        RpCls(Lab_IP0).Width = 180
        RpCls(Lab_Gruppe).Width = 0
        RpCls(Lab_ID3).Width = 150
        IniSetSek "RpCnt5"
        IniSetVal "RpCnt5", "SSpLa5a", RpCo5.SaveSettings
    Else
        SpStr = IniGetBig(GlINI, "RpCnt5", "SSpLa5a")
        RpCo5.LoadSettings SpStr
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo5 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa5a " & Err.Number
Resume Next

End Sub
Public Sub SSpLa5b()
On Error GoTo SpErr
'Labortestdaten

Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns

With RpCo6
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCo6
    .AllowEdit = True
    .EditOnClick = False
    .MultipleSelection = True
    .PaintManager.FixedRowHeight = True
    .PaintManager.SetPreviewIndent 80, -2, 10, 6
    .PaintManager.UseAlternativeBackground = GlZeK
End With

With RpCls
    Set RpCol = .Add(Lbl_ID0, "ID0", 0, False)
    Set RpCol = .Add(Lbl_Ident, "Ident", 0, True)
    Set RpCol = .Add(Lbl_Testbezeichnung, "Testbezeichnung", 0, True)
    If RpCo6.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Lbl_Ergebniswert, "Ergebnis", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Lbl_Grenzwert, "Grenz", 0, True)
    Set RpCol = .Add(Lbl_Einheit, "Einheit", 0, True)
    Set RpCol = .Add(Lbl_NormText, "Normalwert", 0, True)
    Set RpCol = .Add(Lbl_Teststatus, "Teststatus", 0, False)
    Set RpCol = .Add(Lbl_ID4, "ID4", 0, False)
    Set RpCol = .Add(Lbl_IDB, "IDB", 0, False)
    Set RpCol = .Add(Lbl_ID2, "ID2", 0, False)
    Set RpCol = .Add(Lbl_TestID, "TestID", 0, False)
    Set RpCol = .Add(Lbl_Hinweis, "Hinweis", 0, False)
    Set RpCol = .Add(Lbl_GONr, "Ziffer", 0, False)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lbl_Betrag, "Betrag", 0, False)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lbl_Multi, "Faktor", 0, False)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lbl_Gruppe, "Gruppe", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentLeft
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        .Visible = GlGrY
    End With
End With

For Each RpCol In RpCls
    RpCol.Editable = True
    RpCol.Groupable = True
    RpCol.Sortable = False
    RpCol.AutoSize = False
    RpCol.Resizable = False
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(Lbl_Ident).Width = 80
RpCls(Lbl_Testbezeichnung).Width = 100
RpCls(Lbl_Testbezeichnung).AutoSize = True
RpCls(Lbl_Ergebniswert).Width = 75
RpCls(Lbl_Grenzwert).Width = 50
RpCls(Lbl_Einheit).Width = 75
RpCls(Lbl_NormText).Width = 120
RpCls(Lbl_GONr).Width = 70
RpCls(Lbl_Betrag).Width = 70
RpCls(Lbl_Multi).Width = 50
RpCls(Lbl_Gruppe).Width = 120

If GlGrY = False Then 'Gruppierung Laborbericht Anzeigen
    RpCls(Lbl_GONr).Visible = True
    RpCls(Lbl_Betrag).Visible = True
    RpCls(Lbl_Multi).Visible = True
    RpCls(Lbl_Gruppe).Visible = False
Else
    RpCls(Lbl_GONr).Visible = False
    RpCls(Lbl_Betrag).Visible = False
    RpCls(Lbl_Multi).Visible = False
    RpCls(Lbl_Gruppe).Visible = True
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo6 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa5b " & Err.Number
Resume Next

End Sub
Private Sub SSpLa5c()
On Error GoTo SpErr
'Laborauftrag

Dim SpStr As String
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns

With RpCo5
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Lau_ID1, "ID1", 0, False)
    Set RpCol = .Add(Lau_Auftrag, "Auftrag", 0, True)
    Set RpCol = .Add(Lau_Datum, "Datum", 0, True)
    Set RpCol = .Add(Lau_Gedruckt, "D", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Selekt, "R", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Kunde, "Patient", 0, True)
    If RpCo5.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Lau_Typ, "Typ", 0, True)
    Set RpCol = .Add(Lau_Name, "Name", 0, True)
    Set RpCol = .Add(Lau_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Lau_Geboren, "Geboren", 0, True)
    Set RpCol = .Add(Lau_G, "G", 0, False)
    Set RpCol = .Add(Lau_Labor, "Labornummer", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Betrag, "Laborpreis", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_GesBetrag, "Gebühren", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_W, "W", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Währung, "Währung", 0, True)
    Set RpCol = .Add(Lau_ID0, "ID0", 0, True)
    Set RpCol = .Add(Lau_Faktor, "Faktor", 0, True)
    Set RpCol = .Add(Lau_ID3, "ID3", 0, True)
    Set RpCol = .Add(Lau_Provision, "Provision", 0, True)
    Set RpCol = .Add(Lau_Konto, "Konto", 0, True)
    Set RpCol = .Add(Lau_IP0, "IP0", 0, True)
    Set RpCol = .Add(Lau_B, "B", 0, True)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_BefKosten, "BefKosten", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_IDA, "IDA", 0, True)
    Set RpCol = .Add(Lau_AbrTyp, "AbrTyp", 0, True)
    Set RpCol = .Add(Lau_GebTyp, "GebTyp", 0, True)
    Set RpCol = .Add(Lau_Kommentar, "Kommentar", 0, True)
    Set RpCol = .Add(Lau_Multi, "Multi", 0, True)
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

If IniGetSek(GlINI, "RpLab1") = False Then
    RpCls(0).Width = 0
    RpCls(1).Width = 100
    RpCls(2).Width = 80
    RpCls(3).Width = 19
    RpCls(4).Width = 19
    RpCls(5).Width = 220
    RpCls(6).Width = 30
    RpCls(7).Width = 120
    RpCls(8).Width = 100
    RpCls(9).Width = 80
    RpCls(10).Width = 19
    RpCls(11).Width = 80
    RpCls(12).Width = 80
    RpCls(13).Width = 80
    RpCls(14).Width = 19
    RpCls(15).Width = 0
    RpCls(16).Width = 0
    RpCls(17).Width = 0
    RpCls(18).Width = 0
    RpCls(19).Width = 0
    RpCls(20).Width = 0
    RpCls(21).Width = 0
    RpCls(22).Width = 19
    RpCls(23).Width = 80
    RpCls(24).Width = 0
    RpCls(25).Width = 0
    RpCls(26).Width = 0
    RpCls(27).Width = 0
    RpCls(28).Width = 0
    IniSetSek "RpLab1"
    IniSetVal "RpLab1", "SSpLa5c", RpCo5.SaveSettings
Else
    SpStr = IniGetBig(GlINI, "RpLab1", "SSpLa5c")
    RpCo5.LoadSettings SpStr
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo5 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa5c " & Err.Number
Resume Next

End Sub
Public Sub SSpLa5d()
On Error GoTo SpErr

Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns

With RpCo6
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCo6
    .AllowEdit = True
    .EditOnClick = False
    .MultipleSelection = True
    .PaintManager.FixedRowHeight = True
    .PaintManager.SetPreviewIndent 184, -2, 10, 6
    .PaintManager.UseAlternativeBackground = GlZeK
End With

With RpCls
    Set RpCol = .Add(Lba_ID2, "ID0", 0, False)
    Set RpCol = .Add(Lba_IDA, "IDA", 0, False)
    Set RpCol = .Add(Lba_Ident, "Ident", 0, False)
    Set RpCol = .Add(Lba_GONr, "Ziffer", 0, False)
    Set RpCol = .Add(Lba_Faktor, "Faktor", 0, False)
    Set RpCol = .Add(Lba_Bezeichnung, "Testbezeichnung", 0, True)
    If RpCo6.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Lba_x, "x", 0, False)
    Set RpCol = .Add(Lba_Einheit, "Einheit", 0, False)
    Set RpCol = .Add(Lba_Zuweisung, "Arbeitsplatz", 0, False)
    Set RpCol = .Add(Lba_Höchstwert, "H", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lba_Profilwert, "P", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lba_Kettenwert, "K", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lba_Betrag, "Betrag", 0, False)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lba_GesBetrag, "Gesamt", 0, False)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lba_Provision, "Provision", 0, False)
    Set RpCol = .Add(Lba_IDU, "IDU", 0, False)
    Set RpCol = .Add(Lba_ID0, "ID0", 0, False)
    Set RpCol = .Add(Lba_ID1, "ID1", 0, False)
    Set RpCol = .Add(Lba_ID4, "ID4", 0, False)
    Set RpCol = .Add(Lba_Kommentar, "Kommentar", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Editable = True
    RpCol.Groupable = True
    RpCol.Sortable = False
    RpCol.AutoSize = False
    RpCol.Resizable = False
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(Lba_Ident).Width = 80
RpCls(Lba_GONr).Width = 80
RpCls(Lba_Faktor).Visible = False
RpCls(Lba_Bezeichnung).Width = 100
RpCls(Lba_Bezeichnung).AutoSize = True
RpCls(Lbl_Einheit).Width = 60
RpCls(Lba_Höchstwert).Width = 0
RpCls(Lba_Profilwert).Width = 0
RpCls(Lba_Kettenwert).Width = 0
RpCls(Lba_Betrag).Width = 80
RpCls(Lba_GesBetrag).Width = 80

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo6 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa5d " & Err.Number
Resume Next

End Sub
Private Sub SSpLa5e()
On Error GoTo SpErr
'Laborberichte

Dim SpStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Lab_ID0, "ID0", 0, False)
    Set RpCol = .Add(Lab_Auftrag, "Auftrag", 0, True)
    Set RpCol = .Add(Lab_F, "R", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Befundart, "Befundart", 0, True)
    Set RpCol = .Add(Lab_Patient, "Patient", 0, True)
    Set RpCol = .Add(Lab_Name, "Name", 0, True)
    Set RpCol = .Add(Lab_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Lab_Geboren, "Geboren", 0, True)
    Set RpCol = .Add(Lab_B, "B", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Report, "Report", 0, True)
    Set RpCol = .Add(Lab_Labor, "Labornummer", 0, True)
    Set RpCol = .Add(Lab_Eingang, "Eingang", 0, True)
    Set RpCol = .Add(Lab_Ausgang, "Ausgang", 0, True)
    Set RpCol = .Add(Lab_Arztnummer, "Arztnummer", 0, False)
    Set RpCol = .Add(Lab_Geschlecht, "G", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Beschreibung, "Beschreibung", 0, False)
    Set RpCol = .Add(Lab_Berichtsart, "Befundart", 0, True)
    Set RpCol = .Add(Lab_IDP, "IDP", 0, False)
    Set RpCol = .Add(Lab_Befundung, "B", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lab_Selekt, "Selekt", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Lab_Importdatei, "Importdatei", 0, False)
    Set RpCol = .Add(Lab_Telefon5, "Telefon5", 0, False)
    Set RpCol = .Add(Lab_Telefax, "Telefax", 0, False)
    Set RpCol = .Add(Lab_Briefanrede, "Briefanrede", 0, False)
    Set RpCol = .Add(Lab_IDA, "IDA", 0, False)
    Set RpCol = .Add(Lab_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Lab_IP0, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Lab_Gruppe, "Kostenträgeruntergruppe", 0, False)
    Set RpCol = .Add(Lab_ID3, "Katalog", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
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

If GlIdi = True Then 'Idiotenmodus
    RpCls(Lab_ID0).Width = 0
    RpCls(Lab_Auftrag).Width = 100
    RpCls(Lab_F).Width = 19
    RpCls(Lab_Befundart).Width = 80
    RpCls(Lab_Patient).Width = 220
    RpCls(Lab_Name).Width = 120
    RpCls(Lab_Vorname).Width = 100
    RpCls(Lab_Geboren).Width = 80
    RpCls(Lab_B).Width = 19
    RpCls(Lab_Report).Width = 0
    RpCls(Lab_Labor).Width = 80
    RpCls(Lab_Eingang).Width = 80
    RpCls(Lab_Ausgang).Width = 80
    RpCls(Lab_Arztnummer).Width = 160
    RpCls(Lab_Geschlecht).Width = 50
    RpCls(Lab_Beschreibung).Width = 200
    RpCls(Lab_Berichtsart).Width = 80
    RpCls(Lab_IDP).Width = 0
    RpCls(Lab_Befundung).Width = 0
    RpCls(Lab_Selekt).Width = 0
    RpCls(Lab_Importdatei).Width = 0
    RpCls(Lab_Telefon5).Width = 0
    RpCls(Lab_Telefax).Width = 0
    RpCls(Lab_Briefanrede).Width = 0
    RpCls(Lab_IDA).Width = 0
    RpCls(Lab_Kommentar).Width = 0
    RpCls(Lab_IP0).Width = 180
    RpCls(Lab_Gruppe).Width = 0
    RpCls(Lab_ID3).Width = 150
Else
    If IniGetSek(GlINI, "RpCnt6") = False Then
        RpCls(Lab_ID0).Width = 0
        RpCls(Lab_Auftrag).Width = 100
        RpCls(Lab_F).Width = 19
        RpCls(Lab_Befundart).Width = 80
        RpCls(Lab_Patient).Width = 220
        RpCls(Lab_Name).Width = 120
        RpCls(Lab_Vorname).Width = 100
        RpCls(Lab_Geboren).Width = 80
        RpCls(Lab_B).Width = 19
        RpCls(Lab_Report).Width = 0
        RpCls(Lab_Labor).Width = 80
        RpCls(Lab_Eingang).Width = 80
        RpCls(Lab_Ausgang).Width = 80
        RpCls(Lab_Arztnummer).Width = 160
        RpCls(Lab_Geschlecht).Width = 50
        RpCls(Lab_Beschreibung).Width = 200
        RpCls(Lab_Berichtsart).Width = 80
        RpCls(Lab_IDP).Width = 0
        RpCls(Lab_Befundung).Width = 0
        RpCls(Lab_Selekt).Width = 0
        RpCls(Lab_Importdatei).Width = 0
        RpCls(Lab_Telefon5).Width = 0
        RpCls(Lab_Telefax).Width = 0
        RpCls(Lab_Briefanrede).Width = 0
        RpCls(Lab_IDA).Width = 0
        RpCls(Lab_Kommentar).Width = 0
        RpCls(Lab_IP0).Width = 180
        RpCls(Lab_ID3).Width = 150
        RpCls(Lab_Gruppe).Width = 0
        IniSetSek "RpCnt6"
        IniSetVal "RpCnt6", "SSpLa5e", RpCo1.SaveSettings
    Else
        SpStr = IniGetBig(GlINI, "RpCnt6", "SSpLa5e")
        RpCo1.LoadSettings SpStr
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa5e " & Err.Number
Resume Next

End Sub
Private Sub SSpLa5f()
On Error GoTo SpErr
'Laboraufträge

Dim SpStr As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Lau_ID1, "ID1", 0, False)
    Set RpCol = .Add(Lau_Auftrag, "Auftrag", 0, True)
    Set RpCol = .Add(Lau_Datum, "Datum", 0, True)
    Set RpCol = .Add(Lau_Gedruckt, "D", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Selekt, "R", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Kunde, "Patient", 0, True)
    If RpCo1.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Lau_Typ, "Typ", 0, True)
    Set RpCol = .Add(Lau_Name, "Name", 0, True)
    Set RpCol = .Add(Lau_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Lau_Geboren, "Geboren", 0, True)
    Set RpCol = .Add(Lau_G, "G", 0, False)
    Set RpCol = .Add(Lau_Labor, "Labornummer", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Betrag, "Laborpreis", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_GesBetrag, "Gebühren", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_W, "W", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_Währung, "Währung", 0, True)
    Set RpCol = .Add(Lau_ID0, "ID0", 0, True)
    Set RpCol = .Add(Lau_Faktor, "Faktor", 0, True)
    Set RpCol = .Add(Lau_ID3, "ID3", 0, True)
    Set RpCol = .Add(Lau_Provision, "Provision", 0, True)
    Set RpCol = .Add(Lau_Konto, "Konto", 0, True)
    Set RpCol = .Add(Lau_IP0, "IP0", 0, True)
    Set RpCol = .Add(Lau_B, "B", 0, True)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_BefKosten, "BefKosten", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Lau_IDA, "IDA", 0, True)
    Set RpCol = .Add(Lau_AbrTyp, "AbrTyp", 0, True)
    Set RpCol = .Add(Lau_GebTyp, "GebTyp", 0, True)
    Set RpCol = .Add(Lau_Kommentar, "Kommentar", 0, True)
    Set RpCol = .Add(Lau_Multi, "Multi", 0, True)
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

If IniGetSek(GlINI, "RpLab2") = False Then
    RpCls(0).Width = 0
    RpCls(1).Width = 100
    RpCls(2).Width = 80
    RpCls(3).Width = 19
    RpCls(4).Width = 19
    RpCls(5).Width = 220
    RpCls(6).Width = 30
    RpCls(7).Width = 120
    RpCls(8).Width = 100
    RpCls(9).Width = 80
    RpCls(10).Width = 19
    RpCls(11).Width = 80
    RpCls(12).Width = 80
    RpCls(13).Width = 80
    RpCls(14).Width = 19
    RpCls(15).Width = 0
    RpCls(16).Width = 0
    RpCls(17).Width = 0
    RpCls(18).Width = 0
    RpCls(19).Width = 0
    RpCls(20).Width = 0
    RpCls(21).Width = 0
    RpCls(22).Width = 19
    RpCls(23).Width = 80
    RpCls(24).Width = 0
    RpCls(25).Width = 0
    RpCls(26).Width = 0
    RpCls(27).Width = 0
    RpCls(28).Width = 0
    IniSetSek "RpLab2"
    IniSetVal "RpLab2", "SSpLa5f", RpCo1.SaveSettings
Else
    SpStr = IniGetBig(GlINI, "RpLab2", "SSpLa5f")
    RpCo1.LoadSettings SpStr
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa5f " & Err.Number
Resume Next

End Sub
Private Sub SSpLa6()
On Error GoTo SpErr

Dim SpStr As String
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns

With RpCo5
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Rzp_ID0, "ID0", 0, False)
    Set RpCol = .Add(Rzp_ID1, "Beleg", 0, False)
    Set RpCol = .Add(Rzp_Rezepttext, "Rezepttext", 0, False)
    Set RpCol = .Add(Rzp_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Rzp_Datum, "Datum", 0, True)
    Set RpCol = .Add(Rzp_DatVon, "DatVon", 0, False)
    Set RpCol = .Add(Rzp_DatBis, "DatBis", 0, False)
    Set RpCol = .Add(Rzp_DatNeu, "DatNeu", 0, False)
    Set RpCol = .Add(Rzp_ZeiVon, "ZeiVon", 0, False)
    Set RpCol = .Add(Rzp_ZeiBis, "ZeiBis", 0, False)
    Set RpCol = .Add(Rzp_Drucken, "D", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.HeaderAlignment = xtpAlignmentCenter
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Rzp_Type, "Type", 0, False)
    Set RpCol = .Add(Rzp_Unfall, "Unfall", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Arbeitsunfall, "Arbeitsunfall", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Erstbescheinigung, "Erstbescheinigung", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Folgebescheinigung, "Folgebescheinigung", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Durchgangsarzt, "Durchgangsarzt", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Arbeitgeber, "Arbeitgeber", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Sonstige, "Sonstige", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Schulunterricht, "Schulunterricht", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Sportunterricht, "Sportunterricht", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Absender, "Absender", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_GesBetrag, "GesBetrag", 0, False)
    Set RpCol = .Add(Rzp_GebFrei, "GebFrei", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_GebPfl, "GebPfl", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_AutIdem1, "AutIdem1", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_AutIdem2, "AutIdem2", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_AutIdem3, "AutIdem3", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_BVG, "BVG", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_EWRCH, "EWRCH", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Gruppen, "Gruppen", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_Regelfall, "Regelfall", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_HausBes1, "HausBes1", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_HausBes2, "HausBes2", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_TheraBer1, "TheraBer1", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_TheraBer2, "TheraBer2", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_GesZuza, "GesZuza", 0, False)
    Set RpCol = .Add(Rzp_GesBrut, "GesBrut", 0, False)
    Set RpCol = .Add(Rzp_HeilmPos1, "HeilmPos1", 0, False)
    Set RpCol = .Add(Rzp_HeilmPos2, "HeilmPos2", 0, False)
    Set RpCol = .Add(Rzp_Wegegeld, "Wegegeld", 0, False)
    Set RpCol = .Add(Rzp_Hausbet1, "Hausbet1", 0, False)
    Set RpCol = .Add(Rzp_Hausbet2, "Hausbet2", 0, False)
    Set RpCol = .Add(Rzp_Faktor1, "Faktor1", 0, False)
    Set RpCol = .Add(Rzp_Faktor2, "Faktor2", 0, False)
    Set RpCol = .Add(Rzp_Faktor3, "Faktor3", 0, False)
    Set RpCol = .Add(Rzp_Faktor4, "Faktor4", 0, False)
    Set RpCol = .Add(Rzp_Faktor5, "Faktor5", 0, False)
    Set RpCol = .Add(Rzp_RechNr, "RechNr", 0, False)
    Set RpCol = .Add(Rzp_BelegNr, "BelegNr", 0, False)
    Set RpCol = .Add(Rzp_Firma1, "Firma1", 0, False)
    Set RpCol = .Add(Rzp_Anrede, "Anrede", 0, False)
    Set RpCol = .Add(Rzp_Titel, "Titel", 0, False)
    Set RpCol = .Add(Rzp_Name, "Name", 0, True)
    Set RpCol = .Add(Rzp_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Rzp_Straße, "Straße", 0, False)
    Set RpCol = .Add(Rzp_PLZ, "PLZ", 0, False)
    Set RpCol = .Add(Rzp_Ort, "Ort", 0, False)
    Set RpCol = .Add(Rzp_LK, "LK", 0, False)
    Set RpCol = .Add(Rzp_Land, "Land", 0, False)
    Set RpCol = .Add(Rzp_Briefanrede, "Briefanrede", 0, False)
    Set RpCol = .Add(Rzp_Geboren, "Geboren", 0, False)
    Set RpCol = .Add(Rzp_Anschrift, "Anschrift", 0, False)
    Set RpCol = .Add(Rzp_Versicherung, "Versicherung", 0, False)
    Set RpCol = .Add(Rzp_Kartennummer, "Kartennummer", 0, False)
    Set RpCol = .Add(Rzp_Kartengultig, "Kartengultig", 0, False)
    Set RpCol = .Add(Rzp_Kartenstatus, "Kartenstatus", 0, False)
    Set RpCol = .Add(Rzp_KVNummer, "KVNummer", 0, False)
    Set RpCol = .Add(Rzp_Beleg, "Belegvorlage", 0, True)
    Set RpCol = .Add(Rzp_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Rzp_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Rzp_EM_User, "EM_User", 0, False)
    Set RpCol = .Add(Rzp_EM_Pass, "EM_Pass", 0, False)
    Set RpCol = .Add(Rzp_EM_Aut, "EM_Aut", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Rzp_BetrText, "BetrText", 0, False)
    Set RpCol = .Add(Rzp_HeiMtl1, "Heilmittel1", 0, False)
    Set RpCol = .Add(Rzp_HeiMtl2, "Heilmittel2", 0, False)
    Set RpCol = .Add(Rzp_Diagnos, "Diagnose", 0, False)
    Set RpCol = .Add(Rzp_TherZie, "Therapieziele", 0, False)
    Set RpCol = .Add(Rzp_Begründ, "Begründung", 0, False)
    Set RpCol = .Add(Rzp_AnzTerm, "Anzhal", 0, False)
    Set RpCol = .Add(Rzp_AnzWoVo, "AnzWoche1", 0, False)
    Set RpCol = .Add(Rzp_AnzWoBi, "AnzWoche2", 0, False)
    Set RpCol = .Add(Rzp_Storniert, "Storniert", 0, False)
    Set RpCol = .Add(Rzp_TSEString, "TSEString", 0, False)
    Set RpCol = .Add(Rzp_TSELog, "TSELog", 0, False)
    Set RpCol = .Add(Rzp_TSEZahl, "TSEZahl", 0, False)
    Set RpCol = .Add(Rzp_TSESign, "TSESign", 0, False)
    Set RpCol = .Add(Rzp_TSEStrSig, "TSEStrSig", 0, False)
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

RpCls(Rzp_ID1).Width = 70
RpCls(Rzp_Kommentar).Width = 200
RpCls(Rzp_Datum).Width = 80
RpCls(Rzp_Drucken).Width = 19
RpCls(Rzp_Name).Width = 110
RpCls(Rzp_Vorname).Width = 110
RpCls(Rzp_Geboren).Width = 80
RpCls(Rzp_Beleg).Width = 200
RpCls(Rzp_IDP).Width = 180
RpCls(Rzp_IDM).Width = 180

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo5 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa6 " & Err.Number
Resume Next

End Sub
Private Sub SSpLa7()
On Error GoTo SpErr

Dim AktZa As Integer
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set TxMul = FM.txtMulti
Set CmMit = FM.cmbMitar
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns

With RpCo6
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Select Case GlBut
Case RibTab_Abrechnung:
        With RpCo6
            .AllowEdit = True
            .EditOnClick = GlKrE
            .EditOnDoubleClick = True
            .MultipleSelection = True
            .PaintManager.FixedRowHeight = True
            .PaintManager.UseAlternativeBackground = GlZeK
            If GlSpK = True Then 'Katalogspalte
                .PaintManager.SetPreviewIndent 252, -2, 10, 6
            Else
                .PaintManager.SetPreviewIndent 192, -2, 10, 6
            End If
            .ShowHeader = GlSpU
            .ShowFooter = True
        End With
Case RibTab_Tagesproto:
        With RpCo6
            .AllowEdit = False
            .EditOnClick = False
            .EditOnDoubleClick = False
            .MultipleSelection = False
            .PaintManager.FixedRowHeight = True
            .PaintManager.SetPreviewIndent 112, -2, 10, 6
            .PaintManager.UseAlternativeBackground = False
            .ShowHeader = False
            .ShowFooter = False
        End With
End Select

With RpCls
    Set RpCol = .Add(Kra_ID2, "ID2", 0, False)
    Set RpCol = .Add(Kra_ID0, "ID0", 0, False)
    If GlFri = 5 Then
        Set RpCol = .Add(Kra_KatNa, "Tarif", 0, False)
    Else
        Set RpCol = .Add(Kra_KatNa, "Taxe", 0, False)
    End If
    With RpCol
        .EditOptions.SelectTextOnEdit = True
        If GlBut = RibTab_Tagesproto Then
            .Visible = False
        Else
            .Visible = GlSpK 'Katalogspalte
        End If
    End With
    Set RpCol = .Add(Kra_ID3, "ID3", 0, False)
    Set RpCol = .Add(Kra_Provision, "Format", 0, False)
    Set RpCol = .Add(Kra_ID4, "ID4", 0, False)
    Set RpCol = .Add(Kra_KrTyp, "Typ", 0, False)
    Set RpCol = .Add(Kra_IDR, "IDR", 0, False)
    Set RpCol = .Add(Kra_Datum, "Datum", 0, False)
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.SelectTextOnEdit = True
        .EditOptions.AllowEdit = True
    End With
    Set RpCol = .Add(Kra_Uhrzeit, "Uhrzeit", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.SelectTextOnEdit = True
        .Alignment = xtpAlignmentCenter
    End With
    Select Case GlBut
    Case RibTab_Abrechnung: RpCol.Visible = GlSpZ
    Case RibTab_Tagesproto: RpCol.Visible = False
    End Select
    Set RpCol = .Add(Kra_Typ, "Typ", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        .Visible = GlETy
        If GlTyV = True Then 'Krankenblattypen vorhanden
            For AktZa = 1 To UBound(GlKrA)
                .EditOptions.Constraints.Add GlKrA(AktZa, 1), GlKrA(AktZa, 0)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(Kra_Ziffer, "Ziffer", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.SelectTextOnEdit = True
        .FooterFont.Bold = True
    End With
    Set RpCol = .Add(Kra_Bezeichnung, "Bezeichnung", 0, False)
    With RpCol
        If RpCo6.PaintManager.FixedRowHeight = False Then
            .Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If .Editable = True Then
                .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll Or xtpEditStyleMultiline
                .TreeColumn = False
                .EditOptions.MaxLength = 250
                .EditOptions.SelectTextOnEdit = True
            End If
        End If
    End With
    Set RpCol = .Add(Kra_Analog, "A", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        If GlBut = RibTab_Tagesproto Then
            .Visible = False
        Else
            .Visible = GlAnl
        End If
    End With
    Set RpCol = .Add(Kra_Anz, "Anz", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.SelectTextOnEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleNumber
        If GlBut = RibTab_Tagesproto Then
            .Visible = False
        Else
            Select Case GlFri
            Case 4: 'Veterinär (GOT)
                .Visible = False
            Case 5: 'Naturheilpraktiker (Tarif 590)
                .Visible = True
            Case Else:
                .Visible = True
            End Select
        End If
    End With
    Set RpCol = .Add(Kra_Faktor, "Fakt.", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.SelectTextOnEdit = True
        If GlBut = RibTab_Tagesproto Then
            .Visible = False
        Else
            .Visible = GlSpM
            TxMul.Visible = GlSpM
        End If
    End With
    Set RpCol = .Add(Kra_Betrag, "Einzel", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.SelectTextOnEdit = True
        Select Case GlBut
        Case RibTab_Abrechnung: .Visible = True
        Case RibTab_Tagesproto: .Visible = False
        End Select
        .FooterFont.Bold = True
        .FooterFont.SIZE = GlTFt.SIZE
        .FooterAlignment = xtpAlignmentRight
    End With
    Select Case GlBut
    Case RibTab_Abrechnung:
                Set RpCol = .Add(Kra_GesBetrag, "Gesamt", 0, False)
                With RpCol
                    .Alignment = xtpAlignmentRight
                    .HeaderAlignment = xtpAlignmentCenter
                    .Visible = True
                    .FooterFont.Bold = True
                    .FooterFont.SIZE = GlTFt.SIZE
                    .FooterAlignment = xtpAlignmentRight
                End With
    Case RibTab_Tagesproto:
                Set RpCol = .Add(Kra_GesBetrag, "Gesamt", 0, False)
                With RpCol
                    .Alignment = xtpAlignmentRight
                    .HeaderAlignment = xtpAlignmentCenter
                    .Visible = False
                End With
    End Select
    Set RpCol = .Add(Kra_WVBetrag, "Akonto", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        Select Case GlBut
        Case RibTab_Abrechnung: .Visible = False
        Case RibTab_Tagesproto: .Visible = False
        End Select
    End With
    Set RpCol = .Add(Kra_LaBetrag, "Einstand", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        Select Case GlBut
        Case RibTab_Abrechnung: .Visible = GlEns
        Case RibTab_Tagesproto: .Visible = False
        End Select
    End With
    Set RpCol = .Add(Kra_Steuersatz, "Steuer", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.AllowEdit = True
        .Visible = GlSpl
    End With
    Select Case GlBut
    Case RibTab_Abrechnung:
                Set RpCol = .Add(Kra_Zeit, "Min.", 0, False)
                With RpCol
                    .Alignment = xtpAlignmentRight
                    .HeaderAlignment = xtpAlignmentCenter
                    .Visible = GlSpZ
                End With
    Case RibTab_Tagesproto:
                Set RpCol = .Add(Kra_Zeit, "Min.", 0, False)
                With RpCol
                    .Alignment = xtpAlignmentRight
                    .HeaderAlignment = xtpAlignmentCenter
                    .Visible = False
                End With
    End Select
    Set RpCol = .Add(Kra_Einheit, "Einheit", 0, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .HeaderAlignment = xtpAlignmentCenter
        Select Case GlBut
        Case RibTab_Abrechnung: .Visible = GlEnh
        Case RibTab_Tagesproto: .Visible = GlEnh
        End Select
    End With
    Set RpCol = .Add(Kra_Währung, "W", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Kra_IDD, "Diagnosezuord.", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        Select Case GlBut
        Case RibTab_Abrechnung: .Visible = GlKDi
        Case RibTab_Tagesproto: .Visible = False
        End Select
    End With
    If GlMsp = True Then 'Mandantenspalte anstelle von Mitarbeiterspalte in Abrechnung
        Set RpCol = .Add(Kra_IDM, "Mandanten", 0, False)
    Else
        Set RpCol = .Add(Kra_IDM, "Mitarbeiter", 0, False)
    End If
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        Select Case GlBut
        Case RibTab_Abrechnung: .Visible = GlMiZ
        Case RibTab_Tagesproto: .Visible = True
        End Select
    End With
    CmMit.Visible = GlMiZ
    Set RpCol = .Add(Kra_Gedruckt, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Check
        .Tag = 1
    End With
    Set RpCol = .Add(Kra_Selekt, "Selekt", 0, False)
    Set RpCol = .Add(Kra_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Kra_Zusatztext, "Zusatztext", 0, False)
    Set RpCol = .Add(Kra_Typname, "Typname", 0, False)
    Set RpCol = .Add(Kra_Quart, "Quartal", 0, False)
    Set RpCol = .Add(Kra_Monat, "Monat", 0, False)
    Set RpCol = .Add(Kra_Woche, "Woche", 0, False)
    Set RpCol = .Add(Kra_TagSo, "Datum", 0, False)
    Set RpCol = .Add(Kra_Lock, "Lock", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.Icon = IC16_Lock
    RpCol.Tag = 1
    Set RpCol = .Add(Kra_Storniert, "Storniert", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Groupable = True
    RpCol.Sortable = False
    RpCol.AutoSize = False
    RpCol.Resizable = False
    RpCol.AutoSortWhenGrouped = False
Next RpCol

If GlKFt.SIZE > 10 Then
    RpCls(Kra_Datum).Width = 110
    RpCls(Kra_Typ).Width = 50
Else
    RpCls(Kra_Datum).Width = 80
    RpCls(Kra_Typ).Width = 30
End If

RpCls(Kra_Bezeichnung).AutoSize = True

RpCls(Kra_GesBetrag).Editable = False
RpCls(Kra_WVBetrag).Editable = False
RpCls(Kra_Gedruckt).Editable = False

Select Case GlBut
Case RibTab_Abrechnung:
    RpCls(Kra_ID3).Width = 0
    RpCls(Kra_KatNa).Width = 60
    RpCls(Kra_Uhrzeit).Width = 60
    RpCls(Kra_Ziffer).Width = 80
    RpCls(Kra_Anz).Width = 35
    RpCls(Kra_Faktor).Width = 45
    RpCls(Kra_Analog).Width = 30
    If GlKFt.SIZE > 10 Then
        RpCls(Kra_Betrag).Width = 80
        RpCls(Kra_GesBetrag).Width = 80
        RpCls(Kra_LaBetrag).Width = 80
        RpCls(Kra_Steuersatz).Width = 80
        RpCls(Kra_Zeit).Width = 60
    Else
        RpCls(Kra_Betrag).Width = 60
        RpCls(Kra_GesBetrag).Width = 60
        RpCls(Kra_LaBetrag).Width = 60
        RpCls(Kra_Steuersatz).Width = 60
        RpCls(Kra_Zeit).Width = 40
    End If
    RpCls(Kra_Einheit).Width = 40
    RpCls(Kra_IDD).Width = 120
    RpCls(Kra_IDM).Width = 80
    RpCls(Kra_Gedruckt).Width = 18
Case RibTab_Tagesproto:
    RpCls(Kra_KatNa).Width = 130
    RpCls(Kra_Uhrzeit).Width = 0
    RpCls(Kra_Ziffer).Width = 80
    RpCls(Kra_Anz).Width = 0
    RpCls(Kra_Faktor).Width = 0
    RpCls(Kra_Betrag).Width = 0
    RpCls(Kra_GesBetrag).Width = 0
    RpCls(Kra_Steuersatz).Width = 0
    RpCls(Kra_Zeit).Width = 40
    RpCls(Kra_Einheit).Width = 0
    RpCls(Kra_IDD).Width = 0
    RpCls(Kra_IDM).Width = 140
    RpCls(Kra_Gedruckt).Width = 0
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo6 = Nothing

If GlSta = False Then
    S_KrSpl
End If

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa7 " & Err.Number
Resume Next

End Sub
Private Sub SSpLa7b()
On Error GoTo SpErr

Dim SpStr As String
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns

With RpCo6
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

 With RpCo6
    .AllowEdit = False
    .EditOnClick = False
    .EditOnDoubleClick = False
    .MultipleSelection = True
    .PaintManager.FixedRowHeight = True
    .PaintManager.SetPreviewIndent 112, -2, 10, 6
    .PaintManager.UseAlternativeBackground = False
    .ShowFooter = False
End With

With RpCls
    Set RpCol = .Add(Adr_ID0, "ID0", 0, False)
    Set RpCol = .Add(Adr_ID3, "ID3", 0, False)
    Set RpCol = .Add(Adr_IDKurz, "Suchbegriff", 0, True)
    If RpCo6.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Adr_Geboren, "Geboren", 0, True)
    Set RpCol = .Add(Adr_Name, "Name", 0, True)
    Set RpCol = .Add(Adr_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Adr_Straße, "Straße", 0, True)
    Set RpCol = .Add(Adr_PLZ, "PLZ", 0, True)
    Set RpCol = .Add(Adr_Ort, "Ort", 0, True)
    Set RpCol = .Add(Adr_Firma1, "Firma", 0, True)
    Set RpCol = .Add(Adr_Telefon1, "Privat", 0, True)
    Set RpCol = .Add(Adr_Telefon2, "Büro", 0, True)
    Set RpCol = .Add(Adr_Telefon3, "Telefax", 0, True)
    Set RpCol = .Add(Adr_Telefon4, "Mobil", 0, True)
    Set RpCol = .Add(Adr_Telefon5, "Email", 0, True)
    Set RpCol = .Add(Adr_Geschlecht, "G", 0, True)
    Set RpCol = .Add(Adr_Datum, "Datun", 0, False)
    Set RpCol = .Add(Adr_Briefanrede, "Briefanrede", 0, False)
    Set RpCol = .Add(Adr_Anschrift, "Anschrift", 0, False)
    Set RpCol = .Add(Adr_TreKey, "TreKey", 0, False)
    Set RpCol = .Add(Adr_Grafik, "Grafik", 0, False)
    Set RpCol = .Add(Adr_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Adr_Objekt, "Objekt", 0, False)
    Set RpCol = .Add(Adr_IDP, "Mandant", 0, True)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Adr_Mandant, "Datum", 0, True)
    Set RpCol = .Add(Adr_VIP, "VIP", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Adr_Titel, "Titel", 0, False)
    Set RpCol = .Add(Adr_Land, "Tage", 0, False)
    Set RpCol = .Add(Adr_Behindert, "Minimal", 0, False)
    Set RpCol = .Add(Adr_Passiv, "Maximal", 0, False)
    Set RpCol = .Add(Adr_Gruppen, "Datum", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Check
        .Tag = 1
    End With
    Set RpCol = .Add(Adr_Versand, "V", 0, True)
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

RpCls(0).Width = 0
RpCls(1).Width = 0
RpCls(2).Width = 220
If GlTFt.SIZE > 10 Then
    RpCls(3).Width = 110
Else
    RpCls(3).Width = 80
End If
RpCls(4).Width = 100
RpCls(5).Width = 100
RpCls(6).Width = 120
RpCls(7).Width = 60
RpCls(8).Width = 100
RpCls(9).Width = 0
RpCls(10).Width = 0
RpCls(11).Width = 0
RpCls(12).Width = 0
RpCls(13).Width = 0
RpCls(14).Width = 0
RpCls(15).Width = 0
RpCls(16).Width = 0
RpCls(17).Width = 0
RpCls(18).Width = 0
RpCls(19).Width = 0
RpCls(20).Width = 0
RpCls(21).Width = 0
RpCls(22).Width = 0
RpCls(23).Width = 0
RpCls(24).Width = 80
RpCls(25).Width = 0
RpCls(26).Width = 0
RpCls(27).Width = 0
RpCls(28).Width = 80
RpCls(29).Width = 80
RpCls(30).Width = 18

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo6 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa7b " & Err.Number
Resume Next

End Sub
Public Sub SSpLa8()
On Error GoTo SpErr

Dim SpStr As String
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCls = RpCo3.Columns

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
End With

For Each RpCol In RpCls
    With RpCol
        .Editable = False
        .Groupable = True
        .Sortable = True
        .AutoSize = False
        .AutoSortWhenGrouped = False
    End With
Next RpCol

If GlIdi = True Then 'Idiotenmodus
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
    RpCls(Rec_Versand).Width = 20
    RpCls(Rec_Betrag).Width = 75
    RpCls(Rec_Bezahlt).Width = 75
    RpCls(Rec_Differe).Width = 75
    RpCls(Rec_IDKurz).Width = 220
    RpCls(Rec_Offen).Width = 0
    RpCls(Rec_Extrageb).Width = 75
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
    RpCls(Rec_IDP).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_Druckdatum).Width = 110
    Else
        RpCls(Rec_Druckdatum).Width = 80
    End If
    RpCls(Rec_Kopie).Width = 0
    RpCls(Rec_Steuer).Width = 60
    RpCls(Rec_Monat).Width = 0
    RpCls(Rec_Termin).Width = 75
    RpCls(Rec_Storniert).Width = 0
    RpCls(Rec_PKU).Width = 50
    RpCls(Rec_Beendet).Width = 0
    RpCls(Rec_Rabatt).Width = 0
    RpCls(Rec_IDM).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_GuStr).Width = 110
    Else
        RpCls(Rec_GuStr).Width = 80
    End If
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
Else
    If IniGetSek(GlINI, "SplRp8") = False Then
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
        RpCls(Rec_Versand).Width = 20
        RpCls(Rec_Betrag).Width = 75
        RpCls(Rec_Bezahlt).Width = 75
        RpCls(Rec_Differe).Width = 75
        RpCls(Rec_IDKurz).Width = 220
        RpCls(Rec_Offen).Width = 0
        RpCls(Rec_Extrageb).Width = 75
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
        RpCls(Rec_IDP).Width = 180
        If GlTFt.SIZE > 10 Then
            RpCls(Rec_Druckdatum).Width = 110
        Else
            RpCls(Rec_Druckdatum).Width = 80
        End If
        RpCls(Rec_Kopie).Width = 0
        RpCls(Rec_Steuer).Width = 60
        RpCls(Rec_Monat).Width = 0
        RpCls(Rec_Termin).Width = 75
        RpCls(Rec_Storniert).Width = 0
        RpCls(Rec_PKU).Width = 50
        RpCls(Rec_Beendet).Width = 0
        RpCls(Rec_Rabatt).Width = 0
        RpCls(Rec_IDM).Width = 180
        If GlTFt.SIZE > 10 Then
            RpCls(Rec_GuStr).Width = 110
        Else
            RpCls(Rec_GuStr).Width = 80
        End If
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
        IniSetSek "SplRp8"
        IniSetVal "SplRp8", "SSpLa8", RpCo3.SaveSettings
    Else
        SpStr = IniGetBig(GlINI, "SplRp8", "SSpLa8")
        RpCo3.LoadSettings SpStr
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo3 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLa8 " & Err.Number
Resume Next

End Sub
Private Sub SSpLa9()
On Error GoTo SpErr

Dim SpStr As String
Dim AktZa As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Ter_ID0, "ID0", 0, False)
    Set RpCol = .Add(Ter_ID2, "ID2", 0, False)
    Set RpCol = .Add(Ter_IDR, "IDR", 0, False)
    Set RpCol = .Add(Ter_IDSer, "IDSer", 0, False)
    Set RpCol = .Add(Ter_Icon, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Calendar_Day
    End With
    Set RpCol = .Add(Ter_Aufgabe, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Mail_Close
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_Status, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Pin_Gray
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_VonDat, "Startdatum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Ter_BisDat, "BisDat", 0, False)
    Set RpCol = .Add(Ter_ZeiVon, "Von", 0, True)
    Set RpCol = .Add(Ter_ZeiBis, "Bis", 0, True)
    Set RpCol = .Add(Ter_ZeiVor, "Entl.", 0, False) 'Ter_ZeiEnt
    Set RpCol = .Add(Ter_Priorität, "Prio.", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ter_Vorwarn, "Vorwarn", 0, False)
    Set RpCol = .Add(Ter_Farbe, "Farbe", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ter_Anzahl, "Anzahl", 0, False)
    Set RpCol = .Add(Ter_Abgehakt, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Check
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_Erledigt, "Erledigt", 0, True)
    Set RpCol = .Add(Ter_Patient, "Patient", 0, True)
    Set RpCol = .Add(Ter_IDKurz, "Betreff", 0, True)
    Set RpCol = .Add(Ter_Datei, "Datei", 0, False)
    Set RpCol = .Add(Ter_Datum, "Hinzugefügt", 0, False)
    Set RpCol = .Add(Ter_Change, "Geändert", 0, False)
    Set RpCol = .Add(Ter_Farbtyp, "Status", 0, False)
    Set RpCol = .Add(Ter_Folge, "Folge", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Ter_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Ter_Raum, "Raum", 0, True)
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        If GlRaV = True Then
            For AktZa = 1 To UBound(GlRmu)
                .EditOptions.Constraints.Add GlRmu(AktZa, 1), GlRmu(AktZa, 2)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(Ter_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Ter_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Ter_Wiederholung, "Wiederholung", 0, False)
    Set RpCol = .Add(Ter_Selekt, "G", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Editable = False
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_Wochentag, "Tag", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    Set RpCol = .Add(Ter_MasTer, "Serie", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_AbrKom, "Abgerechnet", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_TerBet, "Terminbetrag", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_Monat, "Monat", 0, False)
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
    Set RpCol = .Add(Ter_SerBet, "Serienbetrag", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_BezBet, "Bezahlt", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_BezBet2, "Bezahlt2", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_BetOff, "Offen", 0, True)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Ter_Fallig1, "Fälligkeit", 0, False)
    Set RpCol = .Add(Ter_Fallig2, "Fälligkeit2", 0, False)
    Set RpCol = .Add(Ter_Passiv, vbNullString, 0, False)
    If GlBut = RibTab_Ter_Warte Then
        Set RpCol = .Add(Ter_WartZim, "WaZi", 0, False)
    Else
        Set RpCol = .Add(Ter_WartZim, vbNullString, 0, False)
    End If
    Set RpCol = .Add(Ter_NotiSeti, "Erinnerungswunsch", 0, False)
    Set RpCol = .Add(Ter_NotiSend, "Erinnerungsausgang", 0, False)
    Set RpCol = .Add(Ter_OnlBook, "Onlinebuchung", 0, False)
    Set RpCol = .Add(Ter_OnlSync, "Synchronisation", 0, False)
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

Select Case GlBut
Case RibTab_Ter_Listen:
    RpCls(Ter_ID0).Width = 0
    RpCls(Ter_ID2).Width = 0
    RpCls(Ter_IDR).Width = 0
    RpCls(Ter_IDSer).Width = 0
    RpCls(Ter_Icon).Width = 20
    RpCls(Ter_Aufgabe).Width = 20
    RpCls(Ter_Status).Width = 20
    RpCls(Ter_VonDat).Width = 80
    RpCls(Ter_BisDat).Width = 0
    RpCls(Ter_ZeiVon).Width = 60
    RpCls(Ter_ZeiBis).Width = 60
    RpCls(Ter_ZeiVor).Width = 0
    RpCls(Ter_Priorität).Width = 40
    RpCls(Ter_Vorwarn).Width = 0
    RpCls(Ter_Farbe).Width = 0
    RpCls(Ter_Anzahl).Width = 0
    RpCls(Ter_Abgehakt).Width = 20
    RpCls(Ter_Erledigt).Width = 0
    RpCls(Ter_Patient).Width = 200
    RpCls(Ter_IDKurz).Width = 180
    RpCls(Ter_Datei).Width = 0
    RpCls(Ter_Datum).Width = 120
    RpCls(Ter_Change).Width = 120
    RpCls(Ter_Farbtyp).Width = 0
    RpCls(Ter_Folge).Width = 60
    RpCls(Ter_IDP).Width = 180
    RpCls(Ter_IDM).Width = 180
    RpCls(Ter_Raum).Width = 110
    RpCls(Ter_GuiID).Width = 0
    RpCls(Ter_Kommentar).Width = 0
    RpCls(Ter_Wiederholung).Width = 0
    RpCls(Ter_Selekt).Width = 20
    RpCls(Ter_Wochentag).Width = 30
    RpCls(Ter_MasTer).Width = 60
    RpCls(Ter_AbrKom).Width = 150
    RpCls(Ter_TerBet).Width = 80
    RpCls(Ter_Monat).Width = 0
    RpCls(Ter_SerBet).Width = 80
    RpCls(Ter_BezBet).Width = 0
    RpCls(Ter_BezBet2).Width = 0
    RpCls(Ter_BetOff).Width = 0
    RpCls(Ter_Fallig1).Width = 0
    RpCls(Ter_Fallig2).Width = 0
    RpCls(Ter_Passiv).Width = 0
    RpCls(Ter_WartZim).Width = 0
    RpCls(Ter_NotiSeti).Width = 120
    RpCls(Ter_NotiSend).Width = 120
    RpCls(Ter_OnlBook).Width = 120
    RpCls(Ter_OnlSync).Width = 120
Case RibTab_Ter_Akont:
    RpCls(Ter_ID0).Width = 0
    RpCls(Ter_ID2).Width = 0
    RpCls(Ter_IDR).Width = 0
    RpCls(Ter_IDSer).Width = 0
    RpCls(Ter_Icon).Width = 20
    RpCls(Ter_Aufgabe).Width = 20
    RpCls(Ter_Status).Width = 20
    RpCls(Ter_VonDat).Width = 80
    RpCls(Ter_BisDat).Width = 0
    RpCls(Ter_ZeiVon).Width = 0
    RpCls(Ter_ZeiBis).Width = 0
    RpCls(Ter_ZeiVor).Width = 0
    RpCls(Ter_Priorität).Width = 40
    RpCls(Ter_Vorwarn).Width = 0
    RpCls(Ter_Farbe).Width = 0
    RpCls(Ter_Anzahl).Width = 0
    RpCls(Ter_Abgehakt).Width = 20
    RpCls(Ter_Erledigt).Width = 0
    RpCls(Ter_Patient).Width = 200
    RpCls(Ter_IDKurz).Width = 180
    RpCls(Ter_Datei).Width = 0
    RpCls(Ter_Datum).Width = 120
    RpCls(Ter_Change).Width = 120
    RpCls(Ter_Farbtyp).Width = 0
    RpCls(Ter_Folge).Width = 0
    RpCls(Ter_IDP).Width = 180
    RpCls(Ter_IDM).Width = 180
    RpCls(Ter_Raum).Width = 110
    RpCls(Ter_GuiID).Width = 0
    RpCls(Ter_Kommentar).Width = 0
    RpCls(Ter_Wiederholung).Width = 0
    RpCls(Ter_Selekt).Width = 0
    RpCls(Ter_Wochentag).Width = 0
    RpCls(Ter_MasTer).Width = 60
    RpCls(Ter_AbrKom).Width = 0
    RpCls(Ter_TerBet).Width = 80
    RpCls(Ter_Monat).Width = 0
    RpCls(Ter_SerBet).Width = 80
    RpCls(Ter_BezBet).Width = 80
    RpCls(Ter_BezBet2).Width = 0
    RpCls(Ter_BetOff).Width = 80
    RpCls(Ter_Fallig1).Width = 80
    RpCls(Ter_Fallig2).Width = 80
    RpCls(Ter_Passiv).Width = 0
    RpCls(Ter_WartZim).Width = 0
    RpCls(Ter_NotiSeti).Width = 0
    RpCls(Ter_NotiSend).Width = 0
    RpCls(Ter_OnlBook).Width = 120
    RpCls(Ter_OnlSync).Width = 120
Case RibTab_Ter_Warte:
    RpCls(Ter_ID0).Width = 0
    RpCls(Ter_ID2).Width = 0
    RpCls(Ter_IDR).Width = 0
    RpCls(Ter_IDSer).Width = 0
    RpCls(Ter_Icon).Width = 20
    RpCls(Ter_Aufgabe).Width = 20
    RpCls(Ter_Status).Width = 20
    RpCls(Ter_VonDat).Width = 0
    RpCls(Ter_BisDat).Width = 0
    RpCls(Ter_ZeiVon).Width = 60
    RpCls(Ter_ZeiBis).Width = 60
    RpCls(Ter_ZeiVor).Width = 60
    RpCls(Ter_Priorität).Width = 40
    RpCls(Ter_Vorwarn).Width = 0
    RpCls(Ter_Farbe).Width = 0
    RpCls(Ter_Anzahl).Width = 0
    RpCls(Ter_Abgehakt).Width = 20
    RpCls(Ter_Erledigt).Width = 0
    RpCls(Ter_Patient).Width = 200
    RpCls(Ter_IDKurz).Width = 180
    RpCls(Ter_Datei).Width = 0
    RpCls(Ter_Datum).Width = 120
    RpCls(Ter_Change).Width = 120
    RpCls(Ter_Farbtyp).Width = 0
    RpCls(Ter_Folge).Width = 60
    RpCls(Ter_IDP).Width = 180
    RpCls(Ter_IDM).Width = 180
    RpCls(Ter_Raum).Width = 110
    RpCls(Ter_GuiID).Width = 0
    RpCls(Ter_Kommentar).Width = 0
    RpCls(Ter_Wiederholung).Width = 0
    RpCls(Ter_Selekt).Width = 20
    RpCls(Ter_Wochentag).Width = 30
    RpCls(Ter_MasTer).Width = 60
    RpCls(Ter_AbrKom).Width = 150
    RpCls(Ter_TerBet).Width = 80
    RpCls(Ter_Monat).Width = 0
    RpCls(Ter_SerBet).Width = 80
    RpCls(Ter_BezBet).Width = 0
    RpCls(Ter_BezBet2).Width = 0
    RpCls(Ter_BetOff).Width = 0
    RpCls(Ter_Fallig1).Width = 0
    RpCls(Ter_Fallig2).Width = 0
    RpCls(Ter_Passiv).Width = 0
    RpCls(Ter_WartZim).Width = 0
    RpCls(Ter_NotiSeti).Width = 0
    RpCls(Ter_NotiSend).Width = 0
    RpCls(Ter_OnlBook).Width = 120
    RpCls(Ter_OnlSync).Width = 120
End Select
'Keine INI Speicherung, da sonst die Gruppierung nicht richtig funktioniert

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa9 " & Err.Number
Resume Next

End Sub
Private Sub SSpLa9a()
On Error GoTo SpErr

Dim SpStr As String
Dim AktZa As Integer
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo6 = FM.repCont6
Set RpCls = RpCo6.Columns

With RpCo6
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

 With RpCo6
    .AllowEdit = False
    .EditOnClick = False
    .EditOnDoubleClick = False
    .MultipleSelection = False
    .PaintManager.FixedRowHeight = True
    .PaintManager.SetPreviewIndent 112, -2, 10, 6
    .PaintManager.UseAlternativeBackground = False
    .ShowFooter = False
End With

With RpCls
    Set RpCol = .Add(War_ID0, "ID0", 0, False)
    Set RpCol = .Add(War_ID2, "ID2", 0, False)
    Set RpCol = .Add(War_IDR, "IDR", 0, False)
    Set RpCol = .Add(War_IDSer, "IDSer", 0, False)
    Set RpCol = .Add(War_Icon, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Calendar_Day
    End With
    Set RpCol = .Add(War_Aufgabe, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Mail_Close
        .Tag = 1
    End With
    Set RpCol = .Add(Ter_Status, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Pin_Gray
        .Tag = 1
    End With
    Set RpCol = .Add(War_ZeiAN, "Aufn.", 0, True)
    Set RpCol = .Add(War_ZeiVon, "Von", 0, True)
    Set RpCol = .Add(War_ZeiBis, "Bis", 0, True)
    Set RpCol = .Add(War_ZeiAB, "Entl.", 0, True)
    Set RpCol = .Add(War_Priorität, "Prio.", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(War_Farbe, "Farbe", 0, False)
    Set RpCol = .Add(War_Patient, "Patient", 0, True)
    Set RpCol = .Add(War_Farbtyp, "Farbtyp", 0, False)
    Set RpCol = .Add(War_Verzug, "Verz.", 0, False)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(War_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(War_IDM, "Mitarbeiter", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(War_Raum, "Raum", 0, True)
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        If GlRaV = True Then
            For AktZa = 1 To UBound(GlRmu)
                .EditOptions.Constraints.Add GlRmu(AktZa, 1), GlRmu(AktZa, 2)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(War_IDKurz, "Betreff", 0, True)
    Set RpCol = .Add(War_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(War_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(War_Wiederholung, "Wiederholung", 0, False)
    Set RpCol = .Add(War_Wartezimmer, "WarZim", 0, False)
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

RpCls(War_ID0).Width = 0
RpCls(War_ID2).Width = 0
RpCls(War_IDR).Width = 0
RpCls(War_IDSer).Width = 0
RpCls(War_Icon).Width = 20
RpCls(War_Aufgabe).Width = 20
RpCls(War_Status).Width = 20
RpCls(War_ZeiAN).Width = 60
RpCls(War_ZeiVon).Width = 60
RpCls(War_ZeiBis).Width = 60
RpCls(War_ZeiAB).Width = 0
RpCls(War_Priorität).Width = 40
RpCls(War_Farbe).Width = 0
RpCls(War_Patient).Width = 200
RpCls(War_Farbtyp).Width = 0
RpCls(War_Verzug).Width = 50
RpCls(War_IDP).Width = 180
RpCls(War_IDM).Width = 180
RpCls(War_Raum).Width = 110
RpCls(War_IDKurz).Width = 180
RpCls(War_GuiID).Width = 0
RpCls(War_Kommentar).Width = 0
RpCls(War_Wiederholung).Width = 0
RpCls(War_Wartezimmer).Width = 0

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo6 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLa9a " & Err.Number
Resume Next

End Sub
Private Sub SSpLaB()
On Error GoTo SpErr

Dim SpStr As String
Dim AktZa As Integer
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns

With RpCo5
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Bog_ID5, "ID5", 0, False)
    Set RpCol = .Add(Bog_ID0, "ID0", 0, False)
    Set RpCol = .Add(Bog_Datum, "Datum", 0, True)
    Set RpCol = .Add(Bog_ID3, "Fragebogen", 0, True)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        If GlBoV > 0 Then 'Fragebogen vorhanden
            For AktZa = 1 To GlBoV
                RpCol.EditOptions.Constraints.Add GlFrB(AktZa, 1), GlFrB(AktZa, 0)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(Bog_IDKurz, "Kommentar", 0, True)
    If RpCo5.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Bog_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Bog_Name, "Name", 0, True)
    Set RpCol = .Add(Bog_Vorname, "Vorname", 0, True)
    Set RpCol = .Add(Bog_Geboren, "Geboren", 0, True)
    Set RpCol = .Add(Bog_IDP, "Mandant", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
    End With
    Set RpCol = .Add(Bog_Selekt, "Selekt", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Bog_GuiID, "Selekt", 0, False)
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

If IniGetSek(GlINI, "RpCntB1") = False Then
    RpCls(Bog_ID5).Width = 0
    RpCls(Bog_ID0).Width = 0
    RpCls(Bog_Datum).Width = 100
    RpCls(Bog_ID3).Width = 250
    RpCls(Bog_IDKurz).Width = 250
    RpCls(Bog_Kommentar).Width = 0
    RpCls(Bog_Name).Width = 150
    RpCls(Bog_Vorname).Width = 150
    RpCls(Bog_Geboren).Width = 80
    RpCls(Bog_IDP).Width = 200
    RpCls(Bog_Selekt).Width = 0
    RpCls(Bog_GuiID).Width = 0
    IniSetSek "RpCntB1"
    IniSetVal "RpCntB1", "SSpLaB", RpCo5.SaveSettings
Else
    SpStr = IniGetBig(GlINI, "RpCntB1", "SSpLaB")
    RpCo5.LoadSettings SpStr
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo5 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLaB " & Err.Number
Resume Next

End Sub
Public Sub SSpLaK()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim TreKy As String
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set RpCls = RpCo8.Columns

TreKy = Left$(GlNod, 1)

With RpCo8
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Kat_ID0, "ID0", 0, False)
    If GlNod <> vbNullString Then
        Select Case TreKy
        Case "M": Set RpCol = .Add(Kat_GOID, "Nummer", 0, False) 'Terminbetreffs
        Case "N": Set RpCol = .Add(Kat_GOID, "Nummer", 0, False) 'Fragebogen
        Case "D": Set RpCol = .Add(Kat_GOID, "Kürzel", 100, False) 'Gebührenketten
        Case "F": Set RpCol = .Add(Kat_GOID, "Kürzel", 100, False) 'Diagnoseketten
        Case "H": Set RpCol = .Add(Kat_GOID, "Kürzel", 100, False) 'Laborketten
        Case "J": Set RpCol = .Add(Kat_GOID, "Kürzel", 100, False) 'Arzneiketten
        Case "R": Set RpCol = .Add(Kat_GOID, "Kürzel", 100, False) 'Terminketten
        Case Else: Set RpCol = .Add(Kat_GOID, "Nummer", 100, False)
        End Select
    Else
        If GlBut = RibTab_Kat_Frage Then
            Set RpCol = .Add(Kat_GOID, vbNullString, 0, False)
        Else
            Set RpCol = .Add(Kat_GOID, "Nummer", 100, False)
        End If
    End If
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        If RpCo8.PaintManager.FixedRowHeight = False Then
            .Alignment = xtpAlignmentLeft Or xtpAlignmentIconTop
        Else
            .Alignment = xtpAlignmentLeft
        End If
    End With
    Set RpCol = .Add(Kat_IDKurz, "Bezeichnung", 500, False)
    If RpCo8.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
    End If
     If GlBut = RibTab_Kat_Frage Then
        RpCol.TreeColumn = True
    Else
        RpCol.TreeColumn = False
    End If
    Set RpCol = .Add(Kat_Gruppe, "Gruppe", 0, False)
    If GlNod <> vbNullString Then
        Select Case TreKy
        Case "C": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
        Case "F": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
        Case "K": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
        Case "L": Set RpCol = .Add(Kat_Preis1, "Preis", 0, False)
        Case "I": Set RpCol = .Add(Kat_Preis1, "Packung", 80, False) 'Arzneikatalog
        Case "P": Set RpCol = .Add(Kat_Preis1, "Packung", 80, False) 'Artikelkatalog
        Case "M": Set RpCol = .Add(Kat_Preis1, "Zeit", 80, False)
        Case "N": 'Fragebogen
            If GlBut = RibTab_Kat_Frage Then
                Set RpCol = .Add(Kat_Preis1, "Typ", 130, False)
            Else
                Set RpCol = .Add(Kat_Preis1, vbNullString, 0, False)
            End If
        Case "O": Set RpCol = .Add(Kat_Preis1, vbNullString, 0, False)
        Case "R": Set RpCol = .Add(Kat_Preis1, "Zeit", 80, False)
        Case Else: Set RpCol = .Add(Kat_Preis1, "Preis", 80, False)
        End Select
        If TreKy = "N" Then 'Fragebogen
            RpCol.Alignment = xtpAlignmentLeft
            RpCol.HeaderAlignment = xtpAlignmentLeft
        Else
            RpCol.Alignment = xtpAlignmentRight
            RpCol.HeaderAlignment = xtpAlignmentCenter
        End If
        Select Case TreKy
        Case "G":
            Set RpCol = .Add(Kat_Sorter, "Kennung", 100, False) 'Laborkatalog
            RpCol.HeaderAlignment = xtpAlignmentCenter
        Case "I":
            Set RpCol = .Add(Kat_Sorter, "Einzel", 80, False) 'Arzneikatalog
            RpCol.HeaderAlignment = xtpAlignmentCenter
        Case "M":
            Set RpCol = .Add(Kat_Sorter, "F", 23, False) 'Terminbetreffs
            RpCol.HeaderAlignment = xtpAlignmentCenter
        Case "P":
            Set RpCol = .Add(Kat_Sorter, "Bestand", 80, False) 'Artikelkatalog
            RpCol.HeaderAlignment = xtpAlignmentCenter
        Case Else:
            If GlBut = RibTab_Kat_Frage Then
                Set RpCol = .Add(Kat_Sorter, "Sorter", 50, False)
            End If
        End Select
        If TreKy = "A" Then 'Gebühren
            Set RpCol = .Add(Kat_Typ, "Typ", 0, False) 'Eintragstyp
        End If
    Else
        If GlBut = RibTab_Kat_Frage Then
            Set RpCol = .Add(Kat_Preis1, "Typ", 130, False)
            RpCol.Alignment = xtpAlignmentLeft
            RpCol.HeaderAlignment = xtpAlignmentLeft
            Set RpCol = .Add(Kat_Sorter, "Sorter", 50, False)
        Else
            Set RpCol = .Add(Kat_Preis1, "Preis", 80, False)
            RpCol.Alignment = xtpAlignmentRight
            RpCol.HeaderAlignment = xtpAlignmentCenter
            Set RpCol = .Add(Kat_Sorter, "Sorter", 0, False)
        End If
    End If
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Resizable = False
    RpCol.Sortable = True
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(Kat_IDKurz).AutoSize = True

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo8 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLaK " & Err.Number
Resume Next

End Sub
Public Sub SSpLaM()
On Error GoTo SpErr
'Stellt Spaltenbreiten und Spaltenposition für die Emails

Dim SpStr As String
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo0 = FM.repCont0
Set RpCls = RpCo0.Columns

With RpCo0
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(Mai_IDA, vbNullString, 0, False)
    Set RpCol = .Add(Mai_ID0, vbNullString, 0, False)
    Set RpCol = .Add(Mai_IDM, vbNullString, 0, False)
    Set RpCol = .Add(Mai_GuiID, vbNullString, 0, False)
    Set RpCol = .Add(Mai_TreKey, "TreKey", 0, False)
    Set RpCol = .Add(Mai_Priority, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentIconCenter
        .Icon = IC16_Sign_Info
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Mai_Sensitivity, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentIconCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Mai_Attachment, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentIconCenter
        .Icon = IC16_Paperclip
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Mai_Marker, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentIconCenter
        .Icon = IC16_Pin_Norm
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Mai_SenderName, "Von", 0, True)
    Set RpCol = .Add(Mai_SenderEmail, "Email", 0, True)
    Set RpCol = .Add(Mai_Subject, "Betreff", 0, True)
    Set RpCol = .Add(Mai_Empfaenger, "Empfänger", 0, True)
    Set RpCol = .Add(Mai_Mailsize, "Größe", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(Mai_MailFile, vbNullString, 0, False)
    Set RpCol = .Add(Mai_Maildate, "Datum", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(Mai_Mailtime, "Uhrzeit", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(Mai_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Mai_Spammail, "Spammail", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Mai_Gelesen, "Gelesen", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Mai_Importiert, "Importiert", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Mai_Patient, "Patient / Kommentar", 0, True)
    Set RpCol = .Add(Mai_Passiv, "Passiv", 0, False)
    RpCol.Tag = 1
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Resizable = False
    RpCol.Sortable = True
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(Mai_IDA).Width = 0
RpCls(Mai_ID0).Width = 0
RpCls(Mai_IDM).Width = 0
RpCls(Mai_GuiID).Width = 0
RpCls(Mai_TreKey).Width = 0
RpCls(Mai_Priority).Width = 20
RpCls(Mai_Sensitivity).Width = 0
RpCls(Mai_Attachment).Width = 20
RpCls(Mai_Marker).Width = 20
RpCls(Mai_SenderName).Width = 200
RpCls(Mai_SenderEmail).Width = 210
RpCls(Mai_Subject).Width = 320
RpCls(Mai_Empfaenger).Width = 180
If GlTFt.SIZE > 10 Then
    RpCls(Mai_Mailsize).Width = 90
    RpCls(Mai_MailFile).Width = 0
    RpCls(Mai_Maildate).Width = 110
    RpCls(Mai_Mailtime).Width = 80
Else
    RpCls(Mai_Mailsize).Width = 75
    RpCls(Mai_MailFile).Width = 0
    RpCls(Mai_Maildate).Width = 80
    RpCls(Mai_Mailtime).Width = 70
End If
RpCls(Mai_Kommentar).Width = 0
RpCls(Mai_Spammail).Width = 0
RpCls(Mai_Gelesen).Width = 0
RpCls(Mai_Importiert).Width = 0
RpCls(Mai_Patient).Width = 200
RpCls(Mai_Passiv).Width = 0

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo0 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLaM " & Err.Number
Resume Next

End Sub
Public Sub SSpLaN()
On Error GoTo SpErr

Dim AktZa As Integer
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCoK = FM.repContK
Set RpCls = RpCoK.Columns

With RpCoK
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCoK
    .AllowEdit = True
    .EditOnClick = GlDiB 'Direktbearbeitung
    .EditOnDoubleClick = False
    .PaintManager.FixedRowHeight = Not GlZuK 'Zeilenumbruch Krankenblatt
    .PaintManager.SetPreviewIndent 184, -2, 10, 6
    .ShowFooter = False
End With

With RpCls
    Set RpCol = .Add(Kra_ID2, "ID2", 0, False)
    Set RpCol = .Add(Kra_ID0, "ID0", 0, False)
    If GlFri = 5 Then
        Set RpCol = .Add(Kra_KatNa, "Tarif", 0, False)
    Else
        Set RpCol = .Add(Kra_KatNa, "Taxe", 0, False)
    End If
    With RpCol
        .EditOptions.SelectTextOnEdit = True
        If GlBut = RibTab_Tagesproto Then
            .Visible = False
        Else
            .Visible = GlSpK 'Katalogspalte
        End If
    End With
    Set RpCol = .Add(Kra_ID3, "ID3", 0, False)
    Set RpCol = .Add(Kra_Provision, "Format", 0, False)
    Set RpCol = .Add(Kra_ID4, "ID4", 0, False)
    Set RpCol = .Add(Kra_KrTyp, "Typ", 0, False)
    Set RpCol = .Add(Kra_IDR, "IDR", 0, False)
    Set RpCol = .Add(Kra_Datum, "Datum", 0, False)
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.SelectTextOnEdit = True
        .EditOptions.AllowEdit = True
    End With
    Set RpCol = .Add(Kra_Uhrzeit, "Uhrzeit", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.SelectTextOnEdit = True
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Kra_Typ, "Typ", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        .Visible = GlETy
        If GlTyV = True Then 'Krankenblattypen vorhanden
            For AktZa = 1 To UBound(GlKrA)
                .EditOptions.Constraints.Add GlKrA(AktZa, 1), GlKrA(AktZa, 0)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(Kra_Ziffer, "Nummer", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.SelectTextOnEdit = True
        .Visible = GlZif
        .FooterFont.Bold = True
    End With
    Set RpCol = .Add(Kra_Bezeichnung, "Bezeichnungstext", 0, True)
    With RpCol
        If RpCoK.PaintManager.FixedRowHeight = False Then
            .Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If .Editable = True Then
                .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll Or xtpEditStyleMultiline
                .TreeColumn = True
                .EditOptions.MaxLength = 14000
                .EditOptions.SelectTextOnEdit = False
            End If
        End If
    End With
    Set RpCol = .Add(Kra_Analog, "A", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Visible = False
    End With
    Set RpCol = .Add(Kra_Anz, "Anz", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.SelectTextOnEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleNumber
        .Visible = False
    End With
    Set RpCol = .Add(Kra_Faktor, "Fakt.", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.SelectTextOnEdit = True
        .Visible = False
    End With
    Set RpCol = .Add(Kra_Betrag, "Mini.", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.SelectTextOnEdit = True
        .Visible = False
        .FooterFont.Bold = True
        .FooterFont.SIZE = GlTFt.SIZE
        .FooterAlignment = xtpAlignmentRight
        .Visible = GlBeZ
    End With
    Set RpCol = .Add(Kra_GesBetrag, "Maxi.", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        .Visible = GlBeZ
    End With
    Set RpCol = .Add(Kra_WVBetrag, "Akonto", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        .Visible = GlBeZ
    End With
    Set RpCol = .Add(Kra_LaBetrag, "Einstand", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        .Visible = False
    End With
    Set RpCol = .Add(Kra_Steuersatz, "Steuer", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
        .Visible = False
    End With
    Set RpCol = .Add(Kra_Zeit, "Min.", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Kra_Einheit, "Einheit", 0, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .HeaderAlignment = xtpAlignmentCenter
        .Visible = False
    End With
    Set RpCol = .Add(Kra_Währung, "W", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .HeaderAlignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Kra_IDD, "Diagnosezuord.", 0, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        .Visible = False
    End With
    If GlMsp = True Then 'Mandantenspalte anstelle von Mitarbeiterspalte in Abrechnung
        Set RpCol = .Add(Kra_IDM, "Mandanten", 0, False)
    Else
        Set RpCol = .Add(Kra_IDM, "Mitarbeiter", 0, False)
    End If
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        .Visible = GlMiZ 'Mitarbeiterspalte Abrechnung
    End With
    Set RpCol = .Add(Kra_Gedruckt, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Check
        .Tag = 1
    End With
    Set RpCol = .Add(Kra_Selekt, "Selekt", 0, False)
    Set RpCol = .Add(Kra_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Kra_Zusatztext, "Zusatztext", 0, False)
    Set RpCol = .Add(Kra_Typname, "Typname", 0, False)
    Set RpCol = .Add(Kra_Quart, "Quartal", 0, False)
    Set RpCol = .Add(Kra_Monat, "Monat", 0, False)
    Set RpCol = .Add(Kra_Woche, "Woche", 0, False)
    Set RpCol = .Add(Kra_TagSo, "Datum", 0, False)
    Set RpCol = .Add(Kra_Lock, "Lock", 0, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.Icon = IC16_Lock
    RpCol.Tag = 1
    Set RpCol = .Add(Kra_Storniert, "Storniert", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Groupable = True
    RpCol.Sortable = False
    RpCol.AutoSize = False
    RpCol.Resizable = False
    RpCol.AutoSortWhenGrouped = False
Next RpCol

If GlKFt.SIZE > 10 Then
    RpCls(Kra_Datum).Width = 110
    RpCls(Kra_Typ).Width = 50
Else
    RpCls(Kra_Datum).Width = 80
    RpCls(Kra_Typ).Width = 30
End If
RpCls(Kra_Bezeichnung).AutoSize = True
RpCls(Kra_GesBetrag).Editable = False
RpCls(Kra_WVBetrag).Editable = False
RpCls(Kra_Gedruckt).Editable = False

RpCls(Kra_KatNa).Width = 60
RpCls(Kra_Uhrzeit).Width = 60
RpCls(Kra_Ziffer).Width = 80
RpCls(Kra_Anz).Width = 35
RpCls(Kra_Faktor).Width = 45
If GlKFt.SIZE > 10 Then
    RpCls(Kra_Betrag).Width = 80
    RpCls(Kra_GesBetrag).Width = 80
    RpCls(Kra_WVBetrag).Width = 80
    RpCls(Kra_Steuersatz).Width = 80
    RpCls(Kra_Zeit).Width = 0
Else
    RpCls(Kra_Betrag).Width = 60
    RpCls(Kra_GesBetrag).Width = 60
    RpCls(Kra_WVBetrag).Width = 60
    RpCls(Kra_Steuersatz).Width = 60
    RpCls(Kra_Zeit).Width = 0
End If
RpCls(Kra_Einheit).Width = 40
RpCls(Kra_IDD).Width = 0
RpCls(Kra_IDM).Width = 80
RpCls(Kra_Gedruckt).Width = 18
RpCls(Kra_Lock).Width = 18

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCoK = Nothing

If GlSta = False Then
    S_KrSpk
End If

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLaN " & Err.Number
Resume Next

End Sub
Public Sub SSpLaT()
On Error GoTo SpErr

Dim AktZa As Integer
Dim RpCoT As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCoT = FM.repContT
Set RpCls = RpCoT.Columns

With RpCoT
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

With RpCls
    Set RpCol = .Add(0, vbNullString, 30, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentIconCenter
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        Set RpCol = .Add(1, "Mitarbeiterfilter", 100, True)
    Else
        Set RpCol = .Add(1, "Mandantfilter", 100, True)
    End If
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
        .AutoSize = True
    End With
    If RpCoT.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(2, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentIconCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCoT = Nothing

If GlSta = False Then
    S_KrSpk
End If

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpLaT " & Err.Number
Resume Next

End Sub

Public Sub SSpLaS1()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim Rpc10 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmStart
Set Rpc10 = FM.repCon10
Set RpCls = Rpc10.Columns

Select Case GlUb1
Case "N1":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "Geburtstage heute", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc10.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(2, "Alter", 60, False)
        With RpCol
            .Alignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N2":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Uhrzeit", 100, False)
        With RpCol
            .Alignment = xtpAlignmentIconLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Termine heute", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc10.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(4, "Farbe", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(5, "Datum", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N3":
    With RpCls
        Set RpCol = .Add(0, "ID1", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "Rechnung", 110, False)
         With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(2, "Patient", 200, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = True
            .AutoSize = True
        End With
        If Rpc10.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            RpCol.AutoSize = True
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(3, "Datum", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(4, "Fällig", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(5, "M", 20, False)
        With RpCol
            .Alignment = xtpAlignmentCenter
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(6, "Betrag", 0, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(7, "Offen", 70, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(8, "Mahnfrist", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(9, "Mahnbar", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(10, "Beleg", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(11, "Selekt", 0, False)
        RpCol.Tag = 1
    End With
Case "N4":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Datum", 75, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Uhrzeit", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(4, "Aufgaben", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc10.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(5, "Mitarbeiter", 60, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            If GlMiV = True Then
                For AktZa = 1 To UBound(GlMiK)
                    RpCol.EditOptions.Constraints.Add GlMiK(AktZa, 1), GlMiK(AktZa, 2)
                Next AktZa
            End If
        End With
        Set RpCol = .Add(6, vbNullString, 20, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentIconCenter
            .Icon = IC16_Check
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N5":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Datum", 92, False)
        With RpCol
            .Alignment = xtpAlignmentIconLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Termin", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc10.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(4, "Serienbetrag", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(5, "Bezahlt", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(6, "Bezahlt2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(7, "Offen", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(8, "Fälligkeit", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
        End With
        Set RpCol = .Add(9, "Fällig2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
        End With
        Set RpCol = .Add(10, "Farbe", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set Rpc10 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLaS1 " & Err.Number
Resume Next

End Sub
Public Sub SSpLaS2()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim Rpc11 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmStart
Set Rpc11 = FM.repCon11
Set RpCls = Rpc11.Columns

Select Case GlUb2
Case "N1":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "Geburtstage heute", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc11.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(2, "Alter", 60, False)
        With RpCol
            .Alignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N2":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Uhrzeit", 100, False)
        With RpCol
            .Alignment = xtpAlignmentIconLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Termine heute", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc11.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(4, "Farbe", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(5, "Datum", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N3":
    With RpCls
        Set RpCol = .Add(0, "ID1", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "Rechnung", 110, False)
         With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(2, "Patient", 200, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = True
            .AutoSize = True
        End With
        If Rpc11.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            RpCol.AutoSize = True
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(3, "Datum", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(4, "Fällig", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(5, "M", 20, False)
        With RpCol
            .Alignment = xtpAlignmentCenter
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(6, "Betrag", 0, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(7, "Offen", 70, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(8, "Mahnfrist", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(9, "Mahnbar", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(10, "Beleg", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(11, "Selekt", 0, False)
        RpCol.Tag = 1
    End With
Case "N4":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Datum", 75, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Uhrzeit", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(4, "Aufgaben", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc11.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(5, "Mitarbeiter", 60, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            If GlMiV = True Then
                For AktZa = 1 To UBound(GlMiK)
                    RpCol.EditOptions.Constraints.Add GlMiK(AktZa, 1), GlMiK(AktZa, 2)
                Next AktZa
            End If
        End With
        Set RpCol = .Add(6, vbNullString, 20, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentIconCenter
            .Icon = IC16_Check
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N5":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Datum", 92, False)
        With RpCol
            .Alignment = xtpAlignmentIconLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Termin", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc11.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(4, "Serienbetrag", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(5, "Bezahlt", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(6, "Bezahlt2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(7, "Offen", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(8, "Fälligkeit", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
        End With
        Set RpCol = .Add(9, "Fällig2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
        End With
        Set RpCol = .Add(10, "Farbe", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set Rpc11 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLaS2 " & Err.Number
Resume Next

End Sub
Public Sub SSpLaS3()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim Rpc12 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmStart
Set Rpc12 = FM.repCon12
Set RpCls = Rpc12.Columns

Select Case GlUb3
Case "N1":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "Geburtstage heute", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc12.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(2, "Alter", 60, False)
        With RpCol
            .Alignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N2":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Uhrzeit", 100, False)
        With RpCol
            .Alignment = xtpAlignmentIconLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Termine heute", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc12.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(4, "Farbe", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(5, "Datum", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N3":
    With RpCls
        Set RpCol = .Add(0, "ID1", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "Rechnung", 110, False)
         With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(2, "Patient", 200, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = True
            .AutoSize = True
        End With
        If Rpc12.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            RpCol.AutoSize = True
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(3, "Datum", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(4, "Fällig", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(5, "M", 20, False)
        With RpCol
            .Alignment = xtpAlignmentCenter
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(6, "Betrag", 0, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(7, "Offen", 70, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(8, "Mahnfrist", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(9, "Mahnbar", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(10, "Beleg", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(11, "Selekt", 0, False)
        RpCol.Tag = 1
    End With
Case "N4":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Datum", 75, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Uhrzeit", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(4, "Aufgaben", 100, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc12.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(5, "Mitarbeiter", 60, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            If GlMiV = True Then
                For AktZa = 1 To UBound(GlMiK)
                    RpCol.EditOptions.Constraints.Add GlMiK(AktZa, 1), GlMiK(AktZa, 2)
                Next AktZa
            End If
        End With
        Set RpCol = .Add(6, vbNullString, 20, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentIconCenter
            .Icon = IC16_Check
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case "N5":
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
         With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Datum", 92, False)
        With RpCol
            .Alignment = xtpAlignmentIconLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Termin", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
            .AutoSize = True
        End With
        If Rpc12.PaintManager.FixedRowHeight = False Then
            RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
            If RpCol.Editable = True Then
                RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
            End If
        End If
        Set RpCol = .Add(4, "Serienbetrag", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(5, "Bezahlt", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(6, "Bezahlt2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(7, "Offen", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(8, "Fälligkeit", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
        End With
        Set RpCol = .Add(9, "Fällig2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = True
            .Sortable = False
        End With
        Set RpCol = .Add(10, "Farbe", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set Rpc12 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLaS3 " & Err.Number
Resume Next

End Sub
Public Sub SSpLaX()
On Error GoTo SpErr
'Formatiert die Daten im GridEx

Dim AktZa As Integer
Dim RpCo9 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo9 = FM.repCont9
Set RpCls = RpCo9.Columns

With RpCls
    Set RpCol = .Add(0, "ID0", 0, False)
    Set RpCol = .Add(1, "Adressen", 0, True)
    If RpCo9.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(2, vbNullString, 20, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Printer_Ink
    End With
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Resizable = False
    RpCol.Sortable = False
    RpCol.AutoSortWhenGrouped = False
Next RpCol

RpCls(1).AutoSize = True

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo9 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSpLaX " & Err.Number
Resume Next

End Sub
Public Sub SSplS1()
On Error GoTo SpErr
'Splashscreen schließen

Dim Popu4 As XtremeSuiteControls.PopupControl

If GlApp = False Then 'AppMode
    If GloSp = False Then
        If GlRDP = True Then
            Set FM = frmSplashR
            TimEnde 4
            TimInit 5, 1
        ElseIf GlStF = True Then 'Startform zeigen
            Set FM = frmSplash
            Set Popu4 = FM.popCont4
            Popu4.Hide
            Set Popu4 = Nothing
            TimEnde 4
            TimInit 5, 1
        End If
    End If
End If

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SSplS1 " & Err.Number
Resume Next

End Sub
Public Sub SSplS2()
On Error GoTo SpErr
'Splashscreen schließen

If GlApp = False Then 'AppMode
    If GloSp = False Then
        TimEnde 5
        If GlRDP = True Then
            Unload frmSplashR
        Else
            Unload frmSplash
        End If
    End If
End If

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSplS2 " & Err.Number
Resume Next

End Sub
Public Sub SSpSav()
On Error GoTo SpErr
'Speichert die Einstellungen des GridEx

Dim DocPa As XtremeDockingPane.DockingPane
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmMain
Set DocPa = FM.dcpDoc01
Set RpCo0 = FM.repCont0
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5

If GlIdi = False Then 'Idiotenmodus
    IniSetVal "DocPa3", "PanLay", DocPa.SaveStateToString
    
    Select Case GlBut
    Case RibTab_Adressen:
            IniSetVal "RpCnt1a", "SSpLa1", RpCo2.SaveSettings
    Case RibTab_Mandanten:
            IniSetVal "RpCnt1a", "SSpLa1", RpCo2.SaveSettings
    Case RibTab_Verordner:
            IniSetVal "RpCnt1a", "SSpLa1", RpCo2.SaveSettings
    Case RibTab_Mitarbeit:
            IniSetVal "RpCnt1a", "SSpLa1", RpCo2.SaveSettings
    Case RibTab_Abrechnung:
            IniSetVal "SplRp8", "SSpLa8", RpCo3.SaveSettings
    Case RibTab_Rechnungen:
            IniSetVal "SplRp2", "SSpLa2", RpCo4.SaveSettings
    Case RibTab_Mahnwesen:
            IniSetVal "RpCnt3b", "SSpLa3", RpCo1.SaveSettings
    Case RibTab_Buchungen:
            IniSetVal "Rp4_01", "SSpLa4", RpCo1.SaveSettings
    Case RibTab_HomeBanki:
            IniSetVal "RpCnt7a", "SSpLa4a", RpCo1.SaveSettings
    Case RibTab_Ter_Listen:
            IniSetVal "RpCnt9", "SSpLa9", RpCo1.SaveSettings
    Case RibTab_Ter_Akont:
            IniSetVal "RpCnt9", "SSpLa9", RpCo1.SaveSettings
    Case RibTab_Ter_Warte:
            IniSetVal "RpCnt9", "SSpLa9", RpCo1.SaveSettings
    Case RibTab_LabBericht:
            IniSetVal "RpCnt5", "SSpLa5a", RpCo5.SaveSettings
    Case RibTab_LabBerichte:
            IniSetVal "RpCnt6", "SSpLa5e", RpCo1.SaveSettings
    Case RibTab_LabAuftrag:
            IniSetVal "RpLab1", "SSpLa5c", RpCo5.SaveSettings
    Case RibTab_LabAuftrage:
            IniSetVal "RpLab2", "SSpLa5f", RpCo1.SaveSettings
    End Select
End If

Set RpCo0 = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set DocPa = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "SSpSav " & Err.Number
Resume Next

End Sub
Public Sub TSEDisa()
On Error GoTo SuErr

Dim InTyp As String
Dim SgStr As String
Dim DaNam As String
Dim DaIni As String
Dim CoIni As String
Dim PrNam As String
Dim PaStr As String
Dim TmpSt As String
Dim TmZei As String
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim DisOK As Boolean
Dim AryZe() As String
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

Set clFil = New clsFile

PrNam = App.Path & "\smtse.exe"
CoIni = App.Path & "\smtse.ini"

Screen.MousePointer = vbHourglass
DoEvents

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smtse.exe"
    If clFil.FilVor(PrNam) = False Then
        TxTSE.Text = vbCrLf & "Eine Programmdatei wurde nicht gefunden!"
        Exit Sub
    End If
End If

If clFil.FilVor(CoIni) = False Then
    CoIni = clFil.FilVer(22) & "\SimpliMed\smtse.ini"
    If clFil.FilVor(CoIni) = False Then
        SPopu "Konfigurationsdatei nicht gefunden!", "Die Datei smtse.ini konnte nicht gefunden werden.", IC48_Forbidden
        Exit Sub
    End If
End If

PrNam = Chr$(34) & PrNam & Chr$(34)

If GlTSN <> vbNullString Then 'TSE Kennung
    If GlTSK <> vbNullString Then 'TSE Organisation Key
        DaNam = CreateID("T") & ".ini"
        DaIni = GlTmp & DaNam
        
        PaStr = "disabletss" & Space$(1) & GlTSK & Space$(1) & GlTSS & Space$(1) & GlTSN & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "state": If LCase(Right$(TmZei, Lange - Posit)) = "disabled" Then DisOK = True
                                Case "signaturecounter": TxTSE.Text = TxTSE.Text & vbCrLf & "Anzahl Signaturen : " & Right$(TmZei, Lange - Posit)
                                Case "transactioncounter": TxTSE.Text = TxTSE.Text & vbCrLf & "Anzahl Transaktionen : " & Right$(TmZei, Lange - Posit)
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
            DoEvents
            If GlLog = False Then 'General Logging
                With clFil
                    .DaLoe = GlTmp & "*.ini" & vbNullChar
                    .FilLoe
                End With
            End If
        End With
        
        If DisOK = True Then
            TxTSE.Text = TxTSE.Text & vbCrLf & "Die TSE wurde erfolgreich deaktiviert!" & vbCrLf
            DoEvents
            TxTSE.SelStart = Len(TxTSE.Text)
        End If

        DoEvents
        If GlLog = False Then 'General Logging
            With clFil
                .DaLoe = GlTmp & "*.ini" & vbNullChar
                .FilLoe
            End With
        Else
            Clipboard.Clear
            Clipboard.SetText PrNam & Space$(1) & PaStr
        End If
    End If
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSEDisa " & Err.Number
Resume Next

End Sub
Public Sub TSEExp()
On Error GoTo OrErr

Select Case GlTSe
Case 1: TSEExp1
Case 2: TSEExp2
End Select

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSEExp " & Err.Number
Resume Next

End Sub
Private Sub TSEExp1()
On Error GoTo SuErr

Dim GesZa As Long
Dim RetWe As Long
Dim FiNam As String
Dim NeNam As String
Dim Gefun As Boolean

Set FM = frmMain
Set CoDia = FM.comDialo

Set clFil = New clsFile

With CoDia
    .CancelError = True
    .DialogStyle = 1
    .DialogTitle = "Wohin sollen die Dateien exportiert werden?"
    .FileName = GlEPf
    RetWe = .ShowBrowseFolder
    FiNam = .FileName
    If RetWe = 0 Then Exit Sub
End With

NeNam = FiNam & "\_TAR_Export"

Gefun = clFil.FilVor(GLTSL & "TSE_COMM.DAT")
DoEvents

If Gefun = True Then
    Screen.MousePointer = vbHourglass
    DoEvents

    If clFil.FilDir(NeNam & "\") = False Then
        MkDir NeNam & "\"
        DoEvents
    End If

    GesZa = TSE_TAR(NeNam)
    DoEvents
    
    SPopu "TAR Export", "Es wurden : " & GesZa & " Vorgänge exportiert.", IC48_Information

    DoEvents
    Screen.MousePointer = vbNormal
End If

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSEExp1 " & Err.Number
Resume Next

End Sub


Private Sub TSEExp2()
On Error GoTo SuErr

Dim InTyp As String
Dim SgStr As String
Dim DaNam As String
Dim ExNam As String
Dim DaIni As String
Dim CoIni As String
Dim PrNam As String
Dim PaStr As String
Dim TmpSt As String
Dim TmZei As String
Dim FiNam As String
Dim AnwNa As String
Dim PfaNa As String
Dim AktZe As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim AryZe() As String
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmMain
Set CoDia = FM.comDialo

Set clFil = New clsFile

PrNam = App.Path & "\smtse.exe"
CoIni = App.Path & "\smtse.ini"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smtse.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Sub
    End If
End If

If clFil.FilVor(CoIni) = False Then
    CoIni = clFil.FilVer(22) & "\SimpliMed\smtse.ini"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Konfigurationsdatei nicht gefunden!", "Die Datei smtse.ini konnte nicht gefunden werden.", IC48_Forbidden
        Exit Sub
    End If
End If

PrNam = Chr$(34) & PrNam & Chr$(34)

AnwNa = SNaFi(GlMan(GlSMa, 1), True)
AnwNa = Replace(AnwNa, Space$(1), vbNullString, 1)
ExNam = AnwNa & "_" & Format$(DatePart("d", Date, vbMonday), "00") & Format$(DatePart("m", Date, vbMonday), "00") & DatePart("yyyy", Date, vbMonday) & ".tar"

If GlTSN <> vbNullString Then 'TSE Kennung
    If GlTSK <> vbNullString Then 'TSE Organisation Key
    
        With CoDia
            .CancelError = True
            .DialogStyle = 1
            .DefaultExt = "*.tar"
            .DialogTitle = "Bitte Name und Ordner der Exportdatei angeben"
            .FileName = GlEPf & ExNam
            .Filter = "TAR Dateien (*.tar)|*.tar|Alle Dateien (*.*)|*.*"
            .InitDir = GlEPf
            .ShowSave
            FiNam = .FileName
            If .FileTitle = vbNullString Then
                Set CoDia = Nothing
                Set clFil = Nothing
                Exit Sub
            End If
        End With
        
        If GlLog = False Then 'General Logging
            With clFil
                .FilPfa FiNam
                If .FilVor(FiNam) = True Then
                    .DaLoe = FiNam & vbNullChar
                    .FilLoe
                End If
            End With
        End If
        
        If IsNull(FiNam) = False And FiNam <> vbNullString Then
            Screen.MousePointer = vbHourglass
            DoEvents
        
            DaNam = CreateID("T") & ".ini"
            DaIni = GlTmp & DaNam
            
            PaStr = "createexport" & Space$(1) & GlTSK & Space$(1) & GlTSS & Space$(1) & GlTSN & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--tar=" & Chr$(34) & FiNam & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
            WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
            DoEvents

            With clFil
                If .FilVor(DaIni) = True Then
                    .FilPfa DaIni
                    TmpSt = .FilReSt
                    DoEvents
                    If TmpSt <> vbNullString Then
                        AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                        For AktZe = 0 To UBound(AryZe) - 1
                            If AryZe(AktZe) <> vbNullString Then
                                TmZei = AryZe(AktZe)
                                Lange = Len(TmZei)
                                Posit = InStr(1, TmZei, "=", 1)
                                If Posit > 0 Then
                                        InTyp = LCase(Left$(TmZei, Posit - 1))
                                        Select Case InTyp
                                        Case "state":
                                        If LCase(Right$(TmZei, Lange - Posit)) = "completed" Then
                                            SPopu "TAR Export erfolgreich", "Die TAT Dateien wurden exportiert.", IC48_Information
                                        End If
                                    End Select
                                End If
                            End If
                        Next AktZe
                    End If
                End If
                DoEvents
                If GlLog = False Then 'General Logging
                    With clFil
                        .DaLoe = GlTmp & "*.ini" & vbNullChar
                        .FilLoe
                    End With
                Else
                    Clipboard.Clear
                    Clipboard.SetText PrNam & Space$(1) & PaStr
                End If
            End With

            DoEvents
            Screen.MousePointer = vbNormal
        End If
    End If
End If

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSEExp2 " & Err.Number
Resume Next

End Sub
Public Function TSEGen(ByVal BeGes As Double, Steue As Single, ByVal KliNa As String, Optional ByVal BeSt1 As Double, Optional ByVal BeSt2 As Double, Optional ByVal BeSt3 As Double) As Boolean
On Error GoTo InErr
'Generiert eine fiskaly TSE Buchung

Dim BetS1 As Double
Dim BetS2 As Double
Dim BetS3 As Double
Dim PrNam As String
Dim TssNa As String
Dim DaNam As String
Dim DaIni As String
Dim CoIni As String
Dim PaStr As String
Dim TmpSt As String
Dim TmZei As String
Dim InTyp As String
Dim Betr1 As String
Dim Betr2 As String
Dim Betr3 As String
Dim Betr4 As String
Dim SgStr As String
Dim AktZe As Integer
Dim Posit As Integer
Dim PoSg1 As Integer
Dim PoSg2 As Integer
Dim Lange As Integer
Dim AryZe() As String
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set clFil = New clsFile

GlTSB = GlTSC

PrNam = App.Path & "\smtse.exe"
CoIni = App.Path & "\smtse.ini"

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smtse.exe"
    If clFil.FilVor(PrNam) = False Then
        SPopu "Programmdatei nicht gefunden!", "Eine zur Ausführung benötigte Datei konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

If clFil.FilVor(CoIni) = False Then
    CoIni = clFil.FilVer(22) & "\SimpliMed\smtse.ini"
    If clFil.FilVor(CoIni) = False Then
        SPopu "Konfigurationsdatei nicht gefunden!", "Die Datei smtse.ini konnte nicht gefunden werden.", IC48_Forbidden
        Exit Function
    End If
End If

PrNam = Chr$(34) & PrNam & Chr$(34)

BeGes = Round(BeGes, 2)
If Steue > 0 Then
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        If Steue > 6 Then
            BetS1 = BeGes
            BetS2 = 0
            BetS3 = 0
        Else
            BetS1 = 0
            BetS2 = BeGes
            BetS3 = 0
        End If
    Else
        If Steue > 10 Then
            BetS1 = BeGes
            BetS2 = 0
            BetS3 = 0
        Else
            BetS1 = 0
            BetS2 = BeGes
            BetS3 = 0
        End If
    End If
Else
    BeSt1 = Round(BeSt1, 2)
    BeSt2 = Round(BeSt2, 2)
    BeSt3 = Round(BeGes, 2)
End If

Betr1 = Replace(BeSt1, ",", ".", 1)
Betr2 = Replace(BeSt2, ",", ".", 1)
Betr3 = Replace(BeSt3, ",", ".", 1)
Betr4 = Replace(BeGes, ",", ".", 1)

If GlTSN <> vbNullString Then 'TSE Kennung
    If GlTSK <> vbNullString Then 'TSE Organisation Key
        GlTSB.ZeiSt = Format$(Date, "dd.mm.yyyy") & " " & Format$(Now, "hh:mm:ss")

        DaNam = CreateID("T") & ".ini"
        DaIni = GlTmp & DaNam
        
        PaStr = "createtrans" & Space$(1) & GlTSK & Space$(1) & GlTSS & Space$(1) & GlTSN & Space$(1) & KliNa & Space$(1) & Betr1 & Space$(1) & Betr2 & Space$(1) & Betr3 & Space$(1) & Betr4 & Space$(1) & "Cash" & Space$(1) & "1" & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents
        
        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                DoEvents
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe) - 1
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "transactionid": GlTSB.TraID = Right$(TmZei, Lange - Posit)
                                Case "number": GlTSB.TraZe = Val(CLng(Right$(TmZei, Lange - Posit)))
                                Case "startutc":
                                Case "endutc": GlTSB.ZeiEn = Right$(TmZei, Lange - Posit)
                                Case "qrcodedata": GlTSB.SigQr = Right$(TmZei, Lange - Posit)
                                Case "certificateserial": GlTSB.SigSt = Right$(TmZei, Lange - Posit)
                                Case "signature":
                                    SgStr = Right$(TmZei, Lange - Posit)
                                    PoSg1 = InStr(1, SgStr, "counter:", 1)
                                    If PoSg1 > 0 Then
                                        PoSg2 = InStr(PoSg1 + 8, SgStr, ",", 1)
                                        If PoSg2 > 0 Then
                                            If IsNumeric(Mid$(SgStr, (PoSg1 + 8), (PoSg2) - (PoSg1 + 8))) = True Then
                                                GlTSB.SigZe = CLng(Mid$(SgStr, (PoSg1 + 8), (PoSg2) - (PoSg1 + 8)))
                                            End If
                                        End If
                                    End If
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
            DoEvents
        End With
        
        If GlTSB.SigQr <> vbNullString Then
            GlTSB.ZeLog = GlTSB.ZeiSt & " - " & GlTSB.ZeiEn
            TSEGen = True
        End If
        
        If GlLog = False Then 'General Logging
            With clFil
                .DaLoe = GlTmp & "*.ini" & vbNullChar
                .FilLoe
            End With
        Else
            Clipboard.Clear
            Clipboard.SetText PrNam & Space$(1) & PaStr
        End If
    End If
End If

Set clFil = Nothing

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSEGen " & Err.Number
Resume Next

End Function
Public Sub TSEInfo()
On Error Resume Next

Dim RetWe As Long
Dim TmStr As String
Dim DaStr As String
Dim ZeStr As String
Dim AktZa As Integer
Dim StaWe As Integer
Dim BytAr(0 To 20) As Byte
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

RetWe = DirectReadDriveNT(GLTSL & "TSE_INFO.DAT", 0, 0, RuAry(), 512)

If RetWe = 0 Then
    TxTSE.Text = TxTSE.Text & vbCrLf & "TSE nicht anzusprechen!"
    Exit Sub
End If
StaWe = RuAry(32)

If (StaWe And 4) = 4 Then
    TmStr = TmStr & vbCrLf & "Time Admin Pin changed" & vbCrLf
End If

If (StaWe And 2) = 2 Then
    TmStr = TmStr & vbCrLf & "Pin changed " & vbCrLf
End If

If (StaWe And 1) = 1 Then
    TxTSE.Text = TxTSE.Text & vbCrLf & "Puk changed" & vbCrLf
End If

For AktZa = 0 To 12
    BytAr(AktZa) = RuAry(AktZa)
Next

AktZa = AktZa

TmStr = StrConv(BytAr, vbUnicode)

If RuAry(29) = 0 Then
    TxTSE.Text = TxTSE.Text & vbCrLf & "TSE nicht initialisiert!"
End If

If RuAry(29) = 1 Then
    TxTSE.Text = TxTSE.Text & vbCrLf & "TSE initialisiert!"
End If

If RuAry(29) = 2 Then
    TxTSE.Text = TxTSE.Text & vbCrLf & "TSE stillgelegt!"
End If

If RuAry(30) = 0 Then
  TxTSE.Text = TxTSE.Text & vbCrLf & "TSE Data Import nicht initialisiert"
Else
  TxTSE.Text = TxTSE.Text & vbCrLf & "TSE Data Import gestartet!"
End If

TmStr = vbNullString
StaWe = RuAry(28)

If (StaWe And 8) = 8 Then
    'TmStr = " Bit 4 Set  "
Else
    'TmStr = TmStr & " Bit 4 = 0  "
End If

If (StaWe And 4) = 4 Then
    TmStr = TmStr & vbCrLf & "CTSS Interface activ!" & vbCrLf
Else
    TmStr = TmStr & vbCrLf & "CTSS Interface inactiv!" & vbCrLf
End If

If (StaWe And 2) = 2 Then
    TmStr = TmStr & vbCrLf & "Selbsttest durchgeführt"
    DaStr = vbNullString
    For AktZa = 0 To 3
        If RuAry(36 + AktZa) > 15 Then
            DaStr = DaStr & Hex(RuAry(36 + AktZa))
        Else
            DaStr = DaStr & "0" & Hex(RuAry(36 + AktZa))
        End If
    Next
    TmStr = TmStr & " (" & Format((CLng("&H" & Mid(DaStr, 1, 16)) / 3600), "0.0") & " Std. verbleibend)"
Else
    TmStr = TmStr & vbCrLf & "Selbsttest fehlt! " & vbCrLf
End If

If (StaWe And 1) = 1 Then
    TmStr = TmStr & vbCrLf & "Zeit synchronisiert"
Else
    TmStr = TmStr & vbCrLf & "Zeit nicht synchronisiert!"
End If

DaStr = vbNullString

For AktZa = 0 To 7
    If RuAry(64 + AktZa) > 15 Then
        DaStr = DaStr & Hex(RuAry(64 + AktZa))
    Else
        DaStr = DaStr & "0" & Hex(RuAry(64 + AktZa))
    End If
Next

ZeStr = CLng("&H" & Mid(DaStr, 1, 16))
ZeStr = DateAdd("s", ZeStr, DateSerial(1970, 1, 1))

TxTSE.Text = TxTSE.Text & vbCrLf & "TSE Status gültig bis: " & ZeStr
TxTSE.Text = TxTSE.Text & vbCrLf & TmStr & vbCrLf
TxTSE.Text = TxTSE.Text & vbCrLf & Val(RuAry(43)) & " offene Transaktionen!" & vbCrLf

DaStr = vbNullString

For AktZa = 0 To 7
    If RuAry(72 + AktZa) > 15 Then
        DaStr = DaStr & Hex(RuAry(72 + AktZa))
    Else
        DaStr = DaStr & "0" & Hex(RuAry(72 + AktZa))
    End If
Next

ZeStr = CLng("&H" & Mid(DaStr, 1, 16))

TxTSE.Text = TxTSE.Text & vbCrLf & "Speicher belegt :" & Format(Val(ZeStr) / 1024, "0.0") & " kB " & Val(ZeStr / 512) & " Sectoren" & vbCrLf

End Sub
Public Sub TSESwi()
On Error GoTo SuErr
'Initialisiert und testet die SwissBit TSE

Dim RetWe As Long
Dim KlStr As String
Dim KaNam As String
Dim KaStr As String
Dim SeTes As String
Dim AktZa As Integer
Dim Lange As Integer
Dim StaWe As Integer
Dim Posit As Integer
Dim Gefun As Boolean
Dim TesOK As Boolean
Dim TsAry() As String
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

Screen.MousePointer = vbHourglass
DoEvents

Set clFil = New clsFile

For AktZa = 1 To UBound(GlGeK)
    If CBool(GlGeK(AktZa, 5)) = True Then
        If GlGeK(AktZa, 8) <> vbNullString Then
            KaNam = GlGeK(AktZa, 1)
            KaStr = GlGeK(AktZa, 8)
            Exit For
        End If
    End If
Next AktZa

If KaStr = vbNullString Then
    KaNam = S_KaNa()
    DoEvents
    TxTSE.Text = TxTSE.Text & vbCrLf & "Unter den Geldkonten existiert kein als Kasse gekennzeichneter Eintrag!" & vbCrLf
    DoEvents
    TxTSE.SelStart = Len(TxTSE.Text)
    DoEvents
    Screen.MousePointer = vbNormal
    Exit Sub
End If

Gefun = clFil.FilVor(GLTSL & "TSE_COMM.DAT")

If Gefun = False Then
    TxTSE.Text = vbCrLf & "TSE nicht gefunden!"
Else
    TxTSE.Text = "Laufwerk : " & GLTSL
    TxTSE.Text = TxTSE.Text & vbCrLf & KaNam & " : " & KaStr
    DoEvents
    TsAry = Split(get_status, ";")
    DoEvents
    If InStr(1, TsAry(0), "Fehler", 1) Then
        TxTSE.Text = TxTSE.Text & vbCrLf & TsAry(0)
        Exit Sub
    End If
    If GlSet(1, 93) = vbNullString Or Left$(GlSet(1, 93), 1) = "K" Then
        If TsAry(1) <> vbNullString Then
            GlTSN = TsAry(1) 'TSE Kennung
            DBCmEx2 "qrySetEd1", "@IdxNr", "@IdSet", 94, GlTSN
            DoEvents
        End If
    End If
    GlTSK = TsAry(2) 'TSE Key
    TxTSE.Text = TxTSE.Text & vbCrLf & "Gültigkeit : " & TsAry(0)
    TxTSE.Text = TxTSE.Text & vbCrLf & TsAry(3)
    If CLng(TsAry(4)) > 0 Then
        TxTSE.Text = TxTSE.Text & vbCrLf & "Transaktionen : " & TsAry(4)
    End If
    DoEvents
    RetWe = DirectReadDriveNT(GLTSL & "TSE_INFO.DAT", 0, 0, RuAry(), 512)
    If RetWe = 0 Then
        TxTSE.Text = TxTSE.Text & vbCrLf & "TSE nicht anzusprechen!"
    Else
        StaWe = RuAry(28)
        If (StaWe And 2) <> 2 Then
            DoEvents
            TxTSE.Text = TxTSE.Text & vbCrLf & "TSE Selbsttest bitte warten..."
            DoEvents
            SeTes = TSE_SeTe(KaStr)
            TxTSE.Text = TxTSE.Text & vbCrLf & SeTes
            Posit = InStrRev(SeTes, "fehlgeschlagen", -1, 1)
        End If
        If Posit = 0 Then
            TesOK = True
        End If
    End If
    DoEvents
    If TesOK = True Then
        TxTSE.Text = TxTSE.Text & vbCrLf & "Login : " & TSE_Login(0, 12345)
        TxTSE.Text = TxTSE.Text & vbCrLf & "Zeitsynchronisation : " & TSE_Time()
        DoEvents

        befehl.blen(0) = 0
        befehl.blen(1) = 6
        befehl.befehl(0) = &H43
        befehl.befehl(1) = &H0
        befehl.befehl(2) = 0
        befehl.befehl(3) = 0
        befehl.befehl(4) = 0
        befehl.befehl(5) = 0
        TSE_Send
        DoEvents
        
        WindowSleep 1000
        DoEvents
        
        KlStr = TSEKlie
        Lange = Len(KlStr)
        If InStr(1, KlStr, KaStr, 1) = 0 Then
            TxTSE.Text = TxTSE.Text & vbCrLf & KaNam & " wurde nicht eingerichtet!"
        Else
            TxTSE.Text = TxTSE.Text & vbCrLf & KaNam & " ist bereit."
        End If
    Else
        TxTSE.Text = TxTSE.Text & vbCrLf & KaNam & " wurde nicht eingerichtet!"
    End If
    TxTSE.Text = TxTSE.Text & vbCrLf & "Der Dialog kann geschlossen werden."
    DoEvents
    TxTSE.SelStart = Len(TxTSE.Text)
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSE_Init " & Err.Number
Resume Next

End Sub
Private Function TSEKlie() As String
On Error Resume Next

Dim RetWe As Long
Dim KlStr As String
Dim ResWe As Integer
Dim AktZa As Integer
Dim KlPru As Boolean
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

RetWe = DirectReadDriveNT(GLTSL & "TSE_COMM.DAT", 0, 0, RuAry(), 512)

If RetWe = 0 Then
    TxTSE.Text = TxTSE.Text & vbCrLf & "Der TSE Stick unter " & GLTSL & " kann nicht gefunden werden!"
    Exit Function
End If

ResWe = RuAry(7)

For AktZa = 8 To 8 + ResWe * 32
    If RuAry(AktZa) = 0 Then
        If KlPru = False Then
            'TxTSE.Text = TxTSE.Text & vbCrLf
        End If
        KlPru = True
    End If
    If RuAry(AktZa) > 0 Then
        'TxTSE.Text = TxTSE.Text & Chr$(RuAry(AktZa))
        KlStr = KlStr & Chr$(RuAry(AktZa))
        KlPru = False
    End If
Next

TSEKlie = KlStr

End Function
Public Function TSESet(ByVal GeBet As Double, ByVal Steue As Single, ByVal KliNa As String) As Boolean
On Error Resume Next
'Generiert eine SwissBit TSE Transaktion

Dim TSEBu As Boolean
Dim Gefun As Boolean

GlTSB = GlTSC

Set clFil = New clsFile

If GlTSK <> vbNullString Then
    Gefun = clFil.FilVor(GLTSL & "TSE_COMM.DAT")
    If Gefun = True Then
        If GeBet > 0 Then
            With GlTSB
                .BeTyp = "Beleg"
                If Steue > 0 Then
                    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
                        If Steue > 6 Then
                            .BeSt0 = GeBet
                        Else
                            .BeSt1 = GeBet
                        End If
                    Else
                        If Steue > 10 Then
                            .BeSt0 = GeBet
                        Else
                            .BeSt1 = GeBet
                        End If
                    End If
                Else
                    .BeSt4 = GeBet
                End If
                .BeBar = GeBet
                .BeUnb = 0
            End With
            TSEBu = TSE_Strg(KliNa)
            DoEvents
            TSESet = TSEBu
        End If
    Else
        SPopu "TSE nicht ansprechbar!", "Die TSE kann nicht gefunden werden oder es liegt eine Störung vor.", IC48_Forbidden
    End If
Else
    SPopu "Signierung nicht möglich", "Es wurde noch kein TSE Selbsttest durchgeführt!", IC48_Forbidden
End If

Set clFil = Nothing

End Function

Public Sub TSEWeb()
On Error GoTo SuErr
'Initialisiert und testet die fiskaly TSE

Dim ManNr As Long
Dim DaNam As String
Dim DaIni As String
Dim CoIni As String
Dim PrNam As String
Dim PaStr As String
Dim TssNa As String
Dim DaTab As String
Dim TssID As String
Dim KliID As String
Dim TmpSt As String
Dim TmZei As String
Dim KaNam As String
Dim TssDa As String
Dim InTyp As String
Dim MaKur As String
Dim MaStr As String
Dim MaPLZ As String
Dim MaOrt As String
Dim MaFir As String
Dim MaStu As String
Dim MaLan As String
Dim OrKey As String
Dim OrSec As String
Dim OrIde As String
Dim Posit As Integer
Dim AktZe As Integer
Dim AktZa As Integer
Dim Lange As Integer
Dim KasNr As Integer
Dim OrgOK As Boolean
Dim TSEOK As Boolean
Dim KliOK As Boolean
Dim AryZe() As String
Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

Set clFil = New clsFile

ManNr = GlMan(GlSMa, 2)

PrNam = App.Path & "\smtse.exe"
CoIni = App.Path & "\smtse.ini"

Screen.MousePointer = vbHourglass
DoEvents

If clFil.FilVor(PrNam) = False Then
    PrNam = clFil.FilVer(22) & "\SimpliMed\smtse.exe"
    If clFil.FilVor(PrNam) = False Then
        TxTSE.Text = vbCrLf & "Eine Programmdatei wurde nicht gefunden!"
        Exit Sub
    End If
End If

If clFil.FilVor(CoIni) = False Then
    CoIni = clFil.FilVer(22) & "\SimpliMed\smtse.ini"
    If clFil.FilVor(CoIni) = False Then
        SPopu "Konfigurationsdatei nicht gefunden!", "Die Datei smtse.ini konnte nicht gefunden werden.", IC48_Forbidden
        Exit Sub
    End If
End If

PrNam = Chr$(34) & PrNam & Chr$(34)

'------ Organisation Einrichten ------

If GlSet(1, 97) = vbNullString Then 'TSE Organisation
    If GlThe(GlSMa, 13) <> vbNullString Then
        If Len(GlThe(GlSMa, 13)) >= 3 Then
            If Len(GlThe(GlSMa, 13)) <= 30 Then
                MaKur = Chr$(34) & GlThe(GlSMa, 13) & Chr$(34)
            Else
                MaKur = Chr$(34) & Left$(GlThe(GlSMa, 13), 30) & Chr$(34)
            End If
        Else
            MaKur = Chr$(34) & GlThe(GlSMa, 13) & Space$(3) & Chr$(34)
        End If
    Else
        MaKur = Chr$(34) & "---" & Chr$(34)
    End If

    If GlThe(GlSMa, 3) <> vbNullString Then
        If Len(GlThe(GlSMa, 3)) >= 3 Then
            If Len(GlThe(GlSMa, 3)) <= 30 Then
                MaStr = Chr$(34) & GlThe(GlSMa, 3) & Chr$(34)
            Else
                MaStr = Chr$(34) & Left$(GlThe(GlSMa, 3), 30) & Chr$(34)
            End If
        Else
            MaStr = Chr$(34) & GlThe(GlSMa, 3) & Space$(3) & Chr$(34)
        End If
    Else
        MaStr = Chr$(34) & "---" & Chr$(34)
    End If

    If GlThe(GlSMa, 4) <> vbNullString Then
        If Len(GlThe(GlSMa, 4)) >= 3 Then
            If Len(GlThe(GlSMa, 4)) <= 30 Then
                MaPLZ = Chr$(34) & GlThe(GlSMa, 4) & Chr$(34)
            Else
                MaPLZ = Chr$(34) & Left$(GlThe(GlSMa, 4), 30) & Chr$(34)
            End If
        Else
            MaPLZ = Chr$(34) & GlThe(GlSMa, 4) & Space$(3) & Chr$(34)
        End If
    Else
        MaPLZ = Chr$(34) & "---" & Chr$(34)
    End If
    
    If GlThe(GlSMa, 5) <> vbNullString Then
        If Len(GlThe(GlSMa, 5)) >= 3 Then
            If Len(GlThe(GlSMa, 5)) <= 30 Then
                MaOrt = Chr$(34) & GlThe(GlSMa, 5) & Chr$(34)
            Else
                MaOrt = Chr$(34) & Left$(GlThe(GlSMa, 5), 30) & Chr$(34)
            End If
        Else
            MaOrt = Chr$(34) & GlThe(GlSMa, 5) & Space$(3) & Chr$(34)
        End If
    Else
        MaOrt = Chr$(34) & "---" & Chr$(34)
    End If

    If GlThe(GlSMa, 19) <> vbNullString Then
        If Len(GlThe(GlSMa, 19)) >= 3 Then
            If Len(GlThe(GlSMa, 19)) <= 30 Then
                MaFir = Chr$(34) & GlThe(GlSMa, 19) & Chr$(34)
            Else
                MaFir = Chr$(34) & Left$(GlThe(GlSMa, 19), 30) & Chr$(34)
            End If
        Else
            MaFir = Chr$(34) & "---" & Chr$(34)
        End If
    Else
        If GlThe(GlSMa, 2) <> vbNullString Then
            MaFir = GlThe(GlSMa, 1) & Space$(1) & GlThe(GlSMa, 2)
            If Len(MaFir) > 30 Then
                MaFir = Left$(MaFir, 30)
            End If
        ElseIf GlThe(GlSMa, 1) <> vbNullString Then
            MaFir = GlThe(GlSMa, 1)
            If Len(MaFir) > 30 Then
                MaFir = Left$(MaFir, 30)
            End If
        Else
            MaFir = "---"
        End If
        MaFir = Chr$(34) & MaFir & Chr$(34)
    End If

    MaStu = Chr$(34) & vbNullString & Chr$(34)
    MaLan = Chr$(34) & "DEU" & Chr$(34)
                      
    DaNam = CreateID("T") & ".ini"
    DaIni = GlTmp & DaNam
    
    PaStr = "createorg" & Space$(1) & MaFir & Space$(1) & MaStr & Space$(1) & MaPLZ & Space$(1) & MaOrt & Space$(1) & MaKur & Space$(1) & MaStu & Space$(1) & MaLan & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
    WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
    DoEvents

    With clFil
        If .FilVor(DaIni) = True Then
            .FilPfa DaIni
            TmpSt = .FilReSt
            If TmpSt <> vbNullString Then
                AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                For AktZe = 0 To UBound(AryZe)
                    If AryZe(AktZe) <> vbNullString Then
                        TmZei = AryZe(AktZe)
                        Lange = Len(TmZei)
                        Posit = InStr(1, TmZei, "=", 1)
                        If Posit > 0 Then
                            InTyp = LCase(Left$(TmZei, Posit - 1))
                            Select Case InTyp
                            Case "key":
                                OrKey = Right$(TmZei, Lange - Posit)
                                If OrKey <> vbNullString Then
                                    DBCmEx2 "qrySetEd1", "@IdxNr", "@IdSet", 98, OrKey
                                    GlSet(1, 97) = OrKey
                                    GlTSK = OrKey 'TSE Organisation Key
                                    TxTSE.Text = TxTSE.Text & vbCrLf & "Die Organisation ist jetzt eingerichtet : "
                                    TxTSE.Text = TxTSE.Text & vbCrLf & GlTSK
                                    DoEvents
                                End If
                            Case "secret":
                                OrSec = Right$(TmZei, Lange - Posit)
                                If OrSec <> vbNullString Then
                                    DBCmEx2 "qrySetEd1", "@IdxNr", "@IdSet", 99, OrSec
                                    GlSet(1, 98) = OrSec
                                    GlTSS = OrSec 'TSE Organisation Secret
                                    TxTSE.Text = TxTSE.Text & vbCrLf & GlTSS & vbCrLf
                                    DoEvents
                                End If
                            Case "organizationid": 'TSE Organisation ID
                                OrIde = Right$(TmZei, Lange - Posit)
                                If OrIde <> vbNullString Then
                                    DBCmEx2 "qrySetEd1", "@IdxNr", "@IdSet", 100, OrIde
                                    GlSet(1, 99) = OrIde
                                    GlTSI = OrIde 'TSE Organisation ID
                                    TxTSE.Text = TxTSE.Text & vbCrLf & GlTSI & vbCrLf
                                    DoEvents
                                End If
                            End Select
                        End If
                    End If
                Next AktZe
            End If
        End If
    End With
Else
    If GlThe(GlSMa, 19) <> vbNullString Then
        If Len(GlThe(GlSMa, 19)) >= 3 Then
            If Len(GlThe(GlSMa, 19)) <= 30 Then
                MaFir = Chr$(34) & GlThe(GlSMa, 19) & Chr$(34)
            Else
                MaFir = Chr$(34) & Left$(GlThe(GlSMa, 19), 30) & Chr$(34)
            End If
        Else
            MaFir = Chr$(34) & "---" & Chr$(34)
        End If
    Else
        If GlThe(GlSMa, 2) <> vbNullString Then
            MaFir = GlThe(GlSMa, 1) & Space$(1) & GlThe(GlSMa, 2)
            If Len(MaFir) > 30 Then
                MaFir = Left$(MaFir, 30)
            End If
        ElseIf GlThe(GlSMa, 1) <> vbNullString Then
            MaFir = GlThe(GlSMa, 1)
            If Len(MaFir) > 30 Then
                MaFir = Left$(MaFir, 30)
            End If
        Else
            MaFir = "---"
        End If
        MaFir = Chr$(34) & MaFir & Chr$(34)
    End If
    
    TxTSE.Text = TxTSE.Text & vbCrLf & "Die Praxis : " & MaFir & " ist bereits eingerichtet."
    DoEvents

    If GlTSK <> vbNullString Then 'TSE Organisation Key
        DaNam = CreateID("T") & ".ini"
        DaIni = GlTmp & DaNam
        PaStr = "testorg" & Space$(1) & GlTSI & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe)
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "exists": If Right$(TmZei, Lange - Posit) = "1" Then OrgOK = True
                                Case "displayname": TxTSE.Text = TxTSE.Text & vbCrLf & "Organisation : " & Right$(TmZei, Lange - Posit)
                                Case "name": TxTSE.Text = TxTSE.Text & vbCrLf & "Name : " & Right$(TmZei, Lange - Posit)
                                Case "address": TxTSE.Text = TxTSE.Text & vbCrLf & "Straße : " & Right$(TmZei, Lange - Posit)
                                Case "zipcode": TxTSE.Text = TxTSE.Text & vbCrLf & "PLZ : " & Right$(TmZei, Lange - Posit)
                                Case "town": TxTSE.Text = TxTSE.Text & vbCrLf & "Ort : " & Right$(TmZei, Lange - Posit)
                                Case "countrycode": TxTSE.Text = TxTSE.Text & vbCrLf & "Land : " & Right$(TmZei, Lange - Posit)
                                Case "key": TxTSE.Text = TxTSE.Text & vbCrLf & "Schlüssel : " & Right$(TmZei, Lange - Posit)
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
        End With
    End If
    If OrgOK = True Then
        TxTSE.Text = TxTSE.Text & vbCrLf & "Die Organisation wurde erfolgreich geprüft!" & vbCrLf
    Else
        TxTSE.Text = TxTSE.Text & vbCrLf & "Es ist keine Organisations-ID gespeichert!" & vbCrLf
    End If
End If

'------ TSE Einrichten ------

If GlSet(1, 93) = vbNullString Or Left$(GlSet(1, 93), 1) = "K" Then
    DaTab = IniGetVal("System", "DatTab") & "-" & CreateID("D")
    TssNa = Chr$(39) & DaTab & Chr$(39)
    DaNam = CreateID("T") & ".ini"
    DaIni = GlTmp & DaNam
    
    If GlTSK <> vbNullString Then 'TSE Organisation Key
        PaStr = "createtss" & Space$(1) & GlTSK & Space$(1) & GlTSS & Space$(1) & TssNa & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe)
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "tssid":
                                    TssID = Right$(TmZei, Lange - Posit)
                                    If TssID <> vbNullString Then
                                        DBCmEx2 "qrySetEd1", "@IdxNr", "@IdSet", 94, TssID
                                        GlSet(1, 93) = TssID
                                        GlTSN = TssID 'TSE Kennung
                                        TxTSE.Text = TxTSE.Text & vbCrLf & "Die TSE ist jetzt eingerichtet : "
                                        TxTSE.Text = TxTSE.Text & vbCrLf & GlTSN & vbCrLf
                                        DoEvents
                                    End If
                                Case "createdutc": TxTSE.Text = TxTSE.Text & vbCrLf & "Eingerichtet am: " & Right$(TmZei, Lange - Posit)
                                Case "initializedutc": TxTSE.Text = TxTSE.Text & vbCrLf & "Initialisiert am: " & Right$(TmZei, Lange - Posit)
                                Case "certificateserial": TxTSE.Text = TxTSE.Text & vbCrLf & "TSE Seriennummer : " & Right$(TmZei, Lange - Posit)
                                Case "publickey": TxTSE.Text = TxTSE.Text & vbCrLf & "TSE Public-Key : " & Right$(TmZei, Lange - Posit)
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
        End With
    Else
        TxTSE.Text = TxTSE.Text & vbCrLf & "Es kann keine TSS eingerichtet werden!" & vbCrLf
    End If
Else
    TssNa = GlSet(1, 93)
    DaNam = CreateID("T") & ".ini"
    DaIni = GlTmp & DaNam
    
    If GlTSK <> vbNullString Then 'TSE Organisation Key
        PaStr = "testtss" & Space$(1) & GlTSK & Space$(1) & GlTSS & Space$(1) & TssNa & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
        WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
        DoEvents

        TxTSE.Text = TxTSE.Text & vbCrLf & "Folgende TSE wurde eingerichtet : " & GlTSN
    
        With clFil
            If .FilVor(DaIni) = True Then
                .FilPfa DaIni
                TmpSt = .FilReSt
                If TmpSt <> vbNullString Then
                    AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                    For AktZe = 0 To UBound(AryZe)
                        If AryZe(AktZe) <> vbNullString Then
                            TmZei = AryZe(AktZe)
                            Lange = Len(TmZei)
                            Posit = InStr(1, TmZei, "=", 1)
                            If Posit > 0 Then
                                InTyp = LCase(Left$(TmZei, Posit - 1))
                                Select Case InTyp
                                Case "exists": If Right$(TmZei, Lange - Posit) = "1" Then TSEOK = True
                                Case "createdutc": TxTSE.Text = TxTSE.Text & vbCrLf & "Eingerichtet am: " & Right$(TmZei, Lange - Posit) & " (UTC)"
                                Case "certificateserial": TxTSE.Text = TxTSE.Text & vbCrLf & "Seriennummer : " & Right$(TmZei, Lange - Posit)
                                Case "signaturecounter": TxTSE.Text = TxTSE.Text & vbCrLf & "Anzahl Signaturen : " & Right$(TmZei, Lange - Posit)
                                Case "transactioncounter": TxTSE.Text = TxTSE.Text & vbCrLf & "Anzahl Transaktionen : " & Right$(TmZei, Lange - Posit)
                                End Select
                            End If
                        End If
                    Next AktZe
                End If
            End If
        End With
    End If
    If TSEOK = True Then
        TxTSE.Text = TxTSE.Text & vbCrLf & "Die TSE wurde erfolgreich geprüft!" & vbCrLf
    Else
        TxTSE.Text = TxTSE.Text & vbCrLf & "Es kann keine TSS eingerichtet werden!" & vbCrLf
    End If
End If

'------ Klient Einrichten ------

If GlTSK <> vbNullString Then 'TSE Organisation Key
    If GlTSN <> vbNullString Then
        For AktZa = 1 To UBound(GlGeK)
            If CBool(GlGeK(AktZa, 5)) = True Then 'Kassen
                If GlGeK(AktZa, 8) = vbNullString Or Left$(GlGeK(AktZa, 8), 1) = "K" Then
                    KasNr = GlGeK(AktZa, 0)
                    KaNam = SNaFi(GlGeK(AktZa, 1), True)
                    DaNam = CreateID("T") & ".ini"
                    DaIni = GlTmp & DaNam
                    
                    PaStr = "createclient" & Space$(1) & GlTSK & Space$(1) & GlTSS & Space$(1) & GlTSN & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
                    WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
                    DoEvents

                    With clFil
                        If .FilVor(DaIni) = True Then
                            .FilPfa DaIni
                            TmpSt = .FilReSt
                            If TmpSt <> vbNullString Then
                                AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                                For AktZe = 0 To UBound(AryZe)
                                    If AryZe(AktZe) <> vbNullString Then
                                        TmZei = AryZe(AktZe)
                                        Posit = InStr(1, TmZei, "=", 1)
                                        If Posit > 0 Then
                                            InTyp = LCase(Left$(TmZei, Posit - 1))
                                            If InTyp = "clientid" Then
                                                Lange = Len(TmZei)
                                                KliID = Right$(TmZei, Lange - Posit)
                                                If KliID <> vbNullString Then
                                                    DBCmEx2 "qrySimBaTSE", "@IdStr", "@IdxNr", KliID, KasNr
                                                    GlGeK(AktZa, 8) = KliID
                                                    TxTSE.Text = TxTSE.Text & vbCrLf & "Klient : " & KaNam & " wurde eingerichtet"
                                                    TxTSE.Text = TxTSE.Text & vbCrLf & KliID & vbCrLf
                                                    DoEvents
                                                End If
                                            End If
                                        End If
                                    End If
                                Next AktZe
                            End If
                        End If
                    End With
                    Exit For 'nur eine Kasse einrichten
                Else
                    KaNam = GlGeK(AktZa, 1)
                    KliID = GlGeK(AktZa, 8)
                    DaNam = CreateID("T") & ".ini"
                    DaIni = GlTmp & DaNam
                    
                    PaStr = "testclient" & Space$(1) & GlTSK & Space$(1) & GlTSS & Space$(1) & TssNa & Space$(1) & KliID & Space$(1) & "--file=" & Chr$(34) & DaIni & Chr$(34) & Space$(1) & "--config=" & Chr$(34) & CoIni & Chr$(34)
                    WindowStart PrNam & Space$(1) & PaStr, vbNormalFocus, True, True
                    DoEvents

                    TxTSE.Text = TxTSE.Text & vbCrLf & "Folgender Klient wurde eingerichtet : " & KaNam
        
                    With clFil
                        If .FilVor(DaIni) = True Then
                            .FilPfa DaIni
                            TmpSt = .FilReSt
                            If TmpSt <> vbNullString Then
                                AryZe = Split(TmpSt, vbCrLf) 'Zeilen aufsplitten
                                For AktZe = 0 To UBound(AryZe)
                                    If AryZe(AktZe) <> vbNullString Then
                                        TmZei = AryZe(AktZe)
                                        Lange = Len(TmZei)
                                        Posit = InStr(1, TmZei, "=", 1)
                                        If Posit > 0 Then
                                            InTyp = LCase(Left$(TmZei, Posit - 1))
                                            Select Case InTyp
                                            Case "exists": If Right$(TmZei, Lange - Posit) = "1" Then KliOK = True
                                            Case "serialnumber": TxTSE.Text = TxTSE.Text & vbCrLf & "Seriennummer : " & Right$(TmZei, Lange - Posit)
                                            Case "created": TxTSE.Text = TxTSE.Text & vbCrLf & "Eingerichtet am: " & Right$(TmZei, Lange - Posit) & " (UTC)"
                                            End Select
                                        End If
                                    End If
                                Next AktZe
                            End If
                        End If
                    End With
                    If KliOK = True Then
                        TxTSE.Text = TxTSE.Text & vbCrLf & "Der Klient : " & KaNam & " wurde erfolgreich geprüft!" & vbCrLf
                    Else
                        TxTSE.Text = TxTSE.Text & vbCrLf & "Der Klient : " & KaNam & " ist nicht eingerichtet!" & vbCrLf
                    End If
                    Exit For 'nur eine Kasse einrichten
                End If
            End If
        Next AktZa
    Else
        TxTSE.Text = TxTSE.Text & vbCrLf & "Es kann kein Klient eingerichtet werden!" & vbCrLf
    End If
Else
    TxTSE.Text = TxTSE.Text & vbCrLf & "Es kann kein Klient eingerichtet werden!" & vbCrLf
End If

TxTSE.Text = TxTSE.Text & vbCrLf & "Der Dialog kann geschlossen werden."
DoEvents
TxTSE.SelStart = Len(TxTSE.Text)

If GlLog = False Then 'General Logging
    With clFil
        .DaLoe = GlTmp & "*.ini" & vbNullChar
        .FilLoe
    End With
Else
    Clipboard.Clear
    Clipboard.SetText PrNam & Space$(1) & PaStr
End If

DoEvents
Screen.MousePointer = vbNormal

Set clFil = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSE_Init " & Err.Number
Resume Next

End Sub
Public Sub TSEZeig(ByVal TmStr As String)
On Error Resume Next

Dim TxTSE As XtremeSuiteControls.FlatEdit

Set FM = frmTSEInit
Set TxTSE = FM.txtTSEIn

TxTSE.Text = TxTSE & vbCrLf & TmStr
TxTSE.SelStart = Len(TxTSE.Text)

End Sub
Public Sub STxFr()
On Error GoTo InErr
'Formatanpassungen Textcontrol

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set TxCoN = FM.TexCont1

CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh

If TxCoN.Left <> 0 Then
    TxCoN.Move 0, 0, (ClBre - ClLin), ClHoh
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "STxFr " & Err.Number
Resume Next

End Sub


