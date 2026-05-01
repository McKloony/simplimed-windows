VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmFormular 
   Caption         =   "Formulare"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1455
      Left            =   1200
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
      _Version        =   1048579
      _ExtentX        =   6800
      _ExtentY        =   2566
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
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
      Left            =   360
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFormular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions
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
Private CoDia As XtremeSuiteControls.CommonDialog
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private clFil As clsFile
Private clLis As clsLisLab
Private clFen As clsFenster
Private Sub FDial(Optional ByVal GesUb As Boolean)
On Error GoTo InErr
'Dateiauswahldialog

Dim RetWe As Long
Dim FiNam As String
Dim DaNam As String
Dim DaPfa As String
Dim RegNa As String
Dim DaExt As String
Dim AktZa As Integer
Dim Mld1, Mld2, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set RpCon = Me.repCont1
Set RpRws = RpCon.Rows
Set RpRcs = RpCon.Records
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows
Set CoDia = frmMain.comDialo

Tit1 = "Formulardatei Zuordnen"
Mld1 = "Sie müssen erst eine Formularüberschrift anklicken, damit Sie dieser eine Formulardatei zuordnen können"
Mld2 = "Es können nur Formulare zugeordnet werden, die sich im Standard-Formulareordner befinden"

Set clFil = New clsFile

If GesUb = True Then
    With CoDia
        .CancelError = True
        .DialogStyle = 1
        .DialogTitle = "Wo befindet sich der richtige Formularordner?"
        If GlFrn <> vbNullString Then 'Formulareordner
            .FileName = GlFrn
        Else
            .FileName = GlFrO
        End If
        RetWe = .ShowBrowseFolder
        GlFrO = .FileName
        If RetWe = 0 Then
            Set CoDia = Nothing
            Set RpRec = Nothing
            Set RpRcs = Nothing
            Set RpSel = Nothing
            Set RpRow = Nothing
            Set RpCls = Nothing
            Set RpCon = Nothing
            Set clFil = Nothing
            Exit Sub
        End If
    End With

    If GlFrO <> vbNullString Then
        If Right$(GlFrO, 1) <> "\" Then
            GlFrO = GlFrO & "\"
        End If
        IniSetVal "System", "ForPfa", LCase(GlFrO)
        For AktZa = 0 To 105  'Formularearray
            IniSetVal "Formular", GlFrm(2, AktZa), GlFrO & GlFrm(1, AktZa)
        Next AktZa
    End If
Else
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            RegNa = RpRow.Record(2).Value
            If RpRow.Record(1).Value <> vbNullString Then
                DaExt = Right$(RpRow.Record(1).Value, 3)
            Else
                DaExt = "blg"
            End If

            With CoDia
                .CancelError = True
                .DialogStyle = 1
                Select Case LCase(DaExt)
                Case "blg": .Filter = "Belegvorlagen (*.blg)|*.blg|Alle Dateien (*.*)|*.*"
                Case "lst": .Filter = "Listenvorlagen (*.lst)|*.lst|Alle Dateien (*.*)|*.*"
                Case "lbl": .Filter = "Etikettenvorlagen (*.lbl)|*.lbl|Alle Dateien (*.*)|*.*"
                Case "crd": .Filter = "Kartenvorlagen (*.crd)|*.crd|Alle Dateien (*.*)|*.*"
                Case "asw": .Filter = "Auswertungen (*.asw)|*.asw|Alle Dateien (*.*)|*.*"
                End Select
                .DefaultExt = "*." & DaExt
                .DialogTitle = "Wo befindet sich die richtige Formulardatei?"
                If GlFrn <> vbNullString Then 'Formulareordner
                    .InitDir = GlFrn
                Else
                    .InitDir = GlFrO
                End If
                .FileName = vbNullString
                .ShowOpen
                FiNam = .FileName
                If .FileTitle = vbNullString Then
                    Set CoDia = Nothing
                Set RpRec = Nothing
                Set RpRcs = Nothing
                Set RpSel = Nothing
                Set RpRow = Nothing
                Set RpCls = Nothing
                Set RpCon = Nothing
                Set clFil = Nothing
                    Exit Sub
                End If
            End With

            For Each RpRec In RpRcs
                If RpRec.Item(2).Value = RegNa Then
                    With clFil
                        .FilPfa FiNam
                        DaNam = .DaNam
                        DaPfa = .DaPfa & "\"
                    End With
                    If LCase(DaPfa) = LCase(GlFrO) Then
                        RpRec(1).Value = LCase(DaNam)
                        GlFrm(1, RpRec.Index) = LCase(DaNam)
                        S_FoSe LCase(DaNam), GlFrm(2, RpRec.Index)
                        IniSetVal "Formular", RegNa, GlFrO & LCase(DaNam)
                        RpCon.Redraw
                    Else
                        WindowMess Mld2, Dial2, Tit1, FM.hwnd
                    End If
                    Exit For
                End If
            Next RpRec

        Else
            WindowMess Mld1, Dial2, Tit1, FM.hwnd
        End If
    Else
        WindowMess Mld1, Dial2, Tit1, FM.hwnd
    End If
End If

Set CoDia = Nothing
Set RpRec = Nothing
Set RpRcs = Nothing
Set RpSel = Nothing
Set RpRow = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Set clFil = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDial " & Err.Number
Exit Sub

End Sub
Private Sub FEdit()
On Error GoTo InErr
'Entscheidet, welches Formular gestalltet werden soll

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim FiNam As String
Dim DaNam As String
Dim DaNaO As String
Dim DaExt As String
Dim ForNa As String
Dim NeNam As String
Dim TmpNa As String
Dim AnzDa As Integer
Dim StaNa As Integer
Dim DiNam() As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmFormular
Set RpCon = FM.repCont1
Set RpRws = RpCon.Rows
Set RpRcs = RpCon.Records
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

Set clLis = New clsLisLab
Set clFil = New clsFile
clFil.hwnd = FM.hwnd

GlAkt = True

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
    
        If GlFrn <> vbNullString Then 'Formulareordner
            FiNam = GlFrn & RpRow.Record(1).Value
        Else
            FiNam = GlFrO & RpRow.Record(1).Value
        End If
        ForNa = RpRow.Record(2).Value

        With clFil
            .FilPfa FiNam
            DaNam = .DaNam
            DaNaO = .DaNaO
            DaExt = .DaExt
        End With
        
        TeTit = "Kopie Erstellen?"
        TeMai = "Möchten Sie vorher eine Kopie des Originalformulars erstellen?"
        TeInh = "Bevor das Formular " & DaNam & " verändert wird, ist es möglich, eine Kopie davon zu erstellen. Das Originalformular bleibt auf diese Weise unverändert."
        TeFus = "Falls erforderlich, kann immer wieder auf das Originalformular zurückgegriffen werden, indem dieses über die Zuordnen- Funktion  wider einbezogen wird."

        If InStrRev(DaNaO, "00", -1, 1) = 0 Then
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
            If GlMes = 33565 Then
                If IsNumeric(Right$(DaNaO, 3)) = True Then
                    TmpNa = Left$(DaNaO, Len(DaNaO) - 4)
                Else
                    TmpNa = DaNaO
                End If

                If clFil.FilVor(GlFrO & TmpNa & "_0*.*") = True Then
                    AnzDa = clFil.FilLis(GlFrO, TmpNa & "_*.*", DiNam)
                    If AnzDa > 0 Then
                        StaNa = CInt(Mid$(DiNam(AnzDa), Len(TmpNa) + 2, 3))
                        NeNam = TmpNa & "_" & Format$(StaNa + 1, "000") & "." & DaExt
                    Else
                        NeNam = TmpNa & "_" & "001." & DaExt
                    End If
                Else
                    NeNam = TmpNa & "_" & "001." & DaExt
                End If
                
                With clFil
                    .DaCop = FiNam & ";" & GlFrO & NeNam & vbNullChar
                    .FilCop 1
                End With
                
                For Each RpRec In RpRcs
                    If RpRec.Item(2).Value = ForNa Then
                        RpRec(1).Value = NeNam
                        GlFrm(1, RpRec.Index) = NeNam
                        S_FoSe NeNam, GlFrm(2, RpRec.Index)
                        Exit For
                    End If
                Next RpRec

                IniSetVal "Formular", ForNa, GlFrO & NeNam
                RpCon.Redraw
            End If
        End If
        
        Select Case ForNa
        Case "AdrEti":
            GlBu1 = RibTab_Adressen
            STaSe ShoCut_Adresse, RibTab_Adressen
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 2
            End With
        Case "AdrLis":
            GlBu1 = RibTab_Adressen
            STaSe ShoCut_Adresse, RibTab_Adressen
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 2
            End With
        Case "RechLi":
            GlBu3 = RibTab_Rechnungen
            STaSe ShoCut_Finanz, RibTab_Rechnungen
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 4
            End With
        Case "PostLi":
            GlBu3 = RibTab_Mahnwesen
            STaSe ShoCut_Finanz, RibTab_Mahnwesen
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 1
            End With
        Case "PostGr":
            GlBu3 = RibTab_Mahnwesen
            STaSe ShoCut_Finanz, RibTab_Mahnwesen
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 1
            End With
        Case "EiMahn":
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .MahGeb = 5
                .LLEdit
            End With
        Case "SaMahn":
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .MahGeb = 5
                .LLEdit
            End With
        Case "StMahn": 'Einzelmahnung (Alternativ)
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .MahGeb = 5
                .LLEdit
            End With
        Case "TerLis":
            GlBu4 = RibTab_Ter_Listen
            STaSe ShoCut_Termin, RibTab_Ter_Listen
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 1
            End With
        Case "AbZuRe":
            GlBu8 = RibTab_Vorbereit
            STaSe ShoCut_Abrechn, RibTab_Vorbereit
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 6
            End With
        Case "TagPro":
            GlBu8 = RibTab_Tagesproto
            STaSe ShoCut_Abrechn, RibTab_Tagesproto
            If GlExT = 2 Then 'Gruppierungsexpansion Tagesprotokoll
                STaEx 1
                DoEvents
            End If
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .Gruppe = 1
                .LLEdUn 6
            End With
        Case "RzDiLa":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzDiA6":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzDiA5":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzHoBl":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzQuBl":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzQuGr":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "BsArUn":
            GlBu2 = RibTab_Belegmodul
            STaSe ShoCut_Kranken, RibTab_Belegmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "BsSpre":
            GlBu2 = RibTab_Belegmodul
            STaSe ShoCut_Kranken, RibTab_Belegmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "BsSchu":
            GlBu2 = RibTab_Belegmodul
            STaSe ShoCut_Kranken, RibTab_Belegmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzQuRo":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzQuGW":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "QuiFrm":
            GlBu2 = RibTab_Belegmodul
            STaSe ShoCut_Kranken, RibTab_Belegmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzQuKa":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzQuPr":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzHeil":
            GlBu2 = RibTab_Belegmodul
            STaSe ShoCut_Kranken, RibTab_Belegmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "RzQuNe":
            GlBu2 = RibTab_Rezeptmodul
            STaSe ShoCut_Kranken, RibTab_Rezeptmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "QuiStu":
            GlBu2 = RibTab_Belegmodul
            STaSe ShoCut_Kranken, RibTab_Belegmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "LabMus":
            GlBu2 = RibTab_Belegmodul
            STaSe ShoCut_Kranken, RibTab_Belegmodul
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 5
            End With
        Case "StaEti":
            GlBu3 = RibTab_HomeBanki
            STaSe ShoCut_Finanz, RibTab_HomeBanki
            Unload Me
            DoEvents
            With clLis
                .ForNam = ForNa
                .FilNam = FiNam
                .LLEdUn 1
            End With
        Case Else:
            With clLis
                .ForNam = ForNa
                If NeNam <> vbNullString Then
                    .FilNam = GlFrO & NeNam
                Else
                    .FilNam = FiNam
                End If
                .MandVo = True
                .MitaVo = GlMiV
                .ArztVo = GlArV
                .LLEdit
            End With
        End Select
    End If
End If

GlAkt = False

Set RpRec = Nothing
Set RpRcs = Nothing
Set RpSel = Nothing
Set RpRow = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Set clFil = Nothing
Set clLis = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FEdit " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50241)
TeMai = IniGetOpt("Hilfe", 50242)
TeInh = IniGetOpt("Hilfe", 50243)
TeFus = IniGetOpt("Hilfe", 50244)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FInit()
On Error GoTo InErr

Dim LiKop As Boolean
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set RpCon = Me.repCont1
Set ImMan = frmMain.imgManag

LiKop = True

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
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = False
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
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = True
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
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.FixedRowHeight = False
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = True
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FLoad()
On Error GoTo InErr

Dim AktZa As Integer
Dim GrCap As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmFormular
Set RpCon = FM.repCont1
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set RpRcs = RpCon.Records
Set RpCls = RpCon.Columns

Set clFil = New clsFile

With RpCls
    Set RpCol = .Add(0, "Formularname", 300, False)
    Set RpCol = .Add(1, "Dateiname", 10, False)
    RpCol.AutoSize = True
    If RpCon.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(2, "RegName", 0, False)
    Set RpCol = .Add(3, "Gruppe", 0, False)
End With

For Each RpCol In RpCls
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
    End With
Next RpCol

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    .Populate
End With

For AktZa = 0 To 105 'WICHTIG!
    Set RpRec = RpRcs.Add()
    Set RpItm = RpRec.AddItem(GlFrm(0, AktZa))
    RpItm.Icon = IC16_Doc_Norm
    Set RpItm = RpRec.AddItem(GlFrm(1, AktZa))
    Set RpItm = RpRec.AddItem(GlFrm(2, AktZa))
    Set RpItm = RpRec.AddItem(GlFrm(3, AktZa))
Next AktZa

With RpCon
    .SortOrder.Add .Columns(3)
    .GroupsOrder.Add .Columns(3)
    .GroupsOrder(0).SortAscending = True
    .Populate
End With

Set RpRws = RpCon.Rows
For Each RpRow In RpRws
    If RpRow.GroupRow = True Then
        Set RpGrw = RpRow
        If Len(RpGrw.GroupCaption) > 0 Then
            GrCap = RpGrw.GroupCaption
            RpGrw.GroupCaption = Mid$(GrCap, 11, Len(GrCap) - 1)
        End If
        If RpGrw.Index > 1 Then
            RpGrw.Expanded = False
        End If
    End If
Next RpRow

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpItm = Nothing
Set RpRec = Nothing
Set RpRcs = Nothing
Set RpSel = Nothing
Set RpRow = Nothing
Set RpRws = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Set clFil = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FMenu()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set FM = frmFormular
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set ImMan = frmMain.imgManag

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
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
    .GlobalSettings.App = App
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
    .ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
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
    .ActiveMenuBar.ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .KeyBindings.Add 0, VK_F1, KY_F1
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F8, KY_F8
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

With CmSta
    .Font.SIZE = 8
    Set CmPan = .AddPane(1)
    CmPan.Width = 200
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = GlFrO
    .Visible = True
End With

Set CmBar = CmBrs.Add("ID_Toolbar", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 24, 24
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlSplitButtonPopup, SY_OP_Uebernahme, "Zuordnen")
    With CmCon
        .ToolTipText = "Ein Formular zuordnen"
        .ShortcutText = "F5"
        .IconId = IC24_Folder_Open
        .BeginGroup = True
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_UeberEinz, "Formularzuordnung")
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_UeberNetz, "Ordnerzuordnung")
        CmCon.Enabled = Not GlRDP
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Bearbeiten, "Bearbeiten")
    With CmCon
        .ToolTipText = "Das Formular bearbeiten"
        .ShortcutText = "F2"
        .IconId = IC24_Square_Edit
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlSplitButtonPopup, SY_OP_Zuruck, "Zurücksetzen")
    With CmCon
        .ToolTipText = "Zurücksetzen der Formularzuordnung"
        .IconId = IC24_Folder_Paper
        .BeginGroup = True
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_Zuruck, "Zurücksetzen aktuelles Formular")
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_Reset, "Zurücksetzen aller Formulare")
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .ToolTipText = "Abbrechen"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

Set CmCoS = CmBar.Controls
For Each CmCon In CmCoS
    CmCon.Style = xtpButtonIconAndCaption
Next CmCon

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmFormular
Set RpCon = Me.repCont1
Set CmBrs = FM.comBar02

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    If ClBre > 1000 And ClHoh > 1000 Then
        RpCon.Move ClLin, ClObn, ClBre, ClHoh
    End If
End If

Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FRest()
On Error GoTo MeErr

Dim FiNam As String
Dim AktZa As Integer
Dim Frage As Integer
Dim Mld1, Tit1 As String

Tit1 = "Grundeinstellungen"
Mld1 = "Möchten Sie die Einstellungen der Formulare jetzt zurücksetzen? Beim Zurücksetzen werden alle Pfadzuordnungen auf ihre Grundwerte eingestellt."

Set FM = frmFormular

FiNam = GlFrO & "*.lsp"

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    For AktZa = 0 To 105 'WICHTIG!
        SFoR1 AktZa, True
        S_FoRe AktZa
    Next AktZa
    DoEvents

    With clFil
        If .FilVor(FiNam) = True Then
            .DaLoe = FiNam & vbNullChar
            .FilLoe
        End If
    End With
End If

Set clFil = Nothing

Unload FM

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FRest " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)

Select Case TolId
Case KY_F1: FHilfe
Case KY_F2: FEdit
Case KY_F5: FDial
Case KY_F11: Unload Me
Case SY_OP_Uebernahme: FDial
Case SY_OP_Bearbeiten: FEdit
Case SY_OP_Zuruck: FZuru
Case SY_OP_Reset: FRest
Case SY_OP_Hilfe: FHilfe
Case SY_OP_Abbruch: Unload Me
Case SY_OP_UeberEinz: FDial
Case SY_OP_UeberNetz: FDial True
                      Unload Me
End Select

End Sub
Private Sub FZuru()
On Error GoTo MeErr

Dim SuStr As String
Dim FrIdx As Integer
Dim AktZa As Integer
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Tit1 = "Zurücksetzen"
Mld1 = "Soll die Dateizuordnung für das markierte Formular jetzt zurückgesetzt werden?"

Set FM = frmFormular
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    Set RpSel = RpCon.SelectedRows
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            SuStr = RpRow.Record(2).Value
            For AktZa = 0 To 105 'WICHTIG
                If SuStr = GlFrm(2, AktZa) Then
                    SFoR1 AktZa, True
                    S_FoRe AktZa
                    DoEvents
                    RpRow.Record(0).Value = GlFrm(0, AktZa)
                    RpRow.Record(1).Value = GlFrm(1, AktZa)
                    RpRow.Record(2).Value = GlFrm(2, AktZa)
                    RpRow.Record(3).Value = GlFrm(3, AktZa)
                    RpCon.Redraw
                    Exit For
                End If
            Next AktZa
        End If
    End If
End If

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZuru " & Err.Number
Resume Next

End Sub
Private Sub FOpn()
On Error GoTo InErr

Set FrmEx = Me.frmExtde

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

With FrmEx
    .ClientMinHeight = 5500
    .ClientMinWidth = 5500
End With

With clFen
    If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
        .FeLin = (GlxGr - GlFeB) / 2
        .FeObn = (GlyGr - GlFeH) / 2
        .FeBre = 600
        .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
    Else
        .FeLin = (GlxGr / 2) - (600 / 2)
        .FeObn = 70
        .FeBre = 600
        .FeHoh = GlyGr - 140
    End If
    .FenMov
End With

Set FrmEx = Nothing
Set clFen = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpn " & Err.Number
Resume Next

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
Private Sub comBar02_Resize()
    FPosi
End Sub
Private Sub Form_Load()
On Error Resume Next

FOpn
FInit
FMenu
FLoad
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmFormular = Nothing
End Sub
Private Sub repCont1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmFormular
Set CmBrs = FM.comBar02
Set RpCon = FM.repCont1
Set CmSta = CmBrs.StatusBar
Set RpRcs = RpCon.Records
Set RpCls = RpCon.Columns

Set RpSel = RpCon.SelectedRows
Set RpRow = RpSel(0)
If RpRow.GroupRow = False Then
    CmSta.Pane(0).Text = RpRow.Record(0).Value
End If

Set CmSta = Nothing
Set CmBrs = Nothing
Set RpItm = Nothing
Set RpRec = Nothing
Set RpRcs = Nothing
Set RpSel = Nothing
Set RpRow = Nothing
Set RpRws = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If GlLiz = True Then
        FEdit
    End If
End Sub
