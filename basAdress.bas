Attribute VB_Name = "basAdress"
Option Explicit

Private FM As Form
Private FS As Form
Private AktCo As VB.Control
Private S1L13 As XtremeSuiteControls.Label
Private S1L20 As XtremeSuiteControls.Label
Private S2L01 As XtremeSuiteControls.Label
Private S2L34 As XtremeSuiteControls.Label
Private S2L35 As XtremeSuiteControls.Label
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Rahm7 As XtremeSuiteControls.GroupBox
Private Rahm8 As XtremeSuiteControls.GroupBox
Private TxGes As XtremeSuiteControls.FlatEdit
Private S3F02 As XtremeSuiteControls.FlatEdit
Private S3F03 As XtremeSuiteControls.FlatEdit
Private TxFir As XtremeSuiteControls.FlatEdit
Private TxKur As XtremeSuiteControls.FlatEdit
Private TxBeH As XtremeSuiteControls.FlatEdit
Private TxKop As XtremeSuiteControls.FlatEdit
Private TxGeb As XtremeSuiteControls.FlatEdit
Private TxZGe As XtremeSuiteControls.FlatEdit
Private TxReG As XtremeSuiteControls.FlatEdit
Private TxErs As XtremeSuiteControls.FlatEdit
Private TxOrt As XtremeSuiteControls.FlatEdit
Private TxNum As XtremeSuiteControls.FlatEdit
Private TxTe1 As XtremeSuiteControls.FlatEdit
Private TxTe2 As XtremeSuiteControls.FlatEdit
Private TxTe3 As XtremeSuiteControls.FlatEdit
Private TxTe4 As XtremeSuiteControls.FlatEdit
Private TxBri As XtremeSuiteControls.FlatEdit
Private TxGut As XtremeSuiteControls.FlatEdit
Private TxTel As XtremeSuiteControls.FlatEdit
Private FeAn1 As XtremeSuiteControls.ComboBox
Private FeAn2 As XtremeSuiteControls.ComboBox
Private FeAn3 As XtremeSuiteControls.ComboBox
Private FeLa1 As XtremeSuiteControls.ComboBox
Private FeLa3 As XtremeSuiteControls.ComboBox
Private FeKat As XtremeSuiteControls.ComboBox
Private FeTar As XtremeSuiteControls.ComboBox
Private FeWar As XtremeSuiteControls.ComboBox
Private FeBeh As XtremeSuiteControls.ComboBox
Private FePat As XtremeSuiteControls.ComboBox
Private FeZah As XtremeSuiteControls.ComboBox
Private FeLab As XtremeSuiteControls.ComboBox
Private FeFam As XtremeSuiteControls.ComboBox
Private FeTi1 As XtremeSuiteControls.ComboBox
Private FeTi2 As XtremeSuiteControls.ComboBox
Private FeTi3 As XtremeSuiteControls.ComboBox
Private FeTi4 As XtremeSuiteControls.ComboBox
Private FeTi5 As XtremeSuiteControls.ComboBox
Private FeTi6 As XtremeSuiteControls.ComboBox
Private FeTi7 As XtremeSuiteControls.ComboBox
Private FeGes As XtremeSuiteControls.ComboBox
Private FeBeG As XtremeSuiteControls.ComboBox
Private CmFch As XtremeSuiteControls.ComboBox
Private CmArt As XtremeSuiteControls.ComboBox
Private CmVdo As XtremeSuiteControls.ComboBox
Private CmBrf As XtremeSuiteControls.ComboBox
Private TsDia As XtremeSuiteControls.TaskDialog
Private CoDia As XtremeSuiteControls.CommonDialog
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmBuT As XtremeCommandBars.CommandBarButton
Private CmBuD As XtremeCommandBars.CommandBarButton
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private PoItm As XtremeSuiteControls.PopupControlItem
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
Private TabCo As XtremeSuiteControls.TabControl
Private TabIt As XtremeSuiteControls.TabControlItem
Private TaPa1 As XtremeSuiteControls.TabControlPage
Private TaPa2 As XtremeSuiteControls.TabControlPage
Private TaPa3 As XtremeSuiteControls.TabControlPage
Private PrGr1 As XtremePropertyGrid.PropertyGrid
Private PrGr2 As XtremePropertyGrid.PropertyGrid
Private PrGr3 As XtremePropertyGrid.PropertyGrid

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private Const olContactItem = 2
Private Const olAppointmentItem = 1
Private Const olFolderCalendar = 9
Private Const olFolderContacts = 10

Private Const TAPIERR_NOREQUESTRECIPIENT As Long = -2&
Private Const TAPIERR_REQUESTQUEUEFULL As Long = -3&
Private Const TAPIERR_INVALDESTADDRESS As Long = -4&

Private clFil As clsFile
Private clWor As clsWord
Private clAnw As clsAnwend
Private clFen As clsFenster
Private clDru As clsDruck
Private clNet As clsNetz
Private clLis As clsLisLab
Private clChe As clsChipcard

Private Declare Function SCardComand Lib "SCARD32.dll" (Handle As Long, ByVal Cmd As String, CmdLen As Long, ByVal DataIn As String, DataInLen As Long, ByVal DataOut As String, DataOutLen As Long) As Long
Private Declare Function tapiRequestMakeCall Lib "TAPI32.DLL" (ByVal DestAddress As String, ByVal AppName As String, ByVal CalledParty As String, ByVal Comment As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Sub AAnre()
On Error GoTo CrErr

Dim TeDat As Date
Dim GeDat As Date
Dim TagWe As String
Dim AnTex As String

Set FM = frmAdress
Set FeAn1 = FM.txtS1F02
Set FeGes = FM.cmbS1F08
Set TxGeb = FM.txtS1F13
Set TxGes = FM.txtGesch

If TxGeb.Text <> vbNullString Then
    If IsDate(TxGeb.Text) Then
        GeDat = TxGeb.Text
        TeDat = DateAdd("yyyy", -18, Date)
        If TeDat <= GeDat Then
            If FeAn1.Text <> vbNullString Then
                AnTex = FeAn1.Text
                Select Case AnTex
                Case "Herrn": FeGes.ListIndex = 4
                Case "Frau": FeGes.ListIndex = 5
                Case "Firma": FeGes.ListIndex = 0
                Case "Familie": FeGes.ListIndex = 0
                Case "Herrn und Frau": FeGes.ListIndex = 0
                Case "Herrn/Frau": FeGes.ListIndex = 0
                Case "Sr.": FeGes.ListIndex = 4
                Case "Sra.": FeGes.ListIndex = 5
                Case "Divers": FeGes.ListIndex = 6
                Case Else: FeGes.ListIndex = 0
                End Select
            End If
        Else
            If FeAn1.Text <> vbNullString Then
                AnTex = FeAn1.Text
                Select Case AnTex
                Case "Herrn": FeGes.ListIndex = 1
                Case "Frau": FeGes.ListIndex = 2
                Case "Firma": FeGes.ListIndex = 0
                Case "Familie": FeGes.ListIndex = 0
                Case "Herrn und Frau": FeGes.ListIndex = 0
                Case "Herrn/Frau": FeGes.ListIndex = 0
                Case "Sr.": FeGes.ListIndex = 1
                Case "Sra.": FeGes.ListIndex = 2
                Case "Divers": FeGes.ListIndex = 6
                Case Else: FeGes.ListIndex = 0
                End Select
            End If
        End If
    Else
        If FeAn1.Text <> vbNullString Then
            AnTex = FeAn1.Text
            Select Case AnTex
            Case "Herrn": FeGes.ListIndex = 1
            Case "Frau": FeGes.ListIndex = 2
            Case "Firma": FeGes.ListIndex = 0
            Case "Familie": FeGes.ListIndex = 0
            Case "Herrn und Frau": FeGes.ListIndex = 0
            Case "Herrn/Frau": FeGes.ListIndex = 0
            Case "Sr.": FeGes.ListIndex = 1
            Case "Sra.": FeGes.ListIndex = 2
            Case "Divers": FeGes.ListIndex = 6
            Case Else: FeGes.ListIndex = 0
            End Select
        End If
    End If
Else
    If FeAn1.Text <> vbNullString Then
        AnTex = FeAn1.Text
        Select Case AnTex
        Case "Herrn": FeGes.ListIndex = 1
        Case "Frau": FeGes.ListIndex = 2
        Case "Firma": FeGes.ListIndex = 0
        Case "Familie": FeGes.ListIndex = 0
        Case "Herrn und Frau": FeGes.ListIndex = 0
        Case "Herrn/Frau": FeGes.ListIndex = 0
        Case "Sr.": FeGes.ListIndex = 1
        Case "Sra.": FeGes.ListIndex = 2
        Case "Divers": FeGes.ListIndex = 6
        Case Else: FeGes.ListIndex = 0
        End Select
    End If
End If

If FeGes.Text <> vbNullString Then
    TxGes.Text = FeGes.Text
    TagWe = Mid$(TxGes.Tag, 2, Len(TxGes.Tag) - 1)
    TxGes.Tag = "1" & TagWe
    GlAdS = True
End If

Exit Sub

CrErr:
If GlDbg = True Then SErLog Err.Description & " AAnre " & Err.Number
Resume Next

End Sub
Public Sub AChip()
On Error GoTo CrErr
'Liest die Chipkarte ein

Dim RetBy As Byte
Dim RetWe As Long
Dim TagWe As String
Dim PatGe As String
Dim RetSt As String
Dim KaGul As String
Dim StrFe As String
Dim KasNa As String
Dim KasNr As String
Dim KarNr As String
Dim VerNr As String
Dim KaSta As String
Dim StaEr As String
Dim PaTit As String
Dim PaVor As String
Dim PaZun As String
Dim PaNam As String
Dim PaGeb As String
Dim PaGes As String
Dim PaStr As String
Dim PaLKZ As String
Dim PaPLZ As String
Dim PaOrt As String
Dim KaDat As String
Dim TmStr As String
Dim Lange As Integer
Dim RetAb As Integer
Dim Posit As Integer
Dim FeAnr As XtremeSuiteControls.ComboBox
Dim FeTit As XtremeSuiteControls.FlatEdit
Dim FeVor As XtremeSuiteControls.FlatEdit
Dim FeNam As XtremeSuiteControls.FlatEdit
Dim FeStr As XtremeSuiteControls.FlatEdit
Dim FePLZ As XtremeSuiteControls.FlatEdit
Dim TxOrt As XtremeSuiteControls.FlatEdit
Dim FeGeb As XtremeSuiteControls.FlatEdit
Dim FeVer As XtremeSuiteControls.FlatEdit
Dim FeVNr As XtremeSuiteControls.FlatEdit
Dim FeKar As XtremeSuiteControls.FlatEdit
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrIts As XtremePropertyGrid.PropertyGridItems
Dim Mld1, Tit1 As String

Set FM = frmAdress
Set PrGr1 = FM.prpGrid1
Set FeAnr = FM.txtS1F02
Set FeTit = FM.txtS1F03
Set FeVor = FM.txtS1F04
Set FeNam = FM.txtS1F05
Set FeStr = FM.txtS1F06
Set FePLZ = FM.txtS1F08
Set TxOrt = FM.txtS1F09
Set TxGeb = FM.txtS1F13
Set FeVer = FM.txtVersi
Set FeVNr = FM.txtVerNr
Set FeKar = FM.txtKarGu
Set PrIts = PrGr1.Categories

Set clChe = New clsChipcard

Screen.MousePointer = vbHourglass

If GlChp = 0 Then
    Mld1 = "Im Optionsdialog ist der Chipkartenleser deaktiviert"
    Tit1 = "Chipkartenleser"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
Else
    KasNa = 0
    StrFe = vbNullString
    KasNa = vbNullString
    KasNr = vbNullString
    KarNr = vbNullString
    VerNr = vbNullString
    KaSta = vbNullString
    StaEr = vbNullString
    PaTit = vbNullString
    PaVor = vbNullString
    PaZun = vbNullString
    PaNam = vbNullString
    PaGeb = vbNullString
    PaStr = vbNullString
    PaLKZ = vbNullString
    PaPLZ = vbNullString
    PaOrt = vbNullString
    KaGul = vbNullString
    KaDat = vbNullString

    With clChe
        .SmartReader = GlKvk(GlChp, 1)
        .TerminalPort = GlChR
        
        RetWe = .CarRead()
        StrFe = .Returnstring

        If RetWe = 0 Then
            PaTit = .Patient_Titel
            PaVor = .Patient_Vorname
            PaZun = .Patient_Zusatz
            PaNam = .Patient_Name
            PaGeb = .Patient_Geboren
            If .IsteGK = True Then
                TmStr = .Patient_Strasse
                Lange = Len(TmStr)
                Posit = InStrRev(TmStr, Chr$(32), -1, vbTextCompare)
                If Posit > 0 Then
                    If IsNumeric(Mid$(TmStr, Posit + 1, Lange - Posit)) = True Then
                        PaStr = TmStr
                    Else
                        PaStr = TmStr & Chr$(32) & .Patient_Hausnummer
                    End If
                Else
                    PaStr = .Patient_Strasse & Chr$(32) & .Patient_Hausnummer
                End If
                PaGes = LCase(.Patient_Geschlecht)
            Else
                PaStr = .Patient_Strasse
                PaGes = "w"
            End If
            PaLKZ = .Patient_Landeskennzeichen
            PaPLZ = .Patient_PLZ
            PaOrt = .Patient_Ort
            
            KasNa = .Kostentraegername2
            KasNr = .Kostentraegerkennung
            KarNr = .Kartennummer
            VerNr = .Versicherten_ID
            KaSta = .Kartenstatus
            StaEr = .KVKStatus
            KaGul = .Kartengueltigkeit
            KaDat = .Kartendatum

            FeVor.Text = PaVor
            FeNam.Text = PaNam
            FePLZ.Text = PaPLZ
            TxOrt.Text = PaOrt
            FeStr.Text = PaStr
            FeTit.Text = PaTit
            FeVer.Text = KasNa
            FeVNr.Text = VerNr
            FeKar.Text = KaGul
            
            PatGe = Left$(PaGeb, 2) & "." & Mid$(PaGeb, 3, 2) & "." & Right$(PaGeb, 4)

            TxGeb.Text = PatGe
            If PaGes = "m" Then
                FeAnr.Text = "Herrn"
            Else
                FeAnr.Text = "Frau"
            End If
            
            FeAnr.Tag = 1 & "Anrede"
            FeVor.Tag = 1 & "Vorname"
            FeNam.Tag = 1 & "Name"
            FePLZ.Tag = 1 & "PLZ"
            TxOrt.Tag = 1 & "Ort"
            TxGeb.Tag = 1 & "Geboren"
            FeStr.Tag = 1 & "Straße"
            FeTit.Tag = 1 & "Titel"
            FeVer.Tag = 1 & "Versicherung"
            FeVNr.Tag = 1 & "Kartennummer"
            FeKar.Tag = 1 & "Kartengultig"
    
            AKopi
            DoEvents
            AdBrf
            DoEvents
            
            FM.txtS2F20.Tag = "1" & "R_Briefanrede"
            FM.txtS2F20.Text = FM.cmbS1F10.Text
            
            AErAd
            GlAdS = True
        End If
    End With
    If RetWe = 0 Then
        SPopu KarNr, StrFe, IC48_Warning
    Else
        SPopu "Fehler: " & RetWe, StrFe, IC48_Warning
    End If
End If

Screen.MousePointer = vbNormal

Set clChe = Nothing

Set PrKat = Nothing
Set PrGr1 = Nothing

Exit Sub

CrErr:
If GlDbg = True Then SErLog Err.Description & " AChip " & Err.Number
Resume Next

End Sub
Public Sub AdBrf(Optional ByVal BrStr As String, Optional ByVal BeDat As Boolean)
On Error GoTo LiErr
'Generiert die vier Briefanreden

Dim RetWe As Long
Dim TagWe As String
Dim FBri1 As String
Dim FBri2 As String
Dim FBri3 As String
Dim FBri4 As String
Dim FBri5 As String
Dim FBri6 As String
Dim FAnre As Variant
Dim FTite As Variant
Dim FName As Variant
Dim FVorn As Variant
Dim FDuSi As Integer

If BeDat = True Then
    Set FM = frmMandant
Else
    Set FM = frmAdress
End If
Set CmBrf = FM.cmbS1F10

FAnre = FM.txtS2F12.Text
FTite = FM.txtS2F13.Text
FVorn = FM.txtS2F14.Text
FName = FM.txtS2F15.Text
If FM.txtS1F20.Text <> vbNullString Then
    FDuSi = FM.txtS1F20.Text
Else
    FDuSi = 2
End If

If FAnre = "Herrn und Frau" Then
    FBri1 = "Lieber Herr und Frau " & FName & ","
ElseIf FAnre = "herrn und frau" Then
    FBri1 = "lieber herr und frau " & FName & ","
ElseIf FAnre Like "*Herr*" Then
    FBri1 = "Lieber " & FVorn & ","
ElseIf FAnre Like "*herr*" Then
    FBri1 = "lieber " & FVorn & ","
ElseIf FAnre Like "*Frau*" Then
    FBri1 = "Liebe " & FVorn & ","
ElseIf FAnre Like "*frau*" Then
    FBri1 = "liebe " & FVorn & ","
ElseIf FAnre Like "Familie" Then
    FBri1 = "Liebe Familie " & FName & ","
ElseIf FAnre Like "familie" Then
    FBri1 = "liebe familie " & FName & ","
ElseIf FAnre Like "Divers" Then
    FBri1 = "Liebe(r) " & FVorn & ","
ElseIf FAnre Like "divers" Then
    FBri1 = "liebe(r) " & FVorn & ","
Else
    FBri1 = "Sehr geehrte Damen und Herren,"
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri2 = "Sehr geehrter Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri2 = "sehr geehrter herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri2 = "Sehr geehrter " & Trim$("Herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri2 = "sehr geehrter " & Trim$("herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri2 = "Sehr geehrte " & Trim$("Frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri2 = "sehr geehrte " & Trim$("frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri2 = "Sehr geehrte Familie " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri2 = "sehr geehrte familie " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri2 = "Sehr geehrte(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri2 = "sehr geehrte(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    Else
        FBri2 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri2 = "Sehr geehrter Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri2 = "sehr geehrter herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri2 = "Sehr geehrter " & Trim$("Herr " & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri2 = "sehr geehrter " & Trim$("herr " & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri2 = "Sehr geehrte " & Trim$("Frau " & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri2 = "sehr geehrte " & Trim$("frau " & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri2 = "Sehr geehrte Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri2 = "sehr geehrte familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri2 = "Sehr geehrte(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri2 = "sehr geehrte(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    Else
        FBri2 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri3 = "Guten Tag Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri3 = "guten tag herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri3 = "Guten Tag " & Trim$("Herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri3 = "guten tag " & Trim$("herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri3 = "Guten Tag " & Trim$("Frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri3 = "guten tag " & Trim$("frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri3 = "Guten Tag Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri3 = "guten tag familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri3 = "Guten Tag " & Trim$(FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri3 = "guten tag " & Trim$(FTite & Chr$(32) & FName) & ","
    Else
        FBri3 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri3 = "Gutan Tag Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri3 = "guten tag herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri3 = "Guten Tag Herr " & FName & ","
    ElseIf FAnre Like "*herr*" Then
        FBri3 = "guten tag herr " & FName & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri3 = "Guten Tag Frau " & FName & ","
    ElseIf FAnre Like "*frau*" Then
        FBri3 = "guten tag frau " & FName & ","
    ElseIf FAnre Like "Familie" Then
        FBri3 = "Guten Tag Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri3 = "guten tag familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri3 = "Guten Tag " & FVorn & Chr$(32) & FName & ","
    ElseIf FAnre Like "divers" Then
        FBri3 = "guten tag " & FVorn & Chr$(32) & FName & ","
    Else
        FBri3 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri4 = "Guten Tag Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri4 = "guten tag herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri4 = "Guten Tag Herr " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri4 = "guten tag herr " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri4 = "Guten Tag Frau " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri4 = "guten tag frau" & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri4 = "Guten Tag Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri4 = "guten tag familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri4 = "Guten Tag " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri4 = "guten tag " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    Else
        FBri4 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri4 = "Gutan Tag Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri4 = "guten tag herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri4 = "Hallo Herr " & FName & ","
    ElseIf FAnre Like "*herr*" Then
        FBri4 = "hallo herr " & FName & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri4 = "Hallo Frau " & FName & ","
    ElseIf FAnre Like "*frau*" Then
        FBri4 = "hallo frau " & FName & ","
    ElseIf FAnre Like "Familie" Then
        FBri4 = "Hallo Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri4 = "hallo familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri4 = "Hallo " & FVorn & Chr$(32) & FName & ","
    ElseIf FAnre Like "divers" Then
        FBri4 = "hallo " & FVorn & Chr$(32) & FName & ","
    Else
        FBri4 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri5 = "Lieber Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri5 = "lieber herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri5 = "Lieber " & Trim$("Herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri5 = "lieber " & Trim$("herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri5 = "Liebe " & Trim$("Frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri5 = "liebe " & Trim$("frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri5 = "Liebe Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri5 = "liebe familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri5 = "Liebe(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri5 = "liebe(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    Else
        FBri5 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri5 = "Lieber Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri5 = "lieber herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri5 = "Lieber " & Trim$("Herr " & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri5 = "lieber " & Trim$("herr " & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri5 = "Liebe " & Trim$("Frau " & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri5 = "liebe " & Trim$("frau " & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri5 = "Liebe Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri5 = "liebe familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri5 = "Liebe(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri5 = "liebe(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    Else
        FBri5 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FAnre = "Herrn und Frau" Then
    FBri6 = "Hallo Herr und Frau " & FName & ","
ElseIf FAnre = "herrn und frau" Then
    FBri6 = "hallo herr und frau " & FName & ","
ElseIf FAnre Like "*Herr*" Then
    FBri6 = "Hallo " & FVorn & ","
ElseIf FAnre Like "*herr*" Then
    FBri6 = "hallo " & FVorn & ","
ElseIf FAnre Like "*Frau*" Then
    FBri6 = "Hallo " & FVorn & ","
ElseIf FAnre Like "*frau*" Then
    FBri6 = "hallo " & FVorn & ","
ElseIf FAnre Like "Familie" Then
    FBri6 = "Hallo Familie " & FName & ","
ElseIf FAnre Like "familie" Then
    FBri6 = "hallo familie " & FName & ","
ElseIf FAnre Like "Divers" Then
    FBri6 = "Hallo " & FVorn & ","
ElseIf FAnre Like "Divers" Then
    FBri6 = "hallo " & FVorn & ","
Else
    FBri6 = "Sehr geehrte Damen und Herren,"
End If

With CmBrf
    .Clear
    .AddItem FBri1
    .ItemData(0) = 1
    .AddItem FBri2
    .ItemData(1) = 2
    .AddItem FBri3
    .ItemData(2) = 3
    .AddItem FBri4
    .ItemData(3) = 4
    .AddItem FBri5
    .ItemData(4) = 5
    .AddItem FBri6
    .ItemData(5) = 6
End With

If BrStr <> vbNullString Then
    If BrStr = FBri1 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 0, ByVal 0&)
    ElseIf BrStr = FBri2 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 1, ByVal 0&)
    ElseIf BrStr = FBri3 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 2, ByVal 0&)
    ElseIf BrStr = FBri4 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 3, ByVal 0&)
    ElseIf BrStr = FBri5 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 4, ByVal 0&)
    ElseIf BrStr = FBri6 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 5, ByVal 0&)
    End If
Else
    If FDuSi > 0 Then
        CmBrf.ListIndex = FDuSi - 1
    Else
        CmBrf.ListIndex = 1
    End If
    If BeDat = True Then
        TagWe = Mid$(CmBrf.Tag, 2, Len(CmBrf.Tag) - 1)
        CmBrf.Tag = "1" & TagWe
    End If
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AdBrf " & Err.Number
Resume Next

End Sub

Public Sub AdBri(Optional ByVal BrStr As String)
On Error GoTo LiErr
'Generiert die vier Briefanreden

Dim RetWe As Long
Dim FBri1 As String
Dim FBri2 As String
Dim FBri3 As String
Dim FBri4 As String
Dim FBri5 As String
Dim FBri6 As String
Dim FAnre As Variant
Dim FTite As Variant
Dim FName As Variant
Dim FVorn As Variant
Dim FDuSi As Integer

Set FM = frmAdress
Set CmBrf = FM.cmbS4F11

FAnre = FM.txtS4F02.Text
FTite = FM.txtS4F03.Text
FVorn = FM.txtS4F04.Text
FName = FM.txtS4F05.Text

FDuSi = 2

If FAnre = "Herrn und Frau" Then
    FBri1 = "Lieber Herr und Frau " & FName & ","
ElseIf FAnre = "herrn und frau" Then
    FBri1 = "lieber herr und frau " & FName & ","
ElseIf FAnre Like "*Herr*" Then
    FBri1 = "Lieber " & FVorn & ","
ElseIf FAnre Like "*herr*" Then
    FBri1 = "lieber " & FVorn & ","
ElseIf FAnre Like "*Frau*" Then
    FBri1 = "Liebe " & FVorn & ","
ElseIf FAnre Like "*frau*" Then
    FBri1 = "liebe " & FVorn & ","
ElseIf FAnre Like "Familie" Then
    FBri1 = "Liebe Familie " & FName & ","
ElseIf FAnre Like "familie" Then
    FBri1 = "liebe familie " & FName & ","
ElseIf FAnre Like "Divers" Then
    FBri1 = "Liebe(r) " & FVorn & ","
ElseIf FAnre Like "divers" Then
    FBri1 = "liebe(r) " & FVorn & ","
Else
    FBri1 = "Sehr geehrte Damen und Herren,"
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri2 = "Sehr geehrter Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri2 = "sehr geehrter herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri2 = "Sehr geehrter " & Trim$("Herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri2 = "sehr geehrter " & Trim$("herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri2 = "Sehr geehrte " & Trim$("Frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri2 = "sehr geehrte " & Trim$("frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri2 = "Sehr geehrte Familie " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri2 = "sehr geehrte familie " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri2 = "Sehr geehrte(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri2 = "sehr geehrte(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    Else
        FBri2 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri2 = "Sehr geehrter Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri2 = "sehr geehrter herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri2 = "Sehr geehrter " & Trim$("Herr " & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri2 = "sehr geehrter " & Trim$("herr " & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri2 = "Sehr geehrte " & Trim$("Frau " & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri2 = "sehr geehrte " & Trim$("frau " & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri2 = "Sehr geehrte Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri2 = "sehr geehrte familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri2 = "Sehr geehrte(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri2 = "sehr geehrte(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    Else
        FBri2 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri3 = "Guten Tag Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri3 = "guten tag herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri3 = "Guten Tag " & Trim$("Herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri3 = "guten tag " & Trim$("herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri3 = "Guten Tag " & Trim$("Frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri3 = "guten tag " & Trim$("frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri3 = "Guten Tag Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri3 = "guten tag familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri3 = "Guten Tag " & Trim$(FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri3 = "guten tag " & Trim$(FTite & Chr$(32) & FName) & ","
    Else
        FBri3 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri3 = "Gutan Tag Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri3 = "guten tag herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri3 = "Guten Tag Herr " & FName & ","
    ElseIf FAnre Like "*herr*" Then
        FBri3 = "guten tag herr " & FName & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri3 = "Guten Tag Frau " & FName & ","
    ElseIf FAnre Like "*frau*" Then
        FBri3 = "guten tag frau " & FName & ","
    ElseIf FAnre Like "Familie" Then
        FBri3 = "Guten Tag Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri3 = "guten tag familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri3 = "Guten Tag " & FVorn & Chr$(32) & FName & ","
    ElseIf FAnre Like "divers" Then
        FBri3 = "guten tag " & FVorn & Chr$(32) & FName & ","
    Else
        FBri3 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri4 = "Guten Tag Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri4 = "guten tag herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri4 = "Guten Tag Herr " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri4 = "guten tag herr " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri4 = "Guten Tag Frau " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri4 = "guten tag frau" & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri4 = "Guten Tag Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri4 = "guten tag familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri4 = "Guten Tag " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri4 = "guten tag " & Trim$(FTite & Chr$(32) & FVorn & Chr$(32) & FName) & ","
    Else
        FBri4 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri4 = "Gutan Tag Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri4 = "guten tag herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri4 = "Hallo Herr " & FName & ","
    ElseIf FAnre Like "*herr*" Then
        FBri4 = "hallo herr " & FName & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri4 = "Hallo Frau " & FName & ","
    ElseIf FAnre Like "*frau*" Then
        FBri4 = "hallo frau " & FName & ","
    ElseIf FAnre Like "Familie" Then
        FBri4 = "Hallo Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri4 = "hallo familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri4 = "Hallo " & FVorn & Chr$(32) & FName & ","
    ElseIf FAnre Like "divers" Then
        FBri4 = "hallo " & FVorn & Chr$(32) & FName & ","
    Else
        FBri4 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FTite <> vbNullString Then
    If FAnre = "Herrn und Frau" Then
        FBri5 = "Lieber Herr und Frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri5 = "lieber herr und frau " & FTite & Chr$(32) & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri5 = "Lieber " & Trim$("Herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri5 = "lieber " & Trim$("herr " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri5 = "Liebe " & Trim$("Frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri5 = "liebe " & Trim$("frau " & FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri5 = "Liebe Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri5 = "liebe familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri5 = "Liebe(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri5 = "liebe(r) " & Trim$(FTite & Chr$(32) & FName) & ","
    Else
        FBri5 = "Sehr geehrte Damen und Herren,"
    End If
Else
    If FAnre = "Herrn und Frau" Then
        FBri5 = "Lieber Herr und Frau " & FName & ","
    ElseIf FAnre = "herrn und frau" Then
        FBri5 = "lieber herr und frau " & FName & ","
    ElseIf FAnre Like "*Herr*" Then
        FBri5 = "Lieber " & Trim$("Herr " & FName) & ","
    ElseIf FAnre Like "*herr*" Then
        FBri5 = "lieber " & Trim$("herr " & FName) & ","
    ElseIf FAnre Like "*Frau*" Then
        FBri5 = "Liebe " & Trim$("Frau " & FName) & ","
    ElseIf FAnre Like "*frau*" Then
        FBri5 = "liebe " & Trim$("frau " & FName) & ","
    ElseIf FAnre Like "Familie" Then
        FBri5 = "Liebe Familie " & FName & ","
    ElseIf FAnre Like "familie" Then
        FBri5 = "liebe familie " & FName & ","
    ElseIf FAnre Like "Divers" Then
        FBri5 = "Liebe(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    ElseIf FAnre Like "divers" Then
        FBri5 = "liebe(r) " & Trim$(FVorn & Chr$(32) & FName) & ","
    Else
        FBri5 = "Sehr geehrte Damen und Herren,"
    End If
End If

If FAnre = "Herrn und Frau" Then
    FBri6 = "Hallo Herr und Frau " & FName & ","
ElseIf FAnre = "herrn und frau" Then
    FBri6 = "hallo herr und frau " & FName & ","
ElseIf FAnre Like "*Herr*" Then
    FBri6 = "Hallo " & FVorn & ","
ElseIf FAnre Like "*herr*" Then
    FBri6 = "hallo " & FVorn & ","
ElseIf FAnre Like "*Frau*" Then
    FBri6 = "Hallo " & FVorn & ","
ElseIf FAnre Like "*frau*" Then
    FBri6 = "hallo " & FVorn & ","
ElseIf FAnre Like "Familie" Then
    FBri6 = "Hallo Familie " & FName & ","
ElseIf FAnre Like "familie" Then
    FBri6 = "hallo familie " & FName & ","
ElseIf FAnre Like "Divers" Then
    FBri6 = "Hallo " & FVorn & ","
ElseIf FAnre Like "Divers" Then
    FBri6 = "hallo " & FVorn & ","
Else
    FBri6 = "Sehr geehrte Damen und Herren,"
End If

With CmBrf
    .Clear
    .AddItem FBri1
    .ItemData(0) = 1
    .AddItem FBri2
    .ItemData(1) = 2
    .AddItem FBri3
    .ItemData(2) = 3
    .AddItem FBri4
    .ItemData(3) = 4
    .AddItem FBri5
    .ItemData(4) = 5
    .AddItem FBri6
    .ItemData(5) = 6
End With

If BrStr <> vbNullString Then
    If BrStr = FBri1 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 0, ByVal 0&)
    ElseIf BrStr = FBri2 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 1, ByVal 0&)
    ElseIf BrStr = FBri3 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 2, ByVal 0&)
    ElseIf BrStr = FBri4 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 3, ByVal 0&)
    ElseIf BrStr = FBri5 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 4, ByVal 0&)
    ElseIf BrStr = FBri6 Then
        RetWe = SendMessage(CmBrf.hwnd, CB_SETCURSEL, 5, ByVal 0&)
    End If
Else
    If FDuSi > 0 Then
        CmBrf.ListIndex = FDuSi - 1
    Else
        CmBrf.ListIndex = 1
    End If
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AdBri " & Err.Number
Resume Next

End Sub


Public Sub AdFMa()
On Error GoTo LaErr

If WindowLoad("frmAdrFilt") = True Then
    frmAdrFilt.ZOrder 0
    Exit Sub
End If

GlAsL = True

AdFRe

Load frmAdrFilt

Set FM = frmAdrFilt

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    .FeLin = IniGetVal("Assistent", "FenLin")
    .FeObn = IniGetVal("Assistent", "FenObe")
    .FeBre = IniGetVal("Assistent", "FenBre")
    .FeHoh = IniGetVal("Assistent", "FenHoh")
End With

AFont FM

With clFen
    .FenMov
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If
DoEvents

frmAdrFilt.Show
DoEvents
GlAsL = False

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " AdFMa " & Err.Number
Resume Next

End Sub
Private Sub AdFRe()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "Assistent") = False Then
    xGro = 510
    yGro = 370
    
    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)
     
    IniSetSek "Assistent"
    IniSetVal "Assistent", "FenLin", xPos
    IniSetVal "Assistent", "FenObe", yPos
    IniSetVal "Assistent", "FenBre", xGro
    IniSetVal "Assistent", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " AdFRe " & Err.Number
Resume Next

End Sub
Public Sub AdGru(ByVal Flag As Integer, Optional ByVal NoExp As Boolean = False)
On Error GoTo PeErr
'Lädt die Gruppen in das TreeView

Dim DatKy As String
Dim VorKy As String
Dim KnoKy As String
Dim AkPos As Integer
Dim AltPo As Integer
Dim AktZa As Integer
Dim GruZa As Integer
Dim TrLi1 As XtremeSuiteControls.TreeView
Dim Knote As XtremeSuiteControls.TreeViewNode

Select Case Flag
Case 1: Set FM = frmGrupp
Case 2: Set FM = frmReFilt
Case 3: Set FM = frmAdrFilt
Case 4: Set FM = frmAdrAnpa
Case 5: Set FM = frmMain
Case 6: Set FM = frmOutlook
Case 7: Set FM = frmImport
End Select

Set TrLi1 = FM.trvList1

AktZa = TrLi1.Nodes.Count + 1

For GruZa = 1 To UBound(GlPaG)
    If GlPaG(GruZa, 2) <> vbNullString Then
        DatKy = GlPaG(GruZa, 2)
        KnoKy = "G" & GlPaG(GruZa, 0)
        AkPos = InStrRev(DatKy, ".", Len(DatKy), 1)
        If AkPos > 0 Then
            If AkPos > AltPo Then
                VorKy = TrLi1.Nodes(AktZa - 1).Key
            ElseIf AkPos < AltPo Then
                VorKy = TrLi1.Nodes(AktZa - 1).Parent.Parent.Key
            End If
        Else
            VorKy = "P801"
        End If
        Set Knote = TrLi1.Nodes.Add(VorKy, 4, KnoKy, GlPaG(GruZa, 1), IC16_Folder_Close)
        AltPo = AkPos
        AktZa = AktZa + 1
    End If
Next GruZa

If NoExp = False Then
    For Each Knote In TrLi1.Nodes
        Knote.Expanded = True
    Next Knote
Else
    TrLi1.Nodes("P801").Expanded = True
End If

TrLi1.Nodes.Item(1).EnsureVisible

Exit Sub

PeErr:
If GlDbg = True Then SErLog Err.Description & " AdGru " & Err.Number
Resume Next

End Sub
Public Sub AdTel(TelNr As String, LocSt As String)
On Error GoTo DrErr

Dim RetWe As Long
Dim TmSt1 As String
Dim TmSt2 As String
Dim AkZa1 As Integer
Dim AkZa2 As Integer
Dim Posi1 As Integer
Dim Posi2 As Integer
Dim TapKl As Boolean
Dim Mld1 As String
Dim Tit1 As String

Set FM = frmAdress

TapKl = CBool(IniGetVal("System", "TAPIKl"))

Tit1 = "TAPI Fehler"

TmSt1 = TelNr

For AkZa2 = 1 To 20
    For AkZa1 = 30 To 165
        If AkZa1 < 48 Or AkZa1 > 57 Then
            If AkZa1 <> 43 Then
                TmSt1 = Replace(TmSt1, Chr$(AkZa1), vbNullString, 1)
            End If
        End If
    Next AkZa1
Next AkZa2

If TapKl = True Then
    If Len(TmSt1) > 3 Then
        If Left$(TmSt1, 1) = "+" Then
            TmSt2 = Left$(TmSt1, 3) & "(" & Mid$(TmSt1, 4, 5) & ")" & Mid$(TmSt1, 9, Len(TmSt1) - 8)
        Else
            TmSt2 = "(" & Left$(TmSt1, 5) & ")" & Mid$(TmSt1, 6, Len(TmSt1) - 5)
        End If
    Else
        Exit Sub
    End If
Else
    TmSt2 = TmSt1
End If

RetWe = tapiRequestMakeCall(TmSt2, CStr(GlPrg), LocSt, "")

If RetWe <> 0 Then
    Select Case RetWe
        Case TAPIERR_NOREQUESTRECIPIENT
            Mld1 = "Diw Windows Wählhilfe ist nicht installiert oder konnte nicht gestartet werden!"
        Case TAPIERR_REQUESTQUEUEFULL
            Mld1 = "Die Anrufschlange ist voll!"
        Case TAPIERR_INVALDESTADDRESS
            Mld1 = "Ungültige Telefonnummer!"
        Case Else
            Mld1 = "Sonstiger Fehler!"
    End Select
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

Exit Sub

DrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AdTel " & Err.Number
Resume Next

End Sub
Public Sub AdKop()
On Error GoTo InErr
'Kopiert die aktuelle Adresse

Dim Frage As Integer
Dim Mld1, Tit1 As String

Tit1 = "Eintrag kopieren"
Mld1 = "Möchten Sie diesen Eintrag wirklich kopieren?"

Set FM = frmAdress
Set TxKur = FM.txtS1F11

If GlAdN = False Then
    If TxKur.Text <> vbNullString Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage = 6 Then
            Adr_Kop
            DoEvents
            GlAdL = True 'Formular wird geladen
            ASper True
            ANeue
            GlAdS = False
            GlAdL = False
            Adr_Lad
            Kon_Lis
            DoEvents
            SUpAd True
        End If
    Else
        Mld1 = "Das Feld Suchname muß erst ausgefüllt werden, damit dieser Datensatz gespeichern werden kann"
        Tit1 = "Fehlende Angaben"
        WindowMess Mld1, Dial2, Tit1, FM.hwnd
    End If
Else
    Mld1 = "Die Adresse wurde noch nicht gespeichert"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AdKop " & Err.Number
Resume Next
  
End Sub
Public Sub AEmail(ByVal Flag As Integer)
On Error GoTo EmErr
'Versendet eine Emailnachricht

Dim BrAnr As String
Dim EmAd1 As String
Dim EmAd2 As String

Dim FeEm1 As XtremeSuiteControls.FlatEdit
Dim FeEm2 As XtremeSuiteControls.FlatEdit
Dim FeInt As XtremeSuiteControls.FlatEdit
Dim FeAnr As XtremeSuiteControls.ComboBox

Set FM = frmAdress
Set FeEm1 = FM.txtS1F19 'Email1
Set FeEm2 = FM.txtS2F34 'Email2
Set FeInt = FM.txtS1F27 'Internet
Set FeAnr = FM.cmbS1F10

BrAnr = FeAnr.Text
EmAd1 = FeEm1.Text
EmAd2 = FeEm2.Text

Set clFil = New clsFile

Select Case Flag
Case 1:
        If Len(EmAd1) > 0 Then WindowEml "mailto:" & EmAd1, "kein Betreff...", vbCrLf & BrAnr & vbCrLf & vbCrLf
Case 2:
        If Len(EmAd2) > 0 Then WindowEml "mailto:" & EmAd2, "kein Betreff...", vbCrLf & BrAnr & vbCrLf & vbCrLf
Case 3:
        If Len(FeInt.Text) > 0 Then
            With clFil
                .DaNam = FeInt.Text
                .DaPfa = App.Path
                .FilAusf
            End With
        End If
End Select

Set clFil = Nothing

Exit Sub

EmErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AEmail " & Err.Number
Resume Next

End Sub
Public Sub AEnab(ByVal MeEna As Boolean, Optional ByVal NoDel As Boolean)
On Error GoTo EmErr

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmAdress
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

If GlAdN = False Then
    CmAcs(AM_Patient_Speichern).Enabled = MeEna
Else
    CmAcs(AM_Patient_Speichern).Enabled = True
End If
If GlLiz = False Then
    CmAcs(AM_Drucken).Enabled = False
Else
    CmAcs(AM_Drucken).Enabled = MeEna
End If
CmAcs(AM_Patient_Gruppe).Enabled = MeEna
CmAcs(AM_Patient_Copy).Enabled = MeEna
CmAcs(AM_Patient_Del).Enabled = MeEna
CmAcs(AM_Patient_Clip1).Enabled = MeEna
CmAcs(AM_Patient_Clip2).Enabled = MeEna
CmAcs(AM_Notiz_Neu).Enabled = MeEna
CmAcs(AM_Notiz_Bearbeit).Enabled = MeEna
CmAcs(AM_Notiz_Loeschen).Enabled = MeEna
CmAcs(AM_Extras_Vorlage).Enabled = MeEna

If GlAdN = False Then
    CmAcs(AD_Patienten_Save).Enabled = MeEna
Else
    CmAcs(AD_Patienten_Save).Enabled = True
End If
If GlLiz = False Then
    CmAcs(AD_Adressen_Drucken).Enabled = False
Else
    CmAcs(AD_Adressen_Drucken).Enabled = MeEna
End If
CmAcs(AD_Patienten_Gruppe).Enabled = MeEna
CmAcs(AD_Patient_Copy).Enabled = MeEna
CmAcs(AD_Patient_Del).Enabled = MeEna
CmAcs(AD_Einzelbrief_Word).Enabled = MeEna
CmAcs(AD_Adressen_SMS).Enabled = MeEna
CmAcs(AD_Member_Add).Enabled = MeEna
CmAcs(AD_Member_Del).Enabled = MeEna
CmAcs(AD_Member_Copy).Enabled = MeEna
CmAcs(AD_Member_Save).Enabled = MeEna

If NoDel = True Then
    CmAcs(AD_Patient_Del).Enabled = False
End If

Set CmBrs = Nothing

Exit Sub

EmErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AEnab " & Err.Number
Resume Next

End Sub
Public Sub AErAd(Optional ByVal mAnDa As Boolean)
On Error GoTo ReErr
'Erstellt die Anschrift im Anschriftenfeld

Dim KGebo As String
Dim TagWe As String
Dim FIDKu As String
Dim RAnsh As String
Dim FAnsh As String
Dim KuEnd As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim EmSig As String
Dim FFirm As Variant
Dim FAnre As Variant
Dim FTite As Variant
Dim FName As Variant
Dim FVorn As Variant
Dim FStra As Variant
Dim FPOst As Variant
Dim FOrte As Variant
Dim FLand As Variant
Dim KFirm As Variant
Dim KAnre As Variant
Dim KTite As Variant
Dim KName As Variant
Dim KVorn As Variant
Dim KStra As Variant
Dim KPost As Variant
Dim KOrte As Variant
Dim KLand As Variant
Dim KNumm As Variant
Dim KTele As Variant

If mAnDa = True Then
    Set FM = frmMandant
Else
    Set FM = frmAdress
End If

FFirm = FM.txtS2F11.Text
FAnre = FM.txtS2F12.Text
FTite = FM.txtS2F13.Text
FVorn = FM.txtS2F14.Text
FName = FM.txtS2F15.Text
FStra = FM.txtS2F16.Text
FPOst = FM.txtS2F18.Text
FOrte = FM.txtS2F19.Text
FLand = FM.txtS2F22.Text
KFirm = FM.txtS1F01.Text
KAnre = FM.txtS1F02.Text
KTite = FM.txtS1F03.Text
KName = FM.txtS1F05.Text
KVorn = FM.txtS1F04.Text
KStra = FM.txtS1F06.Text
KPost = FM.txtS1F08.Text
KOrte = FM.txtS1F09.Text
KLand = FM.txtS1F12.Text
KNumm = FM.txtS1F30.Text
KTele = FM.txtS1F16.Text

EmSig = "Mit freundlichen Grüßen" & vbCrLf & vbCrLf

If IsDate(FM.txtS1F13.Text) Then
    KGebo = FM.txtS1F13.Text
Else
    KGebo = vbNullString
End If

If FTite <> vbNullString Then
    If Left$(FTite, 3) = "Dr." Then
        FTite = "Dr."
        EmSig = EmSig & FTite
    ElseIf Left$(FTite, 3) = "Prof" Then
        FTite = "Prof."
        EmSig = EmSig & FTite
    ElseIf Left$(FTite, 3) = "Dip" Then
        FTite = vbNullString
    ElseIf FTite = "HP" Or FTite = "Hp" Then
        FTite = vbNullString
        If FAnre = "Herrn und Frau" Then
            FAnre = "Herrn und Frau HP"
        ElseIf FAnre Like "*Herr*" Then
            FAnre = "Herrn HP"
        ElseIf FAnre Like "*Frau*" Then
            FAnre = "Frau HP"
        End If
    End If
End If

If KTite <> vbNullString Then
    If Left$(KTite, 3) = "Dr." Then
        KTite = "Dr."
    ElseIf Left$(KTite, 3) = "Prof" Then
        KTite = "Prof."
    ElseIf Left$(KTite, 3) = "Dip" Then
        KTite = vbNullString
    ElseIf KTite = "HP" Or KTite = "Hp" Then
        KTite = vbNullString
        If KAnre = "Herrn und Frau" Then
            KAnre = "Herrn und Frau HP"
        ElseIf KAnre Like "*Herr*" Then
            KAnre = "Herrn HP"
        ElseIf KAnre Like "*Frau*" Then
            KAnre = "Frau HP"
        End If
    End If
End If

'Kurbezeichung ###

FIDKu = vbNullString

Select Case GlAdK 'Adressenverkehrsnamens
Case 0:
    If Len(KGebo) > 1 Then
        KuEnd = Mid$(Trim$(KGebo), 1, 10)
    Else
        If Len(KOrte) > 1 Then
            KuEnd = Mid$(Trim$(KOrte), 1, 10)
        Else
            KuEnd = vbNullString
        End If
    End If
Case 1:
    If Len(KNumm) > 0 Then
        KuEnd = Mid$(Trim$(KNumm), 1, 10)
        KuEnd = Format$(KuEnd, "0000")
    Else
        KuEnd = vbNullString
    End If
Case 2:
    If Len(KOrte) > 1 Then
        KuEnd = Mid$(Trim$(KOrte), 1, 10)
    Else
        KuEnd = vbNullString
    End If
End Select

If Len(KName) > 1 Then
    If mAnDa = False Then
        If Len(KFirm) > 1 Then
            FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20) & ", " & "(" & KuEnd & ")"
        Else
            If Len(KuEnd) > 1 Then
                FIDKu = Mid$(Trim$(KName), 1, 25) & ", " & "(" & KuEnd & ")"
            Else
                FIDKu = Mid$(Trim$(KName), 1, 25)
            End If
        End If
    Else
        If Len(KuEnd) > 1 Then
            FIDKu = Mid$(Trim$(KName), 1, 25) & ", " & "(" & KuEnd & ")"
        Else
            FIDKu = Mid$(Trim$(KName), 1, 25)
        End If
    End If
Else
    If Len(KFirm) > 1 Then
        If Len(KuEnd) > 1 Then
            FIDKu = Mid$(Trim$(KFirm), 1, 25) & ", " & "(" & KuEnd & ")"
        Else
            If Len(KVorn) > 1 Then
                If Len(KuEnd) > 1 Then
                    FIDKu = Mid$(Trim$(KVorn), 1, 25) & ", " & "(" & KuEnd & ")"
                Else
                    FIDKu = Mid$(Trim$(KVorn), 1, 25)
                End If
            End If
        End If
    End If
End If

If Len(KVorn) > 1 Then
    If Len(KName) > 1 Then
        EmSig = EmSig & KVorn & Space$(1) & KName & vbCrLf
    Else
        EmSig = EmSig & KVorn & vbCrLf
    End If
    If Len(KFirm) > 1 Then
         EmSig = EmSig & KFirm & vbCrLf
    End If
ElseIf Len(KFirm) > 1 Then
    EmSig = EmSig & KFirm & vbCrLf
End If

If Len(KTele) > 8 Then
    EmSig = EmSig & KTele & vbCrLf
End If

If Len(KVorn) > 1 Then
    If Len(KName) > 1 Then
        If mAnDa = False Then
            If Len(KFirm) > 1 Then
                FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
            Else
                If Len(KuEnd) > 1 Then
                    FIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15) & ", " & "(" & KuEnd & ")"
                Else
                    FIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
                End If
            End If
        Else
            If Len(KuEnd) > 1 Then
                FIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15) & ", " & "(" & KuEnd & ")"
            Else
                FIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
            End If
        End If
    Else
        If mAnDa = False Then
            If Len(KFirm) > 1 Then
                If Len(KuEnd) > 1 Then
                    FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15) & ", " & "(" & KuEnd & ")"
                Else
                    FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
                End If
            Else
                If Len(KuEnd) > 1 Then
                    FIDKu = Mid$(Trim$(KVorn), 1, 30) & ", " & "(" & KuEnd & ")"
                Else
                    FIDKu = Mid$(Trim$(KVorn), 1, 30)
                End If
            End If
        Else
            If Len(KuEnd) > 1 Then
                FIDKu = Mid$(Trim$(KVorn), 1, 30) & ", " & "(" & KuEnd & ")"
            Else
                FIDKu = Mid$(Trim$(KVorn), 1, 30)
            End If
        End If
    End If
End If

'Rechnunmgsanschrift
If GlFZe = False Then
    If Len(FFirm) > 1 Then
        RAnsh = Trim$(FFirm)
    End If
End If

If GlAno = False Then
    If FAnre <> vbNullString Then
        If FFirm <> vbNullString Then
            If Not FAnre = "Firma" Then
                RAnsh = RAnsh & vbCrLf & Trim$(FAnre)
            Else
                RAnsh = RAnsh & vbCrLf
            End If
        Else
            If FAnre <> "Firma" Then
                RAnsh = RAnsh & Trim$(FAnre)
            Else
                RAnsh = RAnsh
            End If
        End If
    End If
End If

If Len(FFirm) > 1 Then
    If FTite <> vbNullString Then
        If GlAno = False Then
            RAnsh = RAnsh & Chr$(32) & Trim$(FTite)
        Else
            RAnsh = RAnsh & Trim$(FTite)
        End If
        If FVorn <> vbNullString Then
            RAnsh = RAnsh & Chr$(32) & Trim$(FVorn)
            If FName <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(FName)
            End If
        Else
            If FName <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(FName)
            End If
        End If
    Else
        If FVorn <> vbNullString Then
            If FAnre <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(FVorn)
            Else
                RAnsh = RAnsh & vbCrLf & Trim$(FVorn)
            End If
            If FName <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(FName)
            End If
        Else
            If FName <> vbNullString Then
                If GlAno = False Then
                    RAnsh = RAnsh & Chr$(32) & Trim$(FName)
                Else
                    RAnsh = RAnsh & Trim$(FName)
                End If
            End If
        End If
    End If
Else
    If FTite <> vbNullString Then
        RAnsh = RAnsh & vbCrLf & Trim$(FTite)
    End If
    If FVorn <> vbNullString Then
        If Not FTite = vbNullString Then
            RAnsh = RAnsh & Chr$(32) & Trim$(FVorn)
        Else
            RAnsh = RAnsh & vbCrLf & Trim$(FVorn)
        End If
    Else
        RAnsh = RAnsh & vbCrLf
    End If
    If FName <> vbNullString Then
        If Not FTite = vbNullString Or Not FVorn = vbNullString Then
            RAnsh = RAnsh & Chr$(32) & Trim$(FName)
        Else
            RAnsh = RAnsh & Trim$(FName)
        End If
    End If
End If

If GlFZe = True Then
    If Len(FFirm) > 1 Then
        RAnsh = RAnsh & vbCrLf & Trim$(FFirm)
    End If
End If

If FStra <> vbNullString Then RAnsh = RAnsh & vbCrLf & Trim$(FStra)
If FPOst <> vbNullString Then RAnsh = RAnsh & vbCrLf & Trim$(FPOst)
If FOrte <> vbNullString Then RAnsh = RAnsh & Chr$(32) & Trim$(FOrte)
If FLand <> vbNullString Then RAnsh = RAnsh & vbCrLf & UCase(Trim$(FLand))

'Patientenschaift
If GlFZe = False Then
    If Len(KFirm) > 1 Then
        FAnsh = Trim$(KFirm)
    End If
End If

If GlAno = False Then
    If KAnre <> vbNullString Then
        If KFirm <> vbNullString Then
            If Not KAnre = "Firma" Then
                FAnsh = FAnsh & vbCrLf & Trim$(KAnre)
            Else
                FAnsh = FAnsh & vbCrLf
            End If
        Else
            If KAnre <> "Firma" Then
                FAnsh = FAnsh & Trim$(KAnre)
            Else
                FAnsh = FAnsh
            End If
        End If
    End If
End If

If Len(KFirm) > 1 Then
    If KTite <> vbNullString Then
        If GlAno = False Then
            FAnsh = FAnsh & Chr$(32) & Trim$(KTite)
        Else
            FAnsh = FAnsh & Trim$(KTite)
        End If
        If KVorn <> vbNullString Then
            FAnsh = FAnsh & Chr$(32) & Trim$(KVorn)
            If KName <> vbNullString Then
                FAnsh = FAnsh & Chr$(32) & Trim$(KName)
            End If
        Else
            If KName <> vbNullString Then
                FAnsh = FAnsh & Chr$(32) & Trim$(KName)
            End If
        End If
    Else
        If KVorn <> vbNullString Then
            If KAnre <> vbNullString Then
                FAnsh = FAnsh & Chr$(32) & Trim$(KVorn)
            Else
                FAnsh = FAnsh & Trim$(KVorn)
            End If
            If KName <> vbNullString Then
                FAnsh = FAnsh & Chr$(32) & Trim$(KName)
            End If
        Else
            If KName <> vbNullString Then
                If GlAno = False Then
                    FAnsh = FAnsh & Chr$(32) & Trim$(KName)
                Else
                    FAnsh = FAnsh & Trim$(KName)
                End If
            End If
        End If
    End If
Else
    If KTite <> vbNullString Then
        FAnsh = FAnsh & vbCrLf & Trim$(KTite)
    End If
    If KVorn <> vbNullString Then
        If Not KTite = vbNullString Then
            FAnsh = FAnsh & Chr$(32) & Trim$(KVorn)
        Else
            FAnsh = FAnsh & vbCrLf & Trim$(KVorn)
        End If
    Else
        FAnsh = FAnsh & vbCrLf
    End If
    If KName <> vbNullString Then
        If Not KTite = vbNullString Or Not KVorn = vbNullString Then
            FAnsh = FAnsh & Chr$(32) & Trim$(KName)
        Else
            FAnsh = FAnsh & Trim$(KName)
        End If
    End If
End If

If GlFZe = True Then
    If Len(KFirm) > 1 Then
        FAnsh = FAnsh & vbCrLf & Trim$(KFirm)
    End If
End If

If KStra <> vbNullString Then FAnsh = FAnsh & vbCrLf & Trim$(KStra)
If KPost <> vbNullString Then FAnsh = FAnsh & vbCrLf & Trim$(KPost)
If KOrte <> vbNullString Then FAnsh = FAnsh & Chr$(32) & Trim$(KOrte)
If KLand <> vbNullString Then FAnsh = FAnsh & vbCrLf & UCase(Trim$(KLand))

If FIDKu <> FM.txtS1F11.Text Then
    If mAnDa = True Then
        If FM.txtS1F11.Text <> vbNullString Then
            Select Case GlBut
            Case RibTab_Mandanten:
                    TeTit = "Anzeigennamenänderung"
                    TeMai = "Soll der Anzeigename des Mandanten angepasst werden?"
                    TeInh = "Die Änderung an den Stammdaten des Mandanten macht eine Angleichung dessen Anzeigenamens erforderlich. Der Anzeigenahme ist für die interne Darstellung des Mandanten erforderlich. Soll diese angepasst werden?"
                    TeFus = "Vielen Einträgen, wie Rechnungen, Buchungen oder Terminen ist ein Mandant zugeordnet. Dieser Mandant wird durch dessen Anzeigenamen repräsentiert. Da es möglich ist, den Anzeigenamen manuell zu vergeben, wird an dieser Stelle nachgefragt, ob dieser verändert werden soll."
            Case RibTab_Mitarbeit:
                    TeTit = "Anzeigennamenänderung"
                    TeMai = "Soll der Anzeigename des Mitarbeiters angepasst werden?"
                    TeInh = "Die Änderung an den Stammdaten des Mitarbeiters macht eine Angleichung dessen Anzeigenamens erforderlich. Der Anzeigenahme ist für die interne Darstellung des Mitarbeiters erforderlich. Soll diese angepasst werden?"
                    TeFus = "Vielen Einträgen, wie Rechnungen, Buchungen oder Terminen ist ein Mitarbeiter zugeordnet. Dieser Mitarbeiter wird durch dessen Anzeigenamen repräsentiert. Da es möglich ist, den Anzeigenamen manuell zu vergeben, wird an dieser Stelle nachgefragt, ob dieser verändert werden soll."
            Case RibTab_Verordner:
                    TeTit = "Anzeigennamenänderung"
                    TeMai = "Soll der Anzeigename des Verordners angepasst werden?"
                    TeInh = "Die Änderung an den Stammdaten des Verordners macht eine Angleichung dessen Anzeigenamens erforderlich. Der Anzeigenahme ist für die interne Darstellung des Verordners erforderlich. Soll diese angepasst werden?"
                    TeFus = "Im Adressenerfassungsdialog ist es möglich, jedem Patienten einen Verordner zuzuordnen. Dieses gilt auch für den Rechnungen. Der Verordner wird durch dessen Anzeigename dargestellt."
            End Select
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
            If GlMes = 33565 Then
                TagWe = Mid$(FM.txtS1F11.Tag, 2, Len(FM.txtS1F11.Tag) - 1)
                FM.txtS1F11.Text = FIDKu
                FM.txtS1F11.Tag = 1 & TagWe
                GlAdS = True
            End If
        Else
            TagWe = Mid$(FM.txtS1F11.Tag, 2, Len(FM.txtS1F11.Tag) - 1)
            FM.txtS1F11.Text = FIDKu
            FM.txtS1F11.Tag = 1 & TagWe
            GlAdS = True
        End If
        If FM.txtEmSig.Text = vbNullString Then
            TagWe = Mid$(FM.txtEmSig.Tag, 2, Len(FM.txtEmSig.Tag) - 1)
            FM.txtEmSig.Text = EmSig
            FM.txtEmSig.Tag = 1 & TagWe
            GlAdS = True
        End If
    Else
        TagWe = Mid$(FM.txtS1F11.Tag, 2, Len(FM.txtS1F11.Tag) - 1)
        FM.txtS1F11.Text = FIDKu
        FM.txtS1F11.Tag = 1 & TagWe
        GlAdS = True
    End If
End If

If RAnsh <> FM.txtS3F01.Text Then
    TagWe = Mid$(FM.txtS3F01.Tag, 2, Len(FM.txtS3F01.Tag) - 1)
    FM.txtS3F01.Text = RAnsh
    FM.txtS3F01.Tag = 1 & TagWe
    GlAdS = True
End If

If mAnDa = False Then
    FM.txtAnsch.Text = FAnsh
End If

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AErAd " & Err.Number
Resume Next

End Sub
Public Function AErKu(ByVal KFirm As String, ByVal KName As String, ByVal KVorn As String, ByVal KNumm As String, Optional ByVal KTite As String) As String
On Error GoTo ReErr
'Erstellt die Kurzbezeichnung des Patienten

Dim FIDKu As String

FIDKu = vbNullString

If Len(KName) > 1 Then
    If Len(KFirm) > 1 Then
        FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20)
    Else
        FIDKu = Mid$(Trim$(KName), 1, 25)
        If KTite <> vbNullString Then
            FIDKu = KTite & Space$(1) & FIDKu
        End If
    End If
Else
    If Len(KFirm) > 1 Then
        FIDKu = Mid$(Trim$(KFirm), 1, 25)
    End If
End If

If Len(KVorn) > 1 Then
    If Len(KName) > 1 Then
        If Len(KFirm) > 1 Then
            FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
        Else
            FIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
        End If
        If KTite <> vbNullString Then
            FIDKu = KTite & Space$(1) & FIDKu
        End If
    Else
        If Len(KFirm) > 1 Then
            FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
        Else
            FIDKu = Mid$(Trim$(KVorn), 1, 30)
        End If
    End If
End If

AErKu = SUmw(FIDKu, False, False, True, True)

Exit Function

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AErKu " & Err.Number
Resume Next

End Function
Public Function AErSu(ByVal KFirm As String, ByVal KName As String, ByVal KVorn As String, ByVal KOrte As String, ByVal IdxNr As Long) As String
On Error GoTo ReErr
'Erstellt das Suchfeld des Patienten

Dim FIDKu As String
Dim KuEnd As String
Dim KNumm As String

If IdxNr > 0 Then
    KNumm = Format$(KNumm, "000000")
End If

FIDKu = vbNullString

Select Case GlAdK 'Adressenkurzbezeichung
Case 0:
    If Len(KOrte) > 1 Then
        KuEnd = Mid$(Trim$(KOrte), 1, 10)
    Else
        KuEnd = vbNullString
    End If
Case 1:
    If Len(KNumm) > 0 Then
        KuEnd = Mid$(Trim$(KNumm), 1, 10)
        KuEnd = Format$(KuEnd, "0000")
    Else
        KuEnd = vbNullString
    End If
Case 2:
    If Len(KOrte) > 1 Then
        KuEnd = Mid$(Trim$(KOrte), 1, 10)
    Else
        KuEnd = vbNullString
    End If
End Select

If Len(KName) > 1 Then
    If Len(KFirm) > 1 Then
        FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20) & ", " & "(" & KuEnd & ")"
    Else
        If Len(KuEnd) > 1 Then
            FIDKu = Mid$(Trim$(KName), 1, 25) & ", " & "(" & KuEnd & ")"
        Else
            FIDKu = Mid$(Trim$(KName), 1, 25)
        End If
    End If
Else
    If Len(KFirm) > 1 Then
        If Len(KuEnd) > 1 Then
            FIDKu = Mid$(Trim$(KFirm), 1, 25) & ", " & "(" & KuEnd & ")"
        Else
            If Len(KVorn) > 1 Then
                If Len(KuEnd) > 1 Then
                    FIDKu = Mid$(Trim$(KVorn), 1, 25) & ", " & "(" & KuEnd & ")"
                Else
                    FIDKu = Mid$(Trim$(KVorn), 1, 25)
                End If
            End If
        End If
    End If
End If

If Len(KVorn) > 1 Then
    If Len(KName) > 1 Then
        If Len(KFirm) > 1 Then
            FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
        Else
            If Len(KuEnd) > 1 Then
                FIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15) & ", " & "(" & KuEnd & ")"
            Else
                FIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
            End If
        End If
    Else
        If Len(KFirm) > 1 Then
            If Len(KuEnd) > 1 Then
                FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15) & ", " & "(" & KuEnd & ")"
            Else
                FIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
            End If
        Else
            If Len(KuEnd) > 1 Then
                FIDKu = Mid$(Trim$(KVorn), 1, 30) & ", " & "(" & KuEnd & ")"
            Else
                FIDKu = Mid$(Trim$(KVorn), 1, 30)
            End If
        End If
    End If
End If

AErSu = FIDKu

Exit Function

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AErSu " & Err.Number
Resume Next

End Function
Public Function AErZu(ByVal IdTyp As Integer) As String
On Error GoTo ReErr
'Erstellt die Kurzbezeichnung im Anschriftenfeld

Dim KuEnd As String
Dim KGebo As String
Dim KIDKu As String
Dim KFirm As Variant
Dim KAnre As Variant
Dim KTite As Variant
Dim KName As Variant
Dim KVorn As Variant
Dim KStra As Variant
Dim KPost As Variant
Dim KOrte As Variant
Dim KLand As Variant

Select Case IdTyp
Case 1: Set FM = frmAdress
Case 2: Set FM = frmTermin
End Select

KFirm = FM.txtS4F01.Text
KAnre = FM.txtS4F02.Text
KTite = FM.txtS4F03.Text
KName = FM.txtS4F05.Text
KVorn = FM.txtS4F04.Text
KStra = FM.txtS4F06.Text
KPost = FM.txtS4F08.Text
KOrte = FM.txtS4F09.Text
KLand = FM.cmbS4F12.Text

If IsDate(FM.txtS4F18.Text) Then
    KGebo = FM.txtS4F18.Text
Else
    KGebo = vbNullString
End If

If KTite <> vbNullString Then
    If Left$(KTite, 3) = "Dr." Then
        KTite = "Dr."
    ElseIf Left$(KTite, 3) = "Prof" Then
        KTite = "Prof."
    ElseIf Left$(KTite, 3) = "Dip" Then
        KTite = vbNullString
    ElseIf KTite = "HP" Or KTite = "Hp" Then
        KTite = vbNullString
        If KAnre = "Herrn und Frau" Then
            KAnre = "Herrn und Frau HP"
        ElseIf KAnre Like "*Herr*" Then
            KAnre = "Herrn HP"
        ElseIf KAnre Like "*Frau*" Then
            KAnre = "Frau HP"
        End If
    End If
End If

'Kurbezeichung ###

KIDKu = vbNullString

Select Case GlAdK 'Adressenkurzbezeichung
Case 0:
    If Len(KGebo) > 1 Then
        KuEnd = Mid$(Trim$(KGebo), 1, 10)
    Else
        If Len(KOrte) > 1 Then
            KuEnd = Mid$(Trim$(KOrte), 1, 10)
        Else
            KuEnd = vbNullString
        End If
    End If
Case 1:
    KuEnd = vbNullString
Case 2:
    If Len(KOrte) > 1 Then
        KuEnd = Mid$(Trim$(KOrte), 1, 10)
    Else
        KuEnd = vbNullString
    End If
End Select

If Len(KName) > 1 Then
    If Len(KFirm) > 1 Then
        KIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20) & ", " & "(" & KuEnd & ")"
    Else
        If Len(KuEnd) > 1 Then
            KIDKu = Mid$(Trim$(KName), 1, 25) & ", " & "(" & KuEnd & ")"
        Else
            KIDKu = Mid$(Trim$(KName), 1, 25)
        End If
    End If
Else
    If Len(KFirm) > 1 Then
        If Len(KuEnd) > 1 Then
            KIDKu = Mid$(Trim$(KFirm), 1, 25) & ", " & "(" & KuEnd & ")"
        Else
            If Len(KVorn) > 1 Then
                If Len(KuEnd) > 1 Then
                    KIDKu = Mid$(Trim$(KVorn), 1, 25) & ", " & "(" & KuEnd & ")"
                Else
                    KIDKu = Mid$(Trim$(KVorn), 1, 25)
                End If
            End If
        End If
    End If
End If

If Len(KVorn) > 1 Then
    If Len(KName) > 1 Then
        If Len(KFirm) > 1 Then
            KIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
        Else
            If Len(KuEnd) > 1 Then
                KIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15) & ", " & "(" & KuEnd & ")"
            Else
                KIDKu = Mid$(Trim$(KName), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
            End If
        End If
    Else
        If Len(KFirm) > 1 Then
            If Len(KuEnd) > 1 Then
                KIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15) & ", " & "(" & KuEnd & ")"
            Else
                KIDKu = Mid$(Trim$(KFirm), 1, 20) & ", " & Mid$(Trim$(KVorn), 1, 15)
            End If
        Else
            If Len(KuEnd) > 1 Then
                KIDKu = Mid$(Trim$(KVorn), 1, 30) & ", " & "(" & KuEnd & ")"
            Else
                KIDKu = Mid$(Trim$(KVorn), 1, 30)
            End If
        End If
    End If
End If

AErZu = KIDKu
GlAdZ = True

Exit Function

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AErZu " & Err.Number
Resume Next

End Function
Public Sub AFont(ByVal FoNam As Form)
On Error GoTo ReErr

Set FM = FoNam

For Each AktCo In FM.Controls
    Select Case TypeName(AktCo)
    Case "GroupBox":
            With AktCo
                .Font.Name = GlTFt.Name
                .Font.SIZE = 8
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "PushButton":
            With AktCo
                .Font.Name = GlTFt.Name
                .Font.SIZE = 8
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "Label":
            With AktCo
                .Font.Name = GlTFt.Name
                .Font.SIZE = 8
            End With
    Case "RadioButton":
            With AktCo
                .Font.Name = GlTFt.Name
                .Font.SIZE = 8
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "CheckBox":
            With AktCo
                .Font.Name = GlTFt.Name
                .Font.SIZE = 8
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    Case "FlatEdit":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
                .Font.Name = GlTFt.Name
                If GlTFt.SIZE > 11 Then
                    .Font.SIZE = 11
                Else
                    .Font.SIZE = GlTFt.SIZE
                End If
                .ForeColor = -2147483641
            End With
    Case "ComboBox":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
                .Font.Name = GlTFt.Name
                If GlTFt.SIZE > 11 Then
                    .Font.SIZE = 11
                Else
                    .Font.SIZE = GlTFt.SIZE
                End If
                .ForeColor = -2147483641
            End With
    Case "ListBox":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
                .Font.Name = GlTFt.Name
                If GlTFt.SIZE > 11 Then
                    .Font.SIZE = 11
                Else
                    .Font.SIZE = GlTFt.SIZE
                End If
                .ForeColor = -2147483641
            End With
    Case "UpDown":
            With AktCo
                Select Case GlSty
                Case 8: .Appearance = xtpAppearanceOffice2013
                Case 7: .Appearance = xtpAppearanceOffice2013
                Case Else: .Appearance = xtpAppearanceResource
                End Select
            End With
    End Select
Next AktCo

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AFont " & Err.Number
Resume Next

End Sub
Public Sub AGebu(Optional ByVal BeDat As Boolean)
On Error GoTo CrErr
'Löscht die Geburtsdaten in der Adresseneingabemaske

Dim TagWe As String
Dim Mld1, Tit1 As String
Dim Frage As Integer

If BeDat = True Then
    Set FM = frmMandant
Else
    Set FM = frmAdress
End If

Set TxGeb = FM.txtS1F13
Set TxReG = FM.txtS2F25

Tit1 = "Geburtsdatum Entfernen"
Mld1 = "Möchten Sie die Geburtsdaten jetzt wirklich löschen?"

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    TxGeb.Text = vbNullString
    TagWe = Mid$(TxGeb.Tag, 2, Len(TxGeb.Tag) - 1)
    TxGeb.Tag = 1 & TagWe
    
    TxReG.Text = vbNullString
    TagWe = Mid$(TxReG.Tag, 2, Len(TxReG.Tag) - 1)
    TxReG.Tag = 1 & TagWe
    
    GlAdS = True
End If

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AGebu " & Err.Number
Resume Next

End Sub
Public Sub AGuth()
On Error GoTo CrErr
'Löscht das Guthaben in der Adresseneingabemaske

Dim TagWe As String
Dim Mld1, Tit1 As String
Dim Frage As Integer

Set FM = frmAdress
Set TxGeb = FM.txtS1F33

Tit1 = "Guthaben Entfernen"
Mld1 = "Möchten Sie das Guthaben jetzt wirklich löschen?"

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    TxGeb.Text = Format$(0, GlWa1)
    TagWe = Mid$(TxGeb.Tag, 2, Len(TxGeb.Tag) - 1)
    TxGeb.Tag = 1 & TagWe
    GlAdS = True
End If

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AGuth " & Err.Number
Resume Next

End Sub

Private Sub AInit()
On Error GoTo InErr
'Initialisierung der Steuerelemente

Dim AktZa As Integer
Dim PoTim As Integer
Dim ZeiUm As Boolean
Dim LiTip As Boolean
Dim FeEm1 As XtremeSuiteControls.FlatEdit
Dim FeEm2 As XtremeSuiteControls.FlatEdit
Dim FeInt As XtremeSuiteControls.FlatEdit
Dim CmBre As XtremeSuiteControls.ComboBox
Dim CmTyp As XtremeSuiteControls.ComboBox
Dim CmLan As XtremeSuiteControls.ComboBox
Dim CmVrs As XtremeSuiteControls.ComboBox
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim PuBu1 As XtremeSuiteControls.PushButton
Dim PuBu2 As XtremeSuiteControls.PushButton
Dim PuBu3 As XtremeSuiteControls.PushButton
Dim PuBu4 As XtremeSuiteControls.PushButton
Dim PuBu5 As XtremeSuiteControls.PushButton
Dim PuBu6 As XtremeSuiteControls.PushButton
Dim PuBu7 As XtremeSuiteControls.PushButton
Dim PuBu8 As XtremeSuiteControls.PushButton
Dim PuBu9 As XtremeSuiteControls.PushButton
Dim PuPo1 As XtremeSuiteControls.PushButton
Dim PuPo2 As XtremeSuiteControls.PushButton
Dim PuPo3 As XtremeSuiteControls.PushButton
Dim PuGut As XtremeSuiteControls.PushButton

Set FM = frmAdress
Set S1L13 = FM.lblS1L13
Set S1L20 = FM.lblS2L20
Set S2L34 = FM.lblS2L34
Set S2L35 = FM.lblS2L35
Set S3F02 = FM.txtS3F02
Set S3F03 = FM.txtS3F03
Set TxNum = FM.txtS1F30
Set FeAn1 = FM.txtS1F02
Set FeAn2 = FM.txtS2F12
Set FeAn3 = FM.txtS4F02
Set FeFam = FM.txtS2F26
Set FeEm1 = FM.txtS1F19 'Email1
Set FeEm2 = FM.txtS2F34 'Email2
Set FeInt = FM.txtS1F27 'Internet
Set FeTi1 = FM.cmbS1F21
Set FeTi2 = FM.cmbS1F22
Set FeTi3 = FM.cmbS1F23
Set FeTi4 = FM.cmbS1F24
Set FeTi5 = FM.cmbS1F25
Set FeTi6 = FM.cmbS1F26
Set FeTi7 = FM.cmbS1F27
Set FePat = FM.cmbS2F10
Set FeKat = FM.cmbS1F06
Set FeTar = FM.cmbS1F07
Set FeBeh = FM.txtS2F08
Set TxGeb = FM.txtS1F13
Set TxZGe = FM.txtS4F18
Set TxReG = FM.txtS2F25
Set TxErs = FM.txtS2F27
Set TxKop = FM.txtS1F14
Set FeGes = FM.cmbS1F08
Set TxTe1 = FM.txtS1F15
Set TxTe2 = FM.txtS1F16
Set TxTe3 = FM.txtS1F17
Set TxTe4 = FM.txtS1F18
Set CmLan = FM.txtS1F22
Set CmBre = FM.cmbS1F10
Set CmTyp = FM.cmbS2F36
Set CmArt = FM.cmbS2F30
Set CmVrs = FM.cmbS1F28
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set PuBu1 = FM.btnTele1
Set PuBu2 = FM.btnTele2
Set PuBu3 = FM.btnTele3
Set PuBu4 = FM.btnTele4
Set PuBu5 = FM.btnTele5
Set PuBu6 = FM.btnTele6
Set PuBu7 = FM.btnTele7
Set PuBu8 = FM.btnTele8
Set PuBu9 = FM.btnTele9
Set PuPo1 = FM.btnPost1
Set PuPo2 = FM.btnPost2
Set PuPo3 = FM.btnPost3
Set PuGut = FM.btnGutha
Set PrGr1 = FM.prpGrid1
Set PrGr2 = FM.prpGrid2
Set PrGr3 = FM.prpGrid3
Set ImMan = frmMain.imgManag

PoTim = IniGetVal("System", "PopTim") * 1000
S2L34.Caption = IniGetVal("Layout", "AdTit1") & " :"
S2L35.Caption = IniGetVal("Layout", "AdTit2") & " :"
LiTip = CBool(IniGetVal("Layout", "GrdTip"))
ZeiUm = False

With RpCo1
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips LiTip
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Notizen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Notizen vorhanden"
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
    If ZeiUm = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
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
    .PreviewMode = True
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo2
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Buchungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Buchungen vorhanden"
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
    If ZeiUm = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
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
    .SelectionEnable = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo3
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine offenen Posten vorhanden"
    .PaintManager.NoItemsText = "Es sind keine offenen Posten vorhanden"
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
    If ZeiUm = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
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
    .SelectionEnable = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo4
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Termine vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Termine vorhanden"
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
    If ZeiUm = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
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
    .SelectionEnable = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo5
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = False 'GlSpS
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = False 'WICHTIG!
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    '.OverrideThemeMetrics = True  no matter what this function do !?
    .MultipleSelection = True
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    '.SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind keine Zugehörigen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Zugehörigen vorhanden"
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
    If ZeiUm = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.ColumnWidthWYSIWYG = False
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .MultiSelectionMode = False
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = GlGKo
    .SelectionEnable = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

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
    .BorderStyle = xtpGridBorderStaticEdge
    .HelpBackColor = -2147483643
    .HelpForeColor = -2147483640
    .HighlightChangedItems = False
    .HideSelection = True
    .HelpVisible = False
    .Font.Name = GlTFt.Name
    .Font.SIZE = GlTFt.SIZE
    .LockRedraw = False
    .NavigateItems = True
    .PropertySort = NoSort
    .ShowInplaceButtonsAlways = False
    .SplitterPos = 0.45
    .ToolBarVisible = False
    .VariableSplitterPos = False
    .VariableItemsHeight = True
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
    .BorderStyle = xtpGridBorderStaticEdge
    .HelpBackColor = -2147483643
    .HelpForeColor = -2147483640
    .HighlightChangedItems = False
    .HideSelection = True
    .HelpVisible = False
    .Font.Name = GlTFt.Name
    .Font.SIZE = GlTFt.SIZE
    .LockRedraw = False
    .NavigateItems = True
    .PropertySort = NoSort
    .ShowInplaceButtonsAlways = False
    .SplitterPos = 0.4
    .ToolBarVisible = False
    .VariableSplitterPos = False
    .VariableItemsHeight = True
    .ViewBackColor = -2147483643
    .ViewCategoryForeColor = -2147483640
    .ViewForeColor = -2147483640
    .ViewReadOnlyForeColor = 8421504
    .Verbs.Clear
End With

With PrGr3
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
    .BorderStyle = xtpGridBorderStaticEdge
    .HelpBackColor = -2147483643
    .HelpForeColor = -2147483640
    .HighlightChangedItems = False
    .HideSelection = True
    .HelpVisible = False
    .Font.Name = GlTFt.Name
    .Font.SIZE = GlTFt.SIZE
    .LockRedraw = False
    .NavigateItems = True
    .PropertySort = NoSort
    .ShowInplaceButtonsAlways = False
    .SplitterPos = 0.4
    .ToolBarVisible = False
    .VariableSplitterPos = False
    .VariableItemsHeight = True
    .ViewBackColor = -2147483643
    .ViewCategoryForeColor = -2147483640
    .ViewForeColor = -2147483640
    .ViewReadOnlyForeColor = 8421504
    .Verbs.Clear
End With

With FeAn1 'Anreden
    For AktZa = 0 To UBound(GlAnr) - 1
        .AddItem GlAnr(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With FeAn2
    For AktZa = 0 To UBound(GlAnr) - 1
        .AddItem GlAnr(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With FeAn3
    For AktZa = 0 To UBound(GlAnr) - 1
        .AddItem GlAnr(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With FeGes 'Geschlecht
    For AktZa = 0 To UBound(GlGes) - 1
        .AddItem GlGes(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With FeTi1
    For AktZa = 0 To UBound(GlTeL) - 1
        .AddItem GlTeL(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 12
End With

With FeTi2
    For AktZa = 0 To UBound(GlTeL) - 1
        .AddItem GlTeL(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 12
End With

With FeTi3
    For AktZa = 0 To UBound(GlTeL) - 1
        .AddItem GlTeL(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 12
End With

With FeTi4
    For AktZa = 0 To UBound(GlTeL) - 1
        .AddItem GlTeL(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 12
End With

With FeTi5
    For AktZa = 0 To UBound(GlTeL) - 1
        .AddItem GlTeL(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 12
End With

With FeTi6
    For AktZa = 0 To UBound(GlTeL) - 1
        .AddItem GlTeL(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 12
End With

With FeTi7
    For AktZa = 0 To UBound(GlTeL) - 1
        .AddItem GlTeL(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 12
End With

With FeBeh
    For AktZa = 0 To UBound(GlBeh) - 1
        .AddItem GlBeh(AktZa, 0)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
    .DropDownItemCount = 13
End With

With FeFam
    .AddItem "Ledig"
    .ItemData(0) = 1
    .AddItem "Verheiratet"
    .ItemData(1) = 2
    .AddItem "Verwitwet"
    .ItemData(2) = 3
    .AddItem "Geschieden"
    .ItemData(3) = 4
    .AddItem "Getrennt"
    .ItemData(4) = 5
    .AddItem "Unbekannt"
    .ItemData(5) = 6
End With

With CmVrs
    .AddItem "Postversand"
    .ItemData(0) = 0
    .AddItem "Emailversand"
    .ItemData(1) = 1
    .AddItem "Downloadlink"
    .ItemData(2) = 2
    .ListIndex = GlRVs - 1 'Standard-Rechnungsversandweg
End With

If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    S1L13.Caption = "Gesetz :"
    S1L20.Caption = "Kanton :"
    With CmLan
        For AktZa = 0 To UBound(GlKtn)
            .AddItem GlKtn(AktZa, 0)
            .ItemData(AktZa) = GlKtn(AktZa, 2)
        Next AktZa
    End With
Else
    S1L13.Caption = "Vertragsart :"
    S1L20.Caption = "Bundesland :"
    With CmLan
        For AktZa = 0 To UBound(GlBsl)
            .AddItem GlBsl(AktZa, 0)
            .ItemData(AktZa) = GlBsl(AktZa, 2)
        Next AktZa
    End With
End If

FeEm1.ForeColor = vbBlue
FeEm2.ForeColor = vbBlue
FeInt.ForeColor = vbBlue

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
S2L34.BackColor = GlBak
S2L35.BackColor = GlBak
FM.chkOpti1.BackColor = GlBak
FM.chkOpti2.BackColor = GlBak
FM.chkOpti3.BackColor = GlBak

With FeKat
    .AutoComplete = False
    .DropDownItemCount = 15
End With

With FeTar
    .AutoComplete = False
    .DropDownItemCount = 15
End With

FePat.Enabled = True
S3F02.Font.Name = GlTFt.Name
S3F02.Font.SIZE = GlTFt.SIZE
S3F03.Font.Name = GlTFt.Name
S3F03.Font.SIZE = GlTFt.SIZE

With TxKop
    .Pattern = "\d*"
    .SetMask "0", "_"
End With

With CmTyp
    .AddItem "Privat Inland"
    .ItemData(0) = 1
    .AddItem "Privat Europa"
    .ItemData(1) = 2
    .AddItem "Privat Ausland"
    .ItemData(2) = 3
    .AddItem "Gewerb. Inland"
    .ItemData(3) = 4
    .AddItem "Gewerb. Europa"
    .ItemData(4) = 5
    .AddItem "Gewerb. Ausland"
    .ItemData(5) = 6
End With

TxNum.Pattern = "\d*"
CmBre.DropDownItemCount = 6

TxGeb.SetMask "00.00.0000", "__.__.____"
TxZGe.SetMask "00.00.0000", "__.__.____"
TxReG.SetMask "00.00.0000", "__.__.____"
TxErs.SetMask "00.00.0000", "__.__.____"
TxZGe.Text = vbNullString

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Telephone, 16)
If GlRDP = True Then PuBu1.Enabled = False
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Telephone, 16)
If GlRDP = True Then PuBu2.Enabled = False
PuBu3.Icon = ImMan.Icons.GetImage(IC16_Telephone, 16)
If GlRDP = True Then PuBu3.Enabled = False
PuBu4.Icon = ImMan.Icons.GetImage(IC16_Phone_Mobil, 16)
PuBu5.Icon = ImMan.Icons.GetImage(IC16_Earth_Mail, 16)
PuBu7.Icon = ImMan.Icons.GetImage(IC16_Earth_Mail, 16)
PuBu6.Icon = ImMan.Icons.GetImage(IC16_Earth_View, 16)
PuBu8.Icon = ImMan.Icons.GetImage(IC16_Telephone, 16)
If GlRDP = True Then PuBu8.Enabled = False
PuBu9.Icon = ImMan.Icons.GetImage(IC16_Earth_Mail, 16)
PuPo1.Icon = ImMan.Icons.GetImage(IC16_Mailbox, 16)
PuPo2.Icon = ImMan.Icons.GetImage(IC16_Mailbox, 16)
PuPo3.Icon = ImMan.Icons.GetImage(IC16_Mailbox, 16)
PuGut.Icon = ImMan.Icons.GetImage(IC16_Money_Coins, 16)

Set PrGr1 = Nothing
Set PrGr2 = Nothing
Set PrGr3 = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set TabCo = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AInit " & Err.Number
Resume Next

End Sub
Public Function AJahr(ByVal GebDa As Date) As Integer
On Error GoTo InErr
'Berechnet das Geburtsdatum in Jahre

Dim AnzJa As Integer

AnzJa = Abs(DateDiff("yyyy", GebDa, Now()))
If Month(GebDa) > Month(Now) Then
    AnzJa = AnzJa - 1
ElseIf Month(GebDa) = Month(Now) Then
    If Day(GebDa) > Day(Now) Then AnzJa = AnzJa - 1
End If
  
AJahr = AnzJa

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AJahr " & Err.Number
Resume Next
  
End Function
Public Sub AKaSt()
On Error GoTo LiErr

Dim GesZa As Long
Dim TmStr As String
Dim AktZa As Integer
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrBol As XtremePropertyGrid.PropertyGridItemBool
Dim PrDat As XtremePropertyGrid.PropertyGridItemDate

Set FM = frmAdress
Set PrGr1 = FM.prpGrid1
Set PrGr2 = FM.prpGrid2

Set PrKat = PrGr1.AddCategory("Abrechnungsdaten")
PrKat.id = 1100
PrKat.Expandable = False
PrKat.Expanded = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Versicherung Anschrift :", vbNullString)
PrItm.Tag = "0Versicherung"
PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
PrItm.MultiLinesCount = 3

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Vertragsnummer :", vbNullString)
PrItm.Tag = "0Zusatz"
PrItm.EditStyle = EditStyleLeft

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Versichertennummer :", vbNullString)
PrItm.Tag = "0Kartennummer"
PrItm.EditStyle = EditStyleLeft

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Karte gültig bis :", vbNullString)
PrItm.Tag = "0Kartengultig"
PrItm.EditStyle = EditStyleLeft

If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "HAV-Nr.:", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Krankenkassennummer :", vbNullString)
End If
PrItm.Tag = "0KVNummer"
PrItm.EditStyle = EditStyleLeft

If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "VEKA-Nr.:", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Kartenstatus :", vbNullString)
End If
PrItm.Tag = "0Kartenstatus"
PrItm.EditStyle = EditStyleLeft

If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Vergütungsart :", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Abrechnungstyp :", vbNullString)
End If
PrItm.Tag = "0AbrTyp"
PrItm.flags = ItemHasComboButton
If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    PrItm.Constraints.Add "TG - Tiers Garant"
    PrItm.Constraints.Add "TP - Tiers Payant"
Else
    PrItm.Constraints.Add "K - Kassenpatient"
    PrItm.Constraints.Add "P - Privatpatient"
    PrItm.Constraints.Add "X - andere Rechnungsempfänger"
    PrItm.Constraints.Add "E - Einsender (persönlich)"
End If

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Versichertenart :", vbNullString)
PrItm.Tag = "0Versichertenart"
PrItm.flags = ItemHasComboButton
PrItm.Constraints.Add "1 - Mitglied"
PrItm.Constraints.Add "3 - Familienversichert"
PrItm.Constraints.Add "5 - Rentner"

If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "Kanton :", vbNullString)
Else
    Set PrItm = PrKat.AddChildItem(PropertyItemString, "KV-Bereich :", vbNullString)
End If
PrItm.Tag = "0KVBereich"
PrItm.flags = ItemHasComboButton
If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    For AktZa = 0 To UBound(GlKtn) 'Kantone
        PrItm.Constraints.Add Format$(AktZa, "00") & " - " & GlKtn(AktZa, 0)
    Next AktZa
Else
    For AktZa = 0 To UBound(GlKVB) 'KV Bezirke
        PrItm.Constraints.Add GlKVB(AktZa, 1) & " - " & GlKVB(AktZa, 0)
    Next AktZa
End If

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Globally Unique Identifier :", vbNullString)
PrItm.Tag = "0GuiID"
PrItm.EditStyle = EditStyleLeft

'---

Set PrKat = PrGr1.AddCategory("Dateiverschlüsselung")
PrKat.id = 1200
PrKat.Expandable = False
PrKat.Expanded = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Verschlüsselungskennwort :", vbNullString)
PrItm.Tag = "0Em_Pass"
PrItm.EditStyle = EditStyleLeft
PrItm.PasswordMask = True
PrItm.id = 9901

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Verschlüsselung aktiviert :", False)
PrBol.CheckBoxStyle = True
PrBol.Tag = "0Em_Aut"

Set PrBol = PrKat.AddChildItem(PropertyItemBool, "Passwort anzeigen :", False)
PrBol.CheckBoxStyle = True
PrBol.Tag = "0Em_User"
PrBol.id = 9902

'---

Set PrKat = PrGr1.AddCategory("Schwangerschaft")
PrKat.id = 1300
PrKat.Expandable = False
PrKat.Expanded = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Anzahl Geburten :", vbNullString)
PrItm.Tag = "0Anz_Geburten"
PrItm.EditStyle = EditStyleNumber

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Anzahl Kinder :", vbNullString)
PrItm.Tag = "0Anz_Kinder"
PrItm.EditStyle = EditStyleNumber

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Anzahl Schwangerschaft. :", vbNullString)
PrItm.Tag = "0Anz_Schwanger"
PrItm.EditStyle = EditStyleNumber

'---

Set PrKat = PrGr1.AddCategory("Nationalität")
PrKat.id = 1400
PrKat.Expandable = False
PrKat.Expanded = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Nationalität :", vbNullString)
PrItm.Tag = "0Nationalität"
PrItm.EditStyle = EditStyleLeft

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Muttersprache :", vbNullString)
PrItm.Tag = "0Muttersprache"
PrItm.EditStyle = EditStyleLeft

'---

Set PrKat = PrGr2.AddCategory("Dokumentation")
PrKat.id = 1500
PrKat.Expandable = False
PrKat.Expanded = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Allergien :", vbNullString)
PrItm.Tag = "0Allergien"
PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
PrItm.MultiLinesCount = 4

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Risikofaktoren :", vbNullString)
PrItm.Tag = "0Risikofaktoren"
PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
PrItm.MultiLinesCount = 4

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Dauertherapie :", vbNullString)
PrItm.Tag = "0Dauertherapie"
PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
PrItm.MultiLinesCount = 4

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Unfälle / Arbeitsunfälle :", vbNullString)
PrItm.Tag = "0Unfälle"
PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
PrItm.MultiLinesCount = 4

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Operationen :", vbNullString)
PrItm.Tag = "0Operationen"
PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
PrItm.MultiLinesCount = 4

'---

Set PrKat = PrGr2.AddCategory("Hausarzt")
PrKat.id = 1600
PrKat.Expandable = False
PrKat.Expanded = True

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Hausarztadresse :", vbNullString)
PrItm.Tag = "0Hausarzt"
PrItm.CaptionMetrics.DrawTextFormat = DrawTextVcenter
PrItm.EditStyle = EditStyleWantReturn Or EditStyleMultiLine Or EditStyleVScroll
PrItm.MultiLinesCount = 3

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Hausarzt-Nr. :", vbNullString)
PrItm.Tag = "0HausarztNr"
PrItm.EditStyle = EditStyleLeft

Set PrItm = PrKat.AddChildItem(PropertyItemString, "Arbeitgeber :", vbNullString)
PrItm.Tag = "0Arbeitgeber"
PrItm.EditStyle = EditStyleLeft

PrGr1.PropertySort = Categorized
PrGr2.PropertySort = Categorized

Set PrKat = Nothing
Set PrGr1 = Nothing
Set PrGr2 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AKaSt " & Err.Number
Resume Next

End Sub
Public Sub Akont(ByVal Flag As Boolean)
On Error GoTo OpErr
'Öffnet das Kontaktformular

Dim AnzPo As Long
Dim IdStr As String
Dim Mld1, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmAdress
Set FS = frmKontakt
Set TxBeH = FS.txtBehan
Set TxKur = FM.txtS1F11
Set RpCon = FM.repCont1
Set ImMan = frmMain.imgManag
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

GlKoL = True
GlKoS = False

If Flag = True Then
    GlKoN = True
    If TxKur.Text <> vbNullString Then
        FS.Caption = "Notiz für: " & TxKur.Text
        TxBeH.Text = GlMiA(GlSmI, 1)
        TxBeH.Tag = 1 & "Behandler"
        frmKontakt.Show vbModal
    Else
        Mld1 = "Es wurde noch keine Adresse ausgewählt oder angelegt"
        Tit1 = "Keine Adresse"
        SPopu Tit1, Mld1, IC48_Forbidden
    End If
Else
    GlKoN = False
    AnzPo = RpCon.Records.Count
    If AnzPo > 0 Then
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            Set RpCol = RpCls.Find(Not_ID2)
            IdStr = RpRow.Record(RpCol.ItemIndex).Value
            Kon_Lad IdStr
            FS.Caption = "Notiz für: " & TxKur.Text
            frmKontakt.Show vbModal
            GlKoS = False
        End If
    Else
        Mld1 = "Es ist kein Notizeintrag vorhanden den Sie öffnen könnten. Legen Sie einen neuen Notizeintrag an"
        Tit1 = "Kein Notizeintrag"
        SPopu Tit1, Mld1, IC48_Forbidden
    End If
End If

GlKoL = False

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AKont " & Err.Number
Resume Next

End Sub
Public Sub AKopi()
On Error GoTo CrErr
'Kopieren der Adressdaten in die Rechungsanschrift

Dim TagWe As String
Dim AdRGe As Boolean
Dim Mld1, Tit1 As String

Set FM = frmAdress

Tit1 = "Adressänderung"
Mld1 = "HINWEIS! Die Anschrift des Rechnungsempfängers ist nicht identisch mit der Anschrift des Patienten."

If FM.txtS2F12.Text = vbNullString Then
    If FM.txtS1F02.Text <> vbNullString Then
        TagWe = Mid$(FM.txtS2F12.Tag, 2, Len(FM.txtS2F12.Tag) - 1)
        FM.txtS2F12.Text = FM.txtS1F02.Text
        FM.txtS2F12.Tag = 1 & TagWe
    End If
End If

If FM.txtS2F12.Text = FM.txtS1F02.Text Then
    If FM.txtS2F13.Text = vbNullString Then If FM.txtS1F03.Text <> vbNullString Then FM.txtS2F13.Text = FM.txtS1F03.Text
    If FM.txtS2F14.Text = vbNullString Then If FM.txtS1F04.Text <> vbNullString Then FM.txtS2F14.Text = FM.txtS1F04.Text
    If FM.txtS2F15.Text = vbNullString Then If FM.txtS1F05.Text <> vbNullString Then FM.txtS2F15.Text = FM.txtS1F05.Text
End If

If FM.txtS2F11.Text = vbNullString Then If FM.txtS1F01.Text <> vbNullString Then FM.txtS2F11.Text = FM.txtS1F01.Text
If FM.txtS2F20.Text = vbNullString Then If FM.cmbS1F10.Text <> vbNullString Then FM.txtS2F20.Text = FM.cmbS1F10.Text
If FM.txtS2F22.Text = vbNullString Then If FM.txtS1F12.Text <> vbNullString Then FM.txtS2F22.Text = FM.txtS1F12.Text
If FM.txtS2F25.Text = vbNullString Then If FM.txtS1F13.Text <> vbNullString Then FM.txtS2F25.Text = FM.txtS1F13.Text

If FM.txtS2F16.Text = vbNullString Then
    If FM.txtS1F06.Text <> vbNullString Then
        FM.txtS2F16.Text = FM.txtS1F06.Text
    End If
Else
    If GlAdS = True Then
        If FM.txtS2F16.Text <> FM.txtS1F06.Text Then
            AdRGe = True
        End If
    End If
End If

If FM.txtS2F18.Text = vbNullString Then
    If FM.txtS1F08.Text <> vbNullString Then
        FM.txtS2F18.Text = FM.txtS1F08.Text
    End If
Else
    If GlAdS = True Then
        If FM.txtS2F18.Text <> FM.txtS1F08.Text Then
            AdRGe = True
        End If
    End If
End If

If FM.txtS2F19.Text = vbNullString Then
    If FM.txtS1F09.Text <> vbNullString Then
        FM.txtS2F19.Text = FM.txtS1F09.Text
    End If
Else
    If GlAdS = True Then
        If FM.txtS2F19.Text <> FM.txtS1F09.Text Then
            AdRGe = True
        End If
    End If
End If

If AdRGe = True Then
    SPopu Tit1, Mld1, IC48_Warning
End If

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AKopi " & Err.Number
Resume Next

End Sub
Public Sub AMain(ByVal PatNr As Long)
On Error GoTo LaErr

Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmAdress") = True Then
    Set FM = frmAdress
    frmAdress.ZOrder 0
    Exit Sub
End If

GlAdL = True 'Formular wird geladen
GlAId = PatNr

Screen.MousePointer = vbHourglass
DoEvents

AReg

Load frmAdress

Set FM = frmAdress

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    If GlBiA = False Then 'Bildschirmaktualisierung
        clFen.FenDsk 1
    Else
        clFen.FenDsk 2
    End If
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (880 / 2)
        .FeObn = (GlyGr / 2) - (720 / 2)
        .FeBre = 1000
        .FeHoh = 720
    Else
        .FeLin = IniGetVal("AdrForm", "FenLin")
        .FeObn = IniGetVal("AdrForm", "FenObe")
        .FeBre = IniGetVal("AdrForm", "FenBre")
        .FeHoh = IniGetVal("AdrForm", "FenHoh")
    End If
End With

AFont FM
AInit
AMenu
AKaSt
AOpen
DoEvents
Adr_EiSt
ASpLa
ASpLu
DoEvents
AMeAc False
DoEvents
Adr_Spl

If GlAId = -1 Then
    ASper False
ElseIf GlAId = 0 Then
    ANeue True
ElseIf GlAId > 0 Then
    ANeue
    Adr_Lad
    Kon_Lis
End If

With clFen
    .FenMov
    DoEvents
    AdPos
    Set CmBrs = FM.comBar01
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    .FenDsk 3
End With

Screen.MousePointer = vbNormal

Set clFen = Nothing

frmAdress.Show
DoEvents
GlAdL = False

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AMain " & Err.Number
Resume Next

End Sub
Public Sub AMeAc(ByVal EnAbl As Boolean)
On Error GoTo LaErr
'Schaltet das Menü ein / aus

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

CmAcs(ME_Adresse_Hinzufuegen).Enabled = EnAbl
CmAcs(ME_Adresse_Bearbeiten).Enabled = EnAbl
CmAcs(SY_AD_Adresse_Hinzufueg).Enabled = EnAbl
CmAcs(SY_AD_Adresse_Bearbeiten).Enabled = EnAbl
CmAcs(SY_AN_AnaBog_AdrSuch).Enabled = EnAbl
CmAcs(SY_AN_AnaBog_AdrBear).Enabled = EnAbl
CmAcs(SY_KB_KraBla_AdrSuch).Enabled = EnAbl
CmAcs(SY_KB_KraBla_AdrBear).Enabled = EnAbl
CmAcs(SY_AB_Abrech_AdrSuch).Enabled = EnAbl
CmAcs(SY_AB_Abrech_AdrBear).Enabled = EnAbl
CmAcs(SY_RZ_Rezept_AdrSuch).Enabled = EnAbl
CmAcs(SY_RZ_Rezept_AdrBear).Enabled = EnAbl
CmAcs(SY_RZ_Beleg_AdrSuch).Enabled = EnAbl
CmAcs(SY_RZ_Beleg_AdrBear).Enabled = EnAbl
CmAcs(SY_BI_Bild_AdrSuch).Enabled = EnAbl
CmAcs(SY_BI_Bild_AdrBear).Enabled = EnAbl
CmAcs(SY_LB_Labor_AdrSuch).Enabled = EnAbl
CmAcs(SY_LB_Labor_AdrBear).Enabled = EnAbl
CmAcs(SY_LA_Auftrag_AdrSuch).Enabled = EnAbl
CmAcs(SY_LA_Auftrag_AdrBear).Enabled = EnAbl
CmAcs(Tex_PaSuch).Enabled = EnAbl

Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AMeAc " & Err.Number
Resume Next

End Sub
Private Sub AMenu()
On Error GoTo CrErr
'Menue erstellen

Dim KeyNa As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbTem As XtremeCommandBars.RibbonTab
Dim MsBar As XtremeCommandBars.MessageBar
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmAdress
Set CmBrs = FM.comBar01
Set ImMan = frmMain.imgManag
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set MsBar = CmBrs.MessageBar
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

KeyNa = "ToolTips"

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(AM_Patient_Speichern, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Gruppe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Drucken, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Copy, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Del, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Clip1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Clip2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Notiz_Neu, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Notiz_Bearbeit, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Notiz_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Extras_Vorlage, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patient_Copy, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patient_Del, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patienten_Save, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patienten_Gruppe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Einzelbrief_Word, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_Drucken, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_SMS, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_GDT_Ex, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_GDT_Im, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Member_Add, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Member_Del, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Member_Copy, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Member_Save, vbNullString, vbNullString, vbNullString, vbNullString)
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Width = 100
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Width = 160
    CmPan.Text = "Alter:"
    Set CmPan = .AddPane(3)
    CmPan.Width = 160
    CmPan.Text = "Ersteintrag:"
    Set CmPan = .AddPane(4)
    CmPan.Width = 160
    CmPan.Text = "Aktualisierung:"
    Set CmPan = .AddPane(59137)
    Set CmPan = .AddPane(59138)
    Set CmPan = .AddPane(59139)
    .Visible = True
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Geburtsdatum, "Geburtsdatum")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Present
    .ToolTipText = "Löscht das Geburtsdatum"
    .Style = xtpButtonIconAndCaption
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Guthaben, "Guthaben")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Garbage
    .ToolTipText = "Löscht das aktuelle Guthaben"
    .Style = xtpButtonIconAndCaption
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Hilfe, "Hilfe")
With CmBuT
    .ToolTipText = "Öffnet die Kurzhilfe"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Beenden, "Schließen")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .ToolTipText = "Schließt den Dialog"
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

'------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Adr_Haupt, "Adressdaten")
With RbTab
    .id = RibTab_Adr_Haupt
    .ToolTip = "Zeigt die Hauptdaten des Patienten"
    .Visible = True
    .Selected = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Adresse Hinzufügen")
With CmCon
    .IconId = IC32_Patient_Add
    .ShortcutText = "F3"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Adresse Kopieren")
With CmCon
    .IconId = IC32_Patient_Copy
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Adresse Entfernen")
With CmCon
    .IconId = IC32_Patient_Del
    .ShortcutText = "Entf"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Adresse Speichern")
With CmCon
    .IconId = IC32_Disk_Patient
    .ShortcutText = "F8"
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Speichern_Close, "Adresse Schließen")
With CmCon
    .IconId = IC32_Disk_Patient
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Patienten_Suchen, "Adresse Suchen")
With CmCon
    .IconId = IC32_Patient_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Adressen_Chipkarte, "Chipkarte Einlesen")
With CmCon
    .IconId = IC32_Smartcard
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Gruppe, "Gruppe Zuordnen")
With CmCon
    .IconId = IC32_Folder_Check
    .ShortcutText = "F6"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("RibGroup", RibGrp_Adr_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Adressen_SMS, "Nachricht Senden")
With CmCon
    .IconId = IC32_Brief_Patient
    .ShortcutText = "F7"
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Ex, "GDT Datenexport")
    CmCon.IconId = IC16_IDCard_Export
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Im, "GDT Datenimport")
    CmCon.IconId = IC16_IDCard_Import
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_SMS, "Neue Nachricht erstellen")
    CmCon.IconId = IC16_Phone_Mobil
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Passw, "Verschlüsselungskennwort")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Warten, "Zur Warteliste hinzufügen")
    CmCon.IconId = IC16_Clipboard_Add
    CmCon.BeginGroup = True
End With

Set RbGrp = RbGps.AddGroup("Notiz", RibGrp_Adr_Termin)
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Neu, "Notiz Hinzufügen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Add
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Bearbeit, "Notiz Bearbeiten")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Edit
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Loeschen, "Notiz Entfernen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Del
End With

'------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Adr_Dokum, "Dokumentation")
With RbTab
    .id = RibTab_Adr_Dokum
    .Visible = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Adresse Hinzufügen")
With CmCon
    .IconId = IC32_Patient_Add
    .ShortcutText = "F3"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Adresse Kopieren")
With CmCon
    .IconId = IC32_Patient_Copy
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Adresse Entfernen")
With CmCon
    .IconId = IC32_Patient_Del
    .ShortcutText = "Entf"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Adresse Speichern")
With CmCon
    .IconId = IC32_Disk_Patient
    .ShortcutText = "F8"
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Speichern_Close, "Adresse Schließen")
With CmCon
    .IconId = IC32_Disk_Patient
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Adresse Suchen")
With CmCon
    .IconId = IC32_Patient_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Adressen_Chipkarte, "Chipkarte Einlesen")
With CmCon
    .IconId = IC32_Smartcard
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Gruppe, "Gruppe Zuordnen")
With CmCon
    .IconId = IC32_Folder_Check
    .ShortcutText = "F6"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("RibGroup", RibGrp_Adr_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Adressen_SMS, "Nachricht Senden")
With CmCon
    .IconId = IC32_Brief_Patient
    .ShortcutText = "F7"
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Ex, "GDT Datenexport")
    CmCon.IconId = IC16_IDCard_Export
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Im, "GDT Datenimport")
    CmCon.IconId = IC16_IDCard_Import
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_SMS, "Neue Nachricht erstellen")
    CmCon.IconId = IC16_Phone_Mobil
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Passw, "Verschlüsselungskennwort")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Warten, "Zur Warteliste hinzufügen")
    CmCon.IconId = IC16_Clipboard_Add
    CmCon.BeginGroup = True
End With

Set RbGrp = RbGps.AddGroup("Notiz", RibGrp_Adr_Termin)
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Neu, "Notiz Hinzufügen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Add
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Bearbeit, "Notiz Bearbeiten")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Edit
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Loeschen, "Notiz Entfernen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Del
End With

'------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Adr_Eigen, "Eigene Daten")
With RbTab
    .id = RibTab_Adr_Eigen
    .ToolTip = "Ermöglich es, eigene Datenfelder zu bearbeiten"
    .Visible = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Adresse Hinzufügen")
With CmCon
    .IconId = IC32_Patient_Add
    .ShortcutText = "F3"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Adresse Kopieren")
With CmCon
    .IconId = IC32_Patient_Copy
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Adresse Entfernen")
With CmCon
    .IconId = IC32_Patient_Del
    .ShortcutText = "Entf"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Adresse Speichern")
With CmCon
    .IconId = IC32_Disk_Patient
    .ShortcutText = "F8"
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Speichern_Close, "Adresse Schließen")
With CmCon
    .IconId = IC32_Disk_Patient
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Adresse Suchen")
With CmCon
    .IconId = IC32_Patient_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Adressen_Chipkarte, "Chipkarte Einlesen")
With CmCon
    .IconId = IC32_Smartcard
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Gruppe, "Gruppe Zuordnen")
With CmCon
    .IconId = IC32_Folder_Check
    .ShortcutText = "F6"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("RibGroup", RibGrp_Adr_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Adressen_SMS, "Nachricht Senden")
With CmCon
    .IconId = IC32_Brief_Patient
    .ShortcutText = "F7"
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Ex, "GDT Datenexport")
    CmCon.IconId = IC16_IDCard_Export
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Im, "GDT Datenimport")
    CmCon.IconId = IC16_IDCard_Import
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_SMS, "Neue Nachricht erstellen")
    CmCon.IconId = IC16_Phone_Mobil
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Passw, "Verschlüsselungskennwort")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Warten, "Zur Warteliste hinzufügen")
    CmCon.IconId = IC16_Clipboard_Add
    CmCon.BeginGroup = True
End With

Set RbGrp = RbGps.AddGroup("Notiz", RibGrp_Adr_Termin)
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Neu, "Notiz Hinzufügen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Add
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Bearbeit, "Notiz Bearbeiten")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Edit
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Loeschen, "Notiz Entfernen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Del
End With

'------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Adr_Booki, "Buchungen")
With RbTab
    .id = RibTab_Adr_Booki
    .ToolTip = "Zeigt alle Buchungsvorgänge des Patienten"
    .Visible = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Adresse Hinzufügen")
With CmCon
    .IconId = IC32_Patient_Add
    .ShortcutText = "F3"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Adresse Kopieren")
With CmCon
    .IconId = IC32_Patient_Copy
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Adresse Entfernen")
With CmCon
    .IconId = IC32_Patient_Del
    .ShortcutText = "Entf"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Adresse Speichern")
With CmCon
    .IconId = IC32_Disk_Patient
    .ShortcutText = "F8"
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Speichern_Close, "Adresse Schließen")
With CmCon
    .IconId = IC32_Disk_Patient
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Adresse Suchen")
With CmCon
    .IconId = IC32_Patient_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Adressen_Chipkarte, "Chipkarte Einlesen")
With CmCon
    .IconId = IC32_Smartcard
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Gruppe, "Gruppe Zuordnen")
With CmCon
    .IconId = IC32_Folder_Check
    .ShortcutText = "F6"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("RibGroup", RibGrp_Adr_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Adressen_SMS, "Nachricht Senden")
With CmCon
    .IconId = IC32_Brief_Patient
    .ShortcutText = "F7"
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Ex, "GDT Datenexport")
    CmCon.IconId = IC16_IDCard_Export
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Im, "GDT Datenimport")
    CmCon.IconId = IC16_IDCard_Import
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_SMS, "Neue Nachricht erstellen")
    CmCon.IconId = IC16_Phone_Mobil
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Passw, "Verschlüsselungskennwort")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Warten, "Zur Warteliste hinzufügen")
    CmCon.IconId = IC16_Clipboard_Add
    CmCon.BeginGroup = True
End With

Set RbGrp = RbGps.AddGroup("Notiz", RibGrp_Adr_Termin)
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Neu, "Notiz Hinzufügen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Add
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Bearbeit, "Notiz Bearbeiten")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Edit
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Loeschen, "Notiz Entfernen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Del
End With

'------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Adr_Membe, "Zugehörige")
With RbTab
    .id = RibTab_Adr_Membe
    .ToolTip = "Verwaltet die zu diesem Patienten zugehörigen Adressen"
    .Visible = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Member_Add, "Zugehörige Hinzufügen")
With CmCon
    .IconId = IC32_IDCard_Add
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Member_Copy, "Zugehörige Kopieren")
With CmCon
    .IconId = IC32_IDCard_Copy
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Member_Copy, "Zugehörigen Kopieren")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Member_Orig, "Hauptadresse Kopieren")
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Member_Del, "Zugehörige Entfernen")
With CmCon
    .IconId = IC32_IDCard_Del
    .Width = GlRib
End With

Set CmCon = RbGrp.Add(xtpControlButton, AD_Member_Save, "Zugehörige Speichern")
With CmCon
    .IconId = IC32_IDCard_Save
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Adresse Suchen")
With CmCon
    .IconId = IC32_Patient_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Adressen_Chipkarte, "Chipkarte Einlesen")
With CmCon
    .IconId = IC32_Smartcard
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Gruppe, "Gruppe Zuordnen")
With CmCon
    .IconId = IC32_Folder_Check
    .ShortcutText = "F6"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("RibGroup", RibGrp_Adr_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Adressen_SMS, "Nachricht Senden")
With CmCon
    .IconId = IC32_Brief_Patient
    .ShortcutText = "F7"
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Ex, "GDT Datenexport")
    CmCon.IconId = IC16_IDCard_Export
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_GDT_Im, "GDT Datenimport")
    CmCon.IconId = IC16_IDCard_Import
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_SMS, "Neue Nachricht erstellen")
    CmCon.IconId = IC16_Phone_Mobil
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Passw, "Verschlüsselungskennwort")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Adressen_Warten, "Zur Warteliste hinzufügen")
    CmCon.IconId = IC16_Clipboard_Add
    CmCon.BeginGroup = True
End With

Set RbGrp = RbGps.AddGroup("Notiz", RibGrp_Adr_Termin)
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Neu, "Notiz Hinzufügen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Add
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Bearbeit, "Notiz Bearbeiten")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Edit
End With
Set CmCon = RbGrp.Add(xtpControlButton, AM_Notiz_Loeschen, "Notiz Entfernen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Doc_Del
End With

'------------------------------------------------------------

Set CmCoS = RbBar.Controls
For Each CmCon In CmCoS
    CmCon.ToolTipText = IniGetOpt(KeyNa, CmCon.id)
Next CmCon

'------------------------------------------------------------

With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
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
    .SetIconSize True, 32, 32
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
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs
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
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
    .KeyBindings.Add FCONTROL, Asc("N"), AM_Notiz_Neu
    .KeyBindings.Add FCONTROL, Asc("Z"), AM_Patient_Clip1
    .KeyBindings.Add FCONTROL, Asc("R"), AM_Patient_Clip2
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
    .FontHeight = GlToF
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
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
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

Set CmPan = Nothing
Set CmSta = Nothing
Set CmPop = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set MsBar = Nothing
Set RbBar = Nothing
Set RbTab = Nothing
Set RbGrp = Nothing
Set RbGps = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AMenu " & Err.Number
Resume Next

End Sub

Public Sub ANeue(Optional ByVal AdNeu As Boolean = False)
On Error GoTo NeErr
'Bereitet die Neueingabe einer Adresse vor

Dim RetWe As Long
Dim IdxNr As Long
Dim AdPIN As Long
Dim TagWe As String
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim PrIts As XtremePropertyGrid.PropertyGridItems
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrBol As XtremePropertyGrid.PropertyGridItemBool
Dim PrDat As XtremePropertyGrid.PropertyGridItemDate
Dim PuGut As XtremeSuiteControls.PushButton
Dim CmTyp As XtremeSuiteControls.ComboBox
Dim CmVrs As XtremeSuiteControls.ComboBox

Set FM = frmAdress
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set TxGut = FM.txtS1F33
Set FeGes = FM.cmbS1F08
Set TxBri = FM.txtS1F20
Set TxFir = FM.txtS1F01
Set TxNum = FM.txtS1F30
Set TxKop = FM.txtS1F14
Set FeTi1 = FM.cmbS1F21
Set FeTi2 = FM.cmbS1F22
Set FeTi3 = FM.cmbS1F23
Set FeTi4 = FM.cmbS1F24
Set FeTi5 = FM.cmbS1F25
Set FeTi6 = FM.cmbS1F26
Set FeTi7 = FM.cmbS1F27
Set FeKat = FM.cmbS1F06
Set FeTar = FM.cmbS1F07
Set FePat = FM.cmbS2F10
Set CmTyp = FM.cmbS2F36
Set CmVrs = FM.cmbS1F28
Set FeFam = FM.txtS2F26
Set FeZah = FM.cmbS2F09
Set FeWar = FM.cmbS2F07
Set CmArt = FM.cmbS2F30
Set TxOrt = FM.txtS1F09
Set TxZGe = FM.txtS4F18
Set TxReG = FM.txtS2F25
Set TxErs = FM.txtS2F27
Set PuGut = FM.btnGutha
Set PrGr1 = FM.prpGrid1
Set PrGr2 = FM.prpGrid2
Set PrGr3 = FM.prpGrid3
Set CmBrs = FM.comBar01
Set CmSta = CmBrs.StatusBar

Tit1 = "Neue Adresse"
Mld1 = "Der Datensatz wurde noch nicht gespeichert. Möchten Sie wirklich eine neue Adresse anlegen?"

If AdNeu = True Then
    If GlAdS = True Then
        Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
        If Frage <> 6 Then
            Exit Sub
        End If
    End If
End If

For Each AktCo In FM.Controls
    If AktCo.Tag <> vbNullString Then
        Select Case TypeName(AktCo)
        Case "FlatEdit":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                AktCo.Text = vbNullString
                AktCo.Tag = 0 & TagWe
        Case "TextBox":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                AktCo.Text = vbNullString
                AktCo.Tag = 0 & TagWe
        Case "CheckBox":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                AktCo.Value = 0
                AktCo.Tag = 0 & TagWe
        Case "ComboBox":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                Select Case Left$(AktCo.Name, 3)
                Case "txt": AktCo.Text = vbNullString
                            AktCo.Tag = 0 & TagWe
                Case "cmb": AktCo.Tag = 1 & TagWe
                End Select
        End Select
    End If
Next AktCo

TxErs.Text = Date

Set PrIts = PrGr1.Categories

For Each PrKat In PrIts
    For Each PrItm In PrKat.Childs
        Select Case PrItm.Type
        Case PropertyItemString:
                PrItm.Value = vbNullString
        Case PropertyItemNumber:
                PrItm.Value = vbNullString
        Case PropertyItemBool:
                Set PrBol = PrItm
                PrBol.Value = False
        Case PropertyItemColor:
                PrItm.Value = RGB(255, 255, 255)
        Case PropertyItemDate:
                PrItm.Value = Date
        End Select
    Next PrItm
Next PrKat

Set PrIts = PrGr2.Categories

For Each PrKat In PrIts
    For Each PrItm In PrKat.Childs
        Select Case PrItm.Type
        Case PropertyItemString:
                PrItm.Value = vbNullString
        Case PropertyItemNumber:
                PrItm.Value = vbNullString
        Case PropertyItemBool:
                Set PrBol = PrItm
                PrBol.Value = False
        Case PropertyItemColor:
                PrItm.Value = RGB(255, 255, 255)
        Case PropertyItemDate:
                PrItm.Value = Date
        End Select
    Next PrItm
Next PrKat

Set PrIts = PrGr3.Categories

For Each PrKat In PrIts
    For Each PrItm In PrKat.Childs
        Select Case PrItm.Type
        Case PropertyItemString:
                PrItm.Value = vbNullString
        Case PropertyItemNumber:
                PrItm.Value = vbNullString
        Case PropertyItemBool:
                Set PrBol = PrItm
                PrBol.Value = False
        Case PropertyItemColor:
                PrItm.Value = RGB(255, 255, 255)
        Case PropertyItemDate:
                PrItm.Value = Date
        End Select
    Next PrItm
Next PrKat

If AdNeu = True Then
    AdPIN = Adr_Let()
    GlAdG = CreateID("A")
    GlAdN = True
    GlAdS = False
    AEnab False
    PrGr3.Enabled = False
    PuGut.Enabled = False
    TxNum.Text = Format$(AdPIN, "000000")
    TxNum.Tag = "1Mandant"
    TxBri.Text = GlBrf 'Standrad-Briefanrede
    TxKop.Text = GlKop
    TxGut.Text = GlWa2
    FeGes.ListIndex = 0
    FeTi1.ListIndex = 0
    FeTi2.ListIndex = 1
    FeTi3.ListIndex = 3
    FeTi4.ListIndex = 4
    FeTi5.ListIndex = 8
    FeTi6.ListIndex = 9
    FeTi7.ListIndex = 10
    FeWar.ListIndex = GlStW - 1
    FeFam.ListIndex = 5
    CmVrs.ListIndex = GlRVs - 1 'Standard-Rechnungsversandweg
    CmTyp.ListIndex = GlStP
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        CmArt.ListIndex = 4
    Else
        CmArt.ListIndex = 0
    End If
    If FePat.ListCount > 0 Then
        If GlSuP.SuMan > 0 Then
            IdxNr = Adr_Cm(FePat, GlSuP.SuMan)
            FePat.ListIndex = IdxNr
        Else
            FePat.ListIndex = GlMaA(GlSMa, 0) - 1
        End If
    End If
    FeTar.ListIndex = 0
    If GlStK - 1 <= FeKat.ListCount Then 'Standardgebührenkatalog
        FeKat.ListIndex = GlStK - 1
    Else
        FeKat.ListIndex = 0
    End If
    If GlStZ - 1 <= FeZah.ListCount Then
        FeZah.ListIndex = GlStZ - 1
    Else
        FeZah.ListIndex = 0
    End If
    CmSta.Pane(0).Text = vbNullString
    CmSta.Pane(2).Text = "Ersteintrag: " & Date
    CmSta.Pane(3).Text = "Aktualisierung: " & Date
    If FM.Visible = True Then
        If Rahm1.Visible = True Then
            TxFir.SetFocus
        End If
    End If
    GlAId = -1
End If

Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ANeue " & Err.Number
Resume Next

End Sub
Private Sub AOpen()
On Error GoTo FiErr
'Füllt die Comboboxen des Adressformulars mit Inhalt aus persistenten Recordsets

Dim RetWe As Long
Dim AktZa As Integer

Set FM = frmAdress
Set FeTi1 = FM.cmbS1F21
Set FeTi2 = FM.cmbS1F22
Set FeTi3 = FM.cmbS1F23
Set FeTi4 = FM.cmbS1F24
Set FeKat = FM.cmbS1F06
Set FeTar = FM.cmbS1F07
Set FePat = FM.cmbS2F10
Set FeZah = FM.cmbS2F09
Set FeWar = FM.cmbS2F07
Set FeBeG = FM.cmbS2F29
Set FeLa1 = FM.txtS1F12
Set FeLa3 = FM.cmbS4F12
Set CmVdo = FM.cmbS2F31
Set CmArt = FM.cmbS2F30

For AktZa = 1 To UBound(GlGKa) 'Gebührenkataloge
    FeKat.AddItem GlGKa(AktZa, 1)
    FeKat.ItemData(AktZa - 1) = GlGKa(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlTar) 'Versicherungstarife
    FeTar.AddItem GlTar(AktZa, 2)
    FeTar.ItemData(AktZa - 1) = GlTar(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlZah)
    FeZah.AddItem GlZah(AktZa, 1)
    FeZah.ItemData(AktZa - 1) = GlZah(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlWar)
    FeWar.AddItem GlWar(AktZa, 1)
    FeWar.ItemData(AktZa - 1) = GlWar(AktZa, 0)
Next AktZa

If GlMaV = True Then 'Mandanten vorhanden
    For AktZa = 1 To UBound(GlMaA)
        FePat.AddItem GlMaA(AktZa, 1)
        FePat.ItemData(AktZa - 1) = GlMaA(AktZa, 2)
    Next AktZa
End If

If GlArV = True Then
   For AktZa = 1 To UBound(GlArz)
        CmVdo.AddItem GlArz(AktZa, 8)
        CmVdo.ItemData(AktZa - 1) = GlArz(AktZa, 0)
    Next AktZa
End If

If GlBgV = True Then
   For AktZa = 1 To UBound(GlBeG)
        FeBeG.AddItem GlBeG(AktZa, 1)
        FeBeG.ItemData(AktZa - 1) = GlBeG(AktZa, 0)
    Next AktZa
End If

For AktZa = 1 To UBound(GlLan)
    With FeLa1
        .AddItem GlLan(AktZa, 1)
        .ItemData(AktZa - 1) = GlLan(AktZa, 0)
    End With
    With FeLa3
        .AddItem GlLan(AktZa, 1)
        .ItemData(AktZa - 1) = GlLan(AktZa, 0)
    End With
Next AktZa

If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
    With CmArt
        For AktZa = 0 To UBound(GlVeG) - 1
            .AddItem GlVeG(AktZa, 0)
            .ItemData(AktZa) = GlVeG(AktZa, 1)
        Next AktZa
        .DropDownItemCount = 6
    End With
Else
    With CmArt
        For AktZa = 0 To UBound(GlVeA) - 1
            .AddItem GlVeA(AktZa, 0)
            .ItemData(AktZa) = GlVeA(AktZa, 1)
        Next AktZa
        .DropDownItemCount = 22
    End With
End If

RetWe = SendMessage(FeTi1.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(FeTi2.hwnd, CB_SETCURSEL, 1, ByVal 0&)
RetWe = SendMessage(FeTi3.hwnd, CB_SETCURSEL, 3, ByVal 0&)
RetWe = SendMessage(FeTi4.hwnd, CB_SETCURSEL, 4, ByVal 0&)
RetWe = SendMessage(FeKat.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(FeZah.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(FeWar.hwnd, CB_SETCURSEL, 0, ByVal 0&)

Exit Sub

FiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AOpen " & Err.Number
Resume Next

End Sub
Public Sub AdPos()
On Error GoTo ReErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim ClHal As Long
Dim ClDri As Long
Dim ClHoD As Long
Dim TbHoh As Long
Dim TbObe As Long
Dim RaBre As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim S1F01 As XtremeSuiteControls.FlatEdit
Dim S1F03 As XtremeSuiteControls.FlatEdit
Dim S1F04 As XtremeSuiteControls.FlatEdit
Dim S1F05 As XtremeSuiteControls.FlatEdit
Dim S1F06 As XtremeSuiteControls.FlatEdit
Dim S1F09 As XtremeSuiteControls.FlatEdit
Dim S1F11 As XtremeSuiteControls.FlatEdit
Dim S1F14 As XtremeSuiteControls.FlatEdit
Dim S1F15 As XtremeSuiteControls.FlatEdit
Dim S1F16 As XtremeSuiteControls.FlatEdit
Dim S1F17 As XtremeSuiteControls.FlatEdit
Dim S1F18 As XtremeSuiteControls.FlatEdit
Dim S1F19 As XtremeSuiteControls.FlatEdit
Dim S1F27 As XtremeSuiteControls.FlatEdit
Dim S1F30 As XtremeSuiteControls.FlatEdit
Dim S1F32 As XtremeSuiteControls.FlatEdit
Dim S1F37 As XtremeSuiteControls.FlatEdit
Dim S2F11 As XtremeSuiteControls.FlatEdit
Dim S2F13 As XtremeSuiteControls.FlatEdit
Dim S2F14 As XtremeSuiteControls.FlatEdit
Dim S2F15 As XtremeSuiteControls.FlatEdit
Dim S2F16 As XtremeSuiteControls.FlatEdit
Dim S2F18 As XtremeSuiteControls.FlatEdit
Dim S2F19 As XtremeSuiteControls.FlatEdit
Dim S2F20 As XtremeSuiteControls.FlatEdit
Dim S2F03 As XtremeSuiteControls.FlatEdit
Dim S2F05 As XtremeSuiteControls.FlatEdit
Dim S2F25 As XtremeSuiteControls.FlatEdit
Dim S2F32 As XtremeSuiteControls.FlatEdit
Dim S2F33 As XtremeSuiteControls.FlatEdit
Dim S2F34 As XtremeSuiteControls.FlatEdit
Dim S2F35 As XtremeSuiteControls.FlatEdit
Dim S2F22 As XtremeSuiteControls.FlatEdit
Dim S4F01 As XtremeSuiteControls.FlatEdit
Dim S4F03 As XtremeSuiteControls.FlatEdit
Dim S4F04 As XtremeSuiteControls.FlatEdit
Dim S4F05 As XtremeSuiteControls.FlatEdit
Dim S4F06 As XtremeSuiteControls.FlatEdit
Dim S4F09 As XtremeSuiteControls.FlatEdit
Dim S4F15 As XtremeSuiteControls.FlatEdit
Dim S4F16 As XtremeSuiteControls.FlatEdit
Dim S4F17 As XtremeSuiteControls.FlatEdit
Dim S4F19 As XtremeSuiteControls.FlatEdit
Dim S1F22 As XtremeSuiteControls.ComboBox
Dim S1F10 As XtremeSuiteControls.ComboBox
Dim S1F12 As XtremeSuiteControls.ComboBox
Dim S1F21 As XtremeSuiteControls.ComboBox
Dim S1F23 As XtremeSuiteControls.ComboBox
Dim S1F24 As XtremeSuiteControls.ComboBox
Dim S1F07 As XtremeSuiteControls.ComboBox
Dim S2F12 As XtremeSuiteControls.ComboBox
Dim S2F06 As XtremeSuiteControls.ComboBox
Dim S2F07 As XtremeSuiteControls.ComboBox
Dim S2F08 As XtremeSuiteControls.ComboBox
Dim S2F09 As XtremeSuiteControls.ComboBox
Dim S2F10 As XtremeSuiteControls.ComboBox
Dim S2F24 As XtremeSuiteControls.ComboBox
Dim S2F29 As XtremeSuiteControls.ComboBox
Dim S2F30 As XtremeSuiteControls.ComboBox
Dim S2F31 As XtremeSuiteControls.ComboBox
Dim S4F11 As XtremeSuiteControls.ComboBox
Dim S4F12 As XtremeSuiteControls.ComboBox
Dim S1F28 As XtremeSuiteControls.ComboBox
Dim S2F36 As XtremeSuiteControls.ComboBox

Set FM = frmAdress
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set CmBrs = FM.comBar01
Set S3F02 = FM.txtS3F02
Set S3F03 = FM.txtS3F03
Set S1F01 = FM.txtS1F01
Set S1F03 = FM.txtS1F03
Set S1F04 = FM.txtS1F04
Set S1F05 = FM.txtS1F05
Set S1F06 = FM.txtS1F06
Set S1F09 = FM.txtS1F09
Set S1F10 = FM.cmbS1F10
Set S1F11 = FM.txtS1F11
Set S1F12 = FM.txtS1F12
Set S1F14 = FM.txtS1F14
Set S1F15 = FM.txtS1F15
Set S1F16 = FM.txtS1F16
Set S1F17 = FM.txtS1F17
Set S1F18 = FM.txtS1F18
Set S1F19 = FM.txtS1F19
Set S2F08 = FM.txtS2F08
Set S1F27 = FM.txtS1F27
Set S1F37 = FM.txtS1F37
Set S1F30 = FM.txtS1F30
Set S1F32 = FM.txtS1F32
Set S2F11 = FM.txtS2F11
Set S2F12 = FM.txtS2F12
Set S2F13 = FM.txtS2F13
Set S2F14 = FM.txtS2F14
Set S2F15 = FM.txtS2F15
Set S2F16 = FM.txtS2F16
Set S2F18 = FM.txtS2F18
Set S2F19 = FM.txtS2F19
Set S2F20 = FM.txtS2F20
Set S2F22 = FM.txtS2F22
Set S2F03 = FM.txtS2F03
Set S2F05 = FM.txtS2F05
Set S2F25 = FM.txtS2F24
Set S2F24 = FM.txtS2F26
Set S2F33 = FM.txtS2F33
Set S2F34 = FM.txtS2F34
Set S2F35 = FM.txtS2F35
Set S4F01 = FM.txtS4F01
Set S4F03 = FM.txtS4F03
Set S4F04 = FM.txtS4F04
Set S4F05 = FM.txtS4F05
Set S4F06 = FM.txtS4F06
Set S4F09 = FM.txtS4F09
Set S4F15 = FM.txtS4F15
Set S4F16 = FM.txtS4F16
Set S4F17 = FM.txtS4F17
Set S4F19 = FM.txtS4F19
Set S1F22 = FM.txtS1F22
Set S4F11 = FM.cmbS4F11
Set S4F12 = FM.cmbS4F12
Set S2F30 = FM.cmbS2F30
Set S1F07 = FM.cmbS1F07
Set S2F06 = FM.cmbS1F06
Set S2F07 = FM.cmbS2F07
Set S2F09 = FM.cmbS2F09
Set S2F10 = FM.cmbS2F10
Set S1F21 = FM.cmbS1F21
Set S1F23 = FM.cmbS1F23
Set S1F24 = FM.cmbS1F24
Set S2F29 = FM.cmbS2F29
Set S2F31 = FM.cmbS2F31
Set S2L34 = FM.lblS2L34
Set S2L35 = FM.lblS2L35
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set PrGr1 = FM.prpGrid1
Set PrGr2 = FM.prpGrid2
Set S1F28 = FM.cmbS1F28
Set S2F36 = FM.cmbS2F36
Set PrGr3 = FM.prpGrid3

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHal = ClBre / 2
    ClDri = ClBre / 3
    ClHoh = ClHoh - ClObn - 20
    TbObe = ClObn + S1F37.Top + S1F37.Height + 220
    TbHoh = (ClHoh + ClObn) - TbObe
    If TbHoh < 0 Then TbHoh = 0
    ClHoD = (ClHoh - TbHoh) / 3
    RaBre = ClDri - 300
    Rahm1.Move 400, ClObn, RaBre, 7200
    Rahm2.Move ClDri + 150, ClObn, RaBre, 7200
    Rahm3.Move (ClDri * 2) - 100, ClObn, RaBre, 7200
    Rahm4.Move 400, ClObn, RaBre, 7200
    Rahm5.Move 0, ClObn + 3020, ClBre, 4160
    PrGr1.Move 0, ClObn, ClHal, ClHoh
    PrGr2.Move ClHal - 20, ClObn, ClHal + 20, ClHoh
    PrGr3.Move 0, ClObn, ClBre, 3000
    RpCo1.Move 0, TbObe, ClBre, TbHoh
    RpCo2.Move 0, ClObn, ClBre, ClHoD
    RpCo3.Move 0, ClObn + ClHoD, ClBre, ClHoD
    RpCo4.Move 0, ClObn + ClHoD + ClHoD, ClBre, ClHoD
    RpCo5.Move ClDri + 150, ClObn, (ClDri * 2) - 150, 7200
    If ClBre < 9900 Then Exit Sub
    S2L35.Left = ClHal + 60
    S3F02.Move 10, 350, ClHal - 20, 3800
    S3F03.Move ClHal + 20, 350, ClHal - 40, 3800
    S1F01.Width = RaBre - 1500
    S1F03.Width = RaBre - 3370
    S1F04.Width = RaBre - 1500
    S1F05.Width = RaBre - 1500
    S1F06.Width = RaBre - 1500
    S1F07.Width = RaBre - 1500
    S1F09.Width = RaBre - 2580
    S1F12.Width = RaBre - 1500
    S1F10.Width = RaBre - 1500
    S1F22.Width = RaBre - 1480
    S1F11.Width = RaBre - 1500
    S2F11.Width = RaBre - 1500
    S2F13.Width = RaBre - 3370
    S2F14.Width = RaBre - 1500
    S2F15.Width = RaBre - 1500
    S2F16.Width = RaBre - 1500
    S2F19.Width = RaBre - 2580
    S2F22.Width = RaBre - 1500
    S2F20.Width = RaBre - 1500
    S2F25.Width = RaBre - 1480
    S2F06.Width = RaBre - 1480
    S2F08.Width = RaBre - 1480
    S2F09.Width = RaBre - 1480
    S2F24.Width = RaBre - 3080
    S1F15.Width = RaBre - 2710
    S1F16.Width = RaBre - 2710
    S1F17.Width = RaBre - 2710
    S1F18.Width = RaBre - 2710
    S1F19.Width = RaBre - 2710
    S2F34.Width = RaBre - 2710
    S1F27.Width = RaBre - 2710
    S2F03.Width = RaBre - 1460
    S2F05.Width = RaBre - 1460
    S2F10.Width = RaBre - 1460
    S1F30.Width = RaBre - 3370
    S1F32.Width = RaBre - 3370
    S2F30.Width = RaBre - 1500
    S2F31.Width = RaBre - 1500
    S2F07.Width = RaBre - 1460
    S2F33.Width = RaBre - 1460
    S2F35.Width = RaBre - 1460
    S2F29.Width = RaBre - 1460
    S4F01.Width = RaBre - 1500
    S4F03.Width = RaBre - 3370
    S4F04.Width = RaBre - 1500
    S4F05.Width = RaBre - 1500
    S4F06.Width = RaBre - 1500
    S4F09.Width = RaBre - 2580
    S4F11.Width = RaBre - 1500
    S4F12.Width = RaBre - 1500
    S4F15.Width = RaBre - 1910
    S4F16.Width = RaBre - 1910
    S4F17.Width = RaBre - 1500
    S4F19.Width = RaBre - 1500
    S1F28.Width = RaBre - 2840
    S2F36.Width = RaBre - 3050
End If

Set CmBrs = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AdPos " & Err.Number
Resume Next

End Sub

Public Sub AOutl()
On Error GoTo ReErr
'Übergibt die aktuelle Adresse an Outlook

Dim IdxNr As Long
Dim FIDKu As Variant
Dim FFirm As Variant
Dim FAnre As Variant
Dim FTite As Variant
Dim FName As Variant
Dim FVorn As Variant
Dim FStra As Variant
Dim FPOst As Variant
Dim FOrte As Variant
Dim FLand As Variant
Dim KFirm As Variant
Dim KAnre As Variant
Dim Krite As Variant
Dim KName As Variant
Dim KVorn As Variant
Dim KStra As Variant
Dim KPost As Variant
Dim KOrte As Variant
Dim KLand As Variant
Dim KBeru As Variant
Dim KGebo As Variant
Dim KBeme As Variant
Dim Tele1 As Variant
Dim Tele2 As Variant
Dim Tele3 As Variant
Dim Tele4 As Variant
Dim Tele5 As Variant
Dim Tele6 As Variant
Dim IntNt As Variant
Dim OutAd As Boolean
Dim OutGb As Boolean
Dim OutAb As Boolean
Dim OutRe As Boolean
Dim Mld1, Tit1 As String
Dim Frage As Integer

Dim NaSpa As Object
Dim MapFo As Object
Dim KoIts As Object
Dim KoItm As Object

Set FM = frmAdress
Set TxNum = FM.txtS1F30

FIDKu = FM.txtS1F11.Text
FFirm = FM.txtS2F11.Text
FAnre = FM.txtS2F12.Text
FTite = FM.txtS2F13.Text
FVorn = FM.txtS2F14.Text
FName = FM.txtS2F15.Text
FStra = FM.txtS2F16.Text
FPOst = FM.txtS2F18.Text
FOrte = FM.txtS2F19.Text
FLand = FM.txtS2F22.Text
KFirm = FM.txtS1F01.Text
KAnre = FM.txtS1F02.Text
Krite = FM.txtS1F03.Text
KName = FM.txtS1F05.Text
KVorn = FM.txtS1F04.Text
KStra = FM.txtS1F06.Text
KPost = FM.txtS1F08.Text
KOrte = FM.txtS1F09.Text
KLand = FM.txtS1F12.Text
KGebo = FM.txtS1F13.Text
KBeru = FM.txtS2F24.Text
KBeme = FM.txtS3F02.Text
Tele1 = FM.txtS1F15.Text
Tele2 = FM.txtS1F16.Text
Tele3 = FM.txtS1F17.Text
Tele4 = FM.txtS1F18.Text
Tele5 = FM.txtS1F19.Text
Tele6 = FM.txtS2F34.Text
IntNt = FM.txtS1F27.Text

Tit1 = "Outlookübergabe"
Mld1 = "Diese Adresse wurde bereits einmal an Outlook übergeben. Möchten Sie diese erneut an Outlook übergeben?"

OutAd = CBool(IniGetVal("System", "OutAdr"))
OutGb = CBool(IniGetVal("System", "OutGeb"))

If GlAId < 1 Then
    WindowMess "Sie müssen die aktuellen Daten erst speichern", Dial2, Tit1, FM.hwnd
    Exit Sub
Else
    S_AdDe GlAId
    With GlADt
        OutAb = .AdEdi
        OutRe = .AdRep
    End With
End If

If OutAd = True Then
    If OutAb = False Then
        WindowMess "Sie können nur Adressen exportieren, die als Outlookadresse gekennzeichnet sind", Dial2, Tit1, FM.hwnd
        Exit Sub
    End If
End If

If OutRe = True Then
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage <> 6 Then
        Exit Sub
    End If
End If

SOuOp 'Outlook Öffnen

Set NaSpa = OutOb.GetNamespace("MAPI") 'NameSpace

If GlMaF = True Then
    Set MapFo = NaSpa.GetDefaultFolder(olFolderContacts)
Else
    Set MapFo = NaSpa.PickFolder
End If
    
If TypeName(MapFo) = "Nothing" Then
    Exit Sub
ElseIf MapFo.DefaultItemType <> olContactItem Then
    Exit Sub
End If

Set KoIts = MapFo.Items

Set KoItm = KoIts.Add(olContactItem)
With KoItm
    If GlAId > 0 Then
        .CustomerID = GlAId
    End If
    If FIDKu <> vbNullString Then
        .FileAs = FIDKu
    End If

    If GlAno = False Then
        If Krite <> vbNullString Then
            If FAnre <> vbNullString Then
                .Title = KAnre & Chr$(32) & Krite
            Else
                .Title = Krite
            End If
        Else
            If FAnre <> vbNullString Then
                .Title = KAnre
            End If
        End If
    End If
    If KFirm <> vbNullString Then .CompanyName = KFirm
    If KVorn <> vbNullString Then .FirstName = KVorn
    If KName <> vbNullString Then .LastName = KName
    If KBeme <> vbNullString Then .Body = KBeme
    If FStra <> vbNullString Then .BusinessAddressStreet = FStra
    If FOrte <> vbNullString Then .BusinessAddressCity = FOrte
    If FPOst <> vbNullString Then .BusinessAddressPostalCode = FPOst
    If FLand <> vbNullString Then .BusinessAddressCountry = FLand
    If Tele1 <> vbNullString Then .PrimaryTelephoneNumber = STele(CStr(Tele1))
    If Tele1 <> vbNullString Then .HomeTelephoneNumber = STele(CStr(Tele1))
    If Tele2 <> vbNullString Then .BusinessTelephoneNumber = STele(CStr(Tele2))
    If Tele3 <> vbNullString Then .BusinessFaxNumber = STele(CStr(Tele3))
    If Tele4 <> vbNullString Then .MobileTelephoneNumber = STele(CStr(Tele4))
    If Tele5 <> vbNullString Then .Email1Address = Tele5
    If Tele6 <> vbNullString Then .Email2Address = Tele6
    If IntNt <> vbNullString Then .BusinessHomePage = IntNt
    If KStra <> vbNullString Then .HomeAddressStreet = KStra
    If KOrte <> vbNullString Then .HomeAddressCity = KOrte
    If KLand <> vbNullString Then .HomeAddressCountry = KLand
    If KPost <> vbNullString Then .HomeAddressPostalCode = KPost
    If KBeru <> vbNullString Then .Profession = KBeru
    If GlAdG <> vbNullString Then .BillingInformation = GlAdG
    If OutGb = True Then
        If KGebo <> vbNullString Then
            .Birthday = KGebo
        End If
    End If
    DoEvents
    .Save
End With

DoEvents
IdxNr = S_AdGui(GlAdG, "ID0")
DBCmEx1 "qrySimAdRe1", "@IdxNr", IdxNr

WindowMess "Die Adresse wurden erfolgreich an Outlook übergeben", Dial2, Tit1, FM.hwnd
        
Set KoItm = Nothing
Set NaSpa = Nothing
Set MapFo = Nothing
Set KoIts = Nothing
Set OutOb = Nothing

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AOutl " & Err.Number
Resume Next

End Sub
Public Sub APaSe(ByVal AdIdx As Long)
On Error GoTo SeErr
'Lädt den ausgewählten Patienten

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCon As XtremeCommandBars.CommandBarControl

Set FM = frmAdress
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

GlAId = AdIdx * (-1)
GlAdL = True
ASper True
ANeue
Adr_Lad
Kon_Lis
GlAdS = False
GlAdL = False

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "APaSe " & Err.Number
Resume Next

End Sub
Public Sub APaSu(ByVal AdIdx As Long, AdStr As String)
On Error GoTo InErr
'Addiert die gefundenen Patienten im Splitbutton

Dim IdVor As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmBuT As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set FM = frmAdress
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

AdIdx = AdIdx * (-1)

IdVor = False
Set CmCon = CmBrs.FindControl(CmCon, AD_Patienten_Suchen, , True)
Set CmCoS = CmCon.CommandBar.Controls
For Each CmBuT In CmCoS
    If CmBuT.id = AdIdx Then
        IdVor = True
        Exit For
    End If
Next CmBuT
If IdVor = False Then
    Set CmBuT = CmCoS.Add(xtpControlButton, AdIdx, AdStr)
    CmBuT.IconId = IC16_Bookmark
End If

Set CmBrs = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "APaSu " & Err.Number
Resume Next

End Sub
Private Sub AReg()
On Error GoTo ReErr
'Anlegen von Registryeinträgen

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

xGro = 1000
yGro = 760
xPos = (GlxGr / 2) - (xGro / 2)
yPos = (GlyGr / 2) - (yGro / 2)

If IniGetSek(GlINI, "AdrForm") = False Then IniSetSek "AdrForm"
If IniGetVal("AdrForm", "FenLin") = vbNullString Then IniSetVal "AdrForm", "FenLin", xPos
If IniGetVal("AdrForm", "FenObe") = vbNullString Then IniSetVal "AdrForm", "FenObe", yPos
If IniGetVal("AdrForm", "FenBre") = vbNullString Then IniSetVal "AdrForm", "FenBre", xGro
If IniGetVal("AdrForm", "FenHoh") = vbNullString Then IniSetVal "AdrForm", "FenHoh", yGro

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AReg " & Err.Number
Resume Next

End Sub
Public Sub ASper(ByVal Flag As Boolean, Optional ByVal BeDat As Boolean)
On Error GoTo NeErr
'Sperrt oder entsperrt die Adressfelder

If BeDat = True Then
    Set FM = frmMandant
Else
    Set FM = frmAdress
End If

For Each AktCo In FM.Controls
    Select Case Left$(AktCo.Name, 3)
    Case "txt": AktCo.Enabled = Flag
    Case "cmb": AktCo.Enabled = Flag
    Case "msk": AktCo.Enabled = Flag
    Case "upd": AktCo.Enabled = Flag
    End Select
Next AktCo

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ASper " & Err.Number
Resume Next

End Sub
Public Sub ASpLa()
On Error GoTo SpErr
'Formratieren der Spalten

Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmAdress
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns

With RpCls
    Set RpCol = .Add(Not_ID2, "", 0, False)
    Set RpCol = .Add(Not_VonDat, "Datum", 100, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    Set RpCol = .Add(Not_ZeiVon, "Von", 50, False)
    Set RpCol = .Add(Not_ZeiBis, "Bis", 50, False)
    Set RpCol = .Add(Not_IDKurz, "Anlass", 100, False)
    RpCol.AutoSize = True
    If RpCon.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Not_Datei, "", 30, False)
    RpCol.Alignment = xtpAlignmentIconLeft
    Set RpCol = .Add(Not_Behandler, "Bearbeitet", 120, False)
    Set RpCol = .Add(Not_Erledigt, "", 20, False)
End With

For Each RpCol In RpCls
    With RpCol
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
    End With
Next RpCol

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ASpLa " & Err.Number
Resume Next

End Sub
Public Sub ASpLu()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmAdress
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5

Set RpCls = RpCo2.Columns
With RpCls
    Set RpCol = .Add(Buh_ID0, "ID0", 0, False)
    Set RpCol = .Add(Buh_Datum, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Buh_Buchtext, "Buchungstext", 0, True)
    If RpCo2.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
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
If GlBuc = True Then 'einfache Buchhaltung verwenden
    RpCls(Buh_Sachkontenbez).Width = 180
    RpCls(Buh_Geldkontenbez).Width = 160
Else
    RpCls(Buh_Sachkontenbez).Width = 0
    RpCls(Buh_Geldkontenbez).Width = 0
End If
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

'---

Set RpCls = RpCo3.Columns
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
    If RpCo3.PaintManager.FixedRowHeight = False Then
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
    Set RpCol = .Add(OPo_Einzahlung, "Einzahlung", 0, True)
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
    Set RpCol = .Add(OPo_Monat, "Monat", 0, True)
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
    RpCls(OPo_IDP).Width = 180
    RpCls(OPo_Berichtdatum).Width = 110
    RpCls(OPo_Steuer).Width = 100
    RpCls(OPo_Monat).Width = 0
    RpCls(OPo_Mahnung1).Width = 110
    RpCls(OPo_Mahnung2).Width = 110
    RpCls(OPo_Mahnung3).Width = 110
    RpCls(OPo_Mahnung4).Width = 110
    RpCls(OPo_Mahnung5).Width = 110
    RpCls(OPo_IDT).Width = 180
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
    RpCls(OPo_IDP).Width = 180
    RpCls(OPo_Berichtdatum).Width = 80
    RpCls(OPo_Steuer).Width = 70
    RpCls(OPo_Monat).Width = 0
    RpCls(OPo_Mahnung1).Width = 80
    RpCls(OPo_Mahnung2).Width = 80
    RpCls(OPo_Mahnung3).Width = 80
    RpCls(OPo_Mahnung4).Width = 80
    RpCls(OPo_Mahnung5).Width = 80
    RpCls(OPo_IDT).Width = 180
    RpCls(OPo_Versand).Width = 20
End If

'---

Set RpCls = RpCo4.Columns
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
    Set RpCol = .Add(Ter_ZeiVor, "ZeiVor", 0, False)
    Set RpCol = .Add(Ter_Priorität, "Prio.", 0, True)
    RpCol.Alignment = xtpAlignmentCenter
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Ter_Vorwarn, "Vorwarn", 0, False)
    Set RpCol = .Add(Ter_Farbe, "Farbe", 0, False)
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

'---

Set RpCls = RpCo5.Columns
With RpCls
    Set RpCol = .Add(Adr_ID0, "IDA", 0, False)
    Set RpCol = .Add(Adr_ID3, "ID3", 0, False)
    Set RpCol = .Add(Adr_IDKurz, "Suchbegriff", 0, True)
    If RpCo2.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Adr_Geboren, "EMail", 0, True)
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
    Set RpCol = .Add(Adr_Telefon5, "Geboren", 0, True)
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
    Set RpCol = .Add(Adr_Mandant, "Nr.", 0, True)
    Set RpCol = .Add(Adr_VIP, "VIP", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Adr_Titel, "Titel", 0, False)
    Set RpCol = .Add(Adr_Land, "Land", 0, False)
    Set RpCol = .Add(Adr_Behindert, "Behindert", 0, False)
    Set RpCol = .Add(Adr_Passiv, "Passiv", 0, False)
    Set RpCol = .Add(Adr_Gruppen, "Gruppen", 0, True)
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

RpCls(Adr_ID0).Width = 0
RpCls(Adr_ID3).Width = 0
RpCls(Adr_IDKurz).Width = 220
RpCls(Adr_Geboren).Width = 160
RpCls(Adr_Name).Width = 100
RpCls(Adr_Vorname).Width = 100
RpCls(Adr_Straße).Width = 120
RpCls(Adr_PLZ).Width = 60
RpCls(Adr_Ort).Width = 100
RpCls(Adr_Firma1).Width = 0
RpCls(Adr_Telefon1).Width = 90
RpCls(Adr_Telefon2).Width = 0
RpCls(Adr_Telefon3).Width = 0
RpCls(Adr_Telefon4).Width = 0
RpCls(Adr_Telefon5).Width = 120
RpCls(Adr_Geschlecht).Width = 0
RpCls(Adr_Datum).Width = 0
RpCls(Adr_Briefanrede).Width = 0
RpCls(Adr_Anschrift).Width = 0
RpCls(Adr_TreKey).Width = 0
RpCls(Adr_Grafik).Width = 0
RpCls(Adr_GuiID).Width = 0
RpCls(Adr_Objekt).Width = 0
RpCls(Adr_IDP).Width = 0
RpCls(Adr_Mandant).Width = 0
RpCls(Adr_VIP).Width = 0
RpCls(Adr_Titel).Width = 0
RpCls(Adr_Land).Width = 0
RpCls(Adr_Behindert).Width = 0
RpCls(Adr_Passiv).Width = 0
RpCls(Adr_Gruppen).Width = 0
RpCls(Adr_Versand).Width = 0

'---

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo2 = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ASpLu " & Err.Number
Resume Next

End Sub
Public Sub ASpSv()
On Error GoTo SpErr
'Speichert die Einstellungen des GridEx

Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl

Set FM = frmAdress
Set RpCo2 = FM.repCont2
Set RpCo3 = FM.repCont3

IniSetVal "RpCnt2a", "SSpLa2a", RpCo2.SaveSettings
IniSetVal "RpCnt3a", "SSpLa3a", RpCo3.SaveSettings

Set RpCo2 = Nothing
Set RpCo3 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ASpSv " & Err.Number
Resume Next

End Sub
Public Sub ATerm(Optional ByVal PasWo As Boolean = False)
On Error GoTo NeErr
'Versenden einer SMS

Dim MitNr As Long
Dim ManNr As Long
Dim NuStr As String
Dim TxStr As String
Dim PaStr As String
Dim PrItm As XtremePropertyGrid.PropertyGridItem
Dim PrKat As XtremePropertyGrid.PropertyGridItem
Dim PrIts As XtremePropertyGrid.PropertyGridItems

Set FM = frmAdress
Set TxTel = FM.txtS1F18
Set PrGr1 = FM.prpGrid1
Set PrIts = PrGr1.Categories

If WindowLoad("frmSMS") = True Then
    Unload frmSMS
    DoEvents
End If

MitNr = GlMiA(GlSmI, 2) 'Standardmitarbeiter
ManNr = GlMan(GlSMa, 2) 'Standardmandant

For Each PrKat In PrIts
    For Each PrItm In PrKat.Childs
        If Mid$(PrItm.Tag, 2, Len(PrItm.Tag) - 1) = "Em_Pass" Then
            If PrItm.Value <> vbNullString Then
                PaStr = PrItm.Value
            Else
                PaStr = vbNullString
            End If
            Exit For
        End If
    Next PrItm
Next PrKat

If TxTel.Text <> vbNullString Then
    NuStr = TxTel.Text
Else
    NuStr = vbNullString
End If

If NuStr <> vbNullString Then
    If PasWo = True Then
        If PaStr <> vbNullString Then
            If UBound(GlEmT) > 10 Then
                If GlEmT(11, 1) <> vbNullString Then
                    With GlTxV
                        .TxStr = GlEmT(11, 1)
                        .MitNr = MitNr
                        .ManNr = ManNr
                        .PatNr = GlAdr
                        .PasWo = PaStr
                    End With
                    TxStr = SEmTx()
                Else
                    TxStr = "Ihr Verschlüsselungskennwort lautet: " & PaStr
                End If
            Else
                TxStr = "Ihr Verschlüsselungskennwort lautet: " & PaStr
            End If
            frmSMS.NaTex = TxStr
            frmSMS.NaNum = NuStr
            frmSMS.Show
        Else
            SPopu "Kein Passwort", "Für den Patienten ist kein Verschlüsselungskennwort vorhanden", IC48_Forbidden
        End If
    Else
        frmSMS.NaTex = vbNullString
        frmSMS.NaNum = NuStr
        frmSMS.Show
    End If
Else
    SPopu "Keine Mobilfunknummer", "Für den Patienten ist keine Mobilfunknummer vorhanden ", IC48_Forbidden
End If

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ATerm " & Err.Number
Resume Next

End Sub
Public Sub AZuLa()
On Error GoTo SpErr
'Lädt Informationen aus der Tabelle in die Felder

Dim S4F01 As XtremeSuiteControls.FlatEdit
Dim S4F03 As XtremeSuiteControls.FlatEdit
Dim S4F04 As XtremeSuiteControls.FlatEdit
Dim S4F05 As XtremeSuiteControls.FlatEdit
Dim S4F06 As XtremeSuiteControls.FlatEdit
Dim S4F08 As XtremeSuiteControls.FlatEdit
Dim S4F09 As XtremeSuiteControls.FlatEdit
Dim S4F15 As XtremeSuiteControls.FlatEdit
Dim S4F16 As XtremeSuiteControls.FlatEdit
Dim S4F17 As XtremeSuiteControls.FlatEdit
Dim S4F18 As XtremeSuiteControls.FlatEdit
Dim S4F19 As XtremeSuiteControls.FlatEdit
Dim S4F02 As XtremeSuiteControls.ComboBox
Dim S4F11 As XtremeSuiteControls.ComboBox
Dim S4F12 As XtremeSuiteControls.ComboBox

Dim GebDa As Date
Dim IdxNr As Long
Dim BriAn As String
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmAdress
Set Rahm4 = FM.frmRahm4
Set S4F01 = FM.txtS4F01
Set S4F02 = FM.txtS4F02
Set S4F03 = FM.txtS4F03
Set S4F04 = FM.txtS4F04
Set S4F05 = FM.txtS4F05
Set S4F06 = FM.txtS4F06
Set S4F08 = FM.txtS4F08
Set S4F09 = FM.txtS4F09
Set S4F15 = FM.txtS4F15
Set S4F16 = FM.txtS4F16
Set S4F17 = FM.txtS4F17
Set S4F18 = FM.txtS4F18
Set S4F19 = FM.txtS4F19
Set S4F11 = FM.cmbS4F11
Set S4F12 = FM.cmbS4F12
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns
Set RpSel = RpCo5.SelectedRows

If RpSel.Count > 0 Then
    If Rahm4.Enabled = False Then
        Rahm4.Enabled = True
    End If
    Set RpRow = RpSel(0)
    Set RpCol = RpCls.Find(Adr_ID0)
    IdxNr = RpRow.Record(RpCol.ItemIndex).Value
    Set RpCol = RpCls.Find(Adr_Firma1)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F01.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_TreKey)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F02.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Titel)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F03.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Vorname)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F04.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Name)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F05.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Straße)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F06.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_PLZ)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F08.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Ort)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F09.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Land)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F12.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_IDKurz)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F19.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Telefon1)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F15.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Geboren)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F16.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Telefon5)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                GebDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
                S4F18.Text = GebDa
            End If
        End If
    End If
    Set RpCol = RpCls.Find(Adr_Anschrift)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        S4F17.Text = RpRow.Record(RpCol.ItemIndex).Value
    End If
    Set RpCol = RpCls.Find(Adr_Briefanrede)
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        BriAn = RpRow.Record(RpCol.ItemIndex).Value
    End If
    If IsNull(RpRow.Record(RpCol.ItemIndex).Value) = False Then
        Set RpCol = RpCls.Find(Adr_GuiID)
        GlAzG = RpRow.Record(RpCol.ItemIndex).Value
    End If
    DoEvents
    If BriAn <> vbNullString Then
        AdBri BriAn
    End If
Else
    If Rahm4.Enabled = True Then
        Rahm4.Enabled = False
    End If
    S4F01.Text = vbNullString
    S4F02.Text = vbNullString
    S4F03.Text = vbNullString
    S4F04.Text = vbNullString
    S4F05.Text = vbNullString
    S4F06.Text = vbNullString
    S4F08.Text = vbNullString
    S4F09.Text = vbNullString
    S4F11.Text = vbNullString
    S4F15.Text = vbNullString
    S4F16.Text = vbNullString
    S4F17.Text = vbNullString
    S4F18.Text = vbNullString
    S4F19.Text = vbNullString
    S4F12.ListIndex = 0
End If

Set RpCo5 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AZuLa " & Err.Number
Resume Next

End Sub
Public Sub AZuNe()
On Error GoTo NeErr
'Bereitet die Neueingabe einer Adresse vor

Dim S4F01 As XtremeSuiteControls.FlatEdit
Dim S4F03 As XtremeSuiteControls.FlatEdit
Dim S4F04 As XtremeSuiteControls.FlatEdit
Dim S4F05 As XtremeSuiteControls.FlatEdit
Dim S4F06 As XtremeSuiteControls.FlatEdit
Dim S4F08 As XtremeSuiteControls.FlatEdit
Dim S4F09 As XtremeSuiteControls.FlatEdit
Dim S4F15 As XtremeSuiteControls.FlatEdit
Dim S4F16 As XtremeSuiteControls.FlatEdit
Dim S4F17 As XtremeSuiteControls.FlatEdit
Dim S4F18 As XtremeSuiteControls.FlatEdit
Dim S4F19 As XtremeSuiteControls.FlatEdit
Dim S4F02 As XtremeSuiteControls.ComboBox
Dim S4F11 As XtremeSuiteControls.ComboBox
Dim S4F12 As XtremeSuiteControls.ComboBox

Set FM = frmAdress
Set Rahm4 = FM.frmRahm4
Set S4F01 = FM.txtS4F01
Set S4F02 = FM.txtS4F02
Set S4F03 = FM.txtS4F03
Set S4F04 = FM.txtS4F04
Set S4F05 = FM.txtS4F05
Set S4F06 = FM.txtS4F06
Set S4F08 = FM.txtS4F08
Set S4F09 = FM.txtS4F09
Set S4F15 = FM.txtS4F15
Set S4F16 = FM.txtS4F16
Set S4F17 = FM.txtS4F17
Set S4F18 = FM.txtS4F18
Set S4F19 = FM.txtS4F19
Set S4F11 = FM.cmbS4F11
Set S4F12 = FM.cmbS4F12

If Rahm4.Enabled = False Then
    Rahm4.Enabled = True
End If

S4F01.Text = vbNullString
S4F02.Text = vbNullString
S4F03.Text = vbNullString
S4F04.Text = vbNullString
S4F05.Text = vbNullString
S4F06.Text = vbNullString
S4F08.Text = vbNullString
S4F09.Text = vbNullString
S4F11.Text = vbNullString
S4F15.Text = vbNullString
S4F16.Text = vbNullString
S4F17.Text = vbNullString
S4F18.Text = vbNullString
S4F19.Text = vbNullString

S4F12.ListIndex = 0

GlAzN = True
GlAzG = CreateID("A")

S4F01.SetFocus

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "AZuNe " & Err.Number
Resume Next

End Sub
Public Function ChipLe(Coma As String) As String
On Error Resume Next
'Liest einen einzelnen Eintrag aus der Chopkarte

Dim ErgL As Long
Dim CoLa As Long
Dim DaLa As Long
Dim CaDa As String
Dim ErgS As String
Dim Eing As String
Dim Kart As Boolean
Dim PoLe As Integer

ChipLe = vbNullString
CoLa = Len(Coma)
Eing = 0&
DaLa = Len(Eing)
ErgS = String$(200, 0)
ErgL = SCardComand(0, Coma, CoLa, Eing, DaLa, ErgS, 200)
If Len(ErgS) > 0 Then
    PoLe = InStr(1, ErgS, vbNullChar, 1)
    If PoLe > 0 Then
        ChipLe = Left$(ErgS, PoLe - 1)
    End If
End If

End Function
Public Sub MInit()
On Error GoTo NeErr
'Bereitet die Neueingabe einer Adresse vor

Dim RetWe As Long
Dim IdxNr As Long
Dim TagWe As String
Dim Frage As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim AkZei As Integer
Dim ZeiUm As Boolean

Dim ImMan As XtremeCommandBars.ImageManager
Dim TxDum As XtremeSuiteControls.FlatEdit
Dim TxRch As XtremeSuiteControls.FlatEdit
Dim ChThe As XtremeSuiteControls.CheckBox
Dim ChOut As XtremeSuiteControls.CheckBox
Dim ChS01 As XtremeSuiteControls.CheckBox
Dim ChS02 As XtremeSuiteControls.CheckBox
Dim ChS03 As XtremeSuiteControls.CheckBox
Dim ChS04 As XtremeSuiteControls.CheckBox
Dim ChS05 As XtremeSuiteControls.CheckBox
Dim ChS06 As XtremeSuiteControls.CheckBox
Dim ChS07 As XtremeSuiteControls.CheckBox
Dim ChS08 As XtremeSuiteControls.CheckBox
Dim ChS09 As XtremeSuiteControls.CheckBox
Dim ChS10 As XtremeSuiteControls.CheckBox
Dim ChS11 As XtremeSuiteControls.CheckBox
Dim ChS12 As XtremeSuiteControls.CheckBox
Dim ChS13 As XtremeSuiteControls.CheckBox
Dim ChS14 As XtremeSuiteControls.CheckBox
Dim ChTeS As XtremeSuiteControls.CheckBox
Dim ChDef As XtremeSuiteControls.CheckBox
Dim ChOnT As XtremeSuiteControls.CheckBox
Dim cmBuLa As XtremeSuiteControls.ComboBox
Dim cmKVBz As XtremeSuiteControls.ComboBox
Dim cmKant As XtremeSuiteControls.ComboBox
Dim cmRas1 As XtremeSuiteControls.ComboBox
Dim cmRas2 As XtremeSuiteControls.ComboBox
Dim cmMaxT As XtremeSuiteControls.ComboBox
Dim cmMaxP As XtremeSuiteControls.ComboBox
Dim cmVorl As XtremeSuiteControls.ComboBox
Dim cmBuRa As XtremeSuiteControls.ComboBox
Dim cmNoti As XtremeSuiteControls.ComboBox
Dim cmKata As XtremeSuiteControls.ComboBox
Dim cmKett As XtremeSuiteControls.ComboBox
Dim cmKet2 As XtremeSuiteControls.ComboBox
Dim cmRahm As XtremeSuiteControls.ComboBox
Dim cmKont As XtremeSuiteControls.ComboBox
Dim cmKon2 As XtremeSuiteControls.ComboBox
Dim cmGeK1 As XtremeSuiteControls.ComboBox
Dim cmGeK2 As XtremeSuiteControls.ComboBox
Dim cmReTy As XtremeSuiteControls.ComboBox
Dim cmSteu As XtremeSuiteControls.ComboBox
Dim cmStKt As XtremeSuiteControls.ComboBox
Dim PuPo1 As XtremeSuiteControls.PushButton
Dim PuSi1 As XtremeSuiteControls.PushButton
Dim MoKal As XtremeCalendarControl.DatePicker
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMandant
Set MoKal = FM.dtpDatu1
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set Rahm8 = FM.frmRahm8
Set RpCo5 = FM.repCont5
Set ChThe = FM.chkOpti2
Set TxNum = FM.txtS1F30
Set TxGeb = FM.txtS1F13
Set TxBri = FM.txtS1F20
Set FeAn1 = FM.txtS1F02
Set FeLa1 = FM.txtS1F12
Set TxFir = FM.txtS1F01
Set TxOrt = FM.txtS1F09
Set TxErs = FM.txtS2F27
Set TxDum = FM.txtDummy
Set TxRch = FM.txtS4F01
Set PuPo1 = FM.btnPost1
Set PuSi1 = FM.btnSign1
Set ChS01 = FM.chkBox01
Set ChS02 = FM.chkBox02
Set ChS03 = FM.chkBox03
Set ChS04 = FM.chkBox04
Set ChS05 = FM.chkBox05
Set ChS06 = FM.chkBox06
Set ChS07 = FM.chkBox07
Set ChS08 = FM.chkBox08
Set ChS09 = FM.chkBox09
Set ChS10 = FM.chkBox10
Set ChS11 = FM.chkBox11
Set ChS12 = FM.chkBox12
Set ChS13 = FM.chkBox13
Set ChS14 = FM.chkBox14
Set CmFch = FM.cmbKatal
Set FePat = FM.cmbS2F10
Set FeGes = FM.cmbS1F08
Set FeFam = FM.txtS2F26
Set ChTeS = FM.chkKaAus
Set ChDef = FM.chkDefra
Set ChOnT = FM.chkOnlTe
Set cmRas1 = FM.cmbRast1
Set cmRas2 = FM.cmbRast2
Set cmMaxT = FM.cmbMaxTe
Set cmMaxP = FM.cmbMaxPa
Set cmVorl = FM.cmbVorLa
Set cmBuRa = FM.cmbBuRad
Set cmNoti = FM.cmbNotVa
Set cmBuLa = FM.cmbBuLnd
Set cmKVBz = FM.cmbAbrBz
Set cmKant = FM.cmbKanto
Set cmKata = FM.cmbGbKat
Set cmKett = FM.cmbGbKet
Set cmKet2 = FM.cmbGbKe2
Set cmRahm = FM.cmbKtoRa
Set cmStKt = FM.cmbKtoSt
Set cmKont = FM.cmbKtoEr
Set cmKon2 = FM.cmbKtoEk
Set cmGeK1 = FM.cmbGeKt1
Set cmGeK2 = FM.cmbGeKt2
Set cmReTy = FM.cmbReTyp
Set cmSteu = FM.cmbSteue
Set RpCls = RpCo5.Columns
Set RpRcs = RpCo5.Records
Set ImMan = frmMain.imgManag

ZeiUm = False

PuPo1.Icon = ImMan.Icons.GetImage(IC16_Mailbox, 16)
PuSi1.Icon = ImMan.Icons.GetImage(IC16_Folder_Open, 16)

For Each AktCo In FM.Controls
    If TypeName(AktCo) = "FlatEdit" Then
        If Left$(AktCo.Name, 5) = "txtSp" Then
            AktCo.SetMask "00:00", "__:__"
        End If
    End If
Next AktCo

With MoKal
    .AllowNoncontinuousSelection = False
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    If GlSty = 8 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    ElseIf GlSty = 7 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    Else
        .BorderStyle = xtpDatePickerBorderOffice
    End If
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    .MaxSelectionCount = 1
    .RightToLeft = False
    .ShowNoneButton = True
    .ShowNonMonthDays = True
    .ShowTodayButton = True
    .ShowWeekNumbers = False
    .TextNoneButton = "Keine"
    .TextTodayButton = "Heute"
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Behandlungstag"
    .MonthDelta = 1
    .YearsTriangle = False
    Select Case GlSty
    Case 8: .VisualTheme = xtpCalendarThemeResource
    Case 7: .VisualTheme = xtpCalendarThemeResource
    Case Else: .VisualTheme = xtpCalendarThemeResource
    End Select
    .PaintManager.ButtonTextColor = -2147483640
    .PaintManager.ControlBackColor = -2147483643
    .PaintManager.DayBackColor = -2147483643
    .PaintManager.DayTextColor = -2147483640
    .PaintManager.DaysOfWeekBackColor = -2147483643
    .PaintManager.DaysOfWeekTextColor = -2147483640
    .PaintManager.ListControlBackColor = -2147483643
    .PaintManager.ListControlTextColor = -2147483640
    .PaintManager.NonMonthDayBackColor = -2147483643
    .PaintManager.NonMonthDayTextColor = -2147483640
    .PaintManager.SelectedDayBackColor = GlFac
    .PaintManager.SelectedDayTextColor = -2147483640
    .PaintManager.WeekNumbersBackColor = -2147483643
    .PaintManager.WeekNumbersTextColor = -2147483640
    .PaintManager.MonthHeaderBackColor = GlMoB
End With

With RpCo5
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = True
    .AllowEditPreview = False
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Rechte vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Rechte vorhanden"
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
    .PaintManager.MaxPreviewLines = 3
    .PaintManager.ThemedInplaceButtons = True
    .PaintManager.HorizontalGridStyle = xtpGridNoLines
    .PaintManager.VerticalGridStyle = xtpGridNoLines
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
    .ShowHeader = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

For AktZa = 0 To UBound(GlFch) 'Fachrichtungen
    With CmFch
        .AddItem GlFch(AktZa, 0)
        .ItemData(AktZa) = GlFch(AktZa, 1)
    End With
Next AktZa

For AktZa = 0 To UBound(GlBsl) 'Bundesländer
    With cmBuLa
        .AddItem GlBsl(AktZa, 1) & " " & GlBsl(AktZa, 0)
        .ItemData(AktZa) = GlBsl(AktZa, 2)
    End With
Next AktZa

For AktZa = 0 To UBound(GlKVB) 'KV Bezirke
    With cmKVBz
        .AddItem GlKVB(AktZa, 1) & " " & GlKVB(AktZa, 0)
        .ItemData(AktZa) = GlKVB(AktZa, 2)
    End With
Next AktZa

For AktZa = 0 To UBound(GlKtn) 'Kantone
    With cmKant
        .AddItem GlKtn(AktZa, 0)
        .ItemData(AktZa) = GlKtn(AktZa, 2)
    End With
Next AktZa

With FeAn1 'Anreden
    For AktZa = 0 To UBound(GlAnr) - 1
        .AddItem GlAnr(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With
For AktZa = 1 To UBound(GlLan)
    With FeLa1
        .AddItem GlLan(AktZa, 1)
        .ItemData(AktZa - 1) = GlLan(AktZa, 0)
    End With
Next AktZa

With FeGes 'Geschlecht
    For AktZa = 0 To UBound(GlGes) - 1
        .AddItem GlGes(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With FeFam
    .AddItem "Ledig"
    .ItemData(0) = 1
    .AddItem "Verheiratet"
    .ItemData(1) = 2
    .AddItem "Verwittwet"
    .ItemData(2) = 3
    .AddItem "Geschieden"
    .ItemData(3) = 4
    .AddItem "Getrennt"
    .ItemData(4) = 5
    .AddItem "Unbekannt"
    .ItemData(5) = 6
End With

With cmRahm
    For AktZa = 1 To UBound(GlKoR) 'Standardkontenrahmen
        .AddItem GlKoR(AktZa, 0)
        .ItemData(AktZa - 1) = GlKoR(AktZa, 1)
    Next AktZa
End With

With cmReTy
    .AddItem "R - Standardrechnung"
    .ItemData(0) = 1
    .AddItem "L - Laborrechnung"
    .ItemData(1) = 2
    .AddItem "A - Abrechnungsstelle"
    .ItemData(2) = 3
    .AddItem "M - Rechnungsauftrag"
    .ItemData(3) = 4
    .AddItem "G - Gewerberechnung"
    .ItemData(4) = 5
    .AddItem "I - Importrechnung"
    .ItemData(5) = 6
End With

For AktZa = 1 To UBound(GlGKa) 'Gebührenkataloge
    cmKata.AddItem GlGKa(AktZa, 1)
    cmKata.ItemData(AktZa - 1) = GlGKa(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlKet)
    cmKett.AddItem GlKet(AktZa, 2)
    cmKet2.AddItem GlKet(AktZa, 2)
    cmKett.ItemData(AktZa - 1) = GlKet(AktZa, 0)
    cmKet2.ItemData(AktZa - 1) = GlKet(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlErK)
    cmKont.AddItem GlErK(AktZa, 1)
    cmKont.ItemData(AktZa - 1) = GlErK(AktZa, 3) '[IDI]
Next AktZa

For AktZa = 1 To UBound(GlErK)
    cmKon2.AddItem GlErK(AktZa, 1)
    cmKon2.ItemData(AktZa - 1) = GlErK(AktZa, 3) '[IDI]
Next AktZa

For AktZa = 1 To UBound(GlSaU) 'Sachkonten mit Steuerkontenzuordnung
    cmStKt.AddItem GlSaU(AktZa, 3)
    cmStKt.ItemData(AktZa - 1) = GlSaU(AktZa, 2)
Next AktZa

If GlBuc = True Then 'einfache Buchhaltung verwenden
    For AktZa = 1 To UBound(GlGeK) 'Geldkonten
        cmGeK1.AddItem GlGeK(AktZa, 3)
        cmGeK1.ItemData(cmGeK1.NewIndex) = GlGeK(AktZa, 0) '[IDB]
        cmGeK2.AddItem GlGeK(AktZa, 3)
        cmGeK2.ItemData(cmGeK2.NewIndex) = GlGeK(AktZa, 0) '[IDB]
    Next AktZa
Else
    For AktZa = 1 To UBound(GlGeK) 'Geldkonten
        For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
            If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                cmGeK1.AddItem GlSaK(AktKo, 3)
                cmGeK1.ItemData(cmGeK1.NewIndex) = GlSaK(AktKo, 6) '[IDB]
                cmGeK2.AddItem GlSaK(AktKo, 3)
                cmGeK2.ItemData(cmGeK2.NewIndex) = GlSaK(AktKo, 6) '[IDB]
                Exit For
            End If
        Next AktKo
    Next AktZa
    If cmGeK1.ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            cmGeK1.AddItem GlGeK(AktZa, 3)
            cmGeK1.ItemData(cmGeK1.NewIndex) = GlGeK(AktZa, 0) '[IDB]
            cmGeK2.AddItem GlGeK(AktZa, 3)
            cmGeK2.ItemData(cmGeK2.NewIndex) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
    End If
End If

For AktZa = 1 To UBound(GlStu)
    cmSteu.AddItem GlStu(AktZa, 2)
    cmSteu.ItemData(cmSteu.NewIndex) = GlStu(AktZa, 0) '[IDS]
Next AktZa

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
    Set RpCol = .Add(1, vbNullString, 100, True)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
        .AutoSize = True
    End With
    If RpCo5.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
End With

For AktZa = 0 To GlZaR - 1 'Rechteanzahl
    Set RpRec = RpRcs.Add()
    Set RpItm = RpRec.AddItem(vbNullString)
    With RpItm
        .HasCheckbox = True
        .Focusable = True
    End With
    Set RpItm = RpRec.AddItem(GlRch(1, AktZa))
    With RpItm
        .Focusable = False
    End With
Next AktZa

If AktZa > 0 Then
    RpCo5.Populate
    Set RpRws = RpCo5.Rows
    RpRws.Row(0).Selected = False
End If

TxNum.Pattern = "\d*"
TxGeb.SetMask "00.00.0000", "__.__.____"

Select Case GlBut
Case RibTab_Mandanten: Rahm7.Caption = "Sprechzeiten"
Case RibTab_Verordner: Rahm7.Caption = "Sprechzeiten"
Case RibTab_Mitarbeit: Rahm7.Caption = "Arbeitszeiten"
Case Else: Rahm7.Caption = "Sprechzeiten"
End Select

With cmRas1
    AkZei = 1
    For AktZa = 1 To UBound(GlTku)
        If GlTku(AktZa, 1) = 0 Then
            .AddItem GlTku(AktZa, 0)
            .ItemData(AkZei - 1) = AktZa
            AkZei = AkZei + 1
        End If
    Next AktZa
    .ListIndex = GlZeR - 1 'Zeitrasterindex
End With

With cmRas2
    AkZei = 1
    For AktZa = 1 To UBound(GlTku)
        If GlTku(AktZa, 1) = 0 Then
            .AddItem GlTku(AktZa, 0)
            .ItemData(AkZei - 1) = AktZa
            AkZei = AkZei + 1
        End If
    Next AktZa
    .ListIndex = GlZeR - 1 'Zeitrasterindex
End With

With cmMaxT
    .AddItem "Alle Term."
    .ItemData(0) = 0
    For AktZa = 1 To 9
        .AddItem AktZa & " Term."
        .ItemData(AktZa) = AktZa
    Next AktZa
    .ListIndex = 0
End With

With cmMaxP
    .AddItem "Alle Term."
    .ItemData(0) = 0
    For AktZa = 1 To 9
        .AddItem AktZa & " Term."
        .ItemData(AktZa) = AktZa
    Next AktZa
    .ListIndex = 1
End With

With cmVorl
    For AktZa = 1 To 24
        .AddItem Format$(AktZa, "00") & " Std."
        .ItemData(AktZa - 1) = AktZa
    Next AktZa
    .ListIndex = 1
End With

With cmBuRa
    For AktZa = 1 To 36
        .AddItem Format$(AktZa, "00") & " Mon."
        .ItemData(AktZa - 1) = AktZa
    Next AktZa
    .ListIndex = 23
End With

With cmNoti
    For AktZa = 0 To 48
        .AddItem AktZa & " Std."
        .ItemData(AktZa) = AktZa
    Next AktZa
    .ListIndex = 24
    .Enabled = GlTeE 'Email-Termin-Erinnerung
End With

If GlMVo = False Then 'mandantenbezogene Vorgaben verwenden
    cmKata.Enabled = False
    cmKett.Enabled = False
    cmKet2.Enabled = False
    cmRahm.Enabled = False
    cmKont.Enabled = False
    cmKon2.Enabled = False
    cmGeK1.Enabled = False
    cmGeK2.Enabled = False
    cmReTy.Enabled = False
    cmSteu.Enabled = False
    cmStKt.Enabled = False
End If

ChOnT.Enabled = GlOTS 'Online-Terminbuchungs System

FePat.Enabled = True
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Rahm6.BackColor = GlBak
Rahm7.BackColor = GlBak
Rahm8.BackColor = GlBak
ChThe.BackColor = GlBak
ChS01.BackColor = GlBak
ChS02.BackColor = GlBak
ChS03.BackColor = GlBak
ChS04.BackColor = GlBak
ChS05.BackColor = GlBak
ChS06.BackColor = GlBak
ChS07.BackColor = GlBak
ChS08.BackColor = GlBak
ChS09.BackColor = GlBak
ChS10.BackColor = GlBak
ChS11.BackColor = GlBak
ChS12.BackColor = GlBak
ChS13.BackColor = GlBak
ChS14.BackColor = GlBak
ChTeS.BackColor = GlBak
ChDef.BackColor = GlBak
ChOnT.BackColor = GlBak
FM.BackColor = GlBak

Set RpCo5 = Nothing

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MInit " & Err.Number
Resume Next

End Sub
Public Sub MKopi()
On Error GoTo CrErr
'Kopieren der Adressdaten in die Rechungsanschrift

Dim TagWe As String
Dim AdRGe As Boolean
Dim Mld1, Tit1 As String

Set FM = frmMandant

If FM.txtS1F02.Text <> vbNullString Then
    TagWe = Mid$(FM.txtS2F12.Tag, 2, Len(FM.txtS2F12.Tag) - 1)
    FM.txtS2F12.Text = FM.txtS1F02.Text
    FM.txtS2F12.Tag = 1 & TagWe
End If

If FM.txtS1F03.Text <> vbNullString Then FM.txtS2F13.Text = FM.txtS1F03.Text
If FM.txtS1F04.Text <> vbNullString Then FM.txtS2F14.Text = FM.txtS1F04.Text
If FM.txtS1F05.Text <> vbNullString Then FM.txtS2F15.Text = FM.txtS1F05.Text
If FM.txtS1F01.Text <> vbNullString Then FM.txtS2F11.Text = FM.txtS1F01.Text
If FM.cmbS1F10.Text <> vbNullString Then FM.txtS2F20.Text = FM.cmbS1F10.Text
If FM.txtS1F12.Text <> vbNullString Then FM.txtS2F22.Text = FM.txtS1F12.Text
If FM.txtS1F13.Text <> vbNullString Then FM.txtS2F25.Text = FM.txtS1F13.Text
If FM.txtS1F06.Text <> vbNullString Then FM.txtS2F16.Text = FM.txtS1F06.Text
If FM.txtS1F08.Text <> vbNullString Then FM.txtS2F18.Text = FM.txtS1F08.Text
If FM.txtS1F09.Text <> vbNullString Then FM.txtS2F19.Text = FM.txtS1F09.Text

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MKopi " & Err.Number
Resume Next

End Sub
Private Sub MLoa()
On Error GoTo NeErr
'Lädt die Benutzerdaten aus der INI

Dim AktZa As Integer
Dim TmIdx As Integer
Dim LiIdx As Integer
Dim IdErB As Integer
Dim IdErk As Long
Dim IdStK As Integer
Dim TxDum As XtremeSuiteControls.FlatEdit
Dim txVorn As XtremeSuiteControls.FlatEdit
Dim txName As XtremeSuiteControls.FlatEdit
Dim txStra As XtremeSuiteControls.FlatEdit
Dim txPost As XtremeSuiteControls.FlatEdit
Dim txOrte As XtremeSuiteControls.FlatEdit
Dim txTele As XtremeSuiteControls.FlatEdit
Dim txFaxe As XtremeSuiteControls.FlatEdit
Dim txBank As XtremeSuiteControls.FlatEdit
Dim txBaLZ As XtremeSuiteControls.FlatEdit
Dim txKont As XtremeSuiteControls.FlatEdit
Dim txSteu As XtremeSuiteControls.FlatEdit
Dim txIKNr As XtremeSuiteControls.FlatEdit
Dim txBeru As XtremeSuiteControls.FlatEdit
Dim txTite As XtremeSuiteControls.FlatEdit
Dim TxEmai As XtremeSuiteControls.FlatEdit
Dim TxIntr As XtremeSuiteControls.FlatEdit
Dim TxPrax As XtremeSuiteControls.FlatEdit
Dim TxIBAN As XtremeSuiteControls.FlatEdit
Dim TxIBA2 As XtremeSuiteControls.FlatEdit
Dim TxBIC1 As XtremeSuiteControls.FlatEdit
Dim TxBIC2 As XtremeSuiteControls.FlatEdit
Dim txGlID As XtremeSuiteControls.FlatEdit
Dim TxAbre As XtremeSuiteControls.FlatEdit
Dim TxLand As XtremeSuiteControls.ComboBox
Dim TxAnre As XtremeSuiteControls.ComboBox
Dim cmFach As XtremeSuiteControls.ComboBox
Dim cmBuLa As XtremeSuiteControls.ComboBox
Dim cmKVBz As XtremeSuiteControls.ComboBox
Dim cmKant As XtremeSuiteControls.ComboBox
Dim cmRas1 As XtremeSuiteControls.ComboBox
Dim cmRas2 As XtremeSuiteControls.ComboBox
Dim cmMaxT As XtremeSuiteControls.ComboBox
Dim cmMaxP As XtremeSuiteControls.ComboBox
Dim cmVorl As XtremeSuiteControls.ComboBox
Dim cmBuRa As XtremeSuiteControls.ComboBox
Dim cmNoti As XtremeSuiteControls.ComboBox
Dim ChTeSp As XtremeSuiteControls.CheckBox
Dim ChDefr As XtremeSuiteControls.CheckBox
Dim ChOnTe As XtremeSuiteControls.CheckBox
Dim cmKata As XtremeSuiteControls.ComboBox
Dim cmKett As XtremeSuiteControls.ComboBox
Dim cmKet2 As XtremeSuiteControls.ComboBox
Dim cmRahm As XtremeSuiteControls.ComboBox
Dim cmKont As XtremeSuiteControls.ComboBox
Dim cmKon2 As XtremeSuiteControls.ComboBox
Dim cmGeK1 As XtremeSuiteControls.ComboBox
Dim cmGeK2 As XtremeSuiteControls.ComboBox
Dim cmReTy As XtremeSuiteControls.ComboBox
Dim cmSteu As XtremeSuiteControls.ComboBox
Dim cmStKt As XtremeSuiteControls.ComboBox

Set FM = frmMandant
Set TxDum = FM.txtDummy
Set TxGeb = FM.txtS1F13
Set txIKNr = FM.txtIKNum
Set TxPrax = FM.txtS2F11
Set txVorn = FM.txtS1F04
Set txName = FM.txtS1F05
Set txStra = FM.txtS1F06
Set txPost = FM.txtS1F08
Set txOrte = FM.txtS1F09
Set txTele = FM.txtS1F16
Set txFaxe = FM.txtS1F17
Set txBank = FM.txtS2F03
Set txBaLZ = FM.txtS2F04
Set txKont = FM.txtS2F05
Set txSteu = FM.txtS1F22
Set txBeru = FM.txtS2F24
Set txTite = FM.txtS1F03
Set TxEmai = FM.txtS1F19
Set TxIntr = FM.txtS1F27
Set TxAnre = FM.txtS1F02
Set TxIBAN = FM.txtS2F33
Set TxIBA2 = FM.txtIBAN2
Set TxBIC1 = FM.txtS2F34
Set TxBIC2 = FM.txtBICN2
Set txGlID = FM.txtGIDNr
Set TxLand = FM.txtS1F12
Set TxAbre = FM.txtS1F23
Set cmBuLa = FM.cmbBuLnd
Set cmKVBz = FM.cmbAbrBz
Set cmKant = FM.cmbKanto
Set cmFach = FM.cmbKatal
Set cmRas1 = FM.cmbRast1
Set cmRas2 = FM.cmbRast2
Set cmMaxT = FM.cmbMaxTe
Set cmMaxP = FM.cmbMaxPa
Set cmVorl = FM.cmbVorLa
Set cmBuRa = FM.cmbBuRad
Set cmNoti = FM.cmbNotVa
Set cmKata = FM.cmbGbKat
Set cmKett = FM.cmbGbKet
Set cmKet2 = FM.cmbGbKe2
Set cmRahm = FM.cmbKtoRa
Set cmKont = FM.cmbKtoEr
Set cmKon2 = FM.cmbKtoEk
Set cmGeK1 = FM.cmbGeKt1
Set cmGeK2 = FM.cmbGeKt2
Set cmStKt = FM.cmbKtoSt
Set cmReTy = FM.cmbReTyp
Set cmSteu = FM.cmbSteue
Set ChTeSp = FM.chkKaAus
Set ChDefr = FM.chkDefra
Set ChOnTe = FM.chkOnlTe

FM.AdAnd = True 'WICHTIG!

TxDum.Text = GlMId

TxAnre.Text = IniGetVal("Adress", "AAnre")
txVorn.Text = IniGetVal("Adress", "AVoNa")
txName.Text = IniGetVal("Adress", "AName")
txStra.Text = IniGetVal("Adress", "AStra")
txPost.Text = IniGetVal("Adress", "APLZ")
txOrte.Text = IniGetVal("Adress", "AOrt")
txTele.Text = IniGetVal("Adress", "ATele")
txFaxe.Text = IniGetVal("Adress", "AFax")
txBank.Text = IniGetVal("Adress", "ABank")
txBaLZ.Text = IniGetVal("Adress", "AnBLZ")
txKont.Text = IniGetVal("Adress", "AnKto")
txSteu.Text = IniGetVal("Adress", "ASteu")
txIKNr.Text = IniGetVal("Adress", "AIKNr")
txBeru.Text = IniGetVal("Adress", "ABeru")
txTite.Text = IniGetVal("Adress", "ATite")
TxEmai.Text = IniGetVal("Adress", "AEmail")
TxIntr.Text = IniGetVal("Adress", "AInter")
TxPrax.Text = IniGetVal("Adress", "APraxis")
TxIBAN.Text = IniGetVal("Adress", "AIBANr")
TxBIC1.Text = IniGetVal("Adress", "ABICNr")
TxIBA2.Text = IniGetVal("Adress", "AIBAN2")
TxBIC2.Text = IniGetVal("Adress", "ABICN2")
txGlID.Text = IniGetVal("Adress", "AGlIDr")
TxLand.Text = IniGetVal("Adress", "ALand")
TxGeb.Text = IniGetVal("Adress", "AGebo")

If IniGetVal("Adress", "AKata") <> vbNullString Then
    cmKata.ListIndex = CInt(IniGetVal("Adress", "AKata"))
Else
    cmKata.ListIndex = GlStK - 1
End If

If IniGetVal("Adress", "AKett") <> vbNullString Then
    cmKett.ListIndex = CInt(IniGetVal("Adress", "AKett"))
Else
    If (GlKe1 - 1) <= (cmKett.ListCount) - 1 Then
        cmKett.ListIndex = GlKe1 - 1
    Else
        cmKett.ListIndex = 0
    End If
End If

If IniGetVal("Adress", "AKet2") <> vbNullString Then
    cmKet2.ListIndex = CInt(IniGetVal("Adress", "AKet2"))
Else
    If (GlKe1 - 1) <= (cmKet2.ListCount) - 1 Then
        cmKet2.ListIndex = GlKe2 - 1
    Else
        cmKet2.ListIndex = 0
    End If
End If

If IniGetVal("Adress", "ARahm") <> vbNullString Then
    cmRahm.ListIndex = CInt(IniGetVal("Adress", "ARahm"))
Else
    If (GlKtR - 1) <= (cmRahm.ListCount) - 1 Then
        cmRahm.ListIndex = GlKtR - 1
    Else
        cmRahm.ListIndex = 0
    End If
End If

If IniGetVal("Adress", "AReTy") <> vbNullString Then
    LiIdx = CInt(IniGetVal("Adress", "AReTy"))
    cmReTy.ListIndex = LiIdx
Else
    Select Case GlReT 'Standardbelegtyp
    Case "R": LiIdx = 0
    Case "V": LiIdx = 1
    Case "L": LiIdx = 2
    Case "A": LiIdx = 3
    Case "U": LiIdx = 4
    Case "M": LiIdx = 5
    Case "G": LiIdx = 6
    Case "I": LiIdx = 7
    End Select
    cmReTy.ListIndex = LiIdx
End If

If IniGetVal("Adress", "ASteu") <> vbNullString Then
    cmSteu.ListIndex = CInt(IniGetVal("Adress", "ASteu"))
Else
    If cmSteu.ListCount >= GlStS Then
        cmSteu.ListIndex = GlStS - 1
    Else
        cmSteu.ListIndex = 0
    End If
End If

If IniGetVal("Adress", "AKont") <> vbNullString Then
    If CInt(IniGetVal("Adress", "AKont")) > 0 Then
        cmKont.ListIndex = CInt(IniGetVal("Adress", "AKont"))
    Else
        IdErk = SCmb(cmKont, GlSE1) 'Standarderlöskonto Kasse
        If IdErk >= 0 Then
            cmKont.ListIndex = IdErk
        Else
            cmKont.ListIndex = 0
        End If
    End If
Else
    IdErk = SCmb(cmKont, GlSE1) 'Standarderlöskonto Kasse
    If IdErk >= 0 Then
        cmKont.ListIndex = IdErk
    Else
        cmKont.ListIndex = 0
    End If
End If

If IniGetVal("Adress", "AKon2") <> vbNullString Then
    If CInt(IniGetVal("Adress", "AKon2")) > 0 Then
        cmKon2.ListIndex = CInt(IniGetVal("Adress", "AKon2"))
    Else
        IdErB = SCmb(cmKon2, GlSE2) 'Standarderlöskonto Bankkonto
        If IdErB >= 0 Then
            cmKon2.ListIndex = IdErB
        Else
            cmKon2.ListIndex = 0
        End If
    End If
Else
    IdErB = SCmb(cmKon2, GlSE2) 'Standarderlöskonto Bankkonto
    If IdErB >= 0 Then
        cmKon2.ListIndex = IdErB
    Else
        cmKon2.ListIndex = 0
    End If
End If

If IniGetVal("Adress", "AStKt") <> vbNullString Then
    If CInt(IniGetVal("Adress", "AStKt")) > 0 Then
        cmStKt.ListIndex = CInt(IniGetVal("Adress", "AStKt"))
    Else
        IdStK = SCmb(cmStKt, GlSKo) 'Standardsteuerkonto
        If IdStK >= 0 Then
            cmStKt.ListIndex = IdStK
        Else
            cmStKt.ListIndex = 0
        End If
    End If
Else
    IdStK = SCmb(cmStKt, GlSKo) 'Standardsteuerkonto
    If IdStK >= 0 Then
        cmStKt.ListIndex = IdStK
    Else
        cmStKt.ListIndex = 0
    End If
End If

If IniGetVal("Adress", "AGeK1") <> vbNullString Then
    If CInt(IniGetVal("Adress", "AGeK1")) > 0 Then
        cmGeK1.ListIndex = CLng(IniGetVal("Adress", "AGeK1"))
    Else
        cmGeK1.ListIndex = 0
    End If
Else
    cmGeK1.ListIndex = 0
End If

If IniGetVal("Adress", "AGeK2") <> vbNullString Then
    If CInt(IniGetVal("Adress", "AGeK2")) > 0 Then
        cmGeK2.ListIndex = CLng(IniGetVal("Adress", "AGeK2"))
    Else
        cmGeK2.ListIndex = 0
    End If
Else
    cmGeK2.ListIndex = 0
End If

If IniGetVal("Adress", "ATmPl") <> vbNullString Then
    If CBool(IniGetVal("Adress", "ATmPl")) = True Then
        ChTeSp.Value = xtpChecked
    Else
        ChTeSp.Value = xtpUnchecked
    End If
Else
    ChTeSp.Value = xtpUnchecked
End If

If IniGetVal("Adress", "ADefr") <> vbNullString Then
    If CBool(IniGetVal("Adress", "ADefr")) = True Then
        ChDefr.Value = xtpChecked
    Else
        ChDefr.Value = xtpUnchecked
    End If
Else
    ChDefr.Value = xtpUnchecked
End If

If IniGetVal("Adress", "AOnTe") <> vbNullString Then
    If CBool(IniGetVal("Adress", "AOnTe")) = True Then
        ChOnTe.Value = xtpChecked
        ChTeSp.Value = xtpUnchecked
    Else
        ChOnTe.Value = xtpUnchecked
    End If
Else
    ChOnTe.Value = xtpUnchecked
End If

TxAbre.Text = IniGetVal("System", "PVSNum")

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MLoa " & Err.Number
Resume Next

End Sub
Public Sub MMain(ByVal IdxNr As Long)
On Error GoTo LaErr

Dim FenBr As Long
Dim FenHo As Long
Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmMandant") = True Then
    Set FM = frmMandant
    frmMandant.ZOrder 0
    Exit Sub
End If

If GlTza = True Then 'Testzeit abgelaufen
    SPopu "Lizenzierung erforderlich!", "Es ist keine bzw. keine gültige Seriennummer vorhanden oder die Testzeit ist abgelaufen.", IC48_Forbidden
    Exit Sub
End If

GlAdL = True 'Formular wird geladen
GlMId = IdxNr

MReg

Load frmMandant

Set FM = frmMandant

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass

With clFen
    If GlBiA = False Then 'Bildschirmaktualisierung
        clFen.FenDsk 1
    Else
        clFen.FenDsk 2
    End If
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (774 / 2)
        .FeObn = (GlyGr / 2) - (700 / 2)
        .FeBre = 774
        .FeHoh = 710
    Else
        .FeLin = IniGetVal("ManForm", "FenLin")
        .FeObn = IniGetVal("ManForm", "FenObe")
        .FeBre = IniGetVal("ManForm", "FenBre")
        .FeHoh = IniGetVal("ManForm", "FenHoh")
    End If
End With

MMenu
MInit
AFont FM

If GlMId < 0 Then
    If GlMaV = True Then 'Mandanten vorhanden
        If GlMaA(GlSMa, 2) = 1 Then
            MNeu
            MLoa
        Else
            GlMId = GlMaA(GlSMa, 2)
            MNeu
            Man_Lad
            GlAdS = False
        End If
    Else
        MNeu
        MLoa
    End If
ElseIf GlMId = 0 Then
    MNeu True
ElseIf GlMId > 0 Then
    MNeu
    Man_Lad
    GlAdS = False
End If
DoEvents
MOpn

With clFen
    .FenMov
    DoEvents
    MPosi
    DoEvents
    Set CmBrs = FM.comBar01
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    .FenDsk 3
End With

Screen.MousePointer = vbNormal

Set clFen = Nothing

frmMandant.Show
DoEvents
GlAdL = False

If GlBut <> RibTab_Verordner Then
    Man_SpL
    DoEvents
End If

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MMain " & Err.Number
Resume Next

End Sub
Private Sub MMenu()
On Error GoTo CrErr
'Menue erstellen

Dim RetWe As Long
Dim TmDat As Date
Dim KeyNa As String
Dim AktWo As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbTem As XtremeCommandBars.RibbonTab
Dim MsBar As XtremeCommandBars.MessageBar
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmDat As XtremeCommandBars.CommandBarComboBox
Dim CmTyp As XtremeCommandBars.CommandBarComboBox
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set ImMan = frmMain.imgManag
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set MsBar = CmBrs.MessageBar
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

KeyNa = "ToolTips"

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(AM_Patient_Speichern, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Gruppe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Drucken, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Copy, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Del, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Clip1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Patient_Clip2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Notiz_Neu, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Notiz_Bearbeit, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Notiz_Loeschen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AM_Extras_Vorlage, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patient_Add, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patient_Copy, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patient_Del, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patienten_Save, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Patienten_Gruppe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Einzelbrief_Word, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_Drucken, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_SMS, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_GDT_Ex, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Adressen_GDT_Im, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Sprechzeit_Add, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Sprechzeit_Save, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Sprechzeit_Del, vbNullString, vbNullString, vbNullString, vbNullString)
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Width = 100
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Width = 160
    CmPan.Text = "Alter:"
    Set CmPan = .AddPane(3)
    CmPan.Width = 160
    CmPan.Text = "Ersteintrag:"
    Set CmPan = .AddPane(4)
    CmPan.Width = 160
    CmPan.Text = "Aktualisierung:"
    Set CmPan = .AddPane(59137)
    Set CmPan = .AddPane(59138)
    Set CmPan = .AddPane(59139)
    .Visible = True
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Geburtsdatum, "Geburtsdatum")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Present
    .ToolTipText = "Löscht das Geburtsdatum"
    .Style = xtpButtonIconAndCaption
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Hilfe, "Hilfe")
With CmBuT
    .ToolTipText = "Öffnet die Kurzhilfe"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Beenden, "Schließen")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .ToolTipText = "Schließt den Dialog"
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

Select Case GlBut
Case RibTab_Mandanten: '------------------------------------------------------------

    Set RbTab = RbBar.InsertTab(RibTab_Adr_Haupt, "Stammdaten")
    With RbTab
        .id = RibTab_Adr_Haupt
        .Visible = True
        .Selected = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mandant Suchen")
    With CmCon
        .IconId = IC32_BusMan_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Dokum, "Sprechzeiten")
    With RbTab
        .id = RibTab_Adr_Dokum
        .Visible = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With

    Set RbGrp = RbGps.AddGroup("Sprechzeiten", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Sprechzeit_Add, "Sprechzeiten Hinzufügen")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Doc_Add
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Sprechzeit_Save, "Sprechzeiten Speichern")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Document_Disk
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Sprechzeit_Del, "Sprechzeiten Entfernen")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Doc_Del
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With
    '---
    Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt2, "Verwendung :")
    With CmCon
        .ToolTipText = vbNullString
        .flags = xtpFlagRightAlign
        .BeginGroup = True
    End With
    Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt1, "Gültigkeit ab :")
    With CmCon
        .ToolTipText = vbNullString
        .flags = xtpFlagRightAlign
    End With
    Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt3, "Sprechzeiten :")
    With CmCon
        .ToolTipText = "Zeigt den aktuelklen Saldo"
        .flags = xtpFlagRightAlign
    End With

    Set CmTyp = RbGrp.Add(xtpControlComboBox, AD_Sprechzeit_Typen, vbNullString)
    With CmTyp
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .DropDownItemCount = 2
        .EditHint = "Sprechzeitentyp"
        .ToolTipText = "Wechselt zwischen starren und flexiblen Sprechzeiten"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 90
        .AddItem "starre Zeiten", 1
        .AddItem "flexible Zeiten", 2
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .ListIndex = 1
        Else
            .ListIndex = 2
        End If
    End With
    Set CmDat = RbGrp.Add(xtpControlComboBox, AD_Sprechzeit_Datum, vbNullString)
    With CmDat
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .DropDownItemCount = 10
        .EditHint = "Startdatum..."
        .ToolTipText = "Ab wann sollen die neuen Sprechzeiten gültig sein?"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 90
        AktWo = 1
        Do
        TmDat = SKaW2(AktWo, vbMonday, Year(Date))
        If TmDat >= Date Then
            .AddItem TmDat
        End If
        AktWo = AktWo + 1
        Loop Until Year(TmDat) > (Year(Date) + 2)
    End With
    Set CmCom = RbGrp.Add(xtpControlComboBox, AD_Sprechzeit_Auswa, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .DropDownItemCount = 5
        .EditHint = "Sprechzeiten..."
        .ToolTipText = "Bitte wählen Sie einen gespeicherten Sprechzeiteneintrag"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 90
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With
    
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Dokum, "Vorgaben")
    With RbTab
        .id = RibTab_Adr_Eigen
        .Visible = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mandant Suchen")
    With CmCon
        .IconId = IC32_BusMan_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Booki, "Buchungszeiten")
    With RbTab
        .id = RibTab_Adr_Booki
        .Visible = GlOTS
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mandant Suchen")
    With CmCon
        .IconId = IC32_BusMan_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
Case RibTab_Verordner: '------------------------------------------------------------

    Set RbTab = RbBar.InsertTab(RibTab_Adr_Haupt, "Stammdaten")
    With RbTab
        .id = RibTab_Adr_Haupt
        .Visible = True
        .Selected = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Verordner Hinzufügen")
    With CmCon
        .IconId = IC32_Doctor_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Verordner Kopieren")
    With CmCon
        .IconId = IC32_Doctor_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Verordner Entfernen")
    With CmCon
        .IconId = IC32_Doctor_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Verordner Speichern")
    With CmCon
        .IconId = IC32_Disk_Doctor
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Verordner Suchen")
    With CmCon
        .IconId = IC32_Doctor_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Dokum, "Sprechzeiten")
    With RbTab
        .id = RibTab_Adr_Dokum
        .Visible = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Verordner Hinzufügen")
    With CmCon
        .IconId = IC32_Doctor_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Verordner Kopieren")
    With CmCon
        .IconId = IC32_Doctor_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Verordner Entfernen")
    With CmCon
        .IconId = IC32_Doctor_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Verordner Speichern")
    With CmCon
        .IconId = IC32_Disk_Doctor
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Verordner Suchen")
    With CmCon
        .IconId = IC32_Doctor_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    
Case RibTab_Mitarbeit: '------------------------------------------------------------

    Set RbTab = RbBar.InsertTab(RibTab_Adr_Haupt, "Stammdaten")
    With RbTab
        .id = RibTab_Adr_Haupt
        .Visible = True
        .Selected = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mitarbeiter Hinzufügen")
    With CmCon
        .IconId = IC32_Woman_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mitarbeiter Kopieren")
    With CmCon
        .IconId = IC32_Woman_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mitarbeiter Entfernen")
    With CmCon
        .IconId = IC32_Woman_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mitarbeiter Speichern")
    With CmCon
        .IconId = IC32_Disk_Woman
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mitarbeiter Suchen")
    With CmCon
        .IconId = IC32_Woman_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Dokum, "Arbeitszeiten")
    With RbTab
        .id = RibTab_Adr_Dokum
        .Visible = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mitarbeiter Hinzufügen")
    With CmCon
        .IconId = IC32_Woman_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mitarbeiter Kopieren")
    With CmCon
        .IconId = IC32_Woman_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mitarbeiter Entfernen")
    With CmCon
        .IconId = IC32_Woman_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mitarbeiter Speichern")
    With CmCon
        .IconId = IC32_Disk_Woman
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    '---
    Set RbGrp = RbGps.AddGroup("Arbeitszeiten", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Sprechzeit_Add, "Arbeitszeiten Hinzufügen")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Doc_Add
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Sprechzeit_Save, "Arbeitszeiten Speichern")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Document_Disk
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Sprechzeit_Del, "Arbeitszeiten Loschen")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Doc_Del
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With
    '---
    Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt2, "Verwendung :")
    With CmCon
        .ToolTipText = vbNullString
        .flags = xtpFlagRightAlign
        .BeginGroup = True
    End With
    Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt1, "Gültigkeit ab :")
    With CmCon
        .ToolTipText = vbNullString
        .flags = xtpFlagRightAlign
    End With
    Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt3, "Arbeitszeiten :")
    With CmCon
        .ToolTipText = "Zeigt den aktuelklen Saldo"
        .flags = xtpFlagRightAlign
    End With

    Set CmTyp = RbGrp.Add(xtpControlComboBox, AD_Sprechzeit_Typen, vbNullString)
    With CmTyp
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .DropDownItemCount = 2
        .EditHint = "Sprechzeitentyp"
        .ToolTipText = "Wechselt zwischen starren und flexiblen Arbeitszeiten"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 90
        .AddItem "starre Zeiten", 1
        .AddItem "flexible Zeiten", 2
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .ListIndex = 1
        Else
            .ListIndex = 2
        End If
    End With
    Set CmDat = RbGrp.Add(xtpControlComboBox, AD_Sprechzeit_Datum, vbNullString)
    With CmDat
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .DropDownItemCount = 10
        .EditHint = "Startdatum..."
        .ToolTipText = "Ab wann sollen die neuen Sprechzeiten gültig sein?"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 90
        AktWo = 1
        Do
        TmDat = SKaW2(AktWo, vbMonday, Year(Date))
        If TmDat >= Date Then
            .AddItem TmDat
        End If
        AktWo = AktWo + 1
        Loop Until Year(TmDat) > (Year(Date) + 2)
    End With
    Set CmCom = RbGrp.Add(xtpControlComboBox, AD_Sprechzeit_Auswa, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .DropDownItemCount = 5
        .EditHint = "Arbeitszeiten..."
        .ToolTipText = "Bitte wählen Sie einen gespeicherten Arbeitszeiteneintrag"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 90
        If GlSpT = False Then 'Starre oder flexible Sprechzeiten verwenden
            .Enabled = False
        End If
    End With

    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Booki, "Buchungszeiten")
    With RbTab
        .id = RibTab_Adr_Booki
        .Visible = GlOTS
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mitarbeiter Hinzufügen")
    With CmCon
        .IconId = IC32_Woman_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mitarbeiter Kopieren")
    With CmCon
        .IconId = IC32_Woman_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mitarbeiter Entfernen")
    With CmCon
        .IconId = IC32_Woman_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mitarbeiter Speichern")
    With CmCon
        .IconId = IC32_Disk_Woman
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mitarbeiter Suchen")
    With CmCon
        .IconId = IC32_Woman_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
Case Else: '------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Adr_Haupt, "Stammdaten")
    With RbTab
        .id = RibTab_Adr_Haupt
        .Visible = True
        .Selected = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mandant Suchen")
    With CmCon
        .IconId = IC32_BusMan_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Dokum, "Sprechzeiten")
    With RbTab
        .id = RibTab_Adr_Dokum
        .Visible = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mandant Suchen")
    With CmCon
        .IconId = IC32_BusMan_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Dokum, "Vorgaben")
    With RbTab
        .id = RibTab_Adr_Eigen
        .Visible = True
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mandant Suchen")
    With CmCon
        .IconId = IC32_BusMan_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
    Set RbTab = RbBar.InsertTab(RibTab_Adr_Booki, "Buchungszeiten")
    With RbTab
        .id = RibTab_Adr_Booki
        .Visible = GlOTS
    End With
    Set RbGps = RbTab.Groups
    Set RbGps = RbTab.Groups
    Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Adr_Bearbeit)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Add, "Mandant Hinzufügen")
    With CmCon
        .IconId = IC32_BusMan_Add
        .ShortcutText = "F3"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Copy, "Mandant Kopieren")
    With CmCon
        .IconId = IC32_BusMan_Copy
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patient_Del, "Mandant Entfernen")
    With CmCon
        .IconId = IC32_BusMan_Del
        .ShortcutText = "Entf"
        .Width = GlRib
    End With
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Save, "Mandant Speichern")
    With CmCon
        .IconId = IC32_Disk_Scientist
        .ShortcutText = "F8"
        .Width = GlRib
    End With
    Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Adr_Suchen)
    Set CmCon = RbGrp.Add(xtpControlButton, AD_Patienten_Suchen, "Mandant Suchen")
    With CmCon
        .IconId = IC32_BusMan_View
        .ShortcutText = "F5"
        .Width = GlRib
    End With
    '---
End Select

Set CmCoS = RbBar.Controls
For Each CmCon In CmCoS
    CmCon.ToolTipText = IniGetOpt(KeyNa, CmCon.id)
Next CmCon

CmAcs(AD_Patient_Copy).Enabled = False
CmAcs(AD_Patient_Del).Enabled = False
If GlMId = -2 Then
    CmAcs(AD_Patient_Add).Enabled = False
End If

'---

With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
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
    .SetIconSize True, 32, 32
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
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs
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
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
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
    .FontHeight = GlToF
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
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
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

Set CmPan = Nothing
Set CmSta = Nothing
Set CmPop = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set MsBar = Nothing
Set RbBar = Nothing
Set RbTab = Nothing
Set RbGrp = Nothing
Set RbGps = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

CrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MMenu " & Err.Number
Resume Next

End Sub
Public Sub MNeu(Optional ByVal DaNeu As Boolean = False)
On Error GoTo NeErr
'Bereitet die Neueingabe einer Adresse vor

Dim RetWe As Long
Dim IdxNr As Long
Dim FoCol As Long
Dim AdPIN As Long
Dim TagWe As String
Dim Frage As Integer
Dim LauZa As Integer
Dim AktZa As Integer
Dim IdErB As Integer
Dim IdErk As Integer
Dim IdStK As Integer
Dim LiIdx As Integer
Dim LiLin As Boolean
Dim Mld1, Tit1 As String

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RpCo5 As XtremeReportControl.ReportControl

Dim TxDum As XtremeSuiteControls.FlatEdit
Dim TxRch As XtremeSuiteControls.FlatEdit
Dim TxZe1 As XtremeSuiteControls.FlatEdit
Dim TxZe2 As XtremeSuiteControls.FlatEdit
Dim cmBuLa As XtremeSuiteControls.ComboBox
Dim cmKVBz As XtremeSuiteControls.ComboBox
Dim cmKant As XtremeSuiteControls.ComboBox
Dim cmRas1 As XtremeSuiteControls.ComboBox
Dim cmRas2 As XtremeSuiteControls.ComboBox
Dim cmMaxT As XtremeSuiteControls.ComboBox
Dim cmMaxP As XtremeSuiteControls.ComboBox
Dim cmVorl As XtremeSuiteControls.ComboBox
Dim cmBuRa As XtremeSuiteControls.ComboBox
Dim cmNoti As XtremeSuiteControls.ComboBox
Dim cmKata As XtremeSuiteControls.ComboBox
Dim cmKett As XtremeSuiteControls.ComboBox
Dim cmKet2 As XtremeSuiteControls.ComboBox
Dim cmRahm As XtremeSuiteControls.ComboBox
Dim cmKont As XtremeSuiteControls.ComboBox
Dim cmKon2 As XtremeSuiteControls.ComboBox
Dim cmGeK1 As XtremeSuiteControls.ComboBox
Dim cmGeK2 As XtremeSuiteControls.ComboBox
Dim cmStKt As XtremeSuiteControls.ComboBox
Dim cmReTy As XtremeSuiteControls.ComboBox
Dim cmSteu As XtremeSuiteControls.ComboBox
Dim ChThe As XtremeSuiteControls.CheckBox
Dim ChOut As XtremeSuiteControls.CheckBox
Dim ChTeS As XtremeSuiteControls.CheckBox
Dim ChDef As XtremeSuiteControls.CheckBox
Dim ChOnT As XtremeSuiteControls.CheckBox

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set RpCo5 = FM.repCont5
Set ChThe = FM.chkOpti2
Set ChTeS = FM.chkKaAus
Set ChDef = FM.chkDefra
Set ChOnT = FM.chkOnlTe
Set TxZe1 = FM.txtZeit1
Set TxZe2 = FM.txtZeit2
Set TxBri = FM.txtS1F20
Set FeAn1 = FM.txtS1F02
Set FeLa1 = FM.txtS1F12
Set TxFir = FM.txtS1F01
Set TxOrt = FM.txtS1F09
Set TxErs = FM.txtS2F27
Set TxNum = FM.txtS1F30
Set TxDum = FM.txtDummy
Set TxRch = FM.txtS4F01
Set cmRas1 = FM.cmbRast1
Set cmRas2 = FM.cmbRast2
Set cmMaxT = FM.cmbMaxTe
Set cmMaxP = FM.cmbMaxPa
Set cmVorl = FM.cmbVorLa
Set cmBuRa = FM.cmbBuRad
Set cmNoti = FM.cmbNotVa
Set cmKata = FM.cmbGbKat
Set cmKett = FM.cmbGbKet
Set cmKet2 = FM.cmbGbKe2
Set cmRahm = FM.cmbKtoRa
Set cmKont = FM.cmbKtoEr
Set cmKon2 = FM.cmbKtoEk
Set cmGeK1 = FM.cmbGeKt1
Set cmGeK2 = FM.cmbGeKt2
Set cmStKt = FM.cmbKtoSt
Set cmReTy = FM.cmbReTyp
Set cmSteu = FM.cmbSteue
Set FePat = FM.cmbS2F10
Set CmFch = FM.cmbKatal
Set FeGes = FM.cmbS1F08
Set FeFam = FM.txtS2F26
Set cmBuLa = FM.cmbBuLnd
Set cmKVBz = FM.cmbAbrBz
Set cmKant = FM.cmbKanto

Tit1 = "Neue Adresse"
Mld1 = "Der Datensatz wurde noch nicht gespeichert. Möchten Sie wirklich einen neuen Mandanten anlegen?"

If GlAdL = False Then
    If DaNeu = True Then
        If GlAdS = True Then
            Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
            If Frage <> 6 Then
                Exit Sub
            End If
        End If
    End If
End If

For Each AktCo In FM.Controls
    If AktCo.Tag <> vbNullString Then
        Select Case TypeName(AktCo)
        Case "FlatEdit":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                AktCo.Text = vbNullString
                AktCo.Tag = 0 & TagWe
        Case "TextBox":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                AktCo.Text = vbNullString
                AktCo.Tag = 0 & TagWe
        Case "CheckBox":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                AktCo.Value = 0
                AktCo.Tag = 0 & TagWe
        Case "ComboBox":
                TagWe = Mid$(AktCo.Tag, 2, Len(AktCo.Tag) - 1)
                Select Case Left$(AktCo.Name, 3)
                Case "txt": AktCo.Text = vbNullString
                            AktCo.Tag = 0 & TagWe
                Case "cmb": AktCo.Tag = 1 & TagWe
                End Select
        End Select
    End If
Next AktCo

If GlMaV = True Then 'Mandanten vorhanden
    For AktZa = 1 To UBound(GlMaA)
        FePat.AddItem GlMaA(AktZa, 1)
        FePat.ItemData(AktZa - 1) = GlMaA(AktZa, 2)
    Next AktZa
End If

TxErs.Text = Date

cmRas1.ListIndex = GlZeR - 1 'Zeitrasterindex
cmRas2.ListIndex = GlZeR - 1 'Zeitrasterindex
cmMaxT.ListIndex = 0
cmMaxP.ListIndex = 1
cmVorl.ListIndex = 1
cmBuRa.ListIndex = 23
cmNoti.ListIndex = 24

If DaNeu = True Then
    IdErk = SCmb(cmKon2, GlSE1) 'Standarderlöskonto Kasse
    IdErB = SCmb(cmKont, GlSE2) 'Standarderlöskonto Bankkonto
    IdStK = SCmb(cmStKt, GlSKo) 'Standardsteuerkonto
    DoEvents
    AdPIN = Adr_Let()
    GlAdG = CreateID("M")
    GlAdN = True
    GlAdS = False
    TxBri.Text = 2
    TxDum.Text = 0
    TxRch.Text = GlStR 'Rechtestring
    TxZe1.Text = GlSZe 'Sprechzietenstring
    TxZe2.Text = GlSZe 'Sprechzietenstring
    TxNum.Text = Format$(AdPIN, "000000")
    TxNum.Tag = "1Mandant"
    CmFch.ListIndex = GlFri - 1
    FeLa1.Enabled = True
    FeLa1.Text = "Deutschland"
    cmBuLa.ListIndex = 0
    cmKVBz.ListIndex = 0
    cmKant.ListIndex = 0
    If GlStK - 1 <= cmKata.ListCount Then 'Standardgebührenkatalog
        cmKata.ListIndex = GlStK - 1
    Else
        cmKata.ListIndex = 0
    End If
    If cmKett.ListCount > 0 Then
        If (GlKe1 - 1) <= cmKett.ListCount Then
            cmKett.ListIndex = GlKe1 - 1
        Else
            cmKett.ListIndex = 0
        End If
    End If
    If cmKet2.ListCount > 0 Then
        If (GlKe2 - 1) <= cmKet2.ListCount Then
            cmKet2.ListIndex = GlKe2 - 1
        Else
            cmKet2.ListIndex = 0
        End If
    End If
    If cmRahm.ListCount > 0 Then
        If (GlKtR - 1) <= cmRahm.ListCount Then
            cmRahm.ListIndex = GlKtR - 1
        Else
            cmRahm.ListIndex = 0
        End If
    End If
    If cmReTy.ListCount > 0 Then
        Select Case GlReT 'Standardbelegtyp
        Case "R": LiIdx = 0
        Case "L": LiIdx = 1
        Case "A": LiIdx = 2
        Case "M": LiIdx = 3
        Case "G": LiIdx = 4
        Case "I": LiIdx = 5
        End Select
        If LiIdx <= cmReTy.ListCount Then
            cmReTy.ListIndex = LiIdx
        Else
            cmReTy.ListIndex = 0
        End If
    End If
    If cmSteu.ListCount > 0 Then
        If (GlStS - 1) <= cmSteu.ListCount Then
            cmSteu.ListIndex = GlStS - 1
        Else
            cmSteu.ListIndex = 0
        End If
    End If
    If cmGeK1.ListCount > 0 Then
        If (GlGkB - 1) < cmGeK1.ListCount Then
            cmGeK1.ListIndex = GlGkB - 1 'Standardgeldkonto Bankkonto
        Else
            cmGeK1.ListIndex = 0
        End If
    End If
    If cmGeK1.ListCount > 0 Then
        If (GlGkK - 1) < cmGeK1.ListCount Then
            cmGeK2.ListIndex = GlGkK - 1 'Standardgeldkonto Kasse
        Else
            cmGeK2.ListIndex = 0
        End If
    End If
    If IdErB >= 0 Then
        cmKont.ListIndex = IdErB
    Else
        cmKont.ListIndex = 0
    End If
    If IdErk >= 0 Then
        cmKon2.ListIndex = IdErk
    Else
        cmKon2.ListIndex = 0
    End If
    If IdStK >= 0 Then
        cmStKt.ListIndex = IdStK
    Else
        cmStKt.ListIndex = 0
    End If
    If GlMId >= 0 Then
        GlMId = -1
    End If
    ChThe.Enabled = False
    ChTeS.Enabled = False
    ChDef.Enabled = False
    If FePat.ListCount > 0 Then
        If GlSuP.SuMan > 0 Then
            IdxNr = Adr_Cm(FePat, GlSuP.SuMan)
            FePat.ListIndex = IdxNr
        Else
            FePat.ListIndex = GlMaA(GlSMa, 0) - 1
        End If
    End If
Else
    If GlMId = -2 Then
        GlAdN = True
        GlAdS = False
        AdPIN = Adr_Let()
        GlAdG = CreateID("M")
        TxDum.Text = GlMId
        TxRch.Text = GlStR 'Rechtestring
        TxZe1.Text = GlSZe 'Sprechzietenstring
        TxZe2.Text = GlSZe 'Sprechzietenstring
        TxNum.Text = Format$(AdPIN, "000000")
        TxNum.Tag = "1Mandant"
        ChThe.Enabled = False
        ChTeS.Enabled = False
        ChDef.Enabled = False
        CmFch.ListIndex = GlFri - 1
    End If
End If

Select Case GlBut
Case RibTab_Mandanten:
    If GlMaV = True Then 'Mandanten vorhanden
        ChTeS.Visible = False
    End If
    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        ChTeS.Enabled = False
        ChOnT.Enabled = False
        cmNoti.Enabled = False
    End If
Case RibTab_Mitarbeit:
    If GlMiV = True Then
        If UBound(GlMiA) <= 1 Then 'Aktive Mitarbeiter
            ChTeS.Enabled = False
            ChTeS.Value = xtpUnchecked
        End If
    End If
Case Else:
    ChTeS.Enabled = False
    ChOnT.Enabled = False
    cmNoti.Enabled = False
End Select

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing
Set RpCo5 = Nothing

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MNeu " & Err.Number
Resume Next

End Sub
Public Sub MOpn()
On Error GoTo SaErr

Dim Recht As String
Dim FaIdx As Integer
Dim AktZa As Integer

Dim Labl01 As XtremeSuiteControls.Label
Dim Labl13 As XtremeSuiteControls.Label
Dim Labl14 As XtremeSuiteControls.Label
Dim Labl15 As XtremeSuiteControls.Label
Dim Labl16 As XtremeSuiteControls.Label
Dim Labl17 As XtremeSuiteControls.Label
Dim Labl18 As XtremeSuiteControls.Label
Dim Labl19 As XtremeSuiteControls.Label
Dim Labl20 As XtremeSuiteControls.Label
Dim Labl21 As XtremeSuiteControls.Label
Dim Labl22 As XtremeSuiteControls.Label
Dim Labl23 As XtremeSuiteControls.Label
Dim Labl24 As XtremeSuiteControls.Label
Dim Labl75 As XtremeSuiteControls.Label
Dim Labl76 As XtremeSuiteControls.Label
Dim txIKNr As XtremeSuiteControls.FlatEdit
Dim txGLNr As XtremeSuiteControls.FlatEdit
Dim txZSRn As XtremeSuiteControls.FlatEdit
Dim txPasw As XtremeSuiteControls.FlatEdit
Dim txPoFa As XtremeSuiteControls.FlatEdit
Dim txAbte As XtremeSuiteControls.FlatEdit
Dim txGlID As XtremeSuiteControls.FlatEdit
Dim txTel1 As XtremeSuiteControls.FlatEdit
Dim txTel6 As XtremeSuiteControls.FlatEdit
Dim txRech As XtremeSuiteControls.FlatEdit
Dim txBLZn As XtremeSuiteControls.FlatEdit
Dim txKont As XtremeSuiteControls.FlatEdit
Dim txGLNn As XtremeSuiteControls.FlatEdit
Dim txZSRr As XtremeSuiteControls.FlatEdit
Dim cmFach As XtremeSuiteControls.ComboBox
Dim cmLand As XtremeSuiteControls.ComboBox
Dim cmBuLa As XtremeSuiteControls.ComboBox
Dim cmKVBz As XtremeSuiteControls.ComboBox
Dim cmKant As XtremeSuiteControls.ComboBox
Dim cmFami As XtremeSuiteControls.ComboBox
Dim cmGsTy As XtremeSuiteControls.ComboBox
Dim cmMand As XtremeSuiteControls.ComboBox
Dim PuBu1 As XtremeSuiteControls.PushButton
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmMandant
Set txIKNr = FM.txtIKNum
Set txGLNr = FM.txtGLNum
Set txZSRn = FM.txtZSRnr
Set txPasw = FM.txtS1F37
Set txPoFa = FM.txtS1F23
Set txGlID = FM.txtGIDNr
Set txAbte = FM.txtS1F22
Set txTel1 = FM.txtS1F15
Set txTel6 = FM.txtS2F23
Set txRech = FM.txtS4F01
Set txBLZn = FM.txtS2F04
Set txKont = FM.txtS2F05
Set txGLNn = FM.txtS1F38
Set txZSRr = FM.txtS1F39
Set cmFami = FM.txtS2F26
Set cmFach = FM.cmbKatal
Set cmLand = FM.txtS1F12
Set cmBuLa = FM.cmbBuLnd
Set cmKVBz = FM.cmbAbrBz
Set cmKant = FM.cmbKanto
Set cmGsTy = FM.cmbS1F08
Set cmMand = FM.cmbS2F10
Set Labl01 = FM.lblLab01
Set Labl13 = FM.lblLab13
Set Labl14 = FM.lblLab14
Set Labl15 = FM.lblLab15
Set Labl16 = FM.lblLab16
Set Labl17 = FM.lblLab17
Set Labl18 = FM.lblLab18
Set Labl19 = FM.lblLab19
Set Labl20 = FM.lblLab20
Set Labl21 = FM.lblLab21
Set Labl22 = FM.lblLab22
Set Labl23 = FM.lblLab23
Set Labl24 = FM.lblLab24
Set Labl75 = FM.lblLab75
Set Labl76 = FM.lblLab76
Set Rahm3 = FM.frmRahm3
Set PuBu1 = FM.btnSign1
Set RpCo5 = FM.repCont5

FaIdx = cmFach.ListIndex

Recht = txRech.Text

If Recht = vbNullString Then
    Recht = GlStR 'Rechtestring
End If

If IsNumeric(Recht) = False Then
    Recht = GlStR 'Rechtestring
End If

If Len(Recht) <> GlZaR Then  'Rechteanzahl
    Recht = GlStR 'Rechtestring
End If

Select Case GlBut
Case RibTab_Mandanten:
    Labl01.Caption = "Folgende Angaben zum Mandanten werden benötigt, um diese auf den Rechnungs- und Mahnungsformularen darzustellen. Sie haben die Möglichkeit, mehrere Mandanten und damit auch gleichzeitig unterschiedliche Briefköpfe anzulegen. Dieser Mandant kann dann z.B. einer Rechnung zugeordnet werden.  Diese Angaben können jederzeit geändert oder ergänzt werden."
    RpCo5.Visible = False
    Labl17.Caption = "Grundvorgabe :"
    Labl19.Caption = "Prax / Firma :"
    Labl20.Caption = "PVS-Nr.:"
    Labl21.Visible = False
    Labl22.Visible = True
    Labl23.Visible = True
    Labl24.Visible = False
    txPoFa.Visible = True
    txGlID.Visible = True
    txAbte.Visible = True
    txTel1.Visible = False
    txTel6.Visible = False
    cmFach.Visible = True
    cmLand.Visible = True
    cmGsTy.Visible = False
    cmMand.Visible = False
    txGLNn.Visible = False
    txZSRr.Visible = False
    PuBu1.Visible = False
    Select Case FaIdx
    Case 0: 'Arzt
        cmLand.Text = "Deutschland"
        txIKNr.Visible = True
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = True
        cmKVBz.Visible = True
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "Bundesland :"
        Labl14.Caption = "LANR :"
        Labl15.Caption = "KV Bezirk :"
        Labl75.Visible = True
        Labl76.Visible = True
        txBLZn.Visible = True
        txKont.Visible = True
    Case 1: 'Heilpraktiker (GebüH)
        cmLand.Text = "Deutschland"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl75.Visible = True
        Labl76.Visible = True
        txBLZn.Visible = True
        txKont.Visible = True
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
    Case 2: 'Zahnarzt
        cmLand.Text = "Deutschland"
        txIKNr.Visible = True
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = True
        cmKVBz.Visible = True
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "Bundesland :"
        Labl14.Caption = "LANR :"
        Labl15.Caption = "KV Bezirk :"
        Labl75.Visible = True
        Labl76.Visible = True
        txBLZn.Visible = True
        txKont.Visible = True
    Case 4: 'Naturheilpraktiker (Tarif 590)
        cmLand.Text = "Schweiz"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = True
        txGLNr.Visible = True
        txZSRn.Visible = True
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "GLN :"
        Labl14.Caption = "ZSR :"
        Labl15.Caption = "Kanton :"
        Labl75.Visible = False
        Labl76.Visible = False
        txBLZn.Visible = False
        txKont.Visible = False
    Case 6: 'Wahlarzt (AT)
        cmLand.Text = "Österreich"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
        Labl75.Visible = False
        Labl76.Visible = False
        txBLZn.Visible = False
        txKont.Visible = False
    Case Else: 'Sonstige
        cmLand.Text = "Deutschland"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
        Labl75.Visible = True
        Labl76.Visible = True
        txBLZn.Visible = True
        txKont.Visible = True
    End Select
Case RibTab_Mitarbeit:
    Labl01.Caption = "Folgende Angaben zum / zur Mitarbeiter(in) werden benötigt, um bestimmte Einträge und Vorgänge zu dokumentieren. Auch dann, wenn keine weiteren Mitarbeiter(innen) vorhanden sind, ist es aus Gründen der passwortgeschützten Zugangskontrolle erforderlich, einen Mitarbeiter einzutragen und ein Passwort festzulegen. Diese Angaben können jederzeit geändert oder ergänzt werden."
    RpCo5.Visible = True
    Labl17.Caption = "Rechte :"
    Labl19.Caption = "Firma/Instit.:"
    Labl20.Caption = "Telefon :"
    Labl21.Visible = True
    Labl22.Visible = False
    Labl23.Visible = False
    Labl24.Visible = True
    txPoFa.Visible = False
    txGlID.Visible = False
    txAbte.Visible = False
    txTel1.Visible = True
    txTel6.Visible = True
    cmFach.Visible = False
    cmGsTy.Visible = True
    cmMand.Visible = True
    cmLand.Visible = False
    txIKNr.Visible = False
    txPasw.Visible = True
    cmFami.Visible = False
    cmBuLa.Visible = False
    cmKVBz.Visible = False
    cmKant.Visible = False
    txGLNr.Visible = False
    txZSRn.Visible = False
    Labl13.Visible = False
    Labl14.Visible = False
    Labl15.Visible = False
    Labl18.Visible = False
    PuBu1.Visible = True
    If GlFri = 5 Then 'Naturheilpraktiker (Tarif 590)
        Rahm3.Caption = "Therapeutendaten"
        Labl75.Caption = "GLN :"
        Labl76.Caption = "ZSR :"
        txBLZn.Visible = False
        txKont.Visible = False
    Else
        txGLNn.Visible = False
        txZSRr.Visible = False
    End If
Case RibTab_Verordner:
    Labl01.Caption = "Folgende Angaben zum Verordner werden benötigt, damit dieser dem Patienten schneller zugeordnet werden kann. Auf diese Weise können der eigenen Akte weitere medizinische Daten und Berichte des Patienten hinzugefügt um im Notfall der Verordner schneller kontaktiert werden. Diese Angaben können jederzeit geändert oder ergänzt werden."
    RpCo5.Visible = False
    Labl17.Caption = "Verordnertyp :"
    Labl19.Caption = "Prax / Firma :"
    Labl20.Caption = "PVS-Nr.:"
    Labl21.Visible = False
    Labl22.Visible = True
    Labl23.Visible = True
    Labl24.Visible = False
    txPoFa.Visible = True
    txGlID.Visible = True
    txAbte.Visible = True
    txTel1.Visible = False
    txTel6.Visible = False
    cmFach.Visible = True
    cmLand.Visible = True
    cmGsTy.Visible = False
    cmMand.Visible = False
    PuBu1.Visible = False
    Select Case FaIdx
    Case 0: 'Arzt
        cmLand.Text = "Deutschland"
        txIKNr.Visible = True
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = True
        cmKVBz.Visible = True
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "Bundesland :"
        Labl14.Caption = "LANR :"
        Labl15.Caption = "KV Bezirk :"
    Case 1: 'Heilpraktiker (GebüH)
        cmLand.Text = "Deutschland"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
    Case 2: 'Zahnarzt
        cmLand.Text = "Deutschland"
        txIKNr.Visible = True
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = True
        cmKVBz.Visible = True
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "Bundesland :"
        Labl14.Caption = "LANR :"
        Labl15.Caption = "KV Bezirk :"
    Case 4: 'Naturheilpraktiker (Tarif 590)
        cmLand.Text = "Schweiz"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = True
        txGLNr.Visible = True
        txZSRn.Visible = True
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "GLN :"
        Labl14.Caption = "ZSR :"
        Labl15.Caption = "Kanton :"
    Case 6: 'Wahlarzt (AT)
        cmLand.Text = "Österreich"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
    Case Else: 'Sonstige
        cmLand.Text = "Deutschland"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
    End Select
Case RibTab_Startseite:
    Labl01.Caption = "Folgende Angaben zum Mandanten werden benötigt, um diese auf den Rechnungs- und Mahnungsformularen darzustellen. Sie haben die Möglichkeit, mehrere Mandanten und damit auch gleichzeitig unterschiedliche Briefköpfe anzulegen. Dieser Mandant kann dann z.B. einer Rechnung zugeordnet werden.  Diese Angaben können jederzeit geändert oder ergänzt werden."
    RpCo5.Visible = False
    Labl17.Caption = "Grundvorgabe :"
    Labl19.Caption = "Prax / Firma :"
    Labl20.Caption = "PVS-Nr.:"
    Labl21.Visible = False
    Labl22.Visible = True
    Labl23.Visible = True
    Labl24.Visible = False
    txPoFa.Visible = True
    txGlID.Visible = True
    txAbte.Visible = True
    txTel1.Visible = False
    txTel6.Visible = False
    cmFach.Visible = True
    cmLand.Visible = True
    cmGsTy.Visible = False
    cmMand.Visible = False
    PuBu1.Visible = False
    Select Case FaIdx
    Case 0: 'Arzt
        cmLand.Text = "Deutschland"
        txIKNr.Visible = True
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = True
        cmKVBz.Visible = True
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "Bundesland :"
        Labl14.Caption = "LANR :"
        Labl15.Caption = "KV Bezirk :"
    Case 1: 'Heilpraktiker (GebüH)
        cmLand.Text = "Deutschland"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
    Case 2: 'Zahnarzt
        cmLand.Text = "Deutschland"
        txIKNr.Visible = True
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = True
        cmKVBz.Visible = True
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "Bundesland :"
        Labl14.Caption = "LANR :"
        Labl15.Caption = "KV Bezirk :"
    Case 4: 'Naturheilpraktiker (Tarif 590)
        cmLand.Text = "Schweiz"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = True
        txGLNr.Visible = True
        txZSRn.Visible = True
        Labl13.Visible = True
        Labl14.Visible = True
        Labl15.Visible = True
        Labl16.Visible = False
        Labl13.Caption = "GLN :"
        Labl14.Caption = "ZSR :"
        Labl15.Caption = "Kanton :"
    Case Else: 'Sonstige
        cmLand.Text = "Deutschland"
        txIKNr.Visible = False
        txPasw.Visible = False
        cmFami.Visible = False
        cmBuLa.Visible = False
        cmKVBz.Visible = False
        cmKant.Visible = False
        txGLNr.Visible = False
        txZSRn.Visible = False
        Labl13.Visible = False
        Labl14.Visible = False
        Labl15.Visible = False
        Labl16.Visible = False
        Labl13.Caption = vbNullString
        Labl14.Caption = vbNullString
        Labl15.Caption = vbNullString
    End Select
End Select

For AktZa = 0 To GlZaR - 1 'Rechteanzahl
    If Mid$(Recht, AktZa + 1, 1) = "1" Then
        RpRcs(AktZa).Item(0).Checked = True
    Else
        RpRcs(AktZa).Item(0).Checked = False
    End If
Next AktZa

Set RpCo5 = Nothing

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MOpn " & Err.Number
Resume Next

End Sub
Public Sub MPosi()
On Error GoTo ReErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim TbHoh As Long
Dim ShLa1 As VB.Shape
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set Rahm8 = FM.frmRahm8
Set TxTe4 = FM.txtS3F02
Set S2L01 = FM.lblLab01
Set ShLa1 = FM.shpLabl1
Set RpCo5 = FM.repCont5

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ShLa1.Move 0, ClObn, ClBre, 1000
    S2L01.Move 400, ClObn + 200, ClBre - 800, 800
    Rahm1.Move 200, ClObn + 1000, 5500, 3560
    Rahm2.Move 200, Rahm1.Top + Rahm1.Height + 100, 5500, 3100
    Rahm3.Move 5800, Rahm1.Top + Rahm1.Height + 100, 5500, 3100
    Rahm4.Move 5800, ClObn + 1000, 5500, 3560
    Rahm5.Move 5800, Rahm1.Top + Rahm1.Height + 100, 5500, 3100
    Rahm6.Move 200, Rahm1.Top + Rahm1.Height + 100, 5500, 3100
    Rahm7.Move 200, ClObn + 1000, 11100, 3560
    Rahm8.Move 200, ClObn + 1000, 11100, 6720
    TxTe4.Move 1800, 3120, 8700, 3400
End If

Set CmBrs = Nothing

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MPosi " & Err.Number
Resume Next

End Sub

Public Sub MRast(Optional ByVal SpStr As String)
On Error GoTo NeErr
'Einstellen des Zeitrasters

Dim TabId As Long
Dim TmStr As String
Dim ZeRas As Integer
Dim AktZa As Integer
Dim AkZei As Integer
Dim CtlZa As Integer
Dim ZeVor As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab

Dim ChS01 As XtremeSuiteControls.CheckBox
Dim ChS02 As XtremeSuiteControls.CheckBox
Dim ChS03 As XtremeSuiteControls.CheckBox
Dim ChS04 As XtremeSuiteControls.CheckBox
Dim ChS05 As XtremeSuiteControls.CheckBox
Dim ChS06 As XtremeSuiteControls.CheckBox
Dim ChS07 As XtremeSuiteControls.CheckBox
Dim ChS08 As XtremeSuiteControls.CheckBox
Dim ChS09 As XtremeSuiteControls.CheckBox
Dim ChS10 As XtremeSuiteControls.CheckBox
Dim ChS11 As XtremeSuiteControls.CheckBox
Dim ChS12 As XtremeSuiteControls.CheckBox
Dim ChS13 As XtremeSuiteControls.CheckBox
Dim ChS14 As XtremeSuiteControls.CheckBox

Dim cmbS01 As XtremeSuiteControls.ComboBox
Dim cmbS02 As XtremeSuiteControls.ComboBox
Dim cmbS03 As XtremeSuiteControls.ComboBox
Dim cmbS04 As XtremeSuiteControls.ComboBox
Dim cmbS05 As XtremeSuiteControls.ComboBox
Dim cmbS06 As XtremeSuiteControls.ComboBox
Dim cmbS07 As XtremeSuiteControls.ComboBox
Dim cmbS08 As XtremeSuiteControls.ComboBox
Dim cmbS09 As XtremeSuiteControls.ComboBox
Dim cmbS10 As XtremeSuiteControls.ComboBox
Dim cmbS11 As XtremeSuiteControls.ComboBox
Dim cmbS12 As XtremeSuiteControls.ComboBox
Dim cmbS13 As XtremeSuiteControls.ComboBox
Dim cmbS14 As XtremeSuiteControls.ComboBox
Dim cmbS15 As XtremeSuiteControls.ComboBox
Dim cmbS16 As XtremeSuiteControls.ComboBox
Dim cmbS17 As XtremeSuiteControls.ComboBox
Dim cmbS18 As XtremeSuiteControls.ComboBox
Dim cmbS19 As XtremeSuiteControls.ComboBox
Dim cmbS20 As XtremeSuiteControls.ComboBox
Dim cmbS21 As XtremeSuiteControls.ComboBox
Dim cmbS22 As XtremeSuiteControls.ComboBox
Dim cmbS23 As XtremeSuiteControls.ComboBox
Dim cmbS24 As XtremeSuiteControls.ComboBox
Dim cmbS25 As XtremeSuiteControls.ComboBox
Dim cmbS26 As XtremeSuiteControls.ComboBox
Dim cmbS27 As XtremeSuiteControls.ComboBox
Dim cmbS28 As XtremeSuiteControls.ComboBox

Dim AktCo As VB.Control
Dim TxZei1 As XtremeSuiteControls.FlatEdit
Dim TxZei2 As XtremeSuiteControls.FlatEdit
Dim cmRas1 As XtremeSuiteControls.ComboBox
Dim cmRas2 As XtremeSuiteControls.ComboBox
Dim cmMaxT As XtremeSuiteControls.ComboBox
Dim cmMaxP As XtremeSuiteControls.ComboBox
Dim cmVorl As XtremeSuiteControls.ComboBox
Dim cmBuRa As XtremeSuiteControls.ComboBox

Set FM = frmMandant
Set CmBrs = FM.comBar01
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Set TxZei1 = FM.txtZeit1
Set TxZei2 = FM.txtZeit2
Set cmRas1 = FM.cmbRast1
Set cmRas2 = FM.cmbRast2
Set cmMaxT = FM.cmbMaxTe
Set cmMaxP = FM.cmbMaxPa
Set cmVorl = FM.cmbVorLa
Set cmBuRa = FM.cmbBuRad

Set ChS01 = FM.chkBox01
Set ChS02 = FM.chkBox02
Set ChS03 = FM.chkBox03
Set ChS04 = FM.chkBox04
Set ChS05 = FM.chkBox05
Set ChS06 = FM.chkBox06
Set ChS07 = FM.chkBox07
Set ChS08 = FM.chkBox08
Set ChS09 = FM.chkBox09
Set ChS10 = FM.chkBox10
Set ChS11 = FM.chkBox11
Set ChS12 = FM.chkBox12
Set ChS13 = FM.chkBox13
Set ChS14 = FM.chkBox14

Set cmbS01 = FM.cmbSpZ01
Set cmbS02 = FM.cmbSpZ02
Set cmbS03 = FM.cmbSpZ03
Set cmbS04 = FM.cmbSpZ04
Set cmbS05 = FM.cmbSpZ05
Set cmbS06 = FM.cmbSpZ06
Set cmbS07 = FM.cmbSpZ07
Set cmbS08 = FM.cmbSpZ08
Set cmbS09 = FM.cmbSpZ09
Set cmbS10 = FM.cmbSpZ10
Set cmbS11 = FM.cmbSpZ11
Set cmbS12 = FM.cmbSpZ12
Set cmbS13 = FM.cmbSpZ13
Set cmbS14 = FM.cmbSpZ14
Set cmbS15 = FM.cmbSpZ15
Set cmbS16 = FM.cmbSpZ16
Set cmbS17 = FM.cmbSpZ17
Set cmbS18 = FM.cmbSpZ18
Set cmbS19 = FM.cmbSpZ19
Set cmbS20 = FM.cmbSpZ20
Set cmbS21 = FM.cmbSpZ21
Set cmbS22 = FM.cmbSpZ22
Set cmbS23 = FM.cmbSpZ23
Set cmbS24 = FM.cmbSpZ24
Set cmbS25 = FM.cmbSpZ25
Set cmbS26 = FM.cmbSpZ26
Set cmbS27 = FM.cmbSpZ27
Set cmbS28 = FM.cmbSpZ28

TabId = RbTab.id

Select Case TabId
Case RibTab_Adr_Dokum: 'Sprechzeiten
    If cmRas1.Text <> vbNullString Then
        ZeRas = cmRas1.ItemData(cmRas1.ListIndex)
    Else
        ZeRas = GlZeR 'Zeitrasterindex
    End If
Case RibTab_Adr_Booki: 'Buchungszeiten
    If cmRas2.Text <> vbNullString Then
        ZeRas = cmRas2.ItemData(cmRas2.ListIndex)
    Else
        ZeRas = GlZeR 'Zeitrasterindex
    End If
End Select

If SpStr <> vbNullString Then
    If Len(SpStr) > 50 Then
        TmStr = SpStr
    Else
        TmStr = GlSZe 'Sprechzietenstring
    End If
Else
    Select Case TabId
    Case RibTab_Adr_Dokum: 'Sprechzeiten
        If TxZei1.Text <> vbNullString Then
            If Len(TxZei1.Text) > 50 Then
                TmStr = TxZei1.Text
            Else
                TmStr = GlSZe 'Sprechzietenstring
            End If
        Else
            TmStr = GlSZe 'Sprechzietenstring
        End If
    Case RibTab_Adr_Booki: 'Buchungszeiten
        If TxZei2.Text <> vbNullString Then
            If Len(TxZei2.Text) > 50 Then
                TmStr = TxZei2.Text
            Else
                TmStr = GlSZe 'Sprechzietenstring
            End If
        Else
            TmStr = GlSZe 'Sprechzietenstring
        End If
    End Select
End If

SRast ZeRas

For Each AktCo In FM.Controls
    If Left$(AktCo.Name, 6) = "cmbSpZ" Then
        AktCo.Clear
        For AkZei = 1 To UBound(GlRas)
            With AktCo
                .AddItem GlRas(AkZei)
            End With
        Next AkZei
        AktCo.DropDownItemCount = 8
    End If
Next AktCo

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 2, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS01.Text = Mid$(TmStr, 2, 5)
Else
    cmbS01.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 8, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS02.Text = Mid$(TmStr, 8, 5)
Else
    cmbS02.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 14, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS03.Text = Mid$(TmStr, 14, 5)
Else
    cmbS03.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 20, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS04.Text = Mid$(TmStr, 20, 5)
Else
    cmbS04.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 26, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS05.Text = Mid$(TmStr, 26, 5)
Else
    cmbS05.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 32, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS06.Text = Mid$(TmStr, 32, 5)
Else
    cmbS06.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 38, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS07.Text = Mid$(TmStr, 38, 5)
Else
    cmbS07.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 44, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS08.Text = Mid$(TmStr, 44, 5)
Else
    cmbS08.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 50, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS09.Text = Mid$(TmStr, 50, 5)
Else
    cmbS09.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 56, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS10.Text = Mid$(TmStr, 56, 5)
Else
    cmbS10.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 62, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS11.Text = Mid$(TmStr, 62, 5)
Else
    cmbS11.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 68, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS12.Text = Mid$(TmStr, 68, 5)
Else
    cmbS12.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 74, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS13.Text = Mid$(TmStr, 74, 5)
Else
    cmbS13.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 80, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS14.Text = Mid$(TmStr, 80, 5)
Else
    cmbS14.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 86, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS15.Text = Mid$(TmStr, 86, 5)
Else
    cmbS15.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 92, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS16.Text = Mid$(TmStr, 92, 5)
Else
    cmbS16.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 98, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS17.Text = Mid$(TmStr, 98, 5)
Else
    cmbS17.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 104, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS18.Text = Mid$(TmStr, 104, 5)
Else
    cmbS18.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 110, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS19.Text = Mid$(TmStr, 110, 5)
Else
    cmbS19.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 116, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS20.Text = Mid$(TmStr, 116, 5)
Else
    cmbS20.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 122, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS21.Text = Mid$(TmStr, 122, 5)
Else
    cmbS21.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 128, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS22.Text = Mid$(TmStr, 128, 5)
Else
    cmbS22.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 134, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS23.Text = Mid$(TmStr, 134, 5)
Else
    cmbS23.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 140, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS24.Text = Mid$(TmStr, 140, 5)
Else
    cmbS24.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 146, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS25.Text = Mid$(TmStr, 146, 5)
Else
    cmbS25.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 152, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS26.Text = Mid$(TmStr, 152, 5)
Else
    cmbS26.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 158, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS27.Text = Mid$(TmStr, 158, 5)
Else
    cmbS27.ListIndex = 1
End If

ZeVor = False
For AktZa = 1 To UBound(GlRas)
    If GlRas(AktZa) = Mid$(TmStr, 164, 5) Then
        ZeVor = True
        Exit For
    End If
Next AktZa
If ZeVor = True Then
    cmbS28.Text = Mid$(TmStr, 164, 5)
Else
    cmbS28.ListIndex = 1
End If

If Mid$(TmStr, 1, 1) = "A" Then
    ChS01.Value = xtpChecked
    cmbS01.Visible = True
    cmbS02.Visible = True
Else
    ChS01.Value = xtpUnchecked
    cmbS01.Enabled = False
    cmbS02.Enabled = False
    cmbS01.Visible = False
    cmbS02.Visible = False
End If

If Mid$(TmStr, 13, 1) = "A" Then
    ChS08.Value = xtpChecked
    cmbS03.Visible = True
    cmbS04.Visible = True
Else
    ChS08.Value = xtpUnchecked
    cmbS03.Enabled = False
    cmbS04.Enabled = False
    cmbS03.Visible = False
    cmbS04.Visible = False
End If

If Mid$(TmStr, 25, 1) = "A" Then
    ChS02.Value = xtpChecked
    cmbS05.Visible = True
    cmbS06.Visible = True
Else
    ChS02.Value = xtpUnchecked
    cmbS05.Enabled = False
    cmbS06.Enabled = False
    cmbS05.Visible = False
    cmbS06.Visible = False
End If

If Mid$(TmStr, 37, 1) = "A" Then
    ChS09.Value = xtpChecked
    cmbS07.Visible = True
    cmbS08.Visible = True
Else
    ChS09.Value = xtpUnchecked
    cmbS07.Enabled = False
    cmbS08.Enabled = False
    cmbS07.Visible = False
    cmbS08.Visible = False
End If

If Mid$(TmStr, 49, 1) = "A" Then
    ChS03.Value = xtpChecked
    cmbS09.Visible = True
    cmbS10.Visible = True
Else
    ChS03.Value = xtpUnchecked
    cmbS09.Enabled = False
    cmbS10.Enabled = False
    cmbS09.Visible = False
    cmbS10.Visible = False
End If

If Mid$(TmStr, 61, 1) = "A" Then
    ChS10.Value = xtpChecked
    cmbS11.Visible = True
    cmbS12.Visible = True
Else
    ChS10.Value = xtpUnchecked
    cmbS11.Enabled = False
    cmbS12.Enabled = False
    cmbS11.Visible = False
    cmbS12.Visible = False
End If

If Mid$(TmStr, 73, 1) = "A" Then
    ChS04.Value = xtpChecked
    cmbS13.Visible = True
    cmbS14.Visible = True
Else
    ChS04.Value = xtpUnchecked
    cmbS13.Enabled = False
    cmbS14.Enabled = False
    cmbS13.Visible = False
    cmbS14.Visible = False
End If

If Mid$(TmStr, 85, 1) = "A" Then
    ChS11.Value = xtpChecked
    cmbS15.Visible = True
    cmbS16.Visible = True
Else
    ChS11.Value = xtpUnchecked
    cmbS15.Enabled = False
    cmbS16.Enabled = False
    cmbS15.Visible = False
    cmbS16.Visible = False
End If

If Mid$(TmStr, 97, 1) = "A" Then
    ChS05.Value = xtpChecked
    cmbS17.Visible = True
    cmbS18.Visible = True
Else
    ChS05.Value = xtpUnchecked
    cmbS17.Enabled = False
    cmbS18.Enabled = False
    cmbS17.Visible = False
    cmbS18.Visible = False
End If

If Mid$(TmStr, 109, 1) = "A" Then
    ChS12.Value = xtpChecked
    cmbS19.Visible = True
    cmbS20.Visible = True
Else
    ChS12.Value = xtpUnchecked
    cmbS19.Enabled = False
    cmbS20.Enabled = False
    cmbS19.Visible = False
    cmbS20.Visible = False
End If

If Mid$(TmStr, 121, 1) = "A" Then
    ChS06.Value = xtpChecked
    cmbS21.Visible = True
    cmbS22.Visible = True
Else
    ChS06.Value = xtpUnchecked
    cmbS21.Enabled = False
    cmbS22.Enabled = False
    cmbS21.Visible = False
    cmbS22.Visible = False
End If

If Mid$(TmStr, 133, 1) = "A" Then
    ChS13.Value = xtpChecked
    cmbS23.Visible = True
    cmbS24.Visible = True
Else
    ChS13.Value = xtpUnchecked
    cmbS23.Enabled = False
    cmbS24.Enabled = False
    cmbS23.Visible = False
    cmbS24.Visible = False
End If

If Mid$(TmStr, 145, 1) = "A" Then
    ChS07.Value = xtpChecked
    cmbS25.Visible = True
    cmbS26.Visible = True
Else
    ChS07.Value = xtpUnchecked
    cmbS25.Enabled = False
    cmbS26.Enabled = False
    cmbS25.Visible = False
    cmbS26.Visible = False
End If

If Mid$(TmStr, 157, 1) = "A" Then
    ChS14.Value = xtpChecked
    cmbS27.Visible = True
    cmbS28.Visible = True
Else
    ChS14.Value = xtpUnchecked
    cmbS27.Enabled = False
    cmbS28.Enabled = False
    cmbS27.Visible = False
    cmbS28.Visible = False
End If

Set RbTab = Nothing
Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

NeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MRast " & Err.Number
Resume Next

End Sub
Private Sub MReg()
On Error GoTo ReErr
'Anlegen von Registryeinträgen

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If GlFnt = True Then
    xGro = 774
    yGro = 710
Else
    xGro = 974
    yGro = 810
End If

xPos = (GlxGr / 2) - (xGro / 2)
yPos = (GlyGr / 2) - (yGro / 2)

If IniGetSek(GlINI, "ManForm") = False Then IniSetSek "ManForm"
If IniGetVal("ManForm", "FenLin") = vbNullString Then IniSetVal "ManForm", "FenLin", xPos
If IniGetVal("ManForm", "FenObe") = vbNullString Then IniSetVal "ManForm", "FenObe", yPos
If IniGetVal("ManForm", "FenBre") = vbNullString Then IniSetVal "ManForm", "FenBre", xGro
If IniGetVal("ManForm", "FenHoh") = vbNullString Then IniSetVal "ManForm", "FenHoh", yGro

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "MReg " & Err.Number
Resume Next

End Sub
