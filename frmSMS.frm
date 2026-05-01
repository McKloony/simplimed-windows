VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmSMS 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "SMS Versand"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5505
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   6
      Top             =   4000
      Width           =   5600
      _Version        =   1048579
      _ExtentX        =   9878
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnClose 
         Height          =   400
         Left            =   3600
         TabIndex        =   9
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
      Begin XtremeSuiteControls.PushButton btnWeite 
         Height          =   400
         Left            =   2200
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
      Begin XtremeSuiteControls.PushButton btnGutha 
         Height          =   400
         Left            =   900
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Guthaben"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   80
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4000
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   4300
      _Version        =   1048579
      _ExtentX        =   7585
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ProgressBar prbStat1 
         Height          =   350
         Left            =   130
         TabIndex        =   5
         Top             =   3500
         Width           =   4100
         _Version        =   1048579
         _ExtentX        =   7232
         _ExtentY        =   617
         _StockProps     =   93
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelef 
         Height          =   350
         Left            =   130
         TabIndex        =   2
         Top             =   840
         Width           =   4100
         _Version        =   1048579
         _ExtentX        =   7232
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtNaTex 
         Height          =   2000
         Left            =   130
         TabIndex        =   3
         Top             =   1300
         Width           =   4100
         _Version        =   1048579
         _ExtentX        =   7232
         _ExtentY        =   3528
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         MaxLength       =   350
         MultiLine       =   -1  'True
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Die folgende Nachricht wird an den gewünschten Empfänger gesendet (max. 160 Zeichen)."
         Height          =   580
         Left            =   200
         TabIndex        =   4
         Top             =   100
         Width           =   4000
      End
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private TxTex As XtremeSuiteControls.FlatEdit
Private TxTel As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCls As XtremeReportControl.ReportColumns
Private RpRow As XtremeReportControl.ReportRow

Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Public mTeNu As Long
Public NaTex As String
Public NaNum As String
Public DoTyp As Integer

Private FoLad As Boolean
Private Sub FInit()
On Error GoTo InErr

Dim AkZa1 As Integer
Dim AkZa2 As Integer

Set FM = frmSMS
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set PrBr1 = FM.prbStat1
Set TxTex = FM.txtNaTex
Set TxTel = FM.txtTelef

With PrBr1
    Select Case GlSty
    Case 8:
        .Appearance = xtpAppearanceOffice2013
        .UseVisualStyle = False
    Case 7:
        .Appearance = xtpAppearanceOffice2013
        .UseVisualStyle = False
    Case Else:
        .Appearance = xtpAppearanceResource
        .UseVisualStyle = True
    End Select
    .Scrolling = xtpProgressBarStandard
End With

TxTel.Text = SRufn(NaNum)
TxTex.Text = NaTex

Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
FM.BackColor = GlBak

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit & Err.Number"
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo InErr

Dim DaSta As Date
Dim DaEnd As Date
Dim ZeSta As Date
Dim ZeEnd As Date
Dim RowNr As Long
Dim AnzTe As Long
Dim AktTe As Long
Dim PatNr As Long
Dim TerNr As Long
Dim TeMit As Long
Dim TeMan As Long
Dim TmGui As String
Dim SMSHt As String
Dim SMSTx As String
Dim SMSNu As String
Dim TelMo As String
Dim TeStr As String
Dim TeGui As String
Dim PaStr As String
Dim DaStL As String
Dim DaStK As String
Dim DaStN As String
Dim ZeStS As String
Dim ZeStE As String
Dim MiNam As String
Dim MeTex As String
Dim SenOk As Boolean
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMain
Set TxTex = Me.txtNaTex
Set TxTel = Me.txtTelef
Set PrBr1 = Me.prbStat1
Set PuBu1 = Me.btnGutha
Set PuBu2 = Me.btnWeite
Set PuBu3 = Me.btnClose
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

Screen.MousePointer = vbHourglass

If GlBut = RibTab_Ter_Listen Then
    PuBu2.Enabled = False
    PuBu3.Enabled = False
    DoEvents

    AnzTe = RpSel.Count
    PrBr1.Min = 0
    PrBr1.Max = AnzTe
    
    If GlMiA(GlSmI, 27) <> vbNullString Then
        MiNam = GlMiA(GlSmI, 27) 'Verkehrsname
    ElseIf GlMiA(GlSmI, 1) <> vbNullString Then
        MiNam = GlMiA(GlSmI, 1) 'Kurzbezeichnung
    Else
        MiNam = GlMiA(GlSmI, 4) & " " & GlMiA(GlSmI, 3) 'Vor- und Nachname
    End If

    For Each RpRow In RpSel
        If RpRow.GroupRow = False Then
            TerNr = 0
            PatNr = 0
            TeMan = 0
            TeMit = 0
            PaStr = vbNullString
            TelMo = vbNullString
            TeStr = vbNullString
            MeTex = vbNullString
            SMSNu = vbNullString

            RowNr = RpRow.Index
            Set RpCol = RpCls.Find(Ter_ID2)
            TerNr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Ter_ID0)
            PatNr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Ter_IDP)
            TeMan = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Ter_IDM)
            TeMit = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Ter_VonDat)
            DaSta = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Ter_BisDat)
            DaEnd = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Ter_ZeiVon)
            ZeSta = TimeValue(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Ter_ZeiBis)
            ZeEnd = TimeValue(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Ter_GuiID)
            TeGui = RpRow.Record(RpCol.ItemIndex).Value

            DaStL = Format$(DaSta, "dddd" & ", " & "dd" & ". " & "mmmm" & Chr$(32) & "yyyy")
            DaStK = Format$(DaSta, "ddd" & ", " & "dd" & ". " & "mmm" & Chr$(32) & "yyyy")
            DaStN = Format$(DaSta, "dd.mm.yyyy")
            ZeStS = Format$(ZeSta, "hh:mm")
            ZeStE = Format$(ZeEnd, "hh:mm")

            If PatNr > 0 Then
                S_AdDe PatNr 'Adressendetails
                With GlADt
                    PaStr = .AdKur
                    TelMo = .AdTe4
                End With

                If TelMo <> vbNullString Then
                    With GlTxV
                        If GlEmN(10, 1) <> vbNullString Then
                            .TxStr = GlEmN(10, 1)
                        Else
                            .TxStr = "Termin"
                        End If
                        If DoTyp < 6 Then
                            .DaStr = DaStL
                        ElseIf DoTyp < 10 Then
                            .DaStr = DaStK
                        Else
                            .DaStr = DaStN
                        End If
                        .MitNr = TeMit
                        .ManNr = TeMan
                        .PatNr = PatNr
                        .PaStr = PaStr
                        .ZeiSt = ZeStS
                        .ZeiEn = ZeStE
                        .TerID = TeGui
                    End With
                    TeStr = SEmTx()

                    TeStr = SUmw(TeStr, False, True)
                    TeStr = Replace(TeStr, vbCrLf, "$$", 1)
                    MiNam = SUmw(MiNam, True, True)
                    TeStr = TeStr & "$$$$" & MiNam
                    DoEvents

                    SMSNu = SRufn(TelMo) 'Formatiert die Rufnummer
                    SMSTx = TeStr
                    DoEvents

                    TxTel.Text = SMSNu
                    DoEvents
                    TxTex.Text = SMSTx
                    DoEvents

                    SenOk = SMSSn(SMSNu, TxTex)
                    If SenOk = True Then
                        TxTex.Text = TxTex.Text & vbCrLf & SMSNu & " OK"
                    Else
                        TxTex.Text = TxTex.Text & vbCrLf & SMSNu & " -"
                    End If
                    DoEvents
                                        
                    If TerNr > 0 Then
                        DBCmEx2 "qryTerOnTe", "@OnlTe", "@IdxNr", -1, TerNr
                    End If

                    If GlEKr = True Then 'Emails in Krankenblatt dokumentieren
                        TmGui = CreateID("M")
                        GlNeK = GlKoX
                        With GlNeK
                            .PatNr = PatNr
                            .IdxNr = 0
                            .EiDat = Format$(Date, "dd.mm.yyyy")
                            .EiZei = TimeValue(Now)
                            .EiTyp = 108 'Email
                            .KoStr = SMSTx
                            .KoGui = TmGui
                            .NeuEi = True
                            .Mitar = GlMiA(GlSmI, 2)
                        End With
                        K_Einf
                    End If
                End If
            End If

            DoEvents
            AktTe = AktTe + 1
            PrBr1.Value = AktTe
        End If
    Next RpRow

    DoEvents
    PuBu3.Enabled = True
    
    FoLad = False
    Screen.MousePointer = vbNormal

Else
    If mTeNu > 0 Then
        TerNr = mTeNu
    End If

    SenOk = SMSSn(TxTel.Text, TxTex.Text)
    DoEvents
    If SenOk = True Then
        SPopu "Nacht Versandt", "Die Nachricht wurde erfolgreich versandt.", IC48_Information
        PuBu2.Enabled = False
    Else
        SPopu "Nachrichtenversand", "Beim Versenden der Nachricht ist ein Fehler aufgetreten!", IC48_Forbidden
    End If
    
    If TerNr > 0 Then
        DBCmEx2 "qryTerOnTe", "@OnlTe", "@IdxNr", -1, TerNr
    End If
    
    FoLad = False
    Screen.MousePointer = vbNormal
    DoEvents
    
    Unload Me
End If

Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then
    MsgBox Err.Description, 48, "FWeit & Err.Number"
    SErLog Err.Description & " SMSSend " & Err.Number & " " & MeTex
End If
Resume Next

End Sub
Private Sub FKon()
On Error GoTo InErr

Dim TmStr As String

Set FM = frmSMS
Set PuBu1 = FM.btnGutha
Set PuBu2 = FM.btnWeite
Set PuBu3 = FM.btnClose

TmStr = SMSGu()

If TmStr <> vbNullString Then
    PuBu1.Caption = TmStr
    PuBu1.Enabled = False
    Me.txtDummy.SetFocus
End If

FoLad = False

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKon & Err.Number"
Resume Next

End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub
Private Sub btnGutha_Click()
    FKon
End Sub
Private Sub btnWeite_Click()
    FWeit
End Sub
Private Sub Form_Activate()
    FoLad = True
End Sub

Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FoLad = True
FInit
AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmSMS = Nothing
End Sub

Private Sub txtNaTex_GotFocus()
    Me.txtNaTex.SelStart = 0
    Me.txtNaTex.SelLength = Len(Me.txtNaTex.Text)
End Sub


Private Sub txtTelef_GotFocus()
    Me.txtTelef.SelStart = 0
    Me.txtTelef.SelLength = Len(Me.txtTelef.Text)
End Sub
