VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Exportieren"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   9
      Top             =   3900
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   4000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schlieﬂen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   2600
         TabIndex        =   11
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
         Left            =   1300
         TabIndex        =   10
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
      Height          =   3900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   6879
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkZuord 
         Height          =   240
         Left            =   1100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2600
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Archivierung"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbForma 
         Height          =   310
         Left            =   1000
         TabIndex        =   2
         Top             =   1140
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7038
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbEmail 
         Height          =   315
         Left            =   1000
         TabIndex        =   3
         Top             =   1940
         Width           =   4000
         _Version        =   1048579
         _ExtentX        =   7038
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Emailversand :"
         Height          =   210
         Left            =   1040
         TabIndex        =   15
         Top             =   1700
         Width           =   1500
      End
      Begin VB.Label lblLab05 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmExport.frx":0000
         Height          =   435
         Left            =   500
         TabIndex        =   13
         Top             =   100
         Width           =   5000
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Exportformat :"
         Height          =   210
         Left            =   1040
         TabIndex        =   14
         Top             =   900
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         BorderStyle     =   0  'Transparent
         Height          =   640
         Left            =   0
         Top             =   0
         Width           =   6000
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
      Top             =   6200
      Width           =   80
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3900
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   6879
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtPraxi 
         Height          =   350
         Left            =   900
         TabIndex        =   7
         Top             =   1940
         Width           =   4100
         _Version        =   1048579
         _ExtentX        =   7232
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   350
         Left            =   900
         TabIndex        =   6
         Top             =   1140
         Width           =   4100
         _Version        =   1048579
         _ExtentX        =   7232
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDocNa 
         Height          =   350
         Left            =   900
         TabIndex        =   8
         Top             =   2740
         Width           =   4100
         _Version        =   1048579
         _ExtentX        =   7232
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   500
         Left            =   500
         TabIndex        =   19
         Top             =   200
         Width           =   5600
         _Version        =   1048579
         _ExtentX        =   9878
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   $"frmExport.frx":0087
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   220
         Left            =   900
         TabIndex        =   18
         Top             =   900
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Emailadresse :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   220
         Left            =   900
         TabIndex        =   17
         Top             =   1700
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Praxisdarstellung :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab09 
         Height          =   220
         Left            =   900
         TabIndex        =   16
         Top             =   2500
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Dokumentenname:"
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private FTeEm As XtremeSuiteControls.FlatEdit
Private FTePr As XtremeSuiteControls.FlatEdit
Private FTeDo As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private CmFma As XtremeSuiteControls.ComboBox
Private CmEml As XtremeSuiteControls.ComboBox
Private ChZuo As XtremeSuiteControls.CheckBox
Private CoDia As XtremeSuiteControls.CommonDialog
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn
Private RpSel As XtremeReportControl.ReportSelectedRows
Private TxCoN As Tx4oleLib.TXTextControl

Private AltNa As String
Private TmGui As String

Public DaNam As String
Public OrgNa As String
Public ExMod As Integer
Public FoIdx As Integer
Public EmIdx As Integer
Public AnzSe As Integer
Public SeUpl As Integer

Private clFil As clsFile
Private Sub FCom()
On Error GoTo LdErr

Dim DatNa As String
Dim DaExt As String
Dim LiIdx As Integer
Dim Posit As Integer

Set CmFma = Me.cmbForma

LiIdx = CmFma.ListIndex

If DaNam <> vbNullString Then
    DatNa = DaNam
    DaExt = LCase(Right$(DatNa, 3))
    Posit = InStrRev(DatNa, ".", -1, 1)
    
    Select Case ExMod
    Case 9: 'Textverarbeitung
        If Posit > 0 Then
            Select Case LiIdx
            Case 0: If DaExt <> "txm" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "txm"
            Case 1: If DaExt <> "pdf" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "pdf"
            Case 2: If DaExt <> "doc" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "doc"
            Case 3: If DaExt <> "ocx" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "docx"
            Case 4: If DaExt <> "rtf" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "rtf"
            Case 5: If DaExt <> "htm" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "htm"
            Case 6: If DaExt <> "txt" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "txt"
            Case 7: If DaExt <> "xml" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "xml"
            End Select
        Else
            Select Case LiIdx
            Case 0: If DaExt <> "txm" Then DatNa = DatNa & ".txm"
            Case 1: If DaExt <> "pdf" Then DatNa = DatNa & ".pdf"
            Case 2: If DaExt <> "doc" Then DatNa = DatNa & ".doc"
            Case 3: If DaExt <> "ocx" Then DatNa = DatNa & ".docx"
            Case 4: If DaExt <> "rtf" Then DatNa = DatNa & ".rtf"
            Case 5: If DaExt <> "htm" Then DatNa = DatNa & ".htm"
            Case 6: If DaExt <> "txt" Then DatNa = DatNa & ".txt"
            Case 7: If DaExt <> "xml" Then DatNa = DatNa & ".xml"
            End Select
        End If
    Case 10: 'PDF-Viewer
        If Posit > 0 Then
            Select Case LiIdx
            Case 0: If DaExt <> "jpg" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "jpg"
            Case 1: If DaExt <> "png" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "png"
            Case 2: If DaExt <> "tif" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "tif"
            Case 3: If DaExt <> "bmp" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "bmp"
            Case 4: If DaExt <> "gif" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "gif"
            Case 5: If DaExt <> "pdf" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "pdf"
            End Select
        Else
            Select Case LiIdx
            Case 0: If DaExt <> "jpg" Then DatNa = DatNa & ".jpg"
            Case 1: If DaExt <> "png" Then DatNa = DatNa & ".png"
            Case 2: If DaExt <> "tif" Then DatNa = DatNa & ".tif"
            Case 3: If DaExt <> "bmp" Then DatNa = DatNa & ".bmp"
            Case 4: If DaExt <> "gif" Then DatNa = DatNa & ".gif"
            Case 5: If DaExt <> "pdf" Then DatNa = DatNa & ".pdf"
            End Select
        End If
    Case 11: 'Imageviewer
        If Posit > 0 Then
            Select Case LiIdx
            Case 0: If DaExt <> "jpg" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "jpg"
            Case 1: If DaExt <> "png" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "png"
            Case 2: If DaExt <> "tif" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "tif"
            Case 3: If DaExt <> "bmp" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "bmp"
            Case 4: If DaExt <> "gif" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "gif"
            Case 5: If DaExt <> "pdf" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "pdf"
            Case 6: If DaExt <> "ocx" Then DatNa = Mid$(DatNa, 1, Len(DatNa) - 3) & "docx"
            End Select
        Else
            Select Case LiIdx
            Case 0: If DaExt <> "jpg" Then DatNa = DatNa & ".jpg"
            Case 1: If DaExt <> "png" Then DatNa = DatNa & ".png"
            Case 2: If DaExt <> "tif" Then DatNa = DatNa & ".tif"
            Case 3: If DaExt <> "bmp" Then DatNa = DatNa & ".bmp"
            Case 4: If DaExt <> "gif" Then DatNa = DatNa & ".gif"
            Case 5: If DaExt <> "pdf" Then DatNa = DatNa & ".pdf"
            Case 6: If DaExt <> "ocx" Then DatNa = DatNa & ".docx"
            End Select
        End If
    End Select
    DaNam = DatNa
End If

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FCom " & Err.Number
Resume Next

End Sub
Private Sub FExpo()
On Error GoTo LdErr
'Exportieren

Dim ManNr As Long
Dim MitNr As Long
Dim DaStK As String
Dim ExFmt As String
Dim FiNam As String
Dim EmAdr As String
Dim EmBrf As String
Dim EmBet As String
Dim EmTex As String
Dim DaExt As String
Dim TypNa As String
Dim DaPas As String
Dim NePas As String
Dim DaNaO As String
Dim TmStr As String
Dim DoLnk As String
Dim PoLnk As String
Dim PaStr As String
Dim MaEma As String
Dim MaBrf As String
Dim DocNa As String
Dim BogNa As String
Dim EiTyp As Integer
Dim EmlSe As Integer
Dim LiIdx As Integer
Dim Lange As Integer
Dim Posit As Integer
Dim Zuord As Boolean
Dim RetWe As Boolean

Set FM = frmMain
Set FTeEm = Me.txtEmail
Set FTePr = Me.txtPraxi
Set FTeDo = Me.txtDocNa
Set CoDia = FM.comDialo
Set CmFma = Me.cmbForma
Set CmEml = Me.cmbEmail
Set ChZuo = Me.chkZuord

If WindowLoad("frmDruVo") = True Then
    Set TxCoN = frmTxVor.TexCont3
Else
    Set TxCoN = FM.TexCont1
End If

MitNr = GlMiA(GlSmI, 2) 'Standardmitarbeiter
ManNr = GlMan(GlSMa, 2) 'Standardmandant

TmGui = CreateID("D")

LiIdx = CmFma.ListIndex
EmlSe = CmEml.ListIndex

If ChZuo.Value = xtpChecked Then
    Zuord = True
End If

If ExMod <> 9 Then
    Unload Me
End If

BogNa = "Neuaufnahme"

Select Case ExMod
Case 1: 'Adresse
    EmTex = GlEmT(9, 1) & vbCrLf
    Select Case LiIdx
    Case 0: S_AdEx "csv"
    Case 1: S_AdEx "dtv"
    Case 2: S_AdEx "txt"
    Case 3: S_AdBDT
    Case 4: S_AdTx "ClientInfo", True
    Case 5: S_AdEx "pst"
    Case 6: S_AdEx "wgm"
    Case 7: S_PaExT
    End Select
Case 2: 'Fragebogen
    EmTex = GlEmT(3, 1) & vbCrLf
    Select Case LiIdx
    Case 0: ExFmt = "HTML"
    Case 1: ExFmt = "MHTML"
    Case 2: ExFmt = "XLS"
    Case 3: ExFmt = "RTF"
    Case 4: ExFmt = "PDF"
    End Select
    SExpo "AnFrBo", ExFmt, EmlSe, False, 0, True
Case 3: 'Tagesprotokoll
    EmTex = GlEmT(9, 1) & vbCrLf
    Select Case LiIdx
    Case 0: ExFmt = "HTML"
    Case 1: ExFmt = "MHTML"
    Case 2: ExFmt = "XLS"
    Case 3: ExFmt = "RTF"
    Case 4: ExFmt = "TXT"
    Case 5: ExFmt = "PDF"
    End Select
    SExpo "TagPro", ExFmt, EmlSe, False, 0, True
Case 4: 'Laborberichte
    EmTex = GlEmT(4, 1) & vbCrLf
    Select Case LiIdx
    Case 0: ExFmt = "LDT"
    Case 1: ExFmt = "HTML"
    Case 2: ExFmt = "MHTML"
    Case 3: ExFmt = "XLS"
    Case 4: ExFmt = "RTF"
    Case 5: ExFmt = "TXT"
    Case 6: ExFmt = "PDF"
    Case 7: ExFmt = "PICTURE_BMP"
    Case 8: ExFmt = "PICTURE_EMF"
    Case 9: ExFmt = "PICTURE_MULTITIFF"
    Case 10: ExFmt = "PICTURE_JPEG"
    End Select
    If ExFmt = "LDT" Then
        STran 1, EmlSe
    Else
        SExpo "LabExp", ExFmt, EmlSe, False, 0, True
    End If
Case 5: 'Laborauftrag
    EmTex = GlEmT(9, 1) & vbCrLf
    ExFmt = "LDT"
    STran 1, EmlSe
Case 6: 'Termine
    EmTex = GlEmT(5, 1) & vbCrLf
    Select Case LiIdx
    Case 0: S_Expor "xls", 0, 0
    Case 1: S_Expor "pst"
    Case 2: S_TeExT
    Case 3: S_BeDat "TerExp", EmlSe, False, 0, False, True
    End Select
Case 7: 'Terminliste
    EmTex = GlEmT(5, 1) & vbCrLf
    Select Case LiIdx
    Case 0: S_Expor "xls", 0, 0
    Case 1: S_Expor "ics", 0, 0
    Case 2: S_Expor "pst"
    Case 3: S_TeExT
    Case 4: S_BeDat "TerExp", EmlSe, False, 0, False, True
    End Select
Case 8: 'Offene Posten
    EmTex = GlEmT(9, 1) & vbCrLf
    Select Case LiIdx
    Case 0: 'DATEV
        S_Expor "csv", EmlSe, 0
    Case 1: 'XLS
        S_BeDat "AbrLis", EmlSe, False, 0, False, True
    Case 2: 'TXT
        S_Expor "txt", 0, 0
    Case 3: 'XML
        S_Expor "xml", 0, 0
    End Select
Case 9: 'Textverarbeitung

    With GlTxV
        Select Case SeUpl
        Case 0: .TxStr = GlEmT(9, 3)
        Case 1: .TxStr = GlEmT(13, 3)
        End Select
        .Datum = Date
        .DaStr = DaStK
        .MitNr = MitNr
        .ManNr = ManNr
        .PatNr = GlAdr
    End With
    EmBet = SEmTx()
    With GlTxV
        Select Case SeUpl
        Case 0: .TxStr = GlEmT(9, 1)
        Case 1: .TxStr = GlEmT(13, 1)
        End Select
        .Datum = Date
        .DaStr = DaStK
        .MitNr = MitNr
        .ManNr = ManNr
        .PatNr = GlAdr
        .DaNam = DaNam
        .DoNam = DaNam
    End With
    EmTex = SEmTx()

    If DaNam <> vbNullString Then
        DaExt = LCase(Right$(DaNam, 3))
        Select Case LiIdx
        Case 0: If DaExt <> "txm" Then DaNam = DaNam & ".txm"
        Case 1: If DaExt <> "pdf" Then DaNam = DaNam & ".pdf"
        Case 2: If DaExt <> "doc" Then DaNam = DaNam & ".doc"
        Case 3: If DaExt <> "ocx" Then DaNam = DaNam & ".docx"
        Case 4: If DaExt <> "rtf" Then DaNam = DaNam & ".rtf"
        Case 5: If DaExt <> "htm" Then DaNam = DaNam & ".htm"
        Case 6: If DaExt <> "txt" Then DaNam = DaNam & ".txt"
        Case 7: If DaExt <> "xml" Then DaNam = DaNam & ".xml"
        End Select
    End If

    If Zuord = False Then
        If SeUpl > 1 Then
            If FTeEm.Text = vbNullString Then
                SPopu "Keine Emailadresse", "Es wurde keine Emailadresse angegeben", IC48_Forbidden
                Exit Sub
            End If
            If FTePr.Text = vbNullString Then
                SPopu "Keine Praxisangaben", "Es wurde keine Praxisdaten angegeben", IC48_Forbidden
                Exit Sub
            End If
            If FTeDo.Text = vbNullString Then
                SPopu "Kein Dokumentenname", "Es wurde kein Dokumentenname angegeben", IC48_Forbidden
                Exit Sub
            End If
        End If

        If EmlSe = 0 Then
            With CoDia
                Select Case LiIdx
                Case 0: .DefaultExt = "*.txm"
                        .Filter = "Textverarbeitung (*.txm)|*.txm|Alle Dateien (*.*)|*.*"
                Case 1: .DefaultExt = "*.pdf"
                        .Filter = "Adobe-Acrobat Dokument (*.pdf)|*.pdf|Alle Dateien (*.*)|*.*"
                Case 2: .DefaultExt = "*.doc"
                        .Filter = "Microsof-Word 2003 Dokument (*.doc)|*.doc|Alle Dateien (*.*)|*.*"
                Case 3: .DefaultExt = "*.docx"
                        .Filter = "Microsof-Word 2007 Dokument (*.docx)|*.docx|Alle Dateien (*.*)|*.*"
                Case 4: .DefaultExt = "*.rtf"
                        .Filter = "Rich Text Dokument (*.rtf)|*.rtf|Alle Dateien (*.*)|*.*"
                Case 5: .DefaultExt = "*.htm"
                        .Filter = "Hypertext Markup Language (*.htm)|*.htm|Alle Dateien (*.*)|*.*"
                Case 6: .DefaultExt = "*.txt"
                        .Filter = "Acsii Textdatei (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
                Case 7: .DefaultExt = "*.xml"
                        .Filter = "XML-Datendatei (*.xml)|*.xml|Alle Dateien (*.*)|*.*"
                End Select
                .CancelError = True
                .DialogStyle = 1
                .DialogTitle = "Bitte Name und Ordner der Datei angeben"
                .FileName = GlEPf & DaNam
                .InitDir = GlEPf
                .ShowSave
                FiNam = .FileName
                If .FileTitle = vbNullString Then Exit Sub
            End With

            DaExt = LCase(Right$(FiNam, 3))
            Select Case LiIdx
            Case 0: If DaExt <> "txm" Then FiNam = FiNam & ".txm"
            Case 1: If DaExt <> "pdf" Then FiNam = FiNam & ".pdf"
            Case 2: If DaExt <> "doc" Then FiNam = FiNam & ".doc"
            Case 3: If DaExt <> "ocx" Then FiNam = FiNam & ".docx"
            Case 4: If DaExt <> "rtf" Then FiNam = FiNam & ".rtf"
            Case 5: If DaExt <> "htm" Then FiNam = FiNam & ".htm"
            Case 6: If DaExt <> "txt" Then FiNam = FiNam & ".txt"
            Case 7: If DaExt <> "xml" Then FiNam = FiNam & ".xml"
            End Select
        Else
            If Left$(DaNam, 2) <> "TD" Then
                PaStr = Format$(GlAdr, "000000")
                DaNaO = Left$(DaNam, Len(DaNam) - 4)
                DaNaO = SNaFi(DaNaO, True, True, True, True)
                DaNam = "TD" & PaStr & "_" & TmGui & "_" & DaNaO & ".pdf"
                S_TxEin
                DoEvents
                STxV2
                DoEvents
            End If
            If SeUpl = 0 Then
                FiNam = GlEPf & DaNam 'Exportordner
            Else
                FiNam = GlTEx & DaNam 'Termineordner
            End If
        End If
    End If

    Set clFil = New clsFile
    With clFil
        If .FilVor(FiNam) = True Then
            .DaLoe = FiNam & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
    DoEvents

    If LiIdx = 1 Then 'PDF
        If SeUpl < 2 Then 'Verschl¸sseln
            DaPas = IniGetVal("System", "DatPas")  'Datenbankpasswort Access
            TxCoN.LoadSaveAttribute(txMasterPassword) = DaPas
            TxCoN.LoadSaveAttribute(txDocAccessPermissions) = &H80
        End If
    End If

    Select Case LiIdx
    Case 0: TxCoN.Save FiNam, 0, 3
    Case 1: TxCoN.Save FiNam, 0, 12 'PDF
    Case 2: TxCoN.Save FiNam, 0, 9
    Case 3: TxCoN.Save FiNam, 0, 13
    Case 4: TxCoN.Save FiNam, 0, 5
    Case 5: TxCoN.Save FiNam, 0, 4
    Case 6: TxCoN.Save FiNam, 0, 1
    Case 7: TxCoN.Save FiNam, 0, 10
    End Select
    DoEvents

    If Zuord = False Then
        DaStK = Format$(Date, "ddd" & ", " & "dd" & ". " & "mmm" & Chr$(32) & "yyyy")
        MaEma = FTeEm.Text
        MaBrf = FTePr.Text
        DocNa = FTeDo.Text

        Unload Me
        DoEvents
        Select Case EmlSe
        Case 1: 'an Patienten
            S_AdDe GlAdr 'Adressendetails
            With GlADt
                ManNr = .AdMan
                EmAdr = .AdTe5
                EmBrf = .AdBrf
            End With
            With GlTxV
                .TxStr = GlEmT(9, 3)
                .MitNr = MitNr
                .ManNr = ManNr
                .PatNr = GlAdr
            End With
            EmBet = SEmTx()
            Select Case SeUpl
            Case 0:
                SMaNe GlAdr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet, FiNam
            Case 1: 'Dokument Downloadlink
                TmStr = SMCUp(FiNam)
                If TmStr <> vbNullString Then
                    Lange = Len(TmStr)
                    Posit = InStr(1, TmStr, ";", 1)
                    If Posit > 0 Then
                        DoLnk = Left$(TmStr, Posit - 1)
                        PoLnk = Mid$(TmStr, Posit + 1, Lange - Posit)
                        With GlTxV
                            .TxStr = GlEmT(13, 3)
                            .Datum = Date
                            .DaStr = DaStK
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = GlAdr
                        End With
                        EmBet = SEmTx()
                        With GlTxV
                            .TxStr = GlEmT(13, 1)
                            .Datum = Date
                            .DaStr = DaStK
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = GlAdr
                            .DoLnk = DoLnk
                            .PoLnk = PoLnk
                            .DoLan = GlVrw
                            .DoNam = DaNam
                        End With
                        EmTex = SEmTx()
                        SMaNe GlAdr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet
                    End If
                End If
            Case 2: 'Dokument Digitalunterschrift
                TmStr = SMDUp(FiNam, GlAdr, MaEma, MaBrf, DocNa)
                If TmStr <> vbNullString Then
                    DoLnk = TmStr
                    With GlTxV
                        .TxStr = GlEmT(14, 3)
                        .Datum = Date
                        .DaStr = DaStK
                        .MitNr = MitNr
                        .ManNr = ManNr
                        .PatNr = GlAdr
                    End With
                    EmBet = SEmTx()
                    With GlTxV
                        .TxStr = GlEmT(14, 1)
                        .Datum = Date
                        .DaStr = DaStK
                        .MitNr = MitNr
                        .ManNr = ManNr
                        .PatNr = GlAdr
                        .DoLnk = DoLnk
                        .PoLnk = PoLnk
                        .DoLan = GlVrw
                        .DoNam = DaNam
                    End With
                    EmTex = SEmTx()
                    SMaNe GlAdr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet
                End If
            Case 3: 'Dokument Neuaufnahmeformular
                TmStr = SMFUp(FiNam, MaEma, MaBrf, BogNa)
            End Select
        Case 2: 'An Mandanten
            S_AdDe GlAdr 'Adressendetails
            ManNr = S_AdIdx(GlAdr, "IDP")
            With GlTxV
                .TxStr = GlEmT(9, 3)
                .MitNr = MitNr
                .ManNr = ManNr
                .PatNr = GlAdr
            End With
            EmBet = SEmTx()
            S_AdDe ManNr 'Adressendetails
            With GlADt
                EmAdr = .AdTe5
                EmBrf = .AdBrf
            End With
            Select Case SeUpl
            Case 0:
                SMaNe ManNr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet, FiNam
            Case 1: 'Dokument Downloadlink
                TmStr = SMCUp(DaNam)
                If TmStr <> vbNullString Then
                    Lange = Len(TmStr)
                    Posit = InStr(1, TmStr, ";", 1)
                    If Posit > 0 Then
                        DoLnk = Left$(TmStr, Posit - 1)
                        PoLnk = Mid$(TmStr, Posit + 1, Lange - Posit)
                        With GlTxV
                            .TxStr = GlEmT(13, 3)
                            .Datum = Date
                            .DaStr = DaStK
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = GlAdr
                        End With
                        EmBet = SEmTx()
                        With GlTxV
                            .TxStr = GlEmT(13, 1)
                            .Datum = Date
                            .DaStr = DaStK
                            .MitNr = MitNr
                            .ManNr = ManNr
                            .PatNr = GlAdr
                            .DoLnk = DoLnk
                            .PoLnk = PoLnk
                            .DoLan = GlVrw
                            .DoNam = DaNam
                        End With
                        EmTex = SEmTx()
                        SMaNe ManNr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet
                    End If
                End If
            Case 2: 'Dokument Digitalunterschrift
                TmStr = SMDUp(FiNam, GlAdr, MaEma, MaBrf, DocNa)
                If TmStr <> vbNullString Then
                    DoLnk = TmStr
                    With GlTxV
                        .TxStr = GlEmT(14, 3)
                        .Datum = Date
                        .DaStr = DaStK
                        .MitNr = MitNr
                        .ManNr = ManNr
                        .PatNr = GlAdr
                    End With
                    EmBet = SEmTx()
                    With GlTxV
                        .TxStr = GlEmT(14, 1)
                        .Datum = Date
                        .DaStr = DaStK
                        .MitNr = MitNr
                        .ManNr = ManNr
                        .PatNr = GlAdr
                        .DoLnk = DoLnk
                        .PoLnk = PoLnk
                        .DoLan = GlVrw
                        .DoNam = DaNam
                    End With
                    EmTex = SEmTx()
                    SMaNe ManNr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet
                End If
            Case 3: 'Dokument Neuaufnahmeformular
                TmStr = SMFUp(FiNam, MaEma, MaBrf, BogNa)
            End Select
        End Select
    Else
        If DaExt = "pdf" Then
            EiTyp = 105
            TypNa = "PDF-Dokument"
        Else
            EiTyp = 102
            TypNa = "Textdokument"
        End If
        GlNeK = GlKoX
        With GlNeK
            .PatNr = GlAdr
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = EiTyp
            .KoStr = DaNam
            .KoGui = TmGui
            .TeStr = TypNa
            .NeuEi = True
            .Mitar = GlMiA(GlSmI, 2)
        End With
        K_Einf
        DoEvents
        SKrVo
        DoEvents
        Unload Me
    End If

Case 10: 'PDF-Viewer
        
    If DaNam <> vbNullString Then
        DaExt = LCase(Right$(DaNam, 3))
        Select Case LiIdx
        Case 0: If DaExt <> "jpg" Then DaNam = DaNam & ".jpg"
        Case 1: If DaExt <> "png" Then DaNam = DaNam & ".png"
        Case 2: If DaExt <> "tif" Then DaNam = DaNam & ".tif"
        Case 3: If DaExt <> "bmp" Then DaNam = DaNam & ".bmp"
        Case 4: If DaExt <> "gif" Then DaNam = DaNam & ".gif"
        Case 5: If DaExt <> "pdf" Then DaNam = DaNam & ".pdf"
        End Select
    End If

    If Zuord = False Then
        If EmlSe = 0 Then
            With CoDia
                Select Case LiIdx
                Case 0: .DefaultExt = "*.jpg"
                        .Filter = "Joint Photographic Experts Group (.jpg)|*.jpg|Alle Dateien (*.*)|*.*"
                Case 1: .DefaultExt = "*.png"
                        .Filter = "Portable Network Graphics (.png)|*.png|Alle Dateien (*.*)|*.*"
                Case 2: .DefaultExt = "*.tif"
                        .Filter = "Tagged Image Format (.tif)|*.tif|Alle Dateien (*.*)|*.*"
                Case 3: .DefaultExt = "*.bmp"
                        .Filter = "Windows Bitmap Format (.bmp)|*.bmp|Alle Dateien (*.*)|*.*"
                Case 4: .DefaultExt = "*.gif"
                        .Filter = "Graphics Interchange (.gif)|*.gif|Alle Dateien (*.*)|*.*"
                Case 5: .DefaultExt = "*.pdf"
                        .Filter = "Adobe-Acrobat Dokument (.pdf)|*.pdf|Alle Dateien (*.*)|*.*"
                End Select
                .CancelError = True
                .DialogStyle = 1
                .DialogTitle = "Bitte Name und Ordner der Datei angeben"
                .FileName = GlEPf & DaNam
                .InitDir = GlEPf
                .ShowSave
                FiNam = .FileName
                If .FileTitle = vbNullString Then Exit Sub
            End With
            
            DaExt = LCase(Right$(FiNam, 3))
            Select Case LiIdx
            Case 0: If DaExt <> "jpg" Then FiNam = FiNam & ".jpg"
            Case 1: If DaExt <> "png" Then FiNam = FiNam & ".png"
            Case 2: If DaExt <> "tif" Then FiNam = FiNam & ".tif"
            Case 3: If DaExt <> "bmp" Then FiNam = FiNam & ".bmp"
            Case 4: If DaExt <> "gif" Then FiNam = FiNam & ".gif"
            Case 5: If DaExt <> "pdf" Then FiNam = FiNam & ".pdf"
            End Select
        Else
            FiNam = GlEPf & DaNam
        End If
    Else
        FiNam = GlBPf & DaNam
    End If

    Set clFil = New clsFile
    With clFil
        If .FilVor(FiNam) = True Then
            .DaLoe = FiNam & vbNullChar
            .FilLoe
        End If
    End With
    Set clFil = Nothing
    DoEvents

    If Zuord = False Then
        Unload Me
        DoEvents
        Select Case EmlSe
        Case 1: 'An Patienten
            S_AdDe GlAdr 'Adressendetails
            With GlADt
                EmAdr = .AdTe5
                EmBrf = .AdBrf
                EmBet = .AdKur
            End With
            SMaNe GlAdr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet, FiNam
        Case 2: 'An Mandanten
            ManNr = S_AdIdx(GlAdr, "IDP")
            EmBet = S_AdIdx(GlAdr, "IDKurz")
            S_AdDe ManNr 'Adressendetails
            With GlADt
                EmAdr = .AdTe5
                EmBrf = .AdBrf
            End With
            SMaNe GlAdr, EmAdr, , EmBrf & vbCrLf & vbCrLf & EmTex, EmBet, FiNam
        End Select
    Else
        EiTyp = 105
        TypNa = "Bilddokument"
        GlNeK = GlKoX
        With GlNeK
            .PatNr = GlAdr
            .IdxNr = 0
            .EiDat = Format$(Date, "dd.mm.yyyy")
            .EiZei = TimeValue(Now)
            .EiTyp = EiTyp
            .KoStr = DaNam
            .KoGui = TmGui
            .TeStr = TypNa
            .NeuEi = True
            .Mitar = GlMiA(GlSmI, 2)
        End With
        K_Einf
        DoEvents
        SKrVo
        DoEvents
        Unload Me
    End If

Case 12: 'Kontoums‰tze

    EmTex = GlEmT(9, 1) & vbCrLf
    Select Case LiIdx
    Case 0: S_Expor "xls", 0, 0
    Case 1: S_Expor "txt", 0, 0
    End Select
    
Case Else:

    EmTex = GlEmT(9, 1) & vbCrLf
    Select Case LiIdx
    Case 0: S_Expor "csv", 0, 0
    Case 1: S_Expor "xls", 0, 0
    Case 2: S_Expor "txt", 0, 0
    Case 3: S_Expor "xml", 0, 0
    Case 4: S_BeDat "BuList", EmlSe, False, 0, False, True
    End Select
    
End Select

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FExpo " & Err.Number
Resume Next

End Sub
Private Sub FLoad()
On Error GoTo LdErr

Dim AkSei As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set FTeEm = Me.txtEmail
Set FTePr = Me.txtPraxi
Set FTeDo = Me.txtDocNa
Set CmFma = Me.cmbForma
Set CmEml = Me.cmbEmail
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set ChZuo = Me.chkZuord

If GlBut = RibTab_Adressen Then
    Set RpCo2 = FM.repCont2
    Set RpCls = RpCo2.Columns
    Set RpSel = RpCo2.SelectedRows
ElseIf GlBut = RibTab_LabBericht Then
    Set RpCo5 = FM.repCont5
    Set RpCls = RpCo5.Columns
    Set RpSel = RpCo5.SelectedRows
Else
    Set RpCo1 = FM.repCont1
    Set RpCls = RpCo1.Columns
    Set RpSel = RpCo1.SelectedRows
    If SeUpl > 1 Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        FPrax
    End If
End If

Select Case ExMod
Case 1: 'Adressen
    With CmFma
        .AddItem "Microsoft Excel (.csv)"
        .ItemData(.NewIndex) = 1
        .AddItem "DATEV 4.0 Datei (.csv)"
        .ItemData(.NewIndex) = 2
        .AddItem "Acsii Textdatei (.txt)"
        .ItemData(.NewIndex) = 3
        .AddItem "GDT-Datendatei (.gdt)"
        .ItemData(.NewIndex) = 4
        .AddItem "THEDEX-Dateien (.thx)"
        .ItemData(.NewIndex) = 5
        .AddItem "Outlook Kontakte (.psx)"
        .ItemData(.NewIndex) = 6
        .AddItem "WEGAMED-Daten (.wgm)"
        .ItemData(.NewIndex) = 7
        .AddItem "SMA Stammdaten (.sma)"
        .ItemData(.NewIndex) = 8
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = False
Case 2: 'Fragebogen
    With CmFma
        .AddItem "HTML-Format (.htm)"
        .ItemData(.NewIndex) = 1
        .AddItem "Multi-Mime-HTML (.mht)"
        .ItemData(.NewIndex) = 2
        .AddItem "Microsoft Excel (.xel)"
        .ItemData(.NewIndex) = 3
        .AddItem "Rich Text-Format (.rtf)"
        .ItemData(.NewIndex) = 4
        .AddItem "Adobe Acrobat (.pdf)"
        .ItemData(.NewIndex) = 5
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = False
Case 3: 'Tagesprotokoll
    With CmFma
        .AddItem "HTML-Format (.htm)"
        .ItemData(.NewIndex) = 1
        .AddItem "Multi-Mime-HTML (.mht)"
        .ItemData(.NewIndex) = 2
        .AddItem "Microsoft Excel (.xls)"
        .ItemData(.NewIndex) = 3
        .AddItem "Rich Text-Format (.rft)"
        .ItemData(.NewIndex) = 4
        .AddItem "Ascii-Text-Format (.txt)"
        .ItemData(.NewIndex) = 5
        .AddItem "Adobe Acrobat (.pdf)"
        .ItemData(.NewIndex) = 6
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = False
Case 4: 'Laborberichte
    With CmFma
        .AddItem "Labordatentr‰ger (.ldt)"
        .ItemData(.NewIndex) = 1
        .AddItem "HTML-Format (.htm)"
        .ItemData(.NewIndex) = 2
        .AddItem "Multi-Mime-HTML (.mht)"
        .ItemData(.NewIndex) = 3
        .AddItem "Microsoft Excel (.xls)"
        .ItemData(.NewIndex) = 4
        .AddItem "Rich Text-Format (.rtf)"
        .ItemData(.NewIndex) = 5
        .AddItem "Ascii-Text-Format (.txt)"
        .ItemData(.NewIndex) = 6
        .AddItem "Adobe Acrobat (.pdf)"
        .ItemData(.NewIndex) = 7
        .AddItem "Windows Bitmap (.bmp)"
        .ItemData(.NewIndex) = 8
        .AddItem "Enhanced Metafile (.emf)"
        .ItemData(.NewIndex) = 9
        .AddItem "Multi-TIFF-Format (.tif)"
        .ItemData(.NewIndex) = 10
        .AddItem "JPEG-Format (.jpg)"
        .ItemData(.NewIndex) = 11
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = True
Case 5: 'Laborauftrag
    With CmFma
        .AddItem "Labordatentr‰ger (.ldt)"
        .ItemData(.NewIndex) = 1
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = True
Case 6: 'Termine
    With CmFma
        .AddItem "Microsoft Excel (.csv)"
        .ItemData(.NewIndex) = 1
        .AddItem "iCalendar Datei (*.ics)"
        .ItemData(.NewIndex) = 2
        .AddItem "Outlook Termine (.psx)"
        .ItemData(.NewIndex) = 3
        .AddItem "SimpliMed Termine (.smt)"
        .ItemData(.NewIndex) = 4
        .AddItem "Terminexportliste (.xls)"
        .ItemData(.NewIndex) = 5
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = False
Case 7: 'Terminliste
    With CmFma
        .AddItem "Microsoft Excel (.csv)"
        .ItemData(.NewIndex) = 1
        .AddItem "iCalendar Datei (*.ics)"
        .ItemData(.NewIndex) = 2
        .AddItem "Outlook Termine (.psx)"
        .ItemData(.NewIndex) = 3
        .AddItem "SimpliMed Termine (.smt)"
        .ItemData(.NewIndex) = 4
        .AddItem "Terminexportliste (.xls)"
        .ItemData(.NewIndex) = 5
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = False
Case 8: 'Offene Posten
    With CmFma
        .AddItem "DATEV 4.0 Dateien (.csv)"
        .ItemData(.NewIndex) = 1
        .AddItem "Microsoft Excel (.xls)"
        .ItemData(.NewIndex) = 2
        .AddItem "Lexware-Dateien (.txt)"
        .ItemData(.NewIndex) = 3
        .AddItem "XML-Datendatei (.xml)"
        .ItemData(.NewIndex) = 4
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = True
Case 9: 'Textverarbeitung
    With CmFma
        Select Case GlBut
        Case RibTab_Tex_Dokumt:
                .AddItem "Textverarbeitung (.txm)"
                .ItemData(.NewIndex) = 1
        Case RibTab_Tex_Vorlag:
                .AddItem "Textverarbeitung (.txm)"
                .ItemData(.NewIndex) = 1
        Case RibTab_Tex_Rezept:
                .AddItem "Langrezepte (.txr)"
                .ItemData(.NewIndex) = 1
        Case RibTab_Tex_NewsLe:
                .AddItem "Newslettervorlage (.txn)"
                .ItemData(.NewIndex) = 1
        Case RibTab_Krankenbla:
                .AddItem "Textverarbeitung (.txm)"
                .ItemData(.NewIndex) = 1
        End Select
        .AddItem "Adobe-Acrobat Dokument (.pdf)"
        .ItemData(.NewIndex) = 2
        .AddItem "Microsof-Word 2003 Dokument (.doc)"
        .ItemData(.NewIndex) = 3
        .AddItem "Microsof-Word 2007 Dokument (.docx)"
        .ItemData(.NewIndex) = 4
        .AddItem "Rich Text Dokument (.rtf)"
        .ItemData(.NewIndex) = 5
        .AddItem "Hypertext Markup Language (.htm)"
        .ItemData(.NewIndex) = 6
        .AddItem "Windows Ansi-Text Format (.txt)"
        .ItemData(.NewIndex) = 7
        .AddItem "Extensible Markup Language (.xml)"
        .ItemData(.NewIndex) = 8
        .ListIndex = FoIdx
    End With
    ChZuo.Enabled = True

    If SeUpl > 1 Then 'Dokumentupload
        CmFma.Enabled = False
        CmEml.Enabled = False
        ChZuo.Enabled = False
    End If

Case 10: 'PDF-Viewer
     With CmFma
        .AddItem "Joint Photographic Experts Group (.jpg)"
        .ItemData(.NewIndex) = 1
        .AddItem "Portable Network Graphics (.png)"
        .ItemData(.NewIndex) = 2
        .AddItem "Tagged Image Format (.tif)"
        .ItemData(.NewIndex) = 3
        .AddItem "Windows Bitmap Format (.bmp)"
        .ItemData(.NewIndex) = 4
        .AddItem "Graphics Interchange (.gif)"
        .ItemData(.NewIndex) = 5
        .AddItem "Adobe Acrobat (.pdf)"
        .ItemData(.NewIndex) = 6
        .ListIndex = FoIdx
    End With
    ChZuo.Enabled = True
    CmEml.Enabled = False
Case 11: 'Imageviewer
     With CmFma
        .AddItem "Joint Photographic Experts Group (.jpg)"
        .ItemData(.NewIndex) = 1
        .AddItem "Portable Network Graphics (.png)"
        .ItemData(.NewIndex) = 2
        .AddItem "Tagged Image Format (.tif)"
        .ItemData(.NewIndex) = 3
        .AddItem "Windows Bitmap Format (.bmp)"
        .ItemData(.NewIndex) = 4
        .AddItem "Graphics Interchange (.gif)"
        .ItemData(.NewIndex) = 5
        .AddItem "Adobe Acrobat (.pdf)"
        .ItemData(.NewIndex) = 6
        .AddItem "Microsof-Word 2007 Dokument (.docx)"
        .ItemData(.NewIndex) = 7
        .ListIndex = FoIdx
    End With
    ChZuo.Enabled = True
    CmEml.Enabled = False
Case 12: 'Kontoums‰tze
    With CmFma
        .AddItem "Microsoft Excel (.csv)"
        .ItemData(.NewIndex) = 1
        .ListIndex = FoIdx
    End With
    CmEml.Enabled = False
Case Else:
    With CmFma
        .AddItem "DATEV 4.0 Dateien (*.csv)"
        .ItemData(.NewIndex) = 1
        .AddItem "Microsoft Excel (*.xls)"
        .ItemData(.NewIndex) = 2
        .AddItem "Lexware-Dateien (*.txt)"
        .ItemData(.NewIndex) = 3
        .AddItem "XML-Datendatei (*.xml)"
        .ItemData(.NewIndex) = 4
        .AddItem "Buchungsexportliste (*.xls)"
        .ItemData(.NewIndex) = 5
        .ListIndex = FoIdx
    End With
End Select

If SeUpl > 0 Then
    With CmEml
        .AddItem "Kein Emailversand"
        .ItemData(0) = 1
        .AddItem "Downloadlink an einen Patienten"
        .ItemData(1) = 2
        .AddItem "Downloadlink an den Mandanten"
        .ItemData(2) = 3
        .AddItem "Downloadlink an alle Patienten"
        .ItemData(3) = 4
        .ListIndex = EmIdx
    End With
Else
    With CmEml
        .AddItem "Kein Emailversand"
        .ItemData(0) = 1
        .AddItem "Emailversand an einen Patienten"
        .ItemData(1) = 2
        .AddItem "Emailversand an den Mandanten"
        .ItemData(2) = 3
        .AddItem "Emailversand an alle Patienten"
        .ItemData(3) = 4
        .ListIndex = EmIdx
    End With
End If

CmFma.MaxLength = 13

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
ChZuo.BackColor = GlBak

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set RpCo5 = Nothing

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FPrax()
On Error GoTo InErr
'Sammelt die Praxisangaben

Dim MitNr As Long
Dim ManNr As Long
Dim MaEma As String
Dim MaNam As String
Dim MaBrf As String
Dim DocNa As String
Dim AktZa As Integer
Dim Lange As Integer

Set FM = frmMain
Set FTeEm = Me.txtEmail
Set FTePr = Me.txtPraxi
Set FTeDo = Me.txtDocNa

MitNr = GlMiA(GlSmI, 2)

If GlEKV = True Then 'Emailkonten vorhanden
    For AktZa = 1 To UBound(GlMkt)
        If CLng(GlMkt(AktZa, 1)) = MitNr Then
            If CBool(GlMkt(AktZa, 20)) = True Then 'Standardemailkonto
                If GlMkt(AktZa, 13) <> vbNullString Then
                    MaEma = GlMkt(AktZa, 13)
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

For AktZa = 1 To UBound(GlThe) 'Mandanten
    If ManNr = GlThe(AktZa, 0) Then
        MaNam = GlThe(AktZa, 13)
        MaBrf = GlThe(AktZa, 36)
        If MaEma = vbNullString Then
            MaEma = GlThe(AktZa, 16)
        End If
        Exit For
    End If
Next AktZa

If DaNam <> vbNullString Then
    DocNa = DaNam
    Lange = Len(DocNa)
    If Left$(DocNa, 1) = "_" Then
        DocNa = Right$(DocNa, Lange - 1)
    End If
    If Left$(DocNa, 2) = "TD" Then
        Lange = Len(DocNa)
        DocNa = Right$(DocNa, Lange - 43)
    End If
Else
    If GlTxK <> vbNullString Then
        DocNa = GlTxK
    End If
End If

FTeDo.Text = DocNa
FTeEm.Text = MaEma
FTePr.Text = MaBrf

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPrax " & Err.Number
Resume Next

End Sub


Private Sub FZuor()
On Error GoTo LdErr
'Zuordnen

Set ChZuo = Me.chkZuord
Set CmFma = Me.cmbForma
Set CmEml = Me.cmbEmail

Dim PaStr As String
Dim NeuDa As String
Dim DaExt As String
Dim LiIdx As Integer

LiIdx = CmFma.ListIndex

If ChZuo = xtpChecked Then
    If DaNam <> vbNullString Then
        AltNa = DaNam
    End If
    
    Select Case ExMod
    Case 9: 'Textverarbeitung
        Select Case LiIdx
        Case 0: DaExt = "txm"
        Case 1: DaExt = "pdf"
        Case 2: DaExt = "doc"
        Case 3: DaExt = "docx"
        Case 4: DaExt = "rtf"
        Case 5: DaExt = "htm"
        Case 6: DaExt = "txt"
        Case 7: DaExt = "xml"
        End Select
    Case 10: 'PDF-Viewer
        Select Case LiIdx
        Case 0: DaExt = "jpg"
        Case 1: DaExt = "png"
        Case 2: DaExt = "tif"
        Case 3: DaExt = "bmp"
        Case 4: DaExt = "gif"
        Case 5: DaExt = "pdf"
        End Select
    Case 11: 'Imageviewer
        Select Case LiIdx
        Case 0: DaExt = "jpg"
        Case 1: DaExt = "png"
        Case 2: DaExt = "tif"
        Case 3: DaExt = "bmp"
        Case 4: DaExt = "gif"
        Case 5: DaExt = "pdf"
        Case 6: DaExt = "docx"
        End Select
    End Select
            
    PaStr = "P" & Format$(GlAdr, "000000")
    TmGui = CreateID("D")
    
    NeuDa = PaStr & "_" & TmGui & "#_" & "Exportdokument" & "." & LCase(DaExt)
    DaNam = NeuDa
    
    CmEml.ListIndex = 0
    CmEml.Enabled = False
Else
    If AltNa <> vbNullString Then
        DaNam = AltNa
    End If
    CmEml.Enabled = True
End If

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZuor " & Err.Number
Resume Next

End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50231)
TeMai = IniGetOpt("Hilfe", 50232)
TeInh = IniGetOpt("Hilfe", 50233)
TeFus = IniGetOpt("Hilfe", 50234)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    
If GlRch(0, 22) = 0 Then
    WindowMess "Sie besitzen keine Berechtigung f¸r diesen Vorgang", Dial3, "Benutzerrechte", FM.hwnd
    Exit Sub
End If
    
FExpo

End Sub

Private Sub chkZuord_Click()
    FZuor
End Sub
Private Sub cmbForma_Click()

If GlKeL = False Then
    Select Case ExMod
    Case 9: FCom
    Case 10: FCom
    Case 11: FCom
    End Select
End If

End Sub

Private Sub Form_Activate()
    GlKeL = False
End Sub
Private Sub Form_Load()
On Error Resume Next

FLoad
AFont Me
SFrame 1, Me.hwnd
GlKeL = True

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmExport = Nothing
End Sub

Private Sub txtDocNa_GotFocus()
    Me.txtDocNa.SelStart = 0
    Me.txtDocNa.SelLength = Len(Me.txtDocNa.Text)
End Sub
Private Sub txtEmail_GotFocus()
    Me.txtEmail.SelStart = 0
    Me.txtEmail.SelLength = Len(Me.txtEmail.Text)
End Sub

Private Sub txtPraxi_GotFocus()
    Me.txtPraxi.SelStart = 0
    Me.txtPraxi.SelLength = Len(Me.txtPraxi.Text)
End Sub
