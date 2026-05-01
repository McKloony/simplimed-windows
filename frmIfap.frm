VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmIfap 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ifap praxisCENTER"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "frmIfap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   10
      Top             =   6600
      Width           =   6600
      _Version        =   1048579
      _ExtentX        =   11642
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4600
         TabIndex        =   14
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
         Left            =   3200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Weiter >"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnZurück 
         Height          =   400
         Left            =   1800
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hausliste"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   500
         TabIndex        =   11
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
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8500
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   5900
      Left            =   400
      TabIndex        =   17
      Top             =   660
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   10407
      _StockProps     =   79
      Appearance      =   6
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbSuch7 
         Height          =   315
         Left            =   1000
         TabIndex        =   9
         Top             =   5200
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch5 
         Height          =   350
         Left            =   1000
         TabIndex        =   7
         Top             =   3760
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch4 
         Height          =   350
         Left            =   1000
         TabIndex        =   6
         Top             =   3040
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch3 
         Height          =   350
         Left            =   1000
         TabIndex        =   5
         Top             =   2320
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   3730
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1600
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch2 
         Height          =   350
         Left            =   1000
         TabIndex        =   2
         Top             =   1060
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtSuch1 
         Height          =   350
         Left            =   1000
         TabIndex        =   1
         Top             =   340
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   2500
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1600
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkIfAbg 
         Height          =   220
         Left            =   2500
         TabIndex        =   21
         Top             =   1600
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Heilmittelabgleich"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox cmbSuch6 
         Height          =   315
         Left            =   1000
         TabIndex        =   8
         Top             =   4480
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.Label lblLabl8 
         Height          =   220
         Left            =   1010
         TabIndex        =   36
         Top             =   4950
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach ATC :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl1 
         Height          =   220
         Left            =   1010
         TabIndex        =   35
         Top             =   90
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach Produktbezeichnung :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl2 
         Height          =   220
         Left            =   1010
         TabIndex        =   34
         Top             =   810
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach PZN (Pharmazentralnummer) :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl7 
         Height          =   220
         Left            =   1010
         TabIndex        =   33
         Top             =   4230
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach ICD Diagnosecode :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl6 
         Height          =   220
         Left            =   1010
         TabIndex        =   32
         Top             =   3510
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "PZN-Schnellauskunft :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl5 
         Height          =   220
         Left            =   1010
         TabIndex        =   31
         Top             =   2790
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach Wirkstoff :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl4 
         Height          =   220
         Left            =   1010
         TabIndex        =   30
         Top             =   2070
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Suche nach Hersteller / Anbieter :"
         Alignment       =   4
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLabl3 
         Height          =   220
         Left            =   1010
         TabIndex        =   20
         Top             =   1640
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Behandlungsdatum :"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPatNr 
      Height          =   200
      Left            =   400
      TabIndex        =   18
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtReNum 
      Height          =   200
      Left            =   800
      TabIndex        =   19
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   2100
      Left            =   100
      TabIndex        =   22
      Top             =   660
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   3704
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ProgressBar prbStat1 
         Height          =   315
         Left            =   460
         TabIndex        =   23
         Top             =   600
         Width           =   4760
         _Version        =   1048579
         _ExtentX        =   8396
         _ExtentY        =   556
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.ProgressBar prbStat2 
         Height          =   315
         Left            =   460
         TabIndex        =   24
         Top             =   1200
         Width           =   4760
         _Version        =   1048579
         _ExtentX        =   8396
         _ExtentY        =   556
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   210
         Left            =   460
         TabIndex        =   26
         Top             =   1600
         Width           =   2955
         _Version        =   1048579
         _ExtentX        =   5221
         _ExtentY        =   370
         _StockProps     =   79
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   220
         Left            =   460
         TabIndex        =   25
         Top             =   310
         Width           =   2955
         _Version        =   1048579
         _ExtentX        =   5212
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Bitte warten..."
         Alignment       =   4
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtPaVor 
      Height          =   200
      Left            =   1200
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPaNam 
      Height          =   200
      Left            =   1600
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPaGeb 
      Height          =   200
      Left            =   2000
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte gebene Sie die gewünschte Produktbezeichnung oder Pharmazentralnummer ein und klicken auf Weiter."
      Height          =   495
      Left            =   1000
      TabIndex        =   16
      Top             =   140
      Width           =   4000
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   660
      Left            =   0
      Top             =   0
      Width           =   6600
   End
End
Attribute VB_Name = "frmIfap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private LaCa1 As XtremeSuiteControls.Label
Private LaCa2 As XtremeSuiteControls.Label
Private Labl1 As XtremeSuiteControls.Label
Private Labl2 As XtremeSuiteControls.Label
Private Labl3 As XtremeSuiteControls.Label
Private Labl4 As XtremeSuiteControls.Label
Private Labl5 As XtremeSuiteControls.Label
Private Labl6 As XtremeSuiteControls.Label
Private Labl7 As XtremeSuiteControls.Label
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private TxDum As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxSu1 As XtremeSuiteControls.FlatEdit
Private TxSu2 As XtremeSuiteControls.FlatEdit
Private TxSu3 As XtremeSuiteControls.FlatEdit
Private TxSu4 As XtremeSuiteControls.FlatEdit
Private TxSu5 As XtremeSuiteControls.FlatEdit
Private TxPNu As XtremeSuiteControls.FlatEdit
Private TxPVo As XtremeSuiteControls.FlatEdit
Private TxPNa As XtremeSuiteControls.FlatEdit
Private TxPGe As XtremeSuiteControls.FlatEdit
Private TxRen As XtremeSuiteControls.FlatEdit
Private CmAus As XtremeSuiteControls.ComboBox
Private CmSu6 As XtremeSuiteControls.ComboBox
Private CmSu7 As XtremeSuiteControls.ComboBox
Private ChAbg As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private ImMan As XtremeCommandBars.ImageManager
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private MoKal As XtremeCalendarControl.DatePicker
Private TxCoN As Tx4oleLib.TXTextControl

Private GesZa As Long

Private RS01 As ADODB.Recordset
Private RS02 As ADODB.Recordset
Private RS03 As ADODB.Recordset
Private RS04 As ADODB.Recordset
Private RS05 As ADODB.Recordset
Private RS06 As ADODB.Recordset
Private RS07 As ADODB.Recordset
Private RS08 As ADODB.Recordset
Private RS09 As ADODB.Recordset

Private WithEvents IfIdx As praxisCENTER3.ApplicationObject
Attribute IfIdx.VB_VarHelpID = -1

'Private IfRzM As praxisCENTER3.PrescriptionMedicament
'Private IfDos As praxisCENTER3.DosageForm
'Private IfHer As praxisCENTER3.Supplier
'Private IfPhy As praxisCENTER3.Physician
'Private IfPts As praxisCENTER3.Patients
'Private IfPat As praxisCENTER3.patient
'Private IfDis As praxisCENTER3.Diagnoses
'Private IfDia As praxisCENTER3.Diagnose
'Private IfMes As praxisCENTER3.Medicaments
'Private IfMed As praxisCENTER3.Medicament

Private IfRzM As Object
Private IfDos As Object
Private IfHer As Object
Private IfPhy As Object
Private IfPts As Object
Private IfPat As Object
Private IfDis As Object
Private IfDia As Object
Private IfMes As Object
Private IfMed As Object

Private clFil As clsFile
Private clFen As clsFenster

Private EiPZN() As String
Private MePZN() As String
Private MeBez() As String
Private HerNa() As String
Private Dosag() As String
Private MePre() As Single
Private EiPre() As Single
Public Sub FDAbg()
On Error GoTo DaErr
'Gleicht den Arzneikatalog ab

Dim KatNr As Long
Dim ManNr As Long
Dim LanNr As String
Dim ManNa As String
Dim ManPl As String
Dim ManVo As String
Dim KatNa As String
Dim AktZa As Integer

Set FM = frmIfap
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set LaCa1 = FM.lblLab01
Set LaCa2 = FM.lblLab02

If GlMiV = True Then 'Mitarbeiter vorhanden
    ManNr = GlMiA(GlSmI, 7) 'zugeordnete Mandantennummer
Else
    ManNr = GlMan(GlSMa, 2) 'Standardmandant
End If

For AktZa = 1 To UBound(GlThe) 'Mandanten
    If ManNr = GlThe(AktZa, 0) Then
        LanNr = GlThe(AktZa, 15) 'Lebenslange Arztnummer
        ManVo = GlThe(AktZa, 1)
        ManNa = GlThe(AktZa, 2)
        ManPl = GlThe(AktZa, 4)
    End If
Next AktZa

Set IfIdx = New praxisCENTER3.ApplicationObject
'Set IfIdx = CreateObject("praxisCENTER3.ApplicationObject")
IfIdx.Activate

Set IfPat = New praxisCENTER3.patient
'Set IfPat = CreateObject("praxisCENTER3.Patient")
With IfPat
    .id = "ID0"
    .LastName = GlPrg
End With

Set IfPhy = New praxisCENTER3.Physician
'Set IfPhy = CreateObject("praxisCENTER3.Physician")

If GlRDP = True Then
    With IfPhy
        .LANR = "999999934"
        .FirstName = "..."
        .LastName = "..."
        .Activate
    End With
Else
    With IfPhy
        .LANR = LanNr
        .FirstName = ManVo
        .LastName = ManNa
        .PostalCode = ManPl
        .Activate
    End With
End If

IfIdx.Patients.Add IfPat

IfPat.Activate

Rahm2.Visible = False
Rahm1.Visible = True

Set RS06 = New ADODB.Recordset
RS06.CursorLocation = adUseClient
Set RS06 = DBCmRe0("qryKat04")
Set RS06.ActiveConnection = Nothing
If RS06.RecordCount > 0 Then
    Do
    KatNr = RS06.Fields("ID3").Value
    KatNa = RS06.Fields("IDKurz").Value
    LaCa2.Caption = KatNa
    FKaOp KatNr 'Alle Einträge eines Kataloges einlesen
    DoEvents
    FdAbL 'Einlesen der Daten aus dem ifap
    DoEvents
    FDAbs KatNr 'Abspeichern in die eigenen Datenbank
    RS06.MoveNext
    Loop Until RS06.EOF
End If
RS06.Close
Set RS06 = Nothing

DoEvents
LaCa1.Caption = vbNullString
LaCa2.Caption = vbNullString
DoEvents

IfIdx.Hide
Set IfIdx = Nothing

DoEvents
Unload Me

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDAbg " & Err.Number
Resume Next

End Sub
Private Sub FDAbs(ByVal KatNr As Long)
On Error GoTo DaErr
'Speichert die Daten in die Datenbank

Dim AktZa As Long
Dim SQL2 As String
Dim Krite As String
Dim SuStr As String

Set FM = frmIfap
Set LaCa1 = FM.lblLab01
Set PrBr1 = FM.prbStat1
Set PrBr2 = FM.prbStat2
Set TxDum = FM.txtDummy

PrBr2.Min = 0
PrBr2.Max = GesZa

If GlTyp < 2 Then
    SQL2 = "SELECT * FROM dbo.qryKat04C WHERE ID3 = " & KatNr
Else
    SQL2 = "SELECT * FROM qryKat04C WHERE [ID3] = " & KatNr & ";"
End If

Set RS04 = New ADODB.Recordset
With RS04
    .CursorLocation = adUseClient
    .Source = SQL2
    .ActiveConnection = DB1
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open Options:=adCmdText
End With

If RS04.RecordCount > 0 Then
    For AktZa = 0 To GesZa - 1
        SuStr = Format$(MePZN(AktZa), "0000000")
    
        Krite = "[GOID] Like '" & SuStr & "'"
        RS04.Filter = Krite

        If RS04.RecordCount > 0 Then
            Do
            If Len(MeBez(AktZa)) < 8 Then
                RS04.Fields("Automatic").Value = -1
                RS04.Fields("IDKurz").Value = "a.V."
                RS04.Fields("Preis1").Value = 0
                RS04.Fields("Preis3").Value = 0
            Else
                RS04.Fields("Automatic").Value = -1
                RS04.Fields("IDKurz").Value = MeBez(AktZa)
                RS04.Fields("Preis1").Value = MePre(AktZa)
                RS04.Fields("Preis3").Value = EiPre(AktZa)
            End If
            RS04.Update
            RS04.MoveNext
            Loop Until RS04.EOF
            LaCa1.Caption = MeBez(AktZa)
        Else
            LaCa1.Caption = vbNullString
        End If
        PrBr2.Value = AktZa
        If TxDum.Text = "B" Then Exit For 'Abbrechen
        RS04.Filter = adFilterNone
        DoEvents
    Next AktZa
End If
RS04.Close
Set RS04 = Nothing

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDAbs " & Err.Number
Resume Next

End Sub
Private Sub FdAbL()
On Error GoTo LaErr

Dim AktZa As Long
Dim PatNr As Long
Dim RetW1 As Long
Dim RetW2 As Long
Dim HerNr As Long
Dim TmpPz As String
Dim TmpSu As String
Dim Einhe As String
Dim Posit As Integer
Dim Anzal As Integer
Dim Lange As Integer
Dim Menge() As String

Set FM = frmIfap
Set LaCa1 = FM.lblLab01
Set LaCa2 = FM.lblLab02
Set PrBr1 = FM.prbStat1
Set PrBr2 = FM.prbStat2
Set TxDum = FM.txtDummy

ReDim Preserve MePZN(GesZa)
ReDim Preserve MeBez(GesZa)
ReDim Preserve MePre(GesZa)
ReDim Preserve EiPre(GesZa)
ReDim Preserve HerNa(GesZa)
ReDim Preserve Menge(GesZa)

PrBr1.Min = 0
PrBr1.Max = GesZa
DoEvents
For AktZa = 0 To GesZa - 1
    TmpSu = EiPZN(AktZa)
    
    Set IfMed = New praxisCENTER3.Medicament
    'Set IfMed = CreateObject("praxisCENTER3.Medicament")
    RetW1 = IfIdx.GetMedicament(TmpSu, IfMed)
    
    If RetW1 = 0 Then
        If Left$(EiPZN(AktZa), 1) = "0" Then
            TmpPz = IfMed.PIC
            Lange = Len(TmpPz)
            If Lange < 7 Then
                Select Case Lange
                Case 4: TmpPz = "000" & IfMed.PIC
                Case 5: TmpPz = "00" & IfMed.PIC
                Case 6: TmpPz = "0" & IfMed.PIC
                End Select
                MePZN(AktZa) = TmpPz
            Else
                MePZN(AktZa) = IfMed.PIC
            End If
        Else
            MePZN(AktZa) = IfMed.PIC
        End If
        
        MeBez(AktZa) = IfMed.DefaultPrintName
        MePre(AktZa) = CSng(Replace(IfMed.PharmacyPrice, ".", ",", 1))
        
        Einhe = IfMed.DosageForm
        HerNr = IfMed.SupplierId
        
        Set IfHer = New praxisCENTER3.Supplier
        'Set IfHer = CreateObject("praxisCENTER3.Supplier")
        RetW2 = IfIdx.GetSupplierByID(HerNr, IfHer)
        If RetW2 = 0 Then
            HerNa(AktZa) = RTrim$(IfHer.Name)
        End If

        If Einhe Like "AMP" Or Einhe Like "ILO" Or Einhe Like "PUL" Or Einhe Like "INF" Then
            Menge(AktZa) = IfMed.Quantity
            Posit = InStr(1, Menge(AktZa), "X", 1)
            If Posit > 0 Then
                Anzal = CInt(Mid$(Menge(AktZa), 1, Posit - 1))
            Else
                Anzal = Val(Menge(AktZa))
            End If
            EiPre(AktZa) = MePre(AktZa) / Anzal
        Else
            EiPre(AktZa) = MePre(AktZa)
        End If
        LaCa1.Caption = MeBez(AktZa)

    Else
        LaCa1.Caption = vbNullString
    End If
    PrBr1.Value = AktZa
    If TxDum.Text = "B" Then Exit For 'Abbrechen
    DoEvents
Next AktZa

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FdAbL " & Err.Number
Resume Next

End Sub

Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    TxDa1.Text = NeuDa
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FDSav()
On Error GoTo DaErr
'Speichert die Daten in die Datenbank

Dim NeuDa As Date
Dim HerID As Long
Dim AktZa As Long
Dim ReNum As Long
Dim RzNum As Long
Dim PatNr As Long
Dim RowNr As Long
Dim RowFi As Long
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim SQL4 As String
Dim ReBet As Single
Dim ReRab As Single
Dim NeReB As Double
Dim CoStr As String
Dim KuStr As String
Dim TmStr As String
Dim MeTex As Variant
Dim NeuHe As Boolean
Dim RetWe As Boolean
Dim ReAbg As Boolean
Dim Mld1, Tit1 As String
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmIfap
Set TxDa1 = FM.txtDatu1
Set TxPNu = Me.txtPatNr
Set TxPVo = FM.txtPaVor
Set TxPNa = FM.txtPaNam
Set TxPGe = FM.txtPaGeb
Set TxRen = FM.txtReNum
Set RpCo1 = frmMain.repCont1
Set RpCo4 = frmMain.repCont4
Set RpCo3 = frmMain.repCont3
Set RpCo5 = frmMain.repCont5
Set RpCo8 = frmMain.repCont8
Set TxCoN = frmMain.TexCont1

If IsDate(TxDa1.Text) = True Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpRws = RpCo3.Rows
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rezeptmodul:
    Set RpRws = RpCo5.Rows
    Set RpCls = RpCo5.Columns
    Set RpSel = RpCo5.SelectedRows
Case RibTab_Kat_Eintrg:
    Set RpRws = RpCo8.Rows
    Set RpCls = RpCo8.Columns
    Set RpSel = RpCo8.SelectedRows
End Select

Select Case GlBut
Case RibTab_Abrechnung:
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Rec_ID1)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    Set RpCol = RpCls.Find(Rec_ID0)
                    PatNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Rec_ID1)
                    ReNum = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Rec_Betrag)
                    ReBet = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Rec_Selekt)
                    ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
                    Set RpCol = RpCls.Find(Rec_Rabatt)
                    ReRab = Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1)
                End If
            End If
        End If
Case RibTab_Rezeptmodul:
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Rzp_ID1)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    Set RpCol = RpCls.Find(Rzp_ID0)
                    PatNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Rzp_ID1)
                    RzNum = RpRow.Record(RpCol.ItemIndex).Value
                End If
            End If
        End If
Case RibTab_Kat_Eintrg:
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        RowNr = RpRow.Index
    End If
Case RibTab_Krankenbla:
    PatNr = GlAdr
Case RibTab_Tex_Rezept:
    PatNr = GlAdr
End Select

Select Case GlBut
Case RibTab_Tex_Rezept:
    For AktZa = 0 To GesZa - 1
        If TmStr = vbNullString Then
            If GlPzn = True Then
                TmStr = "(" & MePZN(AktZa) & ") " & MeBez(AktZa)
            Else
                TmStr = MeBez(AktZa)
            End If
        Else
            If GlPzn = True Then
                TmStr = TmStr & vbCrLf & "(" & MePZN(AktZa) & ") " & MeBez(AktZa)
            Else
                TmStr = TmStr & vbCrLf & MeBez(AktZa)
            End If
        End If
        If Dosag(AktZa) <> vbNullString Then
            TmStr = TmStr & Dosag(AktZa)
        End If
    Next AktZa
    With TxCoN
        .SelText = TmStr & vbCrLf
        .SelStart = Len(TxCoN.Text)
        .SelLength = 0
    End With
Case RibTab_Rezeptmodul:
    If GlTyp < 2 Then
        SQL1 = "SELECT * FROM dbo.qrySimRez WHERE ID1 = " & RzNum
    Else
        SQL1 = "SELECT * FROM qrySimRez WHERE [ID1] = " & RzNum & ";"
    End If
    Set RS01 = New ADODB.Recordset
    With RS01
        .CursorLocation = adUseClient
        .Source = SQL1
        .ActiveConnection = DB1
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open Options:=adCmdText
    End With
    If RS01.RecordCount > 0 Then
        If RS01.Fields("Rezepttext").Value <> vbNullString Then
            MeTex = RS01.Fields("Rezepttext").Value
        End If
        For AktZa = 0 To GesZa - 1
            If TmStr = vbNullString Then
                If GlPzn = True Then
                    TmStr = "(" & MePZN(AktZa) & ") " & MeBez(AktZa)
                Else
                    TmStr = MeBez(AktZa)
                End If
            Else
                If GlPzn = True Then
                    TmStr = TmStr & vbCrLf & "(" & MePZN(AktZa) & ") " & MeBez(AktZa)
                Else
                    TmStr = TmStr & vbCrLf & MeBez(AktZa)
                End If
            End If
            If Dosag(AktZa) <> vbNullString Then
                TmStr = TmStr & Dosag(AktZa)
            End If
        Next AktZa
        If MeTex = vbNullString Then
            MeTex = TmStr
        Else
            MeTex = MeTex & vbCrLf & TmStr
        End If
        If RS01.Supports(adUpdate) Then
            RS01.Fields("Rezepttext").Value = MeTex
            RS01.Update
        End If
    End If
    RS01.Close
    Set RS01 = Nothing
Case RibTab_Abrechnung:
    If ReAbg = False Then
        If GlTyp < 2 Then
            SQL4 = "SELECT * FROM dbo.qrySimAbSav WHERE IDR = " & ReNum
        Else
            SQL4 = "SELECT * FROM qrySimAbSav WHERE [IDR] = " & ReNum & ";"
        End If
        Set RS02 = New ADODB.Recordset
        With RS02
            .CursorLocation = adUseClient
            .Source = SQL4
            .ActiveConnection = DB1
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open Options:=adCmdText
        End With
        If RS02.Supports(adAddNew) Then
            For AktZa = 0 To GesZa - 1
                RS02.AddNew
                RS02.Fields("ID0").Value = PatNr
                RS02.Fields("IDR").Value = ReNum
                RS02.Fields("ID1").Value = 4
                RS02.Fields("Datum").Value = NeuDa
                RS02.Fields("GONr").Value = Format$(MePZN(AktZa), "0000000")
                RS02.Fields("IDKurz").Value = MeBez(AktZa) & " (PZN: " & Format$(MePZN(AktZa), "0000000") & ")"
                RS02.Fields("Multi").Value = 1
                RS02.Fields("x").Value = 1
                RS02.Fields("Betrag").Value = CSng(Round(EiPre(AktZa), 2))
                RS02.Fields("GesBetrag").Value = CSng(Round(EiPre(AktZa), 2))
                RS02.Fields("Währung").Value = 1
                RS02.Update
            Next AktZa
        End If
        RS02.Close
        Set RS02 = Nothing
        
        DBCmEx0 "qryWarAbPos" 'Übertragen [ID2] auf [ID3]
    Else
        Tit1 = "Rechnung Abgeschlossen"
        Mld1 = "Diese Rechnung wurde bereits abgeschlossen"
        WindowMess Mld1, Dial2, Tit1, FM.hwnd
    End If
Case RibTab_Krankenbla:
    For AktZa = 0 To GesZa - 1
        CoStr = Format$(MePZN(AktZa), "0000000")
        KuStr = MeBez(AktZa)
        DoEvents
        DBCmEx6 "qrySimAbDiA2", "@PatNr", "@IdxNr", "@IdCod", "@IdStr", "@IdGrp", "@IdRef", PatNr, 0, CoStr, KuStr, 0, 0
    Next AktZa
Case RibTab_Kat_Eintrg:
    If GlTyp < 2 Then
        SQL1 = "SELECT TOP 5 * FROM dbo.qryKat04C"
    Else
        SQL1 = "SELECT TOP 5 * FROM qryKat04C;"
    End If
    Set RS05 = New ADODB.Recordset
    With RS05
        .CursorLocation = adUseClient
        .Source = SQL1
        .ActiveConnection = DB1
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open Options:=adCmdText
    End With
    If RS05.Supports(adAddNew) Then
        For AktZa = 0 To GesZa - 1
            If GlTyp < 2 Then
                SQL2 = "SELECT * FROM dbo.qryKat04 WHERE IDKurz Like '" & HerNa(AktZa) & "'"
            Else
                SQL2 = "SELECT * FROM qryKat04 WHERE [IDKurz] Like '" & HerNa(AktZa) & "';"
            End If
            Set RS06 = New ADODB.Recordset
            With RS06 'Ist Hersteller schon vorhanden?
                .CursorLocation = adUseClient
                .Source = SQL2
                .ActiveConnection = DB1
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open Options:=adCmdText
            End With
            If RS06.RecordCount = 0 Then
                If RS06.Supports(adAddNew) Then
                    RS06.AddNew
                    RS06.Fields("IDKurz").Value = HerNa(AktZa)
                    RS06.Update
                    NeuHe = True
                End If
            Else
                HerID = RS06.Fields("ID3").Value
            End If
            RS06.Close
            Set RS06 = Nothing
            
            If NeuHe = True Then 'Neuer Hersteller wird angelegt
                Set RS07 = New ADODB.Recordset
                With RS07
                    .CursorLocation = adUseClient
                    .Source = SQL2
                    .ActiveConnection = DB1
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open Options:=adCmdText
                End With
                If RS07.RecordCount > 0 Then
                    HerID = RS07.Fields("ID3").Value
                Else
                    HerID = 0
                End If
                RS07.Close
                Set RS07 = Nothing
            End If
            DoEvents

            If HerID > 0 Then
                If GlTyp < 2 Then
                    SQL3 = "SELECT * FROM dbo.qryKat04C WHERE ID3=" & HerID & " AND GOID Like '" & MePZN(AktZa) & "'"
                Else
                    SQL3 = "SELECT * FROM qryKat04C WHERE [ID3]=" & HerID & " AND [GOID] Like '" & MePZN(AktZa) & "';"
                End If
                Set RS04 = New ADODB.Recordset
                With RS04
                    .CursorLocation = adUseClient
                    .Source = SQL3
                    .ActiveConnection = DB1
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open Options:=adCmdText
                End With
                If RS04.RecordCount = 0 Then
                    RS05.AddNew
                    RS05.Fields("ID3").Value = HerID
                    RS05.Fields("GOID").Value = Format$(MePZN(AktZa), "0000000")
                    RS05.Fields("IDKurz").Value = MeBez(AktZa)
                    RS05.Fields("Preis1").Value = CSng(Round(MePre(AktZa), 2))
                    RS05.Fields("Preis3").Value = CSng(Round(EiPre(AktZa), 2))
                    RS05.Fields("Multi").Value = 1
                    RS05.Update
                Else
                    RS04.Fields("GOID").Value = Format$(MePZN(AktZa), "0000000")
                    RS04.Fields("IDKurz").Value = MeBez(AktZa)
                    RS04.Fields("Preis1").Value = CSng(Round(MePre(AktZa), 2))
                    RS04.Fields("Preis3").Value = CSng(Round(EiPre(AktZa), 2))
                    RS04.Update
                End If
                RS04.Close
                Set RS04 = Nothing
                
                If GlTyp < 2 Then
                    SQL3 = "SELECT * FROM dbo.qryKat04C WHERE ID3=" & 1 & " AND GOID Like '" & MePZN(AktZa) & "'"
                Else
                    SQL3 = "SELECT * FROM qryKat04C WHERE [ID3]=" & 1 & " AND [GOID] Like '" & MePZN(AktZa) & "';"
                End If
                Set RS04 = New ADODB.Recordset
                With RS04
                    .CursorLocation = adUseClient
                    .Source = SQL3
                    .ActiveConnection = DB1
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open Options:=adCmdText
                End With
                If RS04.RecordCount = 0 Then
                    RS05.AddNew
                    RS05.Fields("ID3").Value = 1
                    RS05.Fields("GOID").Value = Format$(MePZN(AktZa), "0000000")
                    RS05.Fields("IDKurz").Value = MeBez(AktZa)
                    RS05.Fields("Preis1").Value = CSng(Round(MePre(AktZa), 2))
                    RS05.Fields("Preis3").Value = CSng(Round(EiPre(AktZa), 2))
                    RS05.Fields("Multi").Value = 1
                    RS05.Update
                Else
                    RS04.Fields("GOID").Value = Format$(MePZN(AktZa), "0000000")
                    RS04.Fields("IDKurz").Value = MeBez(AktZa)
                    RS04.Fields("Preis1").Value = CSng(Round(MePre(AktZa), 2))
                    RS04.Fields("Preis3").Value = CSng(Round(EiPre(AktZa), 2))
                    RS04.Update
                End If
                RS04.Close
                Set RS04 = Nothing
            End If
        Next AktZa
    End If
    RS05.Close
    Set RS05 = Nothing
End Select

Select Case GlBut
Case RibTab_Abrechnung:
        RetWe = S_KrPo()
        NeReB = S_ReBet(ReNum, Round(ReBet, 2), ReAbg, ReRab)
        RowFi = RpCo3.TopRowIndex
        S_KrLi
        If RowNr > 0 Then
            If RpRws.Count > 0 Then
                If RowNr >= RpRws.Count Then
                    RowNr = RpRws.Count - 1
                End If
                RpCo3.TopRowIndex = RowFi
                RpRws.Row(0).Selected = False
                RpRws.Row(RowNr).EnsureVisible
                RpRws.Row(RowNr).Selected = True
            End If
        End If
        Set RpSel = RpCo4.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpRe RowNr
        End If
Case RibTab_Rezeptmodul:
        SUpRz RowNr
Case RibTab_Kat_Eintrg:
        KList
Case RibTab_Krankenbla:
        RetWe = S_KrPk()
End Select

Set RpCo1 = Nothing
Set RpCo4 = Nothing
Set RpCo3 = Nothing
Set RpCo5 = Nothing
Set RpCo8 = Nothing

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDSav " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1
Set Rahm2 = Me.frmRahm2

If IsDate(TxDa1.Text) Then
    NeuDa = TxDa1.Text
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = Rahm2.Top + TxDa1.Top + TxDa1.Height
    .Left = Rahm2.Left + TxDa1.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TxDa1.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKaOp(ByVal KatNr As Long)
On Error GoTo DaErr
'Liest alle Einträge eines bestimmten Kataloges ein

Dim AktZa As Long

Set FM = frmIfap
Set LaCa1 = FM.lblLab01
Set PrBr1 = FM.prbStat1
Set PrBr2 = FM.prbStat2
Set TxDum = FM.txtDummy

Set RS08 = New ADODB.Recordset
RS08.CursorLocation = adUseClient
Set RS08 = DBCmRe1("qryKat04E", "@IdxNr", KatNr)
Set RS08.ActiveConnection = Nothing
If RS08.RecordCount > 0 Then
    RS08.MoveLast
    GesZa = RS08.RecordCount
    ReDim EiPZN(GesZa)
    PrBr1.Min = 0
    PrBr1.Max = GesZa
    DoEvents
    RS08.MoveFirst
    For AktZa = 0 To GesZa - 1
        If RS08.Fields("GOID").Value <> vbNullString Then
            EiPZN(AktZa) = RS08.Fields("GOID").Value
        Else
            EiPZN(AktZa) = "0"
        End If
        If RS08.Fields("IDKurz").Value <> vbNullString Then
            LaCa1.Caption = RS08.Fields("IDKurz").Value
        Else
            LaCa1.Caption = "keine Bezeichnung"
        End If
        PrBr1.Value = AktZa
        If TxDum.Text = "B" Then Exit For 'Abbrechen
        DoEvents
        RS08.MoveNext
    Next AktZa
End If
RS08.Close
Set RS08 = Nothing

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaOp " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim AktZa As Integer

Set FM = frmIfap
Set ChAbg = FM.chkIfAbg
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set TxDa1 = FM.txtDatu1
Set PuBu1 = FM.btnDatu1
Set MoKal = FM.dtpDatu1
Set Labl1 = FM.lblLabl1
Set Labl2 = FM.lblLabl2
Set Labl3 = FM.lblLabl3
Set Labl4 = FM.lblLabl4
Set Labl5 = FM.lblLabl5
Set Labl6 = FM.lblLabl6
Set Labl7 = FM.lblLabl7
Set PrBr1 = FM.prbStat1
Set PrBr2 = FM.prbStat2
Set CmSu6 = FM.cmbSuch6
Set CmSu7 = FM.cmbSuch7
Set ImMan = frmMain.imgManag

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
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Rechnungstag"
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

With PrBr1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .FlatStyle = False
    .Scrolling = xtpProgressBarStandard
    .UseVisualStyle = True
End With

With PrBr2
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .FlatStyle = False
    .Scrolling = xtpProgressBarStandard
    .UseVisualStyle = True
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Date, "dd.mm.yyyy")
End With

If GlDiV > 0 Then
    For AktZa = 0 To GlDiV - 1
        With CmSu6
            .AddItem DiAry(4, AktZa) & " " & DiAry(6, AktZa)
        End With
    Next AktZa
End If

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Select Case GlBut
Case RibTab_Abrechnung:
    PuBu1.Visible = True
    TxDa1.Visible = True
    Labl3.Visible = True
    ChAbg.Visible = False
Case RibTab_Kat_Eintrg:
    PuBu1.Visible = False
    TxDa1.Visible = False
    Labl3.Visible = False
    ChAbg.Visible = Not GlRDP
Case Else:
    PuBu1.Visible = False
    TxDa1.Visible = False
    Labl3.Visible = False
    ChAbg.Visible = False
End Select

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
ChAbg.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FRes()
On Error Resume Next

Set TxSu1 = Me.txtSuch1
Set TxSu2 = Me.txtSuch2
Set TxSu3 = Me.txtSuch3
Set TxSu4 = Me.txtSuch4
Set TxSu5 = Me.txtSuch5
Set CmSu6 = Me.cmbSuch6
Set CmSu7 = Me.cmbSuch7

TxSu1.Text = vbNullString
TxSu2.Text = vbNullString
TxSu3.Text = vbNullString
TxSu4.Text = vbNullString
TxSu5.Text = vbNullString
CmSu6.Text = vbNullString
CmSu7.Text = vbNullString

End Sub
Public Sub FSuch(Optional ByVal HauLi As Boolean = False)
On Error GoTo LaErr

Dim RetWe As Long
Dim ManNr As Long
Dim PatNr As String
Dim PaVor As String
Dim PaNam As String
Dim PaGeb As String
Dim SuStr As String
Dim LanNr As String
Dim ManNa As String
Dim ManPl As String
Dim ManVo As String
Dim MitNa As String
Dim AktZa As Integer
Dim Posit As Integer

Set FM = frmIfap
Set TxSu1 = FM.txtSuch1
Set TxSu2 = FM.txtSuch2
Set TxSu3 = FM.txtSuch3
Set TxSu4 = FM.txtSuch4
Set TxSu5 = FM.txtSuch5
Set TxPNu = FM.txtPatNr
Set TxPVo = FM.txtPaVor
Set TxPNa = FM.txtPaNam
Set TxPGe = FM.txtPaGeb
Set TxRen = FM.txtReNum
Set ChAbg = FM.chkIfAbg
Set CmSu6 = FM.cmbSuch6
Set CmSu7 = FM.cmbSuch7

If TxPNu.Text <> vbNullString Then
    PatNr = "P" & Format$(TxPNu.Text, "000000")
Else
    PatNr = "P000000"
End If
If TxPVo.Text <> vbNullString Then
    PaVor = TxPVo.Text
Else
    PaVor = vbNullString
End If
If TxPNa.Text <> vbNullString Then
    PaNam = TxPNa.Text
Else
    PaNam = vbNullString
End If
If TxPGe.Text <> vbNullString Then
    PaGeb = CDate(TxPGe.Text)
Else
    PaGeb = vbNullString
End If

If GlMiV = True Then 'Mitarbeiter vorhanden
    ManNr = GlMiA(GlSmI, 7) 'zugeordnete Mandantennummer
    MitNa = GlMiA(GlSmI, 1)
Else
    ManNr = GlMan(GlSMa, 2) 'Standardmandant
    MitNa = vbNullString
End If

For AktZa = 1 To UBound(GlThe) 'Mandanten
    If ManNr = GlThe(AktZa, 0) Then
        LanNr = GlThe(AktZa, 7) 'Lebenslange Arztnummer
        ManVo = GlThe(AktZa, 1)
        ManNa = GlThe(AktZa, 2)
        ManPl = GlThe(AktZa, 4)
        Exit For
    End If
Next AktZa

If ChAbg.Value = xtpChecked Then
    FDAbg
Else
    'Set IfPat = New praxisCENTER3.patient
    'Set IfPts = New praxisCENTER3.Patients
    'Set IfDis = New praxisCENTER3.Diagnoses
    'Set IfMes = New praxisCENTER3.Medicaments
    'Set IfPhy = New praxisCENTER3.Physician
    'Set IfIdx = New praxisCENTER3.ApplicationObject
    
    Set IfPat = CreateObject("praxisCENTER3.Patient")
    Set IfPts = CreateObject("praxisCENTER3.Patients")
    Set IfDis = CreateObject("praxisCENTER3.Diagnoses")
    Set IfMes = CreateObject("praxisCENTER3.Medicaments")
    Set IfPhy = CreateObject("praxisCENTER3.Physician")
    Set IfIdx = CreateObject("praxisCENTER3.ApplicationObject")
    
    RetWe = IfIdx.SetLicenceKey("689149F60BD11AF36A8F8E86A3556886575B7EFD")
    RetWe = IfIdx.SetMode("MaxMedicamentsInPrescription", 0)
    RetWe = IfIdx.SetMode("UpdateAfterOnPrescription", 1)
    RetWe = IfIdx.SetMode("DisableStartSortiment", 0)
    RetWe = IfIdx.SetMode("ProductHouselist", 1)
    RetWe = IfIdx.SetMode("PrescriptionType", 0)
    RetWe = IfIdx.SetMode("MultiplePatient", 1) 'Mehrere Patienten gleichzeitig
    RetWe = IfIdx.SetMode("MinimizeToTray", 0)
    RetWe = IfIdx.SetMode("EnableOnEvent", 0)
    RetWe = IfIdx.SetMode("SmartClose", 0)
    RetWe = IfIdx.SetMode("UserRole", 0) 'Kein Arzt
    
    If GlRDP = True Then
        RetWe = IfIdx.SetMode("CanModifyHausliste", 0)
        RetWe = IfIdx.SetMode("AddToHouselistOnPrescription", 0)
        RetWe = IfIdx.SetMode("HauslisteDisabled", 1)
    Else
        RetWe = IfIdx.SelectDepartment(MitNa)
    End If

    IfIdx.Activate
    
    If GlRDP = True Then
        IfIdx.Physician.LANR = "999934"
        IfIdx.Physician.LastName = "SimpliMed"
        IfIdx.Physician.Activate
    Else
        IfIdx.Physician.LANR = LanNr
        IfIdx.Physician.FirstName = ManVo
        IfIdx.Physician.LastName = ManNa
        IfIdx.Physician.PostalCode = ManPl
        IfIdx.Physician.Activate
    End If

    With IfPat
        .id = PatNr
        .FirstName = PaVor
        .LastName = PaNam
        .Birthday = PaGeb
        IfIdx.Patients.Add IfPat
        .Activate
    End With

    If GlDiV > 0 Then
        For AktZa = 0 To GlDiV - 1
            'Set IfDia = New praxisCENTER3.Diagnose
            Set IfDia = CreateObject("praxisCENTER3.Diagnose")
            IfDia.ICD = DiAry(4, AktZa)
            IfDis.Add IfDia
        Next AktZa
    End If
    
    If GlMeV > 0 Then
        For AktZa = 0 To GlMeV - 1
            'Set IfMed = New praxisCENTER3.Medicament
            Set IfMed = CreateObject("praxisCENTER3.Medicament")
            IfMed.PIC = MeAry(4, AktZa)
            IfMes.Add IfMed
        Next AktZa
    End If

    If HauLi = True Then
        RetWe = IfIdx.ShowRange(7)
    Else
        With IfIdx
            If TxSu1.Text <> vbNullString Then 'Arzneimittelname
                SuStr = TxSu1.Text
                RetWe = .ShowMedicament(SuStr)
            ElseIf TxSu2.Text <> vbNullString Then 'PZN
                SuStr = TxSu2.Text
                RetWe = .ShowMedicament(SuStr)
            ElseIf TxSu3.Text <> vbNullString Then 'Hersteller
                SuStr = TxSu3.Text
                RetWe = .ShowMedicamentsBySearchType(700, SuStr)
            ElseIf TxSu4.Text <> vbNullString Then 'Wirkstoff
                SuStr = TxSu4.Text
                RetWe = .ShowMedicamentsBySearchType(300, SuStr)
            ElseIf TxSu5.Text <> vbNullString Then 'PZN Schnellauskunft
                SuStr = TxSu5.Text
                RetWe = .ShowLibrary(CLng(SuStr), 100)
            ElseIf CmSu6.Text <> vbNullString Then 'ICD-10 Name
                Posit = InStr(1, CmSu6.Text, Chr$(32), 1) 'Leerzeichen
                If Posit > 0 Then
                    SuStr = Left$(CmSu6.Text, Posit - 1)
                Else
                    SuStr = CmSu6.Text
                End If
                RetWe = .ShowMedicamentsBySearchType(400, SuStr)
            ElseIf CmSu7.Text <> vbNullString Then 'ATC Code
                Posit = InStr(1, CmSu7.Text, "(", 1)
                If Posit > 0 Then
                    SuStr = Left$(CmSu7.Text, Posit - 2)
                Else
                    SuStr = CmSu7.Text
                End If
                RetWe = .ShowMedicamentsBySearchType(500, SuStr)
            End If
        End With
    End If
End If

If GlDiV > 0 Then
    Erase DiAry
    GlDiV = 0
End If

If GlMeV > 0 Then
    Erase MeAry
    GlMeV = 0
End If

DoEvents
GlNeK = GlKoX 'Protokolleintrag
With GlNeK
    .PatNr = GlMan(GlSMa, 2)
    .IdxNr = 0
    .EiDat = Format$(Date, "dd.mm.yyyy")
    .EiZei = TimeValue(Now)
    .EiTyp = 104
    .TeStr = "ifap3 aufgerufen"
    .ZiStr = Format$(Now, "hh:mm") & " Uhr"
    .NeuEi = True
    .KeiAk = True
    .Mitar = GlMiA(GlSmI, 2)
End With
S_Prot

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuch " & Err.Number
Resume Next

End Sub
Private Sub FUber(ByVal PaStr As String)
On Error GoTo LaErr

Dim RetW1 As Long
Dim RetW2 As Long
Dim HerNr As Long
Dim TmpSt As String
Dim Einhe As String
Dim MeTex As String
Dim DosTm As String
Dim AktZa As Integer
Dim Posit As Integer
Dim Anzal As Integer
Dim Menge() As String

GesZa = IfPat.Prescription.Count

If GesZa > 0 Then
    ReDim Preserve MePZN(GesZa)
    ReDim Preserve MeBez(GesZa)
    ReDim Preserve MePre(GesZa)
    ReDim Preserve EiPre(GesZa)
    ReDim Preserve HerNa(GesZa)
    ReDim Preserve Menge(GesZa)
    ReDim Preserve Dosag(GesZa)
        
    For Each IfRzM In IfPat.Prescription
        TmpSt = IfRzM.PIC
        DosTm = vbNullString
        
        'Set IfMed = New praxisCENTER3.Medicament
        Set IfMed = CreateObject("praxisCENTER3.Medicament")
        RetW1 = IfIdx.GetMedicament(TmpSt, IfMed)

        If RetW1 = 0 Then
            MePZN(AktZa) = IfRzM.PIC
            MeBez(AktZa) = IfRzM.PrintName
            
            If CLng(IfRzM.Dosage.Morning) > 0 Then DosTm = vbCrLf & "Morgens: " & IfRzM.Dosage.Morning
            If CLng(IfRzM.Dosage.Midday) > 0 Then DosTm = DosTm & vbCrLf & "Mittags: " & IfRzM.Dosage.Midday
            If CLng(IfRzM.Dosage.Evening) > 0 Then DosTm = DosTm & vbCrLf & "Abends: " & IfRzM.Dosage.Evening
            If CLng(IfRzM.Dosage.Night) > 0 Then DosTm = DosTm & vbCrLf & "Nachts: " & IfRzM.Dosage.Night
            Dosag(AktZa) = DosTm
            
            Einhe = IfMed.DosageForm
            HerNr = IfMed.SupplierId
            
            MePre(AktZa) = CSng(Replace(IfMed.PharmacyPrice, ".", ",", 1))

            'Set IfHer = New praxisCENTER3.Supplier
            Set IfHer = CreateObject("praxisCENTER3.Supplier")
            RetW2 = IfIdx.GetSupplierByID(HerNr, IfHer)
            If RetW2 = 0 Then
                HerNa(AktZa) = RTrim$(IfHer.Name)
            End If
                                    
            If Einhe Like "AMP" Or Einhe Like "ILO" Or Einhe Like "PUL" Or Einhe Like "INF" Then
                Menge(AktZa) = IfMed.Quantity
                Posit = InStr(1, Menge(AktZa), "X", 1)
                If Posit > 0 Then
                    Anzal = CInt(Mid$(Menge(AktZa), 1, Posit - 1))
                Else
                    Anzal = Val(Menge(AktZa))
                End If
                EiPre(AktZa) = MePre(AktZa) / Anzal
            Else
                EiPre(AktZa) = MePre(AktZa)
            End If

        End If
        AktZa = AktZa + 1
    Next
End If

IfPat.Deactivate

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FUber " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    FKale
End Sub
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
Private Sub btnWeiter_Click()
    FSuch
End Sub

Private Sub btnZurück_Click()
    FSuch True
End Sub
Private Sub chkIfAbg_Click()

Set ChAbg = Me.chkIfAbg
Set TxSu1 = Me.txtSuch1
Set TxSu2 = Me.txtSuch2
    
If ChAbg.Value = xtpChecked Then
    TxSu1.Enabled = False
    TxSu2.Enabled = False
Else
    TxSu1.Enabled = True
    TxSu2.Enabled = True
End If

End Sub

Private Sub cmbSuch6_GotFocus()
    FRes
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd
S_ATC

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmIfap = Nothing
End Sub
Private Sub IfIdx_OnPrescription(ByVal bstrPatientID As String)
On Error Resume Next

FUber bstrPatientID
DoEvents

FDSav
DoEvents

If GlRDP = True Then
    IfIdx.Close
Else
    IfIdx.Hide
End If

Set IfIdx = Nothing
DoEvents

Unload Me

End Sub

Private Sub txtSuch1_GotFocus()
    FRes
End Sub
Private Sub txtSuch2_GotFocus()
    FRes
End Sub

Private Sub txtSuch3_GotFocus()
    FRes
End Sub
Private Sub txtSuch4_GotFocus()
    FRes
End Sub

Private Sub txtSuch5_GotFocus()
    FRes
End Sub
