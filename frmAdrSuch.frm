VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmAdrSuch 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Suchen"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   6
      Top             =   3200
      Width           =   7000
      _Version        =   1048579
      _ExtentX        =   12347
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   5000
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
         Left            =   3600
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
      Begin XtremeSuiteControls.PushButton btnZurück 
         Height          =   400
         Left            =   2200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zurück"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3000
      Left            =   600
      TabIndex        =   0
      Top             =   100
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   5292
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtBemer 
         Height          =   350
         Left            =   1000
         TabIndex        =   5
         Top             =   2430
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   350
         Left            =   1000
         TabIndex        =   1
         Top             =   330
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtNumme 
         Height          =   350
         Left            =   3100
         TabIndex        =   3
         Top             =   1030
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtGebor 
         Height          =   350
         Left            =   1000
         TabIndex        =   2
         Top             =   1030
         Width           =   1900
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPost 
         Height          =   350
         Left            =   1000
         TabIndex        =   4
         Top             =   1730
         Width           =   1900
         _Version        =   1048579
         _ExtentX        =   3351
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.PushButton btnPictu 
         Height          =   550
         Left            =   200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   200
         Width           =   550
         _Version        =   1048579
         _ExtentX        =   970
         _ExtentY        =   970
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   6
         DrawFocusRect   =   0   'False
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   210
         Left            =   1000
         TabIndex        =   18
         Top             =   2180
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Suche nach Bemerkung :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   210
         Left            =   1000
         TabIndex        =   17
         Top             =   1480
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Suche nach Postleitzahl :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   210
         Left            =   3100
         TabIndex        =   16
         Top             =   780
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Suche nach PIN :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   210
         Left            =   1000
         TabIndex        =   15
         Top             =   780
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Suche nach Geburtsdatum :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   210
         Left            =   1000
         TabIndex        =   14
         Top             =   80
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Suche nach Patientenname :"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3000
      Left            =   600
      TabIndex        =   10
      Top             =   100
      Visible         =   0   'False
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   5292
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ListView lstView3 
         Height          =   2450
         Left            =   200
         TabIndex        =   11
         Top             =   340
         Width           =   5300
         _Version        =   1048579
         _ExtentX        =   9349
         _ExtentY        =   4322
         _StockProps     =   77
         BackColor       =   -2147483643
         View            =   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie einen der gefundenen Einträge :"
         Height          =   200
         Left            =   210
         TabIndex        =   12
         Top             =   100
         Width           =   3600
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
Attribute VB_Name = "frmAdrSuch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Pict1 As VB.PictureBox
Private Rahm0 As XtremeSuiteControls.GroupBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private FTex3 As XtremeSuiteControls.FlatEdit
Private FTex4 As XtremeSuiteControls.FlatEdit
Private FTex5 As XtremeSuiteControls.FlatEdit
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private LiVw3 As XtremeSuiteControls.ListView
Private LiItm As XtremeSuiteControls.ListViewItem
Private LiIts As XtremeSuiteControls.ListViewItems
Private CmBar As XtremeCommandBars.CommandBar
Private CmSta As XtremeCommandBars.StatusBar
Private CmAcs As XtremeCommandBars.CommandBarActions
Private ImMan As XtremeCommandBars.ImageManager
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private TxCoN As Tx4oleLib.TXTextControl
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private TagWe As String
Private clFen As clsFenster
Private clFil As clsFile

Public FiNam As String
Private Sub FInit()
On Error Resume Next

Set FM = frmMain
Set ImMan = FM.imgManag

Set Rahm0 = Me.frmRahm0
Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumme
Set FTex3 = Me.txtPost
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor
Set LiVw3 = Me.lstView3
Set PuBu1 = Me.btnPictu
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2

With LiVw3
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .AllowColumnReorder = True
    .Arrange = xtpListViewArrangeAutoLeft
    .Checkboxes = False
    .FlatScrollBar = False
    .Font.SIZE = 10 'GlTFt.size
    .Font.Name = GlTFt.Name
    .ForeColor = vbBlack
    .FullRowSelect = True
    .GridLines = False
    .HideColumnHeaders = False
    .HideSelection = False
    .HotTracking = False
    .HoverSelection = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpListViewLabelManual
    .LabelWrap = True
    .MultiSelect = False
    .OLEDropMode = xtpOLEDropNone
    .View = xtpListViewReport
End With

FTex2.Pattern = "\d*"
FTex3.Pattern = "\d*"
FTex5.SetMask "00.00.0000", "__.__.____"

PuBu1.Icon = ImMan.Icons.GetImage(IC32_View, 32)

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
PuBu1.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak

With LiVw3
    .ColumnHeaders.Add 1, , "Adresse", 3000
    .ColumnHeaders.Add 2, , "Mandant", 1900
    .ColumnHeaders.Add 3, , "Email", 0
    .ColumnHeaders.Add 4, , "Briefanrede", 0
End With

Set ImMan = Nothing

End Sub
Private Sub FMail()
On Error GoTo SeErr

Dim IdxNr As Long
Dim IdStr As String
Dim EmStr As String
Dim EmaNr As String
Dim SelTx As String
Dim SigNa As String
Dim SigDa As String
Dim SigVo As Boolean

Set FTex1 = Me.txtKurz
Set LiVw3 = Me.lstView3
Set LiIts = LiVw3.ListItems

For Each LiItm In LiIts
    If LiItm.Selected = True Then
        IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
        IdStr = LiItm.Text
        If LiItm.SubItems(2) <> vbNullString Then
            EmStr = LiItm.SubItems(2) 'Emailadresse
        Else
            EmStr = vbNullString
        End If
        If LiItm.SubItems(3) <> vbNullString Then
            EmaNr = LiItm.SubItems(3) 'Briefanrede
        Else
            EmaNr = vbNullString
        End If
        Exit For
    End If
Next LiItm
If EmStr <> vbNullString Then
    Set FM = frmMaiView
    Set TxCoN = FM.TexCont3
    Select Case GlMaE 'Emailempfänger
    Case 1:
        If GlMiA(GlSmI, 25) <> vbNullString Then 'Signaturen
            SigDa = GlVor & GlMiA(GlSmI, 25)
            Set clFil = New clsFile
            If clFil.FilVor(SigDa) = True Then
                If Right$(LCase(SigDa), 3) = "txn" Then
                    SigVo = True
                End If
            End If
            Set clFil = Nothing
            If SigVo = False Then
                If GlMiA(GlSmI, 11) <> vbNullString Then
                    SigNa = GlMiA(GlSmI, 11)
                Else
                    SigNa = GlMiA(GlSmI, 1)
                End If
            End If
        ElseIf GlMiA(GlSmI, 11) <> vbNullString Then
            SigNa = GlMiA(GlSmI, 11)
        Else
            SigNa = GlMiA(GlSmI, 1)
        End If
        If GlNaT < 4 Then 'Mailtyp (1=View 2=Neu 3=Antwort)
            SelTx = vbCrLf & EmaNr & vbCrLf & vbCrLf & vbCrLf
            If SigVo = True Then
                With TxCoN
                    .ResetContents
                    .ForeColor = vbBlack
                    .LoadFromMemory SelTx, 1, True
                    If Right$(LCase(SigDa), 3) = "txn" Then
                        .Append SigDa, 0, 3
                    End If
                End With
            Else
                With TxCoN
                    .ResetContents
                    .ForeColor = vbBlack
                    .LoadFromMemory SelTx, 1, True
                    SelTx = vbCrLf & SigNa & vbCrLf & vbCrLf
                    .FontItalic = 1
                    .ForeColor = 8404992
                    .LoadFromMemory SelTx, 1, True
                End With
            End If
        End If
        If FM.cmbEmEmp.Text <> vbNullString Then
            If FM.txtEmCCM.Text <> vbNullString Then
                FM.txtEmCCM.Text = FM.txtEmCCM.Text & ";" & EmStr
            Else
                FM.txtEmCCM.Text = EmStr
            End If
        Else
            FM.cmbEmEmp.Text = EmStr
        End If
    Case 2: FM.txtEmCCM.Text = EmStr
    Case 3: FM.cmbEmBCC.Text = EmStr
    Case Else: S_MaMa 5, IdxNr, IdStr
    End Select
End If

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMail " & Err.Number
Resume Next

End Sub
Private Sub TRes()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtPost
Set FTex3 = Me.txtNumme
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor

FTex1.Text = vbNullString
FTex2.Text = vbNullString
FTex3.Text = vbNullString
FTex4.Text = vbNullString
FTex5.Text = vbNullString

End Sub
Private Sub FSett()
On Error GoTo SeErr
'Lädt die ausgewählte Adresse in das Adressformular

Dim IdxNr As Long
Dim BerNr As Long
Dim IdStr As String
Dim EmStr As String
Dim EmaNr As String
Dim PaStr As String
Dim SelTx As String
Dim SigNa As String
Dim SigDa As String
Dim SigVo As Boolean
Dim GesZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl

Set FTex1 = Me.txtKurz
Set LiVw3 = Me.lstView3
Set LiIts = LiVw3.ListItems

GesZa = LiVw3.ListItems.Count

If GesZa > 0 Then
    If WindowLoad("frmWieder") = True Then
        For Each LiItm In LiIts
            If LiItm.Selected = True Then
                IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                IdStr = LiItm.Text
                Exit For
            End If
        Next LiItm
        Set FM = frmWieder
        FM.txtPatie.Text = IdStr
        FM.txtID0.Text = IdxNr
        FM.txtPatie.Tag = 1 & "Patient"
        FM.txtID0.Tag = 1 & "ID0"
        If FM.txtAnlaß.Text = vbNullString Then
            FM.txtAnlaß.Text = "wv"
            FM.txtAnlaß.Tag = 1 & "IDKurz"
        End If
    ElseIf WindowLoad("frmMandant") = True Then
        For Each LiItm In LiIts
            If LiItm.Selected = True Then
                GlMId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                IdStr = LiItm.Text
                Exit For
            End If
        Next LiItm
        GlAdL = True
        ASper True, True
        MNeu
        Man_Lad
        GlAdS = False
        GlAdL = False
    Else
        Select Case GlBut
        Case RibTab_Startseite:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
        Case RibTab_Adressen:
            If WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            End If
        Case RibTab_Mandanten:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            GlAdL = True
            ASper True
            ANeue
            Adr_Lad
            Kon_Lis
            GlAdS = False
            GlAdL = False
        Case RibTab_Verordner:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            GlAdL = True
            ASper True
            ANeue
            Adr_Lad
            Kon_Lis
            GlAdS = False
            GlAdL = False
        Case RibTab_Mitarbeit:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            GlAdL = True
            ASper True
            ANeue
            Adr_Lad
            Kon_Lis
            GlAdS = False
            GlAdL = False
        Case RibTab_Mahnwesen:
            Set FM = frmOPEdit
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    FM.txtPatie.Text = LiItm.Text
                    Exit For
                End If
            Next LiItm
        Case RibTab_Fragebogen:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            ElseIf WindowLoad("frmZuord") = True Then
                Set FM = frmZuord
                Set RpCo1 = FM.repCont1
                Set RpSel = RpCo1.SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    For Each LiItm In LiIts
                        If LiItm.Selected = True Then
                            IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                            IdStr = LiItm.Text
                            Exit For
                        End If
                    Next LiItm
                    RpRow.Record(16).Value = IdStr
                    RpRow.Record(17).Value = IdxNr
                End If
                Set RpSel = Nothing
                Set RpCo1 = Nothing
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Tagesproto:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Krankenbla:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Abrechnung:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Rezeptmodul:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuP
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Belegmodul:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuP
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Bildmodul:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuP
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Ter_Kalend:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            If WindowLoad("frmKetten") = True Then
                Set FM = frmKetten
                Set CmBrs = FM.comBar02
                Set CmSta = CmBrs.StatusBar
                FM.txtPatNr.Text = IdxNr
                CmSta.Pane(1).Text = IdStr
            ElseIf WindowLoad("frmTermin") = True Then
                TeAdr IdxNr, IdStr
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                TeAdr IdxNr, IdStr
            End If
        Case RibTab_Ter_Raeume:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            If WindowLoad("frmKetten") = True Then
                Set FM = frmKetten
                Set CmBrs = FM.comBar02
                Set CmSta = CmBrs.StatusBar
                FM.txtPatNr.Text = IdxNr
                CmSta.Pane(1).Text = IdStr
            ElseIf WindowLoad("frmTermin") = True Then
                TeAdr IdxNr, IdStr
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                TeAdr IdxNr, IdStr
            End If
        Case RibTab_Ter_Mitarb:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            If WindowLoad("frmKetten") = True Then
                Set FM = frmKetten
                Set CmBrs = FM.comBar02
                Set CmSta = CmBrs.StatusBar
                FM.txtPatNr.Text = IdxNr
                CmSta.Pane(1).Text = IdStr
            ElseIf WindowLoad("frmTermin") = True Then
                TeAdr IdxNr, IdStr
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                TeAdr IdxNr, IdStr
            End If
        Case RibTab_Ter_Listen:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            If WindowLoad("frmTermin") = True Then
                TeAdr IdxNr, IdStr
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            End If
        Case RibTab_Ter_Akont:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            If WindowLoad("frmTermin") = True Then
                TeAdr IdxNr, IdStr
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            End If
        Case RibTab_Ter_Warte:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            If WindowLoad("frmTermin") = True Then
                TeAdr IdxNr, IdStr
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            End If
        Case RibTab_LabBericht:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            Else
                If WindowLoad("frmZuord") = True Then
                    Set FM = frmZuord
                    Set RpCo1 = FM.repCont1
                    Set RpSel = RpCo1.SelectedRows
                    If RpSel.Count > 0 Then
                        Set RpRow = RpSel(0)
                        BerNr = RpRow.Record(0).Value
                        For Each LiItm In LiIts
                            If LiItm.Selected = True Then
                                IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                                IdStr = LiItm.Text
                                Exit For
                            End If
                        Next LiItm
                        T_FeZ IdxNr, BerNr
                        T_FeB
                    End If
                    Set RpSel = Nothing
                    Set RpCo1 = Nothing
                ElseIf WindowLoad("frmMaiView") = True Then
                    FMail
                Else
                    For Each LiItm In LiIts
                        If LiItm.Selected = True Then
                            GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                            IdStr = LiItm.Text
                            Exit For
                        End If
                    Next LiItm
                    With GlSuP
                        .SuIdx = 1
                        .SuNum = GlAdr
                    End With
                    SSuch
                End If
            End If
        Case RibTab_LabBerichte:
            If WindowLoad("frmZuord") = True Then
                Set FM = frmZuord
                Set RpCo1 = FM.repCont1
                Set RpSel = RpCo1.SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    BerNr = RpRow.Record(0).Value
                    For Each LiItm In LiIts
                        If LiItm.Selected = True Then
                            IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                            IdStr = LiItm.Text
                            Exit For
                        End If
                    Next LiItm
                    T_FeZ IdxNr, BerNr
                    T_FeB
                End If
                Set RpSel = Nothing
                Set RpCo1 = Nothing
            End If
        Case RibTab_LabAuftrag:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuP
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Tex_Dokumt:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuV
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                With GlSuP
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Tex_Vorlag:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuV
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                With GlSuP
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Tex_Rezept:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            ElseIf WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAdr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                With GlSuV
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                With GlSuA
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                With GlSuP
                    .SuIdx = 1
                    .SuNum = GlAdr
                End With
                SSuch
            End If
        Case RibTab_Tex_NewsLe:
            If WindowLoad("frmAdress") = True Then
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        Exit For
                    End If
                Next LiItm
                GlAdL = True
                ASper True
                ANeue
                Adr_Lad
                Kon_Lis
                GlAdS = False
                GlAdL = False
            End If
        Case RibTab_Tex_Email:
            If WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        If LiItm.SubItems(2) <> vbNullString Then
                            EmStr = LiItm.SubItems(2) 'Emailadresse
                        Else
                            EmStr = vbNullString
                        End If
                        If LiItm.SubItems(3) <> vbNullString Then
                            EmaNr = LiItm.SubItems(3) 'Briefanrede
                        Else
                            EmaNr = vbNullString
                        End If
                        Exit For
                    End If
                Next LiItm
                S_MaMa 5, IdxNr, IdStr
            End If
        Case RibTab_Kat_Explor:
             For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            If FiNam <> vbNullString Then
                SFilIm FiNam, IdxNr
            End If
        Case RibTab_Rechnungen:
            If WindowLoad("frmMaiView") = True Then
                FMail
            Else
                For Each LiItm In LiIts
                    If LiItm.Selected = True Then
                        IdxNr = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                        IdStr = LiItm.Text
                        If LiItm.SubItems(2) <> vbNullString Then
                            EmStr = LiItm.SubItems(2) 'Emailadresse
                        Else
                            EmStr = vbNullString
                        End If
                        If LiItm.SubItems(3) <> vbNullString Then
                            EmaNr = LiItm.SubItems(3) 'Briefanrede
                        Else
                            EmaNr = vbNullString
                        End If
                        Exit For
                    End If
                Next LiItm
                S_MaMa 5, IdxNr, IdStr
            End If
        Case Else:
            For Each LiItm In LiIts
                If LiItm.Selected = True Then
                    GlAId = CLng(Mid$(LiItm.Key, 2, Len(LiItm.Key) - 1))
                    IdStr = LiItm.Text
                    Exit For
                End If
            Next LiItm
            IdStr = FTex1.Text
            GlAdL = True
            ASper True
            ANeue
            Adr_Lad
            Kon_Lis
            GlAdS = False
            GlAdL = False
        End Select
    End If

    Select Case GlAdU 'Adresssuch Dialog Option
    Case 1:
        Unload Me
        DoEvents
        GlAdU = 0
        Select Case GlAdO
        Case 0: SReZe GlAId
        Case 1: SKrZe GlAId
        Case 2: SKrZe GlAId
        End Select
    Case 2:
        GlAdU = 0
        SKrZe GlAId
        DoEvents
        Unload Me
        DoEvents
        KrMain 22
    Case 3:
        GlAdU = 0
        GlAdr = GlAId
        S_KrLa
        DoEvents
        Unload Me
        DoEvents
        STxRz 1
    Case Else:
        DoEvents
        Unload Me
    End Select
End If

GlTDa = vbNullString 'Wichtig für Textverarbeitung

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

End Sub
Private Sub FSuda()
On Error GoTo SeErr

Dim GesZa As Long
Dim Mld1, Tit1 As String

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtNumme
Set FTex3 = Me.txtPost
Set FTex4 = Me.txtBemer
Set FTex5 = Me.txtGebor
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set LiVw3 = Me.lstView3
Set LiIts = LiVw3.ListItems

If FTex1.Text <> vbNullString Then
    GesZa = Adr_Fil(1, RTrim$(FTex1.Text), 1)
    If GesZa = 0 Then
        GesZa = Adr_Fil(1, SUmw(RTrim$(FTex1.Text)), 1)
    End If
ElseIf RTrim$(FTex2.Text) <> vbNullString Then
    GesZa = Adr_Fil(1, RTrim$(FTex2.Text), 2)
ElseIf RTrim$(FTex3.Text) <> vbNullString Then
    GesZa = Adr_Fil(1, RTrim$(FTex3.Text), 3)
ElseIf RTrim$(FTex4.Text) <> vbNullString Then
    GesZa = Adr_Fil(1, RTrim$(FTex4.Text), 4)
ElseIf RTrim$(FTex5.Text) <> vbNullString Then
    If IsDate(FTex5.Text) = True Then
        GesZa = Adr_Fil(1, vbNullString, 5, RTrim$(FTex5.Text))
    Else
        GesZa = 0
    End If
End If

If GesZa > 0 Then
    Rahm1.Visible = False
    Rahm2.Visible = True
    LiVw3.SetFocus
    LiIts(1).Selected = True
Else
    If FTex1.Text <> vbNullString Then
        FTex1.SelStart = 0
        FTex1.SelLength = Len(FTex1.Text)
    ElseIf FTex2.Text <> vbNullString Then
        FTex2.SelStart = 0
        FTex2.SelLength = Len(FTex2.Text)
    ElseIf FTex3.Text <> vbNullString Then
        FTex3.SelStart = 0
        FTex3.SelLength = Len(FTex3.Text)
    ElseIf FTex4.Text <> vbNullString Then
        FTex4.SelStart = 0
        FTex4.SelLength = Len(FTex4.Text)
    ElseIf FTex5.Text <> vbNullString Then
        FTex5.SelStart = 0
        FTex5.SelLength = Len(FTex5.Text)
    End If
    SPopu "Patient nicht gefunden", "Der von Ihnen gesuchte Patient, konnte nicht gefunden werden", IC48_Forbidden
End If

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuda " & Err.Number
Resume Next

End Sub
Private Sub FZur()
On Error Resume Next

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2

If Rahm2.Visible = True Then
    Rahm2.Visible = False
    Rahm1.Visible = True
End If

End Sub
Private Sub btnPictu_Click()

Set Rahm1 = Me.frmRahm1

If Rahm1.Visible = True Then
    FSuda
Else
    FSett
End If

Me.btnZurück.Enabled = True

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()

Set Rahm1 = Me.frmRahm1

If Rahm1.Visible = True Then
    FSuda
Else
    FSett
End If

Me.btnZurück.Enabled = True
    
End Sub
Private Sub btnZurück_Click()
    FZur
    Me.btnZurück.Enabled = True
End Sub
Private Sub Form_Activate()
On Error Resume Next

Set FTex1 = Me.txtKurz
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2

If Rahm1.Visible = True Then
    FTex1.SetFocus
End If
        
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3: FZur
    Case vbKeyF11: Unload Me
    End Select
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FInit
AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAdrSuch = Nothing
End Sub
Private Sub lstView3_DblClick()
    FSett
End Sub
Private Sub lstView3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub
Private Sub txtBemer_GotFocus()
    TRes
End Sub
Private Sub txtBemer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
        Me.btnZurück.Enabled = True
    End If
End Sub

Private Sub txtGebor_GotFocus()
    TRes
End Sub
Private Sub txtGebor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
        Me.btnZurück.Enabled = True
    End If
End Sub
Private Sub txtKurz_GotFocus()
    TRes
End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
        Me.btnZurück.Enabled = True
    End If
End Sub
Private Sub txtNumme_GotFocus()
    TRes
End Sub
Private Sub txtNumme_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
        Me.btnZurück.Enabled = True
    End If
End Sub

Private Sub txtNumme_Validate(Cancel As Boolean)
    If (Not txtNumme.isValid) Then Cancel = True
End Sub

Private Sub txtPost_GotFocus()
    TRes
End Sub
Private Sub txtPost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
        Me.btnZurück.Enabled = True
    End If
End Sub

Private Sub txtPost_Validate(Cancel As Boolean)
    If (Not txtPost.isValid) Then Cancel = True
End Sub

