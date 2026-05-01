VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmTermAnh 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Gefundene Adressen"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   1
      Top             =   3200
      Width           =   5600
      _Version        =   1048579
      _ExtentX        =   9878
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   3600
         TabIndex        =   4
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
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   2200
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Einf¸gen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   900
         TabIndex        =   2
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
   Begin XtremeSuiteControls.ListBox lstList1 
      Height          =   2520
      Left            =   400
      TabIndex        =   0
      Top             =   480
      Width           =   4700
      _Version        =   1048579
      _ExtentX        =   8290
      _ExtentY        =   4445
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte w‰hlen Sie einen der gefundenen Eintr‰ge :"
      Height          =   200
      Left            =   400
      TabIndex        =   5
      Top             =   200
      Width           =   3600
   End
End
Attribute VB_Name = "frmTermAnh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private FLis1 As XtremeSuiteControls.ListBox
Private CmAcs As XtremeCommandBars.CommandBarActions
Private TxOrt As XtremeSuiteControls.FlatEdit

Private TagWe As String

Private clFen As clsFenster
Private Sub FLoad()
On Error GoTo SeErr
'L‰dt die ausgew‰hlte Adresse in das Adressformular

Set FLis1 = Me.lstList1
Set Rahm0 = Me.frmRahm0

With FLis1
    .Font.Name = GlTFt.Name
    .Font.SIZE = GlTFt.SIZE
End With

Rahm0.BackColor = GlBak

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FSett()
On Error GoTo SeErr
'L‰dt die ausgew‰hlte Adresse in das Adressformular

Dim PatNr As Long
Dim LiIdx As Long
Dim MaNum As Long
Dim IdStr As String
Dim PaTel As String
Dim BrStr As String
Dim TmStr As String
Dim Telef As String
Dim TeWer As Variant
Dim BeVor As Boolean
Dim AktZa As Integer
Dim Gesch As Integer
Dim GesZa As Integer
Dim PaBeh As Variant
Dim CmBrs As XtremeCommandBars.CommandBars

Set FLis1 = Me.lstList1

GesZa = FLis1.ListCount

If GesZa > 0 Then
    If WindowLoad("frmTermin") = True Then
        Set FM = frmTermin
    Else
        Set FM = frmTermVo
    End If
    
    Set TxOrt = FM.txtRaum1
    Set CmBrs = FM.comBar02
    Set CmAcs = CmBrs.Actions
    
    PatNr = FLis1.ItemData(FLis1.ListIndex)
    IdStr = FLis1.Text

    S_AdDe PatNr 'Adressendetails
    With GlADt
        PaTel = .AdTe1
        PaBeh = .AdBeh
        Gesch = .AdGet
    End With

    If TxOrt.Text <> vbNullString Then
        TmStr = LCase(TxOrt.Text)
    End If
    
    If PatNr > 0 Then
        FM.txtID0.Text = PatNr
        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
        FM.txtID0.Tag = 1 & TagWe
    End If
    If IdStr <> vbNullString Then
        FM.txtAdres.Text = IdStr
        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
        FM.txtAdres.Tag = 1 & TagWe
    End If
    If PaBeh > 0 Then
        FM.txtBehin.Text = PaBeh
        TagWe = Mid$(FM.txtBehin.Tag, 2, Len(FM.txtBehin.Tag) - 1)
        FM.txtBehin.Tag = 1 & TagWe
    End If
    If Gesch > 0 Then
        FM.cmbGesch.ListIndex = Gesch - 1
        TagWe = Mid$(FM.cmbGesch.Tag, 2, Len(FM.cmbGesch.Tag) - 1)
        FM.cmbGesch.Tag = 1 & TagWe
    End If

    If WindowLoad("frmTermin") = True Then
        If PatNr > 0 Then
            CmAcs(TE_Adresse_Bearbeit).Enabled = True
        End If

        If Left$(TmStr, 6) <> "online" And Left$(TmStr, 6) <> "storno" Then
            S_AdDe PatNr 'Adressendetails
            With GlADt
                If .AdTe1 <> vbNullString Then
                    Telef = .AdTe1
                ElseIf .AdTe2 <> vbNullString Then
                    Telef = .AdTe2
                ElseIf .AdTe4 <> vbNullString Then
                    Telef = .AdTe4
                End If

                FM.txtS4F01.Text = .AdFir
                TagWe = Mid$(FM.txtS4F01.Tag, 2, Len(FM.txtS4F01.Tag) - 1)
                FM.txtS4F01.Tag = 1 & TagWe
    
                FM.txtS4F02.Text = .AdAnr
                TagWe = Mid$(FM.txtS4F02.Tag, 2, Len(FM.txtS4F02.Tag) - 1)
                FM.txtS4F02.Tag = 1 & TagWe
                    
                FM.txtS4F03.Text = .AdTit
                TagWe = Mid$(FM.txtS4F03.Tag, 2, Len(FM.txtS4F03.Tag) - 1)
                FM.txtS4F03.Tag = 1 & TagWe
                
                FM.txtS4F04.Text = .AdVor
                TagWe = Mid$(FM.txtS4F04.Tag, 2, Len(FM.txtS4F04.Tag) - 1)
                FM.txtS4F04.Tag = 1 & TagWe
                    
                FM.txtS4F05.Text = .AdNam
                TagWe = Mid$(FM.txtS4F05.Tag, 2, Len(FM.txtS4F05.Tag) - 1)
                FM.txtS4F05.Tag = 1 & TagWe
                    
                FM.txtS4F06.Text = .AdStr
                TagWe = Mid$(FM.txtS4F06.Tag, 2, Len(FM.txtS4F06.Tag) - 1)
                FM.txtS4F06.Tag = 1 & TagWe
                    
                FM.txtS4F08.Text = .AdPLZ
                TagWe = Mid$(FM.txtS4F08.Tag, 2, Len(FM.txtS4F08.Tag) - 1)
                FM.txtS4F08.Tag = 1 & TagWe
                    
                FM.txtS4F09.Text = .AdOrt
                TagWe = Mid$(FM.txtS4F09.Tag, 2, Len(FM.txtS4F09.Tag) - 1)
                FM.txtS4F09.Tag = 1 & TagWe
            
                FM.cmbS4F12.Text = .AdLan
                TagWe = Mid$(FM.cmbS4F12.Tag, 2, Len(FM.cmbS4F12.Tag) - 1)
                FM.cmbS4F12.Tag = 1 & TagWe
            
                FM.txtS4F18.Text = .AdGeb
                TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
                FM.txtS4F18.Tag = 1 & TagWe
            
                FM.txtS4F15.Text = Telef
                TagWe = Mid$(FM.txtS4F15.Tag, 2, Len(FM.txtS4F15.Tag) - 1)
                FM.txtS4F15.Tag = 1 & TagWe
            
                FM.txtS4F16.Text = .AdTe5
                TagWe = Mid$(FM.txtS4F16.Tag, 2, Len(FM.txtS4F16.Tag) - 1)
                FM.txtS4F16.Tag = 1 & TagWe
            
                BrStr = .AdBrf
                Ter_Brz BrStr
                TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
                FM.txtS4F18.Tag = 1 & TagWe
                
                If GlCaT = True Then
                    FM.txtKomme.Text = .AdCav
                    TagWe = Mid$(FM.txtKomme.Tag, 2, Len(FM.txtKomme.Tag) - 1)
                    FM.txtKomme.Tag = 1 & TagWe
                End If
            End With
        End If
    End If

    TeWer = S_AdIdi(PatNr, "IDP")
    If TeWer > 0 Then
        MaNum = CLng(TeWer)
        For AktZa = 1 To UBound(GlMan)
            If MaNum = GlMan(AktZa, 2) Then
                BeVor = True
                Exit For
            End If
        Next AktZa
        If BeVor = True Then
            If CBool(GlMan(AktZa, 5)) = False Then 'Passiv / Aktiv
                MaNum = GlThe(AktZa, 0)
            Else
                MaNum = GlMan(GlSMa, 2)
            End If
        Else
            MaNum = GlMan(GlSMa, 2)
        End If
    Else
        MaNum = GlMan(GlSMa, 2)
    End If
    
    If GlBut <> RibTab_Ter_Mitarb Then
        LiIdx = SCmb(FM.cmbBehan, MaNum)
        FM.cmbBehan.ListIndex = LiIdx
        TagWe = Mid$(FM.cmbBehan.Tag, 2, Len(FM.cmbBehan.Tag) - 1)
        FM.cmbBehan.Tag = 1 & TagWe
    End If

    If WindowLoad("frmTermin") = False Then
        CmAcs(AD_Termin_Abrechnen).Enabled = False
        CmAcs(AD_Termin_EintLoe).Enabled = True
        CmAcs(AD_Termin_Ketten).Enabled = True
        CmAcs(AD_Termin_StaKett).Enabled = True
    End If
    GlTSa = True 'Speichern des Termins erforderlich
    Unload Me
End If

Exit Sub

SeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSett " & Err.Number
Resume Next

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
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FSett
End Sub
Private Sub Form_Load()
On Error Resume Next

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Me.BackColor = GlBak

AFont Me

clFen.FenVor

Set clFen = Nothing

SFrame 1, Me.hwnd

FLoad

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTermAnh = Nothing
End Sub
Private Sub lstList1_DblClick()
    FSett
End Sub
Private Sub lstList1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSett
    End If
End Sub
