VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmSQLCon 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Server Verbindung"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   5
      Top             =   3800
      Width           =   5000
      _Version        =   1048579
      _ExtentX        =   8819
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   3000
         TabIndex        =   7
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
      Begin XtremeSuiteControls.PushButton btnWieter 
         Default         =   -1  'True
         Height          =   400
         Left            =   1600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Verbinden"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
   End
   Begin XtremeSuiteControls.ComboBox cmbServe 
      Height          =   315
      Left            =   1000
      TabIndex        =   1
      Top             =   1130
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox4"
      DropDownItemCount=   0
   End
   Begin XtremeSuiteControls.ComboBox cmbDatNa 
      Height          =   315
      Left            =   1000
      TabIndex        =   4
      Top             =   3230
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox5"
   End
   Begin XtremeSuiteControls.FlatEdit txtUseNa 
      Height          =   350
      Left            =   1000
      TabIndex        =   2
      Top             =   1830
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtPassw 
      Height          =   350
      Left            =   1000
      TabIndex        =   3
      Top             =   2530
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      PasswordChar    =   "*"
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   5500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   495
      Left            =   1000
      TabIndex        =   12
      Top             =   200
      Width           =   3600
      _Version        =   1048579
      _ExtentX        =   6350
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Bitte stellen Sie die gewünschten SQL Server Verbindungsdaten ein und klicken auf Verbinden."
      Alignment       =   4
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLab04 
      BackStyle       =   0  'Transparent
      Caption         =   "Passwort :"
      Height          =   210
      Left            =   1000
      TabIndex        =   11
      Top             =   2300
      Width           =   1200
   End
   Begin VB.Label lblLab03 
      BackStyle       =   0  'Transparent
      Caption         =   "Benutzername :"
      Height          =   210
      Left            =   1000
      TabIndex        =   10
      Top             =   1600
      Width           =   1200
   End
   Begin VB.Label lblLab02 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL-Server :"
      Height          =   210
      Left            =   1000
      TabIndex        =   9
      Top             =   900
      Width           =   1200
   End
   Begin VB.Label lblLab05 
      BackStyle       =   0  'Transparent
      Caption         =   "Datenbank :"
      Height          =   210
      Left            =   1000
      TabIndex        =   8
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "frmSQLCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private LbLab As XtremeSuiteControls.Label
Private Rahm0 As XtremeSuiteControls.GroupBox
Private CmSer As XtremeSuiteControls.ComboBox
Private CmDat As XtremeSuiteControls.ComboBox
Private TxUse As XtremeSuiteControls.FlatEdit
Private TxPas As XtremeSuiteControls.FlatEdit

Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const KEYEVENTF_KEYUP = &H2
Private Const CB_FINDSTRING = &H14C&
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private FoLad As Boolean
Private PasWo As Boolean

Private clAbd As clsDaAb

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub FDate()
On Error GoTo SuErr

Dim DatSe As String
Dim DatUs As String
Dim DatPa As String

Set FM = frmSQLCon
Set CmSer = FM.cmbServe
Set CmDat = FM.cmbDatNa
Set TxUse = FM.txtUseNa
Set TxPas = FM.txtPassw

If CmSer.Text <> vbNullString Then
    DatSe = CmSer.Text
End If

Select Case CmSer.ListIndex
Case 0:
    If IniGetVal("RDPSek", "SQLUs0") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs0")
    End If
    If IniGetVal("RDPSek", "SQLPa0") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa0")
    End If
Case 1:
    If IniGetVal("RDPSek", "SQLUs1") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs1")
    End If
    If IniGetVal("RDPSek", "SQLPa1") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa1")
    End If
Case 2:
    If IniGetVal("RDPSek", "SQLUs2") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs2")
    End If
    If IniGetVal("RDPSek", "SQLPa2") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa2")
    End If
Case 3:
    If IniGetVal("RDPSek", "SQLUs3") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs3")
    End If
    If IniGetVal("RDPSek", "SQLPa3") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa3")
    End If
Case 4:
    If IniGetVal("RDPSek", "SQLUs4") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs4")
    End If
    If IniGetVal("RDPSek", "SQLPa4") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa4")
    End If
Case 5:
    If IniGetVal("RDPSek", "SQLUs5") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs5")
    End If
    If IniGetVal("RDPSek", "SQLPa5") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa5")
    End If
Case 6:
    If IniGetVal("RDPSek", "SQLUs6") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs6")
    End If
    If IniGetVal("RDPSek", "SQLPa6") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa6")
    End If
Case 7:
    If IniGetVal("RDPSek", "SQLUs7") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs7")
    End If
    If IniGetVal("RDPSek", "SQLPa7") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa7")
    End If
Case 8:
    If IniGetVal("RDPSek", "SQLUs8") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs8")
    End If
    If IniGetVal("RDPSek", "SQLPa8") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa8")
    End If
Case 9:
    If IniGetVal("RDPSek", "SQLUs9") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs9")
    End If
    If IniGetVal("RDPSek", "SQLPa9") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa9")
    End If
Case 10:
    If IniGetVal("RDPSek", "SQLUs10") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs10")
    End If
    If IniGetVal("RDPSek", "SQLPa10") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa10")
    End If
Case 11:
    If IniGetVal("RDPSek", "SQLUs11") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs11")
    End If
    If IniGetVal("RDPSek", "SQLPa11") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa11")
    End If
Case 12:
    If IniGetVal("RDPSek", "SQLUs12") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs12")
    End If
    If IniGetVal("RDPSek", "SQLPa12") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa12")
    End If
Case 13:
    If IniGetVal("RDPSek", "SQLUs13") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs13")
    End If
    If IniGetVal("RDPSek", "SQLPa13") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa13")
    End If
Case 14:
    If IniGetVal("RDPSek", "SQLUs14") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs14")
    End If
    If IniGetVal("RDPSek", "SQLPa14") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa14")
    End If
Case 15:
    If IniGetVal("RDPSek", "SQLUs15") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs15")
    End If
    If IniGetVal("RDPSek", "SQLPa15") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa15")
    End If
Case 16:
    If IniGetVal("RDPSek", "SQLUs16") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs16")
    End If
    If IniGetVal("RDPSek", "SQLPa16") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa16")
    End If
Case 17:
    If IniGetVal("RDPSek", "SQLUs17") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs17")
    End If
    If IniGetVal("RDPSek", "SQLPa17") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa17")
    End If
Case 18:
    If IniGetVal("RDPSek", "SQLUs18") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs18")
    End If
    If IniGetVal("RDPSek", "SQLPa18") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa18")
    End If
Case 19:
    If IniGetVal("RDPSek", "SQLUs19") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs19")
    End If
    If IniGetVal("RDPSek", "SQLPa19") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa19")
    End If
Case 20:
    If IniGetVal("RDPSek", "SQLUs20") <> vbNullString Then
        TxUse.Text = IniGetVal("RDPSek", "SQLUs20")
    End If
    If IniGetVal("RDPSek", "SQLPa20") <> vbNullString Then
        TxPas.Text = IniGetVal("RDPSek", "SQLPa20")
    End If
End Select

If TxUse.Text <> vbNullString Then
    DatUs = TxUse.Text
End If

If TxPas.Text <> vbNullString Then
    DatPa = TxPas.Text
End If

S_Tabe DatSe, "-", DatUs, DatPa

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDate " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim DaSer As String
Dim AnzSe As Integer
Dim AktSe As Integer
Dim DatSe() As String

Set FM = frmSQLCon
Set CmSer = FM.cmbServe
Set CmDat = FM.cmbDatNa
Set TxUse = FM.txtUseNa
Set TxPas = FM.txtPassw
Set LbLab = FM.lblLab01
Set Rahm0 = FM.frmRahm0

DaSer = IniGetVal("System", "DatSer")

If IniGetVal("RDPSek", "SQLSe0") <> vbNullString Then
    AnzSe = 0
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe0")
End If
If IniGetVal("RDPSek", "SQLSe1") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe1")
End If
If IniGetVal("RDPSek", "SQLSe2") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe2")
End If
If IniGetVal("RDPSek", "SQLSe3") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe3")
End If
If IniGetVal("RDPSek", "SQLSe4") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe4")
End If
If IniGetVal("RDPSek", "SQLSe5") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe5")
End If
If IniGetVal("RDPSek", "SQLSe6") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe6")
End If
If IniGetVal("RDPSek", "SQLSe7") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe7")
End If
If IniGetVal("RDPSek", "SQLSe8") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe8")
End If
If IniGetVal("RDPSek", "SQLSe9") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe9")
End If
If IniGetVal("RDPSek", "SQLSe10") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe10")
End If
If IniGetVal("RDPSek", "SQLSe11") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe11")
End If
If IniGetVal("RDPSek", "SQLSe12") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe12")
End If
If IniGetVal("RDPSek", "SQLSe13") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe13")
End If
If IniGetVal("RDPSek", "SQLSe14") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe14")
End If
If IniGetVal("RDPSek", "SQLSe15") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe15")
End If
If IniGetVal("RDPSek", "SQLSe16") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe16")
End If
If IniGetVal("RDPSek", "SQLSe17") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe17")
End If
If IniGetVal("RDPSek", "SQLSe18") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe18")
End If
If IniGetVal("RDPSek", "SQLSe19") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe19")
End If
If IniGetVal("RDPSek", "SQLSe20") <> vbNullString Then
    AnzSe = AnzSe + 1
    ReDim Preserve DatSe(AnzSe)
    DatSe(AnzSe) = IniGetVal("RDPSek", "SQLSe20")
End If

With CmSer
    .DropDownItemCount = 21

    For AktSe = 0 To AnzSe
        .AddItem (DatSe(AktSe))
        .ItemData(AktSe) = AktSe + 1
    Next AktSe
    
    For AktSe = 0 To AnzSe
        If LCase(DatSe(AktSe)) = LCase(DaSer) Then
            .ListIndex = AktSe
            Exit For
        End If
    Next AktSe
End With

LbLab.BackColor = GlBak
FM.BackColor = GlBak
Rahm0.BackColor = GlBak

DoEvents
FoLad = False

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FVerb()
On Error GoTo SuErr

Dim DatSe As String
Dim DatDa As String
Dim DatUs As String
Dim DatPa As String
Dim UnStr As String

Set FM = frmSQLCon
Set CmSer = FM.cmbServe
Set CmDat = FM.cmbDatNa
Set TxUse = FM.txtUseNa
Set TxPas = FM.txtPassw

If CmSer.Text <> vbNullString Then
    DatSe = CmSer.Text
End If

If CmDat.Text <> vbNullString Then
    DatDa = CmDat.Text
End If

If TxUse.Text <> vbNullString Then
    DatUs = TxUse.Text
End If

If TxPas.Text <> vbNullString Then
    DatPa = TxPas.Text
    If PasWo = True Then
        DatPa = SCrypt(DatPa, True)
    End If
End If

'----

If InStr(1, DatDa, "TeleWorker", 1) > 0 Then
    UnStr = Mid$(DatDa, 12, 1) 'Unitstring
    GlFrn = "\\simplimed.int\simplimed-dfs\groupdrive\" & UnStr & "-unit\" & DatDa & "\Praxisdaten\Formulare\"
Else
    GlFrn = vbNullString 'Formulareordner
End If

'----

Screen.MousePointer = vbHourglass

SReb 1, True, vbNullString, DatSe, DatDa, DatUs, DatPa

Screen.MousePointer = vbNormal

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FVerb " & Err.Number
Resume Next

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub

Private Sub btnWieter_Click()
    FVerb
    DoEvents
    Unload Me
End Sub
Private Sub cmbServe_Click()
    FDate
End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True

FInit
AFont Me
SFrame 1, Me.hwnd
  
End Sub

Private Sub txtPassw_GotFocus()
    Me.txtPassw.SelStart = 0
    Me.txtPassw.SelLength = Len(Me.txtPassw.Text)
End Sub

Private Sub txtPassw_KeyPress(KeyAscii As Integer)
    PasWo = True
End Sub
Private Sub txtUseNa_GotFocus()
    Me.txtUseNa.SelStart = 0
    Me.txtUseNa.SelLength = Len(Me.txtUseNa.Text)
End Sub
