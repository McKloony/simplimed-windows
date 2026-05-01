VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Info"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picPict3 
      Height          =   1500
      Left            =   0
      ScaleHeight     =   1440
      ScaleWidth      =   5940
      TabIndex        =   7
      Top             =   4000
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox picPict2 
      BorderStyle     =   0  'Kein
      Height          =   490
      Left            =   800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   1800
      Width           =   490
   End
   Begin VB.PictureBox picPict1 
      BorderStyle     =   0  'Kein
      Height          =   1500
      Left            =   0
      ScaleHeight     =   1500
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin XtremeSuiteControls.PushButton btnWeiter 
      Default         =   -1  'True
      Height          =   400
      Left            =   2350
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3580
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&OK"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblLabe4 
      BackStyle       =   0  'Transparent
      Height          =   220
      Left            =   1500
      TabIndex        =   5
      Top             =   3070
      Width           =   4000
   End
   Begin VB.Label lblLabe3 
      BackStyle       =   0  'Transparent
      Height          =   220
      Left            =   1500
      TabIndex        =   4
      Top             =   2780
      Width           =   4000
   End
   Begin VB.Label lblLabe2 
      BackStyle       =   0  'Transparent
      Height          =   440
      Left            =   1500
      TabIndex        =   3
      Top             =   2260
      Width           =   4000
   End
   Begin VB.Label lblLabe1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1500
      TabIndex        =   2
      Top             =   1900
      Width           =   4000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6000
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6000
      Y1              =   1510
      Y2              =   1510
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Labl1 As VB.Label
Private Labl2 As VB.Label
Private Labl3 As VB.Label
Private Labl4 As VB.Label
Private Pict1 As VB.PictureBox
Private Pict2 As VB.PictureBox
Private Pict3 As VB.PictureBox

Private GrIco() As Long
Private KlIco() As Long

Private Const SeKey = "Seriennummer"
Private Const VaKey = "Variante"

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private clNet As clsNetz
Private clFen As clsFenster
Private clFil As clsFile
Private Sub ALoad()
On Error Resume Next

Dim RetWe As Long
Dim PhoBr As Long
Dim PhoHo As Long
Dim WinBr As Long
Dim TeiBr As Double
Dim TeiHo As Double
Dim TeiSc As Double
Dim FiNam As String
Dim RegNr As String

Set Labl1 = Me.lblLabe1
Set Labl2 = Me.lblLabe2
Set Labl3 = Me.lblLabe3
Set Labl4 = Me.lblLabe4
Set Pict1 = Me.picPict1
Set Pict2 = Me.picPict2
Set Pict3 = Me.picPict3

Set clFil = New clsFile
Set clFen = New clsFenster
Set clNet = New clsNetz
clFen.hwnd = Me.hwnd

FiNam = gAnPfa & gIniNa
WinBr = Me.ScaleWidth

If clFil.FilVor(gFrmGr) = True Then
    Pict3.Picture = LoadPicture(gFrmGr)
    Pict1.ScaleMode = vbPixels
    Pict3.ScaleMode = vbPixels
    With Pict1
        .AutoRedraw = True
        .Picture = LoadPicture(vbNullString)
        .Width = WinBr
        .Height = 1500
        .Refresh
        .AutoRedraw = False
    End With
    TeiBr = Pict1.ScaleWidth / Pict3.ScaleWidth
    TeiHo = Pict1.ScaleHeight / Pict3.ScaleHeight
    
    If Pict3.ScaleWidth * TeiHo > Pict1.ScaleWidth Then
      TeiSc = TeiBr
    Else
      TeiSc = TeiHo
    End If
    PhoBr = Pict3.ScaleWidth * TeiSc
    PhoHo = Pict3.ScaleHeight * TeiSc
    With Pict1
        .AutoRedraw = True
        .Width = PhoBr * Screen.TwipsPerPixelX
        .Height = PhoHo * Screen.TwipsPerPixelY
        .Refresh
        .PaintPicture Pict3.Picture, 0, 0, PhoBr, PhoHo
        .AutoRedraw = False
    End With
End If

If clFil.FilVor(FiNam) = True Then
    RegNr = IniGetFil(FiNam, gIniAn, SeKey)
End If

If clFil.FilVor(gFoIco) = True Then
    RetWe = ExtractIconEx(gFoIco, -1, 0, 0, 0)
    If RetWe > 0 Then
        ReDim GrIco(RetWe)
        ReDim KlIco(RetWe)
        RetWe = ExtractIconEx(gFoIco, 0, GrIco(RetWe), KlIco(RetWe), 1)
        With Pict2
            .BackColor = GlBak
            .Height = 32 * Screen.TwipsPerPixelY
            .Width = 32 * Screen.TwipsPerPixelY
            Set .Picture = LoadPicture(vbNullString)
            .AutoRedraw = True
            RetWe = DrawIconEx(.hDC, 0, 0, GrIco(1), 32, 32, 0, 0, 3)
            .Refresh
        End With
        DestroyIcon RetWe
    End If
End If

If gProNa <> vbNullString Then Labl1.Caption = gProNa
If gProHe <> vbNullString Then Labl2.Caption = gProHe
If gProVe <> vbNullString Then Labl4.Caption = gProVe

Labl3.Caption = "Benutzer: " & clNet.NetBen & " - " & clNet.NetNam

clFen.FenVor

Me.BackColor = GlBak

Set clFil = New clsFile
Set clFen = Nothing
Set clNet = Nothing

End Sub
Private Sub btnWeiter_Click()
    Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

ALoad
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmInfo = Nothing
End Sub
