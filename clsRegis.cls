VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As Form

Private Const SeKey = "Seriennummer"
Private Const VaKey = "Variante"
Private Const LiKey = "Lizenz"
Private Const MaKey = "M33A1114S88M3338"

Private clLiz As clsLizenz
Private clFil As clsFile
Public Property Get AnPfa() As String
Attribute AnPfa.VB_Description = "Pfad der Anwendung"
    AnPfa = gAnPfa
End Property

Public Property Let AnPfa(ByVal Flag As String)
    gAnPfa = Flag
End Property

Public Property Get BeLis() As String
Attribute BeLis.VB_Description = "Liste der Benutzertypen"
    BeLis = gBeLis
End Property

Public Property Let BeLis(ByVal Flag As String)
    gBeLis = Flag
End Property

Public Property Get BeTyp() As String
Attribute BeTyp.VB_Description = "Behandlertyp"
    BeTyp = gBeTyp
End Property

Public Property Let BeTyp(ByVal Flag As String)
    gBeTyp = Flag
End Property

Private Sub Class_Terminate()

gFrmGr = vbNullString  'Formulargerafik
gFoIco = vbNullString  'Formularicon
gProNa = vbNullString  'Programname
gProVe = vbNullString  'Programmversion
gProHe = vbNullString  'Programmhersteller
gGlVar = vbNullString  'Programmvariante
gBeTyp = vbNullString  'Benutzertyp
gExNam = vbNullString  'Anwendungsname
gAnPfa = vbNullString  'Programmpfad
gIniNa = vbNullString  'INI-Datei
gIniAn = vbNullString  'INI-Anwendungsname
gAnwKy = vbNullString  'Anwendungskeywort
gSerKy = vbNullString  'Seriennummernkey
gHrdKy = vbNullString  'Hardwarenummernkey
gVerKy = vbNullString  'Verschlüsselungskey
gBeLis = vbNullString  'Benutzertypenliste

End Sub

Public Property Get ExNam() As String
Attribute ExNam.VB_Description = "Name der Anwendung für Registry"
    ExNam = gExNam
End Property

Public Property Let ExNam(ByVal Flag As String)
    gExNam = Flag
End Property

Public Property Get FoIco() As String
    FoIco = gFoIco
End Property

Public Property Let FoIco(ByVal Flag As String)
    gFoIco = Flag
End Property

Public Sub FoInf()
Attribute FoInf.VB_Description = "Zeigt das Infoformular"

Set FM = frmInfo

FM.Show vbModal

End Sub

Public Sub FoLiz()
Attribute FoLiz.VB_Description = "Zeigt das Lizenzierungsfornular"

Set FM = frmLizenz

FM.Show vbModal

End Sub

Public Property Get FrmGr() As String
    FrmGr = gFrmGr
End Property

Public Property Let FrmGr(ByVal Flag As String)
    gFrmGr = Flag
End Property

Public Property Get IniAn() As String
Attribute IniAn.VB_Description = "Anednungsname für INI Datei"
    IniAn = gIniAn
End Property

Public Property Let IniAn(ByVal sAnIni As String)
    gIniAn = sAnIni
End Property

Public Property Get IniNa() As String
Attribute IniNa.VB_Description = "Name der INI Datei"
    IniNa = gIniNa
End Property

Public Property Let IniNa(ByVal Flag As String)
    gIniNa = Flag
End Property

Public Property Get KyAnw() As String
Attribute KyAnw.VB_Description = "Rückgabename der Anwendung"
    KyAnw = gAnwKy
End Property

Public Property Let KyAnw(ByVal Flag As String)
    gAnwKy = Flag
End Property

Public Property Get KyHar() As String
Attribute KyHar.VB_Description = "Rückgabewert der Hardwarenummer"
    KyHar = gHrdKy
End Property

Public Property Let KyHar(ByVal Flag As String)
    gHrdKy = Flag
End Property

Public Property Get KySer() As String
Attribute KySer.VB_Description = "Rückgabewert der Seriennummer"
    KySer = gSerKy
End Property

Public Property Let KySer(ByVal Flag As String)
    gSerKy = Flag
End Property

Public Property Get KyVer() As String
Attribute KyVer.VB_Description = "Rückgabe der Verschlüsselungsbestätigung"
    KyVer = gVerKy
End Property

Public Property Let KyVer(ByVal Flag As String)
    gVerKy = Flag
End Property

Public Property Get ProHe() As String
Attribute ProHe.VB_Description = "Programmhersteller"
    ProHe = gProHe
End Property

Public Property Let ProHe(ByVal Flag As String)
    gProHe = Flag
End Property

Public Property Get ProNa() As String
Attribute ProNa.VB_Description = "Programmname"
    ProNa = gProNa
End Property

Public Property Let ProNa(ByVal Flag As String)
    gProNa = Flag
End Property

Public Property Get GlVar() As String
Attribute GlVar.VB_Description = "Programmvariante"
    GlVar = gGlVar
End Property

Public Property Let GlVar(ByVal Flag As String)
    gGlVar = Flag
End Property

Public Property Get ProVe() As String
Attribute ProVe.VB_Description = "Programmversion"
    ProVe = gProVe
End Property

Public Property Let ProVe(ByVal Flag As String)
    gProVe = Flag
End Property

Private Function ReCry(ByVal SuStr As String, ByVal Flag As Boolean) As String
On Error Resume Next
'Ver/- und Entschlüsselt einen String

Dim AktZa As Long
Dim Posit As Long
Dim CptZa As Long
Dim OrgZa As Long
Dim KeyZa As Long
Dim CptSt As String
Dim KeOut As String
  
For AktZa = 1 To Len(SuStr)
    Posit = Posit + 1
    If Posit > Len(gVerKy) Then Posit = 1
    KeyZa = Asc(Mid(gVerKy, Posit, 1))
    If Flag = True Then ' Verschlüsseln
        OrgZa = Asc(Mid(SuStr, AktZa, 1))
        CptZa = OrgZa Xor KeyZa
        CptSt = Hex(CptZa)
        If Len(CptSt) < 2 Then CptSt = "0" & CptSt
        KeOut = KeOut & CptSt
    Else 'Entschlüsseln
        If AktZa > Len(SuStr) \ 2 Then Exit For
        CptZa = CByte("&H" & Mid$(SuStr, AktZa * 2 - 1, 2))
        OrgZa = CptZa Xor KeyZa
        KeOut = KeOut & Chr$(OrgZa)
    End If
Next AktZa
 
ReCry = KeOut

End Function
Public Function Regis() As Boolean
Attribute Regis.VB_Description = "Überprüft, ob Seriennummer kerrekt ist"
On Error Resume Next

Dim FiNam As String
Dim RegNr As String
Dim SerNr As String
Dim VarNr As String
Dim LizNr As String
Dim Mld1, Mld2, Mld3, Mld4, Tit1 As String

Set clLiz = New clsLizenz
Set clFil = New clsFile

Tit1 = "Programmlizenzierung"
FiNam = gAnPfa & gIniNa
Mld1 = "Die in der Lizenzdatei gespeicherte Seriennummer zu dieser Software hat keine 16 Zeichen"
Mld2 = "Die in der Lizenzdatei gespeicherte Seriennummer zu dieser Software ist falsch. " & _
"Dieses kann mehrere Ursachen haben. Entweder wurde die Lizenzdatei bzw. deren Inhalt ausgetauscht" & _
"oder es wurde eine Hardwarekomponente wie CPU, Netzwerkkarte, Festplatte oder Grafikkarte ausgewechselt bzw. diese Software auf einem neuen PC installiert."
Mld3 = "Der in der Lizenzdatei gespeicherte Produktvarianten-Schlüssel hat keine 6 Zeichen"
Mld4 = "Der in der Lizenzdatei gespeicherte Produktvarianten-Schlüssel ist falsch."

With clLiz
    .AnwNa = gAnwKy
    .KyHar = gHrdKy
    .KySer = gSerKy
    .LiGeSer
    SerNr = .SerNu
End With

If clFil.FilVor(FiNam) = True Then
    RegNr = IniGetFil(FiNam, gIniAn, SeKey)
    If Len(RegNr) = 16 Then
        If SerNr = RegNr Then
            If gSoVar = True Then
                VarNr = IniGetFil(FiNam, gIniAn, VaKey)
                If Len(VarNr) = 6 Then
                    Regis = True
                ElseIf Len(VarNr) > 0 Then
                    MsgBox Mld3, vbExclamation, Tit1
                End If
            Else
                Regis = True
            End If
        Else
            LizNr = IniGetFil(FiNam, gIniAn, LiKey)
            If LizNr = "7774" Then
                Regis = True
            Else
                MsgBox Mld2, vbCritical, Tit1
            End If
        End If
    ElseIf Len(RegNr) > 0 Then
        MsgBox Mld1, vbExclamation, Tit1
    End If
End If

Set clFil = Nothing
Set clLiz = Nothing

End Function
Public Property Get SoVar() As Boolean
Attribute SoVar.VB_Description = "Dialog für Sofwtarevariante"
    SoVar = gSoVar
End Property

Public Property Let SoVar(ByVal Flag As Boolean)
    gSoVar = Flag
End Property
