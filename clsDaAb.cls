VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaAb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clCon As clsConn
Attribute clCon.VB_VarHelpID = -1

Private DaPfa As String 'Path Access Database
Private DaUsr As String 'Username SQL-Server
Private DaSer As String 'Servername SQL-Server
Private DaTab As String 'Databasename SQL-Server
Private DaPas As String 'Datenbankpasswort Access
Private DaUsN As String 'Username Sicheerheitsdatei
Private DaUsP As String 'Passwort Sicheerheitsdatei
Private DaSys As String 'Pfad zur Sicherheitsdatei
Private PaAnw As String 'Anwenderpasswort (SQL Server)
Private DaVer As Integer 'SQL-Server Security
Private DaTim As Integer 'Connection Timeout
Private DaAsy As Boolean 'Asynchron Dataconnection
Public Sub FConn(Optional DatSe As String = vbNullString, Optional ByVal DatTa As String = vbNullString, Optional ByVal DatUs As String = vbNullString, Optional ByVal DatPa As String = vbNullString)
On Error GoTo StErr
'Startet den Datenbankzugriff

Dim TesNa As String

DaVer = CInt(Right(IniGetVal("System", "MatVer"), 1))
DaAsy = CBool(IniGetVal("System", "MatLad"))
DaTab = IniGetVal("System", "MatTab")
DaTim = IniGetVal("System", "MatAkt")
DaSer = IniGetVal("System", "MatSer")
DaUsr = IniGetVal("System", "MatUsr")
DaSys = IniGetVal("System", "MaSich")
DaPas = IniGetVal("System", "MatPas")
DaUsN = IniGetVal("System", "MaUsNa")
DaUsP = IniGetVal("System", "MaSiPa")
DaPfa = IniGetVal("SysPfa", "MatPfa")
PaAnw = IniGetVal("System", "MaUsPa")

If DatSe <> vbNullString Then
    DaSer = DatSe 'Servername SQL-Server
End If

If DatTa <> vbNullString Then
    DaTab = DatTa 'Databasename SQL-Server
End If

If DatUs <> vbNullString Then
    DaUsr = DatUs 'Username SQL-Server
End If

If DatPa <> vbNullString Then
    PaAnw = DatPa 'Passwort SQL-Server
End If

If GlVrs = True Then 'Kennwort verschlüsselt
    DaPas = SCrypt(DaPas, False)
    DaUsN = SCrypt(DaUsN, False)
    DaUsP = SCrypt(DaUsP, False)
    PaAnw = SCrypt(PaAnw, False)
End If

With clCon
    .DaPrm = adPromptComplete
    .DaUsr = DaUsr 'Username SQL-Server
    .DaSer = DaSer 'Servername SQL-Server
    .DaSQL = PaAnw 'Passwort SQL-Server
    .DaPas = DaPas 'Datenbankpasswort Access
    .DaUsN = DaUsN 'Username Sicheerheitsdatei
    .DaUsP = DaUsP 'Passwort Sicheerheitsdatei
    .DaTim = DaTim 'Connection Timeout
    .DaSec = DaVer 'SQL-Server Security
    .DaAsy = DaAsy 'Asynchron Dataconnection
    .DaTyp = GlMyp
    .DaPer = True
    .DaMod = 3
    Select Case GlMyp
    Case 0: 'SQL Native
        If DaTab <> "-" Then
            .DaNam = DaTab
        End If
        .FDBOp
    Case 1: 'SQL OLEDB
        If DaTab <> "-" Then
            .DaNam = DaTab
        End If
        .FDBOp
    Case 2: 'Jet 4.0
        TesNa = SDBTes(DaPfa)
        If TesNa <> vbNullString Then
            .DaNam = TesNa
            .FDBOp
        End If
    Case 3: 'Jet 12.0
        TesNa = SDBTes(DaPfa)
        If TesNa <> vbNullString Then
            .DaNam = TesNa
            .FDBOp
        End If
    End Select
End With

Exit Sub

StErr:
If GlDbg = True Then
    SErLog Err.Description & " FConn " & Err.Number & " - " & Err.Source
End If
Resume Next

End Sub
Private Sub Class_Initialize()
    Set clCon = New clsConn
End Sub
Private Sub Class_Terminate()
    Set clCon = Nothing
End Sub

Private Sub clCon_ConCom(adStat As ADODB.EventStatusEnum)
On Error GoTo StErr

Set DB2 = clCon.DaCon

If DB2.State = adStateClosed Then
    DB2.Open
End If

GlSt2 = DB2.State

Exit Sub

StErr:
If GlDbg = True Then
    SErLog Err.Description & " clCon_ConCom " & Err.Number & " - " & Err.Source
End If
Resume Next

End Sub
