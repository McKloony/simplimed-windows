VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDaTy As Long    'Datenbanktyp
Private mTime As Long    'Timeout
Private mMode As Long    'Mode Eigenschaft
Private mOpti As Long    'Option
Private mDaNa As String  'Datenbankname
Private mPass As String  'Datenbankpasswort für Access
Private mPaSQ As String  'Datenbankpasswort für SQL
Private mServ As String  'Servername
Private mUser As String  'Username für SQL-Server
Private mUsNa As String  'Username für Access
Private mUsPa As String  'Userpasswort für Access
Private mSyst As String  'Sicherheitsdatei
Private mSecu As Integer 'Interne Sicherheit
Private mProm As Integer 'Prompt
Private mAsyn As Boolean 'Asynchron
Private mConn As Boolean 'Connection Complete
Private mPers As Boolean 'Persistente Sicherheitsinfo

'Private Const adAsyncExecute = 16
'Private Const adAsyncFetch = 32
'Private Const adAsyncFetchNonBlocking = 64
'Private Const adExecuteNoRecords = 128

Private WithEvents RS100 As ADODB.Recordset
Attribute RS100.VB_VarHelpID = -1
Private WithEvents RS101 As ADODB.Recordset
Attribute RS101.VB_VarHelpID = -1
Private WithEvents RS102 As ADODB.Recordset
Attribute RS102.VB_VarHelpID = -1
Private WithEvents RS103 As ADODB.Recordset
Attribute RS103.VB_VarHelpID = -1

Private WithEvents CN100 As ADODB.Connection
Attribute CN100.VB_VarHelpID = -1

Private CM101 As ADODB.Command
Private PA101 As ADODB.Parameter

Private PR100 As ADODB.ConnectPromptEnum

Public Event ConCom(adStat As ADODB.EventStatusEnum)
Public Event ReSe00(adSta1 As ADODB.EventStatusEnum)
Public Event ReSe01(adSta1 As ADODB.EventStatusEnum)
Public Event ReSe02(adSta2 As ADODB.EventStatusEnum)
Public Event ReSe03(adSta3 As ADODB.EventStatusEnum)

Private Const DB_MODE_READ = 1
Private Const DB_MODE_WRITE = 2
Private Const DB_MODE_READWRITE = 3
Private Const DB_MODE_SHARE_DENY_READ = 4
Private Const DB_MODE_SHARE_DENY_WRITE = 8
Private Const DB_MODE_SHARE_DENY_EXCLUSIVE = 12
Private Const DB_MODE_SHARE_DENY_NONE = 16
Public Property Let DaSQL(Flag As String)
    mPaSQ = Flag
End Property

Public Property Get DaSQL() As String
    DaPas = mPaSQ
End Property
Private Sub Class_Initialize()
    Set CN100 = New ADODB.Connection
End Sub
Private Sub CN100_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
On Error GoTo StErr

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim DaPfa As String
Dim Frage As Integer

If GlTyp < 2 Then
    TeTit = "Fehler beim Öffnen der SQL-Datenbank"
    TeMai = "Beim Öffnen der SQL-Datenbank ist ein Fehler aufgetreten. Sollen die Datenbank-Einstellungen jetzt zurückgesetzt werden?"
    TeInh = "Beim Versucht auf den SQL-Server zuzugreifen ist ein Fehler aufgetreten. Entweder sind die Einstellungen im Abschnitt: Datenbankeinstellungen nicht richtig oder der SQL-Server ist nicht erreichbar"
    TeFus = "SimpliMed wird nun beendet. Bitte starten Sie SimpliMed erneut und korrigieren falls erforderlich die Einstellungen im Abschnitt: Datenbankeinstellungen."
Else
    TeTit = "Fehler beim Öffnen de Datenbankdatei"
    TeMai = "Beim Öffnen der Datenbankdatei ist ein Fehler aufgetreten. Sollen die Datenbankdatei-Einstellungen jetzt zurückgesetzt werden?"
    TeInh = "Die Daten konnten nicht korrekt gelesen werden. Entweder ist die Datenbankdatei beschädigt oder Sie versuchen eine Datei zu öffnen, die keine SimpliMed Datenbankdatei ist. Das Programm wird daher beendet"
    TeFus = "Bitte löschen Sie die Konfigurationsdatei im Ordner Praxisdaten und versuchen Sie das Programm danach erneut zu starten"
End If

If adStatus = adStatusOK Then
    mConn = True
    RaiseEvent ConCom(adStatusCancel)
Else
    mConn = False
    SSplS1 'Splashscreen schließen
    SSplS2 'Splashscreen schließen
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, frmMain.hwnd
    If GlMes = 33565 Then
        If GlTyp < 2 Then
            If GlRDP = False Then
                IniSetVal "System", "DBaTyp", "S3"
            End If
        Else
            IniSetVal "SysPfa", "DatPfa", LCase(GlPfa & "Datenbankdatei.dbv")
        End If
    End If
    Unload frmMain
End If

Exit Sub

StErr:
If GlDbg = True Then SErLog Err.Description & " CN100_ConnectComplete " & Err.Number
SSplS1 'Splashscreen schließen
SSplS2 'Splashscreen schließen
DoEvents
SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, True, frmMain.hwnd
If GlMes = 33565 Then
    If GlTyp < 2 Then
        If GlRDP = False Then
            IniSetVal "System", "DBaTyp", "S3"
        End If
    Else
        IniSetVal "SysPfa", "DatPfa", LCase(GlPfa & "Datenbankdatei.dbv")
    End If
End If
Unload frmMain

End Sub
Public Property Get DaAsy() As Boolean
Attribute DaAsy.VB_Description = "Asynchrone Verbindung ja/nein"
    DaAsy = mAsyn
End Property

Public Property Let DaAsy(Flag As Boolean)
    mAsyn = Flag
End Property

Public Property Let DaCon(Flag As ADODB.Connection)
Attribute DaCon.VB_Description = "ADODB Connection Object"
    Set CN100 = Flag
End Property

Public Property Get DaCon() As ADODB.Connection
    Set DaCon = CN100
End Property

Public Property Let DaKom(Flag As Boolean)
Attribute DaKom.VB_Description = "Wird auf True gesetzt, wenn Verbindung erfolgreich hergestellt wurde"
    mConn = Flag
End Property

Public Property Get DaKom() As Boolean
    DaKom = mConn
End Property

Public Property Let DaMod(Flag As Long)
Attribute DaMod.VB_Description = "Mode Eigenschaft des Recordsets"
    mMode = Flag
End Property

Public Property Get DaMod() As Long
    DaMod = mMode
End Property

Public Property Let DaNam(Flag As String)
Attribute DaNam.VB_Description = "Access: Pfadangabe zur Datenbankdatei / SQL-Server: Datenbankname"
    mDaNa = Flag
End Property

Public Property Get DaNam() As String
    DaNam = mDaNa
End Property

Public Property Let DaOpt(Flag As Long)
Attribute DaOpt.VB_Description = "Zusätzliche Optionen (mom. nicht verwendet)"
    mOpti = Flag
End Property

Public Property Get DaOpt() As Long
    DaOpt = mOpti
End Property

Public Property Let DaPas(Flag As String)
Attribute DaPas.VB_Description = "Passwort der Datenbankdatei"
    mPass = Flag
End Property

Public Property Get DaPas() As String
    DaPas = mPass
End Property

Public Property Get DaPer() As Boolean
Attribute DaPer.VB_Description = "Persistente Sicherheitsinformationen halten"
    DaPer = mPers
End Property

Public Property Let DaPer(Flag As Boolean)
    mPers = Flag
End Property

Public Property Let DaSec(Flag As Integer)
Attribute DaSec.VB_Description = "1=NT-Authentifizierung ; 2=Username/Passwort"
    mSecu = Flag
End Property
Public Property Get DaSec() As Integer
    DaSec = mSecu
End Property
Public Property Let DaPrm(Flag As Integer)
    mProm = Flag
End Property
Public Property Get DaPrm() As Integer
    DaPrm = mProm
End Property
Public Property Let DaSer(Flag As String)
Attribute DaSer.VB_Description = "Name des Datenbankservers"
    mServ = Flag
End Property

Public Property Get DaSer() As String
    DaSer = mServ
End Property

Public Property Let DaSys(Flag As String)
Attribute DaSys.VB_Description = "Name der Sicherheitsdatei"
    mSyst = Flag
End Property

Public Property Get DaSys() As String
    DaSys = mSyst
End Property

Public Property Let DaTim(Flag As Long)
Attribute DaTim.VB_Description = "Datenbanktimeout"
    mTime = Flag
End Property

Public Property Get DaTim() As Long
    DaTim = mTime
End Property

Public Property Let DaTyp(Flag As Long)
Attribute DaTyp.VB_Description = "1:SQL-Server 2:Access"
    mDaTy = Flag
End Property

Public Property Get DaTyp() As Long
    DaTyp = mDaTy
End Property

Public Property Let DaUsN(Flag As String)
Attribute DaUsN.VB_Description = "Username für Sicherhietsdatei"
    mUsNa = Flag
End Property

Public Property Get DaUsN() As String
    DaUsN = mUsNa
End Property

Public Property Let DaUsP(Flag As String)
Attribute DaUsP.VB_Description = "Userpasswort für Sicherheitsdatei"
    mUsPa = Flag
End Property

Public Property Get DaUsP() As String
    DaUsP = mUsPa
End Property

Public Property Let DaUsr(Flag As String)
Attribute DaUsr.VB_Description = "Username für SQL-Server"
    mUser = Flag
End Property

Public Property Get RecSet100() As ADODB.Recordset
    Set RecSet100 = RS100
End Property
Public Property Let RecSet100(Flag As ADODB.Recordset)
    Set RS100 = Flag
End Property
Public Property Get RecSet200() As ADODB.Recordset
    Set RecSet200 = RS102
End Property
Public Property Let RecSet200(Flag As ADODB.Recordset)
    Set RS102 = Flag
End Property
Public Property Get RecSet300() As ADODB.Recordset
    Set RecSet300 = RS103
End Property
Public Property Let RecSet300(Flag As ADODB.Recordset)
    Set RS103 = Flag
End Property
Public Property Get DaUsr() As String
    DaUsr = mUser
End Property

Public Sub FDBOp()
On Error GoTo StErr
'Öffnet die Verbindung zur Datenbankdatei und stellt Connectionobjekt zur Verfügung

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

If GlTyp < 2 Then
    TeTit = "Fehler beim Öffnen der SQL-Datenbank"
    TeMai = "Beim Öffnen der SQL-Datenbank ist ein Fehler aufgetreten. Sollen die Datenbank-Einstellungen jetzt zurückgesetzt werden?"
    TeInh = "Beim Versucht auf den SQL-Server zuzugreifen ist ein Fehler aufgetreten. Entweder sind die Einstellungen im Abschnitt: Datenbankeinstellungen nicht richtig oder der SQL-Server ist nicht erreichbar"
    TeFus = "SimpliMed wird nun beendet. Bitte starten Sie SimpliMed erneut und korrigieren falls erforderlich die Einstellungen im Abschnitt: Datenbankeinstellungen."
Else
    TeTit = "Fehler beim Öffnen de Datenbankdatei"
    TeMai = "Beim Öffnen der Datenbankdatei ist ein Fehler aufgetreten. Sollen die Datenbankdatei-Einstellungen jetzt zurückgesetzt werden?"
    TeInh = "Die Daten konnten nicht korrekt gelesen werden. Entweder ist die Datenbankdatei beschädigt oder Sie versuchen eine Datei zu öffnen, die keine SimpliMed Datenbankdatei ist. Das Programm wird daher beendet"
    TeFus = "Bitte löschen Sie die Konfigurationsdatei im Ordner Praxisdaten und versuchen Sie das Programm danach erneut zu starten"
End If

If mProm > 0 Then
    PR100 = mProm
Else
    PR100 = adPromptComplete
End If

Select Case mDaTy
Case 0: 'Native Client
    With CN100
        .Provider = "SQLNCLI.1"
        .CommandTimeout = mTime
        .ConnectionTimeout = mTime
        .CursorLocation = adUseClient
        .Mode = mMode
        .Properties("Data Source") = mServ
        .Properties("Initial Catalog") = mDaNa
        .Properties("Persist Security Info") = mPers
        Select Case mSecu
        Case 1: .Properties("Integrated Security") = "SSPI"
        Case 2: .Properties("User ID") = mUser
                .Properties("Password") = mPaSQ
        End Select
        .Properties("Prompt") = PR100
        mServ = .Properties("Data Source")
        mDaNa = .Properties("Initial Catalog")
        mPers = .Properties("Persist Security Info")
        mUser = .Properties("User ID")
        mPaSQ = .Properties("Password")
        If mAsyn = True Then
            .Open Options:=adAsyncConnect
        Else
            .Open
        End If
    End With
Case 1: 'OLDE DB
    With CN100
        .Provider = "SQLOLEDB.1"
        .CommandTimeout = mTime
        .ConnectionTimeout = mTime
        .CursorLocation = adUseClient
        .Mode = mMode
        .Properties("Data Source") = mServ
        .Properties("Initial Catalog") = mDaNa
        .Properties("Persist Security Info") = mPers
        Select Case mSecu
        Case 1: .Properties("Integrated Security") = "SSPI"
        Case 2: .Properties("User ID") = mUser
                .Properties("Password") = mPaSQ
        End Select
        .Properties("Prompt") = PR100
        mServ = .Properties("Data Source")
        mDaNa = .Properties("Initial Catalog")
        mPers = .Properties("Persist Security Info")
        mUser = .Properties("User ID")
        mPaSQ = .Properties("Password")
        If mAsyn = True Then
            .Open Options:=adAsyncConnect
        Else
            .Open
        End If
    End With
Case 2: 'Jet 4.0
    With CN100
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .CommandTimeout = mTime
        .ConnectionTimeout = mTime
        .CursorLocation = adUseClient
        .Mode = mMode
        .Properties("Data Source") = mDaNa
        .Properties("Jet OLEDB:Engine Type") = 5
        .Properties("Persist Security Info") = mPers
        If mPass <> vbNullString Then
            .Properties("Jet OLEDB:Database Password") = mPass
        End If
        If mSyst <> vbNullString Then
            .Properties("Jet OLEDB:System Database") = mSyst
            .Properties("User ID") = mUsNa
            .Properties("Password") = mUsPa
        End If
        If mAsyn = True Then
            .Open Options:=adAsyncConnect
        Else
            .Open
        End If
    End With
Case 3: 'Jet 12.0
    With CN100
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .CommandTimeout = mTime
        .ConnectionTimeout = mTime
        .CursorLocation = adUseClient
        .Mode = mMode
        .Properties("Data Source") = mDaNa
        .Properties("Jet OLEDB:Engine Type") = 6
        .Properties("Persist Security Info") = mPers
        If mPass <> vbNullString Then
            .Properties("Jet OLEDB:Database Password") = mPass
        End If
        If mSyst <> vbNullString Then
            .Properties("Jet OLEDB:System Database") = mSyst
            .Properties("User ID") = mUsNa
            .Properties("Password") = mUsPa
        End If
        If mAsyn = True Then
            .Open Options:=adAsyncConnect
        Else
            .Open
        End If
    End With
End Select

Exit Sub

StErr:
If GlDbg = True Then SErLog Err.Description & " FDBOp " & Err.Number
SSplS1 'Splashscreen schließen
SSplS2 'Splashscreen schließen
DoEvents
SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, frmMain.hwnd
If GlMes = 33565 Then
    If GlTyp < 2 Then
        If GlRDP = False Then
            IniSetVal "System", "DBaTyp", "S3"
        End If
    Else
        IniSetVal "SysPfa", "DatPfa", LCase(GlPfa & "Datenbankdatei.dbv")
    End If
End If
Unload frmMain

End Sub
Private Sub RS100_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

If adStatus = adStatusOK Then
    RaiseEvent ReSe00(adStatusCancel)
End If

End Sub

Private Sub RS101_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

If adStatus = adStatusOK Then
    RaiseEvent ReSe01(adStatusCancel)
End If

End Sub

Private Sub RS102_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

If adStatus = adStatusOK Then
    RaiseEvent ReSe02(adStatusCancel)
End If

End Sub

Private Sub RS103_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

If adStatus = adStatusOK Then
    RaiseEvent ReSe03(adStatusCancel)
End If

End Sub

