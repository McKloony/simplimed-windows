VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMailing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private Lab01 As XtremeSuiteControls.Label
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar

Private mAnon As Boolean 'Anonymous
Private mAtFi As String  'Attachmentfile
Private mAlAk As Boolean 'Emailabfrufalter aktivieren
Private mAlTy As Integer 'Emailabfrufalter Typ
Private mAtVo As Boolean 'Attachment vorhanden
Private mAsyn As Integer 'Asynchronous
Private mAuTy As Integer 'Athentifizierungstyp
Private mAuUs As Boolean 'Benutzerauthentifizierung
Private mBest As Boolean 'Emailbestätigung
Private mCoNa As String  'ComputerName
Private mCoIP As String  'LocalIP
Private mFiln As String  'Emailfiestring
Private mFold As String  'Emailfolder
Private mFrmt As Integer 'Format 1=HTML 2=ASCII
Private mGele As Boolean 'als gelesen kennzeichnen
Private mHELO As String  'HELO/EHLO
Private mIMAP As String  'Server IMAP
Private mImag As String  'Signaturimage
Private mNoDe As Boolean 'No Emaildelete on Server
Private mPOP3 As String  'Server POP3
Private mPor1 As Integer 'Port GetMail
Private mPor2 As Integer 'Port SendMail
Private mPosf As Integer 'Postfachtyp
Private mPrio As Integer 'Priority
Private mProg As Boolean 'Progressbar
Private mPxSe As String  'SocksProxyServer
Private mPxPa As String  'SocksProxyPassword
Private mPxPo As String  'SocksProxyPort
Private mPxUs As String  'SocksProxyUser
Private mReAd As String  'Recipientadress
Private mReBc As String  'RecipientBcc
Private mReCc As String  'RecipientCc
Private mReNa As String  'Recipientname
Private mSubj As String  'Emailsubject
Private mSeAd As String  'Senderadress
Private mSeNa As String  'Sendername
Private mSeRe As String  'Senderreplyto
Private mSMTP As String  'Server SMTP
Private mTrKy As String  'TreeKey
Private mTLSg As Integer 'SSL/TLS Getmail
Private mTLSs As Integer 'SSL/TLS Sendmail
Private mTxAc As String  'Bodytext ASCII
Private mTxHt As String  'Bodytext HTML
Private mUsNa As String  'Username
Private mUsPa As String  'Userpassword
Private mUIDn As String  'UIDL Filename

Private Const ConnectNormal = 0
Private Const ConnectSSLAuto = 1
Private Const ConnectSTARTTLS = 2
Private Const ConnectDirectSSL = 3
Private Const ConnectTryTLS = 4

Private Const Notification_None = 0
Private Const Notification_OnSuccess = 2
Private Const Notification_OnFailure = 4
Private Const Notification_Delay = 8
Private Const Notification_Never = 16

Private Const CRYPT_MACHINE_KEYSET = 32
Private Const CRYPT_USER_KEYSET = 4096
Private Const CERT_SYSTEM_STORE_CURRENT_USER = 65536
Private Const CERT_SYSTEM_STORE_LOCAL_MACHINE = 131072
Private Const CERT_STORE_PROV_LDAP_STORE = 16

Private mCanc As Boolean
Private mErro As Boolean
Private mIdle As Boolean

Private Const mChar = "utf-8"
Private Const EmExt = ".eml"

Private RS110 As ADODB.Recordset
Private RS120 As ADODB.Recordset
Private RS130 As ADODB.Recordset
Private RS140 As ADODB.Recordset
Private RS150 As ADODB.Recordset
Private FL101 As ADODB.Field

Private MaiGeTo As EAGetMailObjLib.Tools
Private MaiGeUm As EAGetMailObjLib.UIDLManager
Private MaiGeSe As EAGetMailObjLib.MailServer
Private MaiInCo As EAGetMailObjLib.MailInfoCollection

Private WithEvents MaiSend As EASendMailObjLib.Mail
Attribute MaiSend.VB_VarHelpID = -1
Private WithEvents MaiGeCl As EAGetMailObjLib.MailClient
Attribute MaiGeCl.VB_VarHelpID = -1

Private clFil As clsFile

Public Function MaiGet() As Integer
On Error GoTo StErr

Dim PatNr As Long
Dim MitNr As Long
Dim MaSiz As Long
Dim MaDat As Date
Dim AlDat As Date
Dim MesID As String
Dim MaGui As String
Dim SenNa As String
Dim EmAdr As String
Dim EmBet As String
Dim EmKom As String
Dim EmEmp As String
Dim TrKey As String
Dim PaStr As String
Dim SpStr As String
Dim TmSt1 As String
Dim TmSt2 As String
Dim FiNam As String
Dim DaNam As String
Dim DaPfa As String
Dim EmNam As String
Dim EmPfa As String
Dim UDLFi As String
Dim MaVor As Boolean
Dim FlGef As Boolean
Dim AttVo As Boolean
Dim MaGel As Boolean
Dim MaAkt As Integer
Dim MaGes As Integer
Dim MaNeu As Integer
Dim MaZal As Integer
Dim FltZa As Integer
Dim AdAkt As Integer
Dim AdGes As Integer
Dim AtAkt As Integer
Dim AtGes As Integer
Dim MaPri As Integer

Set FM = frmStatus
Set Lab01 = FM.lblLab01
Set PrBr1 = FM.prbStat1
Set PrBr2 = FM.prbStat2

Set clFil = New clsFile

Set MaiGeUm = New EAGetMailObjLib.UIDLManager
Set MaiGeSe = New EAGetMailObjLib.MailServer
Set MaiGeCl = New EAGetMailObjLib.MailClient
Set MaiGeTo = New EAGetMailObjLib.Tools

MaiGeCl.LicenseCode = "EG-C1653719494-01199-8BVCAA171B372F1D-F797V24BDU55166A"

EmPfa = GlDpf & mFold & "\"
DaPfa = GlDpf & "MailDat" & "\"
UDLFi = DaPfa & mUIDn

If clFil.FilDir(EmPfa) = False Then
    MkDir EmPfa
Else
    If clFil.FilDir(DaPfa) = False Then
        MkDir DaPfa
    End If
End If
DoEvents 'Emailfoldercheck

If GlLog = True Then 'General Logging
    MaiGeCl.LogFileName = DaPfa & Format$(GlMiA(GlSmI, 2), "0000") & "EmGetMail.log"
End If

mCanc = False

With MaiGeSe
    Select Case mPosf
    Case 1:
        .Server = mIMAP
        .protocol = 1
    Case 2:
        .Server = mPOP3
        .protocol = 0
    End Select
    .User = mUsNa
    .Password = mUsPa
    If mTLSg = 1 Then
        .SSLConnection = True
    Else
        .SSLConnection = False
    End If
    .AuthType = mAuTy
    .SSLType = mTLSg
    .Port = mPor1
    .ProxyProtocol = 1 'Socks5
    If GlSSp <> vbNullString Then
        .SocksProxyServer = GlSSp
        If GlSPo <> vbNullString Then
            .SocksProxyPort = GlSPo
        End If
    End If
End With

MaiGeCl.Connect MaiGeSe
DoEvents 'Verbindungsaufbau

MaiGeCl.GetMailInfosParam.Reset
DoEvents

If mPosf = 1 Then 'IMAP
    If mAlAk = True Then
        MaiGeCl.GetMailInfosParam.GetMailInfosOptions = (MaiGeCl.GetMailInfosParam.GetMailInfosOptions Or 2)
    End If
    If mAlTy > 0 Then
        MaiGeCl.GetMailInfosParam.GetMailInfosOptions = (MaiGeCl.GetMailInfosParam.GetMailInfosOptions Or 64)
        Select Case mAlTy
        Case 1: AlDat = DateAdd("d", -7, Now())
        Case 2: AlDat = DateAdd("d", -30, Now())
        Case 3: AlDat = DateAdd("d", -60, Now())
        Case 4: AlDat = DateAdd("d", -90, Now())
        End Select
        MaiGeCl.GetMailInfosParam.DateRangeSINCE = AlDat
        MaiGeCl.GetMailInfosParam.DateRangeBEFORE = DateAdd("d", 1, Now())
    End If
End If

Set MaiInCo = MaiGeCl.GetMailInfoList()
DoEvents

If mUIDn <> vbNullString Then
    MaiGeUm.Load UDLFi 'UIDL Textfile
    MaiGeUm.SyncUIDLEX MaiGeSe, MaiInCo
    MaiGeUm.Update
End If

MaGes = MaiInCo.Count

If MaGes > 0 Then
    For MaAkt = 0 To MaGes - 1
        Dim MaiInf2 As EAGetMailObjLib.MailInfo
        Dim MaiUDI2 As EAGetMailObjLib.UIDLItem
        Set MaiInf2 = MaiInCo.Item(MaAkt)
        If mUIDn <> vbNullString Then
            Set MaiUDI2 = MaiGeUm.FindUIDL(MaiGeSe, MaiInf2.UIDL)
            If MaiUDI2 Is Nothing Then
                MaNeu = MaNeu + 1
            End If
        End If
    Next MaAkt

    If MaNeu > 0 Then
        If mProg = True Then
            Load FM
            FM.Show
            DoEvents
        
            PrBr1.Min = 0
            PrBr1.Max = MaNeu
            PrBr2.Min = 0
            PrBr2.Max = 100
            
            Lab01.Caption = "Emailempfang..."
        End If
    
        Set RS110 = New ADODB.Recordset
        With RS110
            .CursorLocation = adUseClient
            .Source = "qryMailInAd"
            .ActiveConnection = DB1
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open Options:=adCmdTableDirect
        End With
        
        If RS110.Supports(adAddNew) = True Then
            For MaAkt = 0 To MaGes - 1
    
                Dim MaiInfo As EAGetMailObjLib.MailInfo
                Dim MaiUDIt As EAGetMailObjLib.UIDLItem
                Dim MaiAdCo As EAGetMailObjLib.AddressCollection
                Dim MaiAtCo As EAGetMailObjLib.AttachmentCollection
                
                EmKom = vbNullString
                EmNam = vbNullString
                DaNam = vbNullString
                FiNam = vbNullString
                EmAdr = vbNullString
                EmEmp = vbNullString
                EmBet = vbNullString
                PaStr = vbNullString
                SenNa = vbNullString
                TmSt1 = vbNullString
                TmSt2 = vbNullString
                SpStr = vbNullString
                TrKey = mTrKy
                FlGef = False
                PatNr = 0
                                
                Set MaiInfo = MaiInCo.Item(MaAkt)
                MaSiz = CLng(MaiInfo.SIZE)
                If mPosf = 2 Then 'POP3
                    MaGel = False
                Else
                    MaGel = MaiInfo.Read
                End If
        
                If mUIDn <> vbNullString Then
                    Set MaiUDIt = MaiGeUm.FindUIDL(MaiGeSe, MaiInfo.UIDL)
                End If
                
                If MaiInfo.UIDL <> vbNullString Then
                    MesID = MaiInfo.UIDL & MaiGeTo.GenFileName(MaAkt)
                Else
                    MesID = CreateID("M")
                End If
    
                Set RS120 = New ADODB.Recordset 'Email still available
                RS120.CursorLocation = adUseClient
                Set RS120 = DBCmRe1("qrySimMaInId", "@pMailId", MesID)
                If RS120.Fields.Item(0).Value > 0 Then ' count the mail messages with the same id
                    MaVor = True
                Else
                    MaVor = False
                End If
                RS120.Close
                Set RS120 = Nothing
                            
                If MaiUDIt Is Nothing Then
                    If MaVor = False Then
                        If MaGel = False Then
                    
                            MaZal = MaZal + 1
                            Lab01.Caption = "Empfangen: " & Format$(MaZal + 1, "000") & " / " & Format$(MaNeu, "000") & " ..."
                            PrBr1.Value = MaAkt + 1
                            DoEvents
        
                            Dim MaiMail As EAGetMailObjLib.Mail
                            Set MaiMail = MaiGeCl.GetMail(MaiInfo)
                            
                            MaGui = CreateID("M")
                            DaNam = MaGui & EmExt
                            EmNam = MaiGeTo.GenFileName(MaAkt) & EmExt
                            FiNam = EmPfa & DaNam
        
                            MaiMail.SaveAs FiNam, True
                            DoEvents
            
                            If MaiMail.From.Address = vbNullString Then
                                EmAdr = "keine Emailadresse"
                                SenNa = "keine Emailadresse"
                            Else
                                EmAdr = MaiMail.From.Address
                                If Len(EmAdr) > 250 Then
                                    EmAdr = Left$(EmAdr, 250)
                                End If
                                If MaiMail.From.Name <> vbNullString Then
                                    SenNa = MaiMail.From.Name
                                    SenNa = Replace(SenNa, Chr$(124), vbNullString, 1)
                                    If Len(SenNa) > 250 Then
                                        SenNa = Left$(SenNa, 250)
                                    End If
                                Else
                                    SenNa = EmAdr
                                End If
                            End If
        
                            If MaiMail.Subject = vbNullString Then
                                EmBet = "kein Betreff"
                            Else
                                EmBet = MaiMail.Subject
                                If Len(EmBet) > 250 Then
                                    EmBet = Left$(EmBet, 250)
                                End If
                            End If
            
                            If IsDate(MaiMail.ReceivedDate) = True Then
                                MaDat = CDate(MaiMail.ReceivedDate)
                            Else
                                MaDat = Date
                            End If
            
                            If MaiMail.Priority > 0 Then
                                MaPri = MaiMail.Priority
                            Else
                                MaPri = 3
                            End If
                            
                            Set MaiAdCo = MaiMail.ToList
                            AdGes = MaiAdCo.Count
                            
                            Set MaiAtCo = MaiMail.AttachmentList
                            AtGes = MaiAtCo.Count
            
                            If AtGes > 0 Then
                                AttVo = True
                            Else
                                AttVo = False
                            End If
            
                            If EmAdr <> "keine Emailadresse" Then
                                For FltZa = 1 To UBound(GlMft) 'Mailfilter Adresse
                                    If GlMft(FltZa, 3) <> vbNullString Then
                                        If LCase(EmAdr) = LCase(GlMft(FltZa, 3)) Then
                                            TrKey = "o" & GlMft(FltZa, 1) & "o"
                                            SpStr = LCase(GlMft(FltZa, 3))
                                            FlGef = True
                                            Exit For
                                        End If
                                    End If
                                Next FltZa
                            End If
                
                            If FlGef = False Then
                                If EmBet <> "kein Betreff" Then
                                    For FltZa = 1 To UBound(GlMft) 'Mailfilter Betreff
                                        If LCase(GlMft(FltZa, 2)) <> vbNullString Then
                                            If LCase(EmBet) Like LCase(GlMft(FltZa, 2)) Then
                                                TrKey = "o" & GlMft(FltZa, 1) & "o"
                                                SpStr = LCase(GlMft(FltZa, 2))
                                                FlGef = True
                                                Exit For
                                            End If
                                        End If
                                    Next FltZa
                                End If
                            End If
                
                            If EmAdr <> "keine Emailadresse" Then
                                Set RS130 = New ADODB.Recordset 'v
                                RS130.CursorLocation = adUseClient
                                Set RS130 = DBCmRe1("qrySimAdEma", "@IdStr", EmAdr)
                                If RS130.RecordCount > 0 Then
                                    PatNr = RS130.Fields("ID0").Value
                                    PaStr = RS130.Fields("IDKurz").Value
                                End If
                                RS130.Close
                                Set RS130 = Nothing
                
                                If PatNr = 0 Then
                                    Set RS140 = New ADODB.Recordset 'Patientenfilter
                                    RS140.CursorLocation = adUseClient
                                    Set RS140 = DBCmRe1("qrySimAdEm2", "@IdStr", EmAdr)
                                    If RS140.RecordCount > 0 Then
                                        PatNr = RS140.Fields("ID0").Value
                                        PaStr = RS140.Fields("IDKurz").Value
                                    End If
                                    RS140.Close
                                    Set RS140 = Nothing
                                End If
                
                                If PatNr = 0 Then
                                    Set RS150 = New ADODB.Recordset 'Zugehörigenflter
                                    RS150.CursorLocation = adUseClient
                                    Set RS150 = DBCmRe1("qryMailAdFil", "@IdStr", EmAdr)
                                    If RS150.RecordCount > 0 Then
                                        PatNr = RS150.Fields("ID0").Value
                                        PaStr = RS150.Fields("IDKurz").Value
                                    End If
                                    RS150.Close
                                    Set RS150 = Nothing
                                End If
                            End If

                            Select Case TrKey
                            Case "o801o": MitNr = GlMiA(GlSmI, 2)
                            Case "o802o": MitNr = 0
                            Case "o803o": MitNr = 0
                            Case "o804o": MitNr = 0
                            Case "o805o": MitNr = GlMiA(GlSmI, 2)
                            Case Else: MitNr = GlMiA(GlSmI, 2)
                            End Select
                            
                            For AdAkt = 0 To AdGes - 1
                                If EmEmp = vbNullString Then
                                    EmEmp = MaiAdCo.Item(AdAkt).Address
                                Else
                                    EmEmp = EmEmp & ";" & MaiAdCo.Item(AdAkt).Address
                                End If
                            Next AdAkt
            
                            If Len(EmEmp) > 250 Then
                                EmEmp = Left$(EmEmp, 250)
                            End If
            
                            If MaiMail.TextBody = vbNullString Then
                                EmKom = "kein Emailkommentar"
                            Else
                                EmKom = MaiMail.TextBody
                                EmKom = Replace(EmKom, vbCrLf, Chr$(32), 1)
                                EmKom = LTrim$(EmKom)
                                EmKom = RTrim$(EmKom)
                                If Len(EmKom) > 250 Then
                                    EmKom = Left$(EmKom, 250)
                                End If
                            End If
                
                            If SpStr <> vbNullString Then
                                EmBet = "<" & SpStr & "> " & EmBet
                                If Len(EmBet) > 250 Then
                                    EmBet = Left$(EmBet, 250)
                                End If
                            End If
                
                            RS110.AddNew
                            RS110.Fields("IDM").Value = MitNr
                            RS110.Fields("ID0").Value = PatNr
                            RS110.Fields("GuiID").Value = MaGui
                            RS110.Fields("TreKey").Value = TrKey
                            RS110.Fields("SenderName").Value = SenNa
                            RS110.Fields("SenderEMail").Value = EmAdr
                            RS110.Fields("Subject").Value = EmBet
                            RS110.Fields("Empfaenger").Value = EmEmp
                            RS110.Fields("Mailsize").Value = SMSize(MaSiz)
                            RS110.Fields("Maildate").Value = Format$(MaDat, "dd.mm.yyyy")
                            RS110.Fields("Mailtime").Value = Format$(MaDat, "hh:mm")
                            RS110.Fields("Priority").Value = MaPri
                            RS110.Fields("Attachment").Value = AttVo
                            RS110.Fields("MailFile").Value = DaNam
                            RS110.Fields("MailId").Value = MesID
                            RS110.Fields("Marker").Value = 0
                            RS110.Fields("Kommentar").Value = EmKom
                            RS110.Fields("Patient").Value = PaStr
                            If TrKey = "o805o" Then
                                RS110.Fields("Spammail").Value = -1
                            End If
                            RS110.Update
        
                            If mUIDn <> vbNullString Then
                                If mNoDe = True Then 'No Emaildelete on Server
                                    MaiGeUm.AddUIDL MaiGeSe, MaiInfo.UIDL, EmNam
                                    If mGele = True Then 'Als gelesen kennzeichnen
                                        If mPosf = 1 Then 'IMAP4
                                            If MaiInfo.Read = False Then
                                                MaiGeCl.MarkAsRead MaiInfo, True
                                            End If
                                        End If
                                    End If
                                Else
                                    MaiGeCl.Delete MaiInfo
                                    MaiGeUm.RemoveUIDL MaiGeSe, MaiInfo.UIDL
                                End If
                            End If
            
                            Set MaiMail = Nothing

                            If GlEKr = True Then 'Emails in Krankenblatt dokumentieren
                                If PatNr > 0 Then
                                    GlNeK = GlKoX
                                    With GlNeK
                                        .PatNr = PatNr
                                        .IdxNr = 0
                                        .EiDat = Format$(Date, "dd.mm.yyyy")
                                        .EiZei = TimeValue(Now)
                                        .EiTyp = 108 'Email
                                        .KoStr = EmKom
                                        .KoGui = MaGui
                                        .NeuEi = True
                                        .Mitar = GlMiA(GlSmI, 2)
                                    End With
                                    K_Einf
                                End If
                            End If
                        End If
                    End If
                End If
    
                Set MaiUDIt = Nothing
                Set MaiInfo = Nothing
                Set MaiAdCo = Nothing
                Set MaiAtCo = Nothing
            Next
        End If
        
        If mUIDn <> vbNullString Then
            MaiGeUm.Update
        End If
        
        RS110.Close
        Set RS110 = Nothing
        
        If mProg = True Then
            Unload FM
            DoEvents
        End If
    Else
        Select Case mPosf
        Case 1: SPopu mIMAP, mUsNa & " - keine neuen Emails.", IC48_Information
        Case 2: SPopu mPOP3, mUsNa & " - keine neuen Emails.", IC48_Information
        End Select
    End If
Else
    Select Case mPosf
    Case 1: SPopu mIMAP, mUsNa & " - keine neuen Emails.", IC48_Information
    Case 2: SPopu mPOP3, mUsNa & " - keine neuen Emails.", IC48_Information
    End Select
End If

MaiGeCl.Quit

Set MaiGeUm = Nothing
Set MaiGeSe = Nothing
Set MaiGeCl = Nothing
Set MaiGeTo = Nothing

Set clFil = Nothing

MaiGet = 0

Exit Function

StErr:
If GlDbg = True Then
    SErLog Err.Description & " MaiGet " & Err.Number & " - " & Err.Source
End If
Resume Next

End Function
Public Function MaiSen() As Integer
On Error GoTo StErr

Dim TmStr As String
Dim ImStr As String
Dim DaPfa As String
Dim AktZa As Integer
Dim GesZa As Integer

Set FM = frmStatus
Set Lab01 = FM.lblLab01
Set PrBr1 = FM.prbStat1
Set PrBr2 = FM.prbStat2

DaPfa = GlDpf & "MailDat"

If mAtVo = True Then
    GesZa = UBound(GlAtt)
End If

Set MaiSend = New EASendMailObjLib.Mail

If mProg = True Then
    Load FM
    FM.Show
    DoEvents

    PrBr2.Min = 0
    PrBr2.Max = 100
    
    Lab01.Caption = "Emailversand..."
End If

With MaiSend
    .LicenseCode = "ES-E1646458156-01154-3FA5U38BE448BU9D-38EF11A12331U9D9"
    .Reset
    DoEvents
    If GlLog = True Then 'General Logging
        .CreateFolder DaPfa
        .LogFileName = DaPfa & "\" & Format$(GlMiA(GlSmI, 2), "0000") & "EmSendMail.log"
    End If
    If mAnon = True Then
        .Anonymous = 1
    Else
        .Anonymous = 0
    End If
    .AuthType = -1
    .AutoCalendar = 1
    .ConnectType = ConnectSSLAuto
    .KeepConnection = 0
    .EncryptionAlgorithm = 0 'R C2
    .HeaderEncoding = 1 'Base64
    .TransferEncoding = 4 'Base64
    .Timeout = 10
    .Asynchronous = mAsyn
    .RecipientsCerts.Clear
    .protocol = 0 'Standard
    .ServerAddr = mSMTP
    .ServerPort = mPor2
    If mAuUs = True Then
        .Username = mUsNa
        .Password = mUsPa
    End If
    .Priority = mPrio
    .ReadReceipt = mBest
    .DeliveryNotification = Notification_None
    .Charset = mChar
    .From = mSeNa
    .FromAddr = mSeAd
    .Subject = mSubj
    .ReplyTo = mSeRe
    .ReturnPath = mSeRe
    .SaveCopy = False
    .BodyFormat = mFrmt
    '.AddHeader "auto-submitted", "auto-generated"
    .SignerCert.Unload
    .ProxyProtocol = 1 'Socks5
    If GlSSp <> vbNullString Then
        .SocksProxyServer = GlSSp
        If GlSPo <> vbNullString Then
            .SocksProxyPort = GlSPo
        End If
    End If
    If GlSPn <> vbNullString Then
        .ComputerName = GlSPn
        If GlSIP <> vbNullString Then
            .LocalIP = GlSIP
        End If
    ElseIf mSeNa <> vbNullString Then
        .ComputerName = SNaFi(mSeNa, True, True, True)
    End If
    If InStr(1, mReAd, ";", 1) > 0 Then
        .AddRecipientEx mReAd, 0
    Else
        .AddRecipient mReNa, mReAd, 0
    End If
    If InStr(1, mReCc, ";", 1) > 0 Then
        .AddRecipientEx mReCc, 1
    Else
        .AddRecipient vbNullString, mReCc, 1
    End If
    If InStr(1, mReBc, ";", 1) > 0 Then
        .AddRecipientEx mReBc, 2
    Else
        .AddRecipient vbNullString, mReBc, 2
    End If
    If mImag <> vbNullString Then
        ImStr = .AddInline(mImag) 'Signaturdatei
        If ImStr = vbNullString Then
            SPopu "Bilddateifehler", .GetLastErrDescription, IC48_Forbidden
        End If
    End If
    If mFrmt = 1 Then 'HTML
        If mTxHt <> vbNullString Then
            If ImStr <> vbNullString Then
                mTxHt = Replace(mTxHt, "</html>", vbNullString, 1)
                mTxHt = mTxHt & "<br>" & vbCrLf & "<img src=""cid:" & ImStr & """>" & vbCrLf & "</html>"
            End If
            .BodyText = mTxHt
        Else
            If mTxAc <> vbNullString Then
                .BodyText = mTxHt
            End If
        End If
        If mTxAc <> vbNullString Then
            .AltBody = mTxAc
        End If
    Else
        If mTxAc <> vbNullString Then
            .BodyText = mTxAc
        End If
    End If
    If mAtVo = True Then 'Attachments
        If GesZa > 0 Then
            For AktZa = 1 To GesZa
                TmStr = vbNullString
                TmStr = MTrim(GlAtt(AktZa), ";")
                .AddAttachment TmStr
            Next AktZa
        End If
    ElseIf mAtFi <> vbNullString Then
        .AddAttachment mAtFi
    End If
    If mFiln <> vbNullString Then
        .SaveMail mFiln
    End If
    DoEvents
    .SSL_starttls = mTLSs
    DoEvents
    If .SSL_init() <> 0 Then
        Select Case mTLSs
        Case 0: 'SSL/TLS
            .SSL_starttls = 1
            If .SSL_init() <> 0 Then
                .SSL_starttls = 2
                If .SSL_init() <> 0 Then
                    Exit Function
                Else
                    mTLSs = 2
                End If
            Else
                mTLSs = 1
            End If
        Case 1: 'STARTTLS
            .SSL_starttls = 2
            If .SSL_init() <> 0 Then
                .SSL_starttls = 0
                If .SSL_init() <> 0 Then
                    Exit Function
                Else
                    mTLSs = 0
                End If
            Else
                mTLSs = 2
            End If
        Case 2: 'TryTLS
            .SSL_starttls = 0
            If .SSL_init() <> 0 Then
                .SSL_starttls = 1
                If .SSL_init() <> 0 Then
                    Exit Function
                Else
                    mTLSs = 1
                End If
            Else
                mTLSs = 0
            End If
        End Select
    End If
    DoEvents
    mCanc = False
    mErro = False
    mIdle = False
    .SendMail
End With

Do While Not mIdle
    DoEvents
Loop

If mCanc = True Then
    MaiSen = 1
ElseIf mErro = True Then
    MaiSen = 2
Else
    MaiSen = 0
End If

If mProg = True Then
    Unload FM
    DoEvents
End If

With MaiSend
    .ClearAttachment
    .SSL_uninit
End With

Set MaiSend = Nothing

Exit Function

StErr:
If GlDbg = True Then
    MsgBox Err.Description, 48, "MaiSen " & Err.Number
    SErLog Err.Description & " MaiSen " & Err.Number & " - " & Err.Source
End If
Resume Next

End Function
Private Function MTrim(ByVal StrVa As String, ByVal SuStr As String) As String
On Error Resume Next

'Trim the right
Do While Right$(StrVa, 1) Like "[" & SuStr & "]"
    StrVa = Left$(StrVa, Len(StrVa) - 1)
Loop

'Trim the left
Do While Left$(StrVa, 1) Like "[" & SuStr & "]"
    StrVa = Mid(StrVa, 2)
Loop

'Return the result
MTrim = StrVa

End Function
Public Property Let moAnon(Flag As Boolean)
    mAnon = Flag
End Property
Public Property Get moAnon() As Boolean
    moAnon = mAnon
End Property
Public Property Let moAsyn(Flag As Integer)
    mAsyn = Flag
End Property
Public Property Get moAsyn() As Integer
    moAsyn = mAsyn
End Property
Public Property Let moAlAk(Flag As Boolean)
    mAlAk = Flag
End Property
Public Property Get moAlAk() As Boolean
    moAlAk = mAlAk
End Property
Public Property Let moAlTy(Flag As Integer)
    mAlTy = Flag
End Property
Public Property Get moAlTy() As Integer
    moAlTy = mAlTy
End Property
Public Property Let moAtFi(Flag As String)
    mAtFi = Flag
End Property
Public Property Get moAtFi() As String
    moAtFi = mAtFi
End Property
Public Property Let moAtVo(Flag As Boolean)
    mAtVo = Flag
End Property
Public Property Get moAtVo() As Boolean
    moAtVo = mAtVo
End Property
Public Property Let moAuTy(Flag As Integer)
    mAuTy = Flag
End Property
Public Property Get moAuTy() As Integer
    moAuTy = mAuTy
End Property
Public Property Let moAuUs(Flag As Boolean)
    mAuUs = Flag
End Property
Public Property Get moAuUs() As Boolean
    moAuUs = mAuUs
End Property
Public Property Let moBest(Flag As Boolean)
    mBest = Flag
End Property
Public Property Get moBest() As Boolean
    moBest = mBest
End Property
Public Property Let moCoNa(Flag As String)
    mCoNa = Flag
End Property
Public Property Get moCoNa() As String
    moCoNa = mCoNa
End Property
Public Property Let moCoIP(Flag As String)
    mCoIP = Flag
End Property
Public Property Get moCoIP() As String
    moCoIP = mCoIP
End Property

Public Property Let moHELO(Flag As String)
    mHELO = Flag
End Property
Public Property Get moHELO() As String
    moHELO = mHELO
End Property
Public Property Let moFiln(Flag As String)
    mFiln = Flag
End Property
Public Property Get moFiln() As String
    moFiln = mFiln
End Property
Public Property Let moFold(Flag As String)
    mFold = Flag
End Property
Public Property Get moFold() As String
    moFold = mFold
End Property
Public Property Let moIMAP(Flag As String)
    mIMAP = Flag
End Property
Public Property Get moIMAP() As String
    moIMAP = mIMAP
End Property
Public Property Let moImag(Flag As String)
    mImag = Flag
End Property
Public Property Get moImag() As String
    moImag = mImag
End Property
Public Property Let moNoDe(Flag As Boolean)
    mNoDe = Flag
End Property
Public Property Get moNoDe() As Boolean
    moNoDe = mNoDe
End Property
Public Property Let moPOP3(Flag As String)
    mPOP3 = Flag
End Property
Public Property Get moPOP3() As String
    moPOP3 = mPOP3
End Property
Public Property Let moPor1(Flag As Integer)
    mPor1 = Flag
End Property
Public Property Get moPor1() As Integer
    moPor1 = mPor1
End Property
Public Property Let moPor2(Flag As Integer)
    mPor2 = Flag
End Property
Public Property Get moPor2() As Integer
    moPor2 = mPor2
End Property
Public Property Let moPosf(Flag As Integer)
    mPosf = Flag
End Property
Public Property Get moPosf() As Integer
    moPosf = mPosf
End Property
Public Property Let moPrio(Flag As Integer)
    mPrio = Flag
End Property
Public Property Get moPrio() As Integer
    moPrio = mPrio
End Property
Public Property Let moProg(Flag As Boolean)
    mProg = Flag
End Property
Public Property Get moProg() As Boolean
    moProg = mProg
End Property
Public Property Let moPxSe(Flag As String)
    mPxSe = Flag
End Property
Public Property Get moPxSe() As String
    moPxSe = mPxSe
End Property
Public Property Let moPxPa(Flag As Integer)
    mPxPa = Flag
End Property
Public Property Get moPxPa() As Integer
    moPxPa = mPxPa
End Property
Public Property Let moPxPo(Flag As String)
    mPxPo = Flag
End Property
Public Property Get moPxPo() As String
    moPxPo = mPxPo
End Property
Public Property Let moPxUs(Flag As String)
    mPxUs = Flag
End Property
Public Property Get moPxUs() As String
    moPxUs = mPxUs
End Property
Public Property Let moReNa(Flag As String)
    mReNa = Flag
End Property
Public Property Get moReNa() As String
    moReNa = mReNa
End Property
Public Property Let moReAd(Flag As String)
    mReAd = Flag
End Property
Public Property Get moReAd() As String
    moReAd = mReAd
End Property
Public Property Let moReBc(Flag As String)
    mReBc = Flag
End Property
Public Property Get moReBc() As String
    moReBc = mReBc
End Property
Public Property Let moReCc(Flag As String)
    mReCc = Flag
End Property
Public Property Get moReCc() As String
    moReCc = mReCc
End Property
Public Property Let moGele(Flag As Boolean)
    mGele = Flag
End Property
Public Property Get moGele() As Boolean
    moGele = mGele
End Property
Public Property Let moSubj(Flag As String)
    mSubj = Flag
End Property
Public Property Get moSubj() As String
    moSubj = mSubj
End Property
Public Property Let moFrmt(Flag As Integer)
    mFrmt = Flag
End Property
Public Property Get moFrmt() As Integer
    moFrmt = mFrmt
End Property
Public Property Let moSeAd(Flag As String)
    mSeAd = Flag
End Property
Public Property Get moSeAd() As String
    moSeAd = mSeAd
End Property
Public Property Let moSeNa(Flag As String)
    mSeNa = Flag
End Property
Public Property Get moSeNa() As String
    moSeNa = mSeNa
End Property
Public Property Let moSeRe(Flag As String)
    mSeRe = Flag
End Property
Public Property Get moSeRe() As String
    moSeRe = mSeRe
End Property
Public Property Let moSMTP(Flag As String)
    mSMTP = Flag
End Property
Public Property Get moSMTP() As String
    moSMTP = mSMTP
End Property
Public Property Let moTLSg(Flag As Integer)
    mTLSg = Flag
End Property
Public Property Get moTLSg() As Integer
    moTLSg = mTLSg
End Property
Public Property Let moTrKy(Flag As String)
    mTrKy = Flag
End Property
Public Property Get moTrKy() As String
    moTrKy = mTrKy
End Property
Public Property Let moTLSs(Flag As Integer)
    mTLSs = Flag
End Property
Public Property Get moTLSs() As Integer
    moTLSs = mTLSs
End Property
Public Property Let moTxAc(Flag As String)
    mTxAc = Flag
End Property
Public Property Get moTxAc() As String
    moTxAc = mTxAc
End Property
Public Property Let moTxHt(Flag As String)
    mTxHt = Flag
End Property
Public Property Get moTxHt() As String
    moTxHt = mTxHt
End Property
Public Property Let moUIDn(Flag As String)
    mUIDn = Flag
End Property
Public Property Get moUIDn() As String
    moUIDn = mUIDn
End Property
Public Property Let moUsNa(Flag As String)
    mUsNa = Flag
End Property
Public Property Get moUsNa() As String
    moUsNa = mUsNa
End Property
Public Property Let moUsPa(Flag As String)
    mUsPa = Flag
End Property
Public Property Get moUsPa() As String
    moUsPa = mUsPa
End Property

Private Sub Class_Terminate()
    Set FM = Nothing
End Sub

Private Sub MaiGeCl_OnAuthorized(ByVal oSender As Object, Cancel As Boolean)
    Cancel = mCanc
End Sub
Private Sub MaiGeCl_OnConnected(ByVal oSender As Object, Cancel As Boolean)
On Error Resume Next

Dim EmKto As String

If mIMAP <> vbNullString Then
    EmKto = mIMAP
ElseIf mPOP3 <> vbNullString Then
    EmKto = mPOP3
Else
    EmKto = "Kein Mailkonto"
End If

If mProg = True Then
    Lab01.Caption = "Verbunden (" & EmKto & " - " & mUsNa & ")..."
    DoEvents
End If

Cancel = mCanc

End Sub
Private Sub MaiGeCl_OnIdle(ByVal oSender As Object, Cancel As Boolean)
    DoEvents
    Cancel = mCanc
End Sub
Private Sub MaiGeCl_OnQuit(ByVal oSender As Object, Cancel As Boolean)
    mIdle = True
End Sub

Private Sub MaiGeCl_OnReceivingDataStream(ByVal oSender As Object, ByVal oInfo As Object, ByVal Received As Long, ByVal Total As Long, Cancel As Boolean)
On Error Resume Next

If mProg = True Then
    PrBr2.Value = (Received / Total) * 100
    DoEvents
End If

Cancel = mCanc

End Sub
Private Sub MaiGeCl_OnSecuring(ByVal oSender As Object, Cancel As Boolean)
On Error Resume Next

If mProg = True Then
    Lab01.Caption = "Sicherheitsabfrage..."
    DoEvents
End If

Cancel = mCanc

End Sub

Private Sub MaiSend_OnAuthenticated()
    DoEvents
End Sub
Private Sub MaiSend_OnClosed()
    mIdle = True
End Sub
Private Sub MaiSend_OnConnected()
On Error Resume Next

Dim EmKto As String

If mSMTP <> vbNullString Then
    EmKto = mSMTP
Else
    EmKto = "Kein Mailkonto"
End If

If mProg = True Then
    Lab01.Caption = "Verbunden (" & EmKto & ")..."
    DoEvents
End If

End Sub
Private Sub MaiSend_OnError(ByVal lError As Long, ByVal ErrDescription As String)
On Error Resume Next

Dim Mld1, Mld2, Tit1 As String

Set FM = frmStatus

If mProg = True Then
    Lab01.Caption = "Versandfehler..."
    Unload frmStatus
    DoEvents
End If

mErro = True
mIdle = True

Tit1 = "E-Mail-Konto Fehler"
Mld1 = "Das Passwort: " & mUsPa & " passt nicht oder nicht mehr zum Benutzernamen: " & mUsNa & " und muss bei den E-Mail-Konten angepasst werden."
Mld2 = "Eine Einstellung zum SMTP-Server: " & mSMTP & " Ihres E-Mail-Kontos ist nicht oder nicht mehr korrekt und muss bei den E-Mail-Konten angepasst werden."

If InStr(1, LCase(ErrDescription), "username", 1) > 0 Then
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
ElseIf InStr(1, LCase(ErrDescription), "password", 1) > 0 Then
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
ElseIf InStr(1, LCase(ErrDescription), "authentication", 1) > 0 Then
    WindowMess Mld2, Dial2, Tit1, FM.hwnd
Else
    SPopu Tit1, ErrDescription, IC48_Forbidden
End If
DoEvents

SErLog ErrDescription & " - " & lError

End Sub
Private Sub MaiSend_OnSending(ByVal lSent As Long, ByVal lTotal As Long)
On Error Resume Next

If mProg = True Then
    If lSent = 0 Then
        PrBr2.Min = 0
        PrBr2.Max = lTotal
    Else
        PrBr2.Value = lSent
    End If
End If

End Sub
