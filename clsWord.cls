VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHwnd As Long
Private OffOb As Object
Private mPatNr As String 'Patientennummer
Private mRefNr As String 'Referenznummer
Private mTeFax As String 'Telefaxnummer
Private mAnred As String 'Anrede
Private mRAnre As String 'Anrede
Private mAnsch As String 'Anschrift
Private mBrief As String 'Briefanrede
Private mRBrie As String 'Briefanrede
Private mGebur As String 'Geburtsdatum
Private mTelNu As String 'Telefonnummer
Private mFirma As String 'Firma
Private mNaNam As String 'Nachname
Private mRNaNa As String 'Nachname
Private mVorna As String 'Vorname
Private mRVorn As String 'Vorname
Private mStras As String 'Straße
Private mPostl As String 'Postleitzahl
Private mWoOrt As String 'Wohnort
Private mLaKen As String 'Landeskennung
Private mLands As String 'Land
Private mDiagn As String 'Diagnose
Private mAnamn As String 'Anamnese
Private mLabor As String 'Laborbericht
Private mLeben As String 'Lebensalter
Private mRechn As String 'Rechnungsposition
Private mKraPo As String 'Krankenblattpositionen
Private mBefun As String 'Befundung
Private mKopZe As String 'Kopfzeile
Private mFusZe As String 'Fusszeile
Private mTerBe As String 'Terminbetreff
Private mTerBt As String 'Terminberag
Private mTerDa As String 'Termindatum Anfang
Private mTerEn As String 'Termindatum Ende
Private mTeZeV As String 'Terminzeit Anfang
Private mTeZeB As String 'Terminzeit Ende
Private mSerBt As String 'Serienterminberag
Private mSerDa As String 'Serientermindatum Anfang
Private mSerEn As String 'Serientermindatum Ende
Private mVersi As String 'Versicherung
Private mThera As String 'Therapeut
Private mVerFx As String 'Telefax versicherung
Private mVerNr As String 'Versicherungsnummer
Private mKraNr As String 'Krankenkassennummer
Private mVorPf As String
Private mDocPf As String
Private mFiNam As String 'Vorlagenname
Private mDaNam As String 'Speichername
Private mAcPrn As String
Private mEmEmp As String
Private mEmBtr As String
Private mEmTxt As String
Private mOfVer As Integer 'Office Version
Private mOpTyp As Integer 'Arte Word zu starten
Private mAnzFe As Integer
Private mAnzTe As Integer
Private mSerEt As Boolean
Private mSerEm As Boolean
Private mSchut As Boolean
Private mDocSa As Boolean 'Dokument Speichern

Public Event WorEnd(Flag As Boolean)

Private Const wdDialogFileNew = 79
Private Const wdDialogFileOpen = 80
Private Const wdDialogFilePageSetup = 178
Private Const wdDialogFilePrint = 88
Private Const wdDialogFilePrintSetup = 79
Private Const wdDialogFileSaveAs = 84
Private Const wdDialogInsertBookmark = 168
Private Const wdDialogMailMerge = 676 ' Nur wenn Felder eingefügt sind
Private Const wdDialogMailMergeInsertAsk = 4047
Private Const wdDialogMailMergeOpenDataSource = 81
Private Const wdDialogMailMergeHelper = 680
Private Const wdDialogInsertMergeField = 167
Private Const wdSendToNewDocument = 0
Private Const wdSendToEmail = 2
Private Const wdDialogToolsCreateLabels = 489

Private Const wdGoToBookmark = -1
Private Const wdGoToComment = 6
Private Const wdGoToField = 7
Private Const wdGoToGraphic = 8
Private Const wdGoToLine = 3
Private Const wdGoToObject = 9
Private Const wdGoToPage = 1
Private Const wdGoToPercent = 12
Private Const wdGoToSection = 0
Private Const wdGoToTable = 2
Private Const wdGoToRelative = 2

Private Const wdNotAMergeDocument = -1
Private Const wdFormLetters = 0
Private Const wdMailingLabels = 1
Private Const wdEnvelopes = 2
Private Const wdCatalog = 3
Private Const wdNormalDocument = 0
Private Const wdMainDocumentOnly = 1
Private Const wdMainAndDataSource = 2
Private Const wdMainAndHeader = 3
Private Const wdMainAndSourceAndHeader = 4
Private Const wdDataSource = 5
Private Const wdOpenFormatAuto = 0
Private Const wdOpenFormatDocument = 1
Private Const wdOpenFormatTemplate = 2
Private Const wdOpenFormatRTF = 3
Private Const wdOpenFormatText = 4
Private Const wdOpenFormatEncodedText = 5
Private Const wdOpenFormatAllWord = 6
Private Const wdOpenFormatWebPages = 7

Private Const wdCurrentFolderPath = 14
Private Const wdDocumentsPath = 0
Private Const wdPicturesPath = 1
Private Const wdProgramPath = 9
Private Const wdStartupPath = 8
Private Const wdUserTemplatesPath = 2
Private Const wdPrintAllDocument = 0
Private Const wdDoNotSaveChanges = 0
Private Const wdPromptToSaveChanges = -2
Private Const wdSaveChanges = -1
Private Const wdNewEmailMessage = 2
Private Const wdOneAfterAnother = 0
Private Const wdCustomLabelA4 = 2
Private Const xlOneAfterAnother = 1
Private Const xlAllAtOnce = 2
Private Const xlNormal = -4143
Private Const xlWindows = 2
Private Const xlDelimited = 1
Private Const xlFixedWidth = 2
Private Const xlTextQualifierDoubleQuote = 1
Private Const xlTextQualifierSingleQuote = 2
Private Const xlTextQualifierNone = -4142

Private Const wdWindowStateNormal = 0
Private Const wdWindowStateMaximize = 1
Private Const wdWindowStateMinimize = 2

Private Const SW_SHOW = 5
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function SearchTree Lib "imagehlp.dll" Alias "SearchTreeForFile" (ByVal lpRootPath As String, ByVal lpInputPathName As String, ByVal lpOutputPathBuffer As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Property Let ActPrn(ByVal Flag As String)
Attribute ActPrn.VB_Description = "Aktueller Drucker mit Portausgabe"
    mAcPrn = Flag
End Property

Public Property Get ActPrn() As String
    ActPrn = mAcPrn
End Property

Public Property Let AnzFel(Flag As Integer)
Attribute AnzFel.VB_Description = "Anzahl gefundener Seriendruckfelder im Dokument"
    mAnzFe = Flag
End Property

Public Property Get AnzFel() As Integer
    AnzFel = mAnzFe
End Property

Public Property Let AnzTex(Flag As Integer)
Attribute AnzTex.VB_Description = "Anzahl gefundener Textmarken im Dokument"
    mAnzTe = Flag
End Property

Public Property Get AnzTex() As Integer
    AnzTex = mAnzTe
End Property

Public Property Get DatNam() As String
Attribute DatNam.VB_Description = "Seriendrucksteuerdatei"
    DatNam = mDaNam
End Property
Public Property Let DatNam(ByVal Flag As String)
    mDaNam = Flag
End Property
Public Property Get DocSav() As Boolean
Attribute DocSav.VB_Description = "Speichert das Dokument nach übergabe benötigt DaNam"
    DocSav = mDocSa
End Property
Public Property Let DocSav(ByVal Flag As Boolean)
    mDocSa = Flag
End Property
Public Property Get DoSchu() As Boolean
Attribute DoSchu.VB_Description = "Schreibschutz des geöffneten Dokumentes"
    DoSchu = mSchut
End Property
Public Property Let DoSchu(ByVal Flag As Boolean)
    mSchut = Flag
End Property
Public Property Let EmBetr(ByVal Flag As String)
Attribute EmBetr.VB_Description = "Betreff der Email"
    mEmBtr = Flag
End Property
Public Property Get EmBetr() As String
    EmBetr = mEmBtr
End Property

Public Property Let EmEmpf(ByVal Flag As String)
Attribute EmEmpf.VB_Description = "Emailadresse des Empfängers"
    mEmEmp = Flag
End Property

Public Property Get EmEmpf() As String
    EmEmpf = mEmEmp
End Property
Public Property Let EmText(ByVal Flag As String)
Attribute EmText.VB_Description = "Emailtext"
    mEmTxt = Flag
End Property
Public Property Get EmText() As String
    EmText = mEmTxt
End Property

Public Property Get FilNam() As String
Attribute FilNam.VB_Description = "Dateiname"
    FilNam = mFiNam
End Property

Public Property Let FilNam(ByVal Flag As String)
    mFiNam = Flag
End Property

Public Property Let FlAnam(ByVal Flag As String)
Attribute FlAnam.VB_Description = "AAnamnese"
    mAnamn = Flag
End Property

Public Property Get FlAnam() As String
    FlAnam = mAnamn
End Property

Public Property Get FlAnsch() As String
Attribute FlAnsch.VB_Description = "Anschrift"
    FlAnsch = mAnsch
End Property
Public Property Let FlAnsch(ByVal Flag As String)
    mAnsch = Flag
End Property
Public Property Get FlAnred() As String
    FlAnred = mAnred
End Property
Public Property Let FlAnred(ByVal Flag As String)
    mAnred = Flag
End Property
Public Property Get FlRAnre() As String
    FlRAnre = mRAnre
End Property
Public Property Let FlRAnre(ByVal Flag As String)
    mRAnre = Flag
End Property
Public Property Let FlBefun(ByVal Flag As String)
Attribute FlBefun.VB_Description = "Befundung"
    mBefun = Flag
End Property

Public Property Get FlBefun() As String
    FlBefun = mBefun
End Property

Public Property Get FlBrief() As String
Attribute FlBrief.VB_Description = "Briefanrede"
    FlBrief = mBrief
End Property
Public Property Let FlBrief(ByVal Flag As String)
    mBrief = Flag
End Property
Public Property Get FlRBrie() As String
    FlRBrie = mRBrie
End Property
Public Property Let FlRBrie(ByVal Flag As String)
    mRBrie = Flag
End Property
Public Property Let FlDiagn(ByVal Flag As String)
Attribute FlDiagn.VB_Description = "Diagnose"
    mDiagn = Flag
End Property

Public Property Get FlDiagn() As String
    FlDiagn = mDiagn
End Property

Public Property Let FlFirm(ByVal Flag As String)
Attribute FlFirm.VB_Description = "Firma / Institution"
    mFirma = Flag
End Property

Public Property Get FlFirm() As String
    FlFirm = mFirma
End Property

Public Property Let FlFuss(ByVal Flag As String)
Attribute FlFuss.VB_Description = "Fusszeile"
    mFusZe = Flag
End Property

Public Property Get FlFuss() As String
    FlFuss = mFusZe
End Property

Public Property Get FlGebur() As String
Attribute FlGebur.VB_Description = "Geburtsdatum"
    FlGebur = mGebur
End Property

Public Property Let FlGebur(ByVal Flag As String)
    mGebur = Flag
End Property

Public Property Let FlKopf(ByVal Flag As String)
Attribute FlKopf.VB_Description = "Kopfzeile"
    mKopZe = Flag
End Property

Public Property Get FlKopf() As String
    FlKopf = mKopZe
End Property

Public Property Let FlLabor(ByVal Flag As String)
Attribute FlLabor.VB_Description = "Laborbericht"
    mLabor = Flag
End Property
Public Property Get FlLabor() As String
    FlLabor = mLabor
End Property
Public Property Let FlLaKe(ByVal Flag As String)
Attribute FlLaKe.VB_Description = "Landeskennung"
    mLaKen = Flag
End Property
Public Property Get FlLaKe() As String
    FlLaKe = mLaKen
End Property

Public Property Let FlLand(ByVal Flag As String)
Attribute FlLand.VB_Description = "Landesname"
    mLands = Flag
End Property
Public Property Get FlLeben() As String
Attribute FlLeben.VB_Description = "Lebensalter"
    FlLeben = mLeben
End Property
Public Property Let FlLeben(ByVal Flag As String)
    mLeben = Flag
End Property
Public Property Get FlLand() As String
    FlLand = mLands
End Property

Public Property Get FlNaNam() As String
Attribute FlNaNam.VB_Description = "Nachname"
    FlNaNam = mNaNam
End Property
Public Property Let FlNaNam(ByVal Flag As String)
    mNaNam = Flag
End Property
Public Property Get FlRNaNa() As String
    FlRNaNa = mRNaNa
End Property
Public Property Let FlRNaNa(ByVal Flag As String)
    mRNaNa = Flag
End Property
Public Property Get FlPatNr() As String
    FlPatNr = mPatNr
End Property
Public Property Let FlPatNr(ByVal Flag As String)
    mPatNr = Flag
End Property
Public Property Get FlRefNr() As String
    FlRefNr = mRefNr
End Property
Public Property Let FlRefNr(ByVal Flag As String)
    mRefNr = Flag
End Property
Public Property Get FlPostl() As String
Attribute FlPostl.VB_Description = "PLZ"
    FlPostl = mPostl
End Property
Public Property Let FlPostl(ByVal Flag As String)
    mPostl = Flag
End Property
Public Property Let FlRechn(ByVal Flag As String)
Attribute FlRechn.VB_Description = "Rechnungsposition"
    mRechn = Flag
End Property

Public Property Get FlRechn() As String
    FlRechn = mRechn
End Property

Public Property Get FlStras() As String
Attribute FlStras.VB_Description = "Straße"
    FlStras = mStras
End Property

Public Property Let FlStras(ByVal Flag As String)
    mStras = Flag
End Property
Public Property Get FlTerBe() As String
Attribute FlTerBe.VB_Description = "Terminbetreff"
    FlTerBe = mTerBe
End Property
Public Property Let FlTerBe(ByVal Flag As String)
    mTerBe = Flag
End Property
Public Property Get FlTerBt() As String
    FlTerBt = mTerBt
End Property
Public Property Let FlTerBt(ByVal Flag As String)
    mTerBt = Flag
End Property
Public Property Get FlTerDa() As String
Attribute FlTerDa.VB_Description = "Termindatum"
    FlTerDa = mTerDa
End Property
Public Property Let FlTerDa(ByVal Flag As String)
    mTerDa = Flag
End Property
Public Property Get FlSerBt() As String
    FlSerBt = mSerBt
End Property
Public Property Let FlSerBt(ByVal Flag As String)
    mSerBt = Flag
End Property
Public Property Get FlSerDa() As String
    FlSerDa = mSerDa
End Property
Public Property Let FlSerDa(ByVal Flag As String)
    mSerDa = Flag
End Property
Public Property Get FlTerEn() As String
    FlTerEn = mTerEn
End Property
Public Property Let FlTerEn(ByVal Flag As String)
    mTerEn = Flag
End Property
Public Property Get FlSerEn() As String
    FlSerEn = mSerEn
End Property
Public Property Let FlSerEn(ByVal Flag As String)
    mSerEn = Flag
End Property
Public Property Get FlTerZe() As String
Attribute FlTerZe.VB_Description = "Terminzeit"
    FlTerZe = mTeZeV
End Property
Public Property Let FlTerZe(ByVal Flag As String)
    mTeZeV = Flag
End Property
Public Property Get FlTerBi() As String
    FlTerBi = mTeZeB
End Property
Public Property Let FlTerBi(ByVal Flag As String)
    mTeZeB = Flag
End Property
Public Property Get FlTeFax() As String
Attribute FlTeFax.VB_Description = "Telefax"
    FlTeFax = mTeFax
End Property
Public Property Let FlTeFax(ByVal Flag As String)
    mTeFax = Flag
End Property

Public Property Get FlTelNu() As String
Attribute FlTelNu.VB_Description = "Telefon"
    FlTelNu = mTelNu
End Property

Public Property Let FlTelNu(ByVal Flag As String)
    mTelNu = Flag
End Property

Public Property Get FlVorna() As String
Attribute FlVorna.VB_Description = "Vorname"
    FlVorna = mVorna
End Property
Public Property Let FlVorna(ByVal Flag As String)
    mVorna = Flag
End Property
Public Property Get FlRVorn() As String
    FlRVorn = mRVorn
End Property
Public Property Let FlRVorn(ByVal Flag As String)
    mRVorn = Flag
End Property
Public Property Get FlVersi() As String
Attribute FlVersi.VB_Description = "Versicherungsanschrift"
    FlVersi = mVersi
End Property
Public Property Let FlVersi(ByVal Flag As String)
    mVersi = Flag
End Property
Public Property Get FlVerFx() As String
    FlVerFx = mVerFx
End Property
Public Property Let FlVerFx(ByVal Flag As String)
    mVerFx = Flag
End Property
Public Property Get FlVerNr() As String
    FlVerNr = mVerNr
End Property
Public Property Let FlVerNr(ByVal Flag As String)
    mVerNr = Flag
End Property
Public Property Get FlKraNr() As String
    FlKraNr = mKraNr
End Property
Public Property Let FlKraNr(ByVal Flag As String)
    mKraNr = Flag
End Property
Public Property Get FlThera() As String
    FlThera = mThera
End Property
Public Property Let FlThera(ByVal Flag As String)
    mThera = Flag
End Property
Public Property Get FlWoOrt() As String
Attribute FlWoOrt.VB_Description = "Ort"
    FlWoOrt = mWoOrt
End Property

Public Property Let FlWoOrt(ByVal Flag As String)
    mWoOrt = Flag
End Property
Public Function FOutlo() As Object
On Error Resume Next
'Öffnet ein Outlookobjekt

Set OffOb = GetObject(, "Outlook.Application")

If Err.Number = 429 Then
    Err.Clear

    Set OffOb = CreateObject("Outlook.Application")
End If
    
Set OffOb = Nothing

End Function
Public Sub FWoDia(ByVal Flag As Integer)
Attribute FWoDia.VB_Description = "Öffnet einen Worddialog 1:Drucken 2:Serienfelder 3:Etiketten 4:Fileopen 5:Textmarken"
On Error Resume Next
'Öffnet einen Worddialog 1:Drucken 2:Serienfelder 3:Etiketten 4:Fileopen 5:Textmarken

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = GetObject(, "Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

DoEvents
With OffOb
    Select Case Flag
    Case 1: .Dialogs(wdDialogFilePrint).Show
    Case 2: .Dialogs(wdDialogInsertMergeField).Show
    Case 3: .Dialogs(wdDialogMailMergeHelper).Show
    Case 4: .Dialogs(wdDialogFileOpen).Show
    Case 5: .Dialogs(wdDialogInsertBookmark).Show
    End Select
End With

Set OffOb = Nothing

End Sub

Public Sub FWoDru()
Attribute FWoDru.VB_Description = "Öffnet ein Dokument in Word und druckt dieses aus, benötigt FilNam und OPeTyp und optional ActPrn"
On Error Resume Next
'Öffnet ein Dokument in Word und druckt dieses aus, benötigt FilNam

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = GetObject(, "Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

DoEvents
With OffOb
    .Options.PrintBackground = True
    .ActivePrinter = mAcPrn
    .ActiveDocument.ActiveWindow.PrintOut Background:=True, Range:=wdPrintAllDocument
    DoEvents
    If .BackgroundPrintingStatus = 0 Then
        .ActiveDocument.ActiveWindow.Close SaveChanges:=wdDoNotSaveChanges
    End If
End With

Set OffOb = Nothing

End Sub

Public Sub FWoEml()
Attribute FWoEml.VB_Description = "Versendet eine Email mit Word, benötigt EmEmpf, EmBetr, optional EmText und FilNam für Attachment und OPeTyp"
On Error Resume Next
'Versendet eine Email mit Word, benötigt EmEmpf, EmBetr, optional EmText und FilNam für Attachment

Dim ReWe As Long
Dim Posi As Integer
Dim FiTy As String

FiTy = LCase(Right$(mFiNam, 4))
If mEmBtr = vbNullString Then mEmBtr = mFiNam
If mEmTxt = vbNullString Then mEmTxt = mFiNam

If FiTy = "doc" Or FiTy = "dot" Or FiTy = "txt" Or FiTy = "docx" Then
    Select Case mOpTyp
    Case 1: Set OffOb = CreateObject("Word.Application")
    Case 2: Set OffOb = GetObject(, "Word.Application")
    End Select
    DoEvents
    With OffOb
        .Visible = False
        .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        .ActiveDocument.HasRoutingSlip = True
        With .ActiveDocument.RoutingSlip
            .Subject = mEmBtr
            .Message = mEmTxt
            .AddRecipient Recipient:=mEmEmp
            .Delivery = wdOneAfterAnother
        End With
        .ActiveDocument.Route
        .ActiveDocument.ActiveWindow.Close SaveChanges:=wdDoNotSaveChanges
    End With
    Set OffOb = Nothing
ElseIf FiTy = "xls" Or FiTy = "xlt" Or FiTy = "xlsx" Then
    Select Case mOpTyp
    Case 1: Set OffOb = CreateObject("Excel.Application")
    Case 2: Set OffOb = GetObject(, "Excel.Application")
    End Select
    DoEvents
    With OffOb
        .Visible = False
        .Workbooks.Open FileName:=mFiNam
        .ActiveWorkbook.HasRoutingSlip = True
        With .ActiveWorkbook.RoutingSlip
            .Subject = mEmBtr
            .Message = mEmTxt
            .Recipients = mEmEmp
            .Delivery = xlOneAfterAnother
        End With
        .ActiveWorkbook.Route
        .ActiveWindow.Close SaveChanges:=False
    End With
    Set OffOb = Nothing
Else
    ReWe = ShellExecute(mHwnd, "open", "mailto:" & mEmEmp & "?subject=" & mEmBtr & "&body=" & mEmTxt, vbNullString, vbNullString, SW_SHOW)
End If

End Sub
Public Function FWoExp() As Boolean
Attribute FWoExp.VB_Description = "Öffnet die angegebene Ascii-Textdatei in Excel, benötigt FilNam und OPeTyp"
On Error Resume Next
'Öffnet die angegebene Ascii-Textdatei in Excel, benötigt FilNam und OPeTyp

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Excel.Application")
Case 2: Set OffOb = GetObject(, "Excel.Application")
End Select

DoEvents
With OffOb
    .Workbooks.OpenText mFiNam, Origin:=xlWindows, StartRow:=1, dataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True
    DoEvents
    .Visible = True
End With

Set OffOb = Nothing

End Function
Public Function FWoFin() As String
Attribute FWoFin.VB_Description = "Liefert das Installationsverzeichnis von Word"
On Error Resume Next
'Liefert das Installationsverzeichnis von Word

Dim FilNa As String

If regKeyExist("Software\Microsoft\Office\8.0\Word\Options") = False Then
    If regKeyExist("Software\Microsoft\Office\9.0\Word\Options") = False Then
        If regKeyExist("Software\Microsoft\Office\9.0\Word\InstallRoot") = False Then
            If regKeyExist("Software\Microsoft\Office\10.0\Word\Options") = False Then
                If regKeyExist("Software\Microsoft\Office\10.0\Word\InstallRoot") = False Then
                    If regKeyExist("Software\Microsoft\Office\11.0\Word\Options") = False Then
                        If regKeyExist("Software\Microsoft\Office\11.0\Word\InstallRoot") = False Then
                            If regKeyExist("Software\Microsoft\Office\12.0\Word\Options") = False Then
                                If regKeyExist("Software\Microsoft\Office\12.0\Word\InstallRoot") = False Then
                                    If FWoVor("c:\programme\microsoft office\office\", "winword.exe") = vbNullString Then
                                        FilNa = "-"
                                    Else
                                        FilNa = "c:\programme\microsoft office\office\winword.exe"
                                    End If
                                Else
                                    FilNa = regReadKey("Software\Microsoft\Office\12.0\Word\InstallRoot", "PATH") & "winword.exe"
                                End If
                            Else
                                FilNa = regReadKey("Software\Microsoft\Office\12.0\Word\Options", "PROGRAMDIR") & "winword.exe"
                            End If
                        Else
                            FilNa = regReadKey("Software\Microsoft\Office\11.0\Word\InstallRoot", "PATH") & "winword.exe"
                        End If
                    Else
                        FilNa = regReadKey("Software\Microsoft\Office\11.0\Word\Options", "PROGRAMDIR") & "winword.exe"
                    End If
                Else
                    FilNa = regReadKey("Software\Microsoft\Office\10.0\Word\InstallRoot", "PATH") & "winword.exe"
                End If
            Else
                FilNa = regReadKey("Software\Microsoft\Office\10.0\Word\Options", "PROGRAMDIR") & "winword.exe"
            End If
        Else
            FilNa = regReadKey("Software\Microsoft\Office\9.0\Word\InstallRoot", "PATH") & "winword.exe"
        End If
    Else
        FilNa = regReadKey("Software\Microsoft\Office\9.0\Word\Options", "PROGRAMDIR") & "winword.exe"
    End If
Else
    FilNa = regReadKey("Software\Microsoft\Office\8.0\Word\Options", "PROGRAMDIR") & "winword.exe"
End If

FWoFin = FilNa

End Function
Public Sub FWoNeu()
Attribute FWoNeu.VB_Description = "Öffnet Word mit einer neuen bestehenden Vorlage und ohne Vorlage, benötigt  OPeTyp und optional FilNam"
On Error Resume Next
'Öffnet Word mit einer neuen bestehenden Vorlage und ohne Vorlage, benötigt optional FilNam

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = GetObject(, "Word.Application")
End Select
DoEvents

With OffOb
    If LCase(Right$(mFiNam, 3)) = "dot" Then
        .Documents.Add Template:=mFiNam, NewTemplate:=False
    Else
        .Documents.Add
    End If
    DoEvents
    .Visible = True
End With

Set OffOb = Nothing

End Sub

Public Sub FWoOpn(ByVal Flag As Integer)
Attribute FWoOpn.VB_Description = "Öffnet Word über den Shell Befehl, benötigt FilNam"
On Error Resume Next
'Öffnet Word über den Shell Befehl, benötigt FilNam

Dim SheWe As Long
Dim RetWe As Long
Dim FeHan As Long
Dim PfNam As String

RaiseEvent WorEnd(False)

PfNam = FWoFin
If PfNam <> vbNullString Then
    Select Case Flag
    Case 1: SheWe = ShellExecute(mHwnd, "open", PfNam & Chr$(32) & mFiNam & vbNullString, vbNullString, vbNullString, SW_NORMAL)
    Case 2: SheWe = ShellExecute(mHwnd, "open", PfNam & Chr$(32) & mFiNam & vbNullString, vbNullString, vbNullString, SW_MAXIMIZE)
    Case 3: SheWe = ShellExecute(mHwnd, "open", PfNam & Chr$(32) & mFiNam & vbNullString, vbNullString, vbNullString, SW_MINIMIZE)
    End Select
    FeHan = OpenProcess(SYNCHRONIZE, False, RetWe)
    WaitForSingleObject FeHan, INFINITE
    CloseHandle FeHan
End If

RaiseEvent WorEnd(True)

End Sub

Public Sub FWoPfa()
Attribute FWoPfa.VB_Description = "Liefert Informationen über Wordoptionen: wVorPf, wDocPf"
On Error Resume Next
'Liefert Informationen über Wordoptionen: wVorPf, wDocPf

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = GetObject(, "Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

DoEvents
With OffOb
    .Visible = False
    mVorPf = .Options.DefaultFilePath(wdUserTemplatesPath) & "\"
    mDocPf = .Options.DefaultFilePath(wdDocumentsPath) & "\"
    .Quit
End With

Set OffOb = Nothing

End Sub
Public Function FWoRel() As Integer
On Error Resume Next
'Zeigt die Office Releasenummer

Dim OfStr As String
Dim Posit As Integer

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = GetObject(, "Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

If OffOb.Application.Build <> vbNullString Then
    OfStr = OffOb.Application.Build
    Posit = InStr(1, OfStr, ".", 1)
    If Posit > 0 Then
        If IsNumeric(Left$(OfStr, Posit - 1)) Then
            mOfVer = CInt(Left$(OfStr, Posit - 1))
        Else
            mOfVer = 0
        End If
    Else
        If IsNumeric(Left$(OfStr, Posit - 1)) Then
            mOfVer = CInt(Left$(OfStr, Posit - 1))
        Else
            mOfVer = 0
        End If
    End If
Else
    mOfVer = 0
End If

FWoRel = mOfVer

End Function
Public Sub FWoSer()
Attribute FWoSer.VB_Description = "Öffnet Word für den Serienbriefdruck, benötigt FilNam (auch Vorlagen) und DatNam als Steuerdatei sowie OpeTyp. SerEti und SerEml optional"
On Error Resume Next
'Öffnet Word für den Serienbriefdruck, benötigt FilNam (auch Vorlagen) und DatNam

Dim AnzFe As Integer
Dim AktFe As Integer
Dim TesFe As Boolean

TesFe = False

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = CreateObject("Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

DoEvents
With OffOb
    Select Case mOpTyp
    Case 1:
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    Case 2:
        .Documents.Add
    Case 3:
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    End Select
    DoEvents
    With .ActiveDocument.MailMerge
        If mSerEt = True Then
            .MainDocumentType = wdMailingLabels
        Else
            .MainDocumentType = wdFormLetters
        End If
        .OpenDataSource Name:=mDaNam, Format:=wdOpenFormatAuto, ConfirmConversions:=True, ReadOnly:=mSchut, LinkToSource:=True, AddToRecentFiles:=False
        If Err.Number = 0 Then
            .SuppressBlankLines = True
            .ViewMailMergeFieldCodes = True
            AnzFe = .DataSource.FieldNames.Count
            If AnzFe > 0 Then
                For AktFe = 1 To AnzFe
                    If .DataSource.FieldNames(AktFe).Name = "Anschrift" Then
                        TesFe = True
                        Exit For
                    End If
                Next AktFe
                If TesFe = True Then
                    mAnzFe = .Fields.Count
                    If mAnzFe > 0 Then
                        .Destination = wdSendToNewDocument
                        If mSerEm = True Then
                            .MailAddressFieldName = "Email"
                            .MailAsAttachment = True
                            .Destination = wdSendToEmail
                            .MailSubject = mFiNam
                            .Destination = wdSendToEmail
                            .Execute
                        Else
                            .Execute
                        End If
                    End If
                End If
            End If
        End If
    End With
    DoEvents
    If mDocSa = True Then
        If mDaNam <> vbNullString Then
            .ActiveDocument.SaveAs mDaNam
        End If
    End If
    DoEvents
    .Visible = True
    DoEvents
    .WindowState = wdWindowStateMaximize
End With

Set OffOb = Nothing

End Sub
Public Sub FWoSta()
Attribute FWoSta.VB_Description = "Öffnet Word mit einem Dokument und übergibt keine Variablen, benötigt FilNam und OpeTyp"
On Error Resume Next
'Öffnet Word mit einem Dokument

Dim AktBo As Integer

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = CreateObject("Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

DoEvents
With OffOb
    Select Case mOpTyp
    Case 1:
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    Case 2:
        .Documents.Add
        .Selection.GOTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=7
        .Selection.InsertAfter mAnsch
    Case 3:
        DoEvents
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    End Select
    DoEvents
    If mDocSa = True Then
        If mDaNam <> vbNullString Then
            .ActiveDocument.SaveAs mDaNam
        End If
    End If
    DoEvents
    .Visible = True
    DoEvents
    .WindowState = wdWindowStateMaximize
End With

Set OffOb = Nothing

End Sub
Public Sub FWoUbe()
Attribute FWoUbe.VB_Description = "Öffnet Word mit einem Dokument und übergibt Variablen für die Textmarken, benötigt FilNam und OpeTyp"
On Error Resume Next
'Öffnet Word mit einem Dokument und übergibt Variablen für die Textmarken, benötigt FilNam

Dim AktBo As Integer

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = CreateObject("Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

DoEvents
With OffOb
    Select Case mOpTyp
    Case 1:
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    Case 2:
        .Documents.Add
        .Selection.GOTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=7
        .Selection.InsertAfter mAnsch
    Case 3:
        DoEvents
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    End Select
    DoEvents
    mAnzTe = .ActiveDocument.Bookmarks.Count
    If mAnzTe > 0 Then
        For AktBo = 1 To mAnzTe
            Select Case .ActiveDocument.Bookmarks(AktBo).Name
            Case "Nummer":
                    .ActiveDocument.Bookmarks("Nummer").Select
                    .Selection.InsertAfter mPatNr
            Case "Referenznummer":
                    .ActiveDocument.Bookmarks("Referenznummer").Select
                    .Selection.InsertAfter mPatNr
            Case "Referenznummer2":
                    .ActiveDocument.Bookmarks("Referenznummer2").Select
                    .Selection.InsertAfter mPatNr
            Case "Telefax":
                    .ActiveDocument.Bookmarks("Telefax").Select
                    .Selection.InsertAfter mTeFax
            Case "Anschrift":
                    .ActiveDocument.Bookmarks("Anschrift").Select
                    .Selection.InsertAfter mAnsch
            Case "Anschrift2":
                    .ActiveDocument.Bookmarks("Anschrift2").Select
                    .Selection.InsertAfter mAnsch
            Case "Briefanrede":
                    .ActiveDocument.Bookmarks("Briefanrede").Select
                    .Selection.InsertAfter mBrief
            Case "Briefanrede2":
                    .ActiveDocument.Bookmarks("Briefanrede2").Select
                    .Selection.InsertAfter mBrief
            Case "RBriefanrede":
                    .ActiveDocument.Bookmarks("RBriefanrede").Select
                    .Selection.InsertAfter mRBrie
            Case "RBriefanrede2":
                    .ActiveDocument.Bookmarks("RBriefanrede2").Select
                    .Selection.InsertAfter mRBrie
            Case "Anrede":
                    .ActiveDocument.Bookmarks("Anrede").Select
                    .Selection.InsertAfter mAnred
            Case "Anrede2":
                    .ActiveDocument.Bookmarks("Anrede2").Select
                    .Selection.InsertAfter mAnred
            Case "RAnrede":
                    .ActiveDocument.Bookmarks("RAnrede").Select
                    .Selection.InsertAfter mRAnre
            Case "RAnrede2":
                    .ActiveDocument.Bookmarks("RAnrede2").Select
                    .Selection.InsertAfter mRAnre
            Case "Geburtstag":
                    .ActiveDocument.Bookmarks("Geburtstag").Select
                    .Selection.InsertAfter mGebur
            Case "Geburtstag2":
                    .ActiveDocument.Bookmarks("Geburtstag2").Select
                    .Selection.InsertAfter mGebur
            Case "Lebensalter":
                    .ActiveDocument.Bookmarks("Lebensalter").Select
                    .Selection.InsertAfter mLeben
            Case "Telefon":
                    .ActiveDocument.Bookmarks("Telefon").Select
                    .Selection.InsertAfter mTelNu
            Case "Firma":
                    .ActiveDocument.Bookmarks("Firma").Select
                    .Selection.InsertAfter mFirma
            Case "Firma2":
                    .ActiveDocument.Bookmarks("Firma2").Select
                    .Selection.InsertAfter mFirma
            Case "Name":
                    .ActiveDocument.Bookmarks("Name").Select
                    .Selection.InsertAfter mNaNam
            Case "Name2":
                    .ActiveDocument.Bookmarks("Name2").Select
                    .Selection.InsertAfter mNaNam
            Case "RName":
                    .ActiveDocument.Bookmarks("RName").Select
                    .Selection.InsertAfter mRNaNa
            Case "RName2":
                    .ActiveDocument.Bookmarks("RName2").Select
                    .Selection.InsertAfter mRNaNa
            Case "Vorname":
                    .ActiveDocument.Bookmarks("Vorname").Select
                    .Selection.InsertAfter mVorna
            Case "Vorname2":
                    .ActiveDocument.Bookmarks("Vorname2").Select
                    .Selection.InsertAfter mVorna
            Case "Vorname3":
                    .ActiveDocument.Bookmarks("Vorname3").Select
                    .Selection.InsertAfter mVorna
            Case "Vorname4":
                    .ActiveDocument.Bookmarks("Vorname4").Select
                    .Selection.InsertAfter mVorna
            Case "RVorname":
                    .ActiveDocument.Bookmarks("RVorname").Select
                    .Selection.InsertAfter mRVorn
            Case "RVorname2":
                    .ActiveDocument.Bookmarks("RVorname2").Select
                    .Selection.InsertAfter mRVorn
            Case "Straße":
                    .ActiveDocument.Bookmarks("Straße").Select
                    .Selection.InsertAfter mStras
            Case "Straße2":
                    .ActiveDocument.Bookmarks("Straße2").Select
                    .Selection.InsertAfter mStras
            Case "PLZ":
                    .ActiveDocument.Bookmarks("PLZ").Select
                    .Selection.InsertAfter mPostl
            Case "PLZ2":
                    .ActiveDocument.Bookmarks("PLZ2").Select
                    .Selection.InsertAfter mPostl
            Case "Ort":
                    .ActiveDocument.Bookmarks("Ort").Select
                    .Selection.InsertAfter mWoOrt
            Case "Ort2":
                    .ActiveDocument.Bookmarks("Ort2").Select
                    .Selection.InsertAfter mWoOrt
            Case "Land":
                    .ActiveDocument.Bookmarks("Land").Select
                    .Selection.InsertAfter mLands
            Case "Land2":
                    .ActiveDocument.Bookmarks("Land2").Select
                    .Selection.InsertAfter mLands
            Case "LK":
                    .ActiveDocument.Bookmarks("LK").Select
                    .Selection.InsertAfter mLaKen
            Case "Termindatum":
                    .ActiveDocument.Bookmarks("Termindatum").Select
                    .Selection.InsertAfter mTerDa
            Case "Termindatum2":
                    .ActiveDocument.Bookmarks("Termindatum2").Select
                    .Selection.InsertAfter mTerDa
            Case "Terminendet":
                    .ActiveDocument.Bookmarks("Terminendet").Select
                    .Selection.InsertAfter mTerEn
            Case "Terminendet2":
                    .ActiveDocument.Bookmarks("Terminendet2").Select
                    .Selection.InsertAfter mTerEn
            Case "Serienterminbetrag":
                    .ActiveDocument.Bookmarks("Serienterminbetrag").Select
                    .Selection.InsertAfter mSerBt
            Case "Serienterminbetrag2":
                    .ActiveDocument.Bookmarks("Serienterminbetrag2").Select
                    .Selection.InsertAfter mSerBt
            Case "Serienterminbegin":
                    .ActiveDocument.Bookmarks("Serienterminbegin").Select
                    .Selection.InsertAfter mSerDa
            Case "Serienterminbegin2":
                    .ActiveDocument.Bookmarks("Serienterminbegin2").Select
                    .Selection.InsertAfter mSerDa
            Case "Serienterminendet":
                    .ActiveDocument.Bookmarks("Serienterminendet").Select
                    .Selection.InsertAfter mSerEn
            Case "Serienterminendet2":
                    .ActiveDocument.Bookmarks("Serienterminendet2").Select
                    .Selection.InsertAfter mSerEn
            Case "Terminzeit":
                    .ActiveDocument.Bookmarks("Terminzeit").Select
                    .Selection.InsertAfter mTeZeV
            Case "Terminzeit2":
                    .ActiveDocument.Bookmarks("Terminzeit2").Select
                    .Selection.InsertAfter mTeZeV
            Case "Terminende":
                    .ActiveDocument.Bookmarks("Terminende").Select
                    .Selection.InsertAfter mTeZeB
            Case "Terminende2":
                    .ActiveDocument.Bookmarks("Terminende2").Select
                    .Selection.InsertAfter mTeZeB
            Case "Terminbetreff":
                    .ActiveDocument.Bookmarks("Terminbetreff").Select
                    .Selection.InsertAfter mTerBe
            Case "Terminbetreff2":
                    .ActiveDocument.Bookmarks("Terminbetreff2").Select
                    .Selection.InsertAfter mTerBe
            Case "Terminbetrag":
                    .ActiveDocument.Bookmarks("Terminbetrag").Select
                    .Selection.InsertAfter mTerBt
            Case "Terminbetrag2"
                    .ActiveDocument.Bookmarks("Terminbetrag2").Select
                    .Selection.InsertAfter mTerBt
            Case "Versicherung":
                    .ActiveDocument.Bookmarks("Versicherung").Select
                    .Selection.InsertAfter mVersi
            Case "Versicherungsfax":
                    .ActiveDocument.Bookmarks("Versicherungsfax").Select
                    .Selection.InsertAfter mVerFx
            Case "Versicherungsnummer":
                    .ActiveDocument.Bookmarks("Versicherungsnummer").Select
                    .Selection.InsertAfter mVerNr
            Case "Krankenkassennummer":
                    .ActiveDocument.Bookmarks("Krankenkassennummer").Select
                    .Selection.InsertAfter mVerNr
            Case "Diagnose":
                    .ActiveDocument.Bookmarks("Diagnose").Select
                    .Selection.InsertAfter mDiagn
            Case "Anamnese":
                    .ActiveDocument.Bookmarks("Anamnese").Select
                    .Selection.InsertAfter mAnamn
            Case "Bericht":
                    .ActiveDocument.Bookmarks("Bericht").Select
                    .Selection.InsertAfter mLabor
            Case "Rechnung":
                    .ActiveDocument.Bookmarks("Rechnung").Select
                    .Selection.InsertAfter mRechn
            Case "Krankenblatt":
                    .ActiveDocument.Bookmarks("Krankenblatt").Select
                    .Selection.InsertAfter mKraPo
            Case "Befund":
                    .ActiveDocument.Bookmarks("Befund").Select
                    .Selection.InsertAfter mBefun
            Case "Kopfzeile":
                    .ActiveDocument.Bookmarks("Kopfzeile").Select
                    .Selection.InsertAfter mKopZe
            Case "Fusszeile":
                    .ActiveDocument.Bookmarks("Fusszeile").Select
                    .Selection.InsertAfter mFusZe
            Case "Therapeut":
                    .ActiveDocument.Bookmarks("Therapeut").Select
                    .Selection.InsertAfter mThera
            Case "Therapeut2":
                    .ActiveDocument.Bookmarks("Therapeut2").Select
                    .Selection.InsertAfter mThera
            Case "Emailadresse":
                    .ActiveDocument.Bookmarks("Emailadresse").Select
                    .Selection.InsertAfter mEmEmp
            End Select
        Next AktBo
    End If
    DoEvents
    .ActiveDocument.GOTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=22
    DoEvents
    If mDocSa = True Then
        If mDaNam <> vbNullString Then
            .ActiveDocument.SaveAs mDaNam
        End If
    End If
    DoEvents
    .Visible = True
    DoEvents
    .WindowState = wdWindowStateMaximize
End With

Set OffOb = Nothing

End Sub
Public Sub FWoUEr()
On Error Resume Next
'Öffnet Word mit einem Dokument und übergibt Variablen für die Textmarken, benötigt FilNam

Dim AktBo As Integer

Select Case mOpTyp
Case 1: Set OffOb = CreateObject("Word.Application")
Case 2: Set OffOb = CreateObject("Word.Application")
Case 3: Set OffOb = GetObject(, "Word.Application")
End Select

DoEvents
With OffOb
    Select Case mOpTyp
    Case 1:
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    Case 2:
        .Documents.Add
        .Selection.GOTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=7
        .Selection.InsertAfter mAnsch
    Case 3:
        DoEvents
        If LCase(Right$(mFiNam, 3)) = "dot" Then
            .Documents.Add Template:=mFiNam, NewTemplate:=False
        ElseIf LCase(Right$(mFiNam, 3)) = "doc" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        ElseIf LCase(Right$(mFiNam, 4)) = "docx" Then
            .Documents.Open FileName:=mFiNam, ConfirmConversions:=False, ReadOnly:=mSchut, AddToRecentFiles:=True
        End If
    End Select
    DoEvents
    mAnzTe = .ActiveDocument.Bookmarks.Count
    If mAnzTe > 0 Then
        For AktBo = 1 To mAnzTe
            Select Case AktBo
            Case 1: .ActiveDocument.Bookmarks(AktBo).Select
                    .Selection.InsertAfter mTeFax
            Case 2: .ActiveDocument.Bookmarks(AktBo).Select
                    .Selection.InsertAfter mAnsch
            Case 3: .ActiveDocument.Bookmarks(AktBo).Select
                    .Selection.InsertAfter mBrief
            End Select
        Next AktBo
    End If
    DoEvents
    .ActiveDocument.GOTo What:=wdGoToLine, Which:=wdGoToRelative, Count:=22
    DoEvents
    If mDocSa = True Then
        If mDaNam <> vbNullString Then
            .ActiveDocument.SaveAs mDaNam
        End If
    End If
    DoEvents
    .Visible = True
    DoEvents
    .WindowState = wdWindowStateMaximize
End With

Set OffOb = Nothing

End Sub
Public Function FWoVer() As Integer
On Error Resume Next
'Liefert das Installationsverzeichnis von Word

Dim WoVer As Integer

If regKeyExist("Software\Microsoft\Office\8.0\Word\Options") = False Then
    If regKeyExist("Software\Microsoft\Office\9.0\Word\Options") = False Then
        If regKeyExist("Software\Microsoft\Office\9.0\Word\InstallRoot") = False Then
            If regKeyExist("Software\Microsoft\Office\10.0\Word\Options") = False Then
                If regKeyExist("Software\Microsoft\Office\10.0\Word\InstallRoot") = False Then
                    If regKeyExist("Software\Microsoft\Office\11.0\Word\Options") = False Then
                        If regKeyExist("Software\Microsoft\Office\11.0\Word\InstallRoot") = False Then
                            If regKeyExist("Software\Microsoft\Office\12.0\Word\Options") = False Then
                                If regKeyExist("Software\Microsoft\Office\12.0\Word\InstallRoot") = False Then
                                    If FWoVor("c:\programme\microsoft office\office\", "winword.exe") = vbNullString Then
                                        WoVer = 0
                                    Else
                                        WoVer = 0
                                    End If
                                Else
                                    WoVer = 12
                                End If
                            Else
                                WoVer = 12
                            End If
                        Else
                            WoVer = 11
                        End If
                    Else
                        WoVer = 11
                    End If
                Else
                    WoVer = 10
                End If
            Else
                WoVer = 10
            End If
        Else
            WoVer = 9
        End If
    Else
        WoVer = 9
    End If
Else
    WoVer = 8
End If

FWoVer = WoVer

End Function
Private Function FWoVor(ByVal StPf As String, ByVal SuNa As String) As String
On Error Resume Next

Dim ReWe As String
Dim FiHa As Long
  
ReWe = Space(512)
FiHa = SearchTree(StPf, SuNa, ReWe)

If FiHa <> 0 Then
    ReWe = Left$(ReWe, InStr(ReWe, vbNullChar) - 1)
    FWoVor = ReWe
Else
    FWoVor = vbNullString
End If

End Function

Private Sub FWoZus(ByVal Flag As Integer)
On Error Resume Next

Select Case Flag
Case 1: ShowWindow mHwnd, SW_MAXIMIZE
Case 2: ShowWindow mHwnd, SW_MINIMIZE
Case 3: ShowWindow mHwnd, SW_NORMAL
End Select

End Sub

Public Property Get hwnd() As Long
    hwnd = mHwnd
End Property

Public Property Let hwnd(ByVal Flag As Long)
    mHwnd = Flag
End Property

Public Property Get OPeTyp() As Integer
Attribute OPeTyp.VB_Description = "Legt den Typ fest, wie Word angesprochen wird: 1,2 oder 3"
    OPeTyp = mOpTyp
End Property
Public Property Let OPeTyp(ByVal Flag As Integer)
    mOpTyp = Flag
End Property
Public Property Get OffVer() As Integer
    OffVer = mOfVer
End Property
Public Property Let OffVer(ByVal Flag As Integer)
    mOfVer = Flag
End Property
Public Property Let SerEml(Flag As Boolean)
Attribute SerEml.VB_Description = "Sollen Serienemails ausgegeben werden?"
    mSerEm = Flag
End Property

Public Property Get SerEml() As Boolean
    SerEml = mSerEm
End Property

Public Property Let SerEti(Flag As Boolean)
Attribute SerEti.VB_Description = "Sollen Serienetiketten ausgegeben werden?"
    mSerEt = Flag
End Property

Public Property Get SerEti() As Boolean
    SerEti = mSerEt
End Property

Public Property Get wDocPf() As String
    wDocPf = mDocPf
End Property

Public Property Let wDocPf(ByVal Flag As String)
    mDocPf = Flag
End Property

Public Property Get wVorPf() As String
    wVorPf = mVorPf
End Property

Public Property Let wVorPf(ByVal Flag As String)
    mVorPf = Flag
End Property
Public Property Let FlKraPo(ByVal Flag As String)
    mKraPo = Flag
End Property
Public Property Get FlKraPo() As String
    FlKraPo = mKraPo
End Property

