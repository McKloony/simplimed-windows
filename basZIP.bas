Attribute VB_Name = "basZIP"
Option Explicit

Private FM As Form
Private TxSta As VB.TextBox
Private Lab02 As XtremeSuiteControls.Label
Private PrBr1 As XtremeSuiteControls.ProgressBar

Public Declare Function sevZIP_Init Lib "sevZip40.dll" Alias "InitZip" (ByVal SInit As String) As Boolean
Public Declare Sub sevZIP_SetLanguage Lib "sevZip40.dll" Alias "SetLanguage" (ByVal nLanguage As Long)
Public Declare Sub sevZIP_SetCompressionRate Lib "sevZip40.dll" Alias "SetCompressionRate" (ByVal nRate As Long)
Public Declare Sub sevZIP_SetRootDir Lib "sevZip40.dll" Alias "SaveFolderLocation" (ByVal sRootPath As String)
Public Declare Function sevZIP_SetTempPath Lib "sevZip40.dll" Alias "SetTempPath" (ByVal sPath As String) As Boolean
Public Declare Sub sevZIP_IncludeOnlyArchivFiles Lib "sevZip40.dll" Alias "IncludeOnlyArchivFiles" (ByVal ArchivFiles As Boolean)
Public Declare Sub sevZIP_IncludeHiddenFiles Lib "sevZip40.dll" Alias "IncludeHiddenFiles" (ByVal HiddenFiles As Boolean)
Public Declare Sub sevZIP_IncludeSystemFiles Lib "sevZip40.dll" Alias "IncludeSystemFiles" (ByVal SystemFiles As Boolean)
Public Declare Sub sevZip_ResetArchivBit Lib "sevZip40.dll" Alias "ResetArchiveBit" (ByVal ResetArchivBitOnZip As Boolean)
Public Declare Function sevZIP_ZipFile Lib "sevZip40.dll" Alias "ZipFile" (ByVal sZipFile As String, ByVal sSourceFile As String, ByVal sPassword As String, ByVal nOverwrite As Long, ByVal hStatus As Long) As Long
Public Declare Function sevZIP_ZipAddFile Lib "sevZip40.dll" Alias "ZipAddFile" (ByVal sZipFile As String, ByVal sFilesToAdd As String, ByVal nOverwrite As Long, ByVal hStatus As Long) As Long
Public Declare Function sevZIP_ZipAddFileEx Lib "sevZip40.dll" Alias "ZipAddFileEx" (ByVal sZipFile As String, ByVal sFilesToAdd As String, ByVal sPassword As String, ByVal nOverwrite As Long, ByVal hStatus As Long) As Long
Public Declare Function sevZIP_ZipDeleteFile Lib "sevZip40.dll" Alias "ZipDeleteFile" (ByVal sZipFile As String, ByVal sFilesToDelete As String, ByVal hStatus As Long) As Long
Public Declare Function sevZIP_ZipFolderEx Lib "sevZip40.dll" Alias "ZipFolderEx" (ByVal sZipFile As String, ByVal sSourcePath As String, ByVal sFileSpec As String, ByVal nSubFolder As Long, ByVal sPassword As String, ByVal nOverwrite As Long, ByVal hStatus As Long) As Long
Public Declare Function sevZIP_ZipFileCount Lib "sevZip40.dll" Alias "ZipFileCount" (ByVal sZipFile As String) As Long
Public Declare Function sevZIP_ZipFileInfo Lib "sevZip40.dll" Alias "ZipFileInfo" (ByVal sZipFile As String, ByVal nIndex As Long, ByVal hwnd As Long) As Long
Public Declare Function sevZIP_ZipFileInfoEx Lib "sevZip40.dll" Alias "ZipFileInfoEx" (ByVal sZipFile As String, ByVal nIndex As Long, ByRef sBuffer As String) As Long

' Mögliche Rückgabewerte
' 0: Zip-File ist OK
' 1: Zip-File ist beschädigt oder enthält keine Dateien
' 2: Zip-File existiert nicht
' 3: Ungültiges Zip - File

Public Declare Function sevZIP_CheckZipFile Lib "sevZip40.dll" Alias "CheckZipFile" (ByVal sZipFile As String, ByVal sPassword As String) As Long
Public Declare Function sevZIP_UnzipEx Lib "sevZip40.dll" Alias "UnZipEx" (ByVal sZipFile As String, ByVal sDestPath As String, ByVal sFileSpec As String, ByVal nSubFolder As Long, ByVal sPassword As String, ByVal nOverwriteState As Long, ByVal hStatus As Long) As Long
Public Declare Sub sevZIP_CancelZip Lib "sevZip40.dll" Alias "CancelZip" ()
Public Declare Function sevZIP_ZipProgress Lib "sevZip40.dll" Alias "ZipProgress" () As Long

' nur aus Kompatibilitätsgründen zur Version 1.0
Public Declare Function sevZIP_ZipFolder Lib "sevZip40.dll" Alias "ZipFolder" (ByVal sZipFile As String, ByVal sSourcePath As String, ByVal sPassword As String, ByVal nOverwrite As Long, ByVal hStatus As Long) As Long
Public Declare Function sevZIP_UnZip Lib "sevZip40.dll" Alias "UnZip" (ByVal sZipFile As String, ByVal sDestPath As String, ByVal sPassword As String, ByVal nOverwriteState As Long, ByVal hStatus As Long) As Long

Public Declare Sub sevZip_SetEncryption Lib "sevZip40.dll" Alias "SetEncryption" (ByVal EncryptionMode As EncryptMode)
Public Declare Function sevZip_IsPasswordProtected Lib "sevZip40.dll" Alias "IsPasswordProtected" (ByVal sZipFile As String) As Boolean
Public Declare Function sevZip_GetEncryption Lib "sevZip40.dll" Alias "GetEncryption" (ByVal sZipFile As String) As Long

Public Enum EncryptMode
    emStandard = 0
    emAES128 = 1
    emAES192 = 2
    emAES256 = 3
End Enum

Public Declare Function sevZip_IsZipFile Lib "sevZip40.dll" Alias "IsZipFile" (ByVal sZipFile As String) As Boolean
Public Declare Function sevZIP_CheckZipFileEx Lib "sevZip40.dll" Alias "CheckZipFileEx" (ByVal sZipFile As String, ByVal sPassword As String, ByVal hStatus As Long) As Long
Public Declare Function sevZip_GetComment Lib "sevZip40.dll" Alias "GetComment" (ByVal sZipFile As String, ByRef sComment As String) As Long
Public Declare Function sevZip_SetComment Lib "sevZip40.dll" Alias "SetComment" (ByVal sZipFile As String, ByVal sComment As String) As Long
Public Declare Function sevZip_OpenZip Lib "sevZip40.dll" Alias "OpenZip" (ByVal sZipFile As String) As Long
Public Declare Function sevZip_ReadZip Lib "sevZip40.dll" Alias "ReadZip" (ByVal nIndex As Long, ByRef sBuffer As String) As Long
Public Declare Sub sevZip_CloseZip Lib "sevZip40.dll" Alias "CloseZip" ()
Public Declare Sub sevZip_UnZipCurFile Lib "sevZip40.dll" Alias "UnZipGetCurrentFile" (ByRef sFile As String)

Private clFil As clsFile
Private clNet As clsNetz
Public Sub mldDaZIP(ByVal PfaNa As String, ByVal ZipNa As String, ByVal AnzDa As Integer, Optional ByVal KmStr As String)
On Error GoTo DaErr
'Komprimieren

Dim ZipOv As Long
Dim RetWe As Long
Dim AktZa As Long
Dim Posit As Long

Set FM = frmStatus
Set TxSta = FM.txtDummy

If ZipNa <> vbNullString Then
    Posit = InStr(1, ZipNa, "\\", 1)
    If Posit > 0 Then
        ZipNa = Replace(ZipNa, "\\", "\", 1, , 1)
    End If
End If

If AnzDa > 0 Then
    ZipOv = 2
    sevZIP_SetLanguage 1 'Deutsch
    sevZIP_SetCompressionRate 6
    sevZIP_IncludeSystemFiles False
    sevZIP_IncludeHiddenFiles True
    sevZIP_IncludeOnlyArchivFiles False
    sevZip_ResetArchivBit False
    sevZIP_SetTempPath GlTmp
    sevZIP_SetRootDir PfaNa
    
    If GlZip(0) <> vbNullString Then
        RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(0), ZipOv, TxSta.hwnd)
    ElseIf GlZip(1) <> vbNullString Then
        RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(1), ZipOv, TxSta.hwnd)
    End If

    If RetWe >= 0 Then
        If AnzDa > 1 Then
            For AktZa = 2 To AnzDa
                RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(AktZa), ZipOv, TxSta.hwnd)
            Next AktZa
        End If
    End If
    
    If KmStr <> vbNullString Then
        sevZip_SetComment ZipNa, KmStr
    End If
End If

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "mldDaZIP " & Err.Number
Resume Next

End Sub
Public Sub mldOrZIP(ByVal PfaNa As String, ByVal ZipNa As String, Optional ByVal KmStr As String, Optional ByVal VerDa As Boolean = False, Optional ByVal VerPa As String)
On Error GoTo DaErr
'Ordner Komprimieren

Dim ZipOv As Long
Dim RetWe As Long
Dim AktZa As Long
Dim Posit As Long
Dim DaPas As String

Set FM = frmStatus
Set TxSta = FM.txtDummy

If VerDa = True Then
    If VerPa <> vbNullString Then
        DaPas = VerPa
    Else
        DaPas = InputBox("Bitte optional ein Verschlüsselungskennwort eingeben:", "Dateiverschlüsselung", vbNullString)
    End If
Else
    DaPas = vbNullString
End If

If ZipNa <> vbNullString Then
    Posit = InStr(1, ZipNa, "\\", 1)
    If Posit > 0 Then
        ZipNa = Replace(ZipNa, "\\", "\", 1, , 1)
    End If
    
    ZipOv = 2
    sevZIP_SetLanguage 1  ' Deutsch
    sevZIP_SetCompressionRate 6
    sevZIP_IncludeSystemFiles False
    sevZIP_IncludeHiddenFiles True
    sevZIP_IncludeOnlyArchivFiles False
    sevZip_ResetArchivBit False
    sevZIP_SetTempPath GlTmp
    
    sevZip_SetEncryption emAES128 'Verschlüsselung
    sevZIP_SetRootDir PfaNa
    
    RetWe = sevZIP_ZipFolderEx(ZipNa, PfaNa, "*.*", 1, DaPas, ZipOv, TxSta.hwnd)
            
    If KmStr <> vbNullString Then
        sevZip_SetComment ZipNa, KmStr
    End If
End If

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "mldOrZIP " & Err.Number
Resume Next

End Sub
Public Sub mldVeZIP(ByVal PfaNa As String, ByVal ZipNa As String, ByVal AnzDa As Integer, Optional ByVal VerDa As Boolean = False, Optional ByVal VerPa As String, Optional ByVal KmStr As String)
On Error GoTo DaErr
'Komplrimieren und Verschlüsseln

Dim ZipOv As Long
Dim RetWe As Long
Dim AktZa As Long
Dim DaNam As String
Dim DaPas As String
Dim Posit As Integer
Dim Lange As Integer

Set FM = frmStatus
Set TxSta = FM.txtDummy

If VerDa = True Then
    If VerPa <> vbNullString Then
        DaPas = VerPa
    Else
        DaPas = InputBox("Bitte optional ein Verschlüsselungskennwort eingeben:", "Dateiverschlüsselung", vbNullString)
    End If
Else
    DaPas = vbNullString
End If

If AnzDa > 0 Then
    ZipOv = 2
    sevZIP_SetLanguage 1  ' Deutsch
    sevZIP_SetCompressionRate 6
    sevZIP_IncludeSystemFiles False
    sevZIP_IncludeHiddenFiles False
    sevZIP_IncludeOnlyArchivFiles False
    sevZip_ResetArchivBit False
    sevZIP_SetTempPath GlTmp
    sevZIP_SetRootDir PfaNa
    sevZip_SetEncryption emAES128 'Verschlüsselung

    If VerDa = True Then 'Dateiverschlüsselung
        If DaPas <> vbNullString Then
            If GlZip(0) <> vbNullString Then
                RetWe = sevZIP_ZipAddFileEx(ZipNa, GlZip(0), DaPas, ZipOv, TxSta.hwnd)
            ElseIf GlZip(1) <> vbNullString Then
                RetWe = sevZIP_ZipAddFileEx(ZipNa, GlZip(1), DaPas, ZipOv, TxSta.hwnd)
            End If
        Else
            If GlZip(0) <> vbNullString Then
                RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(0), ZipOv, TxSta.hwnd)
            ElseIf GlZip(1) <> vbNullString Then
                RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(1), ZipOv, TxSta.hwnd)
            End If
        End If
    Else
        If GlZip(0) <> vbNullString Then
            RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(0), ZipOv, TxSta.hwnd)
        ElseIf GlZip(1) <> vbNullString Then
            RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(1), ZipOv, TxSta.hwnd)
        End If
    End If
    DoEvents

    If VerDa = True Then 'Datei Verschlüsseln
        If DaPas <> vbNullString Then
            If RetWe >= 0 Then
                If AnzDa > 1 Then
                    frmZIPStat.Show
                    DoEvents
                    Set PrBr1 = frmZIPStat.prbStat1
                    Set Lab02 = frmZIPStat.lblLab02
                    PrBr1.Min = 0
                    PrBr1.Max = AnzDa
                    For AktZa = 2 To AnzDa
                        Lange = Len(GlZip(AktZa))
                        Posit = InStrRev(GlZip(AktZa), "\", 1)
                        DaNam = Mid$(GlZip(AktZa), Posit + 1, Lange - Posit)
                        RetWe = sevZIP_ZipAddFileEx(ZipNa, GlZip(AktZa), DaPas, ZipOv, TxSta.hwnd)
                        PrBr1.Value = AktZa
                        Lab02.Caption = DaNam
                        DoEvents
                    Next AktZa
                    Unload frmZIPStat
                    Set frmZIPStat = Nothing
                End If
            End If
        Else
            If RetWe >= 0 Then
                If AnzDa > 1 Then
                    frmZIPStat.Show
                    DoEvents
                    Set PrBr1 = frmZIPStat.prbStat1
                    Set Lab02 = frmZIPStat.lblLab02
                    PrBr1.Min = 0
                    PrBr1.Max = AnzDa
                    For AktZa = 2 To AnzDa
                        Lange = Len(GlZip(AktZa))
                        Posit = InStrRev(GlZip(AktZa), "\", 1)
                        DaNam = Mid$(GlZip(AktZa), Posit + 1, Lange - Posit)
                        RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(AktZa), ZipOv, TxSta.hwnd)
                        PrBr1.Value = AktZa
                        Lab02.Caption = DaNam
                        DoEvents
                    Next AktZa
                    Unload frmZIPStat
                    Set frmZIPStat = Nothing
                End If
            End If
        End If
    Else
        If RetWe >= 0 Then
            If AnzDa > 1 Then
                frmZIPStat.Show
                DoEvents
                Set PrBr1 = frmZIPStat.prbStat1
                Set Lab02 = frmZIPStat.lblLab02
                PrBr1.Min = 0
                PrBr1.Max = AnzDa
                For AktZa = 2 To AnzDa
                    Lange = Len(GlZip(AktZa))
                    Posit = InStrRev(GlZip(AktZa), "\", 1)
                    DaNam = Mid$(GlZip(AktZa), Posit + 1, Lange - Posit)
                    RetWe = sevZIP_ZipAddFile(ZipNa, GlZip(AktZa), ZipOv, TxSta.hwnd)
                    PrBr1.Value = AktZa
                    Lab02.Caption = DaNam
                    DoEvents
                Next AktZa
                Unload frmZIPStat
                Set frmZIPStat = Nothing
            End If
        End If
    End If

    If KmStr <> vbNullString Then
        sevZip_SetComment ZipNa, KmStr
    End If
End If

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "mldVeZIP " & Err.Number
Resume Next

End Sub
Public Sub mldUnZIP(ByVal PfaNa As String, ByVal ZipNa As String)
On Error GoTo DaErr

Dim ZipOv As Long
Dim ZipOr As Long
Dim RetWe As Long

Set FM = frmStatus
Set TxSta = FM.txtDummy

If ZipNa <> vbNullString Then
    ZipOv = 2
    ZipOr = 0
    sevZIP_SetLanguage 1  ' Deutsch
    RetWe = sevZIP_UnzipEx(ZipNa, PfaNa, "*.*", ZipOr, vbNullString, ZipOv, TxSta.hwnd)
End If

Exit Sub

DaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "mldUnZIP " & Err.Number
Resume Next

End Sub
