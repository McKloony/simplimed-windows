VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private FM As Form
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode

Private mPfa As VBA.Collection
Private mNam As VBA.Collection
Private mDir As VBA.Collection

Private GesZa As Long
Private FelZa As Long
Private mSuPf As String
Private mSuSt As String
Private mUnOr As Boolean
Private mSuOr As Boolean
Private mStop As Boolean
Private mAktu As Boolean

Private Const SHGFI_TYPENAME = &H400&
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 259
Private Const Large = &H100
Private Const Small = &H101
Private Const MxIco = 1000

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type IconType
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSIdType
    id(16) As Byte
End Type

Private Type ShellFileInfoType
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Type DIRLISTTYPE
    Ext As String
    File As String
    Type As String
    Attributes As Long
    FileLen As Long
    LastWrite As FILETIME
End Type

Private Type DRVLISTTYPE
    DrNam As String
    TyNam As String
    VoNam As String
    DrTyp As Integer
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type FILELISTYPE
    Name As String
    Path As String
End Type

Private Type COLLISTTYPE
    FilNa As String
    PfaNa As String
    IsDir As Boolean
    Attrb As Long
End Type

Private FiLis() As DIRLISTTYPE
Private DrLis() As DRVLISTTYPE
Private VeLis() As FILELISTYPE
Private CoLis() As COLLISTTYPE

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As ShellFileInfoType, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, riid As CLSIdType, ByVal fown As Long, lpUnk As Object) As Long

Private Sub Class_Initialize()

Set mPfa = New Collection
Set mNam = New Collection
Set mDir = New Collection

End Sub

Private Sub Class_Terminate()

Dim GesZa As Long

For GesZa = 1 To mPfa.Count
    mPfa.Remove 1
Next GesZa
For GesZa = 1 To mNam.Count
    mNam.Remove 1
Next GesZa
For GesZa = 1 To mDir.Count
    mDir.Remove 1
Next GesZa

Set mPfa = Nothing
Set mNam = Nothing
Set mDir = Nothing

ReDim FiLis(0)
ReDim DrLis(0)
ReDim VeLis(0)
ReDim CoLis(0)

End Sub

Public Property Let SuchAkti(Flag As Boolean)
    mAktu = Flag
End Property

Public Property Get SuchAkti() As Boolean
    SuchAkti = mAktu
End Property

Private Function SuchAttr(Attrb As Long) As String
On Error Resume Next

Dim AtStr As String
  
If Attrb And FILE_ATTRIBUTE_TEMPORARY Then AtStr = "T"
If Attrb And FILE_ATTRIBUTE_READONLY Then AtStr = AtStr & "R"
If Attrb And FILE_ATTRIBUTE_HIDDEN Then AtStr = AtStr & "H"
If Attrb And FILE_ATTRIBUTE_SYSTEM Then AtStr = AtStr & "S"
If Attrb And FILE_ATTRIBUTE_NORMAL Then AtStr = AtStr & "N"
If Attrb And FILE_ATTRIBUTE_COMPRESSED Then AtStr = AtStr & "C"
If Attrb And FILE_ATTRIBUTE_ARCHIVE Then AtStr = AtStr & "A"
SuchAttr = AtStr

End Function

Private Function SuchExt(ByVal FiNam As String) As String
On Error Resume Next

Dim AktZa As Long
    
For AktZa = Len(FiNam) To 1 Step -1
  If Mid$(FiNam, AktZa, 1) = "." Then Exit For
Next AktZa
SuchExt = Mid$(FiNam, AktZa + 1)

End Function
Private Function SuchFile() As Boolean
On Error GoTo FiErr

Dim Gefun As Long
Dim AktZa As Long
Dim RetWe As Long
Dim SpStr As String
Dim FiNam As String
Dim ExStr As String
Dim BeStr As String
Dim StrEx As Integer
Dim IsDir As Boolean
Dim WinDat As WIN32_FIND_DATA

Const extra = "|zico=Symbol|zexe=Anwendung|" & "zlnk=Verknüpfung|zcur=Cursor|"

Set FM = frmMain

SpStr = "|Folder=Verzeichnis|"
ReDim FiLis(0)
If Right$(mSuPf, 1) <> "\" And Right$(mSuPf, 1) <> "/" Then mSuPf = mSuPf & "\"

Gefun = FindFirstFile(mSuPf & mSuSt, WinDat)
If Gefun <> INVALID_HANDLE_VALUE Then
    Do
    FiNam = Left(WinDat.cFileName, InStr(WinDat.cFileName, Chr(0)) - 1)
    If (WinDat.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
        If (FiNam <> ".") And (FiNam <> "..") Then
            With FiLis(UBound(FiLis))
                .Attributes = WinDat.dwFileAttributes
                .FileLen = WinDat.nFileSizeLow
                .LastWrite = WinDat.ftLastWriteTime
                .Ext = "Ordner"
                .File = FiNam
                .Type = "Verzeichnis"
            End With
            ReDim Preserve FiLis(0 To UBound(FiLis) + 1)
            If Not IsDir Then IsDir = True
        End If
    Else
        ExStr = "z" & LCase(SuchExt(FiNam))
        AktZa = InStr(1, SpStr, "|" & ExStr & "=")
        RetWe = InStr(1, extra, "|" & ExStr & "=")
        If AktZa = 0 Or RetWe <> 0 Then
            If RetWe <> 0 Then
                RetWe = RetWe + Len(ExStr) + 2
                AktZa = InStr(RetWe, extra, "|")
                BeStr = Mid$(extra, RetWe, AktZa - RetWe)
                ExStr = ExStr + CStr(StrEx)
                StrEx = StrEx + 1
            Else
                BeStr = SuchType(mSuPf & FiNam)
                SpStr = SpStr & ExStr & "=" & BeStr & "|"
            End If
        Else
            AktZa = AktZa + Len(ExStr) + 2
            RetWe = InStr(AktZa, SpStr, "|")
            BeStr = Mid$(SpStr, AktZa, RetWe - AktZa)
        End If
        With FiLis(UBound(FiLis))
            .Attributes = WinDat.dwFileAttributes
            .FileLen = WinDat.nFileSizeLow
            .LastWrite = WinDat.ftLastWriteTime
            .Ext = ExStr
            .File = FiNam
            .Type = BeStr
        End With
        ReDim Preserve FiLis(0 To UBound(FiLis) + 1)
    End If
    Loop While FindNextFile(Gefun, WinDat)
    FindClose Gefun
End If

If UBound(FiLis) > 0 Then
    SuchNam
    SuchFile = True
End If

Exit Function

FiErr:
If GlDbg Then MsgBox Err.Description, 48, "SuchFile " & Err.Number
Resume Next

End Function
Public Function SuchIcon(lSize As Long, SFile As String) As IPictureDisp
On Error GoTo IcErr

Dim RetWe As Long
Dim ReUnb As IUnknown
Dim Icon1 As IconType
Dim CLSID As CLSIdType
Dim ShInf As ShellFileInfoType

SHGetFileInfo SFile, 0, ShInf, Len(ShInf), lSize

Icon1.cbSize = Len(Icon1)
Icon1.picType = vbPicTypeIcon
Icon1.hIcon = ShInf.hIcon
CLSID.id(8) = &HC0
CLSID.id(15) = &H46

RetWe = OleCreatePictureIndirect(Icon1, CLSID, 1, ReUnb)
Set SuchIcon = ReUnb

Exit Function

IcErr:
If GlDbg Then MsgBox Err.Description, 48, "SuchIcon " & Err.Number
Resume Next

End Function
Private Function SuchNam()
On Error GoTo NaErr

Dim AktZa As Integer
Dim EndZa As Long
Dim DirLs As DIRLISTTYPE

For AktZa = 0 To UBound(FiLis) - 1
    If FiLis(AktZa).Type = "Verzeichnis" Then
        DirLs = FiLis(EndZa)
        FiLis(EndZa) = FiLis(AktZa)
        FiLis(AktZa) = DirLs
        EndZa = EndZa + 1
    End If
Next AktZa

If EndZa <> 0 Then SuchSort 0, EndZa - 1
If EndZa < UBound(FiLis) Then
    SuchSort EndZa, UBound(FiLis) - 1
End If

Exit Function

NaErr:
If GlDbg Then MsgBox Err.Description, 48, "SuchName " & Err.Number
Resume Next

End Function
Private Sub SuchSort(StaZah As Long, EndZah As Long)
On Error GoTo SoErr

Dim StaVe As Long
Dim EndVe As Long
Dim RefVe As String
Dim FilTy As DIRLISTTYPE

StaVe = StaZah
EndVe = EndZah
RefVe = FiLis((StaVe + EndVe) / 2).File

Do
Do While (FiLis(StaVe).File < RefVe)
StaVe = StaVe + 1
Loop
Do While (FiLis(EndVe).File > RefVe)
EndVe = EndVe - 1
Loop
If StaVe <= EndVe Then
    FilTy = FiLis(StaVe)
    FiLis(StaVe) = FiLis(EndVe)
    FiLis(EndVe) = FilTy
    StaVe = StaVe + 1
    EndVe = EndVe - 1
End If
Loop Until (StaVe > EndVe)

If StaZah < EndVe Then SuchSort StaZah, EndVe
If StaVe < EndZah Then SuchSort StaVe, EndZah

Exit Sub

SoErr:
If GlDbg Then MsgBox Err.Description, 48, "SuchSort " & Err.Number
Resume Next

End Sub
Private Function SuchTime(FTime As FILETIME) As String
On Error GoTo TiErr

Dim Stund As String
Dim Minut As String
Dim Sekun As String
Dim WoTag As String
Dim Monat As String
Dim Jahre As String
Dim STime As SYSTEMTIME

FileTimeToSystemTime FTime, STime
With STime
    WoTag = .wDay
    If Len(WoTag) < 2 Then WoTag = "0" & WoTag
    Monat = .wMonth
    If Len(Monat) < 2 Then Monat = "0" & Monat
    Jahre = CStr(.wYear)
    Minut = Trim$(CStr(.wMinute))
    If Len(Minut) < 2 Then Minut = "0" & Minut
    Sekun = Trim$(CStr(.wSecond))
    If Len(Sekun) < 2 Then Sekun = "0" & Sekun
    Stund = Trim$(CStr(.wHour))
    If Len(.wHour) < 2 Then Stund = "0" & Stund
    SuchTime = WoTag & "." & Monat & "." & Jahre & " " & Stund & ":" & Minut & ":" & Sekun
End With

Exit Function

TiErr:
If GlDbg Then MsgBox Err.Description, 48, "SuchTime " & Err.Number
Resume Next

End Function
Public Sub SuchTree()
On Error GoTo AnErr
'Listet alle Filter auf

Dim AktZa As Integer
Dim TrKey As String
Dim FiNam As String

Set FM = frmMain
Set TrLi1 = FM.trvList1

GesZa = 0
FelZa = 0

If SuchFile = True Then
    For AktZa = 0 To UBound(FiLis) - 1
        TrKey = "F" & Format$(AktZa + 1, "0000")
        FiNam = Mid$(FiLis(AktZa).File, 1, Len(FiLis(AktZa).File) - 4)
        Set Knote = TrLi1.Nodes.Add("P04", 4, TrKey, FiNam, IC16_Doc_View)
    Next AktZa
End If

With TrLi1
    .Nodes(5).Expanded = True
    .Nodes(1).EnsureVisible
End With

Exit Sub

AnErr:
If GlDbg Then MsgBox Err.Description, 48, "SuchTree " & Err.Number
Resume Next

End Sub
Private Function SuchType(FiNam As String) As String
On Error Resume Next

Dim RetWe As Long
Dim ShInf As ShellFileInfoType
    
SHGetFileInfo FiNam, 0, ShInf, Len(ShInf), SHGFI_TYPENAME
RetWe = InStr(1, ShInf.szTypeName, Chr$(0)) - 1
SuchType = Left(ShInf.szTypeName, RetWe)

End Function
Public Property Let SuOrd(Flag As Boolean)
    mSuOr = Flag
End Property
Public Property Get SuOrd() As Boolean
    SuOrd = mSuOr
End Property
Public Property Let SuPfa(Flag As String)
    mSuPf = Flag
End Property
Public Property Get SuPfa() As String
    SuPfa = mSuPf
End Property
Public Property Let SuStr(Flag As String)
    mSuSt = Flag
End Property
Public Property Get SuStr() As String
    SuStr = mSuSt
End Property
Public Property Let SuUnt(Flag As Boolean)
    mUnOr = Flag
End Property
Public Property Get SuUnt() As Boolean
    SuUnt = mUnOr
End Property
