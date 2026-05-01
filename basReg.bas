Attribute VB_Name = "basReg"
Option Explicit

Private RetWe As Long
Private Buffr As String

Private m_lngRetVal As Long
Private Const REG_NONE As Long = 0
Private Const REG_SZ As Long = 1
Private Const REG_EXPAND_SZ As Long = 2
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4
Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Private Const REG_DWORD_BIG_ENDIAN As Long = 5
Private Const REG_LINK As Long = 6
Private Const REG_MULTI_SZ As Long = 7
Private Const REG_RESOURCE_LIST As Long = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_NO_MORE_ITEMS As Long = 259
Private Const REG_OPTION_NON_VOLATILE As Long = 0
Private Const REG_OPTION_VOLATILE As Long = &H1

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_USERS As Long = &H80000003
Private Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Private Const HKEY_CURRENT_CONFIG As Long = &H80000005
Private Const HKEY_DYN_DATA As Long = &H80000006

Private Const HELP_CONTENTS = &H3
Private Const HELP_SETCONTENTS = &H5
Private Const HELP_CONTEXTPOPUP = &H8
Private Const HELP_FORCEFILE = &H9
Private Const HELP_COMMAND = &H102
Private Const HELP_PARTIALKEY = &H105
Private Const HELP_SETWINPOS = &H203

Private Const HH_DISPLAY_TOPIC = &H0    ' WinHelp equivalent.
Private Const HH_DISPLAY_TOC = &H1      ' WinHelp equivalent.
Private Const HH_DISPLAY_INDEX = &H2    ' WinHelp equivalent.
Private Const HH_DISPLAY_SEARCH = &H3   ' WinHelp equivalent.

Private Const HH_HELP_CONTEXT = &HF     ' Display mapped numeric.
Private Const HH_CLOSE_ALL = &H12       ' WinHelp equivalent.

Private Type tGUID
    bytes(15) As Byte
End Type

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function StringFromGUID2 Lib "OLE32.dll" (guid As tGUID, ByVal lpszString As String, ByVal lMax As Long) As Long
Private Declare Function CoCreateGuid Lib "OLE32.dll" (guid As tGUID) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal szFilename As String, ByVal dwCommand As Long, ByRef dwData As Any) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function CreateID(Optional ByVal StaCh As String = "G") As String
On Error Resume Next

Dim GuiID As tGUID
Dim RetSt As String
Dim TmStr As String
Dim TmGui As String

RetSt = Space(100)
CoCreateGuid GuiID
RetWe = StringFromGUID2(GuiID, RetSt, Len(RetSt))
TmStr = Left$(StrConv(RetSt, vbFromUnicode), RetWe - 1)

TmGui = Mid$(TmStr, 2, Len(TmStr) - 2)
TmGui = Replace(TmGui, "-", vbNullString, 1)
CreateID = StaCh & TmGui

End Function

Public Sub IniDelKey(ByVal FiNam As String, ByVal InSek As String, ByVal InKey As String)
    RetWe = WritePrivateProfileString(InSek, InKey, 0&, FiNam)
End Sub

Public Sub IniDelSek(ByVal FiNam As String, ByVal InSek As String)
    RetWe = WritePrivateProfileString(InSek, 0&, 0&, FiNam)
End Sub
Public Sub IniGetAry(ByVal FiNam As String, ByVal InSek As String, IniAr() As String)
On Error Resume Next

Dim StaPo As Integer
Dim Posit As Integer
Dim AktZa As Integer

Buffr = Space(32767)
RetWe = GetPrivateProfileSection(InSek, Buffr, Len(Buffr), FiNam)

Buffr = Left$(Buffr, RetWe)

If Buffr <> vbNullString Then
    StaPo = 1
    ReDim IniAr(0)
    Do While StaPo < RetWe
        Posit = InStr(StaPo, Buffr, Chr$(0))
        If Posit = 0 Then Exit Do
        
        IniAr(AktZa) = Mid$(Buffr, StaPo, Posit - StaPo)
        AktZa = AktZa + 1
        ReDim Preserve IniAr(0 To AktZa)
        StaPo = Posit + 1
    Loop
End If
  
End Sub
Public Function IniGetOpt(ByVal InSek As String, ByVal InKey As String) As String
On Error GoTo WiErr

Buffr = Space$(1024)
RetWe = GetPrivateProfileString(InSek, InKey, vbNullString, Buffr, Len(Buffr), GlOpt)
IniGetOpt = Left$(Buffr, RetWe)

Exit Function

WiErr:
If GlDbg = True Then SErLog Err.Description & " IniGetOpt " & Err.Number
Resume Next

End Function

Public Function IniGetBig(ByVal FiNam As String, ByVal InSek As String, ByVal InKey As String) As String
On Error Resume Next

Buffr = Space$(8192)
RetWe = GetPrivateProfileString(InSek, InKey, vbNullString, Buffr, Len(Buffr), FiNam)
IniGetBig = Left$(Buffr, RetWe)

End Function
Public Function IniGetFil(ByVal FiNam As String, InSek As String, ByVal InKey As String) As String
On Error GoTo WiErr

Buffr = Space$(256)
RetWe = GetPrivateProfileString(InSek, InKey, vbNullString, Buffr, Len(Buffr), FiNam)
IniGetFil = Left$(Buffr, RetWe)

Exit Function

WiErr:
If GlDbg = True Then SErLog Err.Description & " IniGetFil " & Err.Number
Resume Next

End Function
Public Function IniGetSek(ByVal FiNam As String, ByVal SuSek As String) As Boolean
On Error Resume Next

Dim AktBu As Long
Dim AktZa As Long

AktBu = 256
Buffr = Space$(AktBu)

Do
AktBu = AktBu * 2
Buffr = String(AktBu, 0)
RetWe = GetPrivateProfileSectionNames(Buffr, AktBu, FiNam)
Loop While (RetWe = AktBu - 2)

Do While AktZa < RetWe
    AktBu = AktZa + 1
    AktZa = InStr(AktBu, Buffr, Chr(0))
    If UCase(Mid(Buffr, AktBu, AktZa - AktBu)) = UCase(SuSek) Then
        IniGetSek = True
        Exit Do
    End If
Loop

End Function
Public Function IniGetTSE(ByVal InSek As String, ByVal InKey As String, ByVal DaIni As String) As String
On Error GoTo WiErr

Buffr = Space$(256)
RetWe = GetPrivateProfileString(InSek, InKey, vbNullString, Buffr, Len(Buffr), DaIni)
IniGetTSE = Left$(Buffr, RetWe)

Exit Function

WiErr:
If GlDbg = True Then SErLog Err.Description & " IniGetVal " & Err.Number
Resume Next

End Function

Public Function IniGetVal(ByVal InSek As String, ByVal InKey As String) As String
On Error GoTo WiErr

Buffr = Space$(256)
RetWe = GetPrivateProfileString(InSek, InKey, vbNullString, Buffr, Len(Buffr), GlINI)
IniGetVal = Left$(Buffr, RetWe)

Exit Function

WiErr:
If GlDbg = True Then SErLog Err.Description & " IniGetVal " & Err.Number
Resume Next

End Function
Public Sub IniSetAry(ByVal FiNam As String, ByVal InSek As String, IniAr() As String)
On Error Resume Next

Dim AktZa As Integer

For AktZa = LBound(IniAr) To UBound(IniAr)
    Buffr = Buffr & IniAr(AktZa) & Chr$(0)
Next AktZa

Buffr = Left$(Buffr, Len(Buffr) - 1)
RetWe = WritePrivateProfileSection(InSek, Buffr, FiNam)
    
End Sub
Public Sub IniSetFil(FiNam As String, ByVal InSek As String, ByVal InKey As String, ByVal Value As String)
On Error GoTo WiErr

RetWe = WritePrivateProfileString(InSek, InKey, Value, FiNam)

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " IniSetFil " & Err.Number
Resume Next

End Sub

Public Sub IniSetSek(ByVal InSek As String)
    RetWe = WritePrivateProfileSection(InSek, vbNullChar, GlINI)
End Sub
Public Sub IniSetVal(ByVal InSek As String, ByVal InKey As String, ByVal Value As String)
On Error GoTo WiErr

RetWe = WritePrivateProfileString(InSek, InKey, Value, GlINI)

Exit Sub

WiErr:
If GlDbg = True Then SErLog Err.Description & " IniSetVal " & Err.Number
Resume Next

End Sub
Public Function IniSetWer(ByVal InFil As String, ByVal InSek As String, ByVal InKey As String, ByVal InVal As String) As Long
On Error GoTo mError

Dim AktZa As Long
Dim Posit As Long
Dim mLine As String
Dim TmStr As String
Dim TmpSt As String
Dim SekGe As Boolean
Dim FiNam As Integer
Dim ArySe() As String
Dim AryWe() As String

FiNam = FreeFile()
TmStr = vbNullString

If Dir(InFil) <> vbNullString Then
    Open InFil For Input As #FiNam
    While Not EOF(FiNam)
        Line Input #FiNam, mLine
        If TmStr = vbNullString Then
            TmStr = mLine
        Else
            TmStr = TmStr & vbCrLf & mLine
        End If
    Wend
    Close #FiNam
    
    ArySe = Split(TmStr, vbCrLf)
    Posit = UBound(ArySe) + 1
    
    For AktZa = LBound(ArySe) To UBound(ArySe)
        If InStr(1, ArySe(AktZa), "[") > 0 Then
         ' Section gefunden
            If ArySe(AktZa) Like "[[]" & InSek & "[]]" Then
            'If LCase(ArySe(AktZa)) = "[[]" & InSek & "[]]" Then
                SekGe = True
            Else
                If SekGe Then
                    If Posit = (UBound(ArySe) + 1) Then
                    ' Eintrag nicht gefunden -> neuen Eintrag erstellen
                        Posit = AktZa
                    End If
                End If
                SekGe = False
            End If
        Else
            If SekGe Then
                If ArySe(AktZa) Like InKey & "*" Then
                'If InStr(1, ArySe(AktZa), InKey, vbTextCompare) Then
                ' Treffer, Eintrag gefunden -> aktualisieren
                    Posit = -1
                    AryWe = Split(ArySe(AktZa), "=")
                    If LBound(AryWe) = 0 And UBound(AryWe) = 1 Then
                        If InStr(1, AryWe(1), ";") > 0 Then
                            TmpSt = Mid(AryWe(1), InStr(1, AryWe(1), ";"))
                        End If
                        AryWe(1) = InVal & TmpSt
                        ArySe(AktZa) = AryWe(0) & "=" & AryWe(1)
                    End If
                End If
            End If
        End If
    Next AktZa
    
    TmStr = vbNullString
    
    For AktZa = LBound(ArySe) To UBound(ArySe)
        If AktZa = Posit Then
            If TmStr = vbNullString Then
                TmStr = InKey & "=" & InVal
            Else
                TmStr = TmStr & vbCrLf & InKey & "=" & InVal
            End If
            Posit = -1
        End If
        If TmStr = vbNullString Then
            TmStr = ArySe(AktZa)
        Else
            TmStr = TmStr & vbCrLf & ArySe(AktZa)
        End If
    Next AktZa
    
    If AktZa = Posit Then
        ' Section gibt es nocht nicht -> erstellen
        If TmStr = vbNullString Then
            TmStr = "[" & InSek & "]" & vbCrLf & InKey & "=" & InVal
        Else
            If SekGe Then
            ' Section gibt es schon, nur Eintrag hinzufügen!
                TmStr = TmStr & vbCrLf & InKey & "=" & InVal
            Else
                ' Section gibt es noch nicht -> erstellen
                TmStr = TmStr & vbCrLf & "[" & InSek & "]" & vbCrLf & InKey & "=" & InVal
            End If
        End If
    End If
Else
    ' Section gibt es nocht nicht -> erstellen
   TmStr = "[" & InSek & "]" & vbCrLf
   TmStr = TmStr & InKey & "=" & InVal
End If

FiNam = FreeFile()

Open InFil For Output As #FiNam

Print #FiNam, TmStr

Close #FiNam

IniSetWer = 0

Exit Function

mError:
Reset
IniSetWer = -Err.Number

End Function
Public Sub ReDimEx(ByRef MyArray As Variant, ByVal iDimX As Integer, ByVal iDimY As Integer)
 
Dim MyTempArray As Variant
Dim i As Integer
Dim J As Integer

MyTempArray = MyArray

ReDim MyArray(iDimX, iDimY)

For i = LBound(MyTempArray, 1) To UBound(MyTempArray, 1)
    For J = LBound(MyTempArray, 2) To UBound(MyTempArray, 2)
        If i <= iDimX And J <= iDimY Then
            MyArray(i, J) = MyTempArray(i, J)
        End If
    Next J
Next i

End Sub
Public Function regCreKey(ByVal RegKeyStr As String)
On Error Resume Next

Dim lngKeyHandle As Long
  
m_lngRetVal = RegCreateKey(HKEY_CURRENT_USER, RegKeyStr, lngKeyHandle)
m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

Public Sub regCreKeyStr(ByVal RegKeyStr As String, ByVal strRegSubKey As String, varRegData As String)
    
Dim lngKeyHandle As Long
Dim lngDataType As Long
Dim lngKeyValue As Long
Dim strKeyValue As String

lngDataType = REG_SZ

m_lngRetVal = RegCreateKey(HKEY_CURRENT_USER, RegKeyStr, lngKeyHandle)

Select Case lngDataType
Case REG_SZ:
    strKeyValue = Trim(varRegData) & Chr(0)
    m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal strKeyValue, Len(strKeyValue))
Case REG_DWORD:
    lngKeyValue = CLng(varRegData)
    m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngKeyValue, 4&)
End Select

m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Sub

Public Sub regCreKeyVal(ByVal RegKeyStr As String, ByVal strRegSubKey As String, varRegData As Variant)
On Error Resume Next

Dim lngKeyHandle As Long
Dim lngDataType As Long
Dim lngKeyValue As Long
Dim strKeyValue As String

If IsNumeric(varRegData) Then
    lngDataType = REG_DWORD
Else
    lngDataType = REG_SZ
End If

m_lngRetVal = RegCreateKey(HKEY_CURRENT_USER, RegKeyStr, lngKeyHandle)

Select Case lngDataType
Case REG_SZ:
    strKeyValue = Trim(varRegData) & Chr(0)
    m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal strKeyValue, Len(strKeyValue))
Case REG_DWORD:
    lngKeyValue = CLng(varRegData)
    m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngKeyValue, 4&)
End Select

m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Sub

Public Function regDelKey(ByVal RegKeyStr As String, ByVal strRegKeyName As String) As Boolean
On Error Resume Next

Dim lngKeyHandle As Long

regDelKey = False

If regKeyExist(RegKeyStr) Then
    m_lngRetVal = RegOpenKey(HKEY_CURRENT_USER, RegKeyStr, lngKeyHandle)
    m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)
    
    If m_lngRetVal = 0 Then
        regDelKey = True
    End If
    
    m_lngRetVal = RegCloseKey(lngKeyHandle)
End If
  
End Function

Public Function regDelSubKey(ByVal RegKeyStr As String, ByVal strRegSubKey As String)
On Error Resume Next

Dim lngKeyHandle As Long

If regKeyExist(RegKeyStr) Then
    m_lngRetVal = RegOpenKey(HKEY_CURRENT_USER, RegKeyStr, lngKeyHandle)
    m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
    m_lngRetVal = RegCloseKey(lngKeyHandle)
End If

End Function

Public Function regKeyExist(ByVal RegKeyStr As String) As Boolean
On Error Resume Next

Dim lngKeyHandle As Long

lngKeyHandle = 0

m_lngRetVal = RegOpenKey(HKEY_CURRENT_USER, RegKeyStr, lngKeyHandle)

If lngKeyHandle = 0 Then
    regKeyExist = False
Else
    regKeyExist = True
End If

m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

Public Function regReadKey(ByVal RegKeyStr As String, ByVal strRegSubKey As String) As Variant
On Error Resume Next

Dim intPosition As Integer
Dim lngKeyHandle As Long
Dim lngDataType As Long
Dim lngBufferSize As Long
Dim lngBuffer As Long
Dim strBuffer As String

lngKeyHandle = 0
lngBufferSize = 0

m_lngRetVal = RegOpenKey(HKEY_CURRENT_USER, RegKeyStr, lngKeyHandle)

If lngKeyHandle = 0 Then
    regReadKey = vbNullString
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
End If

m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal 0&, lngBufferSize)

If lngKeyHandle = 0 Then
    regReadKey = vbNullString
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
End If

Select Case lngDataType
Case REG_SZ:
    strBuffer = Space(lngBufferSize)
    m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, ByVal strBuffer, lngBufferSize)
    If m_lngRetVal <> ERROR_SUCCESS Then
        regReadKey = vbNullString
    Else
        intPosition = InStr(1, strBuffer, Chr(0))
        If intPosition > 0 Then
            regReadKey = Left(strBuffer, intPosition - 1)
        Else
            regReadKey = strBuffer
        End If
    End If
Case REG_DWORD:
    m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngBuffer, 4&)
    If m_lngRetVal <> ERROR_SUCCESS Then
        regReadKey = vbNullString
    Else
        regReadKey = lngBuffer
    End If
Case Else:
        regReadKey = vbNullString
End Select

m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

Public Sub WinHTML(ByVal Thema As String)
On Error Resume Next

Dim FiNa As String

FiNa = App.Path & "\Hilfe\" & App.ProductName & ".chm::/" & Thema & ".htm"

RetWe = HTMLHelp(0&, FiNa, HH_DISPLAY_TOC, 0&)

End Sub
