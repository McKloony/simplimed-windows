VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------------------
' clsInfo: Collects file metadata via Shell32.dll
' -------------------------------------------------------------------

Private m_FilPa As String
Private m_FilNa As String
Private m_DatCr As Date
Private m_DatMo As Date
Private m_FilSz As Currency
Private m_WidPx As Long
Private m_HeiPx As Long
Private m_ResDX As Long
Private m_ResDY As Long
Private m_PagCn As Long
Private m_WidMM As Double
Private m_HeiMM As Double
Private m_Comme As String

Public Property Get FilePath() As String
    FilePath = m_FilPa
End Property

Public Property Let FilePath(ByVal sNewVal As String)
    m_FilPa = sNewVal
End Property

Public Property Get FileName() As String
    FileName = GetShortName(m_FilNa)
End Property

Public Property Get DateCreated() As Date
    DateCreated = m_DatCr
End Property

Public Property Get DateModified() As Date
    DateModified = m_DatMo
End Property

Public Property Get FileSize() As Currency
    FileSize = m_FilSz
End Property

Public Property Get WidthPx() As Long
    WidthPx = m_WidPx
End Property

Public Property Get HeightPx() As Long
    HeightPx = m_HeiPx
End Property

Public Property Get ResDpiX() As Long
    ResDpiX = m_ResDX
End Property

Public Property Get ResDpiY() As Long
    ResDpiY = m_ResDY
End Property

Public Property Get PageCount() As Long
    PageCount = m_PagCn
End Property

Public Property Get WidthMM() As Double
    WidthMM = m_WidMM
End Property

Public Property Get HeightMM() As Double
    HeightMM = m_HeiMM
End Property

Public Property Get Comment() As String
    Comment = m_Comme
End Property

Private Function GetShortName(ByVal sFullName As String) As String
    Dim sRet As String
    sRet = sFullName
    
    If Left$(sRet, 1) = "P" And Len(sRet) > 42 Then
        If Mid$(sRet, 42, 1) = "_" Then
            sRet = Mid$(sRet, 43)
        End If
    ElseIf Left$(sRet, 2) = "TD" And Len(sRet) > 43 Then
        If Mid$(sRet, 43, 1) = "_" Then
            sRet = Mid$(sRet, 44)
        End If
    End If
    
    GetShortName = sRet
End Function

Public Sub ReadData()
    Dim oShell  As Object
    Dim oFolder As Object
    Dim oItem   As Object
    Dim sDir    As String
    Dim sFile   As String
    Dim oFSO    As Object
    Dim oFile   As Object
    Dim vVal    As Variant
    
    If Len(m_FilPa) = 0 Then Exit Sub
    
    On Error GoTo ErrHand
    
    ' 1. Robust Basic Data via FSO (Works on XP - Win11)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO.FileExists(m_FilPa) Then
        Set oFile = oFSO.GetFile(m_FilPa)
        m_FilNa = oFile.Name
        m_FilSz = CCur(oFile.SIZE)
        m_DatMo = oFile.DateLastModified
        m_DatCr = oFile.DateCreated
        
        ' Resolve absolute path for Shell
        sDir = oFSO.GetParentFolderName(oFile.Path)
        sFile = oFile.Name
    Else
        ' File not found - keep path but exit
        m_FilNa = GetFilNam(m_FilPa)
        Set oFSO = Nothing
        Exit Sub
    End If
    Set oFile = Nothing
    Set oFSO = Nothing
    
    ' 2. Extended Info via Shell
    Set oShell = CreateObject("Shell.Application")
    Set oFolder = oShell.Namespace(sDir)
    
    If Not oFolder Is Nothing Then
        Set oItem = oFolder.ParseName(sFile)
        If Not oItem Is Nothing Then
            
            ' A) Try Vista+ API (ExtendedProperty)
            On Error Resume Next
            vVal = oItem.ExtendedProperty("System.Image.HorizontalSize")
            If Err.Number = 0 And Not IsEmpty(vVal) Then
                m_WidPx = GetLngSaf(vVal)
                m_HeiPx = GetLngSaf(oItem.ExtendedProperty("System.Image.VerticalSize"))
                m_ResDX = GetLngSaf(oItem.ExtendedProperty("System.Image.HorizontalResolution"))
                m_ResDY = GetLngSaf(oItem.ExtendedProperty("System.Image.VerticalResolution"))
                m_PagCn = GetLngSaf(oItem.ExtendedProperty("System.Document.PageCount"))
                m_Comme = GetStrSaf(oItem.ExtendedProperty("System.Comment"))
            Else
                ' B) WinXP Fallback (GetDetailsOf)
                Err.Clear
                ' Search for Dimensions
                Dim idxDim As Long
                Dim idxPag As Long
                Dim sDim As String
                
                ' Note: Indexes vary. We search for standard names or use common XP IDs
                ' Simple heuristic: Try standard XP indices if search is too slow/complex
                ' Dimensions often 26 (XP), Pages often ?
                ' Better: Try to parse "Dimensions" string if found.
                
                ' Attempt to find "Dimensions" / "Abmessungen" column index?
                ' Skipping dynamic search for performance, try common fallback slots or rely on FSO/Basic only for XP
                ' But user wants "Dimensions".
                
                ' Try reading Dimensions via GetDetailsOf (Index 31 on Win10, 26 on XP commonly)
                ' We will try a few common indices for Dimensions if 0
                sDim = oFolder.GetDetailsOf(oItem, 31) ' Win7/10 approx
                If sDim = "" Then sDim = oFolder.GetDetailsOf(oItem, 26) ' XP
                
                ParseDimensions sDim ' Helper to extract w/h
                
                m_PagCn = GetLngSaf(oFolder.GetDetailsOf(oItem, 13)) ' Pages? High variance. Often unreliable on XP without specific handlers.
            End If
            On Error GoTo ErrHand
            
            ' Calculations
            If m_ResDX > 0 Then
                m_WidMM = (m_WidPx / m_ResDX) * 25.4
            Else
                 ' Assume 96 DPI if missing for calc
                 m_WidMM = (m_WidPx / 96) * 25.4
            End If
            
            If m_ResDY > 0 Then
                m_HeiMM = (m_HeiPx / m_ResDY) * 25.4
            Else
                m_HeiMM = (m_HeiPx / 96) * 25.4
            End If
            
        End If
    End If
    
    Set oItem = Nothing
    Set oFolder = Nothing
    Set oShell = Nothing
    Exit Sub
    
ErrHand:
    If GlLog = True Then SLogi "clsInfo.ReadData Error: " & Err.Description & " for " & m_FilPa
    ' Ensure cleanup
    Set oItem = Nothing
    Set oFolder = Nothing
    Set oShell = Nothing
    Set oFile = Nothing
    Set oFSO = Nothing
End Sub

Private Sub ParseDimensions(ByVal sDim As String)
    ' Expects format like "800 x 600" or "800 x 600 px"
    ' Very basic parser
    Dim Parts() As String
    Dim sClean As String
    Dim i As Long
    
    If Len(sDim) = 0 Then Exit Sub
    
    sClean = LCase$(sDim)
    sClean = Replace(sClean, "px", "")
    sClean = Replace(sClean, " ", "")
    ' Look for "x"
    If InStr(sClean, "x") > 0 Then
        Parts = Split(sClean, "x")
        If UBound(Parts) >= 1 Then
            m_WidPx = GetLngSaf(Parts(0))
            m_HeiPx = GetLngSaf(Parts(1))
        End If
    End If
End Sub

Public Property Get DisplayInfo() As String
    Dim sRet As String
    Dim sExt As String
    Dim sSize As String
    Dim sSpacer As String
    Dim sDispNa As String
    
    On Error GoTo ErrHand
    
    ' Definition fuer 1.5 Zeilenabstand (Simuliert durch kleine Leerzeile)
    sSpacer = "<LineBreak/><Run FontSize='6'> </Run><LineBreak/>"
    
    sExt = LCase$(Right$(m_FilPa, 4))
    sSize = FormatSize(m_FilSz)
    
    ' Dateinamen-Logik via Helper
    sDispNa = GetShortName(m_FilNa)
    
    ' Start Block (Basic)
    sRet = "<TextBlock>"
    
    ' Header: Filename (Fett, Schriftgröße 13)
    sRet = sRet & "<Bold><Run FontSize='13'>" & sDispNa & "</Run></Bold>" & sSpacer
    
    ' Basis-Informationen
    sRet = sRet & "<Run Foreground='Gray'>Geändert:</Run> <Run Foreground='Black'>" & Format$(m_DatMo, "dd.mm.yyyy - hh:mm") & "</Run>" & sSpacer
    sRet = sRet & "<Run Foreground='Gray'>Erstellt am:</Run> <Run Foreground='Black'>" & Format$(m_DatCr, "dd.mm.yyyy - hh:mm") & "</Run>" & sSpacer
    sRet = sRet & "<Run Foreground='Gray'>Dateigröße:</Run> <Run Foreground='Black'>" & sSize & "</Run>" & sSpacer
    
    ' Variant 1: PDF
    If sExt = ".pdf" Or m_PagCn > 0 Then
        sRet = sRet & "<Run Foreground='Gray'>Typ:</Run> <Bold>PDF Dokument</Bold>" & sSpacer
        If m_PagCn > 0 Then
            sRet = sRet & "<Run Foreground='Gray'>Seiten:</Run> <Run Foreground='Black'>" & m_PagCn & "</Run>" & sSpacer
        End If
        ' Versuch Dimensionen bei PDF anzuzeigen, falls Windows sie liefert
        If m_WidPx > 0 Then
             sRet = sRet & "<Run Foreground='Gray'>Aufloesung:</Run> <Run Foreground='Black'>" & m_WidPx & " x " & m_HeiPx & " px</Run>" & sSpacer
             If m_WidMM > 0 Then
                sRet = sRet & "<Run Foreground='Gray'>Papiergröße:</Run> <Run Foreground='Black'>" & Format$(m_WidMM, "0.0") & " x " & Format$(m_HeiMM, "0.0") & " mm</Run>" & sSpacer
             End If
        End If
        
    ' Variant 2: Bilddatei
    ElseIf m_WidPx > 0 Or IsImageExt(sExt) Then
        sRet = sRet & "<Run Foreground='Gray'>Typ:</Run> <Bold>Bilddatei</Bold>" & sSpacer
        If m_WidPx > 0 Then
            sRet = sRet & "<Run Foreground='Gray'>Auflösung:</Run> <Run Foreground='Black'>" & m_WidPx & " x " & m_HeiPx & " px</Run>" & sSpacer
            sRet = sRet & "<Run Foreground='Gray'>DPI:</Run> <Run Foreground='Black'>" & m_ResDX & " DPI</Run>" & sSpacer
            sRet = sRet & "<Run Foreground='Gray'>Papiergröße:</Run> <Run Foreground='Black'>" & Format$(m_WidMM, "0.0") & " x " & Format$(m_HeiMM, "0.0") & " mm</Run>" & sSpacer
        End If
        
    ' Variant 3: Sonstige
    Else
        sRet = sRet & "<Run Foreground='Gray'>Typ:</Run> <Italic>Sonstige Datei</Italic>" & sSpacer
    End If
    
    ' Kommentar
    If Len(m_Comme) > 0 Then
        sRet = sRet & "<Run Foreground='Gray'>Kommentar:</Run> <Run Foreground='Blue'>" & m_Comme & "</Run>" & sSpacer
    End If
    
    sRet = sRet & "</TextBlock>"
    
    DisplayInfo = sRet
    Exit Property

ErrHand:
    If GlLog = True Then SLogi "clsInfo.DisplayInfo Error: " & Err.Description
    DisplayInfo = "<TextBlock><Run Foreground='Red'>Fehler beim Erstellen der Info.</Run></TextBlock>"
End Property

Private Function FormatSize(ByVal curBytes As Currency) As String
    If curBytes >= 1048576 Then
        FormatSize = Format$(curBytes / 1048576, "0.00") & " MB"
    ElseIf curBytes >= 1024 Then
        FormatSize = Format$(curBytes / 1024, "0.00") & " KB"
    Else
        FormatSize = curBytes & " Bytes"
    End If
End Function

Private Function IsImageExt(ByVal sExt As String) As Boolean
    Select Case sExt
        Case ".bmp", ".jpg", ".png", ".gif", ".tif", "tiff", "jpeg"
            IsImageExt = True
        Case Else
            IsImageExt = False
    End Select
End Function

Private Function GetDirNam(ByVal sPath As String) As String
    Dim iPos As Long
    iPos = InStrRev(sPath, "\")
    If iPos > 0 Then
        GetDirNam = Left$(sPath, iPos - 1)
        If Right$(GetDirNam, 1) = ":" Then GetDirNam = GetDirNam & "\"
    End If
End Function

Private Function GetFilNam(ByVal sPath As String) As String
    Dim iPos As Long
    iPos = InStrRev(sPath, "\")
    If iPos > 0 Then
        GetFilNam = Mid$(sPath, iPos + 1)
    Else
        GetFilNam = sPath
    End If
End Function

Private Function GetLngSaf(ByVal vVal As Variant) As Long
    On Error Resume Next
    GetLngSaf = CLng(vVal)
    On Error GoTo 0
End Function

Private Function GetCurSaf(ByVal vVal As Variant) As Currency
    On Error Resume Next
    GetCurSaf = CCur(vVal)
    On Error GoTo 0
End Function

Private Function GetDatSaf(ByVal vVal As Variant) As Date
    On Error Resume Next
    If IsDate(vVal) Then GetDatSaf = CDate(vVal)
    On Error GoTo 0
End Function

Private Function GetStrSaf(ByVal vVal As Variant) As String
    On Error Resume Next
    GetStrSaf = CStr(vVal)
    On Error GoTo 0
End Function

