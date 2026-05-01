Attribute VB_Name = "basLabor"
Option Explicit

' basLabor.bas
' Fully rewritten LDT Import Module with Debug Logging and Safe Recordset Handling
' Native VB6 implementation for LDT2 and LDT3 including PDF extraction.
' Uses clsFile for robust Windows API file handling.

' --- Windows API Declarations ---
Private Declare Function CryptStringToBinary Lib "crypt32.dll" Alias "CryptStringToBinaryA" ( _
    ByVal pszString As String, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As Long, _
    ByRef pcbBinary As Long, _
    ByRef pdwSkip As Long, _
    ByRef pdwFlags As Long) As Long

Private Const CRYPT_STRING_BASE64 As Long = 1

' --- Internal Recordsets ---
Private RS_Dat As ADODB.Recordset
Private RS_Fil As ADODB.Recordset
Private RS_Rep As ADODB.Recordset
Private RS_Val As ADODB.Recordset
Private RS_Tmp As ADODB.Recordset

' --- Global/Module Variables ---
Private m_TempPDFPath As String
Private m_LogFile As String
Private m_CurrentFile As String ' Store current filename for logging path
Private m_OrderNumber As String ' Order number (Auftragsnummer) from LDT3 for PDF naming

' ==========================================================================================
' LOGGING HELPER
' ==========================================================================================

Private Sub L_Sta(ByVal StepNr As Integer, ByVal StepText As String)
    ' Update progress bar 1 (main steps) and allow UI refresh
    On Error Resume Next
    If Not frmStatus Is Nothing Then
        frmStatus.Caption = StepText
        If StepNr <= frmStatus.prbStat1.Max Then
            frmStatus.prbStat1.Value = StepNr
        End If
        frmStatus.Refresh
        DoEvents
    End If
End Sub

Private Sub L_Pct(ByVal Percent As Long)
    ' Update progress bar 2 (percentage within current step)
    On Error Resume Next
    If Not frmStatus Is Nothing Then
        If Percent >= 0 And Percent <= 100 Then
            frmStatus.prbStat2.Value = Percent
        End If
        frmStatus.Refresh
        DoEvents
    End If
End Sub

Private Sub L_Log(ByVal sMsg As String)
On Error Resume Next

Dim f As Integer
Dim sPath As String

' Only log if debug mode is enabled
If GlDbg = False Then Exit Sub

' Try to determine log path from current file, else fallback to App.Path
If m_CurrentFile <> "" Then
    If InStrRev(m_CurrentFile, "\") > 0 Then
        sPath = Left$(m_CurrentFile, InStrRev(m_CurrentFile, "\"))
    Else
        sPath = App.Path & "\"
    End If
Else
    sPath = App.Path & "\"
End If

' Ensure trailing slash
If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"

m_LogFile = sPath & "LDT_Import_Debug.log"

f = FreeFile
Open m_LogFile For Append As #f
Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & sMsg
Close #f

End Sub

' ==========================================================================================
' PUBLIC ENTRY POINT
' ==========================================================================================

Public Sub L_Imp(ByVal FiNam As String)
    On Error GoTo ErrHandler
    
    m_CurrentFile = FiNam ' Set for logging
    
    L_Log "=== START IMPORT ==="
    L_Log "File: " & FiNam
    
    Dim ImDat As String
    Dim IsLDT3 As Boolean
    Dim oFile As clsFile
    
    ' 1. Read File Content using clsFile
    Set oFile = New clsFile
    oFile.FilPfa FiNam
    
    ' FilReSt reads the entire file via API
    ImDat = oFile.FilReSt()
    
    If Len(ImDat) = 0 Then
        L_Log "Error: File is empty or could not be read."
        MsgBox "Datei ist leer oder konnte nicht gelesen werden: " & FiNam, vbExclamation
        Exit Sub
    End If
    L_Log "File read successfully. Size: " & Len(ImDat)

    ' Initialize Status Form early - show progress during parsing
    If Not frmStatus Is Nothing Then
        frmStatus.Caption = "LDT Import..."
        frmStatus.prbStat1.Min = 0
        frmStatus.prbStat1.Max = 9
        frmStatus.prbStat1.Value = 0
        frmStatus.prbStat2.Min = 0
        frmStatus.prbStat2.Max = 100
        frmStatus.prbStat2.Value = 0
        frmStatus.Show
        DoEvents
    End If

    ' 2. Detect LDT Version (Check first 500 chars)
    If InStr(1, Left$(ImDat, 500), "LDT3") > 0 Or InStr(1, Left$(ImDat, 500), "0138002") > 0 Then
        IsLDT3 = True
        L_Log "Detected Version: LDT3"
    Else
        L_Log "Detected Version: LDT2 (Legacy)"
    End If

    ' 3. Prepare Temporary Database Table (Step 1)
    L_Sta 1, "Lese LDT-Datei..."
    L_Log "Clearing qryLdtLoe..."
    DBCmEx0 "qryLdtLoe"

    Set RS_Dat = New ADODB.Recordset
    L_Log "Opening RS_Dat (qryLdtDat)..."
    With RS_Dat
        .CursorLocation = adUseClient
        .Source = "qryLdtDat"
        .ActiveConnection = DB1
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open Options:=adCmdTable
    End With
    L_Log "RS_Dat opened."

    If IsLDT3 Then
        L_Par3 ImDat, FiNam
    Else
        L_Par2 ImDat
    End If

    L_Log "Parsing finished. Closing RS_Dat."
    RS_Dat.Close
    Set RS_Dat = Nothing
    Set oFile = Nothing ' Clean up

    ' 4. Execute Import Logic (Process the data now in qryLdtDat)
    
    Dim ImpDa As Long, VerNu As Long, ManNr As Long, LaKat As Long
    Dim LaGui As String
    
    If GlLaM > 0 Then ManNr = GlLaM Else ManNr = GlMan(GlSMa, 2)
    LaKat = GlLab(GlStL, 0)
    LaGui = CreateID("L")
    L_Log "Generated LaGui: " & LaGui

    ' Process Headers (Step 2)
    L_Log "Calling L_Hea..."
    L_Sta 2, "Importiere Kopfdaten..."
    L_Hea LaGui, VerNu, ImpDa
    L_Log "L_Hea finished. ImpDa: " & ImpDa & ", VerNu: " & VerNu

    ' Process Reports (Step 3)
    L_Log "Calling L_Rep..."
    L_Sta 3, "Importiere Berichte..."
    L_Rep LaGui, ImpDa, ManNr, LaKat, VerNu
    L_Log "L_Rep finished."

    ' Process Values (Step 4)
    L_Log "Calling L_Val..."
    L_Sta 4, "Importiere Messwerte..."
    L_Val ImpDa, VerNu
    L_Log "L_Val finished."

    ' Process Normal Ranges (Step 5)
    L_Log "Calling L_Nor..."
    L_Sta 5, "Importiere Normwerte..."
    L_Nor ImpDa
    L_Log "L_Nor finished."

    ' Process Maps (Step 6)
    L_Log "Calling L_Map..."
    L_Sta 6, "Verarbeite Zuordnungen..."
    L_Map ImpDa
    L_Log "L_Map finished."

    ' Process Comments (Step 7)
    L_Log "Calling L_Com..."
    L_Sta 7, "Verarbeite Kommentare..."
    L_Com ImpDa
    L_Log "L_Com finished."

    ' Process Patients (Step 8)
    L_Log "Calling L_Pat..."
    L_Sta 8, "Ordne Patienten zu..."
    L_Pat ImpDa
    L_Log "L_Pat finished."

    ' Final DB Cleanups (Step 9)
    L_Sta 9, "Finalisiere Import..."
    DBCmEx0 "qryLabPaNr"
    DBCmEx0 "qryLabSam"

    ' 5. PDF was already saved during parsing with final name (ldt<OrderNumber>.pdf)
    ' No further action needed - user will import manually later
    If m_TempPDFPath <> "" Then
        L_Log "PDF saved at: " & m_TempPDFPath
        m_TempPDFPath = ""
    End If
    
    L_Log "=== IMPORT SUCCESSFUL ==="
    Exit Sub

ErrHandler:
    L_Log "CRITICAL ERROR in L_Imp: " & Err.Description & " (" & Err.Number & ")"
    MsgBox "L_Imp Fehler: " & Err.Description & " (" & Err.Number & ")", vbCritical
    Set oFile = Nothing
End Sub

' ==========================================================================================
' PARSING LOGIC (NATIVE VB6)
' ==========================================================================================

Private Sub L_Par2(ByRef Content As String)
    ' Standard LDT2 Parsing into qryLdtDat
    Dim Posit As Long, StaWe As Long, Lange As Long
    Dim FldKn As Long, FldLa As Long, BerAr As Long
    Dim FldIn As String, AkZei As String
    Dim LineCount As Long
    Dim LastPct As Long, CurPct As Long

    L_Log "Entering L_Par2"

    Lange = Len(Content)
    StaWe = 1
    LineCount = 0
    LastPct = 0
    Posit = InStr(StaWe, Content, Chr$(10))
    If Posit = 0 Then Posit = InStr(StaWe, Content, Chr$(13))

    Do While Posit > 0
        LineCount = LineCount + 1
        ' Update progress every 1% of file parsed
        CurPct = (StaWe * 100) \ Lange
        If CurPct > LastPct Then
            LastPct = CurPct
            L_Pct CurPct
        End If
        AkZei = Mid$(Content, StaWe, Posit - StaWe)
        StaWe = Posit + 1
        If Len(AkZei) > 3 Then
            ' Normalize Line Endings logic
            If Asc(Right$(AkZei, 1)) = 13 Then
                FldIn = Trim$(Mid$(AkZei, 8, Len(AkZei) - 8))
            Else
                FldIn = Trim$(Mid$(AkZei, 8, Len(AkZei) - 7))
            End If
            
            If FldIn <> vbNullString Then
                If GlIFo = "X2" Then FldIn = SZech(FldIn, True)
                
                If IsNumeric(Mid$(AkZei, 4, 4)) Then FldKn = CLng(Mid$(AkZei, 4, 4))
                If FldKn = 8000 Then BerAr = Val(Trim$(FldIn))
                
                If IsNumeric(Left$(AkZei, 3)) Then
                    RS_Dat.AddNew
                    FldLa = CLng(Left$(AkZei, 3))
                    RS_Dat.Fields("L" & Chr$(228) & "nge").Value = FldLa
                    RS_Dat.Fields("Feldkennung").Value = FldKn
                    RS_Dat.Fields("Feldinhalt").Value = FldIn
                    RS_Dat.Fields("Berichtsart").Value = BerAr
                    RS_Dat.Update
                End If
            End If
        End If
        Posit = InStr(StaWe, Content, Chr$(10))
        If Posit = 0 Then Posit = InStr(StaWe, Content, Chr$(13))
    Loop
    L_Log "Leaving L_Par2"
End Sub

Private Sub L_Par3(ByRef Content As String, ByVal OriginalPath As String)
    ' LDT3 Parsing: Extract PDF and Map fields to Legacy codes for qryLdtDat
    Dim Posit As Long, StaWe As Long, Lange As Long
    Dim FldKn As Long
    Dim FldIn As String, AkZei As String
    Dim Base64Buffer As String
    Dim InPdfBlock As Boolean
    Dim Folder As String
    Dim LineCount As Long
    Dim LastPct As Long, CurPct As Long

    L_Log "Entering L_Par3"

    m_TempPDFPath = ""
    m_OrderNumber = ""
    Lange = Len(Content)
    StaWe = 1
    LineCount = 0
    LastPct = 0

    ' Initialize simulated headers for Legacy Logic
    L_Add 8000, "8218" ' Start Report
    L_Add 9212, "0003" ' Version Simulation

    Posit = InStr(StaWe, Content, Chr$(10))
    If Posit = 0 Then Posit = InStr(StaWe, Content, Chr$(13))

    Do While Posit > 0
        LineCount = LineCount + 1
        ' Update progress every 1% of file parsed
        CurPct = (StaWe * 100) \ Lange
        If CurPct > LastPct Then
            LastPct = CurPct
            L_Pct CurPct
        End If
        AkZei = Mid$(Content, StaWe, Posit - StaWe)
        StaWe = Posit + 1

        ' Minimal LDT structure check: Len(3) + Code(4)
        If Len(AkZei) >= 7 Then
             If IsNumeric(Mid$(AkZei, 4, 4)) Then
                FldKn = CLng(Mid$(AkZei, 4, 4))
                ' Extract Content (skip length 3 + code 4 = 7 chars)
                FldIn = Mid$(AkZei, 8)
                ' Cleanup CRLF if stuck at end
                FldIn = Replace(FldIn, vbCr, "")
                FldIn = Replace(FldIn, vbLf, "")
                
                ' --- PDF Extraction Logic ---
                If FldKn = 6329 Then ' Embedded Base64 Content
                    Base64Buffer = Base64Buffer & FldIn
                    InPdfBlock = True
                ElseIf InPdfBlock And FldKn <> 6329 Then
                    ' End of PDF block detected
                    InPdfBlock = False
                End If
                
                ' --- Data Mapping to Legacy Codes ---
                Select Case FldKn
                    ' Patient
                    Case 3101: L_Add 3101, FldIn ' Name
                    Case 3102: L_Add 3102, FldIn ' Vorname
                    Case 3103: L_Add 3103, FldIn ' DOB
                    Case 3110: L_Add 3110, FldIn ' Geschlecht

                    ' Order number (Auftragsnummer) - capture first non-empty value
                    Case 8310, 8311
                        L_Add FldKn, FldIn
                        If m_OrderNumber = "" And FldIn <> "" Then
                            m_OrderNumber = Trim$(FldIn)
                            L_Log "Order number captured: " & m_OrderNumber
                        End If

                    ' Labor
                    Case 1250: L_Add 8300, FldIn ' Lab Name
                    Case 7358: L_Add 203, FldIn  ' Arzt Name (Reporting)
                    
                    ' Values
                    Case 8410: L_Add 8410, FldIn ' Test ID/Kuerzel
                    Case 8411: L_Add 8411, FldIn ' Test Name
                    Case 8420: L_Add 8420, FldIn ' Value
                    Case 8421: L_Add 8421, FldIn ' Unit
                    Case 8460: L_Add 8460, FldIn ' Normal Range / Text
                    Case 8422: L_Add 8422, FldIn ' Status/Flag
                    Case 7278: L_Add 8432, FldIn ' Date
                    
                    ' Comments
                    Case 8470: L_Add 8470, FldIn

                    ' Billing fields
                    Case 5001: L_Add 5001, FldIn ' Fee code (GONr/Gebührenziffer)
                    Case 8406: L_Add 8406, FldIn ' Amount in cents (Betrag)
                End Select
             End If
        End If
        
        Posit = InStr(StaWe, Content, Chr$(10))
        If Posit = 0 Then Posit = InStr(StaWe, Content, Chr$(13))
    Loop
    
    ' Save extracted PDF if found - use order number in filename
    If Len(Base64Buffer) > 0 Then
        Folder = Left(OriginalPath, InStrRev(OriginalPath, "\"))
        ' Build filename: ldt<OrderNumber>.pdf
        Dim OrdNr As String
        OrdNr = m_OrderNumber
        If OrdNr = "" Then OrdNr = Format$(Now, "yyyymmddhhnnss") ' Fallback
        ' Sanitize for filename
        OrdNr = Replace(OrdNr, "/", "_")
        OrdNr = Replace(OrdNr, "\", "_")
        OrdNr = Replace(OrdNr, ":", "_")
        OrdNr = Replace(OrdNr, "*", "_")
        OrdNr = Replace(OrdNr, "?", "_")
        OrdNr = Replace(OrdNr, """", "_")
        OrdNr = Replace(OrdNr, "<", "_")
        OrdNr = Replace(OrdNr, ">", "_")
        OrdNr = Replace(OrdNr, "|", "_")
        m_TempPDFPath = Folder & "ldt" & OrdNr & ".pdf"
        If L_B64(Base64Buffer, m_TempPDFPath) = False Then
            m_TempPDFPath = "" ' Failed
            L_Log "Error: Base64 Decode failed"
        Else
            L_Log "PDF Extracted successfully to: " & m_TempPDFPath
        End If
    End If
    
    L_Add 8003, "8218" ' End Report
    L_Log "Leaving L_Par3"
End Sub

Private Sub L_Add(ByVal code As Long, ByVal ValStr As String)
    ' Helper to populate qryLdtDat
    If RS_Dat Is Nothing Then Exit Sub
    RS_Dat.AddNew
    RS_Dat.Fields("L" & Chr$(228) & "nge").Value = Len(ValStr) + 9
    RS_Dat.Fields("Feldkennung").Value = code
    RS_Dat.Fields("Feldinhalt").Value = ValStr
    RS_Dat.Fields("Berichtsart").Value = 0
    RS_Dat.Update
End Sub

' ==========================================================================================
' PDF / BASE64 HANDLING
' ==========================================================================================

Private Function L_B64(ByVal Base64Str As String, ByVal OutPath As String) As Boolean
    ' Native VB6 Base64 Decode using Crypt32, Writing via clsFile
    Dim bBytes() As Byte
    Dim lLen As Long
    Dim lFlags As Long
    Dim lRet As Long
    Dim oFile As clsFile
    
    lFlags = CRYPT_STRING_BASE64
    
    ' 1. Get Length
    lRet = CryptStringToBinary(Base64Str, Len(Base64Str), lFlags, 0, lLen, 0, 0)
    If lRet = 0 Then Exit Function
    
    ' 2. Resize Buffer
    ReDim bBytes(lLen - 1)
    
    ' 3. Decode
    lRet = CryptStringToBinary(Base64Str, Len(Base64Str), lFlags, VarPtr(bBytes(0)), lLen, 0, 0)
    If lRet = 0 Then Exit Function
    
    ' 4. Write to File using clsFile
    Set oFile = New clsFile
    oFile.FilPfa OutPath
    
    ' FilWeBy handles file creation/overwriting safely
    If oFile.FilWeBy(bBytes, False) Then
        L_B64 = True
    End If
    
    Set oFile = Nothing
End Function
' Import LDT File Header
Private Sub L_Hea(ByVal LaGui As String, ByRef VerNu As Long, ByRef ImpDa As Long)
    Dim FldKn As Long, TmpIn As String
    Dim RS_Src As ADODB.Recordset
    
    On Error GoTo Err_Hea
    
    Set RS_Src = New ADODB.Recordset
    RS_Src.CursorLocation = adUseClient
    If GlTyp < 2 Then
        RS_Src.Open "SELECT * FROM dbo.qryLdtDat ORDER BY ID0", DB1, adOpenKeyset, adLockReadOnly
    Else
        RS_Src.Open "SELECT * FROM qryLdtDat ORDER BY [ID0];", DB1, adOpenKeyset, adLockReadOnly
    End If
    
    If RS_Src.RecordCount = 0 Then
        L_Log "L_Hea: qryLdtDat is empty."
        Exit Sub
    End If
    
    Set RS_Fil = New ADODB.Recordset
    L_Log "L_Hea: Opening qryLdtFile for AddNew..."
    With RS_Fil
        .CursorLocation = adUseClient
        .Source = "qryLdtFile"
        .ActiveConnection = DB1
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open Options:=adCmdTable
    End With
    RS_Fil.AddNew
    RS_Fil.Fields("GuiID").Value = LaGui
    
    RS_Src.MoveFirst
    Do While Not RS_Src.EOF
        TmpIn = Trim$(RS_Src.Fields("Feldinhalt").Value)
        FldKn = CLng(RS_Src.Fields("Feldkennung").Value)
        
        Select Case FldKn
            Case 9211, 9212:
                If TmpIn <> vbNullString Then
                    RS_Fil.Fields("Version").Value = TmpIn
                    If Left$(TmpIn, 3) = "LDT" And IsNumeric(Mid$(TmpIn, 4, 4)) Then
                        VerNu = CLng(Mid$(TmpIn, 4, 4))
                    Else
                        VerNu = IIf(FldKn = 9211, 1011, 1014)
                    End If
                End If
            Case 201: If TmpIn <> "" Then RS_Fil.Fields("Arztnummer").Value = TmpIn
            Case 203: If TmpIn <> "" Then RS_Fil.Fields("Arztname").Value = TmpIn
            Case 8300: If TmpIn <> "" Then RS_Fil.Fields("Labor").Value = Left$(TmpIn, 50)
            Case 8320: If TmpIn <> "" Then RS_Fil.Fields("Laborname").Value = TmpIn
        End Select
        RS_Src.MoveNext
    Loop
    RS_Fil.Update
    RS_Fil.Close
    RS_Src.Close
    
    Set RS_Tmp = New ADODB.Recordset
    RS_Tmp.CursorLocation = adUseClient
    If GlTyp < 2 Then
        RS_Tmp.Open "SELECT ID0 From dbo.qryLdtFile WHERE (GuiID = '" & LaGui & "')", DB1
    Else
        RS_Tmp.Open "SELECT [ID0] From qryLdtFile WHERE ([GuiID] = '" & LaGui & "');", DB1
    End If
    If RS_Tmp.RecordCount > 0 Then ImpDa = RS_Tmp.Fields("ID0").Value
    RS_Tmp.Close
    Exit Sub
    
Err_Hea:
    L_Log "Error in L_Hea: " & Err.Description
    Resume Next
End Sub

' Import Report Headers
Private Sub L_Rep(ByVal LaGui As String, ByVal ImpDa As Long, ByVal ManNr As Long, ByVal LaKat As Long, ByVal VerNu As Long)
    Dim FldKn As Long, TypNr As Integer
    Dim FldNa As String, TmpIn As String
    Dim RS_Src As ADODB.Recordset
    Dim InReport As Boolean

    On Error GoTo Err_Rep

    Set RS_Src = New ADODB.Recordset
    RS_Src.CursorLocation = adUseClient
    If GlTyp < 2 Then
        RS_Src.Open "SELECT * FROM dbo.qryLdtDat ORDER BY ID0", DB1, adOpenForwardOnly, adLockReadOnly
    Else
        RS_Src.Open "SELECT * FROM qryLdtDat ORDER BY [ID0];", DB1, adOpenForwardOnly, adLockReadOnly
    End If

    Set RS_Rep = New ADODB.Recordset
    L_Log "L_Rep: Opening qryLdtIm5 for AddNew..."
    With RS_Rep
        .CursorLocation = adUseClient
        .Source = "qryLdtIm5"
        .ActiveConnection = DB1
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open Options:=adCmdTable
    End With

    InReport = False
    Do While Not RS_Src.EOF
        FldNa = "": TypNr = 0
        TmpIn = Trim$(RS_Src.Fields("Feldinhalt").Value)
        FldKn = CLng(RS_Src.Fields("Feldkennung").Value)

        If FldKn = 8000 Then
            If TmpIn = "8201" Or TmpIn = "8202" Or TmpIn = "8218" Or TmpIn = "8240" Then
                If InReport Then RS_Rep.Update
                RS_Rep.AddNew
                RS_Rep("GuiID").Value = LaGui
                RS_Rep("Berichtsart").Value = TmpIn
                RS_Rep("Importdatei").Value = ImpDa
                RS_Rep("IP0").Value = ManNr
                RS_Rep("ID3").Value = LaKat
                InReport = True
            End If
        ElseIf InReport Then
            Select Case FldKn
                Case 8310: FldNa = "Anforderung"
                Case 8311: FldNa = "LaborNr"
                Case 8301: FldNa = "LaborDat": TypNr = 1
                Case 8302: FldNa = "Berichtsdatum": TypNr = 1
                Case 3101: FldNa = "Pat_Name"
                Case 3102: FldNa = "Pat_Vorname"
                Case 3103: FldNa = "Pat_Geboren": TypNr = 1
                Case 3110: FldNa = "Pat_Geschl"
                           If TmpIn = "1" Then TmpIn = "m" & Chr$(228) & "nnlich"
                           If TmpIn = "2" Then TmpIn = "weiblich"
                Case 8401: FldNa = "Befundart"
                           If TmpIn = "E" Then TmpIn = "Endbefund"
                           If TmpIn = "T" Then TmpIn = "Teilbefund"
            End Select

            If FldNa <> "" And TmpIn <> "" Then
                On Error Resume Next
                If TypNr = 0 Then RS_Rep.Fields(FldNa).Value = TmpIn
                If TypNr = 1 Then RS_Rep.Fields(FldNa).Value = L_Dat(CStr(TmpIn), VerNu)
                On Error GoTo Err_Rep
            End If
        End If
        RS_Src.MoveNext
    Loop
    If InReport Then RS_Rep.Update
    RS_Rep.Close
    RS_Src.Close
    Exit Sub

Err_Rep:
    L_Log "Error in L_Rep: " & Err.Description
    Resume Next
End Sub

' Import Values
Private Sub L_Val(ByVal ImpDa As Long, ByVal VerNu As Long)
    Dim FldKn As Long, FldNa As String, TypNr As Integer
    Dim TmpIn As Variant, BerNr As Long
    Dim RS_Src As ADODB.Recordset
    
    On Error GoTo Err_Val
    
    Set RS_Rep = DBCmRe1("qryLdtIm6", "@IdxNr", ImpDa)
    
    Set RS_Val = New ADODB.Recordset
    L_Log "L_Val: Opening qryLdtIm7..."
    With RS_Val
        .CursorLocation = adUseClient
        .Source = "qryLdtIm7"
        .ActiveConnection = DB1
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open Options:=adCmdTable
    End With
    
    Set RS_Src = New ADODB.Recordset
    RS_Src.CursorLocation = adUseClient
    If GlTyp < 2 Then
        RS_Src.Open "SELECT * FROM dbo.qryLdtDat ORDER BY ID0", DB1, adOpenForwardOnly, adLockReadOnly
    Else
        RS_Src.Open "SELECT * FROM qryLdtDat ORDER BY [ID0];", DB1, adOpenForwardOnly, adLockReadOnly
    End If
    
    BerNr = 0
    Do While Not RS_Src.EOF
        FldNa = "": TypNr = 0
        TmpIn = Trim$(RS_Src.Fields("Feldinhalt").Value)
        FldKn = CLng(RS_Src.Fields("Feldkennung").Value)
        
        If FldKn = 8000 Then
            If TmpIn = "8201" Or TmpIn = "8202" Or TmpIn = "8218" Then
                 If Not RS_Rep.EOF Then
                     BerNr = RS_Rep.Fields("ID0").Value
                     RS_Rep.MoveNext
                 End If
            End If
        ElseIf BerNr > 0 Then
            Select Case FldKn
                Case 8410: ' Test ID
                    RS_Val.AddNew
                    RS_Val("Ident").Value = UCase(Left$(TmpIn, 8))
                    RS_Val("BerNr").Value = BerNr
                Case 8411: FldNa = "Testbezeichnung"
                Case 8420: FldNa = "Ergebniswert"
                Case 8421: FldNa = "Einheit"
                Case 8422: FldNa = "Grenz"
                Case 8432: FldNa = "Datum": TypNr = 1
                Case 8433: FldNa = "Zeit": TypNr = 3
                Case 8406: FldNa = "Betrag": TypNr = 4
                Case 5001: FldNa = "GONr"
            End Select
            
            If FldNa <> "" And TmpIn <> "" Then
                ' On Error Resume Next
                If TypNr = 0 Then RS_Val(FldNa).Value = TmpIn
                If TypNr = 1 Then RS_Val(FldNa).Value = L_Dat(CStr(TmpIn), VerNu)
                If TypNr = 3 Then RS_Val(FldNa).Value = Left$(TmpIn, 2) & ":" & Right$(TmpIn, 2)
                If TypNr = 4 Then RS_Val(FldNa).Value = CSng(TmpIn) / 100
                ' On Error GoTo 0
            End If
        End If
        RS_Src.MoveNext
    Loop
    If RS_Val.EditMode <> adEditNone Then RS_Val.Update
    RS_Src.Close
    RS_Rep.Close
    RS_Val.Close
    Exit Sub

Err_Val:
    L_Log "Error in L_Val: " & Err.Description
    Resume Next
End Sub

' Import Normal Range Texts into Tabelle_Lab_Norm
Private Sub L_Nor(ByVal ImpDa As Long)
    Dim TesNr As Long
    Dim TmpIn As String
    Dim RS_Src As ADODB.Recordset
    Dim RS_Nrm As ADODB.Recordset

    On Error GoTo Err_Nor

    Set RS_Val = DBCmRe1("qryLdtIm8", "@IdxNr", ImpDa)
    If RS_Val Is Nothing Then Exit Sub
    If RS_Val.EOF And RS_Val.BOF Then
        L_Log "L_Nor: No test records found - skipping."
        RS_Val.Close
        Exit Sub
    End If

    Set RS_Nrm = New ADODB.Recordset
    RS_Nrm.CursorLocation = adUseClient
    L_Log "L_Nor: Opening qryLabAuNo for AddNew..."
    If GlTyp < 2 Then
        RS_Nrm.Open "SELECT * FROM dbo.qryLabAuNo", DB1, adOpenDynamic, adLockOptimistic
    Else
        RS_Nrm.Open "SELECT * FROM qryLabAuNo;", DB1, adOpenDynamic, adLockOptimistic
    End If

    Set RS_Src = New ADODB.Recordset
    RS_Src.CursorLocation = adUseClient
    If GlTyp < 2 Then
        RS_Src.Open "SELECT * FROM dbo.qryLdtDat ORDER BY ID0", DB1, adOpenForwardOnly, adLockReadOnly
    Else
        RS_Src.Open "SELECT * FROM qryLdtDat ORDER BY [ID0];", DB1, adOpenForwardOnly, adLockReadOnly
    End If

    TesNr = 0
    Do While Not RS_Src.EOF
        TmpIn = Trim$(RS_Src.Fields("Feldinhalt").Value)
        If RS_Src.Fields("Feldkennung").Value = 8410 Then
            If Not RS_Val.EOF Then
                TesNr = RS_Val("ID0").Value
                RS_Val.MoveNext
            Else
                TesNr = 0
            End If
        ElseIf RS_Src.Fields("Feldkennung").Value = 8460 And TesNr > 0 Then
            If TmpIn <> "" Then
                RS_Nrm.AddNew
                RS_Nrm("Normwert").Value = Left$(TmpIn, 250)
                RS_Nrm("TID").Value = TesNr
                RS_Nrm.Update
            End If
        End If
        RS_Src.MoveNext
    Loop

    RS_Src.Close
    RS_Val.Close
    RS_Nrm.Close
    L_Log "L_Nor: Normal range import completed."
    Exit Sub

Err_Nor:
    L_Log "Error in L_Nor: " & Err.Description
    Resume Next
End Sub

' Import Maps
Private Sub L_Map(ByVal ImpDa As Long)
    Dim TesNr As Long, TmpIn As String
    Dim RS_Src As ADODB.Recordset
    Dim RS_Erg As ADODB.Recordset

    On Error GoTo Err_Map

    Set RS_Val = DBCmRe1("qryLdtIm8", "@IdxNr", ImpDa)
    If RS_Val Is Nothing Then Exit Sub
    If RS_Val.EOF And RS_Val.BOF Then
        RS_Val.Close
        Exit Sub
    End If

    Set RS_Erg = New ADODB.Recordset
    RS_Erg.CursorLocation = adUseClient
    L_Log "L_Map: Opening qryLabErgTe..."
    RS_Erg.Open "SELECT * FROM qryLabErgTe WHERE ID0 = -1", DB1, adOpenDynamic, adLockOptimistic

    Set RS_Src = New ADODB.Recordset
    RS_Src.CursorLocation = adUseClient
    RS_Src.Open "SELECT * FROM qryLdtDat ORDER BY ID0", DB1, adOpenForwardOnly, adLockReadOnly

    Do While Not RS_Src.EOF
        TmpIn = Trim$(RS_Src.Fields("Feldinhalt").Value)
        If RS_Src.Fields("Feldkennung").Value = 8410 Then
            If Not RS_Val.EOF Then
                TesNr = RS_Val("ID0").Value
                RS_Val.MoveNext
            End If
        ElseIf RS_Src.Fields("Feldkennung").Value = 8480 And TesNr > 0 Then
            RS_Erg.AddNew
            RS_Erg("Ergebnis").Value = TmpIn
            RS_Erg("TestID").Value = TesNr
            RS_Erg.Update
        End If
        RS_Src.MoveNext
    Loop
    RS_Src.Close
    RS_Val.Close
    RS_Erg.Close
    Exit Sub

Err_Map:
    L_Log "Error in L_Map: " & Err.Description
    Resume Next
End Sub

' Import Comments
Private Sub L_Com(ByVal ImpDa As Long)
    Dim RS_Com As ADODB.Recordset
    On Error GoTo Err_Com
    
    ' Replaced DBCmRe1 with manual open for updateable cursor
    Set RS_Com = New ADODB.Recordset
    RS_Com.CursorLocation = adUseClient
    L_Log "L_Com: Opening qryLdtIm9 for Update..."
    
    ' Construct the SQL string manually based on what qryLdtIm9 likely does (Select from qryLdtIm9/Table where Importdatei = ImpDa)
    If GlTyp < 2 Then
        RS_Com.Open "SELECT * FROM dbo.qryLdtIm9 WHERE Importdatei = " & ImpDa, DB1, adOpenDynamic, adLockOptimistic
    Else
        RS_Com.Open "SELECT * FROM qryLdtIm9 WHERE [Importdatei] = " & ImpDa, DB1, adOpenDynamic, adLockOptimistic
    End If
    
    Do While Not RS_Com.EOF
        If RS_Com("Grenz").Value & "@" = "@" And RS_Com("Ergebniswert").Value <> "0" Then
             RS_Com("Grenz").Value = "(*)"
             RS_Com.Update
        End If
        RS_Com.MoveNext
    Loop
    RS_Com.Close
    Set RS_Com = Nothing
    Exit Sub
    
Err_Com:
    L_Log "Error in L_Com: " & Err.Description
    Resume Next
End Sub

' Import Patient Match
Private Sub L_Pat(ByVal ImpDa As Long)
    Dim RS_Ber As ADODB.Recordset, RS_Auf As ADODB.Recordset
    Dim IdxSt As String
    Dim AufOK As Boolean

    On Error GoTo Err_Pat

    ' Open qryLdtIm9 for updateable cursor
    Set RS_Ber = New ADODB.Recordset
    RS_Ber.CursorLocation = adUseClient
    L_Log "L_Pat: Opening qryLdtIm9 for Update..."

    If GlTyp < 2 Then
        RS_Ber.Open "SELECT * FROM dbo.qryLdtIm9 WHERE Importdatei = " & ImpDa, DB1, adOpenDynamic, adLockOptimistic
    Else
        RS_Ber.Open "SELECT * FROM qryLdtIm9 WHERE [Importdatei] = " & ImpDa, DB1, adOpenDynamic, adLockOptimistic
    End If

    ' Open qryLdtAuf as SQL query (works for both Access and SQL Server)
    Set RS_Auf = New ADODB.Recordset
    RS_Auf.CursorLocation = adUseClient
    L_Log "L_Pat: Opening qryLdtAuf..."
    On Error Resume Next
    If GlTyp < 2 Then
        RS_Auf.Open "SELECT * FROM dbo.qryLdtAuf", DB1, adOpenKeyset, adLockReadOnly, adCmdText
    Else
        RS_Auf.Open "SELECT * FROM qryLdtAuf", DB1, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    If Err.Number <> 0 Then
        L_Log "L_Pat: qryLdtAuf could not be opened (" & Err.Description & ") - skipping patient match"
        Err.Clear
        On Error GoTo Err_Pat
        RS_Ber.Close
        Set RS_Ber = Nothing
        Set RS_Auf = Nothing
        Exit Sub
    End If
    Err.Clear
    On Error GoTo Err_Pat

    If RS_Auf.EOF And RS_Auf.BOF Then
        L_Log "L_Pat: qryLdtAuf is empty - skipping patient match"
        RS_Ber.Close
        RS_Auf.Close
        Set RS_Ber = Nothing
        Set RS_Auf = Nothing
        Exit Sub
    End If

    L_Log "L_Pat: Processing " & RS_Ber.RecordCount & " records against " & RS_Auf.RecordCount & " orders..."
    Do While Not RS_Ber.EOF
        If Not IsNull(RS_Ber("Auftrag").Value) Then
            IdxSt = Format$(RS_Ber("Auftrag").Value, "00000000")
            RS_Auf.MoveFirst
            RS_Auf.Find "[IDA] = '" & IdxSt & "'", 0, adSearchForward
            If Not RS_Auf.EOF Then
                RS_Ber("IDP").Value = RS_Auf("ID0").Value
                RS_Ber.Update
            End If
        End If
        RS_Ber.MoveNext
    Loop
    RS_Ber.Close
    RS_Auf.Close
    Set RS_Ber = Nothing
    Set RS_Auf = Nothing
    Exit Sub

Err_Pat:
    L_Log "Error in L_Pat: " & Err.Description
    Resume Next
End Sub

' Helpers
Private Function L_Dat(ByVal DaStr As String, Optional ByVal VerNu As Long) As Date
    ' Convert LDT date string (YYYYMMDD or DDMMYYYY) to Date type
    ' Returns Date type for SQL Server compatibility
    ' Supports both formats: YYYYMMDD (LDT3) and DDMMYYYY (LDT2)
    Dim yyyy As Integer, mm As Integer, dd As Integer

    On Error GoTo ErrDate
    DaStr = Trim$(DaStr)

    If Len(DaStr) = 8 And IsNumeric(DaStr) Then
        ' Try YYYYMMDD format first
        yyyy = CInt(Left$(DaStr, 4))
        mm = CInt(Mid$(DaStr, 5, 2))
        dd = CInt(Right$(DaStr, 2))
        ' Validate ranges
        If yyyy >= 1900 And yyyy <= 2100 And mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 Then
            L_Dat = DateSerial(yyyy, mm, dd)
            If GlDbg Then L_Log "Date parsed (YYYYMMDD): " & DaStr & " -> " & Format$(L_Dat, "yyyy-mm-dd")
            Exit Function
        End If

        ' If YYYYMMDD failed, try DDMMYYYY format
        dd = CInt(Left$(DaStr, 2))
        mm = CInt(Mid$(DaStr, 3, 2))
        yyyy = CInt(Right$(DaStr, 4))
        ' Validate ranges
        If yyyy >= 1900 And yyyy <= 2100 And mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 Then
            L_Dat = DateSerial(yyyy, mm, dd)
            If GlDbg Then L_Log "Date parsed (DDMMYYYY): " & DaStr & " -> " & Format$(L_Dat, "yyyy-mm-dd")
            Exit Function
        End If
    End If

ErrDate:
    L_Dat = DateSerial(1900, 1, 1)
    If GlDbg Then L_Log "Date parse FAILED: " & DaStr & " -> using fallback 1900-01-01"
End Function
