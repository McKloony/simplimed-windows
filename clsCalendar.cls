VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function SyncLogFile(ByVal pLogType As String, ByVal pUsername As String, ByVal pPassword As String, ByRef pFile As String, ByRef pStatus As String, ByRef pexception As String, ByRef pexceptiontrace As String, ByVal DaEmp As Boolean, Optional ByVal empId As String = vbNullString) As Boolean
On Error GoTo ErrorGenerated '~~~

SyncLogFile = False

Dim xmlDoc As New MSXML2.DOMDocument60
Dim xmlhttp As New MSXML2.ServerXMLHTTP60

Dim success As Long
Dim strServer As String
Dim responseText As String
Dim strUrl As String

strUrl = vbNullString

If DaEmp = True Then
    GenerateUrl pLogType, "receive", pUsername, pPassword, vbNullString, strUrl, empId
Else
    GenerateUrl pLogType, "send", pUsername, pPassword, pFile, strUrl, empId
End If

xmlhttp.Open "POST", strUrl, False
xmlhttp.setRequestHeader "Content-Type", "text/xml"
If GlPxV = True Then
    xmlhttp.setProxy 2, GlPxN
End If

xmlhttp.send
xmlhttp.waitForResponse 60
DoEvents

If xmlhttp.Status = 200 Then
    Dim statusStr As String
    Dim retExceptionStr As String
    Dim retExceptionTracestr As String
    Dim retFileStr As String

    ParseResponce xmlhttp.responseText, statusStr, retExceptionStr, retExceptionTracestr, retFileStr, xmlDoc
    pStatus = statusStr
    pexception = retExceptionStr
    pexceptiontrace = retExceptionTracestr
    pFile = retFileStr
    SyncLogFile = True
Else
    pStatus = "False"
    pexception = xmlhttp.Status & " " & xmlhttp.statusText
End If

Exit Function

ErrorGenerated:

pStatus = "False"
pexception = Err.Description
pexceptiontrace = Err.HelpFile
SyncLogFile = False

SPopu "Verbindungsfehler", pexception & " " & pexceptiontrace, IC48_Forbidden

If GlDbg = True Then
    SErLog Err.Description & " " & Err.HelpFile & " SyncLogFile " & Err.Number
    SErLog pexception & " " & pexceptiontrace & " " & pFile & " " & strUrl
End If
Resume Next
    
End Function
Public Function SendLogFile(ByVal pLogType As String, ByVal pUsername As String, ByVal pPassword As String, ByVal pFile As String, ByRef pStatus As String, ByRef pexception As String, ByRef pexceptiontrace As String, Optional empId As String = vbNullString) As Boolean
    SendLogFile = SyncLogFile(pLogType, pUsername, pPassword, pFile, pStatus, pexception, pexceptiontrace, False, empId)
End Function
Public Function GetLogFile(ByVal pLogType As String, ByVal pUsername As String, ByVal pPassword As String, ByRef pFile As String, ByRef pStatus As String, ByRef pexception As String, ByRef pexceptiontrace As String, Optional empId As String = vbNullString) As Boolean
    GetLogFile = SyncLogFile(pLogType, pUsername, pPassword, pFile, pStatus, pexception, pexceptiontrace, True, empId)
End Function
Public Function ParseResponce(ByVal html As String, ByRef statusStr As String, ByRef retExceptionStr As String, ByRef retExceptionTracestr As String, ByRef retFileStr As String, ByVal xmlDoc As MSXML2.DOMDocument60) As Boolean
On Error Resume Next
    
Dim Status As MSXML2.IXMLDOMElement

xmlDoc.loadXML (html)

Set Status = xmlDoc.selectSingleNode("//success")

statusStr = Status.nodeTypedValue

Dim retException As MSXML2.IXMLDOMElement

Set retException = xmlDoc.selectSingleNode("//exception")
retExceptionStr = retException.nodeTypedValue

Dim retExceptionTrace As MSXML2.IXMLDOMElement

Set retExceptionTrace = xmlDoc.selectSingleNode("//exceptiontrace")
retExceptionTracestr = retExceptionTrace.nodeTypedValue

Dim retFile As MSXML2.IXMLDOMElement

Set retFile = xmlDoc.selectSingleNode("//file")
retFileStr = retFile.nodeTypedValue

End Function
Private Function GenerateUrl(ByVal pLogType As String, ByVal pDirection As String, ByVal pUsername As String, ByVal pPassword As String, ByVal pFile As String, ByRef generatedUrl As String, Optional ByVal empId As String = vbNullString) As Boolean
On Error Resume Next '~~~

Dim direction As String
Dim protocol As String
Dim file As String
Dim logType As String
Dim Username As String
Dim Password As String
Dim employeeId As String

direction = "?_direction_=" & pDirection

logType = "&_type_=" & pLogType

Username = "&_username_=" & pUsername

Password = "&_password_=" & pPassword

file = "&_file_=" & URLEncode(pFile)

If (pFile <> vbNullString) Then
    logType = "&_type_=" & Left$(pFile, 1)
End If

protocol = "https://"
'protocol = "http://"

If (empId <> vbNullString) Then
    employeeId = "&_employee_=" & empId
End If

generatedUrl = protocol & GlSeN & direction & logType & Username & Password & employeeId & file

If GlLog = True Then 'General Logging
    SOTLog generatedUrl, True
End If

GenerateUrl = True

End Function
Public Function WriteFile(ByVal filecontent As String, FileName As String, fileextension As String) As Boolean
    
On Error GoTo ErrorGenerated
    
Dim strBackSlash As String
Dim strFullPath As String

strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
strFullPath = App.Path & strBackSlash & "\LogFiles\" & FileName & "." & fileextension

Dim iFile As Integer

iFile = FreeFile
Open strFullPath For Output As #iFile
Print #iFile, filecontent
Close #iFile
WriteFile = True
       
ErrorGenerated:     WriteFile = False

End Function


Public Function URLEncode(sRawURL As String) As String
    Dim lngA As Long, strChar As String
    For lngA = 1 To Len(sRawURL)
        strChar = Mid$(sRawURL, lngA, 1)
        If strChar Like "[A-Za-z0-9]" Then
        ElseIf strChar = " " Then
            strChar = "+"
        Else
            strChar = "%" & Right$("0" & Hex$(Asc(strChar)), 2)
        End If
        URLEncode = URLEncode & strChar
    Next lngA
End Function
