VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDaSt As String
Private mDaEn As String
Private mDaAd As String
Private mZeSt As String
Private mZeEn As String
Private mKaNa As String
Private mTeBe As String
Private mTeID As String
Private mTeOr As String
Private mTeKo As String
Private mTePr As Integer
Private mTeGa As Boolean

Public Function ICSBody() As String
On Error GoTo LiErr
'iCalendar ICS Export

Dim TmStr As String

TmStr = "BEGIN:VEVENT" & vbCrLf
TmStr = TmStr & "UID:" & mTeID & vbCrLf
TmStr = TmStr & "CREATED:" & Format(mDaAd, "yyyymmdd") & "T000000" & vbCrLf
TmStr = TmStr & "DTSTART:" & Format(mDaSt, "yyyymmdd") & "T" & Format(mZeSt, "hhmmss") & vbCrLf
TmStr = TmStr & "DTEND:" & Format(mDaEn, "yyyymmdd") & "T" & Format(mZeEn, "hhmmss") & vbCrLf
TmStr = TmStr & "DTSTAMP:" & Format(Date, "yyyymmdd") & "T" & Format(Now, "hhmmss") & vbCrLf
TmStr = TmStr & "TRANSP:OPAQUE" & vbCrLf
TmStr = TmStr & "SUMMARY:" & mTeBe & vbCrLf
TmStr = TmStr & "LOCATION:" & mTeOr & vbCrLf
TmStr = TmStr & "DESCRIPTION:" & mTeKo & vbCrLf
TmStr = TmStr & "PRIORITY:" & mTePr & vbCrLf
TmStr = TmStr & "END:VEVENT" & vbCrLf

ICSBody = TmStr

Exit Function

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ICSBody " & Err.Number
Exit Function

End Function

Public Function ICSFoot() As String
On Error GoTo LiErr
'iCalendar ICS Export

Dim TmStr As String

TmStr = "END:VCALENDAR" & vbCrLf

ICSFoot = TmStr

Exit Function

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ICSFoot " & Err.Number
Exit Function

End Function

Public Function ICSHead() As String
On Error GoTo LiErr
'iCalendar ICS Export

Dim TmStr As String

TmStr = "BEGIN:VCALENDAR" & vbCrLf
TmStr = TmStr & "VERSION:2.0" & vbCrLf
TmStr = TmStr & "PRODID:SimpliMed23 Kalender" & vbCrLf
TmStr = TmStr & "METHOD:PUBLISH" & vbCrLf
TmStr = TmStr & "X-WR-CALNAME:" & mKaNa & vbCrLf
TmStr = TmStr & "X-WR-TIMEZONE:Europe/Berlin" & vbCrLf
TmStr = TmStr & "CALSCALE:GREGORIAN" & vbCrLf

' Define timezone with DST rules
TmStr = TmStr & "BEGIN:VTIMEZONE" & vbCrLf
TmStr = TmStr & "TZID:Europe/Berlin" & vbCrLf
TmStr = TmStr & "X-LIC-LOCATION:Europe/Berlin" & vbCrLf

' Daylight saving time
TmStr = TmStr & "BEGIN:DAYLIGHT" & vbCrLf
TmStr = TmStr & "DTSTART:19810329T020000" & vbCrLf
TmStr = TmStr & "TZOFFSETFROM:+0100" & vbCrLf
TmStr = TmStr & "TZOFFSETTO:+0200" & vbCrLf
TmStr = TmStr & "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU" & vbCrLf
TmStr = TmStr & "TZNAME:MESZ" & vbCrLf
TmStr = TmStr & "END:DAYLIGHT" & vbCrLf

' Standard time
TmStr = TmStr & "BEGIN:STANDARD" & vbCrLf
TmStr = TmStr & "DTSTART:19961027T030000" & vbCrLf
TmStr = TmStr & "TZOFFSETFROM:+0200" & vbCrLf
TmStr = TmStr & "TZOFFSETTO:+0100" & vbCrLf
TmStr = TmStr & "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU" & vbCrLf
TmStr = TmStr & "TZNAME:MEZ" & vbCrLf
TmStr = TmStr & "END:STANDARD" & vbCrLf

TmStr = TmStr & "END:VTIMEZONE" & vbCrLf

ICSHead = TmStr

Exit Function

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "ICSHead " & Err.Number
Exit Function

End Function

Public Property Let moDaAd(Flag As String)
    mDaAd = Flag
End Property
Public Property Get moDaAd() As String
    moDaAd = mDaAd
End Property

Public Property Let moDaEn(Flag As String)
    mDaEn = Flag
End Property
Public Property Get moDaEn() As String
    moDaEn = mDaEn
End Property

Public Property Let moDaSt(Flag As String)
    mDaSt = Flag
End Property
Public Property Get moDaSt() As String
    moDaSt = mDaSt
End Property

Public Property Let moKaNa(Flag As String)
    mKaNa = Flag
End Property
Public Property Get moKaNa() As String
    moKaNa = mKaNa
End Property

Public Property Let moTeGa(Flag As Boolean)
    mTeGa = Flag
End Property
Public Property Get moTeGa() As Boolean
    moTeGa = mTeGa
End Property

Public Property Let moTeBe(Flag As String)
    mTeBe = Flag
End Property
Public Property Get moTeBe() As String
    moTeBe = mTeBe
End Property

Public Property Let moTeID(Flag As String)
    mTeID = Flag
End Property
Public Property Get moTeID() As String
    moTeID = mTeID
End Property

Public Property Let moTeKo(Flag As String)
    mTeKo = Flag
End Property
Public Property Get moTeKo() As String
    moTeKo = mTeKo
End Property

Public Property Let moTeOr(Flag As String)
    mTeOr = Flag
End Property
Public Property Get moTeOr() As String
    moTeOr = mTeOr
End Property

Public Property Let moTePr(Flag As Integer)
    mTePr = Flag
End Property
Public Property Get moTePr() As Integer
    moTePr = mTePr
End Property

Public Property Let moZeEn(Flag As String)
    mZeEn = Flag
End Property
Public Property Get moZeEn() As String
    moZeEn = mZeEn
End Property

Public Property Let moZeSt(Flag As String)
    mZeSt = Flag
End Property
Public Property Get moZeSt() As String
    moZeSt = mZeSt
End Property

