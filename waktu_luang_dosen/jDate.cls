VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "jDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2319
Option Explicit

Public Enum jDateTypes
    jDateMMGG
    jDateGGMM
End Enum

Private pDateSeparator As String
Public DateFormat As jDateTypes
Attribute DateFormat.VB_VarHelpID = 2320
Public LongYear As Boolean
Attribute LongYear.VB_VarHelpID = 2321

Private Const validSeps As String = "-./"
Private pYear As Long
Private pMonth As Long
Private pDay As Long

Public Property Get DateSeparator() As String
Attribute DateSeparator.VB_HelpID = 2322

DateSeparator = pDateSeparator

End Property

Public Property Let DateSeparator(ByVal NewDateSep As String)
If Len(DateSeparator) = 1 Then
    pDateSeparator = NewDateSep
Else
    Throw Me, "DateSeparator.Let", "Invalid date separator"
End If

End Property

Public Property Get jDate() As String
Attribute jDate.VB_HelpID = 2323
' ritorna una stringa del tipo GG/MM/YYYY in base
' alle impostazioni dell'oggetto
Dim s As String
Select Case DateFormat
    Case jDateMMGG
        s = "MM" & DateSeparator & "GG" & DateSeparator
    Case jDateGGMM
        s = "GG" & DateSeparator & "MM" & DateSeparator
    Case Else
        Throw Me, "jDate.Get", "Formato data non valido"
End Select

If LongYear Then
    s = s & "AAAA"
Else
    s = s & "AA"
End If

jDate = s

End Property


Public Function CreateDateMask(Optional ByVal PlaceHolder As String = "_") As String
Attribute CreateDateMask.VB_HelpID = 2324
Dim s As String
    
If Len(PlaceHolder) = 1 Then
    s = String(2, PlaceHolder) & DateSeparator
    s = s & s
    If LongYear Then
        s = s & String(4, PlaceHolder)
    Else
        s = s & String(2, PlaceHolder)
    End If
    CreateDateMask = s
Else
    Throw Me, "CreateDateMask", "Invalid PlaceHolder"
End If
End Function

Private Sub Class_Initialize()
' carica valori di default del sistema
Dim s As String
Dim sYear As String
Dim vNum  As String, i As Long

vNum = "0123456789"

s = Date
For i = 1 To 10
    s = Replace(s, Mid(vNum, i, 1), "")
Next i

pDateSeparator = Left(s, 1)

If InStr(validSeps, pDateSeparator) = 0 Then
    Throw Me, "Initialize", "Formato data sistema non supportato dall'applicazione"
End If

LongYear = Len(Date) > 8
' verifica formato data
If LongYear Then
    sYear = "2001"
Else
    sYear = "01"
End If
s = "01" & pDateSeparator & "02" & pDateSeparator & sYear

If s = DateSerial(2001, 2, 1) Then
    DateFormat = jDateGGMM
Else
    s = "02" & pDateSeparator & "01" & pDateSeparator & sYear
    If s = DateSerial(2001, 2, 1) Then
        DateFormat = jDateMMGG
    Else
        Throw Me, "Initialize", "This system date format is not supported by the application"
    End If
End If
pYear = Year(Date)
pMonth = Month(Date)
pDay = Day(Date)
End Sub

Public Sub SetDate(ByVal jYear As Long, ByVal jMonth As Long, ByVal jDay As Long)
Attribute SetDate.VB_HelpID = 2325
Dim d As Date

d = DateSerial(jYear, jMonth, jDay) ' verifica correttezza dati
pYear = jYear
pMonth = jMonth
pDay = jDay

End Sub

Public Function GetDateLong() As Long
Attribute GetDateLong.VB_HelpID = 2326
GetDateLong = pYear * 10000 + pMonth * 100 + pDay
End Function



Public Sub SetDateLong(ByVal jLongDate As Long)
Attribute SetDateLong.VB_HelpID = 2327
Dim y As Long
Dim M As Long
Dim d As Long, dt As Date
y = jLongDate / 10000
M = (jLongDate - y * 10000) / 100
d = jLongDate - y * 10000 - M * 100
dt = DateSerial(y, M, d) ' verifica correttezza dati
If EqualDate(d, M, y, dt) Then
    pYear = Year(dt)
    pMonth = Month(dt)
    pDay = Day(dt)
Else
    Throw Me, "SetDateLong", "Invalid date"
End If

End Sub

Public Function GetDate() As String
Attribute GetDate.VB_HelpID = 2328
Attribute GetDate.VB_UserMemId = 0
' predefinito
Dim y As String
y = Format(pYear, "0000")
If Not LongYear Then y = Right(y, 2)
    
If DateFormat = jDateGGMM Then
    GetDate = Format(pDay, "00") & pDateSeparator _
        & Format(pMonth, "00") & pDateSeparator _
        & y
Else
    GetDate = Format(pMonth, "00") & pDateSeparator _
        & Format(pDay, "00") & pDateSeparator _
        & y
End If

End Function

Public Sub SetDateString(ByVal StrDate As String)
Attribute SetDateString.VB_HelpID = 2329
' si suppone di avere una stringa nel formato corrente
' dell'oggetto
Dim v As Variant, ErrON As Boolean
Dim dt As Date, dDay As Long, dYear As Long, dMonth As Long
v = Split(StrDate, pDateSeparator)
If UBound(v) = 2 Then
    If DateFormat = jDateGGMM Then
        dDay = v(0)
        dMonth = v(1)
    Else
        dDay = v(1)
        dMonth = v(0)
    End If
    dYear = v(2)
    Select Case Len(v(2))
        Case 2: ErrON = LongYear
        Case 4: ErrON = Not LongYear
        Case Else: ErrON = True
    End Select
    dt = DateSerial(dYear, dMonth, dDay)
    ErrON = ErrON Or Not EqualDate(dDay, dMonth, dYear, dt)
    If Not ErrON Then
        pDay = dDay
        pMonth = dMonth
        pYear = dYear
    End If
Else
    ErrON = True
End If

If ErrON Then
    Throw Me, "SetDateString", "Invalid Date"
End If
End Sub

Public Property Let ShortDateFormatDescription(ByVal SDFD As String)
Attribute ShortDateFormatDescription.VB_HelpID = 2331
Dim Gmese As String
Dim Lyear As String
Dim SpChar As String
If Len(SDFD) = 3 Then
    Gmese = Mid(SDFD, 1, 1)
    Lyear = Mid(SDFD, 2, 1)
    SpChar = Mid(SDFD, 3, 1)
    Select Case Gmese
        Case "G":   DateFormat = jDateGGMM
        Case "M":   DateFormat = jDateMMGG
        Case Else:  Throw Me, "ShortDateFormatDescription.Let", "Invalid Month descriptor"
    End Select
    Select Case Lyear
        Case "L":   LongYear = True
        Case "S":   LongYear = False
        Case Else:  Throw Me, "ShortDateFormatDescription.Let", "Invalid Year descriptor"
    End Select
    pDateSeparator = SpChar
Else
    Throw Me, "ShortDateFormatDescription.Let", "Invalid ShortDescription"
End If
End Property
Public Property Get ShortDateFormatDescription() As String
Dim s As String

s = IIf(DateFormat = jDateGGMM, "G", "M")
s = s & IIf(LongYear, "L", "S")
s = s & pDateSeparator
If Len(s) <> 3 Then Throw Me, "ShortDateFormatDescription.get", "Unattended exception"
ShortDateFormatDescription = s
End Property

Public Function ParseDateLong(ByVal LongDate As Long) As String
Attribute ParseDateLong.VB_HelpID = 2332
Dim jDT As New jDate

With jDT
    .ShortDateFormatDescription = Me.ShortDateFormatDescription
End With ' si clona il formato
jDT.SetDateLong (LongDate)
ParseDateLong = jDT.GetDate
End Function

Private Function EqualDate(ByVal cDay As Long, ByVal cMonth As Long, ByVal cYear As Long, ByVal EqDate As Date) As Boolean

EqualDate = (Day(EqDate) = cDay And _
            Month(EqDate) = cMonth And _
            Year(EqDate) Mod 100 = cYear Mod 100)
End Function



