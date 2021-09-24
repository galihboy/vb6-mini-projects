VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   4080
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   6000
      Top             =   3600
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6000
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------
' Developed by Galih Hermawan
' https://Galih.EU
'-------------------------------
Dim nama As String ' tulisan default pada label1
Dim iTimer1 As Integer ' speed atau interval awal timer1
Dim status1 As Boolean

Dim nama2 As String ' tulisan default pada label2
Dim iTimer2 As Integer ' speed atau interval awal timer2
Dim status2 As Boolean


Private Sub Command1_Click()
    If Timer1.Interval < 2000 Then Timer1.Interval = Timer1.Interval + 100
End Sub

Private Sub Command2_Click()
    If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 100
End Sub

Private Sub Command3_Click()
    Call Aktifkan(Timer1, Command3, Command1, Command2, Command4)
End Sub

Private Sub Command4_Click()
    Label3.Caption = "Speed: 0 ms."
    Timer1.Interval = iTimer1
    Call Aktifkan(Timer1, Command3, Command1, Command2, Command4, True)
    status1 = False
    Label1.Caption = nama
End Sub

Private Sub Command5_Click()
    If Timer2.Interval < 2000 Then Timer2.Interval = Timer2.Interval + 100
End Sub

Private Sub Command6_Click()
    If Timer2.Interval > 100 Then Timer2.Interval = Timer2.Interval - 100
End Sub

Private Sub Command7_Click()
    Call Aktifkan(Timer2, Command7, Command5, Command6, Command8)
End Sub

Private Sub Command8_Click()
    Label4.Caption = "Speed: 0 ms."
    Timer2.Interval = iTimer2
    Call Aktifkan(Timer2, Command7, Command5, Command6, Command8, True)
    status2 = False
    Label2.Caption = nama2
End Sub

Private Sub Form_Load()
    Form1.Caption = "Animasi Teks - by Galih Hermawan"

' Setting untuk keperluan label1 pada tulisan Galih Hermawan
    nama = "Galih Hermawan"
    iTimer1 = 1000
    
    Label1.Caption = nama
    Label1.FontSize = 36
    Label1.ForeColor = vbBlue
    Label1.Alignment = 2 ' rata tengah (center)
    Label3.Caption = "Speed: " & Timer1.Interval & " ms."
    
    Command1.Caption = "- Perlambat"
    Command1.Enabled = False
    Command2.Caption = "+ Percepat"
    Command2.Enabled = False
    Command3.Caption = "Start"
    Command4.Caption = "Stop && Reset"
    Command4.Enabled = False
    'status1 = False
    Command4.Value = True
    
    ' Setting untuk keperluan label1 pada tulisan Galih Hermawan
    nama2 = "Forum Informatika"
    iTimer2 = 1000
    
    Label2.Caption = nama2
    Label2.FontSize = 32
    Label2.FontName = "Courier"
    Label2.ForeColor = RGB(150, 40, 10)
    Label2.Alignment = 2 ' rata tengah (center)
    Label4.Caption = "Speed: " & Timer2.Interval & " ms."
    
    Command5.Caption = "- Perlambat"
    Command5.Enabled = False
    Command6.Caption = "+ Percepat"
    Command6.Enabled = False
    Command7.Caption = "Start"
    Command8.Caption = "Stop && Reset"
    Command8.Enabled = False
    Timer2.Interval = iTimer2
    Check1.Caption = "Perkarakter"
    Option1.Caption = "Kiri"
    Option1.Value = True
    Option2.Caption = "Tengah"
    Option3.Caption = "Kanan"
End Sub

Private Sub Label1_Click()
    Call Aktifkan(Timer1, Command3, Command1, Command2, Command4)
End Sub

Private Sub Label2_Click()
    Call Aktifkan(Timer2, Command7, Command5, Command6, Command8)
End Sub

Private Sub Option1_Click()
    Label2.Alignment = 0
End Sub

Private Sub Option2_Click()
    Label2.Alignment = 2
End Sub

Private Sub Option3_Click()
    Label2.Alignment = 1
End Sub

Private Sub Timer1_Timer()
    Dim p As Integer, t As String
    Dim sAwal As String, sAkhir As String
    Static i As Integer
    
    t = nama
    p = Len(t)
    
    If status1 = False Then
        i = 0
        status1 = True
    End If
    
    i = i + 1
    sAwal = Mid(t, i + 1, p - i)
    sAkhir = Mid(t, 1, i)
    Label1 = sAwal & " " & sAkhir
    If i = p Then i = 0
    
    Label3.Caption = "Speed: " & Timer1.Interval & " ms."
End Sub

Private Sub Aktifkan(tmr As Timer, cmd As CommandButton, cmdMin As CommandButton, cmdPlus As CommandButton, cmdReset As CommandButton, Optional nyala As Boolean)
    If tmr.Enabled Or nyala Then
        cmd.Caption = "Start"
        cmdMin.Enabled = False
        cmdPlus.Enabled = False
        cmdReset.Enabled = False
        tmr.Enabled = False
    Else
        cmd.Caption = "Pause"
        cmdMin.Enabled = True
        cmdPlus.Enabled = True
        cmdReset.Enabled = True
        tmr.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    Dim p As Integer, t As String
    Dim sAwal As String, sAkhir As String
    Static i As Integer
    
    t = nama2
    p = Len(t)
    
    If status2 = False Then
        i = 0
        status2 = True
    End If
    
    i = i + 1
    
    If Check1.Value = False Then
        sAwal = Mid(t, 1, i) ' & " "
        Label2 = sAwal & " "
        If i = p Then i = 0
    Else
        If Option3.Value = True Then
            sAwal = Mid(t, p + 1 - (i - 1), 1)
            Label2 = sAwal & Space(i)
        ElseIf Option2.Value = True Then
            sAwal = Mid(t, i, 1)
            Label2 = Space(i) & sAwal & Space(i)
        Else
            sAwal = Mid(t, i, 1)
            Label2 = Space(i) & sAwal
        End If
        If i = p Then
            Label2.Caption = nama2
        ElseIf i = p + 1 Then
            i = 0
        End If
    End If
    
    Label4.Caption = "Speed: " & Timer2.Interval & " ms."
End Sub
